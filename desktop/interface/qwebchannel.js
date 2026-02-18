"use strict";

var QWebChannelMessageTypes = {
    signal: 1,
    propertyUpdate: 2,
    init: 3,
    idle: 4,
    debug: 5,
    invokeMethod: 6,
    connectToSignal: 7,
    disconnectFromSignal: 8,
    setProperty: 9,
    response: 10,
};

var QWebChannel = function(transport, initCallback)
{
    if (typeof transport !== "object" || typeof transport.send !== "function") {
        console.error("The QWebChannel expects a transport object with a send function and onmessage callback property." +
                      " Given is: transport: " + typeof(transport) + ", transport.send: " + typeof(transport.send));
        return;
    }

    var channel = this;
    this.transport = transport;

    this.send = function(data)
    {
        if (typeof(data) !== "string") {
            data = JSON.stringify(data);
        }
        channel.transport.send(data);
    }

    this.execCallbacks = {};
    this.execId = 1;
    
    this.execCallbacks[0] = function(data) {
        for (var objectName in data) {
            var object = new QObject(objectName, data[objectName], channel);
        }
        if (initCallback) {
            initCallback(channel);
        }
    };

    this.objects = {};

    this.handleSignal = function(message)
    {
        var object = channel.objects[message.object];
        if (object) {
            object.signalEmitted(message.signal, message.args);
        } else {
            console.warn("Unhandled signal: " + message.object + "::" + message.signal);
        }
    }

    this.handleResponse = function(message)
    {
        if (!message.hasOwnProperty("id")) {
            console.error("Invalid response message received: ", message);
            return;
        }
        if (channel.execCallbacks[message.id]) {
            channel.execCallbacks[message.id](message.data);
            delete channel.execCallbacks[message.id];
        } else {
             console.warn("No callback found for response id: " + message.id);
        }
    }

    this.handlePropertyUpdate = function(message)
    {
        for (var i in message.data) {
            var data = message.data[i];
            var object = channel.objects[data.object];
            if (object) {
                object.propertyUpdate(data.signals, data.properties);
            } else {
                console.warn("Unhandled property update: " + data.object + "::" + data.signal);
            }
        }
        // Removed incorrect callback call here
    }

    this.transport.onmessage = function(message)
    {
        var data = message.data;
        if (typeof data === "string") {
            data = JSON.parse(data);
        }
        switch (data.type) {
            case QWebChannelMessageTypes.signal:
                channel.handleSignal(data);
                break;
            case QWebChannelMessageTypes.response:
                channel.handleResponse(data);
                break;
            case QWebChannelMessageTypes.propertyUpdate:
                channel.handlePropertyUpdate(data);
                break;
            default:
                console.error("Invalid message received: ", data);
                break;
        }
    }

    channel.send({type: QWebChannelMessageTypes.init, id: 0});
};

var QObject = function(name, data, webChannel)
{
    this.__id__ = name;
    webChannel.objects[name] = this;

    // List of callbacks that get invoked upon signal emission
    this.__objectSignals__ = {};

    // Cache of all properties, updated when a notify signal is emitted
    this.__propertyCache__ = {};

    var object = this;

    // ----------------------------------------------------------------------

    this.unwrapQObject = function(response)
    {
        if (response instanceof Array) {
            // support list of objects
            var ret = new Array(response.length);
            for (var i = 0; i < response.length; ++i) {
                ret[i] = object.unwrapQObject(response[i]);
            }
            return ret;
        }
        if (!response
            || !response["__QObject*__"]
            || response.id === undefined) {
            return response;
        }

        var objectId = response.id;
        if (webChannel.objects[objectId])
            return webChannel.objects[objectId];

        if (!response.data) {
            console.error("Cannot unwrap unknown QObject " + objectId + " without data.");
            return;
        }

        var qObject = new QObject( objectId, response.data, webChannel );
        qObject.destroyed.connect(function() {
            if (webChannel.objects[objectId] === qObject) {
                delete webChannel.objects[objectId];
                // reset the now deleted QObject to a null QObject
                var property = {};
                property[objectId] = null;
                webChannel.handlePropertyUpdate({data: [{object: objectId, properties: property}]});
            }
        });
        return qObject;
    }

    this.unwrapQObjectProperties = function(props)
    {
        for (var prop in props) {
            props[prop] = object.unwrapQObject(props[prop]);
        }
    }

    this.propertyUpdate = function(signals, propertyMap)
    {
        // update property cache
        for (var propertyIndex in propertyMap) {
            var propertyValue = propertyMap[propertyIndex];
            object.__propertyCache__[propertyIndex] = propertyValue;
        }

        for (var signalName in signals) {
            var args = signals[signalName];
            object.__objectSignals__[signalName].forEach(function(callback) {
                callback.apply(callback, args);
            });
        }
    }

    this.signalEmitted = function(signalName, signalArgs)
    {
        object.__objectSignals__[signalName].forEach(function(callback) {
            callback.apply(callback, signalArgs);
        });
    }

    // ----------------------------------------------------------------------

    var addSignal = function(signalData, isPropertyNotifySignal)
    {
        var signalName = signalData[0];
        var signalIndex = signalData[1];

        if (!object.__objectSignals__[signalIndex]) {
            object.__objectSignals__[signalIndex] = [];
        }

        object[signalName] = {
            connect: function(callback) {
                if (typeof(callback) !== "function") {
                    console.error("Bad callback given to connect to signal " + signalName);
                    return;
                }

                object.__objectSignals__[signalIndex].push(callback);

                if (!isPropertyNotifySignal && signalName !== "destroyed") {
                    // only required for "pure" signals, handled separately for properties in propertyUpdate
                    // also destroyed signal doesn't need this because the object is already gone
                    webChannel.send({
                        type: QWebChannelMessageTypes.connectToSignal,
                        object: object.__id__,
                        signal: signalIndex
                    });
                }
            },
            disconnect: function(callback) {
                if (typeof(callback) !== "function") {
                    console.error("Bad callback given to disconnect from signal " + signalName);
                    return;
                }
                object.__objectSignals__[signalIndex] = object.__objectSignals__[signalIndex].filter(function(c) {
                    return c !== callback;
                });
                if (!isPropertyNotifySignal && signalName !== "destroyed") {
                    // only required for "pure" signals, handled separately for properties in propertyUpdate
                    // also destroyed signal doesn't need this because the object is already gone
                    webChannel.send({
                        type: QWebChannelMessageTypes.disconnectFromSignal,
                        object: object.__id__,
                        signal: signalIndex
                    });
                }
            }
        };
    }

    // ----------------------------------------------------------------------

    var addMethod = function(methodData)
    {
        var methodName = methodData[0];
        var methodIndex = methodData[1];
        object[methodName] = function() {
            var args = [];
            var callback;
            for (var i = 0; i < arguments.length; ++i) {
                var argument = arguments[i];
                if (typeof argument === "function")
                    callback = argument;
                else if (argument instanceof QObject && webChannel.objects[argument.__id__])
                    args.push({
                        "__QObject*__": true,
                        "id": argument.__id__
                    });
                else
                    args.push(argument);
            }

            var execId = webChannel.execId++;
            webChannel.execCallbacks[execId] = function(data) {
                if (callback) {
                    var result = object.unwrapQObject(data);
                    callback(result);
                }
            };

            webChannel.send({
                type: QWebChannelMessageTypes.invokeMethod,
                object: object.__id__,
                method: methodIndex,
                args: args,
                id: execId
            });
        };
    }

    // ----------------------------------------------------------------------

    var addProperty = function(propertyData)
    {
        var propertyName = propertyData[0];
        var propertyIndex = propertyData[1];
        var propertyNotifySignal = propertyData[2];

        // initialize property cache with current value
        // NOTE: if this is an object, it is not yet unwrapped.
        object.__propertyCache__[propertyIndex] = propertyData[3];

        if (propertyNotifySignal) {
            if (object.__objectSignals__[propertyNotifySignal] === undefined) {
                object.__objectSignals__[propertyNotifySignal] = [];
                addSignal([propertyNotifySignal + "Changed", propertyNotifySignal], true);
            }
        }

        Object.defineProperty(object, propertyName, {
            configurable: true,
            get: function () {
                var propertyValue = object.__propertyCache__[propertyIndex];
                if (propertyValue === undefined) {
                    // This shouldn't happen
                    console.warn("Undefined value in property cache for " + propertyName);
                }
                return object.unwrapQObject(propertyValue);
            },
            set: function(value) {
                if (value === undefined) {
                    console.warn("Property setter for " + propertyName + " called with undefined");
                    return;
                }
                object.__propertyCache__[propertyIndex] = value;
                webChannel.send({
                    type: QWebChannelMessageTypes.setProperty,
                    object: object.__id__,
                    property: propertyIndex,
                    value: value
                });
            }
        });
    }

    // ----------------------------------------------------------------------

    data.methods.forEach(addMethod);
    data.properties.forEach(addProperty);
    data.signals.forEach(function(signal) { addSignal(signal, false); });

    for (var name in data.enums) {
        object[name] = data.enums[name];
    }
};

// Required for use with node.js
if (typeof module === 'object') {
    module.exports = {
        QWebChannel: QWebChannel
    };
}
