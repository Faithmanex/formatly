// Utility: Escape HTML to prevent XSS
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// State
let currentView = 'dashboard';
let stats = {
    documents: 0,
    processing: 0,
    formatted: 0,
    plan_usage: 0
};
let currentUserPrefs = {};
let bridge = null; // Qt Bridge

// DOM Elements
const views = document.querySelectorAll('.view');
const navItems = document.querySelectorAll('.nav-item[data-view]');
const fileInput = document.getElementById('fileInput');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    // Initialize Qt WebChannel
    if (typeof QWebChannel !== 'undefined') {
        new QWebChannel(qt.webChannelTransport, function(channel) {
            bridge = channel.objects.bridge;
            
            // Connect Signals
            bridge.progressUpdated.connect(window.updateProgress);
            bridge.formattingFinished.connect(window.formattingFinished);
            bridge.notificationsUpdated.connect(loadNotifications); 
            bridge.aiResponseReceived.connect(window.handleAIResponse); 
            bridge.aiResponseStart.connect(window.startAIStream); // New
            bridge.aiResponseChunk.connect(window.handleAIChunk); // New

            // Initial Data Load
            loadDashboardData();
            loadPreferences();
            loadNotifications();
        });
    } else {
        console.error("QWebChannel not found!");
    }

    setupNavigation();
    setupButtons();
    setupTabs();
    setupSettings();
    setupSidebar();
    setupSettings();
    setupSidebar();
    setupSettings();
    setupSidebar();
    setupChat(); // Initialize Chat Logic
    setupCustomSelects(); // Transform selects to glassmorphism
    
    // Check view state
    if (currentView === 'documents') {
        loadDocumentsTable(); // Will defer if bridge not ready
    }
});

function loadNotifications() {
     if (bridge) {
        bridge.get_notifications(function(notifications) {
            // TODO: Render notifications if UI exists
            // For now just logging or simple badge update could go here
        });
    }
}


// Navigation Logic
function setupNavigation() {
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const targetView = item.getAttribute('data-view');
            navigateTo(targetView);
        });
    });
}

function navigateTo(viewId) {
    // Update State
    currentView = viewId;

    // Update UI
    views.forEach(view => {
        view.style.display = view.id === `view-${viewId}` ? 'block' : 'none';
    });

    navItems.forEach(item => {
        if (item.getAttribute('data-view') === viewId) {
            item.classList.add('active');
        } else {
            item.classList.remove('active');
        }
    });

    // Refresh Data if needed
    if (viewId === 'documents') {
        loadDocumentsTable();
    }
}

// Button Actions
function setupButtons() {
    const uploadBtns = [
        document.getElementById('quickUploadBtn'),
        document.getElementById('mainUploadBtn'),
        document.getElementById('docViewUploadBtn'),
        document.getElementById('actionUpload')
    ];

    uploadBtns.forEach(btn => {
        if(btn) {
            btn.addEventListener('click', () => {
                if (bridge) {
                    bridge.browse_file(function(filePath) {
                        if (filePath) {
                            showUploadOptions(filePath);
                        }
                    });
                }
            });
        }
    });

    document.getElementById('saveStylesBtn').addEventListener('click', saveStyles);
    
    // Upload Modal Buttons
    document.getElementById('cancelUploadBtn').addEventListener('click', cancelUpload);
    document.getElementById('closeUploadModalBtn').addEventListener('click', cancelUpload);
    document.getElementById('confirmUploadBtn').addEventListener('click', confirmUpload);
    
    // Formatting Modal Buttons
    document.getElementById('cancelFormatBtn').addEventListener('click', cancelFormatting);
    document.getElementById('minimizeFormatBtn').addEventListener('click', minimizeFormatting);
    document.getElementById('restoreFormatBtn').addEventListener('click', restoreFormatting);
}

function setupSidebar() {
    const toggleBtn = document.getElementById('sidebarToggle');
    const sidebar = document.querySelector('.sidebar');
    
    if (toggleBtn && sidebar) {
        toggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
        });
    }
}

// Tab Logic
function setupTabs() {
    const segmentBtns = document.querySelectorAll('.segment-btn');
    segmentBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            // Remove active from all buttons
            segmentBtns.forEach(b => b.classList.remove('active'));
            // Add active to clicked
            e.target.classList.add('active');
            
            // Hide all tab content
            document.querySelectorAll('.tab-content').forEach(tab => tab.style.display = 'none');
            
            // Show target
            const tabId = e.target.getAttribute('data-tab');
            document.getElementById(tabId).style.display = 'block';
        });
    });
}

// Upload Options Logic
let selectedUploadPath = null;

function showUploadOptions(filePath) {
    selectedUploadPath = filePath;
    const modal = document.getElementById('uploadOptionsModal');
    
    // Show filename (extract from path)
    const filename = filePath.split(/[\\/]/).pop();
    document.getElementById('uploadFileName').textContent = `Selected: ${filename}`;
    
    // Load defaults from saved prefs
    const prefs = currentUserPrefs;
    if (prefs) {
        if(prefs.style) document.getElementById('uploadStyle').value = prefs.style;
        if(prefs.backend) document.getElementById('uploadBackend').value = prefs.backend;
        if(prefs.englishVariant) document.getElementById('uploadEnglishVariant').value = prefs.englishVariant;
        if(prefs.trackChanges !== undefined) document.getElementById('uploadTrackChanges').checked = prefs.trackChanges;
    }

    modal.style.display = 'flex';
    syncAllCustomSelects(); // Sync the custom dropdowns after updating native selects
}

function cancelUpload() {
    selectedUploadPath = null;
    document.getElementById('uploadOptionsModal').style.display = 'none';
}

function confirmUpload() {
    if (!selectedUploadPath) return;

    const options = {
        input: selectedUploadPath,
        style: document.getElementById('uploadStyle').value,
        english_variant: document.getElementById('uploadEnglishVariant').value,
        backend: document.getElementById('uploadBackend').value,
        track_changes: document.getElementById('uploadTrackChanges').checked,
        overwrite: true
    };

    // Close options modal
    document.getElementById('uploadOptionsModal').style.display = 'none';
    
    // Start processing
    startFormattingProcess(options);
}


// Custom Modals
function showAlert(title, message) {
    const modal = document.getElementById('alertModal');
    document.getElementById('alertTitle').textContent = title;
    document.getElementById('alertMessage').textContent = message;
    
    modal.style.display = 'flex';
    
    // One time listener for OK
    const okBtn = document.getElementById('alertOkBtn');
    const closeHandler = () => {
        modal.style.display = 'none';
        okBtn.removeEventListener('click', closeHandler);
    };
    okBtn.addEventListener('click', closeHandler);
}

function showConfirm(title, message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmMessage').textContent = message;
    
    modal.style.display = 'flex';
    
    const okBtn = document.getElementById('confirmOkBtn');
    const cancelBtn = document.getElementById('confirmCancelBtn');
    
    // Clear old listeners by cloning
    // Helper to safely replace element to remove listeners
    const replace = (el) => {
        const newEl = el.cloneNode(true);
        el.parentNode.replaceChild(newEl, el);
        return newEl;
    };
    
    const newOk = replace(okBtn);
    const newCancel = replace(cancelBtn);
    
    newOk.addEventListener('click', () => {
        modal.style.display = 'none';
        if (onConfirm) onConfirm();
    });
    
    newCancel.addEventListener('click', () => {
        modal.style.display = 'none';
    });
}


// Settings
function setupSettings() {
    // Theme Switch Listener
    const themeSelect = document.getElementById('prefTheme');
    if (themeSelect) {
        themeSelect.addEventListener('change', (e) => {
            if (e.target.value === 'Light Mode') {
                document.body.classList.add('light-mode');
            } else {
                document.body.classList.remove('light-mode');
            }
        });
    }
}

function loadPreferences() {
    if (bridge) {
        bridge.get_preferences(function(prefs) {
            currentUserPrefs = prefs || {};
            applyPreferencesToUI(currentUserPrefs);
        });
    }
}

function applyPreferencesToUI(prefs) {
    if (!prefs) return;
    
    if (document.getElementById('prefStyle') && prefs.style) {
        document.getElementById('prefStyle').value = prefs.style;
    }
    if (document.getElementById('prefBackend') && prefs.backend) {
        document.getElementById('prefBackend').value = prefs.backend;
    }
    if (document.getElementById('prefEnglishVariant') && prefs.englishVariant) {
        document.getElementById('prefEnglishVariant').value = prefs.englishVariant;
    }
    if (document.getElementById('prefCitationStyle') && prefs.citationStyle) {
        document.getElementById('prefCitationStyle').value = prefs.citationStyle;
    }
    if (document.getElementById('prefDefaultFont') && prefs.defaultFont) {
        document.getElementById('prefDefaultFont').value = prefs.defaultFont;
    }
    if (document.getElementById('prefTrackChanges') && prefs.trackChanges !== undefined) {
        document.getElementById('prefTrackChanges').checked = prefs.trackChanges;
    }
    
    // Theme
    if (prefs.theme === 'light') {
        document.body.classList.add('light-mode');
        const themeSelect = document.getElementById('prefTheme');
        if (themeSelect) themeSelect.value = 'Light Mode';
    } else {
        document.body.classList.remove('light-mode'); // Default to dark
         const themeSelect = document.getElementById('prefTheme');
        if (themeSelect) themeSelect.value = 'Dark Mode';
    }
    
    syncAllCustomSelects(); // Sync UI after loading prefs
}

function saveStyles() {
    const themeSelect = document.getElementById('prefTheme');
    const isLight = themeSelect && themeSelect.value === 'Light Mode';
    
    const prefs = {
        style: document.getElementById('prefStyle').value,
        backend: document.getElementById('prefBackend').value,
        englishVariant: document.getElementById('prefEnglishVariant').value,
        citationStyle: document.getElementById('prefCitationStyle').value,
        defaultFont: document.getElementById('prefDefaultFont').value,
        trackChanges: document.getElementById('prefTrackChanges').checked,
        theme: isLight ? 'light' : 'dark'
    };
    
    localStorage.setItem('formatly_prefs', JSON.stringify(prefs));
    currentUserPrefs = prefs;
    
    // Save to Backend
    if (bridge) {
        bridge.save_preferences(prefs, function() {
            showAlert('Success', 'Preferences saved to disk!');
        });
    } else {
        showAlert('Success', 'Preferences saved (Simulation)!');
    }

    // Apply theme
    if (isLight) {
        document.body.classList.add('light-mode');
    } else {
        document.body.classList.remove('light-mode');
    }
}

// Formatting Process
function startFormattingProcess(options) {
    // Show Modal
    const modal = document.getElementById('formattingModal');
    modal.style.display = 'flex';
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('progressText').textContent = 'Starting...';

    // Call Backend
    if (bridge) {
        bridge.start_formatting(options);
    }
}

function cancelFormatting() {
    if (bridge) {
        bridge.cancel_formatting();
    }
    document.getElementById('formattingModal').style.display = 'none';
    document.getElementById('floatingProgress').style.display = 'none';
}

function minimizeFormatting() {
    document.getElementById('formattingModal').style.display = 'none';
    const fp = document.getElementById('floatingProgress');
    fp.style.display = 'block';
    
    // Sync current state
    const text = document.getElementById('progressText').textContent;
    const width = document.getElementById('progressFill').style.width;
    
    document.getElementById('fpProgressText').textContent = text;
    document.getElementById('fpProgressFill').style.width = width;
}

function restoreFormatting() {
    document.getElementById('floatingProgress').style.display = 'none';
    document.getElementById('formattingModal').style.display = 'flex';
}

// Called from Backend via Signal
window.updateProgress = function(message, progress) {
    // Update Modal
    const modal = document.getElementById('formattingModal');
    // If neither is visible, show modal (initial start)
    const fp = document.getElementById('floatingProgress');
    
    if (modal.style.display === 'none' && fp.style.display === 'none') {
        modal.style.display = 'flex';
    }
    
    document.getElementById('progressText').textContent = message;
    document.getElementById('progressFill').style.width = `${progress}%`;
    
    // Update Floating
    document.getElementById('fpProgressText').textContent = message;
    document.getElementById('fpProgressFill').style.width = `${progress}%`;
};

window.formattingFinished = function(success, message, outputPath) {
    // Add small delay to let progress bar fill
    setTimeout(() => {
        document.getElementById('formattingModal').style.display = 'none';
        document.getElementById('floatingProgress').style.display = 'none';
        
        // Refresh stats immediately
        loadDashboardData();
        
        if (success) {
            showConfirm(
                'Formatting Complete!', 
                'Open file?', 
                () => {
                    if (bridge) bridge.open_file(outputPath);
                }
            );
        } else {
            showAlert('Error', message);
        }
    }, 500);
};

// Data Loading Functions
function loadDashboardData() {
    if (bridge) {
        bridge.get_history(function(history) {
            renderRecentDocs(history);
        });
        
        bridge.get_stats(function(data) {
             renderStats(data);
        });

        bridge.get_user_info(function(user) {
            renderUserInfo(user);
        });
    }
}

function loadDocumentsTable() {
    if (bridge) {
        bridge.get_history(function(history) {
            renderDocumentsTable(history);
        });
    }
}

function renderStats(data) {
    const ids = ['statDocuments', 'statProcessing', 'statFormatted', 'statPlanUsage', 'statPlan', 'docCountText', 'sidebarUsageText'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.classList.remove('loading');
            if (id === 'statDocuments') el.textContent = data.documents;
            if (id === 'statProcessing') el.textContent = data.processing;
            if (id === 'statFormatted') el.textContent = data.formatted;
            if (id === 'statPlanUsage') el.textContent = data.plan_usage;
            if (id === 'statPlan') el.textContent = 'Pro'; 
            if (id === 'docCountText') el.textContent = `${data.documents} Documents`;
            if (id === 'sidebarUsageText') el.textContent = data.plan_usage;
        }
    });

    // Update Sidebar progress bar
    const usageFill = document.getElementById('sidebarUsageFill');
    if (usageFill && data.plan_usage) {
        const parts = data.plan_usage.split('/');
        if (parts.length === 2) {
            const count = parseInt(parts[0]);
            const total = parseInt(parts[1]);
            const percent = Math.min(100, Math.round((count / total) * 100));
            usageFill.style.width = `${percent}%`;
        }
    }
}

function renderUserInfo(user) {
    const nameEl = document.getElementById('sidebarUserName');
    const planEl = document.getElementById('sidebarPlan');
    const welcomeEl = document.getElementById('welcomeName');

    if (nameEl) {
        nameEl.classList.remove('loading');
        nameEl.textContent = user.name;
    }
    if (planEl) {
        planEl.classList.remove('loading');
        planEl.textContent = user.plan;
    }
    if (welcomeEl) {
        welcomeEl.classList.remove('loading');
        const firstName = user.name.split(' ')[0];
        welcomeEl.textContent = `Welcome, ${firstName}`;
        
        // Update Chat Greeting
        const chatGreeting = document.getElementById('chatGreeting');
        if (chatGreeting) {
            chatGreeting.textContent = `Hi ${firstName}`;
        }
    }
}

function renderRecentDocs(history) {
    const list = document.getElementById('recentDocsList');
    if (!list) return;
    
    list.innerHTML = '';
    
    // Show only the 5 most recent documents
    const recent = history.slice(0, 5);
    
    if (recent.length === 0) {
        list.innerHTML = `
            <div class="empty-state">
                <div class="icon">📂</div>
                <p>No recent documents</p>
            </div>
        `;
        return;
    }

    recent.forEach(doc => {
        const safePath = doc.path ? doc.path.replace(/\\/g, '\\\\') : '';
        const safeFilename = escapeHtml(doc.filename);
        const safeDate = escapeHtml(doc.date);
        const safeStatus = escapeHtml(doc.status);
        
        const item = document.createElement('div');
        item.className = 'doc-item';
        item.innerHTML = `
            <div class="doc-icon">📄</div>
            <div class="doc-info">
                <div class="doc-name">${safeFilename}</div>
                <div class="doc-meta">${safeDate}</div>
            </div>
            <div class="doc-status">
                <span class="badge ${doc.status === 'Completed' ? 'badge-blue' : 'badge-pink'}">${safeStatus}</span>
            </div>
            <div class="doc-actions">
                <button class="icon-btn" onclick="openFile('${safePath}')" title="Open File">📂</button>
            </div>
        `;
        list.appendChild(item);
    });
}

function renderDocumentsTable(history) {
    const tbody = document.getElementById('allDocsTableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    // Global click listener to close context menu
    document.addEventListener('click', () => {
        const menu = document.getElementById('contextMenu');
        if (menu) menu.style.display = 'none';
    });
    
    history.forEach(doc => {
        const safePath = doc.path ? doc.path.replace(/\\/g, '\\\\') : '';
        const safeFilename = escapeHtml(doc.filename);
        const styleName = escapeHtml(doc.style || '-');
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="checkbox"></td>
            <td title="${safeFilename}">${safeFilename}</td>
            <td><span class="badge badge-blue" style="font-size: 11px;">${styleName}</span></td>
            <td style="text-align: center;">
                <button class="icon-btn" onclick="openFile('${safePath}')" title="Open Document">📂</button>
            </td>
        `;

        // Context Menu Handler
        tr.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            showContextMenu(e, doc);
        });

        tbody.appendChild(tr);
    });
}

function timeAgo(dateStr) {
    if (!dateStr || dateStr === '-') return 'Unknown';
    // Assuming dateStr is "YYYY-MM-DD HH:MM"
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    
    const now = new Date();
    const seconds = Math.floor((now - date) / 1000);
    
    let interval = Math.floor(seconds / 31536000);
    if (interval >= 1) return interval + " year" + (interval === 1 ? "" : "s") + " ago";
    
    interval = Math.floor(seconds / 2592000);
    if (interval >= 1) return interval + " month" + (interval === 1 ? "" : "s") + " ago";
    
    interval = Math.floor(seconds / 86400);
    if (interval >= 1) return interval + " day" + (interval === 1 ? "" : "s") + " ago";
    
    interval = Math.floor(seconds / 3600);
    if (interval >= 1) return interval + " hour" + (interval === 1 ? "" : "s") + " ago";
    
    interval = Math.floor(seconds / 60);
    if (interval >= 1) return interval + " minute" + (interval === 1 ? "" : "s") + " ago";
    
    return "just now";
}

function showContextMenu(e, doc) {
    const menu = document.getElementById('contextMenu');
    if (!menu) return;

    const safeFilename = escapeHtml(doc.filename);
    const timeAgoStr = escapeHtml(timeAgo(doc.date));
    const sizeStr = escapeHtml(doc.size || '-');
    const safeStatus = escapeHtml(doc.status);
    
    // Format date for tooltip: Jan 6, 2026 at 3:12 PM
    let tooltipDate = doc.date || 'Unknown';
    const dateObj = new Date(doc.date);
    if (!isNaN(dateObj.getTime())) {
        tooltipDate = dateObj.toLocaleDateString('en-US', { 
            month: 'short', 
            day: 'numeric', 
            year: 'numeric' 
        }) + ' at ' + dateObj.toLocaleTimeString('en-US', { 
            hour: 'numeric', 
            minute: '2-digit'
        });
    }
    const safeTooltipDate = escapeHtml(tooltipDate);

    menu.innerHTML = `
        <div class="context-menu-item" style="cursor: default; opacity: 0.7; font-size: 11px;">
            ${safeFilename}
        </div>
        <div class="context-menu-item separator"></div>
        <div class="context-menu-item">
            <span class="context-menu-label">Date</span>
            <span class="context-menu-val" style="position: relative;">
                ${timeAgoStr}
                <div class="context-date-tooltip">${safeTooltipDate}</div>
            </span>
        </div>
        <div class="context-menu-item">
            <span class="context-menu-label">Size</span>
            <span class="context-menu-val">${sizeStr}</span>
        </div>
        <div class="context-menu-item">
            <span class="context-menu-label">Status</span>
            <span class="context-menu-val">${safeStatus}</span>
        </div>
        <div class="context-menu-item separator"></div>
        <div class="context-menu-item" onclick="openFile('${doc.path ? doc.path.replace(/\\/g, '\\\\') : ''}')">
            <span>Open File</span>
            <span>📂</span>
        </div>
    `;
    
    // Position
    const x = e.pageX;
    const y = e.pageY;
    
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.style.display = 'block';
}

function openFile(path) {
    if (bridge) {
        bridge.open_file(path);
    }
}

function setupCustomSelects() {
    const selects = document.querySelectorAll('.form-select');
    
    selects.forEach(select => {
        // Create container
        const container = document.createElement('div');
        container.className = 'form-select-container';
        select.parentNode.insertBefore(container, select);
        
        // Create custom trigger
        const trigger = document.createElement('div');
        trigger.className = 'form-select-custom';
        trigger.tabIndex = 0; // Make focusable
        trigger.textContent = select.options[select.selectedIndex]?.text || 'Select...';
        container.appendChild(trigger);
        
        // Hide original select but keep it for data
        select.style.display = 'none';
        container.appendChild(select);
        
        // Create options list
        const optionsList = document.createElement('div');
        optionsList.className = 'custom-options';
        
        Array.from(select.options).forEach((opt, index) => {
            const customOpt = document.createElement('div');
            customOpt.className = 'custom-option';
            if (index === select.selectedIndex) customOpt.classList.add('selected');
            customOpt.textContent = opt.text;
            customOpt.addEventListener('click', (e) => {
                e.stopPropagation();
                select.selectedIndex = index;
                trigger.textContent = opt.text;
                
                // Update selected class
                optionsList.querySelectorAll('.custom-option').forEach(o => o.classList.remove('selected'));
                customOpt.classList.add('selected');
                
                container.classList.remove('open');
                
                // Trigger change event
                select.dispatchEvent(new Event('change'));
            });
            optionsList.appendChild(customOpt);
        });
        
        container.appendChild(optionsList);
        
        // Toggle open
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            // Close other selects
            document.querySelectorAll('.form-select-container').forEach(c => {
                if (c !== container) c.classList.remove('open');
            });
            container.classList.toggle('open');
        });
    });
    
    // Close on outside click
    document.addEventListener('click', () => {
        document.querySelectorAll('.form-select-container').forEach(c => c.classList.remove('open'));
    });
}

function syncAllCustomSelects() {
    document.querySelectorAll('.form-select').forEach(select => {
        const container = select.closest('.form-select-container');
        if (!container) return;
        
        const trigger = container.querySelector('.form-select-custom');
        const optionsList = container.querySelector('.custom-options');
        
        if (trigger) {
            trigger.textContent = select.options[select.selectedIndex]?.text || 'Select...';
        }
        
        if (optionsList) {
            const options = optionsList.querySelectorAll('.custom-option');
            options.forEach((opt, idx) => {
                if (idx === select.selectedIndex) {
                    opt.classList.add('selected');
                } else {
                    opt.classList.remove('selected');
                }
            });
        }
    });
}

// Chat Logic
// Chat Logic
function setupChat() {
    const inputs = ['chatInputMain', 'chatInputInterface'];
    
    inputs.forEach(id => {
        const input = document.getElementById(id);
        if (!input) return;

        // Auto-resize
        input.addEventListener('input', function() {
            this.style.height = 'auto';
            if (id === 'chatInputInterface') {
                 // Smaller resize logic for interface mode if needed, or keep generic
                 this.style.height = (this.scrollHeight) + 'px';
                 if (this.value === '') this.style.height = 'auto'; // Reset
            } else {
                 this.style.height = (this.scrollHeight) + 'px';
            }
        });

        // Send on Enter
        input.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage(id);
            }
        });
    });

    // Buttons
    const btnMain = document.getElementById('sendChatBtnMain');
    const btnInterface = document.getElementById('sendChatBtnInterface');

    if (btnMain) btnMain.addEventListener('click', () => sendMessage('chatInputMain'));
    if (btnInterface) btnInterface.addEventListener('click', () => sendMessage('chatInputInterface'));

    // Quick Actions
    const quickActions = document.querySelectorAll('.quick-action-pill');
    quickActions.forEach(btn => {
        btn.addEventListener('click', () => {
             const question = btn.getAttribute('data-question');
             if (question) {
                 sendQuickMessage(question);
             }
        });
    });
}

function sendQuickMessage(text) {
     // Transition to interface automatically
    const welcomeScreen = document.getElementById('chatWelcomeScreen');
    const chatInterface = document.getElementById('chatInterface');
    
    if (welcomeScreen.style.display !== 'none') {
        welcomeScreen.style.display = 'none';
        chatInterface.style.display = 'flex';
    }
    
    // Add User Message
    appendMessage('user', text);
    
    // Call AI
    if (bridge) {
        bridge.ask_ai(text);
        showTypingIndicator(); // Show "Formatly is typing..."
    } else {
        // Fallback for dev without bridge
        setTimeout(() => appendMessage('system', 'Backend bridge not connected. Message: ' + text), 1000);
    }
}

function sendMessage(inputId) {
    const input = document.getElementById(inputId);
    if (!input) return;
    
    const message = input.value.trim();
    if (!message) return;

    // Transition to Interface Mode if needed
    const welcomeScreen = document.getElementById('chatWelcomeScreen');
    const chatInterface = document.getElementById('chatInterface');
    
    if (welcomeScreen.style.display !== 'none') {
        welcomeScreen.style.display = 'none';
        chatInterface.style.display = 'flex';
        // Focus the other input for continuity if they want to type more
        document.getElementById('chatInputInterface').focus();
    }

    // Add User Message
    appendMessage('user', message);
    input.value = '';
    input.style.height = 'auto';
    
    // Clear the other input too just in case
    if (inputId === 'chatInputMain') {
         document.getElementById('chatInputInterface').value = '';
    }

    // Call Backend
    if (bridge) {
        bridge.ask_ai(message);
        showTypingIndicator();
    } else {
         setTimeout(() => appendMessage('system', "I'm Formatly AI. (Bridge Disconnected)"), 600);
    }
}

// Typing Indicator
function showTypingIndicator() {
     const messagesContainer = document.getElementById('chatMessages');
     const existing = document.getElementById('typingIndicator');
     if (existing) return;
     
     const msgDiv = document.createElement('div');
     msgDiv.className = `message system`;
     msgDiv.id = 'typingIndicator';
     
     msgDiv.innerHTML = `
        <div class="message-content">
            <div class="message-avatar" style="background: linear-gradient(135deg, #4285f4, #9b72cb, #ea4335); color: white;">✨</div>
            <div class="message-text">Thinking...</div>
        </div>
    `;
    messagesContainer.appendChild(msgDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

function hideTypingIndicator() {
    const el = document.getElementById('typingIndicator');
    if (el) el.remove();
}

// Global Stream State
let currentMessageEl = null;
let fullStreamedText = "";

// AI Signal Handlers
window.startAIStream = function() {
    hideTypingIndicator(); // Remove "Thinking..."
    fullStreamedText = "";
    
    // Create the dummy message container immediately with empty text
    // We use a specific ID or just grabbing the last one?
    // Let's modify appendMessage to return the text element if possible, 
    // but for now relying on DOM order is safe enough in single-threaded JS.
    appendMessage('system', ''); 
    
    const msgs = document.querySelectorAll('.message.system .message-text');
    if (msgs.length > 0) {
        currentMessageEl = msgs[msgs.length - 1];
    }
};

window.handleAIChunk = function(chunk) {
    if (!currentMessageEl) return;
    
    fullStreamedText += chunk;
    currentMessageEl.innerHTML = parseMarkdown(fullStreamedText);
    
    // Auto scroll content
    const container = document.getElementById('chatMessages');
    // Only scroll if we were already near bottom to avoid annoying user if reading up history? 
    // For now, always scroll (chat GPT style usually locks to bottom generating)
    container.scrollTop = container.scrollHeight;
};

window.handleAIResponse = function(success, message) {
    hideTypingIndicator(); // Ensure hidden
    
    if (!success) {
        // Error case
        appendMessage('system', `❌ Error: ${message}`);
    } else {
        // Finished successfully
        if (currentMessageEl) {
             // Final parse to ensure everything is perfect (startAIStream/handleAIChunk handled live)
             // Use the 'message' from the signal which is the full text
             currentMessageEl.innerHTML = parseMarkdown(message);
        }
    }
    currentMessageEl = null; // Reset
};

function parseMarkdown(text) {
    if (!text) return '';

    // 1. Escape HTML (Basic)
    let html = text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    // 2. Code Blocks
    html = html.replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>');

    // 3. Headers
    html = html.replace(/^### (.*$)/gm, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gm, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gm, '<h1>$1</h1>');

    // 4. Bold & Italic
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
    
    // 5. Lists (Unordered)
    // Wrap items in <li>
    html = html.replace(/^\s*[\-\*]\s+(.*)$/gm, '<li>$1</li>');
    // Wrap sequences of <li> in <ul>
    html = html.replace(/(<li>.*<\/li>)/gm, '<ul>$1</ul>');
    html = html.replace(/<\/ul>\s*<ul>/g, ''); 

    // 6. Paragraphs / Line Breaks
    html = html.replace(/\n/g, '<br>');
    html = html.replace(/(<\/h[1-6]>|<\/ul>|<\/pre>)\s*<br>/g, '$1');
    
    return html;
}

function appendMessage(role, text) {
    const messagesContainer = document.getElementById('chatMessages');
    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${role}`;
    
    // Avatar Logic
    let avatarHtml = '';
    if (role === 'system') {
        // Formatly Logo
        avatarHtml = `
        <div class="message-avatar" style="background: none; padding: 0;">
             <img src="../assets/logo.png" alt="🤖" style="width: 24px; height: 24px; object-fit: contain;">
        </div>`;
    } else {
        // User avatar hidden by CSS, but we keep structure
        avatarHtml = `<div class="message-avatar"></div>`;
    }
    
    // Parse Text
    const contentHtml = role === 'system' ? parseMarkdown(text) : text.replace(/\n/g, '<br>');

    msgDiv.innerHTML = `
        <div class="message-content">
            ${avatarHtml}
            <div style="flex: 1; min-width: 0;">
                <div class="message-text">${contentHtml}</div>
            </div>
        </div>
    `;
    
    messagesContainer.appendChild(msgDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}



