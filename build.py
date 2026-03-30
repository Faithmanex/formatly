import PyInstaller.__main__
import os
import shutil

def clean_build():
    """Removes previous build artifacts."""
    for folder in ['build', 'dist']:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"Removed {folder} directory.")
    
    spec_file = "Formatly.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"Removed {spec_file} file.")

def build():
    clean_build()
    
    # Define the PyInstaller arguments
    pyinstaller_args = [
        'desktop/GUI_pyside.py',                 # Entry point
        '--name=Formatly',                       # Name of the executable
        '--noconsole',                           # Suppress the console window (for GUI apps)
        '--onefile',                             # Bundle everything into a single executable
        '--clean',                               # Clean PyInstaller cache
        
        # Add data files: source;destination
        # Note: Windows uses ; as separator
        '--add-data=desktop/interface;interface',
        '--add-data=desktop/logo;logo',
        
        # Icon
        '--icon=desktop/interface/favicon.ico',
        
        # Internal imports that might be missed
        '--hidden-import=win32com.client',
        '--hidden-import=engineio.async_drivers.threading', 
        
        # Collect all PySide6 components to be safe
        '--collect-all=PySide6',
    ]

    print("Starting PyInstaller build...")
    PyInstaller.__main__.run(pyinstaller_args)
    print("Build complete. Executable is in the 'dist' folder.")

if __name__ == "__main__":
    build()
