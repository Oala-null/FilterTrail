"""
FilterTrail Optimized Build Script
This script creates a standalone executable for FilterTrail.
"""

import PyInstaller.__main__
import os
import json
import shutil
import sys

def create_empty_data_file():
    """Create empty data file for inclusion"""
    print("Creating empty filter data file...")
    with open("filter_data.json", "w") as f:
        json.dump([], f)
    print("  ✅ Created empty filter_data.json")

def clean_build_artifacts():
    """Clean any existing build artifacts"""
    print("Cleaning build artifacts...")
    folders_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = [f for f in os.listdir() if f.endswith(".spec")]
    
    # Clean folders
    for folder in folders_to_remove:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"  ✅ Removed {folder} directory")
            except Exception as e:
                print(f"  ❌ Error removing {folder}: {e}")
    
    # Clean files
    for file in files_to_remove:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"  ✅ Removed {file}")
            except Exception as e:
                print(f"  ❌ Error removing {file}: {e}")
    print("Cleanup complete.")

def build_executable():
    """Build lightweight executable for FilterTrail"""
    print("Building FilterTrail executable...")
    
    # Use OS-specific path separator
    path_separator = ';' if os.name == 'nt' else ':'
    data_path = f'--add-data=filter_data.json{path_separator}.'
    
    # Create resources folder if it doesn't exist
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources")
    if not os.path.exists(resources_dir):
        os.makedirs(resources_dir)
        print(f"  ✅ Created resources directory")
    
    # Check for logo files
    logo_path = os.path.join(resources_dir, "logo.png")
    favicon_path = os.path.join(resources_dir, "favicon.ico")
    
    # Create and execute PyInstaller command
    cmd = [
        'main.py',  # Main script
        '--name=FilterTrail',
        '--windowed',  # No console window
        '--onefile',   # Single executable file
        '--clean',     # Clean PyInstaller cache
        data_path,     # Include empty data file
    ]
    
    # Add resources folder
    resources_data_path = f'--add-data=resources{path_separator}resources'
    cmd.append(resources_data_path)
    
    # Add icon if available
    if os.path.exists(favicon_path):
        cmd.append(f'--icon={favicon_path}')
        print(f"  ✅ Using favicon.ico as application icon")
    
    # Add remaining options
    cmd.extend([
        '--hidden-import=win32com',
        '--hidden-import=win32com.client',
        '--hidden-import=pythoncom',
        '--hidden-import=PyQt5.QtWebEngineWidgets',
        # Exclude unnecessary libraries to reduce size
        '--exclude-module=matplotlib',
        '--exclude-module=scipy',
        '--exclude-module=pandas.plotting',
        '--exclude-module=holoviews',
        '--exclude-module=numpy.random',
        '--exclude-module=pydantic',
        '--exclude-module=IPython',
        '--exclude-module=PySide2',
        '--exclude-module=nltk',         # Exclude NLTK explicitly
        # Use UPX directory if available but don't use the ambiguous --upx flag
        '--upx-dir=.'
    ])
    
    # Run PyInstaller with the command
    PyInstaller.__main__.run(cmd)
    
    # Verify build was successful
    exe_path = os.path.abspath(os.path.join('dist', 'FilterTrail.exe'))
    if os.path.exists(exe_path):
        print(f"\n✅ Build completed successfully!")
        print(f"Executable created at: {exe_path}")
        print("Note: This executable uses CDN for Plotly to reduce file size.")
        print("\nTo run the application:")
        print("1. Make sure Excel is open with a worksheet")
        print("2. Double-click FilterTrail.exe")
        print("3. Click 'Start Monitoring' and apply filters in Excel")
    else:
        print(f"\n❌ Build failed! Executable not found at: {exe_path}")
        print("Please check the error messages above.")

if __name__ == "__main__":
    print("=" * 60)
    print("FilterTrail Build Process")
    print("=" * 60)
    
    # Prepare for build
    create_empty_data_file()
    clean_build_artifacts()
    
    # Build executable
    build_executable()
    
    # Done
    print("\n" + "=" * 60)
    if sys.platform == 'win32':
        input("Press Enter to exit...")