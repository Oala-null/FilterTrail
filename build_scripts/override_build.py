"""
FilterTrail Override Build Script - Uses hook overrides to solve Anaconda environment issues
"""

import PyInstaller.__main__
import os
import json
import shutil
import sys
import tempfile

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

def create_hook_override():
    """Create a hook override for pydantic"""
    hooks_dir = "hooks"
    
    if not os.path.exists(hooks_dir):
        os.makedirs(hooks_dir)
        print(f"  ✅ Created hooks directory")
    
    # Create an empty hook for pydantic that overrides the default one
    hook_file = os.path.join(hooks_dir, "hook-pydantic.py")
    with open(hook_file, "w") as f:
        f.write("""
# Empty hook for pydantic to override the default
# This prevents PyInstaller from trying to analyze pydantic

# Define empty lists to prevent issues
hiddenimports = []
excludedimports = []
        """)
    print(f"  ✅ Created pydantic hook override")
    
    return os.path.abspath(hooks_dir)

def build_executable():
    """Build executable with hook overrides"""
    print("Building FilterTrail Override Version...")
    
    # Create hook overrides
    hooks_dir = create_hook_override()
    
    # Use OS-specific path separator
    path_separator = ';' if os.name == 'nt' else ':'
    data_path = f'--add-data=filter_data.json{path_separator}.'
    
    # Create command with hook overrides
    cmd = [
        'main.py',  # Main script
        '--name=FilterTrail_Override',
        '--windowed',                # No console window
        '--onedir',                  # Directory based (faster startup)
        '--clean',                   # Clean PyInstaller cache
        data_path,                   # Include empty data file
        f'--additional-hooks-dir={hooks_dir}',  # Use our hooks
        '--hidden-import=win32com.client',
        '--hidden-import=pythoncom',
        '--hidden-import=PyQt5.QtWebEngineWidgets',
        '--exclude-module=pydantic',  # Explicitly exclude problematic module
        '--exclude-module=nltk',      # Also exclude nltk
        '--noupx',
    ]
    
    # Run PyInstaller with the command
    PyInstaller.__main__.run(cmd)
    
    # Verify build was successful
    exe_path = os.path.abspath(os.path.join('dist', 'FilterTrail_Override'))
    
    if os.path.exists(exe_path):
        print(f"\n✅ Override build completed successfully!")
        print(f"Application folder created at: {exe_path}")
        print("Note: This version uses hook overrides to solve Anaconda environment issues.")
        print("\nTo run the application:")
        print(f"1. Navigate to: {exe_path}")
        print("2. Run the FilterTrail_Override executable")
        print("3. Make sure Excel is open with a worksheet")
    else:
        print(f"\n❌ Build failed! Application folder not found at: {exe_path}")
        print("Please check the error messages above.")

if __name__ == "__main__":
    print("=" * 60)
    print("FilterTrail Override Build Process")
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