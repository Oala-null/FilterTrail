"""
FilterTrail - Monitor Excel filter operations and visualize filter flow
Main launcher script
"""

import sys
import os

def main():
    """
    Launch the FilterTrail application
    """
    # Import necessary modules only when needed
    try:
        from main import FilterTrailApp, QApplication
        
        # Create and start the application
        app = QApplication(sys.argv)
        window = FilterTrailApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error launching FilterTrail: {e}")
        input("Press Enter to exit...")
        sys.exit(1)

def build_executable():
    """
    Build standalone executable for FilterTrail by calling optimized_build.py
    """
    try:
        print("Launching build process...")
        
        # Try to import and run the optimized build script
        try:
            from optimized_build import create_empty_data_file, clean_build_artifacts, build_executable
            create_empty_data_file()
            clean_build_artifacts()
            build_executable()
        except ImportError:
            # If we can't directly import, run the script as a subprocess
            import subprocess
            subprocess.run([sys.executable, "optimized_build.py"])
        
        print("Build process completed.")
    except Exception as e:
        print(f"Error during build process: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--build":
        build_executable()
    else:
        main()