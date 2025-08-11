"""
Build script for creating standalone executable
"""
import sys
from pathlib import Path
import PyInstaller.__main__

def build_executable():
    """Build standalone executable"""

    # Project root
    project_root = Path(__file__).parent

    # PyInstaller arguments
    args = [
        str(project_root/"src"/"main.py"),   # Main script
        "--onefile",                         # Create a standalone .exe file
        "--windowed",                        # No console window
        "--name=CV-Converter",               # Executable name
        f"--distpath={project_root/"dist"}", # Output directory (.exe is stored here)
        f"--workpath={project_root/"build"}",# Work directory (temporary files while building .exe is stored here)
        f"--paths={project_root/"src"}",     # For correct referencing of local modules

        # Include data files in .exe
        f"--add-data={project_root/"templates"};templates",  # Template files stored to .exe
        f"--add-data={project_root/"assets"};assets",        # Asset files stored to .exe 
        f"--add-data={project_root/"config"};config",        # config files stored to .exe

        # Icon (if available)
        # f"--icon={project_root / 'assets' / 'icon.ico'}",

        # Hidden imports
        "--hidden-import=docx",
        "--hidden-import=PIL",

        # Exclude unnecessary modules
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=pandas",
    ]

    print("Building executable...")
    PyInstaller.__main__.run(args)
    print("Build completed")

if __name__ == "__main__":
    build_executable()