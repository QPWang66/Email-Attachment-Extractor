"""
Setup script for Email Attachment Extractor
Creates an executable using cx_Freeze or PyInstaller
"""

import sys
from pathlib import Path

try:
    from cx_Freeze import setup, Executable
    USE_CX_FREEZE = True
except ImportError:
    USE_CX_FREEZE = False
    print("cx_Freeze not found. Install with: pip install cx_Freeze")

# Build options
build_exe_options = {
    "packages": [
        "tkinter", 
        "win32com.client", 
        "pythoncom",
        "src.core",
        "src.ui", 
        "src.utils"
    ],
    "include_files": [
        ("README.md", "README.md"),
    ],
    "excludes": ["test", "unittest"],
    "include_msvcrt": True,
}

# Create executable
executable = Executable(
    script="main.py",
    base="Win32GUI" if sys.platform == "win32" else None,
    icon=None,  # Add icon file path if you have one
    target_name="EmailAttachmentExtractor.exe",
)

if USE_CX_FREEZE:
    setup(
        name="Email Attachment Extractor",
        version="2.0.0",
        description="Professional email attachment extraction tool for Microsoft Outlook",
        author="Open Source Contributors",
        options={"build_exe": build_exe_options},
        executables=[executable],
    )
else:
    print("Setup script requires cx_Freeze to build executable")
    print("Install with: pip install cx_Freeze")
    print("Then run: python setup.py build")