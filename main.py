#!/usr/bin/env python3
"""
Email Attachment Extractor - Main Entry Point
Professional email automation tool for extracting attachments from Microsoft Outlook.
"""

import sys
import os
from pathlib import Path

# Add src directory to Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.core.config_manager import ConfigManager
from src.core.outlook_manager import OutlookManager
from src.ui.main_window import MainWindow


def main():
    """Main application entry point"""
    try:
        # Initialize core components
        config_manager = ConfigManager()
        outlook_manager = OutlookManager()
        
        # Create and run main window
        app = MainWindow(config_manager, outlook_manager)
        app.run()
        
    except Exception as e:
        print(f"Error starting application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()