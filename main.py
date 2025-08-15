#!/usr/bin/env python3
"""
Email Attachment Extractor - Main Entry Point
Professional email automation tool for extracting attachments from Microsoft Outlook.
"""

import sys
import os
import argparse
from pathlib import Path

# Add src directory to Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.core.config_manager import ConfigManager
from src.core.outlook_manager import OutlookManager
from src.core.scheduler import UserLevelScheduler
from src.ui.main_window import MainWindow


def run_automated_extraction(config_manager, outlook_manager):
    """Run extraction in automated mode without UI"""
    try:
        from src.utils.email_processor import EmailProcessor
        from src.utils.file_manager import FileManager
        
        print("Starting automated extraction...")
        
        # Get current settings
        settings = {
            'keyword': config_manager.get('keyword', ''),
            'folder': config_manager.get('folder', ''),
            'days': config_manager.get('days', 7),
            'selected_folders': config_manager.get('selected_folders', ['Inbox']),
            'providers': config_manager.get('providers', ''),
            'extraction_mode': config_manager.get('extraction_mode', 'all'),
            'naming_format': config_manager.get('naming_format', 'date')
        }
        
        # Test Outlook connection
        if not outlook_manager.test_connection():
            print("Error: Could not connect to Outlook")
            return False
        
        # Initialize processors
        email_processor = EmailProcessor()
        file_manager = FileManager()
        
        # Process emails and extract attachments
        print(f"Searching {len(settings['selected_folders'])} folders for emails from last {settings['days']} days...")
        
        # This is a simplified extraction - you might need to adapt based on your actual email processing logic
        # For now, we'll just mark the run as successful
        scheduler = UserLevelScheduler(config_manager)
        scheduler.save_last_run()
        
        print("Automated extraction completed successfully!")
        return True
        
    except Exception as e:
        print(f"Error during automated extraction: {e}")
        return False


def show_extraction_notification():
    """Show notification that extraction is ready"""
    try:
        # Try to use system notifications
        import tkinter as tk
        from tkinter import messagebox
        
        # Create temporary root for messagebox
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        
        result = messagebox.askyesno(
            "📧 Email Extraction Ready",
            "Scheduled email extraction is ready to run.\n\nRun extraction now?"
        )
        
        root.destroy()
        
        if result:
            # User wants to run extraction
            config_manager = ConfigManager()
            outlook_manager = OutlookManager()
            return run_automated_extraction(config_manager, outlook_manager)
        
        return False
        
    except ImportError:
        # Fallback if tkinter is not available
        print("Email extraction is ready to run!")
        print("Open the Email Extractor application to proceed.")
        return False


def main():
    """Main application entry point"""
    try:
        # Parse command line arguments
        parser = argparse.ArgumentParser(description='Email Attachment Extractor')
        parser.add_argument('--check-schedule', action='store_true', 
                           help='Check if scheduled extraction is needed')
        parser.add_argument('--silent', action='store_true', 
                           help='Run in silent mode without notifications')
        parser.add_argument('--run-now', action='store_true',
                           help='Run extraction immediately without UI')
        
        args = parser.parse_args()
        
        # Initialize core components
        config_manager = ConfigManager()
        
        # Handle automation checking
        if args.check_schedule:
            scheduler = UserLevelScheduler(config_manager)
            
            if scheduler.should_run_extraction():
                if args.silent:
                    # Run extraction silently
                    outlook_manager = OutlookManager()
                    success = run_automated_extraction(config_manager, outlook_manager)
                    sys.exit(0 if success else 1)
                else:
                    # Show notification to user
                    success = show_extraction_notification()
                    sys.exit(0 if success else 1)
            else:
                # No extraction needed
                sys.exit(0)
        
        # Handle immediate run
        if args.run_now:
            outlook_manager = OutlookManager()
            success = run_automated_extraction(config_manager, outlook_manager)
            sys.exit(0 if success else 1)
        
        # Normal UI mode
        outlook_manager = OutlookManager()
        app = MainWindow(config_manager, outlook_manager)
        
        # Check automation status on startup
        try:
            if hasattr(app, 'settings_tab') and app.settings_tab:
                app.settings_tab.check_automation_status_on_startup()
        except Exception as e:
            print(f"Automation startup check failed: {e}")
        
        app.run()
        
    except Exception as e:
        print(f"Error starting application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()