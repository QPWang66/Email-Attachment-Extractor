"""
Main Window UI Component
Contains the primary application window and tab management.
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, Any

from .main_tab import MainTab
from .settings_tab import SettingsTab
from .styles import StyleManager


class MainWindow:
    """Main application window with tabbed interface"""
    
    def __init__(self, config_manager, outlook_manager):
        """Initialize main window"""
        self.config_manager = config_manager
        self.outlook_manager = outlook_manager
        
        # Create root window
        self.root = tk.Tk()
        self.root.title("📧 Email Attachment Extractor")
        
        # Get saved geometry or use default
        geometry = self.config_manager.get('window_geometry', '1000x750')
        self.root.geometry(geometry)
        self.root.minsize(800, 600)
        
        # Initialize style manager
        self.style_manager = StyleManager()
        self.colors = self.style_manager.get_colors()
        
        self.root.configure(bg=self.colors['bg'])
        
        # Setup styles
        self.style_manager.setup_styles()
        
        # Create UI components
        self.setup_ui()
        
        # Bind window events
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind("<Configure>", self.on_window_resize)
    
    def setup_ui(self):
        """Setup the main UI components"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create main tab
        self.main_tab = MainTab(
            self.notebook, 
            self.config_manager, 
            self.outlook_manager,
            self.colors
        )
        self.notebook.add(self.main_tab.frame, text="📊 Extract Reports")
        
        # Create settings tab
        self.settings_tab = SettingsTab(
            self.notebook,
            self.config_manager,
            self.outlook_manager, 
            self.colors
        )
        self.notebook.add(self.settings_tab.frame, text="⚙️ Settings")
        
        # Connect tabs for communication
        self.main_tab.set_settings_tab(self.settings_tab)
        self.settings_tab.set_main_tab(self.main_tab)
    
    def on_closing(self):
        """Handle window closing"""
        # Save current window geometry
        geometry = self.root.geometry()
        self.config_manager.set('window_geometry', geometry)
        self.config_manager.save_config()
        
        # Close the application
        self.root.destroy()
    
    def on_window_resize(self, event):
        """Handle window resize events"""
        if event.widget == self.root:
            # Notify tabs about resize
            if hasattr(self, 'main_tab'):
                self.main_tab.on_resize(event)
            if hasattr(self, 'settings_tab'):
                self.settings_tab.on_resize(event)
    
    def show_message(self, message: str, level: str = "INFO"):
        """Show message in the main tab log"""
        if hasattr(self, 'main_tab'):
            self.main_tab.log_message(message, level)
    
    def run(self):
        """Start the application"""
        # Load saved settings
        if hasattr(self, 'settings_tab'):
            self.settings_tab.load_saved_settings()
        
        # Auto-run if enabled
        if self.config_manager.get('auto_run', False):
            self.root.after(2000, lambda: self.main_tab.run_extraction())
        
        # Start the main loop
        self.root.mainloop()
    
    def get_root(self):
        """Get the root window"""
        return self.root