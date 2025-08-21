"""
Main Tab UI Component
Contains the primary extraction interface and logging.
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime
import threading
from typing import Optional


class MainTab:
    """Main tab for email extraction operations"""
    
    def __init__(self, parent, config_manager, outlook_manager, colors):
        """Initialize main tab"""
        self.parent = parent
        self.config_manager = config_manager
        self.outlook_manager = outlook_manager
        self.colors = colors
        self.settings_tab = None
        
        # Create main frame
        self.frame = tk.Frame(parent, bg=colors['bg'])
        
        # Setup UI
        self.setup_ui()
        
        # Set outlook manager callback
        self.outlook_manager.progress_callback = self.log_message
    
    def setup_ui(self):
        """Setup the main tab UI"""
        # Main container with padding
        main_container = tk.Frame(self.frame, bg=self.colors['bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_container, text="📧 Email Attachment Extractor", style='Heading.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Quick Settings Card
        self.create_quick_settings(main_container)
        
        # Action Buttons
        self.create_action_buttons(main_container)
        
        # Progress Section
        self.create_progress_section(main_container)
        
        # Activity Log
        self.create_log_section(main_container)
    
    def create_quick_settings(self, parent):
        """Create quick settings card"""
        settings_card = ttk.Frame(parent, style='Card.TFrame')
        settings_card.pack(fill=tk.X, pady=(0, 15))
        
        settings_inner = tk.Frame(settings_card, bg=self.colors['card'])
        settings_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(settings_inner, text="🚀 Quick Settings", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Keyword input
        keyword_frame = tk.Frame(settings_inner, bg=self.colors['card'])
        keyword_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(keyword_frame, text="Search keyword:", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT)
        
        self.keyword_entry = tk.Entry(keyword_frame, width=20, font=('Segoe UI', 10))
        self.keyword_entry.pack(side=tk.LEFT, padx=(10, 0))
        self.keyword_entry.insert(0, self.config_manager.get('keyword', 'report'))
        
        # Auto-save keyword when changed
        def on_keyword_change(*args):
            self.config_manager.set('keyword', self.keyword_entry.get())
        
        self.keyword_entry.bind('<KeyRelease>', on_keyword_change)
        
        # Folder selection
        folder_frame = tk.Frame(settings_inner, bg=self.colors['card'])
        folder_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(folder_frame, text="Save to:", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT)
        
        self.folder_path = tk.StringVar(value=self.config_manager.get('folder', ''))
        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path, width=40, font=('Segoe UI', 10))
        folder_entry.pack(side=tk.LEFT, padx=(10, 5))
        
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side=tk.LEFT)
    
    def create_action_buttons(self, parent):
        """Create action buttons"""
        action_frame = tk.Frame(parent, bg=self.colors['bg'])
        action_frame.pack(fill=tk.X, pady=15)
        
        self.extract_btn = ttk.Button(action_frame, text="🚀 Extract Reports Now", 
                                     command=self.run_extraction, style='Primary.TButton')
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(action_frame, text="🔧 Test Connection", 
                  command=self.test_connection).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(action_frame, text="📁 Discover Folders", 
                  command=self.discover_folders).pack(side=tk.LEFT)
    
    def create_progress_section(self, parent):
        """Create progress section"""
        self.progress_label = tk.Label(parent, text="Ready to extract reports...", 
                                      font=('Segoe UI', 10), bg=self.colors['bg'],
                                      fg=self.colors['text'])
        self.progress_label.pack(pady=(10, 5))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(parent, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
    
    def create_log_section(self, parent):
        """Create activity log section"""
        log_frame = tk.Frame(parent, bg=self.colors['bg'])
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        tk.Label(log_frame, text="📋 Activity Log", font=('Segoe UI', 12, 'bold'),
                bg=self.colors['bg'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=15, 
                                                 font=('Consolas', 9), wrap=tk.WORD,
                                                 bg=self.colors['card'], fg=self.colors['text'])
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure log colors
        log_colors = {
            "INFO": self.colors['text'],
            "SUCCESS": self.colors['success'],
            "WARNING": self.colors['warning'],
            "ERROR": self.colors['danger']
        }
        
        for level, color in log_colors.items():
            self.log_text.tag_configure(level, foreground=color)
        
        # Welcome message
        self.log_message("Welcome to Email Attachment Extractor!", "SUCCESS")
        self.log_message("Configure your settings in the Settings tab, then click 'Extract Reports Now'", "INFO")
    
    def set_settings_tab(self, settings_tab):
        """Set reference to settings tab"""
        self.settings_tab = settings_tab
    
    def browse_folder(self):
        """Browse for folder selection"""
        folder = filedialog.askdirectory(initialdir=self.folder_path.get())
        if folder:
            self.folder_path.set(folder)
            self.config_manager.set('folder', folder)
    
    def test_connection(self):
        """Test Outlook connection"""
        def test_in_thread():
            self.progress_bar.start()
            self.extract_btn.config(state='disabled')
            
            try:
                success = self.outlook_manager.test_connection()
                if success and hasattr(self.settings_tab, 'refresh_folder_list'):
                    # Also discover folders for settings tab
                    self.root.after(0, lambda: self.settings_tab.refresh_folder_list())
            finally:
                self.progress_bar.stop()
                self.extract_btn.config(state='normal')
        
        thread = threading.Thread(target=test_in_thread, daemon=True)
        thread.start()
    
    def discover_folders(self):
        """Discover Outlook folders"""
        def discover_callback(folders):
            self.log_message(f"Discovered {len(folders)} folders", "SUCCESS")
            if hasattr(self.settings_tab, 'update_folder_checkboxes'):
                self.settings_tab.update_folder_checkboxes(folders)
        
        self.progress_bar.start()
        self.outlook_manager.discover_folders_async(discover_callback)
        self.progress_bar.stop()
    
    def run_extraction(self):
        """Run email extraction"""
        def extraction_in_thread():
            try:
                self.progress_bar.start()
                self.extract_btn.config(state='disabled', text="⏳ Extracting...")
                
                # Get settings from UI directly (not from saved config)
                keyword = self.keyword_entry.get()
                save_folder = self.folder_path.get()
                
                if not keyword:
                    self.log_message("Please enter a search keyword", "ERROR")
                    return
                
                if not save_folder:
                    self.log_message("Please select a save folder", "ERROR")
                    return
                
                # Get current settings from UI (not saved config)
                if self.settings_tab:
                    settings = self.settings_tab.get_current_settings()
                    days = settings['days']
                    selected_folders = settings['selected_folders']
                    naming_format = settings['naming_format']
                    custom_suffix = settings['custom_suffix']
                    extraction_mode = settings['extraction_mode']
                    providers_text = settings['providers']
                    conversion_enabled = settings.get('conversion_enabled', False)
                    convert_to_format = settings.get('convert_all_to_format', '')
                    custom_format = settings.get('custom_format_extension', 'xlsx')
                else:
                    # Fallback to saved config if settings tab not available
                    days = self.config_manager.get('days', 7)
                    selected_folders = self.config_manager.get('selected_folders', ['Inbox'])
                    naming_format = self.config_manager.get('naming_format', 'date')
                    custom_suffix = self.config_manager.get('custom_suffix', '')
                    extraction_mode = self.config_manager.get('extraction_mode', 'all')
                    providers_text = self.config_manager.get('providers', '')
                    conversion_enabled = self.config_manager.get('conversion_enabled', False)
                    convert_to_format = self.config_manager.get('convert_all_to_format', '')
                    custom_format = self.config_manager.get('custom_format_extension', 'xlsx')
                
                self.log_message(f"Using settings: {days} days, {len(selected_folders)} folders, {extraction_mode} mode", "INFO")
                
                # Get messages
                messages = self.outlook_manager.get_messages_from_folders(
                    selected_folders, days, keyword, providers_text
                )
                
                if not messages:
                    self.log_message("No matching messages found", "WARNING")
                    return
                
                # Process messages
                self.log_message(f"Found {len(messages)} matching messages", "SUCCESS")
                
                # Use custom format if "custom" is selected
                final_format = custom_format if convert_to_format == 'custom' else convert_to_format
                
                saved_files = self.outlook_manager.save_attachments(
                    messages, save_folder, naming_format, custom_suffix, extraction_mode,
                    conversion_enabled, final_format
                )
                
                self.log_message(f"Extraction completed! Saved {len(saved_files)} files", "SUCCESS")
                
            except Exception as e:
                self.log_message(f"Extraction failed: {str(e)}", "ERROR")
            finally:
                self.progress_bar.stop()
                self.extract_btn.config(state='normal', text="🚀 Extract Reports Now")
        
        thread = threading.Thread(target=extraction_in_thread, daemon=True)
        thread.start()
    
    def log_message(self, message: str, level: str = "INFO"):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Insert message with color
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.log_text.see(tk.END)
        
        # Update progress label for important messages
        if level in ["SUCCESS", "ERROR", "WARNING"]:
            self.progress_label.config(text=message)
    
    def on_resize(self, event):
        """Handle resize events"""
        # Implement any resize-specific logic for main tab
        pass