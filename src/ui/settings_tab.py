"""
Settings Tab UI Component
Contains all configuration options and folder selection.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from typing import Dict, Any, Optional
from datetime import datetime


class SettingsTab:
    """Settings tab for configuration management"""
    
    def __init__(self, parent, config_manager, outlook_manager, colors):
        """Initialize settings tab"""
        self.parent = parent
        self.config_manager = config_manager
        self.outlook_manager = outlook_manager
        self.colors = colors
        self.main_tab = None
        
        # Create main frame
        self.frame = tk.Frame(parent, bg=colors['bg'])
        
        # Initialize variables
        self.folder_checkboxes = {}
        self.days_var = tk.IntVar(value=self.config_manager.get('days', 7))
        self.auto_run_var = tk.BooleanVar(value=self.config_manager.get('auto_run', False))
        self.naming_format = tk.StringVar(value=self.config_manager.get('naming_format', 'date'))
        self.custom_suffix_var = tk.StringVar(value=self.config_manager.get('custom_suffix', ''))
        self.extraction_mode = tk.StringVar(value=self.config_manager.get('extraction_mode', 'all'))
        
        # Setup UI
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the settings tab UI"""
        # Create scrollable settings container
        self.create_scrollable_container()
        
        # Email Settings Card
        self.create_email_settings()
        
        # Folder Selection Card
        self.create_folder_selection()
        
        # File Naming Card
        self.create_file_naming()
        
        # Provider Settings Card
        self.create_provider_settings()
        
        # Auto-run option
        self.create_auto_run_option()
        
        # Save button
        self.create_save_button()
        
        # Initialize visibility
        self.toggle_naming_format()
    
    def create_scrollable_container(self):
        """Create scrollable container for settings"""
        # Create canvas and scrollbar
        self.settings_canvas = tk.Canvas(self.frame, bg=self.colors['bg'], highlightthickness=0)
        settings_scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.settings_canvas.yview)
        self.settings_container = tk.Frame(self.settings_canvas, bg=self.colors['bg'])
        
        self.settings_container.bind(
            "<Configure>",
            lambda e: self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all"))
        )
        
        self.settings_canvas.create_window((0, 0), window=self.settings_container, anchor="nw")
        self.settings_canvas.configure(yscrollcommand=settings_scrollbar.set)
        
        # Pack canvas and scrollbar
        self.settings_canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=20)
        settings_scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=20)
        
        # Mouse wheel binding
        def on_mousewheel(event):
            self.settings_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.settings_canvas.bind("<MouseWheel>", on_mousewheel)
        
        # Bind canvas resize
        def on_canvas_configure(event):
            self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all"))
            canvas_width = event.width
            if self.settings_canvas.find_all():
                self.settings_canvas.itemconfig(self.settings_canvas.find_all()[0], width=canvas_width)
        
        self.settings_canvas.bind("<Configure>", on_canvas_configure)
    
    def create_email_settings(self):
        """Create email settings card"""
        email_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        email_card.pack(fill=tk.X, pady=(0, 15))
        
        email_inner = tk.Frame(email_card, bg=self.colors['card'])
        email_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(email_inner, text="📧 Email Settings", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Days to check
        days_frame = tk.Frame(email_inner, bg=self.colors['card'])
        days_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(days_frame, text="Check emails from last:", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT)
        
        days_spin = ttk.Spinbox(days_frame, from_=1, to=30, textvariable=self.days_var, 
                               width=5, font=('Segoe UI', 10))
        days_spin.pack(side=tk.LEFT, padx=(10, 5))
        
        tk.Label(days_frame, text="days", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT)
    
    def create_folder_selection(self):
        """Create folder selection card"""
        folder_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        folder_card.pack(fill=tk.X, pady=(0, 15))
        
        folder_inner = tk.Frame(folder_card, bg=self.colors['card'])
        folder_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(folder_inner, text="📁 Outlook Folders to Search", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Instructions
        instr_text = "Select which Outlook folders to search for reports. Click 'Discover Folders' to refresh the list."
        tk.Label(folder_inner, text=instr_text, font=('Segoe UI', 9),
                fg=self.colors['text_secondary'], bg=self.colors['card']).pack(anchor=tk.W, pady=(0, 10))
        
        # Discover folders button
        discover_frame = tk.Frame(folder_inner, bg=self.colors['card'])
        discover_frame.pack(fill=tk.X, pady=(0, 10))
        
        discover_btn = ttk.Button(discover_frame, text="🔍 Discover Folders", command=self.discover_folders)
        discover_btn.pack(side=tk.LEFT)
        
        # Selected folders info label
        self.selected_folders_label = tk.Label(folder_inner, text="", font=('Segoe UI', 9),
                                              bg=self.colors['card'], fg=self.colors['success'], wraplength=600)
        self.selected_folders_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Folder selection area
        self.folder_selection_frame = tk.Frame(folder_inner, bg=self.colors['card'])
        self.folder_selection_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Initially show default folders
        self.create_folder_checkboxes()
    
    def create_file_naming(self):
        """Create file naming card"""
        naming_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        naming_card.pack(fill=tk.X, pady=(0, 15))
        
        naming_inner = tk.Frame(naming_card, bg=self.colors['card'])
        naming_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(naming_inner, text="📝 File Extraction & Naming", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Extraction mode options
        mode_frame = tk.Frame(naming_inner, bg=self.colors['card'])
        mode_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(mode_frame, text="Extraction Mode:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 5))
        
        all_radio = tk.Radiobutton(mode_frame, text="Extract all qualified files", 
                                  variable=self.extraction_mode, value="all",
                                  command=self.toggle_naming_format,
                                  bg=self.colors['card'], fg=self.colors['text'],
                                  font=('Segoe UI', 10), anchor='w',
                                  selectcolor=self.colors['card'], relief='flat',
                                  wraplength=500)
        all_radio.pack(anchor=tk.W, fill=tk.X, pady=2)
        
        latest_radio = tk.Radiobutton(mode_frame, text="Extract only latest file per type", 
                                    variable=self.extraction_mode, value="latest",
                                    command=self.toggle_naming_format,
                                    bg=self.colors['card'], fg=self.colors['text'],
                                    font=('Segoe UI', 10), anchor='w',
                                    selectcolor=self.colors['card'], relief='flat',
                                    wraplength=500)
        latest_radio.pack(anchor=tk.W, fill=tk.X, pady=2)
        
        # Separator
        separator = ttk.Separator(naming_inner, orient='horizontal')
        separator.pack(fill=tk.X, pady=(15, 10))
        
        # Naming format options (only for latest mode)
        self.naming_format_frame = tk.Frame(naming_inner, bg=self.colors['card'])
        
        tk.Label(self.naming_format_frame, text="File Naming Format (Latest Mode Only):", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card'], fg=self.colors['text']).pack(anchor=tk.W, pady=(0, 5))
        
        format_frame = tk.Frame(self.naming_format_frame, bg=self.colors['card'])
        format_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(format_frame, text="Use message received date (YYYY-MM-DD)", 
                       variable=self.naming_format, value="date").pack(anchor=tk.W, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="Use message received year (YYYY)", 
                       variable=self.naming_format, value="year").pack(anchor=tk.W, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="Keep original filename", 
                       variable=self.naming_format, value="original").pack(anchor=tk.W, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="Custom suffix", 
                       variable=self.naming_format, value="custom",
                       command=self.toggle_custom_suffix).pack(anchor=tk.W, padx=(10, 0))
        
        # Custom suffix entry (initially hidden)
        self.custom_suffix_frame = tk.Frame(self.naming_format_frame, bg=self.colors['card'])
        
        tk.Label(self.custom_suffix_frame, text="Custom suffix:", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT)
        
        custom_entry = tk.Entry(self.custom_suffix_frame, textvariable=self.custom_suffix_var,
                               width=20, font=('Segoe UI', 10))
        custom_entry.pack(side=tk.LEFT, padx=(10, 0))
    
    def create_provider_settings(self):
        """Create provider settings card"""
        providers_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        providers_card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        providers_inner = tk.Frame(providers_card, bg=self.colors['card'])
        providers_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        ttk.Label(providers_inner, text="🏢 Service Providers", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Provider text area
        self.providers_text = scrolledtext.ScrolledText(providers_inner, width=70, height=8, 
                                                       font=('Consolas', 9), wrap=tk.WORD,
                                                       bg=self.colors['card'], fg=self.colors['text'])
        self.providers_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Add example text if empty
        if not self.providers_text.get(1.0, tk.END).strip():
            self.providers_text.insert(1.0, "# Enter service providers (one per line)\n# Format: email@example.com = prefix\n# Example:\n# client1@companyA.com = A\n")
        
        # Provider buttons
        provider_btn_frame = tk.Frame(providers_inner, bg=self.colors['card'])
        provider_btn_frame.pack(fill=tk.X)
        
        ttk.Button(provider_btn_frame, text="🔍 Auto-Detect", command=self.auto_detect_providers).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(provider_btn_frame, text="🗑️ Clear", command=self.clear_providers).pack(side=tk.LEFT)
    
    def create_auto_run_option(self):
        """Create auto-run option"""
        auto_run_frame = tk.Frame(self.settings_container, bg=self.colors['bg'])
        auto_run_frame.pack(fill=tk.X, pady=10)
        
        ttk.Checkbutton(auto_run_frame, text="🔄 Auto-run extraction on startup", 
                       variable=self.auto_run_var).pack(anchor=tk.W)
    
    def create_save_button(self):
        """Create save and action buttons"""
        save_frame = tk.Frame(self.settings_container, bg=self.colors['bg'])
        save_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(save_frame, text="💾 Save Settings", command=self.save_settings, 
                  style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(save_frame, text="🚀 Save & Run Extraction", command=self.save_and_run, 
                  style='Primary.TButton').pack(side=tk.LEFT)
    
    def set_main_tab(self, main_tab):
        """Set reference to main tab"""
        self.main_tab = main_tab
    
    def discover_folders(self):
        """Discover Outlook folders"""
        def discover_callback(folders):
            self.update_folder_checkboxes(folders)
            if self.main_tab:
                self.main_tab.log_message(f"Discovered {len(folders)} folders", "SUCCESS")
        
        self.outlook_manager.discover_folders_async(discover_callback)
    
    def create_folder_checkboxes(self):
        """Create checkboxes for folder selection"""
        # Clear existing checkboxes
        for widget in self.folder_selection_frame.winfo_children():
            widget.destroy()
        self.folder_checkboxes = {}
        
        # Get discovered folders or use defaults
        folders = self.outlook_manager.folders
        if not folders:
            default_folders = ["Inbox"]
        else:
            default_folders = sorted(folders.keys())
        
        selected_folders = self.config_manager.get('selected_folders', ['Inbox'])
        
        # Create scrollable area if many folders
        if len(default_folders) > 6:
            canvas_frame = tk.Frame(self.folder_selection_frame, bg=self.colors['card'])
            canvas_frame.pack(fill=tk.BOTH, expand=True)
            
            canvas = tk.Canvas(canvas_frame, height=150, bg=self.colors['card'], highlightthickness=0)
            scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=self.colors['card'])
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar.grid(row=0, column=1, sticky="ns")
            
            canvas_frame.grid_rowconfigure(0, weight=1)
            canvas_frame.grid_columnconfigure(0, weight=1)
            
            # Mouse wheel binding
            def on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind("<MouseWheel>", on_mousewheel)
            
            container = scrollable_frame
        else:
            container = self.folder_selection_frame
        
        # Create checkboxes
        for folder_name in default_folders:
            var = tk.BooleanVar()
            if folder_name in selected_folders:
                var.set(True)
            elif not selected_folders and folder_name == "Inbox":
                var.set(True)
            
            checkbox = ttk.Checkbutton(
                container, 
                text=folder_name, 
                variable=var,
                command=self.update_selected_folders
            )
            checkbox.pack(anchor=tk.W, padx=(20, 0), pady=2)
            self.folder_checkboxes[folder_name] = var
        
        # Update display after creating all checkboxes
        self.update_selected_folders_display()
    
    def update_folder_checkboxes(self, folders: Dict[str, Any]):
        """Update folder checkboxes with discovered folders"""
        self.outlook_manager.folders = folders
        self.create_folder_checkboxes()
        # Load saved folder selections after creating checkboxes
        self.load_saved_settings()
    
    def update_selected_folders(self):
        """Update the list of selected folders (for UI validation only)"""
        selected_folders = [folder for folder, var in self.folder_checkboxes.items() if var.get()]
        if not selected_folders:
            # Ensure at least Inbox is selected
            if "Inbox" in self.folder_checkboxes:
                self.folder_checkboxes["Inbox"].set(True)
                selected_folders = ["Inbox"]
        
        # Update the display label
        self.update_selected_folders_display(selected_folders)
        
        # Don't automatically save to config - let user decide when to save
    
    def update_selected_folders_display(self, selected_folders=None):
        """Update the selected folders display label"""
        if selected_folders is None:
            selected_folders = [folder for folder, var in self.folder_checkboxes.items() if var.get()]
        
        if selected_folders:
            if len(selected_folders) <= 3:
                folder_text = ", ".join(selected_folders)
            else:
                folder_text = f"{', '.join(selected_folders[:2])} and {len(selected_folders)-2} more"
            
            self.selected_folders_label.config(
                text=f"✅ Selected folders ({len(selected_folders)}): {folder_text}",
                fg=self.colors['success']
            )
        else:
            self.selected_folders_label.config(
                text="⚠️ No folders selected - will use Inbox as default",
                fg=self.colors['warning']
            )
    
    def toggle_custom_suffix(self):
        """Toggle visibility of custom suffix entry"""
        if self.naming_format.get() == "custom":
            self.custom_suffix_frame.pack(fill=tk.X, pady=(5, 0))
        else:
            self.custom_suffix_frame.pack_forget()
    
    def toggle_naming_format(self):
        """Toggle visibility of naming format options based on extraction mode"""
        if self.extraction_mode.get() == "latest":
            self.naming_format_frame.pack(fill=tk.X, pady=(0, 10))
        else:
            self.naming_format_frame.pack_forget()
    
    def clear_providers(self):
        """Clear all providers"""
        self.providers_text.delete(1.0, tk.END)
        self.providers_text.insert(1.0, "# Enter service providers (one per line)\n# Format: email@example.com = prefix\n# Example:\n# client1@companyA.com = A\n")
    
    def auto_detect_providers(self):
        """Auto-detect providers (placeholder)"""
        if self.main_tab:
            self.main_tab.log_message("Auto-detection feature coming soon!", "INFO")
    
    def get_current_settings(self):
        """Get current settings from UI components"""
        selected_folders = [folder for folder, var in self.folder_checkboxes.items() if var.get()]
        if not selected_folders:
            selected_folders = ["Inbox"]
        
        return {
            'days': self.days_var.get(),
            'selected_folders': selected_folders,
            'naming_format': self.naming_format.get(),
            'custom_suffix': self.custom_suffix_var.get(),
            'extraction_mode': self.extraction_mode.get(),
            'providers': self.providers_text.get(1.0, "end-1c"),
            'auto_run': self.auto_run_var.get()
        }
    
    def save_settings(self):
        """Save current UI settings to config file (for persistence only)"""
        # Get current folder selections
        selected_folders = [folder for folder, var in self.folder_checkboxes.items() if var.get()]
        if not selected_folders:
            selected_folders = ["Inbox"]
        
        # Update configuration with current UI values (complete settings)
        config_data = {
            'days': self.days_var.get(),
            'auto_run': self.auto_run_var.get(),
            'naming_format': self.naming_format.get(),
            'custom_suffix': self.custom_suffix_var.get(),
            'extraction_mode': self.extraction_mode.get(),
            'selected_folders': selected_folders,
            'providers': self.providers_text.get(1.0, tk.END),
            'discovered_folders': list(self.outlook_manager.folders.keys()) if self.outlook_manager.folders else [],
            'last_saved': datetime.now().isoformat(),
            'version': '1.0'
        }
        
        self.config_manager.update(config_data)
        
        # Save to file
        if self.config_manager.save_config():
            if self.main_tab:
                self.main_tab.log_message("Settings saved to file successfully! (Runtime settings use current UI values)", "SUCCESS")
            return True
        else:
            if self.main_tab:
                self.main_tab.log_message("Error saving settings to file!", "ERROR")
            return False
    
    def save_and_run(self):
        """Save settings and automatically run extraction"""
        try:
            # First, save settings
            if not self.save_settings():
                return
            
            # Validate required settings
            providers_text = self.providers_text.get(1.0, "end-1c").strip()
            if not providers_text:
                if self.main_tab:
                    self.main_tab.log_message("⚠️  Please configure service providers before running extraction", "ERROR")
                return
            
            # Check if we have any selected folders
            selected_folders = [folder for folder, var in self.folder_checkboxes.items() if var.get()]
            if not selected_folders:
                if self.main_tab:
                    self.main_tab.log_message("⚠️  Please select at least one folder to search", "ERROR")
                return
            
            # Test connection first
            if self.main_tab:
                self.main_tab.log_message("🔗 Testing Outlook connection before extraction...", "INFO")
                
            if not self.outlook_manager.test_connection():
                if self.main_tab:
                    self.main_tab.log_message("❌ Connection failed! Cannot run extraction.", "ERROR")
                return
            
            # Trigger extraction in main tab
            if self.main_tab:
                self.main_tab.log_message("✅ Connection verified! Starting extraction...", "SUCCESS")
                self.main_tab.run_extraction()
            
        except Exception as e:
            if self.main_tab:
                self.main_tab.log_message(f"Error in save and run: {str(e)}", "ERROR")
    
    def load_saved_settings(self):
        """Load and apply saved settings to UI"""
        try:
            # Load providers text
            providers_text = self.config_manager.get('providers', '')
            if providers_text and providers_text.strip():
                self.providers_text.delete(1.0, tk.END)
                self.providers_text.insert(1.0, providers_text)
            
            # Load saved folder selections
            saved_folders = self.config_manager.get('selected_folders', [])
            if saved_folders and self.folder_checkboxes:
                # First, uncheck all
                for var in self.folder_checkboxes.values():
                    var.set(False)
                
                # Then check the saved ones
                for folder_name in saved_folders:
                    if folder_name in self.folder_checkboxes:
                        self.folder_checkboxes[folder_name].set(True)
                
                # Update the display
                self.update_selected_folders_display(saved_folders)
                        
                if self.main_tab:
                    self.main_tab.log_message(f"✅ Loaded {len(saved_folders)} saved folder selections", "SUCCESS")
            
            # Show custom suffix field if needed
            if self.config_manager.get('naming_format') == 'custom':
                self.toggle_custom_suffix()
                
        except Exception as e:
            if self.main_tab:
                self.main_tab.log_message(f"Error loading settings: {str(e)}", "WARNING")
    
    def refresh_folder_list(self):
        """Refresh the folder list after connection test"""
        self.discover_folders()
    
    def on_resize(self, event):
        """Handle resize events"""
        # Implement any resize-specific logic for settings tab
        pass