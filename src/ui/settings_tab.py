"""
Settings Tab UI Component
Contains all configuration options and folder selection.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from typing import Dict, Any, Optional
from datetime import datetime
from ..core.scheduler import UserLevelScheduler, AutomationSetupDialog


class SettingsTab:
    """Settings tab for configuration management"""
    
    def __init__(self, parent, config_manager, outlook_manager, colors):
        """Initialize settings tab"""
        self.parent = parent
        self.config_manager = config_manager
        self.outlook_manager = outlook_manager
        self.colors = colors
        self.main_tab = None
        
        # Initialize scheduler
        self.scheduler = UserLevelScheduler(config_manager)
        
        # Create main frame
        self.frame = tk.Frame(parent, bg=colors['bg'])
        
        # Initialize variables
        self.folder_checkboxes = {}
        self.days_var = tk.IntVar(value=self.config_manager.get('days', 7))
        self.auto_run_var = tk.BooleanVar(value=self.config_manager.get('auto_run', False))
        self.naming_format = tk.StringVar(value=self.config_manager.get('naming_format', 'date'))
        self.custom_suffix_var = tk.StringVar(value=self.config_manager.get('custom_suffix', ''))
        self.extraction_mode = tk.StringVar(value=self.config_manager.get('extraction_mode', 'all'))
        self.conversion_enabled_var = tk.BooleanVar(value=self.config_manager.get('conversion_enabled', False))
        self.convert_format_var = tk.StringVar(value=self.config_manager.get('convert_all_to_format', 'csv'))
        self.custom_format_var = tk.StringVar(value=self.config_manager.get('custom_format_extension', 'xlsx'))
        
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
        
        # File Conversion Card
        self.create_file_conversion()
        
        # Provider Settings Card
        self.create_provider_settings()
        
        # Automation Settings Card
        self.create_automation_settings()
        
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
    
    def create_file_conversion(self):
        """Create file conversion settings card"""
        conversion_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        conversion_card.pack(fill=tk.X, pady=(0, 15))
        
        conversion_inner = tk.Frame(conversion_card, bg=self.colors['card'])
        conversion_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(conversion_inner, text="🔄 File Format Conversion", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Enable conversion checkbox
        conversion_check = ttk.Checkbutton(conversion_inner, 
                                         text="Convert all downloaded files to a specific format",
                                         variable=self.conversion_enabled_var,
                                         command=self.toggle_conversion_options)
        conversion_check.pack(anchor=tk.W, pady=(0, 10))
        
        # Conversion format selection (initially hidden)
        self.conversion_format_frame = tk.Frame(conversion_inner, bg=self.colors['card'])
        
        tk.Label(self.conversion_format_frame, text="Convert all files to:", font=('Segoe UI', 10),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT, padx=(20, 10))
        
        format_options = ['csv', 'txt', 'pdf', 'custom']
        format_combo = ttk.Combobox(self.conversion_format_frame, textvariable=self.convert_format_var,
                                   values=format_options, state='readonly', width=10)
        format_combo.pack(side=tk.LEFT, padx=(0, 10))
        format_combo.bind('<<ComboboxSelected>>', self.toggle_custom_format)
        
        # Custom format entry (initially hidden)
        self.custom_format_frame = tk.Frame(self.conversion_format_frame, bg=self.colors['card'])
        
        tk.Label(self.custom_format_frame, text="Custom extension:", font=('Segoe UI', 9),
                bg=self.colors['card'], fg=self.colors['text']).pack(side=tk.LEFT, padx=(10, 5))
        
        self.custom_format_var = tk.StringVar(value='xlsx')
        custom_format_entry = tk.Entry(self.custom_format_frame, textvariable=self.custom_format_var,
                                      width=8, font=('Segoe UI', 9))
        custom_format_entry.pack(side=tk.LEFT)
        
        # Info text
        info_text = "Note: Original file type filtering still applies. This converts the saved files to your chosen format."
        info_label = tk.Label(conversion_inner, text=info_text, font=('Segoe UI', 9),
                             fg=self.colors['text_secondary'], bg=self.colors['card'], 
                             wraplength=600, justify=tk.LEFT)
        info_label.pack(anchor=tk.W, pady=(10, 0))
        
        # Initially hide the format selection if conversion is disabled
        self.toggle_conversion_options()
    
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
    
    def create_automation_settings(self):
        """Create automation settings card"""
        automation_card = ttk.Frame(self.settings_container, style='Card.TFrame')
        automation_card.pack(fill=tk.X, pady=(0, 15))
        
        automation_inner = tk.Frame(automation_card, bg=self.colors['card'])
        automation_inner.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Label(automation_inner, text="🤖 Auto-Extraction", style='Subheading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Status display
        status_frame = tk.Frame(automation_inner, bg=self.colors['card'])
        status_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.automation_status_label = tk.Label(status_frame, 
                                               text=self.get_automation_status(),
                                               font=('Segoe UI', 10),
                                               bg=self.colors['card'],
                                               fg=self.colors['text'])
        self.automation_status_label.pack(anchor=tk.W)
        
        self.next_action_label = tk.Label(status_frame,
                                         text=self.get_next_action_text(),
                                         font=('Segoe UI', 9),
                                         bg=self.colors['card'],
                                         fg=self.colors['text_secondary'])
        self.next_action_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Control buttons
        button_frame = tk.Frame(automation_inner, bg=self.colors['card'])
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        if self.config_manager.get('automation_enabled', False):
            # Show disable and settings buttons
            ttk.Button(button_frame, text="⚙️ Change Settings", 
                      command=self.setup_automation).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Button(button_frame, text="❌ Disable Auto-Extraction", 
                      command=self.disable_automation).pack(side=tk.LEFT)
        else:
            # Show main setup button
            setup_btn = ttk.Button(button_frame, text="🚀 Enable Auto-Extraction", 
                                 command=self.setup_automation,
                                 style='Primary.TButton')
            setup_btn.pack(side=tk.LEFT)
    
    def get_automation_status(self) -> str:
        """Get automation status text"""
        if self.config_manager.get('automation_enabled', False):
            schedule_time = self.config_manager.get('schedule_time', '09:00')
            schedule_days = self.config_manager.get('schedule_days', [])
            
            if len(schedule_days) == 7:
                days_text = "Daily"
            elif len(schedule_days) == 5 and all(day in ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'] for day in schedule_days):
                days_text = "Weekdays"
            else:
                days_text = f"{len(schedule_days)} days/week"
            
            return f"✅ Auto-extraction: ON ({days_text} at {schedule_time})"
        else:
            return "⚪ Auto-extraction: OFF"
    
    def get_next_action_text(self) -> str:
        """Get next action text"""
        if self.config_manager.get('automation_enabled', False):
            return f"Next: {self.scheduler.get_next_scheduled_action()}"
        else:
            return "Click 'Enable Auto-Extraction' to set up scheduled runs"
    
    def setup_automation(self):
        """Show automation setup dialog"""
        try:
            # Show setup dialog
            setup_dialog = AutomationSetupDialog(self.parent)
            
            # Wait for dialog to complete
            self.parent.wait_window(setup_dialog.dialog)
            
            if setup_dialog.result:
                settings = setup_dialog.get_settings()
                
                # Perform one-click setup
                success = self.scheduler.one_click_setup(
                    settings['schedule_time'],
                    settings['schedule_days'],
                    settings['show_notifications']
                )
                
                if success:
                    # Update UI
                    self.update_automation_display()
                    
                    messagebox.showinfo("✅ Success!", 
                        "Auto-extraction is now enabled!\n\n"
                        "The system will:\n"
                        "• Check daily when you log in\n"
                        "• Notify you when extraction is ready\n"
                        "• Handle missed runs automatically\n\n"
                        "You can change these settings anytime.")
                    
                    if self.main_tab:
                        self.main_tab.log_message("✅ Auto-extraction enabled successfully!", "SUCCESS")
                else:
                    messagebox.showerror("Setup Failed", 
                        "Could not enable automation. Check permissions and try again.")
                    
                    if self.main_tab:
                        self.main_tab.log_message("❌ Auto-extraction setup failed", "ERROR")
            
        except Exception as e:
            messagebox.showerror("Error", f"Setup failed: {str(e)}")
            if self.main_tab:
                self.main_tab.log_message(f"❌ Automation setup error: {str(e)}", "ERROR")
    
    def disable_automation(self):
        """Disable automation"""
        if messagebox.askyesno("Disable Auto-Extraction", 
                              "Stop automatic email extraction?\n\nThis will remove all scheduled tasks and startup integration."):
            try:
                success = self.scheduler.disable_automation()
                
                if success:
                    self.update_automation_display()
                    messagebox.showinfo("✅ Disabled", "Auto-extraction has been disabled successfully.")
                    
                    if self.main_tab:
                        self.main_tab.log_message("✅ Auto-extraction disabled", "SUCCESS")
                else:
                    messagebox.showerror("Error", "Could not fully disable automation. Some components may remain.")
                    
                    if self.main_tab:
                        self.main_tab.log_message("⚠️ Automation disable incomplete", "WARNING")
                        
            except Exception as e:
                messagebox.showerror("Error", f"Failed to disable automation: {str(e)}")
                if self.main_tab:
                    self.main_tab.log_message(f"❌ Disable automation error: {str(e)}", "ERROR")
    
    def update_automation_display(self):
        """Update automation status display"""
        try:
            # Recreate the automation settings section to reflect current state
            # Find and destroy the current automation card
            for widget in self.settings_container.winfo_children():
                if isinstance(widget, ttk.Frame):
                    # Check if this is the automation card by looking for the automation label
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for grandchild in child.winfo_children():
                                if (isinstance(grandchild, ttk.Label) and 
                                    hasattr(grandchild, 'cget') and
                                    "🤖 Auto-Extraction" in str(grandchild.cget('text'))):
                                    widget.destroy()
                                    break
            
            # Recreate automation settings
            self.create_automation_settings()
            
        except Exception as e:
            print(f"Error updating automation display: {e}")
    
    def check_automation_status_on_startup(self):
        """Check if automation should run on startup"""
        try:
            if self.scheduler.should_run_extraction():
                if self.config_manager.get('show_notifications', True):
                    # Show notification that extraction is ready
                    self.show_extraction_ready_notification()
                    
        except Exception as e:
            if self.main_tab:
                self.main_tab.log_message(f"Startup automation check failed: {str(e)}", "WARNING")
    
    def show_extraction_ready_notification(self):
        """Show notification that extraction is ready to run"""
        result = messagebox.askyesno("📧 Auto-Extraction Ready", 
                                    "Scheduled email extraction is ready to run.\n\n"
                                    "Run extraction now?")
        if result and self.main_tab:
            self.main_tab.run_extraction()
    
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
    
    def toggle_conversion_options(self):
        """Toggle visibility of conversion format options"""
        if self.conversion_enabled_var.get():
            self.conversion_format_frame.pack(fill=tk.X, pady=(5, 0))
            self.toggle_custom_format()
        else:
            self.conversion_format_frame.pack_forget()
    
    def toggle_custom_format(self, event=None):
        """Toggle visibility of custom format entry"""
        if self.convert_format_var.get() == "custom":
            self.custom_format_frame.pack(side=tk.LEFT, padx=(10, 0))
        else:
            self.custom_format_frame.pack_forget()
    
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
            'auto_run': self.auto_run_var.get(),
            'conversion_enabled': self.conversion_enabled_var.get(),
            'convert_all_to_format': self.convert_format_var.get(),
            'custom_format_extension': self.custom_format_var.get()
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
            'version': '1.0',
            'conversion_enabled': self.conversion_enabled_var.get(),
            'convert_all_to_format': self.convert_format_var.get(),
            'custom_format_extension': self.custom_format_var.get()
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