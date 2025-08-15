"""
User-Level Scheduler for Email Extractor
Handles automation without requiring admin privileges.
"""

import os
import json
import subprocess
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta, time
from typing import Dict, List, Optional
from pathlib import Path


class UserLevelScheduler:
    """Manages user-level automation and scheduling"""
    
    def __init__(self, config_manager):
        self.config = config_manager
        self.state_file = os.path.expanduser("~/AppData/Local/EmailExtractor/automation_state.json")
        self.startup_folder = os.path.expanduser(
            "~/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup"
        )
        self.ensure_state_directory()
    
    def ensure_state_directory(self):
        """Ensure automation state directory exists"""
        os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
    
    def should_run_extraction(self) -> bool:
        """Determine if extraction should run based on schedule and last run"""
        if not self.config.get('automation_enabled', False):
            return False
        
        schedule_time = self.config.get('schedule_time', '09:00')
        schedule_days = self.config.get('schedule_days', ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'])
        last_run = self.get_last_run_time()
        current_time = datetime.now()
        
        # Check if today is a scheduled day
        current_day = current_time.strftime('%A').lower()
        if current_day not in schedule_days:
            return False
        
        # Calculate when extraction should have happened today
        schedule_hour, schedule_minute = map(int, schedule_time.split(':'))
        today_schedule = datetime.combine(
            current_time.date(), 
            time(schedule_hour, schedule_minute)
        )
        
        # If we're past scheduled time and haven't run today
        if current_time > today_schedule and (not last_run or last_run.date() < current_time.date()):
            return True
        
        # Check for missed runs from previous days
        return self.check_missed_runs(last_run, current_time)
    
    def check_missed_runs(self, last_run: Optional[datetime], current_time: datetime) -> bool:
        """Check if any scheduled runs were missed"""
        if not last_run:
            return True  # Never run before
        
        schedule_days = self.config.get('schedule_days', ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'])
        check_date = last_run.date() + timedelta(days=1)
        
        while check_date < current_time.date():
            if check_date.strftime('%A').lower() in schedule_days:
                return True  # Found a missed day
            check_date += timedelta(days=1)
        
        return False
    
    def get_last_run_time(self) -> Optional[datetime]:
        """Get the last successful run time"""
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r') as f:
                    state = json.load(f)
                    last_run_str = state.get('last_run')
                    if last_run_str:
                        return datetime.fromisoformat(last_run_str)
        except Exception as e:
            print(f"Error reading last run time: {e}")
        return None
    
    def save_last_run(self):
        """Record successful extraction time"""
        try:
            state = self.load_automation_state()
            state['last_run'] = datetime.now().isoformat()
            state['last_check'] = datetime.now().isoformat()
            self.save_automation_state(state)
        except Exception as e:
            print(f"Error saving last run time: {e}")
    
    def load_automation_state(self) -> Dict:
        """Load automation state from file"""
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading automation state: {e}")
        
        return {
            'last_run': None,
            'last_check': None,
            'pending_notifications': [],
            'missed_runs': []
        }
    
    def save_automation_state(self, state: Dict):
        """Save automation state to file"""
        try:
            with open(self.state_file, 'w') as f:
                json.dump(state, f, indent=4)
        except Exception as e:
            print(f"Error saving automation state: {e}")
    
    def one_click_setup(self, schedule_time: str, schedule_days: List[str], 
                       show_notifications: bool = True) -> bool:
        """Complete automation setup with one operation"""
        try:
            # 1. Save automation settings
            self.config.update({
                'automation_enabled': True,
                'schedule_time': schedule_time,
                'schedule_days': schedule_days,
                'show_notifications': show_notifications
            })
            self.config.save_config()
            
            # 2. Create startup integration
            self.create_startup_checker()
            
            # 3. Set up user task scheduler (optional, may fail on restricted systems)
            try:
                self.setup_user_task_scheduler(schedule_time)
            except Exception as e:
                print(f"Task scheduler setup failed (using startup mode): {e}")
                self.config.set('automation_mode', 'startup_only')
                self.config.save_config()
            
            # 4. Initialize tracking files
            self.initialize_automation_state()
            
            return True
            
        except Exception as e:
            print(f"Automation setup failed: {e}")
            return False
    
    def create_startup_checker(self):
        """Create startup integration script"""
        try:
            os.makedirs(self.startup_folder, exist_ok=True)
            
            # Create batch file that checks schedule on startup
            script_path = os.path.abspath('main.py')
            batch_content = f'''@echo off
cd /d "{os.getcwd()}"
python "{script_path}" --check-schedule --silent > nul 2>&1
'''
            
            startup_file = os.path.join(self.startup_folder, "EmailExtractorCheck.bat")
            with open(startup_file, 'w') as f:
                f.write(batch_content)
                
        except Exception as e:
            raise Exception(f"Could not create startup checker: {e}")
    
    def setup_user_task_scheduler(self, schedule_time: str):
        """Create user-level scheduled task"""
        try:
            script_path = os.path.abspath('main.py')
            task_name = "EmailExtractorDaily"
            
            # Delete existing task if it exists
            subprocess.run(
                f'schtasks /delete /tn "{task_name}" /f',
                shell=True, capture_output=True
            )
            
            # Create new task
            cmd = [
                'schtasks', '/create', '/f',
                '/tn', task_name,
                '/tr', f'python "{script_path}" --check-schedule',
                '/sc', 'daily',
                '/st', schedule_time,
                '/ru', os.environ.get('USERNAME', 'SYSTEM')
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise Exception(f"Task creation failed: {result.stderr}")
                
        except Exception as e:
            raise Exception(f"Task scheduler setup failed: {e}")
    
    def initialize_automation_state(self):
        """Initialize automation state tracking"""
        initial_state = {
            'last_run': None,
            'last_check': datetime.now().isoformat(),
            'pending_notifications': [],
            'missed_runs': [],
            'setup_date': datetime.now().isoformat()
        }
        self.save_automation_state(initial_state)
    
    def disable_automation(self) -> bool:
        """Disable all automation features"""
        try:
            # Remove startup file
            startup_file = os.path.join(self.startup_folder, "EmailExtractorCheck.bat")
            if os.path.exists(startup_file):
                os.remove(startup_file)
            
            # Remove scheduled task
            subprocess.run(
                'schtasks /delete /tn "EmailExtractorDaily" /f',
                shell=True, capture_output=True
            )
            
            # Update config
            self.config.set('automation_enabled', False)
            self.config.save_config()
            
            return True
            
        except Exception as e:
            print(f"Error disabling automation: {e}")
            return False
    
    def get_next_scheduled_action(self) -> str:
        """Get description of next scheduled action"""
        if not self.config.get('automation_enabled', False):
            return "Automation disabled"
        
        schedule_time = self.config.get('schedule_time', '09:00')
        schedule_days = self.config.get('schedule_days', [])
        
        if self.should_run_extraction():
            return "Ready to run now!"
        
        # Find next scheduled day
        current_time = datetime.now()
        for i in range(1, 8):  # Check next 7 days
            check_date = current_time + timedelta(days=i)
            if check_date.strftime('%A').lower() in schedule_days:
                return f"{check_date.strftime('%A')} at {schedule_time}"
        
        return "No upcoming schedule"


class AutomationSetupDialog:
    """Simple dialog for automation setup"""
    
    def __init__(self, parent):
        self.parent = parent
        self.result = False
        self.dialog = None
        self.schedule_time = tk.StringVar(value="09:00")
        self.show_notifications = tk.BooleanVar(value=True)
        
        # Working days
        self.working_days = {
            'monday': tk.BooleanVar(value=True),
            'tuesday': tk.BooleanVar(value=True),
            'wednesday': tk.BooleanVar(value=True),
            'thursday': tk.BooleanVar(value=True),
            'friday': tk.BooleanVar(value=True),
            'saturday': tk.BooleanVar(value=False),
            'sunday': tk.BooleanVar(value=False)
        }
        
        self.create_dialog()
    
    def create_dialog(self):
        """Create the setup dialog"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Setup Auto-Extraction")
        self.dialog.geometry("400x300")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (
            self.parent.winfo_rootx() + 50,
            self.parent.winfo_rooty() + 50
        ))
        
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="🚀 Setup Email Auto-Extraction", 
                              font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Schedule time
        time_frame = tk.Frame(main_frame)
        time_frame.pack(fill='x', pady=5)
        tk.Label(time_frame, text="⏰ Run daily at:", font=('Arial', 10)).pack(side='left')
        time_entry = tk.Entry(time_frame, textvariable=self.schedule_time, width=8)
        time_entry.pack(side='right')
        
        # Working days
        days_frame = tk.LabelFrame(main_frame, text="📅 Working Days", font=('Arial', 10))
        days_frame.pack(fill='x', pady=10)
        
        days_grid = tk.Frame(days_frame)
        days_grid.pack(padx=10, pady=5)
        
        day_names = [
            ('M', 'monday'), ('T', 'tuesday'), ('W', 'wednesday'), 
            ('T', 'thursday'), ('F', 'friday'), ('S', 'saturday'), ('S', 'sunday')
        ]
        
        for i, (short_name, full_name) in enumerate(day_names):
            cb = tk.Checkbutton(days_grid, text=short_name, variable=self.working_days[full_name])
            cb.grid(row=0, column=i, padx=5)
        
        # Notifications
        notif_frame = tk.Frame(main_frame)
        notif_frame.pack(fill='x', pady=10)
        tk.Checkbutton(notif_frame, text="🔔 Show notifications", 
                      variable=self.show_notifications, font=('Arial', 10)).pack(side='left')
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(20, 0))
        
        tk.Button(button_frame, text="🚀 Setup Now", command=self.setup_clicked,
                 bg='#4CAF50', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=(5, 0))
        tk.Button(button_frame, text="❌ Cancel", command=self.cancel_clicked,
                 font=('Arial', 10)).pack(side='right')
    
    def setup_clicked(self):
        """Handle setup button click"""
        self.result = True
        self.dialog.destroy()
    
    def cancel_clicked(self):
        """Handle cancel button click"""
        self.result = False
        self.dialog.destroy()
    
    def get_settings(self) -> Dict:
        """Get the selected settings"""
        selected_days = [day for day, var in self.working_days.items() if var.get()]
        
        return {
            'schedule_time': self.schedule_time.get(),
            'schedule_days': selected_days,
            'show_notifications': self.show_notifications.get()
        }