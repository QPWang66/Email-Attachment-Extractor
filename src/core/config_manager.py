"""
Configuration Manager for Email Extractor
Handles loading, saving, and validating configuration settings.
"""

import json
import os
from typing import Dict, Any, List
from pathlib import Path
from datetime import datetime


class ConfigManager:
    """Manages application configuration settings"""
    
    DEFAULT_CONFIG = {
        'keyword': 'report',
        'folder': os.path.expanduser("~/Downloads"),
        'days': 7,
        'providers': '# Enter service providers (one per line)\n# Format: email@example.com = Provider Name\n# Or: Subject keyword = Provider Name\n',
        'auto_run': False,
        'naming_format': 'date',
        'custom_suffix': '',
        'selected_folders': ['Inbox'],
        'window_geometry': '850x750',
        'theme': 'default',
        'convert_all_to_format': '',
        'conversion_enabled': False,
        'custom_format_extension': 'xlsx'
    }
    
    def __init__(self, config_file: str = "email_extractor_config.json"):
        """Initialize configuration manager"""
        self.config_file = config_file
        self.config = self.load_config()
        
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from file or return default"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    config = self.DEFAULT_CONFIG.copy()
                    config.update(loaded_config)
                    return config
        except Exception as e:
            print(f"Error loading config: {e}")
        
        return self.DEFAULT_CONFIG.copy()
    
    def save_config(self, config: Dict[str, Any] = None) -> bool:
        """Save configuration to file"""
        try:
            config_to_save = config if config is not None else self.config
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    def get(self, key: str, default=None):
        """Get configuration value"""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any):
        """Set configuration value"""
        self.config[key] = value
    
    def update(self, updates: Dict[str, Any]):
        """Update multiple configuration values"""
        self.config.update(updates)
    
    def validate_config(self) -> List[str]:
        """Validate configuration and return list of errors"""
        errors = []
        
        # Validate folder path
        folder_path = self.config.get('folder', '')
        if folder_path and not os.path.exists(os.path.dirname(folder_path)):
            errors.append(f"Invalid folder path: {folder_path}")
        
        # Validate days
        days = self.config.get('days', 7)
        if not isinstance(days, int) or days < 1 or days > 365:
            errors.append("Days must be between 1 and 365")
        
        # Validate selected folders
        selected_folders = self.config.get('selected_folders', [])
        if not selected_folders or not isinstance(selected_folders, list):
            errors.append("At least one folder must be selected")
        
        return errors
    
    def reset_to_defaults(self):
        """Reset configuration to default values"""
        self.config = self.DEFAULT_CONFIG.copy()
        return self.save_config()
    
    def backup_config(self, backup_path: str = None) -> bool:
        """Create a backup of current configuration"""
        try:
            if backup_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = f"{self.config_file}.backup_{timestamp}"
            
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error creating backup: {e}")
            return False