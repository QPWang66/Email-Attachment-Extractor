"""
File Management Utilities
Handles file operations, organization, and backup functionality.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional
import zipfile


class FileManager:
    """Manages file operations and organization"""
    
    def __init__(self, base_folder: str):
        """Initialize file manager"""
        self.base_folder = Path(base_folder)
        self.ensure_folder_exists(self.base_folder)
    
    def ensure_folder_exists(self, folder_path: Path) -> bool:
        """Ensure a folder exists, create if necessary"""
        try:
            folder_path.mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e:
            print(f"Error creating folder {folder_path}: {e}")
            return False
    
    def organize_by_date(self, enable: bool = True) -> Path:
        """Create date-based folder structure"""
        if not enable:
            return self.base_folder
        
        today = datetime.now()
        date_folder = self.base_folder / today.strftime("%Y") / today.strftime("%m-%B")
        self.ensure_folder_exists(date_folder)
        return date_folder
    
    def organize_by_provider(self, provider: str) -> Path:
        """Create provider-based folder structure"""
        # Clean provider name for folder
        clean_provider = "".join(c for c in provider if c.isalnum() or c in (' ', '-', '_')).strip()
        clean_provider = clean_provider.replace(' ', '_')
        
        provider_folder = self.base_folder / clean_provider
        self.ensure_folder_exists(provider_folder)
        return provider_folder
    
    def organize_by_category(self, category: str) -> Path:
        """Create category-based folder structure"""
        category_folder = self.base_folder / category.title()
        self.ensure_folder_exists(category_folder)
        return category_folder
    
    def get_organized_path(self, provider: str = "", category: str = "", 
                          use_date: bool = True) -> Path:
        """Get organized folder path based on preferences"""
        current_path = self.base_folder
        
        if use_date:
            current_path = self.organize_by_date(True)
        
        if provider:
            current_path = current_path / self.clean_name(provider)
            self.ensure_folder_exists(current_path)
        
        if category:
            current_path = current_path / category.title()
            self.ensure_folder_exists(current_path)
        
        return current_path
    
    def clean_name(self, name: str) -> str:
        """Clean name for use in file/folder names"""
        # Remove invalid characters
        cleaned = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_', '.')).strip()
        # Replace spaces with underscores
        cleaned = cleaned.replace(' ', '_')
        # Remove multiple underscores
        while '__' in cleaned:
            cleaned = cleaned.replace('__', '_')
        return cleaned
    
    def save_file_safely(self, source_path: str, destination_folder: Path, 
                        filename: str) -> Optional[Path]:
        """Safely save a file with conflict resolution"""
        try:
            destination_folder = Path(destination_folder)
            self.ensure_folder_exists(destination_folder)
            
            # Clean filename
            safe_filename = self.clean_name(filename)
            destination_path = destination_folder / safe_filename
            
            # Handle filename conflicts
            counter = 1
            original_path = destination_path
            while destination_path.exists():
                name_part = original_path.stem
                ext_part = original_path.suffix
                destination_path = destination_folder / f"{name_part}_{counter}{ext_part}"
                counter += 1
            
            # Copy the file
            shutil.copy2(source_path, destination_path)
            return destination_path
            
        except Exception as e:
            print(f"Error saving file {filename}: {e}")
            return None
    
    def create_backup(self, files: List[Path], backup_name: str = None) -> Optional[Path]:
        """Create a backup archive of files"""
        try:
            if not backup_name:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_name = f"email_reports_backup_{timestamp}.zip"
            
            backup_path = self.base_folder / "Backups"
            self.ensure_folder_exists(backup_path)
            
            backup_file = backup_path / backup_name
            
            with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in files:
                    if file_path.exists():
                        # Add file to zip with relative path
                        arcname = file_path.relative_to(self.base_folder)
                        zipf.write(file_path, arcname)
            
            return backup_file
            
        except Exception as e:
            print(f"Error creating backup: {e}")
            return None
    
    def get_file_statistics(self) -> Dict[str, int]:
        """Get statistics about saved files"""
        stats = {
            'total_files': 0,
            'total_size_mb': 0,
            'file_types': {},
            'folders': 0
        }
        
        try:
            for item in self.base_folder.rglob('*'):
                if item.is_file():
                    stats['total_files'] += 1
                    stats['total_size_mb'] += item.stat().st_size / (1024 * 1024)
                    
                    ext = item.suffix.lower()
                    if ext:
                        stats['file_types'][ext] = stats['file_types'].get(ext, 0) + 1
                elif item.is_dir():
                    stats['folders'] += 1
        
        except Exception as e:
            print(f"Error calculating statistics: {e}")
        
        return stats
    
    def cleanup_empty_folders(self) -> int:
        """Remove empty folders"""
        removed_count = 0
        
        try:
            # Walk through folders in reverse order (deepest first)
            for folder in sorted(self.base_folder.rglob('*'), key=lambda p: len(p.parts), reverse=True):
                if folder.is_dir() and folder != self.base_folder:
                    try:
                        if not any(folder.iterdir()):  # Folder is empty
                            folder.rmdir()
                            removed_count += 1
                    except OSError:
                        # Folder not empty or can't be removed
                        pass
        except Exception as e:
            print(f"Error during cleanup: {e}")
        
        return removed_count
    
    def get_recent_files(self, days: int = 7) -> List[Path]:
        """Get files modified in the last N days"""
        recent_files = []
        cutoff_time = datetime.now().timestamp() - (days * 24 * 60 * 60)
        
        try:
            for file_path in self.base_folder.rglob('*'):
                if file_path.is_file() and file_path.stat().st_mtime > cutoff_time:
                    recent_files.append(file_path)
        except Exception as e:
            print(f"Error getting recent files: {e}")
        
        return sorted(recent_files, key=lambda p: p.stat().st_mtime, reverse=True)