"""
Outlook Manager for Email Extractor
Handles all Outlook connections, folder discovery, and email processing.
"""

import win32com.client
import pythoncom
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Callable, Any, Optional
from dataclasses import dataclass


@dataclass
class EmailMessage:
    """Data class for email message information"""
    subject: str
    sender: str
    received_time: datetime
    body: str
    attachments: List[str]
    folder_name: str


class OutlookManager:
    """Manages Outlook connections and operations"""
    
    def __init__(self, progress_callback: Callable[[str, str], None] = None):
        """Initialize Outlook manager"""
        self.progress_callback = progress_callback or self._default_callback
        self.folders = {}
        self.is_connected = False
        self._outlook_app = None
        self._namespace = None
    
    def _default_callback(self, message: str, level: str = "INFO"):
        """Default progress callback"""
        print(f"[{level}] {message}")
    
    def _log(self, message: str, level: str = "INFO"):
        """Log message using callback"""
        self.progress_callback(message, level)
    
    def test_connection(self) -> bool:
        """Test Outlook connection"""
        try:
            self._log("Testing Outlook connection...", "INFO")
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon()
            
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
            count = inbox.Items.Count
            
            self._log(f"Connection successful! Found {count} emails in inbox", "SUCCESS")
            self.is_connected = True
            return True
            
        except Exception as e:
            self._log(f"Connection failed: {str(e)}", "ERROR")
            self.is_connected = False
            return False
        finally:
            pythoncom.CoUninitialize()
    
    def discover_folders(self) -> Dict[str, Any]:
        """Discover all available Outlook folders"""
        self.folders = {}
        
        try:
            pythoncom.CoInitialize()
            self._log("Discovering Outlook folders...", "INFO")
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon()
            
            # Get the default store (main mailbox)
            default_store = namespace.DefaultStore
            root_folder = default_store.GetRootFolder()
            
            def traverse_folders(folder, path=""):
                try:
                    folder_name = folder.Name
                    full_path = f"{path}/{folder_name}" if path else folder_name
                    
                    # Skip system folders we don't want
                    skip_folders = {"Deleted Items", "Drafts", "Outbox", "Journal", 
                                  "Notes", "Tasks", "Calendar", "Contacts"}
                    
                    if folder_name not in skip_folders:
                        self.folders[full_path] = {
                            'folder_object': folder,
                            'name': folder_name,
                            'path': full_path,
                            'item_count': getattr(folder.Items, 'Count', 0)
                        }
                    
                    # Recursively traverse subfolders
                    if hasattr(folder, 'Folders') and folder.Folders.Count > 0:
                        for subfolder in folder.Folders:
                            traverse_folders(subfolder, full_path)
                except Exception as e:
                    self._log(f"Error processing folder {folder_name}: {e}", "WARNING")
            
            # Start traversal from root
            traverse_folders(root_folder)
            
            # Also add some common default folders directly
            try:
                inbox = namespace.GetDefaultFolder(6)
                sent_items = namespace.GetDefaultFolder(5)
                
                self.folders["Inbox"] = {
                    'folder_object': inbox,
                    'name': "Inbox",
                    'path': "Inbox",
                    'item_count': inbox.Items.Count
                }
                
                self.folders["Sent Items"] = {
                    'folder_object': sent_items,
                    'name': "Sent Items", 
                    'path': "Sent Items",
                    'item_count': sent_items.Items.Count
                }
                
            except Exception as e:
                self._log(f"Error adding default folders: {e}", "WARNING")
            
            self._log(f"Discovered {len(self.folders)} folders", "SUCCESS")
            return self.folders
            
        except Exception as e:
            self._log(f"Error discovering folders: {str(e)}", "ERROR")
            return {}
        finally:
            pythoncom.CoUninitialize()
    
    def discover_folders_async(self, callback: Callable[[Dict[str, Any]], None]):
        """Discover folders asynchronously"""
        def discover_in_thread():
            folders = self.discover_folders()
            callback(folders)
        
        thread = threading.Thread(target=discover_in_thread, daemon=True)
        thread.start()
    
    def get_messages_from_folders(self, folder_names: List[str], 
                                 days_back: int = 7,
                                 keyword: str = "") -> List[EmailMessage]:
        """Get messages from specified folders"""
        messages = []
        
        try:
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon()
            
            # Calculate date range
            start_date = datetime.now() - timedelta(days=days_back)
            
            self._log(f"Searching {len(folder_names)} folder(s) for messages...", "INFO")
            
            for folder_name in folder_names:
                try:
                    folder_info = self.folders.get(folder_name)
                    if folder_info:
                        folder = folder_info['folder_object']
                    elif folder_name == "Inbox":
                        folder = namespace.GetDefaultFolder(6)
                    elif folder_name == "Sent Items":
                        folder = namespace.GetDefaultFolder(5)
                    else:
                        self._log(f"Folder '{folder_name}' not found", "WARNING")
                        continue
                    
                    folder_messages = folder.Items
                    folder_messages.Sort("[ReceivedTime]", True)
                    
                    folder_count = 0
                    for message in folder_messages:
                        try:
                            # Check date range
                            if message.ReceivedTime.date() < start_date.date():
                                continue
                            
                            # Check keyword if specified
                            if keyword and keyword.lower() not in message.Subject.lower():
                                continue
                            
                            # Get attachments
                            attachments = []
                            try:
                                for attachment in message.Attachments:
                                    attachments.append(attachment.FileName)
                            except:
                                pass
                            
                            email_msg = EmailMessage(
                                subject=message.Subject,
                                sender=getattr(message, 'SenderEmailAddress', 'Unknown'),
                                received_time=message.ReceivedTime,
                                body=getattr(message, 'Body', ''),
                                attachments=attachments,
                                folder_name=folder_name
                            )
                            
                            messages.append(email_msg)
                            folder_count += 1
                            
                        except Exception as e:
                            self._log(f"Error processing message: {e}", "WARNING")
                            continue
                    
                    self._log(f"Found {folder_count} matching messages in '{folder_name}'", "INFO")
                    
                except Exception as e:
                    self._log(f"Error accessing folder '{folder_name}': {e}", "ERROR")
                    continue
            
            # Sort all messages by received time (newest first)
            messages.sort(key=lambda msg: msg.received_time, reverse=True)
            
            self._log(f"Total messages found: {len(messages)}", "SUCCESS")
            return messages
            
        except Exception as e:
            self._log(f"Error retrieving messages: {str(e)}", "ERROR")
            return []
        finally:
            pythoncom.CoUninitialize()
    
    def get_messages_async(self, folder_names: List[str], 
                          days_back: int = 7,
                          keyword: str = "",
                          callback: Callable[[List[EmailMessage]], None] = None):
        """Get messages asynchronously"""
        def get_in_thread():
            messages = self.get_messages_from_folders(folder_names, days_back, keyword)
            if callback:
                callback(messages)
        
        thread = threading.Thread(target=get_in_thread, daemon=True)
        thread.start()
    
    def save_attachments(self, messages: List[EmailMessage], 
                        save_folder: str,
                        file_naming_format: str = "date",
                        custom_suffix: str = "") -> Dict[str, str]:
        """Save attachments from messages"""
        saved_files = {}
        
        try:
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon()
            
            for message in messages:
                try:
                    # Find the actual Outlook message object
                    # This is a simplified approach - in practice you'd need to match by properties
                    # For now, we'll use the folder and search for matching subject/time
                    
                    # Generate filename based on format
                    if file_naming_format == "date":
                        suffix = datetime.now().strftime("%Y-%m-%d")
                    elif file_naming_format == "year":
                        suffix = datetime.now().strftime("%Y")
                    else:  # custom
                        suffix = custom_suffix or "custom"
                    
                    # Save attachments logic would go here
                    # This is a placeholder for the actual implementation
                    
                except Exception as e:
                    self._log(f"Error saving attachments for message '{message.subject}': {e}", "ERROR")
                    continue
            
            return saved_files
            
        except Exception as e:
            self._log(f"Error saving attachments: {str(e)}", "ERROR")
            return {}
        finally:
            pythoncom.CoUninitialize()
    
    def get_folder_statistics(self) -> Dict[str, Dict[str, int]]:
        """Get statistics for discovered folders"""
        stats = {}
        
        for folder_name, folder_info in self.folders.items():
            stats[folder_name] = {
                'total_items': folder_info.get('item_count', 0),
                'name': folder_info.get('name', folder_name)
            }
        
        return stats