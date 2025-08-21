"""
Outlook Manager for Email Extractor
Handles all Outlook connections, folder discovery, and email processing.
"""

import win32com.client
import pythoncom
import threading
import os
from datetime import datetime, timedelta
from typing import Dict, List, Callable, Any, Optional
from dataclasses import dataclass
from ..utils.email_processor import EmailProcessor


@dataclass
class EmailMessage:
    """Data class for email message information"""
    subject: str
    sender: str
    received_time: datetime
    body: str
    attachments: List[str]
    folder_name: str
    provider_name: str = ""  # Provider name from service provider mapping


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
    
    def _find_folder_by_path(self, namespace, folder_path: str):
        """Find a folder by its path using fresh connection"""
        try:
            # For email-based paths, we need to search all stores, not just the default store
            # Outlook often puts shared mailboxes and additional accounts as separate stores
            
            for store in namespace.Stores:
                try:
                    root_folder = store.GetRootFolder()
                    found_folder = self._search_folder_recursive(root_folder, folder_path)
                    if found_folder:
                        return found_folder
                except Exception as e:
                    continue
            
            self._log(f"Folder '{folder_path}' not found in any store", "WARNING")
            return None
            
        except Exception as e:
            self._log(f"Error finding folder '{folder_path}': {e}", "WARNING")
            return None
    
    def _search_folder_recursive(self, parent_folder, target_path: str):
        """Recursively search for a folder by path"""
        try:
            # Check if this folder matches the target path
            current_path = self._get_folder_path(parent_folder)
            if current_path == target_path:
                return parent_folder
            
            # Search in subfolders
            if hasattr(parent_folder, 'Folders'):
                for subfolder in parent_folder.Folders:
                    try:
                        result = self._search_folder_recursive(subfolder, target_path)
                        if result:
                            return result
                    except:
                        continue
            
            return None
            
        except Exception as e:
            return None
    
    def _get_folder_path(self, folder):
        """Get the full path of a folder"""
        try:
            path_parts = []
            current = folder
            
            # Walk up the folder hierarchy
            while current and hasattr(current, 'Name'):
                name = current.Name
                # Skip the root folder name (usually the store name)
                if hasattr(current, 'Parent') and current.Parent:
                    path_parts.insert(0, name)
                    current = current.Parent
                else:
                    break
            
            return '/'.join(path_parts) if path_parts else folder.Name
            
        except:
            try:
                return folder.Name
            except:
                return ""
    
    def _parse_provider_settings(self, providers_text: str) -> Dict[str, str]:
        """Parse provider settings text into email -> provider mapping"""
        providers = {}
        
        if not providers_text or not providers_text.strip():
            self._log("📋 No provider filters configured - including all senders", "INFO")
            return providers
        
        self._log("🔧 Parsing provider filters...", "INFO")
        
        for line_num, line in enumerate(providers_text.split('\n'), 1):
            line = line.strip()
            
            # Skip comments and empty lines
            if not line or line.startswith('#'):
                continue
            
            # Parse "email@domain.com = Provider Name" format
            if '=' in line:
                email_part, provider_part = line.split('=', 1)
                email = email_part.strip().lower()
                provider = provider_part.strip()
                
                if email and provider:
                    providers[email] = provider
                    self._log(f"📌 Rule {len(providers)}: '{email}' → '{provider}'", "SUCCESS")
                else:
                    self._log(f"⚠️  Line {line_num}: Invalid format '{line}' (missing email or provider)", "WARNING")
            else:
                self._log(f"⚠️  Line {line_num}: Invalid format '{line}' (missing '=')", "WARNING")
        
        if providers:
            self._log(f"✅ Configured {len(providers)} provider filter(s) - will ONLY process emails from these senders", "SUCCESS")
        else:
            self._log("⚠️  No valid provider filters found - including all senders", "WARNING")
        
        return providers
    
    def _should_include_message(self, message, providers: Dict[str, str]) -> tuple[bool, str]:
        """Check if message should be included based on provider filters"""
        if not providers:
            return True, ""  # No filters, include all
        
        try:
            # Get sender email (try multiple methods to get actual SMTP address)
            sender_email = ""
            
            # Method 1: Try SenderEmailAddress first
            if hasattr(message, 'SenderEmailAddress') and message.SenderEmailAddress:
                addr = message.SenderEmailAddress.lower()
                # Skip Exchange internal paths - look for actual email format
                if '@' in addr and not addr.startswith('/o='):
                    sender_email = addr
            
            # Method 2: Try Sender.Address if we don't have a good email yet
            if not sender_email and hasattr(message, 'Sender') and message.Sender:
                try:
                    addr = message.Sender.Address.lower()
                    if '@' in addr and not addr.startswith('/o='):
                        sender_email = addr
                except:
                    pass
            
            # Method 3: Try to get SMTP address from PropertyAccessor
            if not sender_email:
                try:
                    # PR_SENDER_SMTP_ADDRESS property
                    smtp_addr = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
                    if smtp_addr and '@' in smtp_addr:
                        sender_email = smtp_addr.lower()
                except:
                    pass
            
            # Method 4: Try SenderName as fallback for internal addresses
            if not sender_email and hasattr(message, 'SenderName'):
                try:
                    name = message.SenderName
                    if '@' in name:
                        sender_email = name.lower()
                except:
                    pass
            
            if not sender_email:
                self._log(f"⚠️  Message '{message.Subject}' has no sender email - including by default", "WARNING")
                return True, ""
            
            # Check if sender matches any configured providers
            for configured_email, provider_name in providers.items():
                if configured_email in sender_email or sender_email in configured_email:
                    self._log(f"✅ MATCH: '{message.Subject}' from {sender_email} → Provider: '{provider_name}'", "SUCCESS")
                    return True, provider_name
            
            return False, ""
            
        except Exception as e:
            self._log(f"Error checking provider filter for '{message.Subject}': {e}", "WARNING")
            return True, ""  # Include on error
    
    
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
            
            # Debug: Log all discovered folder paths
            self._log("Available folders:", "INFO")
            for folder_path in sorted(self.folders.keys()):
                item_count = self.folders[folder_path].get('item_count', 0)
                self._log(f"  - '{folder_path}' ({item_count} items)", "INFO")
            
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
                                 keyword: str = "",
                                 providers_text: str = "") -> List[EmailMessage]:
        """Get messages from specified folders"""
        messages = []
        
        try:
            pythoncom.CoInitialize()
            
            # Test connection first
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            try:
                namespace.Logon()
                # Test if we can access default folder
                test_folder = namespace.GetDefaultFolder(6)  # Inbox
                test_count = test_folder.Items.Count
                self._log(f"✅ Connected to Outlook - {test_count} items in inbox", "SUCCESS")
            except Exception as e:
                self._log(f"❌ Failed to connect to Outlook: {str(e)}", "ERROR")
                raise Exception(f"Outlook connection failed: {str(e)}")
            
            # Parse provider settings
            providers = self._parse_provider_settings(providers_text)
            
            # Calculate date range
            start_date = datetime.now() - timedelta(days=days_back)
            
            self._log(f"🔍 Searching {len(folder_names)} folder(s) for messages...", "INFO")
            self._log(f"📅 Date range: {start_date.date()} to {datetime.now().date()}", "INFO")
            if keyword:
                self._log(f"🔤 Keyword filter: '{keyword}'", "INFO")
            else:
                self._log("🔤 No keyword filter - checking all subjects", "INFO")
            
            for folder_name in folder_names:
                try:
                    folder = None
                    
                    # Always get fresh folder references instead of using cached objects
                    if folder_name == "Inbox":
                        folder = namespace.GetDefaultFolder(6)
                    elif folder_name == "Sent Items":
                        folder = namespace.GetDefaultFolder(5)
                    else:
                        # Search for folder by path using fresh connection
                        folder = self._find_folder_by_path(namespace, folder_name)
                    
                    if not folder:
                        self._log(f"Folder '{folder_name}' not found", "WARNING")
                        continue
                    
                    folder_messages = folder.Items
                    folder_messages.Sort("[ReceivedTime]", True)
                    
                    folder_count = 0
                    total_checked = 0
                    date_filtered = 0
                    keyword_filtered = 0
                    provider_filtered = 0
                    
                    for message in folder_messages:
                        try:
                            total_checked += 1
                            
                            # Check date range
                            if hasattr(message, 'ReceivedTime') and message.ReceivedTime:
                                if message.ReceivedTime.date() < start_date.date():
                                    date_filtered += 1
                                    continue
                            else:
                                # Skip messages without valid date
                                continue
                            
                            # Check keyword if specified
                            if keyword and keyword.lower() not in message.Subject.lower():
                                keyword_filtered += 1
                                continue
                            
                            # Check provider filter
                            include_message, provider_name = self._should_include_message(message, providers)
                            if not include_message:
                                provider_filtered += 1
                                continue
                            
                            # Get attachments (only count actual document attachments)
                            attachments = []
                            try:
                                for attachment in message.Attachments:
                                    # Skip embedded/inline attachments
                                    if hasattr(attachment, 'Type') and attachment.Type != 1:
                                        continue
                                    
                                    filename = attachment.FileName
                                    if not filename:
                                        continue
                                    
                                    # Skip inline images and signatures
                                    skip_patterns = [
                                        'image001', 'image002', 'image003', 'image004', 'image005',
                                        'outlook-logo', 'brandimage', 'coverimage', 'openlinkicon',
                                        'signature', 'logo', 'banner'
                                    ]
                                    
                                    if any(pattern in filename.lower() for pattern in skip_patterns):
                                        continue
                                    
                                    # Only count actual document attachments
                                    doc_extensions = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', 
                                                    '.txt', '.csv', '.zip', '.rar', '.7z', '.xml', '.json']
                                    _, ext = os.path.splitext(filename)
                                    
                                    if ext.lower() in doc_extensions:
                                        attachments.append(filename)
                            except:
                                pass
                            
                            email_msg = EmailMessage(
                                subject=message.Subject,
                                sender=getattr(message, 'SenderEmailAddress', 'Unknown'),
                                received_time=message.ReceivedTime,
                                body=getattr(message, 'Body', ''),
                                attachments=attachments,
                                folder_name=folder_name,
                                provider_name=provider_name
                            )
                            
                            messages.append(email_msg)
                            folder_count += 1
                            
                        except Exception as e:
                            self._log(f"Error processing message: {e}", "WARNING")
                            continue
                    
                    self._log(f"📊 Folder '{folder_name}': {folder_count} MATCHES from {total_checked} total (❌ {date_filtered} too old, ❌ {keyword_filtered} wrong keyword, ❌ {provider_filtered} wrong sender)", "INFO")
                    
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
                          providers_text: str = "",
                          callback: Callable[[List[EmailMessage]], None] = None):
        """Get messages asynchronously"""
        def get_in_thread():
            messages = self.get_messages_from_folders(folder_names, days_back, keyword, providers_text)
            if callback:
                callback(messages)
        
        thread = threading.Thread(target=get_in_thread, daemon=True)
        thread.start()
    
    def save_attachments(self, messages: List[EmailMessage], 
                        save_folder: str,
                        file_naming_format: str = "date",
                        custom_suffix: str = "",
                        extraction_mode: str = "all",
                        conversion_enabled: bool = False,
                        convert_to_format: str = "") -> Dict[str, str]:
        """Save attachments from messages"""
        saved_files = {}
        
        if not save_folder:
            self._log("No save folder specified", "ERROR")
            return saved_files
        
        # For "latest" mode, track the latest file by base name
        latest_files = {} if extraction_mode == "latest" else None
        
        try:
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon()
            
            # Create save folder if it doesn't exist
            os.makedirs(save_folder, exist_ok=True)
            
            # Sort messages by received time (newest first) for latest mode
            if extraction_mode == "latest":
                messages = sorted(messages, key=lambda msg: msg.received_time, reverse=True)
            
            for message in messages:
                try:
                    if not message.attachments:
                        continue
                    
                    # Find the actual Outlook message object by searching in the folder
                    folder = None
                    if message.folder_name == "Inbox":
                        folder = namespace.GetDefaultFolder(6)
                    elif message.folder_name == "Sent Items":
                        folder = namespace.GetDefaultFolder(5)
                    else:
                        folder = self._find_folder_by_path(namespace, message.folder_name)
                    
                    if not folder:
                        self._log(f"Cannot find folder '{message.folder_name}' for message '{message.subject}'", "WARNING")
                        continue
                    
                    # Search for the message in the folder
                    outlook_message = None
                    for msg in folder.Items:
                        try:
                            if (msg.Subject == message.subject and 
                                msg.ReceivedTime.date() == message.received_time.date() and
                                msg.Attachments.Count > 0):
                                outlook_message = msg
                                break
                        except:
                            continue
                    
                    if not outlook_message:
                        self._log(f"Cannot find Outlook message object for '{message.subject}'", "WARNING")
                        continue
                    
                    # Generate filename suffix based on format
                    if file_naming_format == "date":
                        suffix = message.received_time.strftime("%Y-%m-%d")
                    elif file_naming_format == "year":
                        suffix = message.received_time.strftime("%Y")
                    else:  # custom
                        suffix = custom_suffix or "custom"
                    
                    # Save each attachment (filter out embedded images and inline content)
                    for i, attachment in enumerate(outlook_message.Attachments):
                        try:
                            # Skip embedded/inline attachments (usually images in email signatures/body)
                            if hasattr(attachment, 'Type'):
                                # Type 1 = olByValue (actual attachments)
                                # Type 5 = olEmbeddeditem 
                                # Type 6 = olOLE (embedded objects)
                                if attachment.Type != 1:  # Skip non-file attachments
                                    continue
                            
                            # Skip common embedded image patterns
                            original_name = attachment.FileName
                            if not original_name:
                                continue
                                
                            # Skip inline images and email signatures
                            skip_patterns = [
                                'image001', 'image002', 'image003', 'image004', 'image005',
                                'outlook-logo', 'brandimage', 'coverimage', 'openlinkicon',
                                'signature', 'logo', 'banner'
                            ]
                            
                            if any(pattern in original_name.lower() for pattern in skip_patterns):
                                self._log(f"Skipping embedded content: {original_name}", "INFO")
                                continue
                            
                            # Only process actual document attachments
                            doc_extensions = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', 
                                            '.txt', '.csv', '.zip', '.rar', '.7z', '.xml', '.json']
                            name, ext = os.path.splitext(original_name)
                            
                            if ext.lower() not in doc_extensions:
                                self._log(f"Skipping non-document file: {original_name}", "INFO")
                                continue
                            
                            # For "latest" mode, check if we already have a newer version of this file type
                            if extraction_mode == "latest":
                                base_name = name.split('_')[0] if '_' in name else name  # Get base name without suffixes
                                file_key = f"{base_name}{ext}".lower()
                                
                                if file_key in latest_files:
                                    # Skip this file as we already have a newer version
                                    self._log(f"Skipping older version: {original_name}", "INFO")
                                    continue
                                else:
                                    latest_files[file_key] = message.received_time
                            
                            # Create filename using provider name + year/date format
                            provider_prefix = message.provider_name if message.provider_name else "unknown"
                            
                            if extraction_mode == "latest":
                                # For latest mode, use provider + year/date + extension
                                if file_naming_format == "year":
                                    final_filename = f"{provider_prefix}{suffix}{ext}"  # e.g., test2025.pdf
                                elif file_naming_format == "date":
                                    final_filename = f"{provider_prefix}{suffix}{ext}"  # e.g., test2025-08-15.pdf
                                else:  # original or custom
                                    final_filename = f"{provider_prefix}_{original_name}"
                            else:
                                # For all mode, use provider + naming format
                                if file_naming_format == "original":
                                    final_filename = f"{provider_prefix}_{original_name}"
                                elif file_naming_format == "year":
                                    final_filename = f"{provider_prefix}{suffix}{ext}"  # e.g., test2025.pdf
                                elif file_naming_format == "date":
                                    final_filename = f"{provider_prefix}{suffix}{ext}"  # e.g., test2025-08-15.pdf
                                else:  # custom
                                    final_filename = f"{provider_prefix}{custom_suffix}{ext}"
                                
                                # Handle duplicate filenames for "all" mode
                                counter = 1
                                full_path = os.path.join(save_folder, final_filename)
                                
                                while os.path.exists(full_path):
                                    name_part, ext_part = os.path.splitext(final_filename)
                                    final_filename = f"{name_part}_{counter}{ext_part}"
                                    full_path = os.path.join(save_folder, final_filename)
                                    counter += 1
                            
                            # Final path
                            full_path = os.path.join(save_folder, final_filename)
                            
                            # Save the attachment
                            attachment.SaveAsFile(full_path)
                            
                            # Apply file conversion if enabled
                            final_saved_path = full_path
                            if conversion_enabled and convert_to_format:
                                try:
                                    processor = EmailProcessor()
                                    converted_path = processor.convert_file_format(full_path, convert_to_format)
                                    if converted_path and converted_path != full_path:
                                        final_saved_path = converted_path
                                        # Update the final filename to reflect the new extension
                                        final_filename = os.path.basename(converted_path)
                                        self._log(f"Converted to {convert_to_format.upper()}: {final_filename}", "INFO")
                                except Exception as e:
                                    self._log(f"Conversion failed for {final_filename}: {e}", "WARNING")
                            
                            saved_files[final_filename] = final_saved_path
                            
                            mode_text = "latest" if extraction_mode == "latest" else "all"
                            self._log(f"Saved attachment ({mode_text} mode): {final_filename}", "SUCCESS")
                            
                        except Exception as e:
                            self._log(f"Error saving attachment '{attachment.FileName}': {e}", "ERROR")
                            continue
                    
                except Exception as e:
                    self._log(f"Error processing message '{message.subject}': {e}", "ERROR")
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