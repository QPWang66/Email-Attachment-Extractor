"""
Email Processing Utilities
Advanced email processing and filtering capabilities.
"""

import os
import re
from datetime import datetime
from typing import List, Dict, Tuple, Optional
from pathlib import Path

from ..core.outlook_manager import EmailMessage


class EmailProcessor:
    """Advanced email processing and filtering"""
    
    def __init__(self, providers_config: str = ""):
        """Initialize email processor"""
        self.providers = self._parse_providers_config(providers_config)
    
    def _parse_providers_config(self, config_text: str) -> Dict[str, str]:
        """Parse providers configuration text"""
        providers = {}
        
        for line in config_text.split('\n'):
            line = line.strip()
            if line and not line.startswith('#'):
                if '=' in line:
                    key, value = line.split('=', 1)
                    providers[key.strip()] = value.strip()
        
        return providers
    
    def filter_messages_by_providers(self, messages: List[EmailMessage]) -> Dict[str, List[EmailMessage]]:
        """Filter and group messages by providers"""
        grouped = {}
        
        for message in messages:
            provider = self.identify_provider(message)
            if provider not in grouped:
                grouped[provider] = []
            grouped[provider].append(message)
        
        return grouped
    
    def identify_provider(self, message: EmailMessage) -> str:
        """Identify the provider for a message"""
        # Check sender email first
        for pattern, provider in self.providers.items():
            if '@' in pattern:  # Email pattern
                if pattern.lower() in message.sender.lower():
                    return provider
            else:  # Subject keyword pattern
                if pattern.lower() in message.subject.lower():
                    return provider
        
        # Default provider based on sender domain
        if '@' in message.sender:
            domain = message.sender.split('@')[-1]
            return f"Provider ({domain})"
        
        return "Unknown Provider"
    
    def get_latest_message_per_provider(self, messages: List[EmailMessage]) -> Dict[str, EmailMessage]:
        """Get the latest message for each provider"""
        grouped = self.filter_messages_by_providers(messages)
        latest = {}
        
        for provider, provider_messages in grouped.items():
            if provider_messages:
                # Sort by received time and get the latest
                latest_message = max(provider_messages, key=lambda m: m.received_time)
                latest[provider] = latest_message
        
        return latest
    
    def filter_messages_with_attachments(self, messages: List[EmailMessage]) -> List[EmailMessage]:
        """Filter messages that have attachments"""
        return [msg for msg in messages if msg.attachments]
    
    def get_attachment_statistics(self, messages: List[EmailMessage]) -> Dict[str, int]:
        """Get statistics about attachments"""
        stats = {
            'total_messages': len(messages),
            'messages_with_attachments': 0,
            'total_attachments': 0,
            'file_types': {}
        }
        
        for message in messages:
            if message.attachments:
                stats['messages_with_attachments'] += 1
                stats['total_attachments'] += len(message.attachments)
                
                for attachment in message.attachments:
                    ext = Path(attachment).suffix.lower()
                    if ext:
                        stats['file_types'][ext] = stats['file_types'].get(ext, 0) + 1
        
        return stats
    
    def generate_filename(self, message: EmailMessage, provider: str, 
                         naming_format: str = "date", custom_suffix: str = "") -> str:
        """Generate filename for saved attachment"""
        # Clean provider name for filename
        clean_provider = re.sub(r'[^\w\s-]', '', provider).strip()
        clean_provider = re.sub(r'[-\s]+', '_', clean_provider)
        
        # Generate suffix based on format
        if naming_format == "date":
            suffix = datetime.now().strftime("%Y-%m-%d")
        elif naming_format == "year":
            suffix = datetime.now().strftime("%Y")
        elif naming_format == "custom" and custom_suffix:
            suffix = custom_suffix
        else:
            suffix = "report"
        
        # Include message date for uniqueness
        msg_date = message.received_time.strftime("%Y%m%d")
        
        return f"{clean_provider}_{suffix}_{msg_date}"
    
    def should_save_attachment(self, filename: str, allowed_extensions: List[str] = None) -> bool:
        """Determine if an attachment should be saved"""
        if allowed_extensions is None:
            # Default allowed extensions for reports
            allowed_extensions = ['.pdf', '.xlsx', '.xls', '.docx', '.doc', '.csv', '.txt']
        
        ext = Path(filename).suffix.lower()
        return ext in allowed_extensions
    
    def clean_filename(self, filename: str) -> str:
        """Clean filename for safe saving"""
        # Remove or replace invalid characters
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # Remove extra spaces and trim
        filename = re.sub(r'\s+', ' ', filename).strip()
        
        # Ensure it's not too long
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
        
        return filename


class ReportAnalyzer:
    """Analyze and categorize report types"""
    
    REPORT_PATTERNS = {
        'financial': [
            r'financial\s+report', r'balance\s+sheet', r'income\s+statement',
            r'profit\s+loss', r'cash\s+flow', r'revenue', r'earnings'
        ],
        'sales': [
            r'sales\s+report', r'revenue\s+report', r'performance\s+report',
            r'sales\s+summary', r'quarterly\s+sales'
        ],
        'analytics': [
            r'analytics\s+report', r'data\s+report', r'metrics\s+report',
            r'kpi\s+report', r'dashboard', r'statistics'
        ],
        'operational': [
            r'operational\s+report', r'operations\s+report', r'status\s+report',
            r'weekly\s+report', r'monthly\s+report', r'daily\s+report'
        ],
        'compliance': [
            r'compliance\s+report', r'audit\s+report', r'regulatory\s+report',
            r'security\s+report', r'risk\s+report'
        ]
    }
    
    def categorize_message(self, message: EmailMessage) -> str:
        """Categorize a message based on subject and content"""
        text = f"{message.subject} {message.body}".lower()
        
        for category, patterns in self.REPORT_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    return category
        
        return 'general'
    
    def get_report_summary(self, messages: List[EmailMessage]) -> Dict[str, int]:
        """Get summary of report types"""
        summary = {}
        
        for message in messages:
            category = self.categorize_message(message)
            summary[category] = summary.get(category, 0) + 1
        
        return summary