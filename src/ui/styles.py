"""
Style Manager for UI Components
Centralizes all styling and theming for the application.
"""

from tkinter import ttk
from typing import Dict


class StyleManager:
    """Manages application styles and themes"""
    
    def __init__(self, theme: str = "default"):
        """Initialize style manager"""
        self.theme = theme
        self.colors = self._get_color_scheme()
    
    def _get_color_scheme(self) -> Dict[str, str]:
        """Get color scheme based on theme"""
        if self.theme == "dark":
            return {
                'bg': '#2b2b2b',
                'card': '#3c3c3c',
                'primary': '#4a9eff',
                'success': '#5cb85c',
                'danger': '#d9534f',
                'warning': '#f0ad4e',
                'text': '#ffffff',
                'text_secondary': '#cccccc',
                'border': '#555555'
            }
        else:  # default light theme
            return {
                'bg': '#f8f9fa',
                'card': '#ffffff',
                'primary': '#0066cc',
                'success': '#28a745',
                'danger': '#dc3545',
                'warning': '#ffc107',
                'text': '#212529',
                'text_secondary': '#6c757d',
                'border': '#dee2e6'
            }
    
    def get_colors(self) -> Dict[str, str]:
        """Get current color scheme"""
        return self.colors
    
    def setup_styles(self):
        """Setup ttk styles"""
        style = ttk.Style()
        
        # Use appropriate theme
        if self.theme == "dark":
            style.theme_use('alt')
        else:
            style.theme_use('clam')
        
        # Configure custom styles
        style.configure('Card.TFrame', 
                       background=self.colors['card'], 
                       relief='flat', 
                       borderwidth=1)
        
        style.configure('Heading.TLabel', 
                       font=('Segoe UI', 18, 'bold'), 
                       background=self.colors['bg'],
                       foreground=self.colors['text'])
        
        style.configure('Subheading.TLabel', 
                       font=('Segoe UI', 12, 'bold'), 
                       background=self.colors['card'],
                       foreground=self.colors['text'])
        
        style.configure('Body.TLabel', 
                       font=('Segoe UI', 10), 
                       background=self.colors['card'],
                       foreground=self.colors['text'])
        
        style.configure('Secondary.TLabel',
                       font=('Segoe UI', 9),
                       background=self.colors['card'],
                       foreground=self.colors['text_secondary'])
        
        style.configure('Primary.TButton', 
                       font=('Segoe UI', 10, 'bold'))
        
        style.configure('Secondary.TButton', 
                       font=('Segoe UI', 10))
        
        style.configure('Success.TLabel', 
                       foreground=self.colors['success'], 
                       background=self.colors['card'])
        
        style.configure('Warning.TLabel', 
                       foreground=self.colors['warning'], 
                       background=self.colors['card'])
        
        style.configure('Danger.TLabel',
                       foreground=self.colors['danger'],
                       background=self.colors['card'])
    
    def get_log_colors(self) -> Dict[str, str]:
        """Get colors for log messages"""
        return {
            "INFO": self.colors['text'],
            "SUCCESS": self.colors['success'],
            "WARNING": self.colors['warning'],
            "ERROR": self.colors['danger']
        }
    
    def update_theme(self, theme: str):
        """Update the current theme"""
        self.theme = theme
        self.colors = self._get_color_scheme()
        self.setup_styles()