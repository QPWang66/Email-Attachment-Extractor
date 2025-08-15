# Email Attachment Extractor

> **Automated email attachment extraction tool for Microsoft Outlook**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Windows](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

A desktop application that automates the extraction of document attachments from Microsoft Outlook emails. Features smart filtering by sender, keyword search, and organized file management with a modern, user-friendly interface.

**Platform**: Windows with Microsoft Outlook (COM automation). Contributions welcome for additional platforms and email clients.

## Features

- **Multi-Folder Search**: Discover and search across all Outlook folders including shared mailboxes
- **Smart Filtering**: Filter emails by keywords, date ranges, and sender providers
- **Extraction Modes**: Extract all files with dates, or latest files only per type
- **Provider-Based Naming**: Automatically name files using sender-based prefixes (e.g., `client2025.pdf`)
- **Flexible Date Ranges**: Search emails from last 7 days, 30 days, or custom timeframes
- **Persistent Settings**: Save configurations including folder selections and provider mappings
- **Document Focus**: Extracts only actual documents (PDF, Office files) while filtering out embedded images
- **Real-time Progress**: Live activity logging with detailed extraction statistics
- **One-Click Operation**: Save settings and run extraction in a single action

## Requirements

- **Operating System**: Windows 10/11
- **Email Client**: Microsoft Outlook (classic desktop version, installed and configured)
- **Runtime**: Python 3.8+ (for development) or standalone executable

## Quick Start

### Option 1: Standalone Executable
1. Download the latest release from [Releases](../../releases)
2. Extract and run `EmailAttachmentExtractor.exe`
3. Configure settings and start extracting attachments

### Option 2: Python Environment
```bash
# Clone the repository
git clone https://github.com/your-username/Email-Attachment-Extractor.git
cd Email-Attachment-Extractor

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

## Building from Source

### Create Executable
```bash
# Install build tools
pip install cx_Freeze

# Build executable
python setup.py build

# Output will be in build/ directory
```

### Alternative with PyInstaller
```bash
# Install PyInstaller
pip install pyinstaller

# Create standalone executable
pyinstaller --windowed --onefile --name EmailAttachmentExtractor main.py
```

## Project Structure

```
Email-Attachment-Extractor/
├── main.py                     # Application entry point
├── src/
│   ├── core/                   # Core business logic
│   │   ├── config_manager.py   # Configuration management
│   │   └── outlook_manager.py  # Outlook integration & COM handling
│   ├── ui/                     # User interface components
│   │   ├── main_window.py      # Main application window
│   │   ├── main_tab.py         # Extraction interface
│   │   ├── settings_tab.py     # Configuration interface
│   │   └── styles.py           # UI theming and styling
│   └── utils/                  # Utility modules
│       ├── email_processor.py  # Email processing & filtering
│       └── file_manager.py     # File operations & organization
├── requirements.txt            # Python dependencies
├── setup.py                   # Build configuration
└── README.md                  # This documentation
```

## Configuration

Settings are automatically saved in `email_extractor_config.json`:

```json
{
    "keyword": "report",
    "folder": "C:/Users/Username/Documents/ExtractedAttachments",
    "days": 7,
    "selected_folders": ["Inbox", "Reports", "Archive"],
    "naming_format": "date",
    "providers": "# Email provider configurations...",
    "auto_run": false
}
```

### Provider Configuration
Configure email providers in the Settings tab using this format:
```
# Email-based rules
reports@company.com = Company Reports
noreply@service.com = Service Provider

# Subject-based rules
invoice = Invoice Documents
statement = Financial Statements
report = General Reports
```

## Usage Guide

### Basic Workflow
1. **Initial Setup**
   - Run the application and navigate to Settings tab
   - Click "Discover Folders" to detect all Outlook folders
   - Select which folders to monitor

2. **Configure Extraction**
   - Set search keywords (e.g., "report", "invoice", "statement")
   - Choose date range for email scanning
   - Select output folder for extracted attachments

3. **Run Extraction**
   - Return to main tab and click "Extract Reports Now"
   - Monitor progress in real-time activity log
   - Review extracted files in the configured output folder

### Advanced Features

#### File Organization Options
- **Date-based**: `2024/01-January/Provider_Report_20240115.pdf`
- **Provider-based**: `Company_Reports/Invoice_20240115.pdf`
- **Category-based**: `Financial/Company_Reports_20240115.pdf`

#### Automated Processing
- Enable auto-run for scheduled processing
- Configure provider-specific rules for smart categorization
- Set up backup workflows for important documents

## Architecture

### Modular Design
- **Separation of Concerns**: Clean separation between UI, business logic, and utilities
- **Extensible Framework**: Easy to add new features and functionality
- **Testable Components**: Individual modules can be tested in isolation
- **Professional Standards**: Follows enterprise software development practices

### Key Components
- **ConfigManager**: Handles all configuration persistence and validation
- **OutlookManager**: Manages COM interactions with Microsoft Outlook
- **EmailProcessor**: Advanced email filtering and categorization
- **FileManager**: Handles file operations, organization, and backup

## Platform Support & Contributions

### Current Support
- **Windows**: Full support for classic Microsoft Outlook desktop application
- **Outlook Versions**: Tested with Outlook 2016, 2019, and Microsoft 365 desktop

### Seeking Contributions For
We actively welcome contributions to extend platform and application support:

#### New Outlook Versions
- **Outlook for Web**: Web-based Outlook interface support
- **New Outlook for Windows**: Modern Outlook app support
- **Microsoft Graph API**: Cloud-based email access

#### macOS Support
- **Microsoft Outlook for Mac**: Native macOS Outlook integration
- **Mail.app Integration**: Apple Mail application support
- **Cross-platform Architecture**: IMAP/Exchange connectivity

#### Additional Platforms
- **Linux**: Email client integration for Linux distributions
- **Alternative Clients**: Thunderbird, Apple Mail, etc.

### How to Contribute
If you're interested in adding support for new platforms or Outlook versions:
1. Check existing issues for platform-specific requests
2. Create a new issue describing your intended contribution
3. Fork the repository and create a feature branch
4. Implement changes following the existing modular architecture
5. Submit a pull request with comprehensive testing

## Troubleshooting

### Common Issues

**Connection Failed**
- Ensure Microsoft Outlook is installed and running
- Try opening Outlook manually before running the extractor
- Check Windows security settings and firewall

**No Folders Found**
- Verify Outlook is properly configured with email accounts
- Run the application as administrator if needed
- Restart Outlook and try again

**Permission Errors**
- Run as administrator
- Check folder permissions for output directory
- Verify antivirus software isn't blocking the application

### Debug Information
Enable detailed logging in the activity log to diagnose issues:
- Connection status and COM object initialization
- Folder discovery progress and results
- Email processing statistics and errors
- File operation status and conflicts

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

### Development Setup
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes following the existing code style
4. Test thoroughly with the modular architecture
5. Commit changes (`git commit -m 'Add amazing feature'`)
6. Push to branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Code Standards
- Follow PEP 8 style guidelines
- Add type hints for new functions
- Include docstrings for classes and methods
- Test with both real Outlook and mock objects
- Maintain backward compatibility

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Roadmap

- [ ] **New Outlook Support**: Modern Outlook for Windows integration
- [ ] **macOS Support**: Native macOS Outlook and Mail.app integration
- [ ] **Cross-platform Email**: IMAP/POP3 server connectivity
- [ ] **Microsoft Graph API**: Cloud-based Office 365 integration
- [ ] **Advanced Filtering**: Rule-based email processing engine
- [ ] **Scheduled Tasks**: Automated extraction workflows
- [ ] **Cloud Storage**: Integration with OneDrive, Google Drive, etc.
- [ ] **Machine Learning**: AI-powered email categorization
- [ ] **REST API**: Automation and integration endpoints

## Support

- **Documentation**: Check this README and inline code documentation
- **Bug Reports**: Open an issue with detailed information
- **Feature Requests**: Submit enhancement suggestions via issues
- **Platform Requests**: Request support for new platforms or email clients
- **Technical Support**: Review troubleshooting section first

---

<div align="center">

**Built for email automation professionals**

[Star this repo](../../stargazers) | [Report bug](../../issues) | [Request feature](../../issues)

</div>