<div align="center">

### 👇

[![Download My Installation Script](https://img.shields.io/badge/Download_My_Script-brightgreen?style=for-the-badge&logo=powershell&logoColor=white)](https://github.com/EXLOUD/Office-Telemetry-Disabler/archive/refs/heads/main.zip)

---

# Office Privacy and Telemetry Disabler

### 👀 Repository Views

<img alt="count" src="https://count.getloli.com/get/@:EXLOUD-Office-Telemetry-Disable?theme=rule34" />

**⭐ If this tool helped you, please consider giving it a star! ⭐**

---

<img src="assets/preview.gif" alt="Office Privacy and Telemetry Disabler Logo" width="600" height="400">

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5%2B-blue.svg)
![Windows](https://img.shields.io/badge/Windows-7%2F8%2F10%2F11-blue.svg)
![Office](https://img.shields.io/badge/Office-2010--2024-orange.svg)

A comprehensive PowerShell script to disable Microsoft Office logging, telemetry, and privacy features across all Office versions (2010-2024) and Windows versions (7-11).

</div>

## 🚀 Features

- **Multi-version support**: Works with Office 2010, 2013, 2016, 2019, 2021, and 2024
- **Cross-Windows compatibility**: Supports Windows 7, 8, 10, and 11
- **Automatic version detection**: Launcher automatically detects Windows version and runs appropriate script
- **Comprehensive privacy protection**: Disables logging, telemetry, and data collection
- **Scheduled task management**: Disables Office telemetry and update tasks
- **Hosts file blocking**: Optional blocking of Microsoft telemetry servers
- **User-friendly interface**: Colored output with clear status indicators
- **Safe execution**: Backup creation and error handling

## 📋 What it disables

### Office Logging & Telemetry
- Microsoft Office application logging
- Client telemetry collection
- Verbose logging features
- OSM (Office Service Manager) logging and uploads

### Privacy Features
- Customer Experience Improvement Program (CEIP)
- Feedback collection
- Connected Experiences (Office 2016+)
- Online content downloads
- Watson error reporting

### Scheduled Tasks
- Office telemetry agents
- Subscription heartbeat tasks
- Background task handlers
- Automatic update tasks

### Network Communication
- Blocks telemetry hosts via hosts file (optional)
- Disables automatic updates
- Prevents data uploads to Microsoft servers

## 🛠️ Installation & Usage

### Method 1: Using the Launcher (Recommended)
1. Download the launcher (`Launcher.bat`) and both PowerShell scripts:
   - `office_privacy_telemetry_disabler.ps1` (for Windows 10/11)
   - `office_privacy_telemetry_disabler_win7+.ps1` (for Windows 7/8/10/11)
2. Place all files in the same directory
3. Right-click on `Launcher.bat` and select "Run as administrator"
4. The launcher will automatically detect your Windows version and run the appropriate script
5. Follow the on-screen prompts

### Method 2: Direct PowerShell Execution
1. Download the appropriate PowerShell script for your Windows version:
   - For Windows 10/11: `office_privacy_telemetry_disabler.ps1`
   - For Windows 7/8/10/11: `office_privacy_telemetry_disabler_win7+.ps1`
2. Open PowerShell as Administrator
3. Run: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process`
4. Execute the downloaded script: `.\office_privacy_telemetry_disabler.ps1` or `.\office_privacy_telemetry_disabler_win7+.ps1`

## 🎯 Supported Versions


### Office Versions
| Version | Year | Status |
|---------|------|--------|
| 14.0 | Office 2010 | ✅ Supported |
| 15.0 | Office 2013 | ✅ Supported |
| 16.0 | Office 2016/2019 | ✅ Supported |
| 17.0 | Office 2021 | ✅ Supported |
| 18.0 | Office 2024 | ✅ Supported |


### Windows Versions

| Version | Status | Script Used |
|---------|--------|-------------|
| Windows 7 | ✅ Supported | `office_privacy_telemetry_disabler_win7+.ps1` |
| Windows 8/8.1 | ✅ Supported | `office_privacy_telemetry_disabler_win7+.ps1` |
| Windows 10 | ✅ Supported | `office_privacy_telemetry_disabler.ps1` but you can run both |
| Windows 11 | ✅ Supported | `office_privacy_telemetry_disabler.ps1` but you can run both |


## 🔧 Requirements

- **Operating System**: Windows 7, 8, 10, or 11
- **PowerShell**: Version 5.1 or higher
- **Privileges**: Administrator rights required for registry and scheduled task modifications
- **Office**: Any version from 2010 to 2024

## 🤖 Automatic Version Detection

The `Launcher.bat` includes intelligent Windows version detection:
- Automatically identifies your Windows version
- Selects the most compatible PowerShell script
- Ensures optimal performance across all Windows versions
- Provides fallback options for maximum compatibility

## 📸 Screenshots

### Main Interface
The script provides a clean, colored interface showing:
- Windows and Office version detection
- Registry modifications
- Scheduled task management
- Progress indicators

### Output Legend
- ✅ **Green**: Successfully completed actions
- 🔄 **Magenta**: Settings changed
- ℹ️ **Blue**: Information messages
- ⚠️ **Yellow**: Warnings
- ❌ **Red**: Errors
- ➡️ **Gray**: Items not found/skipped

## 🛡️ Safety Features

- **Backup creation**: Automatic backup of hosts file before modification
- **Error handling**: Comprehensive error catching and reporting
- **Registry validation**: Checks for existing registry paths before modification
- **Reversible changes**: Most changes can be reversed manually if needed
- **Non-destructive**: Only modifies privacy-related settings
- **Version-specific optimization**: Different scripts optimized for different Windows versions

## 🚨 Important Notes

1. **Administrator Rights**: Required for modifying system registry and scheduled tasks
2. **Office Restart**: Some changes require restarting Office applications
3. **Windows Defender**: The script temporarily adds hosts file to exclusions
4. **Backup**: Always backup your system before running system modification scripts
5. **Version Compatibility**: Use the launcher for automatic version detection and optimal compatibility

## 🔄 What happens after running?

After successful execution:
- Office telemetry and logging are disabled
- Privacy-invasive features are turned off
- Scheduled telemetry tasks are disabled
- (Optional) Telemetry hosts are blocked
- Office applications may need to be restarted
- Changes are optimized for your specific Windows version

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development
- The script uses PowerShell with colored output
- Registry modifications use proper error handling
- Scheduled task management includes comprehensive logging
- Code is modular and well-documented
- Version-specific optimizations for different Windows versions

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This script modifies Windows registry and system settings. While designed to be safe, use at your own risk. Always backup your system before running. The authors are not responsible for any damage or data loss.

## 🙏 Acknowledgments

- Microsoft for providing comprehensive Office documentation
- PowerShell community for best practices
- Privacy advocates for highlighting the importance of telemetry control

## 📞 Support

If you encounter any issues:
1. Check the [Issues](../../issues) section
2. Ensure you're running as Administrator
3. Verify your Office version is supported
4. Check your Windows version compatibility
5. Use the launcher for automatic version detection
6. Check the console output for specific error messages

---

<div align="center">

**Made with ❤️ by [EXLOUD](https://github.com/EXLOUD)**

*Protecting your privacy across all Windows versions, one script at a time.*

</div>
