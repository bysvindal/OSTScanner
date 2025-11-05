# Outlook OST/PST File Scanner

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![.NET](https://img.shields.io/badge/.NET-6.0%20%7C%208.0-512BD4)
![Platform](https://img.shields.io/badge/platform-Windows-0078D4)

A low-level PST/OST file corruption detector and validator for Windows environments. Performs deep structural validation based on Microsoft's PST file format specification.

---

## ⚠️ DISCLAIMER

**USE AT YOUR OWN RISK**

This software is provided "AS IS" without warranty of any kind, either expressed or implied. The author(s) and contributors:

- **Take NO responsibility** for any data loss, corruption, or damage resulting from the use of this software
- **Do NOT guarantee** the accuracy of corruption detection or file integrity validation
- **Are NOT liable** for any consequences of misuse, improper usage, or software defects
- **Strongly recommend** creating full backups before running any file validation or deletion operations

**THIS TOOL CAN DELETE FILES.** Always ensure you have proper backups before using the `--delete` option.

By using this software, you acknowledge that:
- You understand the risks involved in file manipulation operations
- You have adequate backups of your data
- You accept full responsibility for any outcomes resulting from the use of this tool
- The software is intended for IT professionals and advanced users who understand file system operations

**The author(s) are not responsible for:**
- Lost or corrupted data
- System downtime or disruptions
- Misuse or unauthorized use of this tool
- Any direct, indirect, incidental, or consequential damages

---

## 🎯 Features

- **Deep Structural Validation**: Validates PST/OST file headers, CRC checksums, and internal structures
- **Low-Level Binary Parsing**: Direct analysis of file format based on MS-PST specification
- **Corruption Detection**: Identifies corrupted files through multiple validation checks
- **Safe Deletion**: Creates timestamped backups before removing corrupted files
- **Multi-Format Support**: Works with both ANSI and Unicode PST/OST formats
- **Automatic Discovery**: Scans common Outlook data file locations
- **Windows 365 Compatible**: Designed for cloud-based Windows environments

---

## 📋 Requirements

- **Operating System**: Windows 10, Windows 11
- **.NET Runtime**: .NET Runtime 8.0.19 or higher
- **Architecture**: x64 or ARM64
- **Permissions**: User-level access to Outlook data folders

---

## 📦 Installation

### Option 1: Build from Source

```powershell
# Clone the repository
git clone https://github.com/yourusername/OutlookOSTScanner.git
cd OutlookOSTScanner

# Publish self-contained executable
dotnet publish -c Release -f net8.0 -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish/win-x64
dotnet publish -c Release -f net8.0 -r win-arm64 --self-contained true -p:PublishSingleFile=true -o ./publish/win-arm64
```

---

## 🚀 Usage

### Basic Commands

```powershell
# Scan all OST/PST files in default Outlook locations
OutlookOSTScanner.exe

# Show help and all available options
OutlookOSTScanner.exe --help

# Scan a specific file
OutlookOSTScanner.exe --file "C:\Users\username\AppData\Local\Microsoft\Outlook\email@domain.ost"

# Scan and delete corrupted files (creates .bak backup)
OutlookOSTScanner.exe --delete

# Permanently delete without creating backup (⚠️ DANGEROUS)
OutlookOSTScanner.exe --delete --no-backup
```

### Command-Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--help` | `-h` | Display help information |
| `--delete` | `-d` | Delete corrupted files (creates backup by default) |
| `--file <path>` | `-f` | Scan specific file instead of auto-discovery |
| `--no-backup` | | Permanently delete without creating backup ⚠️ |

---

## 📊 Output Examples

### Healthy File
```
╔══════════════════════════════════════════════════════════╗
║ Outlook OST/PST Low-Level File Scanner                  ║
║ Deep Structure Validation                                ║
╚══════════════════════════════════════════════════════════╝

Found 1 file(s) to scan

────────────────────────────────────────────────────────────
Scanning: user@domain.ost
Path: C:\Users\username\AppData\Local\Microsoft\Outlook\user@domain.ost
Size: 2.45 GB
Modified: 11/5/2025 10:30:15 AM

✓ HEALTHY - File validation passed

════════════════════════════════════════════════════════════
Summary:
  Scanned: 1 file(s)
  Corrupted: 0 file(s)
  Healthy: 1 file(s)
════════════════════════════════════════════════════════════
```

### Corrupted File
```
────────────────────────────────────────────────────────────
Scanning: corrupted.ost
Path: C:\Users\username\AppData\Local\Microsoft\Outlook\corrupted.ost
Size: 1.23 GB
Modified: 11/4/2025 3:15:22 PM

✗ CORRUPTED - File validation failed

Errors:
  • Header CRC validation failed
  • Invalid root structure detected

Use --delete to remove this corrupted file
```

---

## ⚙️ How It Works

The scanner performs the following validation steps:

1. **File Accessibility Check**: Verifies file exists and is not locked
2. **Header Validation**: Checks magic signature (`!BDN`), version, and structure
3. **CRC32 Verification**: Validates header checksum integrity
4. **Format Detection**: Identifies ANSI vs Unicode PST format
5. **Root Structure Analysis**: Validates file metadata and pointers
6. **B-Tree Integrity**: Checks Node BTree (NBT) and Block BTree (BBT) structures
7. **Allocation Map Validation**: Verifies internal space allocation structures

Based on: [MS-PST: Outlook Personal Folders (.pst) File Format](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst)

---

## 🔍 Default Scan Locations

The scanner automatically searches these locations:

- `C:\Users\<username>\AppData\Local\Microsoft\Outlook\` (Primary OST location)
- `C:\Users\<username>\AppData\Roaming\Microsoft\Outlook\` (PST files)
- `C:\Users\<username>\Documents\Outlook Files\` (User PST files)

---

## ⚠️ Important Warnings

### Before Using `--delete`:

1. ✅ **ALWAYS create a full backup** of your Outlook data
2. ✅ **Close Outlook completely** before scanning
3. ✅ **Test on non-production data first**
4. ✅ **Verify you have Exchange/Microsoft 365 sync** (OST files can be regenerated)
5. ❌ **NEVER use `--no-backup`** unless you're absolutely certain

### OST vs PST Files:

- **OST files** (Offline Storage): Cached copies of Exchange mailboxes - can be safely deleted and will regenerate
- **PST files** (Personal Storage): Standalone data files - deleting these means **permanent data loss** if no backup exists

### When NOT to Use This Tool:

- ❌ On PST files without backups
- ❌ On production systems without testing
- ❌ When Outlook is running
- ❌ Without understanding the consequences
- ❌ On files you don't own or have permission to modify

---

## 🐛 Known Limitations

- Cannot repair corrupted files (detection only)
- Requires Outlook to be closed for accurate validation
- B-Tree and allocation map validation is simplified
- Does not validate email content or folder structures
- Large files (>10GB) may take several minutes to validate

---

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License - see below:

```
MIT License

Copyright (c) 2025 [Your Name]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 🔗 References

- [MS-PST File Format Specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst)
- [Microsoft Outlook Documentation](https://support.microsoft.com/outlook)
- [ScanPST.exe (Inbox Repair Tool)](https://support.microsoft.com/office/repair-outlook-data-files-pst-and-ost-25663bc3-11ec-4412-86c4-60458afc5253)

---

## 📧 Support

**This is provided as-is with no official support.**

For issues or questions:
- Check existing [Issues](https://github.com/yourusername/OutlookOSTScanner/issues)
- Create a new issue with detailed information
- Include log output and file information (redact sensitive data)

---

## 🙏 Acknowledgments

- Based on Microsoft's PST file format specification (MS-PST)
- Thanks to the open-source community for .NET tools and libraries

---

**Remember: Always backup your data before using any file manipulation tools!**
