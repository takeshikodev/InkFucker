# InkFucker

![VBS Script](https://img.shields.io/badge/Language-VBScript-8A2BE2) 
![Windows Support](https://img.shields.io/badge/Platform-Windows-0078D6)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/takeshikodev/InkFucker)
![Downloads](https://img.shields.io/github/downloads/takeshikodev/InkFucker/total)

Automated tool to recover virus-hidden files from USB drives

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Functionality](#functionality)
- [Limitations](#limitations)
- [Security](#security)
- [License](#license)

## Features

- 🔍 Automatic detection of virus-hidden files
- ♻️ Recovery from common malware folders (`_` and `WindowsServices`)
- 🗑️ Malware cleanup:
  - Fake shortcuts (.lnk)
  - Suspicious executables
  - Hidden system folders
- 📄 Detailed recovery report generation
- 📊 Visual progress tracking
- 🌐 Dual-language support (English/Russian)
- 🛡️ Safe original file deletion after recovery

## Requirements

- Windows 7 or newer
- .NET Framework 3.5+
- Enabled VBScript execution
- Administrator privileges (recommended)

## Installation Options

1. Visit [Releases Page](https://github.com/takeshikodev/InkFucker/releases)
2. Download `InkFucker.vbs`
3. Right-click → Properties → Unblock file (if needed)
4. Open file / Right-click -> Open with Command Prompt

## Usage

1. **Run script** by double-clicking
2. Enter drive letter (e.g., `F`)
3. Choose operations:
   - Quick recovery
   - Full drive scan
   - Malware removal
4. Post-recovery actions:
   - Open `Recovered` folder
   - View `RecoveryReport.txt`

**Sample Workflow:**

1. Enter drive letter: F
2. Found 142 hidden items
3. Recovering... [||||||||||--------] 65%
4. Removed 3 malicious files
5. Report saved to F:\RecoveryReport.txt

## Functionality

### Core Components

| Component           | Description                              |
|---------------------|------------------------------------------|
| `CountItems`        | Recursive item counting                  |
| `RecoverItems`      | File/folder restoration                  |
| `RemoveMalicious`   | Malware cleanup                          |
| `FullScan`          | Deep hidden file search                  |
| `ProgressUI`        | Visual progress display                  |

### Supported Formats
- All file types (including system/hidden)
- Full directory structure preservation
- Detailed operation logging

## Limitations

- ❗ Recovers hidden files only (not deleted)
- ❗ No encrypted file support
- ❗ Manual drive letter input required
- ❗ Network drives unsupported

## Security

- **Precautions**:
  - Scan drive with antivirus first
  - Maintain backup copies
  - Disable autorun.inf

- **Implementation**:
  - No internet connection required
  - Zero registry modifications
  - 100% local execution

## License

Released under [MIT License](LICENSE). Attribution appreciated.
