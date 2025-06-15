# ClipSon
Nextcloud clipboard synchronization tool for Linux (Python) and Windows (PowerShell)

**ClipSon works over Nextcloud's WebDAV protocol, supporting proxy configurations and secure connections to synchronize your clipboard content across devices.**

## Features

- **Nextcloud WebDAV protocol integration** - Uses Nextcloud's WebDAV API for reliable file transfers
- **Proxy support** - Works through corporate proxies and network configurations
- **Cross-platform clipboard monitoring** - Works on both Windows and Linux
- **Real-time synchronization** - Automatically monitors clipboard changes
- **Text and image support** - Handles both text content and images
- **Multi-MIME clipboard support** - Supports rich text formats (HTML, RTF) with copyq on Linux
- **Duplicate prevention** - Avoids saving identical clipboard content
- **Multi-device synchronization** - All instances must use the same Nextcloud account to see each other's updates

## Usage

0. copy the `config_template.json` file to `config.json` and edit it to set your Nextcloud WebDAV URL, username, and password, proxy, etc.

**Important:** All ClipSon instances across different devices (Windows, Linux, etc.) must be configured with the same Nextcloud account credentials to synchronize clipboard content between them. Each device will create its own clipboard file (e.g., `clipboard-hostname.txt`) in the shared Nextcloud folder, allowing all instances to monitor and sync changes from other devices.

### Windows (PowerShell)

1. Run the PowerShell script:
   ```powershell
   .\clipson.ps1
   ```
2. Copy any text or image to clipboard
3. Content will be automatically synced to Nextcloud
4. Peers' clipboard syncs are regularly checked (parameter `remote_check_interval_seconds`), retrieved and pushed to local system clipboard

### Linux (Python)

1. Install dependencies:
   ```bash
   ./setup-python.sh
   # or manually: pip install -r requirements.txt
   ```

2. **Optional: Install copyq for enhanced multi-MIME support**
   ```bash
   sudo apt install copyq
   # Start copyq daemon
   copyq &
   ```

3. Run the Python script:
   ```bash
   python clipson.py
   ```
4. Copy any text or image to clipboard
5. Content will be automatically saved and synced to Nextcloud
6. Peers' clipboard syncs are regularly checked (parameter `remote_check_interval_seconds`), retrieved and pushed to local system clipboard

**Note:** If copyq is installed and running as a daemon, ClipSon will automatically use it for enhanced multi-MIME clipboard support. If copyq is not available or not running, ClipSon will fall back to xclip.

## Configuration

### Proxy Configuration
- See the section "proxy" in `config.json`

### Nextcloud Configuration
- Ensure your Nextcloud instance is accessible via WebDAV.
### Nextcloud credentials and WebDAV URL configuration 
- Edit the `config.json` file to set your Nextcloud WebDAV URL, username, and password.
  
**Note:** For multi-device synchronization, ensure all ClipSon instances use the same Nextcloud account and remote folder configuration.

### Application Configuration
- The `config.json` file contains application level tweaks:

```  
"app": {
        "max_history": 200, // Maximum number of clipboard entries in local cache (directory clipboard-captures)
        "remote_check_interval_seconds": 1, // Interval to check for remote changes
        "debug_enabled": false, // Enable debug logging
        "use_copyq": true // Use copyq for enhanced multi-MIME support on Linux (requires copyq daemon running)
    }
```

## Requirements

### Windows
- Windows PowerShell 5.1 or later
- .NET Framework (usually pre-installed on Windows)
- Nextcloud account and credentials

### Linux
- Python 3.6+
- Required Python packages (see [requirements.txt](requirements.txt)):
  - requests==2.31.0
- System dependencies:
  - xclip (required for basic clipboard operations)
  - notify-send/libnotify-bin (required for desktop notifications)
  - copyq (optional, for enhanced multi-MIME clipboard support)
- Nextcloud account and credentials

## Known Issues

### Clipboard Event Loopback
The Windows clipboard change event handler triggers whenever the clipboard content changes, which can happen during paste operations (e.g., in MS Word). This may cause a loopback where ClipSon detects its own clipboard changes when setting content from remote peers, potentially triggering unnecessary processing.

### Text/HTML Encoding Issues
- **RTF Support Disabled**: text/rtf synchronization from Windows is currently disabled due to UTF-8 encoding issues. HTML format (text/html) works correctly.
- **Diacritics Encoding**: text/html encoding of diacritics and special characters is broken when retrieving clipboard content on the Windows side, resulting in corrupted Unicode characters in the synchronized content.

### CopyQ Dependency
On Linux, enhanced multi-MIME clipboard support requires copyq to be running as a daemon. If copyq is not available or the daemon is not running, ClipSon will automatically fall back to xclip with limited format support.

## Files

- [`config_template.json`](config_template.json) - Configuration file for Nextcloud credentials and WebDAV URL
- [`clipson.ps1`](clipson.ps1) - Main PowerShell script for Windows
- [`clipson.py`](clipson.py) - Main Python script for Linux (uses xclip/copyq for clipboard manipulation and notify-send for desktop notification)
- [`requirements.txt`](requirements.txt) - Python dependencies
- [`setup-python.sh`](setup-python.sh) - Linux dependencies setup script
- `clipboard-captures/` - local clipboard captures history directory (created automatically)

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.
