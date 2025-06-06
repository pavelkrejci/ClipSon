# ClipSon
Nextcloud clipboard synchronization tool for Linux (Python) and Windows (PowerShell)

**ClipSon works over Nextcloud's WebDAV protocol, supporting proxy configurations and secure connections to synchronize your clipboard content across devices.**

## Features

- **Nextcloud WebDAV protocol integration** - Uses Nextcloud's WebDAV API for reliable file transfers
- **Proxy support** - Works through corporate proxies and network configurations
- **Cross-platform clipboard monitoring** - Works on both Windows and Linux
- **Real-time synchronization** - Automatically monitors clipboard changes
- **Text and image support** - Handles both text content and images
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
2. Run the Python script:
   ```bash
   python clipson.py
   ```
3. Copy any text or image to clipboard
4. Content will be automatically saved and synced to Nextcloud
5. Peers' clipboard syncs are regularly checked (parameter `remote_check_interval_seconds`), retrieved and pushed to local system clipboard

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
        "debug_enabled": false // Enable debug logging  
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
- Nextcloud account and credentials

## Files

- [`config_template.json`](config_template.json) - Configuration file for Nextcloud credentials and WebDAV URL
- [`clipson.ps1`](clipson.ps1) - Main PowerShell script for Windows
- [`clipson.py`](clipson.py) - Main Python script for Linux (uses xclip for clipboard manipulation and notify-send for desktop notification)
- [`requirements.txt`](requirements.txt) - Python dependencies
- [`setup-python.sh`](setup-python.sh) - Linux dependencies setup script
- `clipboard-captures/` - local clipboard captures history directory (created automatically)

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.
