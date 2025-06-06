#!/usr/bin/env python3
"""
ClipSon - Advanced Clipboard Synchronization Tool for Linux
Dependencies: python3, xclip, notify-send (libnotify-bin)
Install with: sudo apt install python3 xclip libnotify-bin
"""

import subprocess
import time
import os
import sys
import json
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
import requests
from requests.auth import HTTPBasicAuth
import signal
import getpass

def load_configuration():
    """Load configuration from JSON file"""
    config_path = Path('./config.json')
    if not config_path.exists():
        print("Configuration file not found: config.json")
        print("Please create config.json file with your Nextcloud settings.")
        sys.exit(1)
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print(f"Failed to parse configuration file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Failed to load configuration file: {e}")
        sys.exit(1)

# Load configuration
CONFIG_DATA = load_configuration()

# Global debug constant (now from config)
DEBUG = CONFIG_DATA['app']['debug_enabled']


# Configuration (now loaded from config.json)
CONFIG = {
    'server_url': CONFIG_DATA['nextcloud']['server_url'],
    'username': CONFIG_DATA['nextcloud']['username'],
    'password': CONFIG_DATA['nextcloud']['password'],
    'remote_folder': CONFIG_DATA['nextcloud']['remote_folder']
}

def get_password_if_needed():
    """Prompt for password if not configured"""
    if not CONFIG['password'].strip():
        print(f"No password configured for user: {CONFIG['username']}")
        password = getpass.getpass("Please enter your Nextcloud password: ")
        if not password.strip():
            print("Password cannot be empty. Exiting.")
            sys.exit(1)
        CONFIG['password'] = password
        print("Password configured successfully.")

class ClipSon:
    def __init__(self):
        # Get password if needed before setting up WebDAV
        get_password_if_needed()
        
        self.hostname = os.uname().nodename
        self.output_dir = Path(f'./clipboard-captures')
        self.output_dir.mkdir(exist_ok=True)
        
        self.max_history = CONFIG_DATA['app']['max_history']
        self.file_counter = 0
        self.last_clipboard_content = ""
        self.remote_file_timestamps = {}  # Track timestamps for each remote file
        self.last_remote_check = 0
        self.remote_check_interval = CONFIG_DATA['app']['remote_check_interval_seconds']  # seconds
        
        # WebDAV setup
        self.webdav_base_url = f"{CONFIG['server_url'].rstrip('/')}/remote.php/dav/files/{CONFIG['username']}/"
        self.auth = HTTPBasicAuth(CONFIG['username'], CONFIG['password'])
        
        # File paths
        self.local_sync_file = Path(f'./clipboard-{self.hostname}.txt')
        self.local_upload_file = f"clipboard-{self.hostname}.txt"
        
        # Setup signal handler
        signal.signal(signal.SIGINT, self.signal_handler)
    
    def signal_handler(self, signum, frame):
        print("\nClipSon stopped.")
        sys.exit(0)
    
    def check_dependencies(self):
        """Check if required system dependencies are available"""
        missing_deps = []
        
        for cmd in ['xclip', 'notify-send']:
            if subprocess.run(['which', cmd], capture_output=True).returncode != 0:
                missing_deps.append(cmd)
        
        if missing_deps:
            print(f"Missing dependencies: {' '.join(missing_deps)}")
            print(f"Install with: sudo apt install {' '.join(missing_deps)}")
            sys.exit(1)
    
    def show_notification(self, title, message, icon='info'):
        """Show desktop notification"""
        try:
            subprocess.run([
                'notify-send', title, message,
                f'--icon={icon}', '--expire-time=3000'
            ], check=False)
        except Exception:
            print(f"NOTIFICATION: {title} - {message}")
    
    def get_clipboard_content(self):
        """Get current clipboard content (text only)"""
        try:
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                  capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except Exception:
            pass
        return None

    def has_clipboard_image(self):
        """Check if clipboard contains an image"""
        try:
            # Check if clipboard has image data
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                targets = result.stdout.strip().split('\n')
                # Look for image MIME types
                image_targets = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff']
                return any(target.strip() in image_targets for target in targets)
        except Exception:
            pass
        return False

    def save_clipboard_image(self):
        """Save clipboard image to file and upload"""
        try:
            # Get image data from clipboard
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-o'], 
                                  capture_output=True)
            
            if result.returncode == 0 and result.stdout:
                # Calculate hash from image data
                import hashlib
                current_image_hash = hashlib.md5(result.stdout).hexdigest()
                
                # Check if this is the same image as last time
                if hasattr(self, 'last_image_hash') and current_image_hash == self.last_image_hash:
                    if DEBUG:
                        print(f"DEBUG: Image hash matches previous - skipping save and upload")
                    return False
                
                # Hash is different, proceed with saving
                file_number = self.get_next_file_number()
                filename = self.output_dir / f"clipboard_image_{file_number:03d}.png"
                
                with open(filename, 'wb') as f:
                    f.write(result.stdout)
                
                print(f"{datetime.now().strftime('%H:%M:%S')} - Image saved: {filename}")
                
                # Update hash tracking
                self.last_image_hash = current_image_hash
                
                # Upload image to WebDAV (use hostname-based file name)
                remote_image_file = f"clipboard-{self.hostname}.png"
                if self.upload_to_webdav(filename, remote_image_file):
                    print(f"{datetime.now().strftime('%H:%M:%S')} - Uploaded image to WebDAV: {remote_image_file}")
                
                # Show notification
                self.show_notification("ClipSon", f"Image captured: {filename}")
                
                # Update last clipboard content to prevent re-capture
                self.last_clipboard_content = f"__IMAGE_CONTENT_{len(result.stdout)}_BYTES__"
                
                return True
        except Exception as e:
            print(f"Error saving clipboard image: {e}")
        return False

    def test_webdav_connection(self):
        """Test WebDAV connection"""
        try:
            response = requests.request('PROPFIND', self.webdav_base_url, 
                                      auth=self.auth, headers={'Depth': '0'}, 
                                      timeout=10)
            return response.status_code == 207
        except Exception as e:
            print(f"Failed to connect to Nextcloud: {e}")
            return False
    
    def get_remote_clipboard_files(self):
        """Discover remote clipboard files"""
        try:
            webdav_url = f"{self.webdav_base_url}{CONFIG['remote_folder']}"
            
            propfind_body = '''<?xml version="1.0" encoding="utf-8" ?>
<D:propfind xmlns:D="DAV:">
    <D:prop>
        <D:displayname/>
        <D:getlastmodified/>
    </D:prop>
</D:propfind>'''
            
            response = requests.request('PROPFIND', webdav_url,
                                      auth=self.auth,
                                      headers={'Content-Type': 'application/xml', 'Depth': '1'},
                                      data=propfind_body,
                                      timeout=10)
            
            if response.status_code != 207:
                print(f"Failed to discover remote clipboard files: {response.status_code} {response.reason}")
                if DEBUG:
                    print(f"DEBUG: Response text: {response.text}")
                return []                    
            
            # Parse XML response
            root = ET.fromstring(response.text)
            files = []
            
            for response_elem in root.findall('.//{DAV:}response'):
                displayname_elem = response_elem.find('.//{DAV:}displayname')
                lastmodified_elem = response_elem.find('.//{DAV:}getlastmodified')
                
                if displayname_elem is not None and displayname_elem.text:
                    filename = displayname_elem.text
                    if filename.startswith('clipboard-') and (filename.endswith('.txt') or filename.endswith('.png')):
                        last_modified = None
                        if lastmodified_elem is not None and lastmodified_elem.text:
                            try:
                                from datetime import timezone
                                # Parse without timezone first, then make it UTC
                                dt_str = lastmodified_elem.text.replace(' GMT', '')
                                last_modified_utc = datetime.strptime(dt_str, '%a, %d %b %Y %H:%M:%S')
                                # Make it UTC aware
                                last_modified_utc = last_modified_utc.replace(tzinfo=timezone.utc)
                                # Convert to local time
                                last_modified = last_modified_utc.astimezone()
                            except ValueError:
                                pass
                        
                        files.append({
                            'name': filename,
                            'last_modified': last_modified
                        })
            
            return files
        except Exception as e:
            print(f"Failed to discover remote clipboard files: {e}")
            return []
    
    def discover_remote_peers(self):
        """Discover and track all remote clipboard files"""
        print("Discovering remote clipboard files...")
        remote_files = self.get_remote_clipboard_files()
        
        # Filter out our own files (both text and image)
        base = f"clipboard-{self.hostname}"
        peer_files = [f for f in remote_files if f['name'] not in (f"{base}.txt", f"{base}.png")]
        
        if not peer_files:
            print("No remote clipboard files from other machines found.")
            print(f"Will only upload to: {base}.txt and {base}.png")
            return []
        
        print(f"\nFound {len(peer_files)} remote peer(s):")
        for file_info in peer_files:
            time_str = file_info['last_modified'].strftime('%Y-%m-%d %H:%M:%S') if file_info['last_modified'] else 'Unknown'
            print(f"  - {file_info['name']} (Modified: {time_str})")
            
            # Initialize timestamp tracking
            if file_info['last_modified']:
                self.remote_file_timestamps[file_info['name']] = file_info['last_modified'].timestamp()
            else:
                self.remote_file_timestamps[file_info['name']] = 0
        
        return [f['name'] for f in peer_files]
    
    def get_remote_file_timestamp(self, remote_file):
        """Get remote file timestamp"""
        try:
            webdav_url = f"{self.webdav_base_url}{CONFIG['remote_folder']}{remote_file}"
            
            propfind_body = '''<?xml version="1.0" encoding="utf-8" ?>
<D:propfind xmlns:D="DAV:">
    <D:prop>
        <D:getlastmodified/>
    </D:prop>
</D:propfind>'''
            
            response = requests.request('PROPFIND', webdav_url,
                                      auth=self.auth,
                                      headers={'Content-Type': 'application/xml', 'Depth': '0'},
                                      data=propfind_body,
                                      timeout=10)
            
            if response.status_code == 207:
                root = ET.fromstring(response.text)
                lastmodified_elem = root.find('.//{DAV:}getlastmodified')
                if lastmodified_elem is not None and lastmodified_elem.text:
                    from datetime import timezone
                    # Parse without timezone first, then make it UTC
                    dt_str = lastmodified_elem.text.replace(' GMT', '')
                    last_modified_utc = datetime.strptime(dt_str, '%a, %d %b %Y %H:%M:%S')
                    # Make it UTC aware
                    last_modified_utc = last_modified_utc.replace(tzinfo=timezone.utc)
                    # Convert to local time
                    return last_modified_utc.astimezone()
        except Exception:
            pass
        return None
    
    def download_remote_file(self, remote_file, local_file):
        """Download remote file"""
        try:
            webdav_url = f"{self.webdav_base_url}{CONFIG['remote_folder']}{remote_file}"
            response = requests.get(webdav_url, auth=self.auth, timeout=10)
            
            if response.status_code == 200:
                if local_file.endswith('.png'):
                    with open(local_file, 'wb') as f:
                        f.write(response.content)
                else:
                    with open(local_file, 'w', encoding='utf-8') as f:
                        f.write(response.text)
                return True
        except Exception:
            pass
        return False
    
    def upload_to_webdav(self, local_file, remote_file):
        """Upload file to WebDAV"""
        try:
            webdav_url = f"{self.webdav_base_url}{CONFIG['remote_folder']}{remote_file}"
            
            with open(local_file, 'rb') as f:
                response = requests.put(webdav_url, auth=self.auth, data=f, timeout=10)
            
            if response.status_code in [201, 204]:
                print(f"{datetime.now().strftime('%H:%M:%S')} - Uploaded to WebDAV: {remote_file}")
                return True
            else:
                print(f"{datetime.now().strftime('%H:%M:%S')} - Error uploading to WebDAV: {remote_file}")
                return False
        except Exception as e:
            print(f"{datetime.now().strftime('%H:%M:%S')} - Error uploading to WebDAV: {e}")
            return False
    
    def check_all_remote_files_for_updates(self):
        """Check all remote files for updates and return the most recent one"""
        current_time = time.time()
        if current_time - self.last_remote_check < self.remote_check_interval:
            return None
        
        self.last_remote_check = current_time
        
        # Get current list of remote files
        remote_files = self.get_remote_clipboard_files()
        base = f"clipboard-{self.hostname}"
        peer_files = [f for f in remote_files if f['name'] not in (f"{base}.txt", f"{base}.png")]
        
        most_recent_content = None
        most_recent_timestamp = 0
        most_recent_filename = None
        
        for file_info in peer_files:
            filename = file_info['name']
            
            # Check if this is a new file or an updated file
            is_new_file = filename not in self.remote_file_timestamps
            
            if is_new_file:
                if file_info['last_modified']:
                    self.remote_file_timestamps[filename] = 0  # Set to 0 so it will be processed as an update
                    print(f"{datetime.now().strftime('%H:%M:%S')} - New remote peer discovered: {filename}")
                else:
                    self.remote_file_timestamps[filename] = 0
                    continue
            
            if not file_info['last_modified']:
                continue
                
            remote_timestamp = file_info['last_modified'].timestamp()
            last_known_timestamp = self.remote_file_timestamps[filename]
            
            # Check if this file has been updated (or is newly discovered)
            if remote_timestamp > last_known_timestamp:
                if not is_new_file:
                    print(f"{datetime.now().strftime('%H:%M:%S')} - Remote file updated: {filename}")
                
                # Download and check if it's the most recent
                file_ext = Path(filename).suffix
                temp_download_file = f'./temp-remote-download-{filename.replace(file_ext, "")}{file_ext}'
                
                if self.download_remote_file(filename, temp_download_file):
                    try:
                        if file_ext == '.txt':
                            with open(temp_download_file, 'r', encoding='utf-8') as f:
                                file_content = f.read()
                            has_content = file_content.strip()
                        elif file_ext == '.png':
                            with open(temp_download_file, 'rb') as f:
                                file_content = f.read()
                            has_content = len(file_content) > 0
                        else:
                            has_content = False
                        
                        if has_content and remote_timestamp > most_recent_timestamp:
                            most_recent_content = file_content
                            most_recent_timestamp = remote_timestamp
                            most_recent_filename = filename
                        
                        # Update timestamp regardless
                        self.remote_file_timestamps[filename] = remote_timestamp
                        
                        os.remove(temp_download_file)
                    except Exception as e:
                        print(f"Error processing {filename}: {e}")
                        try:
                            os.remove(temp_download_file)
                        except:
                            pass
        
        # Apply the most recent update if found
        if most_recent_content:
            if most_recent_filename.endswith('.png'):
                # Set clipboard image and show notification
                if self.set_clipboard_image(most_recent_content):
                    self.show_notification("ClipSon", f"Remote image update from {most_recent_filename}")
                    # Use the binary content as our "last clipboard content" to prevent re-upload
                    self.last_clipboard_content = f"__IMAGE_CONTENT_{len(most_recent_content)}_BYTES__"
                else:
                    self.show_notification("ClipSon", f"Remote image update from {most_recent_filename} (clipboard set failed, saved locally)")
                    self.last_clipboard_content = f"__IMAGE_CONTENT_{len(most_recent_content)}_BYTES__"
            else:
                if self.set_clipboard_content(most_recent_content):
                    # Show notification
                    preview = most_recent_content[:50] + "..." if len(most_recent_content) > 50 else most_recent_content
                    self.show_notification("ClipSon", f"Remote update from {most_recent_filename}: {preview}")
                    self.last_clipboard_content = most_recent_content  # Prevent re-capture of this content
            
            return most_recent_content
        
        return None
    
    def get_next_file_number(self):
        """Get next file number for rotation"""
        self.file_counter += 1
        if self.file_counter > 999:
            self.file_counter = 1
        return self.file_counter
    
    def save_clipboard_text(self, content):
        """Save clipboard text to numbered file and upload"""
        # Save to numbered file
        file_number = self.get_next_file_number()
        filename = self.output_dir / f"clipboard_text_{file_number:03d}.txt"
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"{datetime.now().strftime('%H:%M:%S')} - Text saved: {filename}")
        
        # Save to local sync file and upload
        with open(self.local_sync_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        self.upload_to_webdav(self.local_sync_file, self.local_upload_file)
        
        # Show notification
        preview = content[:50] + "..." if len(content) > 50 else content
        self.show_notification("ClipSon", f"Captured: {preview}")
    
    def set_clipboard_content(self, content):
        """Set clipboard content"""
        try:
            subprocess.run(['xclip', '-selection', 'clipboard'], 
                          input=content, text=True, check=True)
            return True
        except Exception:
            return False

    def set_clipboard_image(self, image_data):
        """Set clipboard image content"""
        try:
            # Update hash tracking immediately to prevent re-capture
            import hashlib
            self.last_image_hash = hashlib.md5(image_data).hexdigest()
            if DEBUG:
                print(f"DEBUG: Updated last_image_hash to prevent re-capture: {self.last_image_hash}")
            
            # Use xclip to set image data to clipboard
            subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png'], 
                          input=image_data, check=True)
            return True
        except Exception as e:
            print(f"Error setting clipboard image: {e}")
            return False

    def run(self):
        """Main run loop"""
        # Check dependencies
        self.check_dependencies()
        
        # Test WebDAV connection
        print("Connecting to Nextcloud...")
        if not self.test_webdav_connection():
            print("Failed to connect to Nextcloud. Please check your configuration.")
            sys.exit(1)
        
        # Discover remote peers
        peer_files = self.discover_remote_peers()
        
        print("ClipSon started. Press Ctrl+C to stop.")
        print(f"Captured content will be saved to: {self.output_dir}")
        print(f"Maximum entries: {self.max_history} (older files will be automatically deleted)")
        print(f"Local sync file: {self.local_sync_file}")
        print(f"Local upload file (upload): {self.local_upload_file}")
        print(f"Remote peer files (download): {len(peer_files)} peer(s)")
        for peer_file in peer_files:
            print(f"  - {peer_file}")
        print(f"Remote check interval: {self.remote_check_interval} seconds")
        
        while True:
            try:
                # Check all remote files for updates
                remote_content = self.check_all_remote_files_for_updates()
                # Note: last_clipboard_content is already updated inside check_all_remote_files_for_updates()
                
                # Check for clipboard image first (higher priority)
                if self.has_clipboard_image():
                    # Check if this is a new image (different from our last captured content)
                    if not self.last_clipboard_content.startswith("__IMAGE_CONTENT_"):
                        self.save_clipboard_image()
                else:
                    # Check clipboard text content
                    current_content = self.get_clipboard_content()
                    if current_content and current_content != self.last_clipboard_content:
                        self.save_clipboard_text(current_content)
                        self.last_clipboard_content = current_content
                
                # Wait before next check
                time.sleep(0.5)
                
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"Error in main loop: {e}")
                time.sleep(1)

if __name__ == "__main__":
    clipson = ClipSon()
    clipson.run()
