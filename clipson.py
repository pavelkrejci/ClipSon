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
from datetime import timezone
from pathlib import Path
import requests
from requests.auth import HTTPBasicAuth
import signal
import getpass
import gzip
import base64
import hashlib

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
        
        # File paths - now using .json.gz extension for compressed transfer
        self.local_sync_file = Path(f'./clipboard-{self.hostname}.json')
        self.local_sync_file_gz = Path(f'./clipboard-{self.hostname}.json.gz')
        self.local_upload_file = f"clipboard-{self.hostname}.json.gz"
        
        # Setup signal handler
        signal.signal(signal.SIGINT, self.signal_handler)
        
        # Use copyq or fallback to xclip based on config
        self.use_copyq = CONFIG_DATA['app'].get('use_copyq', True)
        print(f"Using {'copyq' if self.use_copyq else 'xclip'} for clipboard operations.")
    
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
        """Get current clipboard content (text only) using copyq or xclip"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'clipboard', 'text/plain'], capture_output=True, text=True)
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except Exception:
            pass
        return None

    def has_clipboard_image(self):
        """Check if clipboard contains an image using copyq or xclip"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'clipboard', '?'], capture_output=True, text=True)
                if result.returncode == 0:
                    targets = result.stdout.strip().split('\n')
                    image_targets = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff']
                    return any(target.strip() in image_targets for target in targets)
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'], capture_output=True, text=True)
                if result.returncode == 0:
                    targets = result.stdout.strip().split('\n')
                    image_targets = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff']
                    return any(target.strip() in image_targets for target in targets)
        except Exception:
            pass
        return False

    def compress_json_file(self, json_file, gz_file):
        """Compress JSON file to .gz format"""
        try:
            with open(json_file, 'rb') as f_in:
                with gzip.open(gz_file, 'wb', compresslevel=1) as f_out:  # Fast compression
                    original_size = 0
                    while True:
                        chunk = f_in.read(8192)
                        if not chunk:
                            break
                        f_out.write(chunk)
                        original_size += len(chunk)
            
            compressed_size = gz_file.stat().st_size
            if DEBUG:
                ratio = (1 - compressed_size / original_size) * 100 if original_size > 0 else 0
                print(f"DEBUG: File compression {original_size} -> {compressed_size} bytes ({ratio:.1f}% saved)")
            
            return True
        except Exception as e:
            if DEBUG:
                print(f"DEBUG: File compression failed: {e}")
            return False

    def decompress_gz_file(self, gz_file, json_file):
        """Decompress .gz file to JSON format"""
        try:
            with gzip.open(gz_file, 'rb') as f_in:
                with open(json_file, 'wb') as f_out:
                    while True:
                        chunk = f_in.read(8192)
                        if not chunk:
                            break
                        f_out.write(chunk)
            return True
        except Exception as e:
            if DEBUG:
                print(f"DEBUG: File decompression failed: {e}")
            return False

    def save_clipboard_image(self):
        """Save clipboard image to file and upload"""
        try:
            # Get image data from clipboard
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-o'], 
                                  capture_output=True)
            
            if result.returncode == 0 and result.stdout:
                # Calculate hash from image data
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
                
                # Create unified JSON format for image

                image_b64 = base64.b64encode(result.stdout).decode('utf-8')
                upload_content = {
                    "type": "CLIPBOARD_IMAGE",
                    "data": image_b64,
                    "format": "png",
                    "size": len(result.stdout)
                }
                upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
                
                # Save to local sync file, compress, and upload
                with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                    f.write(upload_json)
                
                if self.compress_json_file(self.local_sync_file, self.local_sync_file_gz):
                    if self.upload_to_webdav(self.local_sync_file_gz, self.local_upload_file):
                        print(f"{datetime.now().strftime('%H:%M:%S')} - Uploaded compressed image to WebDAV: {self.local_upload_file}")
                
                # Show notification
                self.show_notification("ClipSon", f"Image captured: {filename}")
                
                # Update last clipboard content for comparison (without data field to save memory)
                self.last_clipboard_content = json.dumps({
                    "type": "CLIPBOARD_IMAGE",
                    "format": "png",
                    "size": len(result.stdout),
                    "hash": current_image_hash
                }, ensure_ascii=False, sort_keys=True)
                
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
                    # Now looking for .json.gz files
                    if filename.startswith('clipboard-') and filename.endswith('.json.gz'):
                        last_modified = None
                        if lastmodified_elem is not None and lastmodified_elem.text:
                            try:
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
        
        # Filter out our own file (compressed json.gz now)
        base = f"clipboard-{self.hostname}"
        peer_files = [f for f in remote_files if f['name'] != f"{base}.json.gz"]
        
        if DEBUG:
            print(f"DEBUG: Total remote files found: {len(remote_files)}")
            print(f"DEBUG: My hostname: {self.hostname}")
            print(f"DEBUG: My file: {base}.json.gz")
            print(f"DEBUG: All remote files: {[f['name'] for f in remote_files]}")
            print(f"DEBUG: Filtered peer files count: {len(peer_files)}")
        
        if not peer_files:
            print("No remote clipboard files from other machines found.")
            print(f"Will only upload to: {base}.json.gz")
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
                # Download compressed file
                with open(local_file, 'wb') as f:
                    f.write(response.content)
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
        # Only look for .json.gz files now
        peer_files = [f for f in remote_files if f['name'] != f"{base}.json.gz" and f['name'].startswith('clipboard-') and f['name'].endswith('.json.gz')]
        
        if DEBUG:
            print(f"DEBUG: Remote check - Total files: {len(remote_files)}, My file: {base}.json.gz, Peer files: {len(peer_files)}")
        
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
                
                # Download and decompress
                temp_download_file_gz = f'./temp-remote-download-{filename.replace(".json.gz", "")}.json.gz'
                temp_download_file_json = f'./temp-remote-download-{filename.replace(".json.gz", "")}.json'
                
                if self.download_remote_file(filename, temp_download_file_gz):
                    try:
                        # Decompress the downloaded file
                        if self.decompress_gz_file(temp_download_file_gz, temp_download_file_json):
                            with open(temp_download_file_json, 'r', encoding='utf-8') as f:
                                file_content = f.read()
                            has_content = file_content.strip()
                            
                            if has_content and remote_timestamp > most_recent_timestamp:
                                most_recent_content = file_content
                                most_recent_timestamp = remote_timestamp
                                most_recent_filename = filename
                        
                        # Update timestamp regardless
                        self.remote_file_timestamps[filename] = remote_timestamp
                        
                        # Clean up temp files
                        if not DEBUG:
                            for temp_file in [temp_download_file_gz, temp_download_file_json]:
                                try:
                                    os.remove(temp_file)
                                except:
                                    pass
                        else:
                            print(f"DEBUG: Keeping temp files for inspection: {temp_download_file_gz}, {temp_download_file_json}")
                                
                    except Exception as e:
                        print(f"Error processing {filename}: {e}")
                        if not DEBUG:
                            for temp_file in [temp_download_file_gz, temp_download_file_json]:
                                try:
                                    os.remove(temp_file)
                                except:
                                    pass
                        else:
                            print(f"DEBUG: Keeping temp files for inspection after error: {temp_download_file_gz}, {temp_download_file_json}")
        
        # Apply the most recent update if found
        if most_recent_content:
            if self.set_clipboard_content_unified(most_recent_content):
                # Show notification
                preview = most_recent_content[:50] + "..." if len(most_recent_content) > 50 else most_recent_content
                self.show_notification("ClipSon", f"Remote update from {most_recent_filename}: {preview}")
                
                # Set last_clipboard_content based on content type to prevent re-capture
                self.update_last_clipboard_content_from_remote(most_recent_content)
            
            return most_recent_content
        
        return None

    def set_clipboard_content_unified(self, content):
        """Set clipboard content from unified JSON format"""
        try:
            # Always try to parse JSON if it looks like JSON            
            if DEBUG:
                print(f"DEBUG: set_clipboard_content_unified received JSON: {content[:120]}{'...' if len(content) > 120 else ''}")
            # Handle possible BOM (Byte Order Mark) at the start of content
            if content and ord(content[0]) == 0xFEFF:
                if DEBUG:
                    print("DEBUG: Detected UTF-8 BOM, stripping it before parsing JSON.")
                content = content.lstrip('\ufeff')
            try:
                data = json.loads(content)
                content_type = data.get("type")
                if DEBUG:
                    print(f"DEBUG: Parsed JSON type: {content_type}, keys: {list(data.keys())}")
                
                if content_type == "CLIPBOARD_IMAGE":
                    # Handle image content
                    image_data = base64.b64decode(data["data"])
                    if DEBUG:
                        print(f"DEBUG: Decoded image data, length: {len(image_data)}")
                    return self.set_clipboard_image(image_data)
                
                elif content_type == "MULTI_FORMAT_CLIPBOARD":
                    if "formats" in data:
                        if DEBUG:
                            print(f"DEBUG: MULTI_FORMAT_CLIPBOARD formats: {list(data['formats'].keys())}")
                        return self.set_clipboard_multiple_formats(data["formats"])
                
                elif content_type == "PLAIN_TEXT":
                    if "content" in data:
                        if DEBUG:
                            print(f"DEBUG: PLAIN_TEXT content: {repr(data['content'])}")
                        return self.set_clipboard_content(data["content"])
                    else:
                        if DEBUG:
                            print("DEBUG: PLAIN_TEXT but no content field, setting empty string")
                        return self.set_clipboard_content("")
            except Exception as e:
                print(f"DEBUG: Failed to parse JSON clipboard content: {e}")
                # Fallback to legacy or plain text handling below

            # Fall back to legacy or plain text handling
            if DEBUG:
                print(f"DEBUG: set_clipboard_content_unified fallback, content: {content[:120]}{'...' if len(content) > 120 else ''}")
            return self.set_clipboard_rich_content(content)
        except Exception as e:
            print(f"Error setting unified clipboard content: {e}")
            return self.set_clipboard_content(content)

    def update_last_clipboard_content_from_remote(self, content):
        """Update last_clipboard_content from remote content to prevent re-capture"""
        try:
            if content.strip().startswith('{'):
                data = json.loads(content)
                content_type = data.get("type")
                
                if content_type == "CLIPBOARD_IMAGE":
                    # Store comparison format for images (without data to save memory)
                    image_data = base64.b64decode(data["data"])
                    image_hash = hashlib.md5(image_data).hexdigest()
                    self.last_image_hash = image_hash  # Update hash tracking
                    self.last_clipboard_content = json.dumps({
                        "type": "CLIPBOARD_IMAGE",
                        "format": data.get("format", "png"),
                        "size": data.get("size", len(image_data)),
                        "hash": image_hash
                    }, ensure_ascii=False, sort_keys=True)
                
                elif content_type == "MULTI_FORMAT_CLIPBOARD" and "formats" in data:
                    # Store in comparison format (same as we use in main loop)
                    self.last_clipboard_content = json.dumps({"formats": data["formats"]}, ensure_ascii=False, sort_keys=True)
                
                elif content_type == "PLAIN_TEXT" and "content" in data:
                    # Store plain text content for comparison
                    self.last_clipboard_content = data["content"]
                
                else:
                    self.last_clipboard_content = content
            else:
                # Regular text or legacy rich content
                self.last_clipboard_content = content
                
        except Exception:
            self.last_clipboard_content = content

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
        
        # Save to local sync file and upload (content wrapped in JSON for consistency)
        upload_content = {
            "type": "PLAIN_TEXT",
            "content": content
        }
        upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
        
        with open(self.local_sync_file, 'w', encoding='utf-8') as f:
            f.write(upload_json)
        
        # Compress and upload
        if self.compress_json_file(self.local_sync_file, self.local_sync_file_gz):
            self.upload_to_webdav(self.local_sync_file_gz, self.local_upload_file)
        
        # Show notification
        preview = content[:50] + "..." if len(content) > 50 else content
        self.show_notification("ClipSon", f"Captured: {preview}")
    
    def set_clipboard_content(self, content):
        """Set clipboard content using copyq or xclip"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'copy', 'text/plain', content], check=True)
                return result.returncode == 0
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard'], input=content, text=True, check=True)
                return result.returncode == 0
        except Exception:
            return False

    def set_clipboard_image(self, image_data):
        """Set clipboard image content using copyq or xclip"""
        try:
            # Update hash tracking immediately to prevent re-capture
            self.last_image_hash = hashlib.md5(image_data).hexdigest()
            if DEBUG:
                print(f"DEBUG: Updated last_image_hash to prevent re-capture: {self.last_image_hash}")
            if self.use_copyq:
                result = subprocess.run(['copyq', 'copy', 'image/png', '-'], input=image_data, check=True)
                return result.returncode == 0
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png'], input=image_data, check=True)
                return result.returncode == 0
        except Exception as e:
            print(f"Error setting clipboard image: {e}")
            return False

    def get_clipboard_formats(self):
        """Get available clipboard formats using copyq or xclip"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'clipboard', '?'], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip().split('\n')
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'], capture_output=True, text=True)
                if result.returncode == 0:
                    return result.stdout.strip().split('\n')
        except Exception:
            pass
        return []

    def has_clipboard_rich_text(self):
        """Check if clipboard contains rich text formats"""
        try:
            formats = self.get_clipboard_formats()
            # Look for rich text MIME types
            rich_text_formats = [
                'text/html', 'text/rtf', 'text/richtext',
                'application/rtf', 'application/x-rtf',
                'text/x-moz-url', 'text/uri-list',
                'application/x-color'
            ]
            return any(fmt.strip() in rich_text_formats for fmt in formats)
        except Exception:
            pass
        return False

    def save_clipboard_rich_content(self):
        """Save all available clipboard formats to files"""
        try:
            formats = self.get_clipboard_formats()
            if not formats:
                return False

            file_number = self.get_next_file_number()
            base_filename = self.output_dir / f"clipboard_rich_{file_number:03d}"

            saved_files = []
            format_data = {}  # Store all format data for multi-format upload
            seen_content = set()  # Track content to avoid duplicates

            # Priority order for formats with unique filenames
            # text/plain first for better cross-platform compatibility
            format_priority = [
                ('text/plain', '.txt'),        # Best cross-platform compatibility
                ('text/html', '.html'),
                ('text/rtf', '.rtf'), 
                ('application/rtf', '.app-rtf'),
                ('application/x-rtf', '.x-rtf'),
                ('text/richtext', '.richtext'),
                ('text/uri-list', '.uri'),
                ('text/x-moz-url', '.url'),
                ('UTF8_STRING', '.utf8'),      # Linux-specific but good fallback
                ('STRING', '.string'),         # Linux-specific
                ('TEXT', '.text')              # Linux-specific
            ]

            for format_name, extension in format_priority:
                if format_name.strip() in formats:
                    try:
                        # Get content for this format
                        result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', format_name, '-o'], 
                                              capture_output=True, text=True, errors='replace')
                        
                        if result.returncode == 0 and result.stdout.strip():
                            content = result.stdout.strip()
                            
                            # Skip if we've already seen this exact content
                            if content in seen_content:
                                if DEBUG:
                                    print(f"DEBUG: Skipping duplicate content for {format_name}")
                                continue
                            
                            seen_content.add(content)
                            filename = f"{base_filename}{extension}"
                            
                            with open(filename, 'w', encoding='utf-8') as f:
                                f.write(content)
                            
                            saved_files.append(filename)
                            format_data[format_name] = content
                            
                            if DEBUG:
                                print(f"DEBUG: Saved {format_name} to {filename}")
                                
                    except Exception as e:
                        if DEBUG:
                            print(f"DEBUG: Failed to save format {format_name}: {e}")
                        continue

            if saved_files and format_data:
                print(f"{datetime.now().strftime('%H:%M:%S')} - Rich content saved: {len(saved_files)} unique formats")
                for filename in saved_files:
                    print(f"  - {filename}")
                
                # Create multi-format upload content
                upload_content = {
                    "type": "MULTI_FORMAT_CLIPBOARD",
                    "formats": format_data  # Ensure consistent structure
                }
                upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
                
                # Save to local sync file, compress, and upload
                with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                    f.write(upload_json)
                
                if self.compress_json_file(self.local_sync_file, self.local_sync_file_gz):
                    self.upload_to_webdav(self.local_sync_file_gz, self.local_upload_file)
                
                # Show notification with unique format types
                unique_extensions = list(set([ext for filename in saved_files for ext in [Path(filename).suffix]]))
                format_list = ', '.join(unique_extensions[:3])
                self.show_notification("ClipSon", f"Rich content captured: {format_list}")
                
                # Store only the format data for comparison
                self.last_clipboard_content = json.dumps({"formats": format_data}, ensure_ascii=False, sort_keys=True)
                
                return True
                
        except Exception as e:
            print(f"Error saving rich clipboard content: {e}")
        
        return False

    def set_clipboard_multiple_formats(self, format_data):
        """Set multiple clipboard formats simultaneously using copyq"""
        try:
            if DEBUG:
                print(f"DEBUG: Setting {len(format_data)} clipboard formats (multi-mime)")
            # Build the argument list for copyq: [mime1, value1, mime2, value2, ...]
            args = ['copyq', 'copy']

            # Use a priority order for consistency
            set_priority = [
                'text/plain', 'text/html', 'text/rtf', 'application/rtf', 'application/x-rtf',
                'text/richtext', 'text/uri-list', 'text/x-moz-url',
                'UTF8_STRING', 'STRING', 'TEXT'
            ]
            used = set()
            
            for format_name in set_priority:
                if format_name in format_data and format_name not in used:
                    args.append(format_name)
                    args.append(format_data[format_name])
                    used.add(format_name)
            # Add any remaining formats not in priority list
            for format_name in format_data:
                if format_name not in used:
                    args.append(format_name)
                    args.append(format_data[format_name])
                    used.add(format_name)

            if DEBUG:
                print(f"DEBUG: copyq arguments: {args[:200]}...")

            # Check if total argument length would exceed system limits (e.g., 2MB)
            total_args_length = sum(len(str(arg)) for arg in args)
            if total_args_length > 2 * 1024 * 1024:  # 2MB limit
                if DEBUG:
                    print(f"DEBUG: Total arguments length ({total_args_length} bytes) exceeds 2MB limit, falling back to text/plain only")
                # Fall back to setting only text/plain format
                if 'text/plain' in format_data:
                    return self.set_clipboard_content(format_data['text/plain'])
                else:
                    # Use the first available format
                    first_format = next(iter(format_data.values()))
                    return self.set_clipboard_content(first_format)
                
            # Call copyq with all formats at once
            result = subprocess.run(args, check=True)
            if result.returncode == 0:
                if DEBUG:
                    print("DEBUG: Successfully set multiple formats with copyq")
                return True
            else:
                if DEBUG:
                    print("DEBUG: copyq returned non-zero exit code")
                return False
        except Exception as e:
            print(f"Error setting multiple clipboard formats: {e}")
            return False

    def set_clipboard_format(self, content, format_type):
        """Set clipboard content with specific format using copyq or xclip"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'copy', format_type, content], check=True)
                return result.returncode == 0
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', format_type], input=content, text=True, check=True)
                return result.returncode == 0
        except Exception as e:
            print(f"Error setting clipboard format {format_type}: {e}")
            return False

    def set_clipboard_rich_content(self, content):
        """Set clipboard content for legacy rich content format"""
        try:
            # Check if this is legacy rich content format
            if content.startswith("RICH_CONTENT_FORMATS:"):
                lines = content.split('\n', 2)
                if len(lines) >= 3:
                    formats_line = lines[0]
                    actual_content = lines[2]
                    
                    # Extract format information
                    format_str = formats_line.replace("RICH_CONTENT_FORMATS:", "").strip()
                    available_formats = [fmt.strip() for fmt in format_str.split(",")]
                    
                    if DEBUG:
                        print(f"DEBUG: Setting legacy rich content with formats: {', '.join(available_formats)}")
                    
                    # Set content based on available format (prioritize HTML, then RTF, then plain text)
                    if "HTML" in available_formats:
                        return self.set_clipboard_format(actual_content, "text/html")
                    elif "RTF" in available_formats:
                        return self.set_clipboard_format(actual_content, "text/rtf")
                    else:
                        # Fallback to plain text
                        return self.set_clipboard_content(actual_content)
            else:
                # Regular text content
                return self.set_clipboard_content(content)
                
        except Exception as e:
            print(f"Error setting legacy rich clipboard content: {e}")
            return self.set_clipboard_content(content)

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
        print("Supported formats: Multi-format Text, Images (unified JSON), HTML, RTF, URLs")
        
        while True:
            try:
                # Import json at the top to avoid scope issues                                
                # Check all remote files for updates
                remote_content = self.check_all_remote_files_for_updates()
                # Note: last_clipboard_content is already updated inside check_all_remote_files_for_updates()
                
                # Check for clipboard image first (highest priority)
                if self.has_clipboard_image():
                    # Get current image data and check hash
                    result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-o'], capture_output=True)
                    if result.returncode == 0 and result.stdout:
                        current_image_hash = hashlib.md5(result.stdout).hexdigest()
                        
                        # Create comparison format
                        current_image_comparison = json.dumps({
                            "type": "CLIPBOARD_IMAGE",
                            "format": "png",
                            "size": len(result.stdout),
                            "hash": current_image_hash
                        }, ensure_ascii=False, sort_keys=True)
                        
                        if current_image_comparison != self.last_clipboard_content:
                            self.save_clipboard_image()
                        elif DEBUG:
                            print(f"DEBUG: Image unchanged, skipping...")
                            
                # Check for rich text formats (medium priority)
                elif self.has_clipboard_rich_text():
                    # Get the current rich content and compare with last captured
                    current_rich_formats = self.get_clipboard_formats()
                    
                    # Build current multi-format data for comparison
                    # Updated priority order with text/plain first
                    format_priority = [
                        ('text/plain', '.txt'), ('text/html', '.html'), ('text/rtf', '.rtf'), 
                        ('application/rtf', '.app-rtf'), ('application/x-rtf', '.x-rtf'), 
                        ('text/richtext', '.richtext'), ('text/uri-list', '.uri'), 
                        ('text/x-moz-url', '.url'), ('UTF8_STRING', '.utf8'), 
                        ('STRING', '.string'), ('TEXT', '.text')
                    ]
                    
                    # Get all available format content
                    current_format_data = {}
                    seen_content = set()
                    
                    for format_name, ext in format_priority:
                        if format_name.strip() in current_rich_formats:
                            try:
                                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', format_name, '-o'], 
                                                      capture_output=True, text=True, errors='replace')
                                if result.returncode == 0 and result.stdout.strip():
                                    content = result.stdout.strip()
                                    if content not in seen_content:
                                        current_format_data[format_name] = content
                                        seen_content.add(content)
                            except Exception:
                                continue
                    
                    if current_format_data:
                        # Create comparison string using content hashes instead of raw content for dynamic formats
                        comparison_data = {}
                        for fmt, content in current_format_data.items():
                            if fmt in ['text/html', 'text/rtf', 'application/rtf', 'application/x-rtf']:
                                # For formats that might have dynamic content, use content hash for comparison
                                # but still store full content for upload
                                content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                                comparison_data[fmt] = content_hash
                            else:
                                # For stable formats, use actual content
                                comparison_data[fmt] = content
                        
                        current_comparison_content = json.dumps({"formats": comparison_data}, ensure_ascii=False, sort_keys=True)
                        
                        # Check if we have a previous comparison
                        has_meaningful_change = True
                        if self.last_clipboard_content and self.last_clipboard_content.startswith('{"formats"'):
                            try:
                                last_comparison_data = json.loads(self.last_clipboard_content)["formats"]
                                
                                # Compare stable text formats for meaningful changes
                                # Now including text/plain as the primary stable format
                                stable_formats = ['text/plain', 'UTF8_STRING', 'STRING', 'TEXT']
                                current_stable = {k: v for k, v in comparison_data.items() if k in stable_formats}
                                last_stable = {k: v for k, v in last_comparison_data.items() if k in stable_formats}
                                
                                # If stable content hasn't changed, consider it unchanged
                                if current_stable == last_stable:
                                    has_meaningful_change = False
                                    if DEBUG:
                                        print(f"DEBUG: Only dynamic content changed (HTML/RTF timestamps), ignoring...")                                
                                            
                            except Exception as e:
                                if DEBUG:
                                    print(f"DEBUG: Error comparing content: {e}")
                        
                        # Only save if there's a meaningful change
                        if has_meaningful_change:
                            if DEBUG and not (self.last_clipboard_content and self.last_clipboard_content.startswith('{"formats"')):
                                print(f"DEBUG: Rich content changed, saving...")
                            self.save_clipboard_rich_content()
                            # Store the comparison data for next time
                            self.last_clipboard_content = current_comparison_content
                        elif DEBUG:
                            print(f"DEBUG: Rich content unchanged (ignoring dynamic changes), skipping...")
                else:
                    # Check clipboard text content (lowest priority)
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
