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
import base64
import hashlib
import re
import urllib.parse
import mimetypes
import tempfile
import shutil
import uuid
import gzip
import atexit

CPSN_MAGIC = b'CPSN'

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

def debug_print(message):
    """Print debug message with timestamp if DEBUG is enabled"""
    if DEBUG:
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]  # Include milliseconds
        print(f"[DEBUG {timestamp}] {message}")

class ClipSon:
    def __init__(self):
        # Get password if needed before setting up WebDAV
        get_password_if_needed()
        
        self.hostname = os.uname().nodename
        self.use_7z_encryption = CONFIG_DATA['app'].get('use_7z_encryption', True)
        self.archive_extension = '.cpsn' if self.use_7z_encryption else '.gz'
        self.my_sync_filenames = {
            f"clipboard-{self.hostname}.cpsn",
            f"clipboard-{self.hostname}.gz"
        }

        self.max_history = CONFIG_DATA['app']['max_history']
        self.file_counter = 0
        self.last_clipboard_content = ""
        self.remote_file_timestamps = {}  # Track timestamps for each remote file
        self.last_remote_check = 0
        self.remote_check_interval = CONFIG_DATA['app']['remote_check_interval_seconds']  # seconds
        self.first_clipboard_capture = True  # Flag to skip sync on first capture
        
        # WebDAV setup
        self.webdav_base_url = f"{CONFIG['server_url'].rstrip('/')}/remote.php/dav/files/{CONFIG['username']}/"
        self.auth = HTTPBasicAuth(CONFIG['username'], CONFIG['password'])
        
        # File paths - archive extension depends on compression mode
        self.local_sync_file = Path(f'./clipboard-{self.hostname}.json')
        self.local_sync_file_archive = Path(f'./clipboard-{self.hostname}{self.archive_extension}')
        self.local_upload_file = f"clipboard-{self.hostname}{self.archive_extension}"

        # Temporary directory for restored files (cleaned up on exit)
        self.restored_temp_dir = Path(tempfile.mkdtemp(prefix=f"clipson-restored-{self.hostname}-"))
        atexit.register(self._cleanup_restored_temp_dir)
        
        # Setup signal handler
        signal.signal(signal.SIGINT, self.signal_handler)
        
        # Use copyq or fallback to xclip based on config
        self.use_copyq = CONFIG_DATA['app'].get('use_copyq', True)
        
        # Check copyq daemon connection if enabled
        if self.use_copyq:
            if self.check_copyq_daemon():
                print(f"Using copyq for clipboard operations (daemon running).")
            else:
                print("Warning: copyq daemon not responding, falling back to xclip.")
                print("Tip: You can manually start copyq daemon with: copyq &")
                self.use_copyq = False
        
        if not self.use_copyq:
            print(f"Using xclip for clipboard operations.")

    def signal_handler(self, signum, frame):
        self._cleanup_restored_temp_dir()
        print("\nClipSon stopped.")
        sys.exit(0)

    def _cleanup_restored_temp_dir(self):
        try:
            if self.restored_temp_dir and self.restored_temp_dir.exists():
                shutil.rmtree(self.restored_temp_dir, ignore_errors=True)
        except Exception:
            pass

    def _clear_restored_temp_dir(self):
        try:
            if self.restored_temp_dir and self.restored_temp_dir.exists():
                for item in self.restored_temp_dir.iterdir():
                    if item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                    else:
                        item.unlink(missing_ok=True)
        except Exception:
            pass
    
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
            # Avoid blocking on DBus/notify-send hangs; fire-and-forget.
            process = subprocess.Popen(
                [
                    'notify-send', title, message,
                    f'--icon={icon}', '--expire-time=3000'
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            try:
                process.communicate(timeout=1)
            except subprocess.TimeoutExpired:
                process.kill()
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

    def has_clipboard_file_uris(self):
        """Check if clipboard contains file URIs (text/uri-list)"""
        try:
            formats = self.get_clipboard_formats()
            return 'text/uri-list' in [fmt.strip() for fmt in formats]
        except Exception:
            pass
        return False

    def get_clipboard_file_uris(self):
        """Get file URIs from clipboard"""
        try:
            if self.use_copyq:
                result = subprocess.run(['copyq', 'clipboard', 'text/uri-list'], capture_output=True, text=True)
            else:
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'text/uri-list', '-o'], 
                                      capture_output=True, text=True)
            
            if result.returncode == 0 and result.stdout.strip():
                # Parse URI list - each URI on a separate line
                uri_list = []
                for line in result.stdout.strip().split('\n'):
                    line = line.strip()
                    if line and line.startswith('file://'):
                        # Decode URI to get actual file path
                        file_path = urllib.parse.unquote(line[7:])  # Remove 'file://' prefix
                        uri_list.append(file_path)
                return uri_list
        except Exception as e:
            debug_print(f"Error getting file URIs from clipboard: {e}")
        return []

    def save_clipboard_files(self):
        """Save clipboard files to Nextcloud and local storage"""
        try:
            file_uris = self.get_clipboard_file_uris()
            if not file_uris:
                return False

            files_data = {}

            for file_path in file_uris:
                try:
                    # Check if file exists and is readable
                    if not os.path.exists(file_path):
                        debug_print(f"File not found: {file_path}")
                        continue
                    
                    if not os.path.isfile(file_path):
                        debug_print(f"Not a regular file: {file_path}")
                        continue

                    # Get file info
                    file_stat = os.stat(file_path)
                    file_size = file_stat.st_size
                    
                    # Limit file size to prevent excessive memory usage (default 50MB)
                    max_file_size = CONFIG_DATA['app'].get('max_file_copy_size_mb', 50) * 1024 * 1024
                    if file_size > max_file_size:
                        print(f"File too large to copy: {file_path} ({file_size / 1024 / 1024:.1f} MB > {max_file_size / 1024 / 1024:.1f} MB)")
                        continue

                    # Read file content
                    with open(file_path, 'rb') as f:
                        file_content = f.read()
                    
                    # Encode file content as base64
                    file_b64 = base64.b64encode(file_content).decode('utf-8')
                    
                    # Get file name and extension
                    file_name = os.path.basename(file_path)
                    
                    # Store file data
                    files_data[file_name] = {
                        "original_path": file_path,
                        "content": file_b64,
                        "size": file_size,
                        "mime_type": self.get_mime_type(file_path)
                    }
                    
                    debug_print(f"Processed file: {file_name} ({file_size} bytes)")
                    
                except Exception as e:
                    debug_print(f"Error processing file {file_path}: {e}")
                    continue

            if files_data:
                print(f"{datetime.now().strftime('%H:%M:%S')} - File(s) captured: {len(files_data)} file(s)")
                for filename in files_data.keys():
                    print(f"  - {filename}")
                
                # Skip sync on first clipboard capture after app start
                if self.first_clipboard_capture:
                    self.first_clipboard_capture = False
                    print(f"{datetime.now().strftime('%H:%M:%S')} - Skipping sync for first clipboard capture")
                else:
                    # Create file upload content
                    upload_content = {
                        "type": "CLIPBOARD_FILES",
                        "files": files_data
                    }
                    upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
                    
                    # Save to local sync file, compress, and upload
                    with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                        f.write(upload_json)
                    
                        if self.compress_json_file(self.local_sync_file, self.local_sync_file_archive):
                            if self.upload_to_webdav(self.local_sync_file_archive, self.local_upload_file):
                                print(f"{datetime.now().strftime('%H:%M:%S')} - Uploaded encrypted files to WebDAV: {self.local_upload_file}")
                
                # Show notification
                file_names = list(files_data.keys())[:3]
                file_list = ', '.join(file_names)
                if len(files_data) > 3:
                    file_list += f" (+{len(files_data) - 3} more)"
                self.show_notification("ClipSon", f"Files captured: {file_list}")
                
                return True
                
        except Exception as e:
            print(f"Error saving clipboard files: {e}")
        return False

    def get_mime_type(self, file_path):
        """Get MIME type for a file"""
        try:
            # Use python's mimetypes module
            mime_type, _ = mimetypes.guess_type(file_path)
            if mime_type:
                return mime_type
        except:
            pass
        
        # Fallback to application/octet-stream
        return "application/octet-stream"

    def _get_7z_executable(self):
        """Return path to 7z executable (Linux)."""
        preferred = Path('/usr/bin/7z')
        if preferred.exists():
            return str(preferred)

        found = shutil.which('7z')
        if found:
            return found

        raise FileNotFoundError("7z executable not found (expected /usr/bin/7z or 7z in PATH)")

    def compress_json_file(self, json_file, archive_file):
        """Compress JSON file using 7z (encrypted) or gzip (unencrypted)."""
        try:
            if self.use_7z_encryption:
                try:
                    if self._compress_with_7z(json_file, archive_file):
                        return True
                except FileNotFoundError as e:
                    debug_print(f"7z not available ({e}); falling back to gzip")

            return self._compress_with_gzip(json_file, archive_file)
        except Exception as e:
            debug_print(f"File compression failed: {e}")
            return False

    def _compress_with_7z(self, json_file, archive_file):
        """Encrypt JSON file into a 7z archive using the Nextcloud password."""
        json_path = Path(json_file).resolve()
        archive_path = Path(archive_file).resolve()

        seven_zip = self._get_7z_executable()
        password = CONFIG.get('password', '')
        if not str(password).strip():
            raise ValueError("Nextcloud password is empty; cannot encrypt")

        original_size = json_path.stat().st_size
        tmp_archive_path = None
        try:
            # Create payload in a temp directory, but create the archive temp file
            # in the destination directory for atomic replace.
            with tempfile.TemporaryDirectory(prefix='clipson-7z-') as tmpdir:
                tmpdir_path = Path(tmpdir)
                payload_name = 'payload.json'
                payload_path = tmpdir_path / payload_name
                shutil.copyfile(json_path, payload_path)

                tmp_archive_path = archive_path.parent / f"{archive_path.name}.{uuid.uuid4().hex}.tmp"
                if tmp_archive_path.exists():
                    tmp_archive_path.unlink()

                args = [
                    seven_zip, 'a', '-t7z', '-mhe=on', f'-p{password}',
                    '-y', '-bd', str(tmp_archive_path), payload_name
                ]

                result = subprocess.run(
                    args,
                    cwd=str(tmpdir_path),
                    capture_output=True,
                    text=True
                )

                if result.returncode != 0:
                    combined = (result.stderr or '').strip()
                    if result.stdout and result.stdout.strip():
                        combined = (combined + "\n" + result.stdout.strip()).strip()
                    debug_print(f"7z encryption failed (exit {result.returncode}): {combined}")
                    return False

            # Some 7z builds may append .7z when the output name has no .7z.
            if not tmp_archive_path.exists():
                candidate1 = tmp_archive_path.with_name(tmp_archive_path.name + '.7z')
                candidate2 = tmp_archive_path.with_suffix(tmp_archive_path.suffix + '.7z')
                if candidate1.exists():
                    tmp_archive_path = candidate1
                elif candidate2.exists():
                    tmp_archive_path = candidate2

            os.replace(str(tmp_archive_path), str(archive_path))
        finally:
            if tmp_archive_path and tmp_archive_path.exists():
                try:
                    tmp_archive_path.unlink()
                except Exception:
                    pass

        # Prepend CPSN signature to the resulting archive.
        # (Remote sync files are CPSN + raw 7z bytes.)
        try:
            with open(archive_path, 'rb') as f_in:
                head = f_in.read(4)

            if head != CPSN_MAGIC:
                prefixed_tmp = archive_path.parent / f"{archive_path.name}.{uuid.uuid4().hex}.cpsn"
                with open(archive_path, 'rb') as f_in, open(prefixed_tmp, 'wb') as f_out:
                    f_out.write(CPSN_MAGIC)
                    shutil.copyfileobj(f_in, f_out, length=1024 * 1024)
                os.replace(str(prefixed_tmp), str(archive_path))
        except Exception as e:
            debug_print(f"Warning: failed to prepend CPSN header: {e}")
            return False

        encrypted_size = archive_path.stat().st_size
        ratio = (1 - encrypted_size / original_size) * 100 if original_size > 0 else 0
        debug_print(f"File encryption {original_size} -> {encrypted_size} bytes ({ratio:.1f}% smaller)")

        return True

    def _compress_with_gzip(self, json_file, archive_file):
        """Compress JSON file using gzip (no encryption)."""
        json_path = Path(json_file).resolve()
        archive_path = Path(archive_file).resolve()
        tmp_archive_path = archive_path.parent / f"{archive_path.name}.{uuid.uuid4().hex}.tmp"

        try:
            if tmp_archive_path.exists():
                tmp_archive_path.unlink()

            with open(json_path, 'rb') as f_in, gzip.open(tmp_archive_path, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out, length=1024 * 1024)

            os.replace(str(tmp_archive_path), str(archive_path))
            original_size = json_path.stat().st_size
            compressed_size = archive_path.stat().st_size
            ratio = (1 - compressed_size / original_size) * 100 if original_size > 0 else 0
            debug_print(f"Gzip compression {original_size} -> {compressed_size} bytes ({ratio:.1f}% smaller)")
            return True
        except Exception as e:
            debug_print(f"Gzip compression failed: {e}")
            return False
        finally:
            if tmp_archive_path.exists():
                try:
                    tmp_archive_path.unlink()
                except Exception:
                    pass

    def decompress_gz_file(self, archive_file, json_file):
        """Decrypt CPSN/7z archives or decompress gzip archives back into JSON."""
        archive_path = Path(archive_file).resolve()
        json_path = Path(json_file).resolve()

        try:
            with open(archive_path, 'rb') as f_in:
                magic = f_in.read(6)
        except Exception as e:
            debug_print(f"Failed to read archive header: {e}")
            return False

        prefer = None
        if magic.startswith(CPSN_MAGIC) or magic.startswith(b'\x37\x7a\xbc\xaf\x27\x1c'):
            prefer = '7z'
        elif magic.startswith(b'\x1f\x8b'):
            prefer = 'gzip'
        else:
            prefer = '7z' if self.use_7z_encryption else 'gzip'

        methods = [prefer, 'gzip' if prefer == '7z' else '7z']

        for method in methods:
            if method == '7z':
                try:
                    if self._decompress_with_7z(archive_path, json_path):
                        return True
                except FileNotFoundError as e:
                    debug_print(f"7z not available ({e})")
                    continue
                except Exception as e:
                    debug_print(f"7z decompression failed: {e}")
                    continue
            else:
                if self._decompress_with_gzip(archive_path, json_path):
                    return True

        return False

    def _decompress_with_7z(self, archive_path, json_path):
        seven_zip = self._get_7z_executable()
        password = CONFIG.get('password', '')
        if not str(password).strip():
            raise ValueError("Nextcloud password is empty; cannot decrypt")

        with tempfile.TemporaryDirectory(prefix='clipson-7z-') as tmpdir:
            tmpdir_path = Path(tmpdir)
            out_dir_arg = f"-o{tmpdir_path}"  # 7z expects -o<dir>

            archive_to_extract = archive_path
            with open(archive_path, 'rb') as f_in:
                head = f_in.read(4)
                if head == CPSN_MAGIC:
                    stripped_path = tmpdir_path / 'archive.7z'
                    with open(stripped_path, 'wb') as f_out:
                        shutil.copyfileobj(f_in, f_out, length=1024 * 1024)
                    archive_to_extract = stripped_path

            args = [
                seven_zip, 'x', '-y', '-bd', f'-p{password}', out_dir_arg, str(archive_to_extract)
            ]
            result = subprocess.run(
                args,
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                debug_print(f"7z decryption failed (exit {result.returncode}): {result.stderr.strip()}")
                return False

            payload_path = tmpdir_path / 'payload.json'
            if payload_path.exists():
                shutil.copyfile(payload_path, json_path)
                return True

            extracted_files = [p for p in tmpdir_path.rglob('*') if p.is_file()]
            if not extracted_files:
                debug_print("7z decryption produced no files")
                return False

            shutil.copyfile(extracted_files[0], json_path)
            return True

    def _decompress_with_gzip(self, archive_path, json_path):
        try:
            with gzip.open(archive_path, 'rb') as f_in, open(json_path, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out, length=1024 * 1024)
            return True
        except Exception as e:
            debug_print(f"Gzip decompression failed: {e}")
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
                    debug_print(f"Image hash matches previous - skipping save and upload")
                    return False
                
                # Hash is different, proceed with upload (no local capture file)
                print(f"{datetime.now().strftime('%H:%M:%S')} - Image captured")
                
                # Update hash tracking
                self.last_image_hash = current_image_hash
                
                # Skip sync on first clipboard capture after app start
                if self.first_clipboard_capture:
                    self.first_clipboard_capture = False
                    print(f"{datetime.now().strftime('%H:%M:%S')} - Skipping sync for first clipboard capture")
                else:
                    # Create unified JSON format for image
                    image_b64 = base64.b64encode(result.stdout).decode('utf-8')
                    upload_content = {
                        "type": "CLIPBOARD_IMAGE",
                        "data": image_b64,
                        "format": "png",
                        "size": len(result.stdout)
                    }
                    upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
                    
                    # Save to local sync file, encrypt, and upload
                    with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                        f.write(upload_json)
                    
                    if self.compress_json_file(self.local_sync_file, self.local_sync_file_archive):
                        if self.upload_to_webdav(self.local_sync_file_archive, self.local_upload_file):
                            print(f"{datetime.now().strftime('%H:%M:%S')} - Uploaded encrypted image to WebDAV: {self.local_upload_file}")
                
                # Show notification
                self.show_notification("ClipSon", "Image captured")
                
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
                debug_print(f"Response text: {response.text}")
                return []                    
            
            # Parse XML response
            root = ET.fromstring(response.text)
            files = []
            
            for response_elem in root.findall('.//{DAV:}response'):
                displayname_elem = response_elem.find('.//{DAV:}displayname')
                lastmodified_elem = response_elem.find('.//{DAV:}getlastmodified')
                
                if displayname_elem is not None and displayname_elem.text:
                    filename = displayname_elem.text
                    if filename.startswith('clipboard-') and (filename.endswith('.cpsn') or filename.endswith('.gz')):
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
        
        base = f"clipboard-{self.hostname}"
        peer_files = [
            f for f in remote_files
            if f['name'] not in self.my_sync_filenames and f['name'].startswith('clipboard-') and (f['name'].endswith('.cpsn') or f['name'].endswith('.gz'))
        ]
        
        debug_print(f"Total remote files found: {len(remote_files)}")
        debug_print(f"My hostname: {self.hostname}")
        debug_print(f"My file candidates: {sorted(self.my_sync_filenames)}")
        debug_print(f"All remote files: {[f['name'] for f in remote_files]}")
        debug_print(f"Filtered peer files count: {len(peer_files)}")
        
        if not peer_files:
            print("No remote clipboard files from other machines found.")
            print(f"Will only upload to: {base}{self.archive_extension}")
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
        my_remote_files = sorted(
            f["name"]
            for f in remote_files
            if f["name"] in self.my_sync_filenames
        )
        peer_files = [
            f for f in remote_files
            if f['name'] not in self.my_sync_filenames
            and f['name'].startswith('clipboard-')
            and (f['name'].endswith('.cpsn') or f['name'].endswith('.gz'))
        ]
        
        debug_print(f"Remote check - Total files: {len(remote_files)}, My files: {my_remote_files}")
        for file_info in peer_files:
            filename = file_info['name']
            timestamp = file_info['last_modified'].strftime('%Y-%m-%d %H:%M:%S') if file_info['last_modified'] else 'Unknown'
            debug_print(f"Peer file: {filename}, Modified: {timestamp}")
        
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
                
                temp_download_file_archive = Path(f'./temp-remote-download-{filename}')
                temp_download_file_json = temp_download_file_archive.with_suffix('.json')
                
                if self.download_remote_file(filename, temp_download_file_archive):
                    try:
                        # Decompress the downloaded file
                        if self.decompress_gz_file(temp_download_file_archive, temp_download_file_json):
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
                            for temp_file in [temp_download_file_archive, temp_download_file_json]:
                                try:
                                    os.remove(temp_file)
                                except:
                                    pass
                        else:
                            debug_print(f"Keeping temp files for inspection: {temp_download_file_archive}, {temp_download_file_json}")
                                
                    except Exception as e:
                        print(f"Error processing {filename}: {e}")
                        if not DEBUG:
                            for temp_file in [temp_download_file_archive, temp_download_file_json]:
                                try:
                                    os.remove(temp_file)
                                except:
                                    pass
                        else:
                            debug_print(f"Keeping temp files for inspection after error: {temp_download_file_archive}, {temp_download_file_json}")
        
        # Apply the most recent update if found
        if most_recent_content:
            if self.set_clipboard_content_unified(most_recent_content):
                # Get the actual clipboard fingerprint after setting content
                # This ensures our comparison data matches what's actually in the clipboard
                actual_fingerprint = self.get_current_clipboard_fingerprint()
                
                self.last_clipboard_content = actual_fingerprint
                debug_print(f"Updated last_clipboard_content with actual fingerprint: {actual_fingerprint}")
                
                # Show notification
                preview = most_recent_content[:50] + "..." if len(most_recent_content) > 50 else most_recent_content
                self.show_notification("ClipSon", f"Remote update from {most_recent_filename}: {preview}")
            return most_recent_content
        
        return None

    
    def get_next_file_number(self):
        """Get next file number for rotation"""
        self.file_counter += 1
        if self.file_counter > 999:
            self.file_counter = 1
        return self.file_counter
    
    def save_clipboard_text(self, content):
        """Save clipboard text and upload"""
        print(f"{datetime.now().strftime('%H:%M:%S')} - Text captured")
        
        # Skip sync on first clipboard capture after app start
        if self.first_clipboard_capture:
            self.first_clipboard_capture = False
            print(f"{datetime.now().strftime('%H:%M:%S')} - Skipping sync for first clipboard capture")
        else:
            # Save to local sync file and upload (content wrapped in JSON for consistency)
            upload_content = {
                "type": "PLAIN_TEXT",
                "content": content
            }
            upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
            
            with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                f.write(upload_json)
            
            # Encrypt and upload
            if self.compress_json_file(self.local_sync_file, self.local_sync_file_archive):
                self.upload_to_webdav(self.local_sync_file_archive, self.local_upload_file)
        
        # Show notification
        preview = content[:50] + "..." if len(content) > 50 else content
        self.show_notification("ClipSon", f"Captured: {preview}")
    
    def set_clipboard_content(self, content):
        """Set clipboard content using copyq or xclip"""
        try:
            if self.use_copyq:
                # Use -- to prevent copyq from expanding escape sequences like \n, \t, \\
                result = subprocess.run(['copyq', 'copy', '--', content], check=True)
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
            debug_print(f"Updated last_image_hash to prevent re-capture: {self.last_image_hash}")
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
            
            # Check if any rich text formats are present
            has_rich = any(fmt.strip() in rich_text_formats for fmt in formats)
            
            # If text/uri-list is present, check if it contains actual file URIs
            # If it does, don't consider this as rich text
            if has_rich and 'text/uri-list' in [fmt.strip() for fmt in formats]:
                file_uris = self.get_clipboard_file_uris()
                if file_uris:  # If we have actual file URIs, this is file content, not rich text
                    # Check if we have other rich text formats besides text/uri-list
                    other_rich_formats = [f for f in rich_text_formats if f != 'text/uri-list']
                    return any(fmt.strip() in other_rich_formats for fmt in formats)
            
            return has_rich
        except Exception:
            pass
        return False

    def save_clipboard_rich_content(self):
        """Save all available clipboard formats to JSON and upload"""
        try:
            formats = self.get_clipboard_formats()
            if not formats:
                return False

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
                                debug_print(f"Skipping duplicate content for {format_name}")
                                continue
                            
                            seen_content.add(content)
                            format_data[format_name] = content
                            debug_print(f"Captured {format_name} content")
                                
                    except Exception as e:
                        debug_print(f"Failed to save format {format_name}: {e}")
                        continue

            if format_data:
                print(f"{datetime.now().strftime('%H:%M:%S')} - Rich content captured: {len(format_data)} unique formats")
                
                # Skip sync on first clipboard capture after app start
                if self.first_clipboard_capture:
                    self.first_clipboard_capture = False
                    print(f"{datetime.now().strftime('%H:%M:%S')} - Skipping sync for first clipboard capture")
                else:
                    # Create multi-format upload content
                    upload_content = {
                        "type": "MULTI_FORMAT_CLIPBOARD",
                        "formats": format_data  # Ensure consistent structure
                    }
                    upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
                    
                    # Save to local sync file, encrypt, and upload
                    with open(self.local_sync_file, 'w', encoding='utf-8') as f:
                        f.write(upload_json)
                    
                    if self.compress_json_file(self.local_sync_file, self.local_sync_file_archive):
                        self.upload_to_webdav(self.local_sync_file_archive, self.local_upload_file)
                
                # Show notification with unique format types
                format_list = ', '.join(list(format_data.keys())[:3])
                self.show_notification("ClipSon", f"Rich content captured: {format_list}")
                
                return True
                
        except Exception as e:
            print(f"Error saving rich clipboard content: {e}")
        
        return False

    def set_clipboard_multiple_formats(self, format_data):
        """Set multiple clipboard formats simultaneously using copyq or xclip"""
        try:
            if self.use_copyq:
                debug_print(f"Setting {len(format_data)} clipboard formats (multi-mime) with copyq")
                
                set_priority = [
                    'text/plain', 'text/html', 'text/rtf', 'application/rtf', 'application/x-rtf',
                    'text/richtext', 'text/uri-list', 'text/x-moz-url'                    
                ]
                
                # Use temp files to avoid command-line payload limits
                with tempfile.TemporaryDirectory(prefix='clipson-formats-') as tmpdir:
                    tmpdir_path = Path(tmpdir)
                    first_format = None
                    used = set()
                    
                    for format_name in set_priority:
                        if format_name in format_data and format_name not in used:
                            used.add(format_name)
                            content = format_data[format_name]
                            
                            # Write content to temp file
                            temp_file = tmpdir_path / f"format_{len(used)}.txt"
                            with open(temp_file, 'w', encoding='utf-8') as f:
                                f.write(content)
                            
                            # First format: use 'write', subsequent formats: use 'change'
                            if first_format is None:
                                # copyq write 0 text/plain - < file
                                cmd = ['copyq', 'write', '0', format_name, '-']
                                debug_print(f"Writing first format {format_name}")
                                with open(temp_file, 'rb') as f:
                                    result = subprocess.run(cmd, stdin=f, check=True)
                                if result.returncode != 0:
                                    debug_print(f"Failed to write format {format_name}")
                                    return False
                                first_format = format_name
                            else:
                                # copyq change 0 text/html - < file
                                cmd = ['copyq', 'change', '0', format_name, '-']
                                debug_print(f"Adding format {format_name}")
                                with open(temp_file, 'rb') as f:
                                    result = subprocess.run(cmd, stdin=f, check=True)
                                if result.returncode != 0:
                                    debug_print(f"Failed to change format {format_name}")
                                    return False
                    
                    if first_format:
                        # Select the clipboard item to activate it
                        result = subprocess.run(['copyq', 'select', '0'], check=True)
                        if result.returncode == 0:
                            debug_print(f"Successfully set {len(used)} formats with copyq (file-based)")
                            return True
                        else:
                            debug_print("Failed to select clipboard item")
                            return False
                    else:
                        debug_print("No formats to set")
                        return False
            else:
                # xclip: only supports one format at a time, prefer text/html > text/rtf > text/plain
                debug_print(f"Setting clipboard formats with xclip (no multi-mime support)")
                for fmt in ['text/html', 'text/rtf', 'text/plain']:
                    if fmt in format_data:
                        return self.set_xclip_format(format_data[fmt], fmt)
                # fallback: set first available format
                for fmt, value in format_data.items():
                    return self.set_xclip_format(value, fmt)
                return False
        except Exception as e:
            print(f"Error setting multiple clipboard formats: {e}")
            return False

    def set_xclip_format(self, content, format_type):
        """Set clipboard content with specific format using xclip"""
        try:            
            debug_print(f"Setting {format_type} format with xclip")
            debug_print(f"Content excerpt: {content[:120]}{'...' if len(content) > 120 else ''}")
            
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', format_type], 
                input=content, text=True, check=True)
            return result.returncode == 0
        except Exception as e:
            print(f"Error setting clipboard format {format_type}: {e}")
            return False

    def set_clipboard_files(self, files_data):
        """Set clipboard with file URIs and restore files to temp location"""
        try:
            # Clear temporary directory for restored files
            self._clear_restored_temp_dir()
            
            restored_files = []
            uri_list = []
            
            for filename, file_info in files_data.items():
                try:
                    # Decode file content from base64
                    file_content = base64.b64decode(file_info["content"])
                    
                    # Create safe filename
                    safe_filename = self.sanitize_filename(filename)
                    temp_file_path = self.restored_temp_dir / safe_filename
                    
                    # Write file to temp location
                    with open(temp_file_path, 'wb') as f:
                        f.write(file_content)
                    
                    # Add to URI list
                    file_uri = f"file://{temp_file_path.absolute()}"
                    uri_list.append(file_uri)
                    restored_files.append(temp_file_path)
                    
                    debug_print(f"Restored file: {safe_filename} ({len(file_content)} bytes) to {temp_file_path}")
                    
                except Exception as e:
                    debug_print(f"Error restoring file {filename}: {e}")
                    continue
            
            if uri_list:
                # Set clipboard with file URI list
                uri_content = '\n'.join(uri_list)
                
                if self.use_copyq:
                    # Set both text/uri-list and text/plain for better compatibility
                    # Use -- to prevent escape sequence expansion in file paths
                    args = ['copyq', 'copy', '--', 'text/uri-list', uri_content, 'text/plain', uri_content]
                    result = subprocess.run(args, check=True)
                    success = result.returncode == 0
                else:
                    # Use xclip to set text/uri-list
                    success = self.set_xclip_format(uri_content, 'text/uri-list')
                
                if success:
                    print(f"Restored {len(restored_files)} file(s) to clipboard:")
                    for file_path in restored_files:
                        print(f"  - {file_path}")
                    return True
                else:
                    debug_print("Failed to set clipboard with file URIs")
                    return False
            
            return False
            
        except Exception as e:
            print(f"Error setting clipboard files: {e}")
            return False

    def sanitize_filename(self, filename):
        """Sanitize filename for safe file system usage"""
        # Remove or replace problematic characters
        safe_chars = '-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
        safe_filename = ''.join(c if c in safe_chars else '_' for c in filename)
        
        # Ensure it's not empty and doesn't start with dot
        if not safe_filename or safe_filename.startswith('.'):
            safe_filename = 'restored_file_' + safe_filename
        
        return safe_filename[:255]  # Limit length

    def set_clipboard_content_unified(self, content):
        """Set clipboard content from unified JSON format"""
        try:
            debug_print(f"set_clipboard_content_unified received JSON: {content[:120]}{'...' if len(content) > 120 else ''}")
            # Handle possible BOM (Byte Order Mark) at the start of content
            if content and ord(content[0]) == 0xFEFF:
                debug_print("Detected UTF-8 BOM, stripping it before parsing JSON.")
                content = content.lstrip('\ufeff')
            try:
                data = json.loads(content)
                content_type = data.get("type")
                debug_print(f"Parsed JSON type: {content_type}, keys: {list(data.keys())}")
                if content_type == "CLIPBOARD_IMAGE":
                    image_data = base64.b64decode(data["data"])
                    debug_print(f"Decoded image data, length: {len(image_data)}")
                    return self.set_clipboard_image(image_data)
                elif content_type == "MULTI_FORMAT_CLIPBOARD":
                    if "formats" in data:
                        debug_print(f"MULTI_FORMAT_CLIPBOARD formats: {list(data['formats'].keys())}")
                        # Use format data directly - JSON parsing already handles Unicode escaping
                        return self.set_clipboard_multiple_formats(data["formats"])
                elif content_type == "PLAIN_TEXT":
                    if "content" in data:
                        debug_print(f"PLAIN_TEXT content: {repr(data['content'])}")
                        # Use content directly - JSON parsing already handles Unicode escaping
                        return self.set_clipboard_content(data["content"])
                    else:
                        debug_print("PLAIN_TEXT but no content field, setting empty string")
                        return self.set_clipboard_content("")
                elif content_type == "CLIPBOARD_FILES":
                    if "files" in data:
                        debug_print(f"CLIPBOARD_FILES: {list(data['files'].keys())}")
                        return self.set_clipboard_files(data["files"])
                    else:
                        debug_print("CLIPBOARD_FILES but no files field")
                        return False
            except Exception as e:
                print(f"DEBUG: Failed to parse JSON clipboard content: {e}")            
        except Exception as e:
            print(f"Error setting unified clipboard content: {e}")
            return self.set_clipboard_content(content)

    def get_current_clipboard_fingerprint(self):
        """Get current clipboard content fingerprint for comparison"""
        try:
            # Check for clipboard image first (highest priority)
            has_image = self.has_clipboard_image()
            if has_image:
                debug_print(f"Getting IMAGE fingerprint...")
                result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'image/png', '-o'], capture_output=True)
                if result.returncode == 0 and result.stdout:
                    current_image_hash = hashlib.md5(result.stdout).hexdigest()
                    fingerprint = json.dumps({
                        "type": "CLIPBOARD_IMAGE",
                        "format": "png",
                        "size": len(result.stdout),
                        "hash": current_image_hash
                    }, ensure_ascii=False, sort_keys=True)
                    debug_print(f"Image fingerprint: {fingerprint}")
                    return fingerprint
            
            # Check for file URIs (high priority, after images)
            has_files = self.has_clipboard_file_uris()
            if has_files:
                debug_print(f"Getting FILE URIs fingerprint...")
                file_uris = self.get_clipboard_file_uris()
                if file_uris:
                    # Create fingerprint based on file paths and their modification times
                    files_info = {}
                    for file_path in file_uris:
                        try:
                            if os.path.exists(file_path) and os.path.isfile(file_path):
                                stat = os.stat(file_path)
                                files_info[file_path] = {
                                    "size": stat.st_size,
                                    "mtime": stat.st_mtime
                                }
                        except:
                            files_info[file_path] = {"error": "inaccessible"}
                    
                    fingerprint = json.dumps({
                        "type": "CLIPBOARD_FILES", 
                        "files": files_info
                    }, ensure_ascii=False, sort_keys=True)
                    debug_print(f"File URIs fingerprint: {fingerprint}")
                    return fingerprint
            
            # Check for rich text formats (medium priority)
            has_rich = self.has_clipboard_rich_text()
            if has_rich:
                debug_print(f"Getting RICH TEXT fingerprint...")
                current_rich_formats = self.get_clipboard_formats()
                
                # Build current multi-format data for comparison
                format_priority = [
                    ('text/plain', '.txt'), ('text/html', '.html'), ('text/rtf', '.rtf'), 
                    ('application/rtf', '.app-rtf'), ('application/x-rtf', '.x-rtf'), 
                    ('text/richtext', '.richtext'), ('text/uri-list', '.uri'), 
                    ('text/x-moz-url', '.url')
                ]
                
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
                    # Create comparison string using content hashes for dynamic formats
                    comparison_data = {}
                    for fmt, content in current_format_data.items():
                        if fmt in ['text/html', 'text/rtf', 'application/rtf', 'application/x-rtf']:
                            # Normalize and hash dynamic formats to handle minor variations
                            if fmt == 'text/html':
                                # Normalize HTML by removing extra whitespace and line breaks
                                import re
                                normalized_content = re.sub(r'\s+', ' ', content.strip())
                                normalized_content = re.sub(r'>\s+<', '><', normalized_content)
                                content_hash = hashlib.md5(normalized_content.encode('utf-8')).hexdigest()
                                debug_print(f"HTML content normalized for hashing: length {len(content)} -> {len(normalized_content)}")
                            else:
                                # Use content hash for other dynamic formats
                                content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                            comparison_data[fmt] = content_hash
                        elif fmt == 'text/plain':
                            # Hash text/plain content for consistency with standalone plain text fingerprinting
                            content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                            comparison_data[fmt] = content_hash
                        else:
                            # Use actual content for other stable formats
                            comparison_data[fmt] = content
                    
                    fingerprint = json.dumps({"formats": comparison_data}, ensure_ascii=False, sort_keys=True)                    
                    
                    # Additional check: if we have the same plain text as before, don't re-capture
                    # This prevents duplicate captures when only HTML formatting metadata changes
                    if 'text/plain' in comparison_data:
                        current_plain_text = comparison_data['text/plain']
                        if hasattr(self, 'last_plain_text_content') and current_plain_text == self.last_plain_text_content:
                            debug_print(f"Plain text content unchanged, using previous fingerprint to prevent duplicate capture")
                            return self.last_clipboard_content if self.last_clipboard_content else fingerprint
                        self.last_plain_text_content = current_plain_text
                    
                    return fingerprint
            
            # Check for plain text (lowest priority)
            current_content = self.get_clipboard_content()
            if current_content:
                # Use hash for plain text content for consistency with other formats
                content_hash = hashlib.md5(current_content.encode('utf-8')).hexdigest()
                fingerprint = json.dumps({
                    "type": "PLAIN_TEXT",
                    "hash": content_hash
                }, ensure_ascii=False, sort_keys=True)
                # debug_print(f"Plain text fingerprint: {fingerprint}")
                return fingerprint
            
            debug_print(f"No clipboard content found")
            return ""
            
        except Exception as e:
            debug_print(f"Error getting clipboard fingerprint: {e}")
            return ""

    def check_copyq_daemon(self):
        """Check if copyq daemon is running and accepting connections"""
        try:
            # First check if copyq command exists
            result = subprocess.run(['which', 'copyq'], capture_output=True, text=True)
            if result.returncode != 0:
                debug_print("copyq command not found in PATH")
                return False
            
            # Test actual clipboard functionality - try to get clipboard content
            # This is a better test than 'copyq info' which doesn't validate daemon functionality
            result = subprocess.run(['copyq', 'clipboard'], capture_output=True, text=True, timeout=3)
            if result.returncode == 0:
                debug_print("copyq daemon clipboard test successful")
                return True
            else:
                debug_print(f"copyq daemon clipboard test failed with return code: {result.returncode}")
                debug_print(f"copyq stderr: {result.stderr}")
                return False
                    
        except subprocess.TimeoutExpired:
            debug_print("copyq daemon test timed out after 3 seconds")
            return False
        except FileNotFoundError:
            debug_print("copyq command not found")
            return False
        except Exception as e:
            debug_print(f"copyq daemon test failed: {e}")
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
        print(f"Maximum entries: {self.max_history} (older files will be automatically deleted)")
        print(f"Local sync file: {self.local_sync_file}")
        print(f"Local upload file (upload): {self.local_upload_file}")
        print(f"Remote peer files (download): {len(peer_files)} peer(s)")
        for peer_file in peer_files:
            print(f"  - {peer_file}")
        print(f"Remote check interval: {self.remote_check_interval} seconds")
        print("Supported formats: Multi-format Text, Images (unified JSON), HTML, RTF, URLs, Files")
        print("Note: First clipboard capture after startup will not be synced.")
        
        while True:
            try:
                # Import json at the top to avoid scope issues                                
                # Check all remote files for updates
                remote_content = self.check_all_remote_files_for_updates()
                # Note: last_clipboard_content is already updated inside check_all_remote_files_for_updates()
                
                # Debug clipboard detection
                has_image = self.has_clipboard_image()
                has_files = self.has_clipboard_file_uris()
                has_rich = self.has_clipboard_rich_text()
                debug_print(f"Clipboard detection: image={has_image}, files={has_files}, rich_text={has_rich}")
                
                # Get current clipboard fingerprint for comparison
                current_fingerprint = self.get_current_clipboard_fingerprint()
                debug_print(f"Current fingerprint: {repr(current_fingerprint) if current_fingerprint else 'None'}")
                debug_print(f"Last fingerprint: {repr(self.last_clipboard_content) if self.last_clipboard_content else 'None'}")
                debug_print(f"Fingerprints equal: {current_fingerprint == self.last_clipboard_content}")
                
                # Only process if fingerprint has changed and we have actual content
                if current_fingerprint and current_fingerprint != self.last_clipboard_content:
                    debug_print(f"Clipboard content changed, processing...")
                    
                    # Determine which save method to use based on content type
                    if has_image:
                        debug_print(f"Saving as IMAGE...")
                        self.save_clipboard_image()
                    elif has_files:
                        debug_print(f"Saving as FILES...")
                        self.save_clipboard_files()
                    elif has_rich:
                        debug_print(f"Saving as RICH TEXT...")
                        self.save_clipboard_rich_content()
                    else:
                        debug_print(f"Saving as PLAIN TEXT...")
                        current_text = self.get_clipboard_content()
                        if current_text:
                            self.save_clipboard_text(current_text)
                    
                    # Update the last known fingerprint
                    self.last_clipboard_content = current_fingerprint
                elif current_fingerprint:
                    debug_print(f"Clipboard content unchanged, skipping...")
                else:
                    debug_print(f"No clipboard content detected.")
                
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
