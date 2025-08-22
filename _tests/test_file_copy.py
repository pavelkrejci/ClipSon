#!/usr/bin/env python3
"""
Test script for file copy functionality in ClipSon
"""

import subprocess
import tempfile
import os
import time

def test_file_copy():
    print("Testing ClipSon file copy functionality...")
    
    # Create a test file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as test_file:
        test_file.write("This is a test file for ClipSon file copy functionality.\nIt contains multiple lines.\n")
        test_file_path = test_file.name
    
    try:
        print(f"Created test file: {test_file_path}")
        
        # Copy file to clipboard using file manager simulation
        file_uri = f"file://{test_file_path}"
        print(f"Setting clipboard with file URI: {file_uri}")
        
        # Set clipboard with file URI using xclip
        result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'text/uri-list'], 
                              input=file_uri, text=True)
        
        if result.returncode == 0:
            print("✓ File URI set to clipboard successfully")
            
            # Check if clipboard contains the file URI
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'text/uri-list', '-o'], 
                                  capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                print(f"✓ Clipboard contains URI: {result.stdout.strip()}")
            else:
                print("✗ Failed to read URI from clipboard")
                
            # Check clipboard formats
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                formats = result.stdout.strip().split('\n')
                print(f"Available clipboard formats: {formats}")
                print(f"Has text/uri-list: {'text/uri-list' in formats}")
            
        else:
            print("✗ Failed to set file URI to clipboard")
            
    finally:
        # Clean up test file
        if os.path.exists(test_file_path):
            os.unlink(test_file_path)
            print(f"Cleaned up test file: {test_file_path}")

if __name__ == "__main__":
    test_file_copy()
