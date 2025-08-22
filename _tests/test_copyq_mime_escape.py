#!/usr/bin/env python3
"""
Test copyq with multiple MIME types and backslashes
"""

import subprocess

def test_copyq_mime_types():
    text_with_backslashes = r"Path: C:\Users\test\file.txt"
    html_with_backslashes = f"<p>Path: C:\\Users\\test\\file.txt</p>"
    
    print(f"Text content: {repr(text_with_backslashes)}")
    print(f"HTML content: {repr(html_with_backslashes)}")
    
    # Test 1: Regular copyq with MIME types
    print(f"\n--- Test 1: Regular copyq with MIME types ---")
    try:
        result = subprocess.run(['copyq', 'copy', 'text/plain', text_with_backslashes, 'text/html', html_with_backslashes], check=True)
        print(f"copyq returned: {result.returncode}")
        
        # Read back plain text
        clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                        capture_output=True, text=True)
        
        if clipboard_result.returncode == 0:
            retrieved_text = clipboard_result.stdout
            print(f"Retrieved text: {repr(retrieved_text)}")
            
            if text_with_backslashes == retrieved_text:
                print("✓ Text content preserved")
            else:
                print("✗ Text content corrupted")
        
    except Exception as e:
        print(f"Error: {e}")
    
    # Test 2: copyq with -- and MIME types (probably won't work but let's try)
    print(f"\n--- Test 2: copyq with -- and MIME types ---")
    try:
        result = subprocess.run(['copyq', 'copy', '--', 'text/plain', text_with_backslashes, 'text/html', html_with_backslashes], check=True)
        print(f"copyq returned: {result.returncode}")
        
        # Read back plain text
        clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                        capture_output=True, text=True)
        
        if clipboard_result.returncode == 0:
            retrieved_text = clipboard_result.stdout
            print(f"Retrieved text: {repr(retrieved_text)}")
            
            if text_with_backslashes == retrieved_text:
                print("✓ Text content preserved")
            else:
                print("✗ Text content corrupted or unexpected format")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_copyq_mime_types()
