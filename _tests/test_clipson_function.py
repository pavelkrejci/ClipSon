#!/usr/bin/env python3
"""
Test the actual ClipSon set_clipboard_content_unified function
"""

import subprocess
import json
import os
import sys

# Add the current directory to the path to import clipson
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the ClipSon class (we'll need to set some basic properties)
from clipson import ClipSon

def test_clipson_set_content():
    print("Testing ClipSon set_clipboard_content_unified function...")
    
    # Original text that causes the problem
    original_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Original text: {repr(original_text)}")
    print(f"Original text (display): {original_text}")
    
    # Create the JSON that would be created by the sender
    upload_content = {
        "type": "PLAIN_TEXT",
        "content": original_text
    }
    upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
    print(f"\nJSON that would be uploaded:")
    print(upload_json)
    
    # Create a minimal ClipSon instance for testing
    clipson = ClipSon()
    
    # Set the clipboard using the unified function
    print(f"\n--- Setting clipboard using ClipSon function ---")
    result = clipson.set_clipboard_content_unified(upload_json)
    print(f"Function returned: {result}")
    
    if result:
        # Read back what was actually set
        try:
            clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                            capture_output=True, text=True)
            
            if clipboard_result.returncode == 0:
                retrieved_text = clipboard_result.stdout
                print(f"\nRetrieved from clipboard: {repr(retrieved_text)}")
                print(f"Retrieved from clipboard (display): {retrieved_text}")
                
                # Compare
                if original_text == retrieved_text:
                    print("✓ SUCCESS: Text matches exactly")
                else:
                    print("✗ FAILURE: Text does not match")
                    print(f"  Original:  {repr(original_text)}")
                    print(f"  Retrieved: {repr(retrieved_text)}")
                    
                    # Character by character comparison
                    print("  Character-by-character comparison:")
                    min_len = min(len(original_text), len(retrieved_text))
                    for i in range(min_len):
                        if original_text[i] != retrieved_text[i]:
                            print(f"    Position {i}: original={repr(original_text[i])}, retrieved={repr(retrieved_text[i])}")
                    
                    if len(original_text) != len(retrieved_text):
                        print(f"    Length difference: original={len(original_text)}, retrieved={len(retrieved_text)}")
            else:
                print("✗ Failed to read from clipboard")
                
        except Exception as e:
            print(f"✗ Error reading clipboard: {e}")
    else:
        print("✗ ClipSon function failed to set clipboard")

if __name__ == "__main__":
    test_clipson_set_content()
