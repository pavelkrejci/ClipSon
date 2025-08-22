#!/usr/bin/env python3
"""
Test copyq with -- option to disable escape sequence expansion
"""

import subprocess

def test_copyq_with_no_escape():
    original_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Original text: {repr(original_text)}")
    print(f"Original text (display): {original_text}")
    
    print(f"\n--- Testing copyq copy -- (no escape expansion) ---")
    try:
        # Use -- to disable escape sequence expansion
        result = subprocess.run(['copyq', 'copy', '--', original_text], check=True)
        print(f"Command returned: {result.returncode}")
        
        # Read back with xclip
        clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                        capture_output=True, text=True)
        
        if clipboard_result.returncode == 0:
            retrieved_text = clipboard_result.stdout
            print(f"Retrieved: {repr(retrieved_text)}")
            print(f"Display:   {retrieved_text}")
            
            if original_text == retrieved_text:
                print("✓ SUCCESS: Backslashes preserved with copyq copy --")
            else:
                print("✗ FAILURE: Backslashes still corrupted")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_copyq_with_no_escape()
