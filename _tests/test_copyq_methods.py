#!/usr/bin/env python3
"""
Test different copyq usage methods
"""

import subprocess

def test_copyq_methods():
    original_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Original text: {repr(original_text)}")
    print(f"Original text (display): {original_text}")
    
    methods = [
        ("Method 1: copyq copy text/plain", ['copyq', 'copy', 'text/plain', original_text]),
        ("Method 2: copyq copy (just text)", ['copyq', 'copy', original_text]),
    ]
    
    for method_name, command in methods:
        print(f"\n--- {method_name} ---")
        try:
            result = subprocess.run(command, check=True)
            print(f"Command returned: {result.returncode}")
            
            # Read back with xclip
            clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                            capture_output=True, text=True)
            
            if clipboard_result.returncode == 0:
                retrieved_text = clipboard_result.stdout
                print(f"Retrieved: {repr(retrieved_text)}")
                print(f"Display:   {retrieved_text}")
                
                if original_text == retrieved_text:
                    print("✓ Backslashes preserved")
                else:
                    print("✗ Backslashes corrupted")
            
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    test_copyq_methods()
