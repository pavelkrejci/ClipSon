#!/usr/bin/env python3
"""
Test copyq vs xclip directly
"""

import subprocess

def test_clipboard_tools():
    original_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Original text: {repr(original_text)}")
    print(f"Original text (display): {original_text}")
    
    print(f"\n--- Testing copyq ---")
    try:
        # Test copyq
        result = subprocess.run(['copyq', 'copy', 'text/plain', original_text], check=True)
        print(f"copyq returned: {result.returncode}")
        
        # Read back with xclip
        clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                        capture_output=True, text=True)
        
        if clipboard_result.returncode == 0:
            retrieved_text = clipboard_result.stdout
            print(f"Retrieved from clipboard: {repr(retrieved_text)}")
            print(f"Retrieved from clipboard (display): {retrieved_text}")
            
            if original_text == retrieved_text:
                print("✓ copyq preserves backslashes")
            else:
                print("✗ copyq corrupts backslashes")
        
    except Exception as e:
        print(f"Error with copyq: {e}")
    
    print(f"\n--- Testing xclip ---")
    try:
        # Test xclip
        result = subprocess.run(['xclip', '-selection', 'clipboard'], 
                               input=original_text, text=True, check=True)
        print(f"xclip returned: {result.returncode}")
        
        # Read back with xclip
        clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                        capture_output=True, text=True)
        
        if clipboard_result.returncode == 0:
            retrieved_text = clipboard_result.stdout
            print(f"Retrieved from clipboard: {repr(retrieved_text)}")
            print(f"Retrieved from clipboard (display): {retrieved_text}")
            
            if original_text == retrieved_text:
                print("✓ xclip preserves backslashes")
            else:
                print("✗ xclip corrupts backslashes")
        
    except Exception as e:
        print(f"Error with xclip: {e}")

if __name__ == "__main__":
    test_clipboard_tools()
