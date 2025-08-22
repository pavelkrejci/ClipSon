#!/usr/bin/env python3
"""
Test the complete ClipSon workflow for backslash preservation
"""

import subprocess
import json
import os
import sys
import tempfile

# Add the current directory to the path to import clipson
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the ClipSon class
from clipson import ClipSon

def test_complete_workflow():
    print("Testing complete ClipSon workflow for backslash preservation...")
    
    # Test cases with problematic backslash content
    test_cases = [
        r"xsmbclient \\192.168.10.144\Users -U Martin",
        r"C:\Users\John\Documents\file.txt",
        r"echo 'This is a \n newline escape'",
        r"Regex: \d+ \w* \s*",
    ]
    
    clipson = ClipSon()
    
    for i, original_text in enumerate(test_cases):
        print(f"\n{'='*60}")
        print(f"Test Case {i+1}: {original_text}")
        print(f"{'='*60}")
        
        # Step 1: Simulate sender - create JSON like ClipSon would
        upload_content = {
            "type": "PLAIN_TEXT",
            "content": original_text
        }
        upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
        print(f"1. JSON created by sender:")
        print(f"   Content: {repr(upload_content['content'])}")
        
        # Step 2: Simulate file transfer - write to temp file and read back
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', delete=False, suffix='.json') as f:
            f.write(upload_json)
            temp_file = f.name
        
        try:
            with open(temp_file, 'r', encoding='utf-8') as f:
                received_json = f.read()
            
            print(f"2. JSON received by peer (matches): {upload_json == received_json}")
            
            # Step 3: Parse JSON and apply to clipboard
            result = clipson.set_clipboard_content_unified(received_json)
            print(f"3. ClipSon function returned: {result}")
            
            if result:
                # Step 4: Read back from clipboard
                clipboard_result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                                capture_output=True, text=True)
                
                if clipboard_result.returncode == 0:
                    final_content = clipboard_result.stdout
                    print(f"4. Final clipboard content: {repr(final_content)}")
                    
                    # Step 5: Verify complete round-trip
                    if original_text == final_content:
                        print("✓ SUCCESS: Complete workflow preserves backslashes")
                    else:
                        print("✗ FAILURE: Backslashes lost in workflow")
                        print(f"   Original: {repr(original_text)}")
                        print(f"   Final:    {repr(final_content)}")
                else:
                    print("✗ Failed to read final clipboard content")
            else:
                print("✗ ClipSon function failed")
                
        finally:
            # Clean up
            try:
                os.unlink(temp_file)
            except:
                pass
    
    print(f"\n{'='*60}")
    print("Complete workflow test finished")

if __name__ == "__main__":
    test_complete_workflow()
