#!/usr/bin/env python3
"""
Test script to specifically test the backslash issue with ClipSon JSON handling
"""

import json
import subprocess

def test_clipson_json_handling():
    print("Testing ClipSon JSON backslash handling...")
    
    # Original text that would be copied to clipboard
    original_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Original text: {repr(original_text)}")
    print(f"Original text (display): {original_text}")
    
    # Step 1: Simulate what ClipSon sender does - create JSON
    upload_content = {
        "type": "PLAIN_TEXT",
        "content": original_text
    }
    
    # This is what gets saved to the JSON file
    upload_json = json.dumps(upload_content, ensure_ascii=False, indent=2)
    print(f"\nJSON created by sender:")
    print(upload_json)
    
    # Step 2: Simulate what ClipSon receiver does - parse JSON
    print(f"\n--- Simulating receiver side ---")
    
    # Parse the JSON (this is what happens when reading the remote file)
    try:
        data = json.loads(upload_json)
        content_type = data.get("type")
        print(f"Parsed content type: {content_type}")
        
        if content_type == "PLAIN_TEXT":
            if "content" in data:
                retrieved_content = data["content"]
                print(f"Retrieved content: {repr(retrieved_content)}")
                print(f"Retrieved content (display): {retrieved_content}")
                
                # Check if they match
                if original_text == retrieved_content:
                    print("✓ SUCCESS: Content matches exactly")
                else:
                    print("✗ FAILURE: Content does not match")
                    print(f"  Original:  {repr(original_text)}")
                    print(f"  Retrieved: {repr(retrieved_content)}")
                    
                    # Character by character comparison
                    print("  Character-by-character comparison:")
                    min_len = min(len(original_text), len(retrieved_content))
                    for i in range(min_len):
                        if original_text[i] != retrieved_content[i]:
                            print(f"    Position {i}: original={repr(original_text[i])}, retrieved={repr(retrieved_content[i])}")
                    
                    if len(original_text) != len(retrieved_content):
                        print(f"    Length difference: original={len(original_text)}, retrieved={len(retrieved_content)}")
                        
    except Exception as e:
        print(f"✗ ERROR parsing JSON: {e}")

def test_actual_clipboard():
    print(f"\n{'='*60}")
    print("Testing actual clipboard operations...")
    
    test_text = r"xsmbclient \\192.168.10.144\Users -U Martin"
    print(f"Test text: {repr(test_text)}")
    
    try:
        # Set clipboard
        result = subprocess.run(['xclip', '-selection', 'clipboard'], 
                              input=test_text, text=True, check=True)
        print("✓ Set to clipboard")
        
        # Read back
        result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                              capture_output=True, text=True)
        
        if result.returncode == 0:
            retrieved = result.stdout
            print(f"Retrieved from clipboard: {repr(retrieved)}")
            
            if test_text == retrieved:
                print("✓ Clipboard preserves backslashes correctly")
            else:
                print("✗ Clipboard altered the text")
        else:
            print("✗ Failed to read from clipboard")
            
    except Exception as e:
        print(f"✗ Error with clipboard: {e}")

if __name__ == "__main__":
    test_clipson_json_handling()
    test_actual_clipboard()
