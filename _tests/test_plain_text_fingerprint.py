#!/usr/bin/env python3

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from clipson import ClipSon, CONFIG_DATA
import subprocess

def test_plain_text_fingerprint():
    # Create ClipSon instance
    clipson = ClipSon()
    
    # Set test text to clipboard
    test_text = "Test plain text content for fingerprinting"
    subprocess.run(['xclip', '-selection', 'clipboard'], input=test_text, text=True, check=True)
    
    # Get fingerprint
    fingerprint = clipson.get_current_clipboard_fingerprint()
    print(f"Plain text fingerprint: {fingerprint}")
    
    # Test that it's JSON and contains hash
    import json
    try:
        data = json.loads(fingerprint)
        if data.get("type") == "PLAIN_TEXT" and "hash" in data:
            print("✓ Plain text fingerprint format is correct")
            print(f"  Type: {data['type']}")
            print(f"  Hash: {data['hash']}")
        else:
            print("✗ Plain text fingerprint format is incorrect")
            print(f"  Data: {data}")
    except json.JSONDecodeError:
        print("✗ Fingerprint is not valid JSON")
        print(f"  Raw: {fingerprint}")

if __name__ == "__main__":
    test_plain_text_fingerprint()
