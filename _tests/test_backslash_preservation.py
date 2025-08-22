#!/usr/bin/env python3
"""
Test script to verify backslash preservation in ClipSon
"""

import subprocess
import tempfile
import os
import time

def test_backslash_preservation():
    print("Testing ClipSon backslash preservation...")
    
    # Test cases with various backslash scenarios
    test_cases = [
        r"C:\Users\John\Documents\file.txt",  # Windows paths
        r"\\server\share\folder",             # UNC paths
        r"This is a \n newline escape",       # Escape sequences in text
        r"Regex pattern: \d+ \w* \s*",        # Regex patterns
        r"Math formula: f(x) = x\^2 + 2x",      # Mathematical expressions
        r"JSON: {\"key\": \"value\"}",        # JSON with escaped quotes
        r"C:\Program Files\App\config.json",  # Complex file path
        r"SELECT * FROM table WHERE col = '\''",  # SQL with escaped quotes
    ]
    
    for i, test_text in enumerate(test_cases):
        print(f"\n--- Test Case {i+1}: {test_text[:50]}{'...' if len(test_text) > 50 else ''} ---")
        
        # Set clipboard with test text using xclip
        try:
            result = subprocess.run(['xclip', '-selection', 'clipboard'], 
                                  input=test_text, text=True, check=True)
            print(f"✓ Set to clipboard: {repr(test_text)}")
            
            # Small delay to ensure clipboard is set
            time.sleep(0.1)
            
            # Read back from clipboard
            result = subprocess.run(['xclip', '-selection', 'clipboard', '-o'], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                retrieved_text = result.stdout
                print(f"✓ Retrieved from clipboard: {repr(retrieved_text)}")
                
                # Check if backslashes are preserved
                if test_text == retrieved_text:
                    print("✓ PASS: Backslashes preserved correctly")
                else:
                    print("✗ FAIL: Backslashes not preserved")
                    print(f"  Expected: {repr(test_text)}")
                    print(f"  Got:      {repr(retrieved_text)}")
                    
                    # Show character-by-character comparison for debugging
                    print("  Character comparison:")
                    for j, (expected, actual) in enumerate(zip(test_text, retrieved_text)):
                        if expected != actual:
                            print(f"    Position {j}: expected {repr(expected)}, got {repr(actual)}")
                            break
            else:
                print("✗ Failed to read from clipboard")
                
        except Exception as e:
            print(f"✗ Error in test case: {e}")
    
    print("\n" + "="*60)
    print("Test completed. Check results above.")

if __name__ == "__main__":
    test_backslash_preservation()
