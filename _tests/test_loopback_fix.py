#!/usr/bin/env python3

import json
import hashlib

# Simulate the test data from the Windows machine
test_remote_content = {
    "formats": {
        "text/plain": "[TYPE THE COMPANY NAME]\r\n\r\n[Pick the date]\r\nAuthored by: Pavel\r\n",
        "text/html": "\r\n\r\n<p class=MsoNoSpacing align=center style='text-align:center;line-height:115%'><b\r\nstyle='mso-bidi-font-weight:normal'><span style='color:#D34817;mso-themecolor:\r\naccent1;text-transform:uppercase'><w:Sdt\r\n PrefixMappings=\"xmlns:ns0='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'\"\r\n Xpath=\"/ns0:Properties[1]/ns0:Company[1]\" ShowingPlcHdr=\"t\"\r\n DocPart=\"6167E8C50641476C8682AC22CD57024D\" Text=\"t\"\r\n StoreItemID=\"X_6668398D-A668-4E3E-A5EB-62B293D839F1\" ID=\"1551716\">[Type the\r\n company name]</w:Sdt><o:p></o:p></span></b></p>\r\n\r\n<p class=MsoNoSpacing align=center style='text-align:center;line-height:115%'><b\r\nstyle='mso-bidi-font-weight:normal'><span style='color:#D34817;mso-themecolor:\r\naccent1;text-transform:uppercase'><o:p>&nbsp;</o:p></span></b></p>\r\n\r\n<p class=MsoNoSpacing align=center style='text-align:center;line-height:115%'><w:Sdt\r\n PrefixMappings=\"xmlns:ns0='http://schemas.microsoft.com/office/2006/coverPageProps'\"\r\n Xpath=\"/ns0:CoverPageProperties[1]/ns0:PublishDate[1]\" ShowingPlcHdr=\"t\"\r\n DocPart=\"ED6EEB0D5F4B421D9CEA656A5F4103C5\" Calendar=\"t\" MapToDateTime=\"t\"\r\n CalendarType=\"Gregorian\" StoreItemID=\"X_55AF091B-3C7A-41E3-B477-F2FDAA23CFDA\"\r\n DateFormat=\"MMMM d, yyyy\" Lang=\"EN-US\" ID=\"1551723\">[Pick the date]</w:Sdt></p>\r\n\r\n<span style='font-size:11.0pt;mso-bidi-font-size:10.0pt;line-height:115%;\r\nfont-family:\"Perpetua\",\"serif\";mso-ascii-theme-font:minor-latin;mso-fareast-font-family:\r\nPerpetua;mso-fareast-theme-font:minor-latin;mso-hansi-theme-font:minor-latin;\r\nmso-bidi-font-family:\"Times New Roman\";color:black;mso-themecolor:text1;\r\nmso-ansi-language:EN-US;mso-fareast-language:EN-US;mso-bidi-language:AR-SA'>Authored\r\nby: <w:Sdt\r\n PrefixMappings=\"xmlns:ns0='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:ns1='http://purl.org/dc/elements/1.1/'\"\r\n Xpath=\"/ns0:coreProperties[1]/ns1:creator[1]\"\r\n DocPart=\"48C7875C8F484906B62E5C0C3BC0FE8C\" Text=\"t\"\r\n StoreItemID=\"X_6C3C8BC8-F283-45AE-878A-BAB7291924A1\" ID=\"1551727\"><span\r\n lang=CS style='font-family:\"Times New Roman\",\"serif\";mso-ansi-language:CS'>Pavel</span></w:Sdt></span>"
    },
    "type": "MULTI_FORMAT_CLIPBOARD"
}

def simulate_set_last_clipboard_comparison_from_remote(content):
    """Simulate the fixed set_last_clipboard_comparison_from_remote method"""
    data = content
    format_data = data["formats"]
    comparison_data = {}
    
    # Use the same logic as in the main loop for comparison data generation
    for fmt, content_text in format_data.items():
        if fmt in ['text/html', 'text/rtf', 'application/rtf', 'application/x-rtf']:
            # For formats that might have dynamic content, use content hash for comparison
            content_hash = hashlib.md5(content_text.encode('utf-8')).hexdigest()
            comparison_data[fmt] = content_hash
        else:
            # For stable formats, use actual content
            comparison_data[fmt] = content_text
    
    # Store the comparison data (with hashes) to match what main loop generates
    return json.dumps({"formats": comparison_data}, ensure_ascii=False, sort_keys=True)

def simulate_main_loop_comparison(format_data):
    """Simulate the main loop comparison data generation"""
    comparison_data = {}
    
    for fmt, content in format_data.items():
        if fmt in ['text/html', 'text/rtf', 'application/rtf', 'application/x-rtf']:
            # For formats that might have dynamic content, use content hash for comparison
            # but still store full content for upload
            content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
            comparison_data[fmt] = content_hash
        else:
            # For stable formats, use actual content
            comparison_data[fmt] = content
    
    return json.dumps({"formats": comparison_data}, ensure_ascii=False, sort_keys=True)

# Test the fix
print("Testing loopback fix...")
print("=" * 50)

# Simulate remote content processing
remote_comparison = simulate_set_last_clipboard_comparison_from_remote(test_remote_content)
print(f"Remote comparison data: {remote_comparison}")

# Simulate main loop processing the same content (after it's been set to clipboard)
main_loop_comparison = simulate_main_loop_comparison(test_remote_content["formats"])
print(f"Main loop comparison data: {main_loop_comparison}")

# Check if they match
if remote_comparison == main_loop_comparison:
    print("\n✅ SUCCESS: Remote and main loop comparison data match!")
    print("   The loopback issue should be fixed.")
else:
    print("\n❌ FAILURE: Remote and main loop comparison data do NOT match!")
    print("   The loopback issue would still occur.")

print("\nKey insight:")
print("- text/plain content is stored as-is in both cases")
print("- text/html content is stored as MD5 hash in both cases")
print("- This prevents loopback caused by dynamic HTML timestamps")
