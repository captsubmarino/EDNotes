import pyperclip
from bs4 import BeautifulSoup
import os
import time

def get_clipboard_with_retry(retries=5, delay=0.1):
    """Tries to access the clipboard multiple times before failing."""
    for i in range(retries):
        try:
            return pyperclip.paste()
        except pyperclip.PyperclipWindowsException as e:
            if "OpenClipboard" in str(e):
                print(f"Clipboard is busy, retrying in {delay}s... ({i+1}/{retries})")
                time.sleep(delay)
            else:
                raise # Re-raise other pyperclip errors
    # If all retries fail, raise the final error
    raise pyperclip.PyperclipWindowsException(
        "Failed to access clipboard after several retries. "
        "Another program (like Remote Desktop or a virtual machine) may be locking it."
    )

def clean_onenote_table(html_content):
    """Uses BeautifulSoup to clean HTML copied from OneNote."""
    print("Cleaning HTML...")
    soup = BeautifulSoup(html_content, 'lxml')
    table = soup.find('table')
    if not table: return None
    
    attributes_to_remove = ['style', 'class', 'lang', 'width', 'height', 'border', 'cellspacing', 'cellpadding']
    for tag in table.find_all(True):
        for attr in attributes_to_remove:
            if tag.has_attr(attr): del tag[attr]
    for p_tag in table.find_all('p'): p_tag.unwrap()
    for span_tag in table.find_all('span'): span_tag.unwrap()
    return table.prettify()

def create_html_file(cleaned_table_html, filename):
    """Wraps the cleaned table in a basic HTML5 boilerplate."""
    if not filename.lower().endswith('.html'):
        filename += '.html'

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{os.path.splitext(filename)[0]}</title>
    <style>
        body {{ font-family: sans-serif; margin: 2em; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
        th {{ background-color: #e0e0e0; }}
    </style>
</head>
<body>
    {cleaned_table_html}
</body>
</html>"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_template)
        print(f"✅ Successfully created '{filename}'")
    except Exception as e:
        print(f"❌ Error saving file: {e}")

if __name__ == "__main__":
    print("--- OneNote HTML Table Cleaner ---")
    try:
        clipboard_content = get_clipboard_with_retry()
        if not clipboard_content or not clipboard_content.strip().startswith('<'):
            print("❌ No HTML content found on clipboard. Please copy a table from OneNote first.")
        else:
            cleaned_table = clean_onenote_table(clipboard_content)
            if cleaned_table:
                output_filename = input("Enter a filename for the output (e.g., table1.html): ")
                if output_filename:
                    create_html_file(cleaned_table, output_filename)
                else:
                    print("No filename entered. Exiting.")
            else:
                print("❌ Could not find a <table> element in the clipboard content.")
    except Exception as e:
        print(f"❌ An error occurred: {e}")
