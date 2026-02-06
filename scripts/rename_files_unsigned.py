import os
import sys
import unicodedata
import re

def remove_accents(input_str):
    if not input_str:
        return ""
    s1 = unicodedata.normalize('NFD', input_str)
    s2 = ''.join(c for c in s1 if unicodedata.category(c) != 'Mn')
    s3 = re.sub(r'đ', 'd', s2)
    s4 = re.sub(r'Đ', 'D', s3)
    return s4

def rename_files_in_folder(folder_path):
    folder_path = os.path.abspath(folder_path)
    if not os.path.exists(folder_path):
        print(f"Directory not found: {folder_path}")
        return

    files = os.listdir(folder_path)
    print(f"Found {len(files)} files in {folder_path}")

    # Sort files to rename longer names first or just by name to avoid conflicts? 
    # Actually just iterating is fine if we check target existence
    
    for filename in files:
        # Skip directories
        if os.path.isdir(os.path.join(folder_path, filename)):
            continue
            
        # Get name and extension
        name, ext = os.path.splitext(filename)
        
        # Remove accents from name part
        new_name = remove_accents(name)
        
        # Clean up special characters characters that might be unsafe for web
        # Keep alphanumeric, spaces, hyphens, underscores, dots, parenthesis
        new_name = re.sub(r'[^a-zA-Z0-9\s\-\_\(\)\.]', '', new_name)
        
        # Collapse multiple spaces
        new_name = re.sub(r'\s+', ' ', new_name).strip()
        
        new_filename = new_name + ext
        
        if new_filename != filename:
            old_path = os.path.join(folder_path, filename)
            new_path = os.path.join(folder_path, new_filename)
            
            # Handle potential collision
            if os.path.exists(new_path):
                print(f"Skipping (Target exists): {filename} -> {new_filename}")
                continue
                
            try:
                os.rename(old_path, new_path)
                # Print mapping for JSON update manually if needed, or just list later
                # Print using repr to handle potential encoding display issues
                print(f"Renamed: {filename} -> {new_filename}")
            except Exception as e:
                print(f"Error renaming {filename}: {e}")
        else:
            print(f"No change needed: {filename}")

if __name__ == "__main__":
    # Force UTF-8 stdout
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
        
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]
    else:
        print("Usage: python rename_files_unsigned.py <folder_path>")
        sys.exit(1)
        
    rename_files_in_folder(target_folder)
