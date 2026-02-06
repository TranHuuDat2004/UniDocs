import os
import sys
import comtypes.client

# Force UTF-8 encoding for stdout to prevent cp1252 errors on Windows
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def convert_folder(folder_path):
    folder_path = os.path.abspath(folder_path)
    if not os.path.exists(folder_path):
        print(f"Directory not found: {folder_path}")
        return

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0

    try:
        files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
        print(f"Found {len(files)} Word files in {folder_path}")

        for filename in files:
            file_path = os.path.join(folder_path, filename)
            file_name_no_ext = os.path.splitext(filename)[0]
            pdf_path = os.path.join(folder_path, file_name_no_ext + ".pdf")

            if os.path.exists(pdf_path):
                print(f"Skipping (PDF exists): {filename}")
                continue

            print(f"Converting: {filename}...")
            try:
                doc = word.Documents.Open(file_path)
                # 17 = wdFormatPDF
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()
                print(f"Success: {pdf_path}")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")
                
    finally:
        word.Quit()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]
    else:
        print("Usage: python convert_docx_to_pdf.py <folder_path>")
        sys.exit(1)
        
    convert_folder(target_folder)
