import os
import sys
import comtypes.client

def convert_folder(folder_path):
    folder_path = os.path.abspath(folder_path)
    if not os.path.exists(folder_path):
        print(f"Directory not found: {folder_path}")
        return

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.ppt', '.pptx'))]
        print(f"Found {len(files)} PowerPoint files in {folder_path}")

        for filename in files:
            file_path = os.path.join(folder_path, filename)
            file_name_no_ext = os.path.splitext(filename)[0]
            pdf_path = os.path.join(folder_path, file_name_no_ext + ".pdf")

            if os.path.exists(pdf_path):
                print(f"Skipping (PDF exists): {filename}")
                continue

            print(f"Converting: {filename}...")
            try:
                deck = powerpoint.Presentations.Open(file_path)
                # 32 = ppSaveAsPDF
                deck.SaveAs(pdf_path, 32)
                deck.Close()
                print(f"Success: {pdf_path}")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")
                
    finally:
        powerpoint.Quit()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]
    else:
        # Default or prompt
        print("Usage: python convert_to_pdf.py <folder_path>")
        sys.exit(1)
        
    convert_folder(target_folder)
