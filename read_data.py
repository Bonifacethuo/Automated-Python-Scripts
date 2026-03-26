import pandas as pd
import os
import zipfile
import re

directory = r"c:\Users\USER\Desktop\China"

# Read excel
excel_file = os.path.join(directory, "Exercise 2.2 (1).xlsx")
try:
    df_dict = pd.read_excel(excel_file, sheet_name=None)
    print("Excel Sheets:")
    for sheet_name, df in df_dict.items():
        print(f"\nSheet '{sheet_name}' shape: {df.shape}")
        print(df.head())
        # Print a bit more if there are empty rows
        print(df.dropna(how='all').head(10))
except Exception as e:
    print(f"Error reading excel: {e}")

print("\n--- PPTX Summaries ---")
for file_name in os.listdir(directory):
    if file_name.endswith('.pptx'):
        filepath = os.path.join(directory, file_name)
        print(f"\n--- {file_name} ---")
        try:
            with zipfile.ZipFile(filepath, 'r') as slide_zip:
                # Find all slides
                slide_files = [f for f in slide_zip.infolist() if f.filename.startswith('ppt/slides/slide') and f.filename.endswith('.xml')]
                for info in slide_files:
                    xml_content = slide_zip.read(info.filename).decode('utf-8')
                    # Very simple text extraction
                    text_matches = re.findall(r'<a:t>(.*?)</a:t>', xml_content)
                    slide_text = " ".join(text_matches)
                    if "Exercise" in slide_text or "exercise" in slide_text.lower() or "2.2" in slide_text:
                        print(f"  {info.filename} (Potential Exercise Text): {slide_text}")
        except Exception as e:
            print(f"Error reading {file_name}: {e}")
