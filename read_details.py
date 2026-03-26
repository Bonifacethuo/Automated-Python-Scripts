import pandas as pd
import os
import zipfile
import re

directory = r"c:\Users\USER\Desktop\China"

# Read excel
excel_file = os.path.join(directory, "Exercise 2.2 (1).xlsx")
df = pd.read_excel(excel_file, sheet_name='Data')
print("Columns:", df.columns.tolist())
print(df.head())

# Also read Exercise 2.2 Solution (1).xlsx to see what's expected
try:
    sol_file = os.path.join(directory, "Exercise 2.2 Solution (1).xlsx")
    sol_df = pd.read_excel(sol_file, sheet_name=None)
    for sheet_name, sheet_df in sol_df.items():
        print(f"\nSolution Sheet '{sheet_name}':")
        print(sheet_df.head(10))
except Exception as e:
    print(f"Error reading solution: {e}")

# Read PPTX "Brand Performance Measures (1).pptx"
pptx_file = os.path.join(directory, "Brand Performance Measures (1).pptx")
print(f"\n--- {pptx_file} ---")
try:
    with zipfile.ZipFile(pptx_file, 'r') as slide_zip:
        slide_files = [f for f in slide_zip.infolist() if f.filename.startswith('ppt/slides/slide') and f.filename.endswith('.xml')]
        # Sort slides by number (slide1.xml, slide2.xml etc)
        slide_files.sort(key=lambda x: int(re.search(r'slide(\d+)\.xml', x.filename).group(1)))
        
        for info in slide_files:
            xml_content = slide_zip.read(info.filename).decode('utf-8')
            text_matches = re.findall(r'<a:t>(.*?)</a:t>', xml_content)
            slide_text = " ".join(text_matches)
            if slide_text.strip():
                print(f"Slide {re.search(r'slide(\d+)', info.filename).group(1)}: {slide_text.strip()}")
except Exception as e:
    print(f"Error reading pptx: {e}")
