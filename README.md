# Data Analysis and Extraction Scripts

## What it does
This project consists of automated Python scripts designed to parse, extract, and analyze data from business or educational materials. Specifically, it processes PowerPoint presentations (`.pptx`) and Excel spreadsheets (`.xlsx`) to extract relevant text and tabular data (such as data and solutions for specific exercises like "Exercise 2.2").

## Problem it solves
Manually reading and comparing data across multiple slide decks and spreadsheet solutions can be tedious and error-prone. This project automates the extraction of slide text content directly from the underlying XML and loads Excel data into structured formats, making it significantly easier to review, analyze, and compare exercises and their solutions programmatically.

## Tech stack
* **Language:** Python 3.x
* **Libraries:** `pandas` (for data manipulation), `openpyxl` (for reading Excel files), `zipfile` & `re` (for parsing PPTX XML contents), `os` (for file management)


## Setup instructions
1. Ensure you have Python installed on your Windows machine.
2. Clone this repository or download the project files into a folder (e.g., `Desktop\China`).
3. Install the required Python packages:
   ```bash
   pip install pandas openpyxl
   ```
4. Ensure the relevant `.xlsx` and `.pptx` files are located in the exact same directory as the scripts.
5. Run any of the extraction scripts from your terminal, for example:
   ```bash
   python read_data.py
   python read_details.py
   python analyze_exercise.py
   ```

## Features
* **PPTX Text Extraction:** Reads raw XML from PowerPoint slides to extract text content without requiring the heavy `python-pptx` dependency.
* **Excel Data Parsing:** Leverages `pandas` to read multiple sheets, analyze shapes, and handle missing data.
* **Automated Log Generation:** Outputs the structured analysis and extracted text into generated files like `output8.txt` for easy reference.
* **Solution Verification Check:** Dynamically reads solution workbook sheet names and structures using `openpyxl` to compare with the given exercise data.

## Challenges & learnings
* **Learning:** Gained a deep understanding of reading PowerPoint slide contents by digging into the underlying XML structure of `.pptx` files using Python's native `zipfile` and `re` modules.
* **Challenge:** Handling robust Excel reading through `pandas` when sheet structures vary, requiring fallback mechanisms and metadata extraction using `openpyxl`.
