# 📄 PDF Data Extraction

## 📋 Introduction
This script provides a **scalable and reusable** solution for **extracting structured data** from PDFs and **annotating PDFs** for offset identification as part of the interview task for the WerkStudent in python.

It automates the process of:
- Extracting **structured data** from PDF invoices.
- **Standardizing values** (currency formatting, date conversion, USD to EUR conversion).
- **Saving extracted data** in structured formats like **Excel (with Pivot Table)** & **CSV**.
- **Marking indexed text** in PDFs for easy reference.
- Allowing users to configure extractions dynamically using **JSON configurations**.
- Works for **multi-page PDFs** and supports **dynamic field extraction**.
---
## 📌 **Extracted Data: Excel & CSV Files**

📊 **1. Excel File (`output.xlsx`)**
The script generates an **Excel file** with two sheets:

**Sheet 1: Raw Extracted Data**
| File Name               | Date        | Original Value | Value in EUR |
|-------------------------|------------|---------------|--------------|
| sample_invoice_1.pdf    | 2024-03-01 | 453,53 €      | 453.53      |
| sample_invoice_2.pdf    | 2016-11-26 | 950.00 $   | 921.50     |

📌 **What this table shows**:
- **"Original Value"** retains the extracted value in its original format.
- **"Value in EUR"** converts USD to EUR using a fixed exchange rate (i.e. 1 USD = 0.97 EUR).


**Sheet 2: Pivot Table (Summarized Data)**
The pivot table helps **summarize** extracted values **by date and document**.

| Date       | sample_invoice_1.pdf | sample_invoice_2.pdf | Total |
|------------|----------------------|----------------------|-------------|
| 2016-11-26 | 0 €                  | 921.50 €            | 921.50 €    |
| 2024-03-01 | 453.53 €              | 0 €                 | 453.53 €    |

📌 **Key Points**:
- **Summarizes** invoice values by **date**.
- **Automatically converts** all values to EUR.
- Adds a `"Total"` column.

---

📂 **2. CSV File (`output.csv`)**
A **CSV file** is also generated, with the same data as **Sheet 1 in Excel**, but uses **a semicolon (`;`) as the delimiter**.

Example content of `output.csv`:
```csv
File Name;Date;Original Value;Value in EUR 
sample_invoice_1.pdf;2024-03-01;453,53 €;453.53 
sample_invoice_2.pdf;2016-11-26;950.00 $;921.50
```
---

## 📚 Table of Contents
1. [Features](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#-features)
2. [Project Structure](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#-project-structure)
3. [Installation & Running the Code](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#-installation--running-the-code)
4. [Building Executable from Source](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#%EF%B8%8F-building-executable-from-source)
5. [Code Structure & How it Works](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#-code-structure--how-it-works)
6. [JSON Configuration](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#3%EF%B8%8F-pdf_configjson)
7. [Troubleshooting](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#-troubleshooting)
8. [Limitations](https://github.com/sleipnir029/Rahat_Rakibuzzaman_WerkStudent_Python?tab=readme-ov-file#%EF%B8%8F-limitations)

---

## 🚀 Features
- **Dynamic PDF Data Extraction**: Extracts predefined values from PDFs based on `pdf_config.json`.
- **PDF Annotation**: Generates PDFs with **green-marked text** and **index numbers**, making it easy to adjust `offset` values.
- **Excel & CSV Output**: Saves extracted data into structured formats.
- **Pivot Table Generation**: Automatically generates a summary table in Excel.
- **Multi-page PDF Support**: Can extract data from **specific pages** as defined in JSON.
- **Configurable Extraction**: Modify `pdf_config.json` to adjust extraction fields without changing code.
- **Runs as Standalone Executable (.exe)**: No need for Python installation.

---

## 📂 Project Structure
```
├── pdf_to_text_extraction.py  # Extracts data from PDFs
├── pdf_to_text_extraction.exe # Executable file
├── pdf_annotation.py          # Annotates PDFs for easy offset identification
├── pdf_config.json            # JSON config file for dynamic extraction
├── sample_invoice_1.pdf       # Sample invoice
├── sample_invoice_2.pdf       # Sample invoice
├── sample_output/             # Provided examples of the Excel and CSV format
├── output.xlsx                # Extracted data in Excel format
├── output.csv                 # Extracted data in CSV format
└── annotated_documents/       # Folder containing annotated PDFs
```

---

## 🔧 Installation & Running the Code

### 1️⃣ Install Dependencies
Ensure you have **Python 3.10+** and install required packages:
```bash
pip install -r requirements.txt
```

### 2️⃣ Running the **PDF Data Extraction Script**
```bash
python pdf_to_text_extraction.py
```
📂 **Output Files Generated:**
- `output.xlsx` (Excel with Pivot Table)
- `output.csv` (CSV with extracted values)

### 3️⃣ Running the **PDF Annotation Script**
```bash
python pdf_annotation.py --output_folder annotated_documents
```
📂 **Annotated PDFs Saved in:**
```
annotated_documents/sample_invoice_1_annotated.pdf
annotated_documents/sample_invoice_2_annotated.pdf
```
These annotated PDFs help users determine **correct offsets** when configuring `pdf_config.json`.

---

## 🏗️ Building Executable from Source

### For **Windows**:
```bash
pyinstaller --onefile --add-data "pdf_config.json;." pdf_to_text_extraction.py
```

### For **Linux** and **macOS (ARM64)**:
```bash
pyinstaller --onefile --add-data "pdf_config.json:." pdf_to_text_extraction.py
```
💡 **Note:** The provided `.exe` file works only for macOS **ARM64 (M1/M2 chips)**.

---

## 📜 Code Structure & How it Works

### **1️. pdf_to_text_extraction.py**
- Reads PDFs **line-by-line**.
- Extracts predefined values using **keywords and offsets** from `pdf_config.json`.
- Standardizes **dates** (`YYYY-MM-DD` format) and **currency values** including USD to EUR conversion.
- Saves extracted data into **Excel and CSV**.

### **2️. pdf_annotation.py**
- Reads PDFs and marks **indexed text** inside a **green box**.
- Places **index numbers** as superscripts.
- Generates an **Text Index Table** listing all detected values.

The script **`pdf_annotation.py`** provides several command-line arguments to **customize** the annotation process. These parameters allow users to adjust colors, font sizes, spacing, and other configurations when marking up PDFs.

**📝 List of CLI Parameters**

| **Parameter**             | **Type** | **Default Value**       | **Description** |
|---------------------------|----------|-------------------------|----------------|
| `--output_folder`         | `str`    | `"annotated_documents"` | Folder where annotated PDFs will be saved. |
| `--rectangle_color`       | `float`  | `[0, 1, 0]` (Green)     | RGB values for bounding boxes around text. |
| `--rectangle_line_width`  | `float`  | `0.5`                   | Line thickness of bounding boxes. |
| `--index_color`           | `float`  | `[1, 0, 0]` (Red)       | RGB values for index number annotations. |
| `--index_font_size`       | `int`    | `8`                     | Font size for index numbers. |
| `--index_offset`          | `int`    | `[5, -5]`               | Offset for positioning index numbers (superscript). |
| `--table_font_size`       | `int`    | `8`                     | Font size for Text Index Table at the beginning of PDF. |
| `--line_spacing`          | `int`    | `12`                    | Line spacing in the Text Index Table. |
| `--max_entries_per_page`  | `int`    | `50`                    | Maximum entries before a new Text Index Table page is created. |

 **Run the script with default parameters:**
```bash
python pdf_annotation.py
```

**Customize output folder, colors, font sizes, and more:**
```bash
python pdf_annotation.py --output_folder "custom_annotated_pdfs" --rectangle_color 0 0 1 --index_color 1 0.5 0 --index_font_size 10 --table_font_size 12 --line_spacing 15 --max_entries_per_page 30
```
**What this command does:**
- Saves annotated PDFs to `"custom_annotated_pdfs"`
- Sets blue bounding boxes (`rectangle_color 0 0 1`)
- Sets orange index numbers (`index_color 1 0.5 0`)
- Increases `font sizes` and `spacing` in the Text Index Table.

### **3️. pdf_config.json**
📌 **Understanding `pdf_config.json`**

The **`pdf_config.json`** file is a key component that tells the script **what data** to extract from each PDF. Users can specify the **keywords**, the **offset (relative position of values)**, and the **page number** to scan.


### 📝 **JSON Configuration**
Here’s an example `pdf_config.json`:

```json
{
    "sample_invoice_1.pdf": {
        "Date": { "keyword": "Date", "offset": 1, "page": 1 },
        "Value": { "keyword": "Gross Amount incl. VAT", "offset": 0, "page": 1 }
    },
    "sample_invoice_2.pdf": {
        "Date": { "keyword": "Invoice Date", "offset": 0, "page": 1 },
        "Value": { "keyword": "Total", "offset": 0, "page": 1 }
    }
}
```
🔍 **Breaking Down the JSON Keys**
|**Key**	  | **Description** |
|-------------|-----------------------------------------------------------------------------|
| PDF Filename|	Each section starts with the name of the PDF file (e.g., "sample_invoice_1.pdf").|
|Field Name   |	Inside each PDF section, fields like "Date" and "Value" define what data to extract.|
|`keyword`    |	The exact text that appears in the PDF. The script looks for this keyword.|
|`offset`     |	The relative position of the value to extract after finding the keyword.|
|`page`       |	The page number where the script should search for the keyword.|

🔍 **Understanding the `offset` Value**
The **offset** determines how far below the keyword the required value is.

📌 Example 1 – `offset: 1`
```json
"Date": {
  "keyword": "Date",
  "offset": 1,
  "page": 1
}
```
- The script finds the word `"Date"` on page `1`.
- It extracts the text one line below (`offset: 1`).
- Useful for cases like:
```python
Date:
1. März 2024  ← Extract this line (offset 1)
```
📌 Example 2 – `offset: 0`
```json
"Value": {
  "keyword": "Total",
  "offset": 0,
  "page": 1
}
```
- The script finds the word `"Total"` and extracts the same line (`offset: 0`).
- Useful for cases like:
```python
Total: USD $950.00  ← Extract this text (offset 0)
```
🔍 **Understanding the `page` Parameter**
|**Value**	|**Effect**|
|-----------|----------------------------------|
|`1`        |	Extracts data only from Page 1.|
|`2`        |	Extracts data only from Page 2.|

🛠 **Modifying `pdf_config.json` for New Data Extraction**
If a user wants to extract new fields, they only need to update `pdf_config.json` and do not need to modify the script. For this, they can check the `annotated_documents` folder for the desired annotated PDF file running the `pdf_annotation.py` script and check both `text region` and `index value` along with the `Text Index Table` at the beginning of the annotated PDF.

✅ **Example: Adding a New Field**
If a PDF contains `"Tax Amount"`, and the value appears two lines below, modify `pdf_config.json`:
```json
{
    "sample_invoice_2.pdf": {
        "Due Date": { 
            "keyword": "Payment due:",
            "offset": 1,
            "page": 1
        }
    }
}
```
This will result in the following extraction which can be seen from the `Text Index Table` of the annotated PDF of `sample_invoice_2_annotated.pdf`
```python
Payment due:
30 days after invoice date  ← Extract this line (offset 1)
```

---

## 🔍 Troubleshooting

| Issue | Possible Solution |
|--------------------------------------|--------------------------------------------------------|
| No PDFs found | Ensure PDFs are in the same folder as the script |
| Configuration file not found | Place `pdf_config.json` in the script/executable folder |
| Extracted values are incorrect | Adjust `offset` and `page` values in `pdf_config.json` |
| Pivot Table is missing in Excel | Ensure Excel automatically refreshes pivot tables |

---

## ⚠️ Limitations
1. **Executable is limited to macOS (ARM64) systems.**
2. **While the code is dynamic, additional cleanup functions are required** for extracting new fields in a structured format.
3. **Multi-page PDFs need `page` values specified in `pdf_config.json`** for precise extraction.

---

## 🔄 Next Steps
- Extend support for **Windows & Linux executables**.
- Add **more dynamic cleanup functions** for extracted data.

**✅ Ready to extract data from your PDFs effortlessly! 🚀**
