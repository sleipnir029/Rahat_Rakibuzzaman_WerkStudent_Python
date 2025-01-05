#### Import necessary libraries ####
import fitz  
import pandas as pd
from datetime import datetime
import json
import os
import sys
import re


#### Constant values ####

# Exchange rate (Example: 1 USD = 0.97 EUR)
USD_TO_EUR = 0.97   # At the time of writing  

# Dictionary to translate German month names to English
GERMAN_MONTHS = {
    "Januar": "January",
    "Februar": "February",
    "März": "March",
    "April": "April",
    "Mai": "May",
    "Juni": "June",
    "Juli": "July",
    "August": "August",
    "September": "September",
    "Oktober": "October",
    "November": "November",
    "Dezember": "December"
}




####  Extrating value/text by keyword ####
def extract_value_by_keyword(file_path, keyword, offset=0, target_page=None):
    """Extract a value from PDFs by looking for a keyword in a specific page (if provided)."""
    try:
        doc = fitz.open(file_path)

        # If a specific page is given, restrict search to that page
        pages = [doc[target_page - 1]] if target_page else doc  # Convert 1-based page number to 0-based index
        
        for page in pages:
            text = page.get_text("text")
            lines = text.splitlines()
            
            for i, line in enumerate(lines):
                if keyword in line:
                    if i + offset < len(lines):
                        return lines[i + offset].strip()
        return "N/A"    # Default to "N/A" if no value is found
    except Exception as e:
        print(f"Error extracting value for keyword '{keyword}' in file '{file_path}': {e}")
        return "N/A"





#### Cleanup and standarization ####

# Standardize date
def standardize_date(date_str):
    """Convert date strings to a standard format (YYYY-MM-DD)."""
    try:
        for german, english in GERMAN_MONTHS.items():
            if german in date_str:
                date_str = date_str.replace(german, english)
                break
        return datetime.strptime(date_str, "%d. %B %Y").strftime("%Y-%m-%d")    # As date present in invoice_1 -> 1. März 2024
    except ValueError:
        try:
            return datetime.strptime(date_str, "%b %d, %Y").strftime("%Y-%m-%d")    # As date present in invoice_2 -> Nov  26, 2016
        except ValueError as e:
            print(f"Error standardizing date '{date_str}': {e}")
            return "Unknown Format"

# Standardize currency value
def standardize_currency_value(value_str):
    """Extract currency symbol and standardize formatting using regex."""
    try:
        # Define regex patterns for EUR and USD (case insensitive)
        eur_pattern = re.compile(r"(€|euro|euros|eur|eurs)", re.IGNORECASE)
        usd_pattern = re.compile(r"(\$|usd|dollars|dollar|usds)", re.IGNORECASE)

        # Detect currency type (only detect once)
        if eur_pattern.search(value_str) and usd_pattern.search(value_str):
            return "Error: Mixed currencies in value!"
        
        if eur_pattern.search(value_str):
            currency = "€"
            decimal_separator = ","
            thousand_separator = "."
        elif usd_pattern.search(value_str):
            currency = "$"
            decimal_separator = "."
            thousand_separator = ","
        else:
            return "Unknown Format"  # If no valid currency is found

        # Remove all known currency symbols, names, and extra spaces
        cleaned_value = eur_pattern.sub("", value_str)
        cleaned_value = usd_pattern.sub("", cleaned_value)
        cleaned_value = cleaned_value.strip()

        # Remove thousands separator ('.' for EUR, ',' for USD)
        cleaned_value = cleaned_value.replace(thousand_separator, "")

        # Convert decimal separator to '.' for float conversion
        cleaned_value = cleaned_value.replace(decimal_separator, ".")

        # Convert to float
        numeric_value = float(cleaned_value)

        # Convert back to correct string format
        if currency == "€":
            formatted_value = f"{numeric_value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            formatted_value = f"{numeric_value:,.2f}"

        return f"{formatted_value} {currency}"
    except ValueError as e:
        return f"Error processing value '{value_str}': {e}"

# Extract numeric value
def extract_numeric_value(value_str):
    """Extract numeric value from currency string and convert USD to EUR if applicable."""
    try:
        cleaned_value = standardize_currency_value(value_str).split()[0]  # Get only the numeric part
        numeric_value = float(cleaned_value.replace(",", "."))  # Convert to float

        # Convert USD to EUR
        if "$" in value_str or "USD" in value_str:
            return round(numeric_value * USD_TO_EUR, 2)  # Convert to EUR
        return numeric_value  # Already in EUR
    except Exception as e:
        print(f"Error extracting numeric value from '{value_str}': {e}")
        return None



#### Get the path of the configuration file ####
def get_config_path():
    """Ensure the script finds pdf_config.json no matter where the .exe is run."""
    if getattr(sys, 'frozen', False):
        # If running from PyInstaller .exe, use the extracted temp directory
        base_path = sys._MEIPASS
    else:
        # If running from source, use the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(base_path, "pdf_config.json")
    
    # Debugging: Print the detected path
    print(f"Looking for pdf_config.json at: {config_path}")
    
    if not os.path.exists(config_path):
        print("⚠️ Configuration file not found!")
    
    return config_path


def get_pdf_directory():
    """Ensure the script looks for PDFs in the same directory as the script or executable."""
    if getattr(sys, 'frozen', False):
        # Running as a PyInstaller .exe
        return os.path.dirname(sys.executable)
    else:
        # Running as a script
        return os.path.dirname(os.path.abspath(__file__))


def get_output_path(file_name):
    """Ensure output files are saved in the same directory as the script or .exe."""
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), file_name)
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)




#### Process the PDFs ####
def process_pdfs(pdf_dir, config_file):
    """Process PDFs using the correct directory and configuration file."""
    pdf_dir = get_pdf_directory()
    
    # Debugging: Print the detected PDF directory
    print(f"🔍 Looking for PDFs in: {pdf_dir}")

    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]
    
    # Debugging: Print detected PDF files
    print(f"📂 PDFs found: {pdf_files}")
    
    if not pdf_files:
        print("❌ No PDF files found! Exiting.")
        return
    
    
    config_file = get_config_path()
    
    if not os.path.exists(config_file):
        print(f"Error: Configuration file not found at {config_file}")
        return
    
    
    try:
        with open(config_file, "r") as f:
            config = json.load(f)
    except Exception as e:
        print(f"Error loading configuration file '{config_file}': {e}")
        return
    

    all_extracted_data = []

    pdf_files = sorted(
        [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")],
        # key=lambda x: 0 if x == "sample_invoice_1.pdf" else 1
    )
    
    # Debugging: Print detected PDF files
    print(f"🔍 PDFs found in {pdf_dir}: {pdf_files}")
    
    if not pdf_files:
        print("❌ No PDF files found! Exiting.")
        return

    for file_name in pdf_files:
        file_path = os.path.join(pdf_dir, file_name)
        print(f"Processing: {file_name}")

        if file_name in config:
            pdf_config = config[file_name]
            record = {"File Name": file_name}
            for field, details in pdf_config.items():
                # value = extract_value_by_keyword(file_path, details["keyword"], details.get("offset", 0))
                value = extract_value_by_keyword(file_path, details["keyword"],
                                                details.get("offset", 0),
                                                details.get("page")  # Read page number from JSON
                                                )   

                if not value:
                    print(f"Warning: '{field}' not found in '{file_name}'. Defaulting to 'N/A'.")
                    value = "N/A"
                
                # # Store the extracted value, further processing will be done later
                # # Depending on the required field
                # record[field] = value

                if field == "Date":
                    value = standardize_date(value)
                    record["Date"] = value  # Store the date

                elif field == "Value":
                    original_value = standardize_currency_value(value)
                    value_numeric = extract_numeric_value(value)  # Convert to EUR if needed
                    record["Original Value"] = original_value
                    record["Value in EUR"] = value_numeric  # Store converted value separately

            all_extracted_data.append(record)
        else:
            print(f"Warning: No configuration found for '{file_name}'.")

    if not all_extracted_data:
        print("No data extracted. Exiting.")
        return

    df = pd.DataFrame(all_extracted_data)

    
    df = df[["File Name", "Date", "Original Value", "Value in EUR"]]  # Ensure correct order of columns

    # Create Excel file
    try:
        excel_output_path = get_output_path("output.xlsx")
        with pd.ExcelWriter(excel_output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")  # First sheet with original values

            # **Create Pivot Table**
            pivot_table = pd.pivot_table(
                df,
                values="Value in EUR",
                index="Date",   # Date as index
                columns="File Name",
                aggfunc="sum",
                fill_value=0
            )

            # "Total" Column for Summation
            pivot_table["Total"] = pivot_table.sum(axis=1)
            
            # Write the Pivot Table to Excel
            pivot_table.to_excel(writer, sheet_name="Sheet2", startrow=2)
            worksheet = writer.sheets["Sheet2"]
            worksheet["A1"] = "Pivot Table - Sum of Values in EUR"
            worksheet.auto_filter.ref = worksheet.dimensions  # Apply filtering to the table

        print("Excel file 'output.xlsx' created successfully.")
        print(f"✅ Excel file saved: {excel_output_path}")
    except Exception as e:
        print(f"Error creating Excel file: {e}")

    # Save CSV File
    try:
        csv_output_path = get_output_path("output.csv")
        df.to_csv(csv_output_path, index=False, sep=";")
        print(f"✅ CSV file saved: {csv_output_path}")
        print("CSV file 'output.csv' created successfully.")
    except Exception as e:
        print(f"Error creating CSV file: {e}")


#### Main function to process PDFs ####
if __name__ == "__main__":
    pdf_dir = "."       # Directory containing PDFs
    config_file = "pdf_config.json"         # JSON configuration file
    process_pdfs(pdf_dir, config_file)




