# Rightmove Property Data Scraper

## Overview
This script automates the extraction of property details from Rightmove using Selenium and Python. It reads property URLs from an Excel sheet, scrapes the required data, and updates the Excel sheet with the extracted information.

---

## Features
1. **Automated Data Extraction**:
   - Property address and postcode.
   - Asking price, property type, agent details, and size.
   - Bedrooms, bathrooms, lease status, garden, parking, and broadband speed.
   - Energy Performance Certificate (EPC) rating and council tax band.
   - Tenanted status based on the description.

2. **Excel Integration**:
   - Reads property URLs from an Excel file (`Rightmove scrape template.xlsx`).
   - Updates extracted data directly into the Excel sheet.

3. **Error Handling**:
   - Skips invalid URLs or missing data gracefully.
   - Logs errors for debugging purposes.

4. **Dynamic Element Interaction**:
   - Handles dynamic page elements like modals, buttons, and dropdowns.

---

## Prerequisites
1. **Python Libraries**:
   - `selenium`: For browser automation.
   - `openpyxl`: For Excel manipulation.

   Install them using:
   ```bash
   pip install selenium openpyxl
   ```

2. **Browser Driver**:
   - Google Chrome with the corresponding ChromeDriver installed.
   - Ensure ChromeDriver is added to your system's PATH.

3. **Excel File**:
   - A file named `Rightmove scrape template.xlsx` containing property URLs in the first column (starting from row 4).

---

## Setup and Usage
1. **Prepare Your Environment**:
   - Place the Excel file in the same directory as the script.
   - Ensure URLs are correctly formatted in the first column of the Excel file.

2. **Run the Script**:
   ```bash
   python rightmove_scraper.py
   ```

3. **Output**:
   - The Excel file will be updated with extracted data, including:
     - Property address, postcode, and price.
     - Agent information, size, bedrooms, bathrooms, and more.

---

## Customization
- **Data Fields**: Modify the script to extract additional fields if needed.
- **Error Handling**: Enhance error logging for detailed debugging.
- **Browser Options**: Adjust Chrome options for specific requirements.

---

## Notes
- The script is designed to work with the current structure of Rightmove's website. Changes to the website layout may require updates to the XPath selectors.
- Ensure stable internet connectivity for smooth execution.
