**Excel Transaction Automation with Python**
This Python script automates the process of loading an Excel file, manipulatin_g the data in it, and creating a bar chart using the openpyxl library.

**Table of Contents**
*Overview
*Features
*Requirements
*Setup
*Usage
*Code Explanation
*License
*Overview
This script reads transaction data from an Excel sheet (transactions.xlsx), applies a 10% discount to the prices in column C, and writes the updated prices to column D. Additionally, a bar chart is generated based on the corrected prices and added to the sheet.

**Features**
Reads data from an Excel file.
Processes prices by applying a discount.
Adds new data (corrected prices) to the Excel sheet.
Generates a bar chart based on the corrected prices.
Saves the updated Excel sheet with a new bar chart.
Requirements
Python 3.x
openpyxl Python package
Setup
Clone this repository:


git clone https://github.com/yourusername/excel-automation.git
Navigate to the project directory:


cd excel-automation
Install the required Python package:


pip install openpyxl
Place your transactions.xlsx file in the specified location (or update the path in the script as needed):


C:\\Users\\srira\\Downloads\\PYTHON PROJECTS\\Automation with python\\transactions.xlsx\\transactions.xlsx
Usage
Ensure your transactions.xlsx file is properly formatted. It should contain the transaction data, with at least the following structure:

Column C: Price of the transaction.
Run the Python script:


python transaction_automation.py
After running the script, an updated Excel file with corrected prices and a bar chart will be saved as transactions.xlsx.

**Code Explanation**
Main Functionality:
Loading Workbook: The script uses openpyxl to load the Excel file (transactions.xlsx) and accesses the active sheet.


wb = xl.load_workbook("path_to_your_file")
sheet = wb['Sheet1']
Processing Each Row: The script iterates through each row, starting from row 2 (excluding headers), and applies a 10% discount to the values in column C. The corrected price is stored in column D.


for row in range(2, sheet.max_row+1):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value * 0.9
    sheet.cell(row, 4).value = corrected_price
Adding Bar Chart: After processing the data, a bar chart is created using the corrected prices in column D.


values = Reference(sheet, min_row=2, max_row=sheet.max_row, min_col=4, max_col=4)
chart = BarChart()
chart.add_data(values)
sheet.add_chart(chart, 'E2')
Saving the Workbook: The updated Excel file is saved with the chart.


wb.save("transactions.xlsx")
License
This project is licensed under the MIT License - see the LICENSE file for details.

