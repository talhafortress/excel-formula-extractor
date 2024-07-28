# excel-formula-extractor
Excel Formula Extractor
This Python script utilizes the openpyxl library to extract and preserve formulas from an Excel workbook while removing the cell values. It creates a new workbook where each sheet from the original workbook is replicated, containing only the formulas and leaving all other cell values empty.

Features
Sheet Replication: Copies all sheets from the original workbook to the new one.
Formula Preservation: Keeps all cell formulas intact.
Value Removal: Clears the content of all cells, leaving only the formulas.
Requirements
Python 3.x
openpyxl library
Installation
Install the openpyxl library if you haven't already:

>pip install openpyxl

Usage
Place your original Excel file in the same directory as the script and name it original_file.xlsx.
Run the script

Contributing
Feel free to fork this repository and submit pull requests if you have any improvements or bug fixes. Contributions are welcome!

License
This project is licensed under the MIT License - see the LICENSE file for details.
