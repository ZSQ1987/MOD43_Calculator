# MOD43 Calculator

A GUI application for calculating MOD43 check digits with a user-friendly interface.

## Features

- **MOD43 Check Digit Calculation**: Calculate check digits for 14-character strings
- **Detailed Calculation Display**: Shows step-by-step calculation process
- **Character Mapping Table**: Displays MOD32 character set with 8 columns and 6 rows
- **Excel Export**: Export results to Excel with two sheets (character set and calculation results)
- **Result Highlighting**: Highlights the calculated result in the character mapping table
- **Input Validation**: Enforces 14-character input validation
- **Mouse Hover Highlighting**: Highlights character mapping table entries on mouse hover

## Requirements

- Python 3.x
- tkinter (standard library)
- openpyxl (for Excel export)

## Installation

1. Clone the repository
2. Install required dependencies:
   ```bash
   pip install openpyxl
   ```

## Usage

### Running from source
1. Run the application:
   ```bash
   python MOD43_GUI.py
   ```
2. Enter a 14-character string in the input field
3. Click "计算" to calculate the check digit
4. View the calculation process and result
5. Click "导出 Excel" to export the results to an Excel file

### Building an executable
1. Run the packing script:
   ```bash
   .\packing.ps1
   ```
2. The executable will be created in the `dist` directory
3. Run `dist\MOD43_GUI.exe` to launch the application

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.