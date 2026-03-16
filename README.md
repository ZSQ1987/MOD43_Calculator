# MOD43 Calculator

A professional GUI application for calculating MOD43 check digits with a user-friendly interface, designed for data integrity verification across various industries.

## Features

- **MOD43 Check Digit Calculation**: Calculate check digits for custom-length strings
- **Detailed Calculation Display**: Shows step-by-step calculation process with formulas
- **Character Mapping Table**: Displays complete MOD43 character set with 5 columns and 9 rows
- **Excel Export**: Export results to Excel with two sheets (character set and calculation results)
- **Result Highlighting**: Highlights the calculated result in the character mapping table
- **Input Validation**: Enforces configured string length validation with centered error messages
- **Configuration System**: Customizable string length with persistent settings
- **Icon Support**: Application icon for professional appearance
- **Cross-Platform**: Works on Windows systems

## Requirements

- Python 3.6+
- tkinter (standard library)
- openpyxl (for Excel export)

## Installation

### From Source
1. Clone the repository
2. Install required dependencies:
   ```bash
   pip install openpyxl
   ```

### From Executable
1. Download the pre-built executable from the [releases page](https://github.com/ZSQ1987/MOD43_Calculator/releases)
2. Run `MOD43_GUI.exe` directly (no installation required)

## Usage

### Running from Source
1. Run the application:
   ```bash
   python MOD43_GUI.py
   ```
2. Enter a string in the input field
3. Click "计算" to calculate the check digit
4. View the calculation process and result
5. Click "配置" to adjust string length settings
6. Click "导出Excel" to export the results to an Excel file

### Building an Executable
1. Run the packing script:
   ```bash
   .\packing.ps1
   ```
2. The executable will be created in the `dist` directory
3. Run `dist\MOD43_GUI.exe` to launch the application

## Configuration

- **String Length**: Customize the required string length via the "配置" button
- **Settings Persistence**: Configuration is automatically saved to `config.ini`
- **Error Handling**: Centered error messages for invalid inputs

## Technical Details

- **GUI Framework**: Tkinter/TTK
- **Excel Export**: openpyxl library
- **Configuration**: configparser
- **Packaging**: PyInstaller
- **Icon Support**: app_ico.ico

## Application Structure

```
MOD43_Calculator/
├── MOD43_GUI.py         # Main application
├── packing.ps1          # Build script
├── app_ico.ico          # Application icon
├── version_info.txt     # Version information
├── README.md            # This file
├── LICENSE              # MIT License
└── .gitignore          # Git ignore file
```

## Use Cases

- **Logistics**: Validate shipping and tracking numbers
- **Manufacturing**: Verify product serial numbers
- **Finance**: Ensure account number integrity
- **Software Development**: Integrate check digit validation
- **Education**: Learn MOD43 algorithm principles

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Project Link

[https://github.com/ZSQ1987/MOD43_Calculator](https://github.com/ZSQ1987/MOD43_Calculator)
