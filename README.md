# Electrochemistry Data Processing Tool v1.0.0

## Update Date: May 21, 2025

The Electrochemistry Data Processing Tool is a Python script designed to automate the processing of various electrochemical test data. It simplifies the transformation from raw data files to structured Excel reports and extracts key electrochemical parameters.

Supported data types include:

- **Cyclic Voltammetry (CV)**: Automatically identifies CV curves and calculates parameters such as double-layer capacitance (Cdl).
- **Linear Sweep Voltammetry (LSV)**: Processes LSV data and extracts relevant parameters.
- **Electrochemical Impedance Spectroscopy (EIS)**: Analyzes EIS data, calculates solution resistance (Rs), and generates ZView-compatible plain data files.

## Key Features

- **Automatic File Recognition**: Automatically identifies CV, LSV, and EIS data files based on file content characteristics (e.g., specific keywords).
- **Data Extraction and Processing**: Precisely extracts relevant data columns from text files.
  - CV: Current and voltage data, current density difference (Δj) for Cdl calculation.
  - LSV: Current and voltage data.
  - EIS: Frequency (Freq), real impedance (Z'), imaginary impedance (Z''). Outputs -Z''.
- **Parameter Calculation**:
  - CV: Double-layer capacitance (Cdl).
  - EIS: Solution resistance (Rs), calculated by finding the intersection or closest point of -Z'' and the Z' axis.
- **Tafel Plot Data Generation**: Combines LSV data and Rs values derived from EIS to calculate `log(j)` and `Overpotential`, outputting the results to a dedicated "Tafel Data" worksheet.
- **Formatted Excel Reports**:
  - Consolidates all processed data and calculated results into a single Excel workbook, including detailed data worksheets (`CV Data`, `LSV Data`, `EIS Data`, `Tafel Data`) and a summary `Analysis Report` worksheet.
  - The `Analysis Report` worksheet summarizes key parameters from the analysis and is set as the default active worksheet when the Excel file is opened.
  - Standardized header format: 3 header rows + 1 blank row with the same format as the header, with data starting from row 5 (for data sheets).
  - In the EIS data sheet, Z' and -Z'' data at the Rs intersection are highlighted with a yellow background and bold font.
  - Supports side-by-side display of multiple datasets of the same type in the same worksheet (applicable to `CV Data`, `LSV Data`, `EIS Data`, `Tafel Data`).
- **ZView-Compatible File Generation**: For EIS data, generates plain data files in `.txt` format (named `OriginalFileName-ZView.txt`) for easy import into ZView and other professional fitting software. These files are saved in the user's selected raw data folder.
- **Logging**: Logs detailed program execution, warnings, and errors to log files in the `logs` folder for tracking and debugging.
- **Executable Packaging**: Supports packaging the Python application into a single executable file (`ElectrochemistryTool.exe`) for use on machines without a Python environment.
- **User-Friendly Interaction**:
  - Select data folders via a graphical interface.
  - Clear terminal progress prompts.

## Installation and Usage

### System Requirements

- Python 3.x
- `pip` (Python package manager)

### Install Dependencies

Ensure all necessary libraries are installed in your Python environment. Open a terminal in the project root directory (`cursor` directory) and run:

```bash
pip install -r requirements.txt
```

The `requirements.txt` file should include at least the following:

```text
openpyxl
numpy
tqdm # Optional, for progress bars
```

### Run the Program

Open a terminal in the project root directory (`cursor` directory) and run:

```bash
python run_electrochemistry.py
```

Alternatively, if an executable version (`ElectrochemistryTool.exe`) is provided, you can run the program directly.

## Usage Instructions

1. **Start the Program**: Run `python run_electrochemistry.py` or double-click `ElectrochemistryTool.exe`.
2. **Select Folder**: The program will prompt you to select a folder. Choose the folder containing raw electrochemical data files (`.txt` format).
3. **Automatic Processing**: The program will automatically scan the selected folder for files:
    - Identify CV, LSV, and EIS files.
    - Extract and calculate data from identified files.
    - Generate ZView-compatible files (for EIS data) in the selected folder.
    - Generate an Excel file containing all processed results in the `processed_data` subdirectory of the selected folder. The file name format is `FolderName_processed_data_Timestamp.xlsx`.
4. **View Results**:
    - Open the generated Excel file to view detailed data and calculated parameters. The default worksheet is "Analysis Report".
    - Check the `-ZView.txt` files generated in the raw data folder.
    - If any issues arise, check the terminal output or log files in the `logs` folder.

## File Structure

```text
cursor/
├── electrochemistry/         # Core processing module package
│   ├── common/               # Common utility modules (Excel, file operations)
│   │   ├── __init__.py
│   │   ├── excel_utils.py
│   │   └── file_utils.py
│   ├── __init__.py
│   ├── cv.py                 # CV data processing module
│   ├── eis.py                # EIS data processing module
│   ├── lsv.py                # LSV data processing module
│   ├── tafel.py              # Tafel analysis module
│   └── main.py               # Main control logic
├── logs/                     # Log file storage directory
├── README.md                 # This documentation file
├── requirements.txt          # Python dependency list
├── run_electrochemistry.py   # Main entry script
└── ElectrochemistryTool.spec # PyInstaller configuration file

# Example of user-selected folder
selected_data_folder/
├── cv_data_1.txt
├── lsv_data_1.txt
├── eis_data_1.txt
├── eis_data_1-ZView.txt      # <--- ZView files will be generated here
└── processed_data/           # <--- Excel reports will be generated here
    └── selected_data_folder_processed_data_YYYYMMDD_HHMMSS.xlsx
```

## Notes

- **Input File Format**: Currently supports raw data files in `.txt` format. Ensure your data file structure is compatible with the parsing logic of each module (CV, LSV, EIS).
  - CV/LSV: Typically requires clear voltage and current data columns.
  - EIS: Requires the keyword "A.C. Impedance" for identification, and the data section should have headers like "Freq/Hz, Z'/ohm, Z"/ohm,...".
- **Dependency Installation**: Ensure all dependencies are installed via `pip install -r requirements.txt` before running (if not using the `.exe` version).
- **Excel File Write Permissions**: Ensure the program has write permissions for the target output directory (i.e., the `processed_data` subdirectory under the selected folder). If the Excel file is open, saving may fail.
- **Antivirus False Positives**: If using the `.exe` version, some antivirus software may flag it (e.g., if UPX compression is used). If this occurs, try using an uncompressed `.exe` version or run from source.
- **Error Troubleshooting**: If issues arise, first check the terminal error messages. For more detailed information, check the log files in the `logs` folder.

## Future Improvements

- Support for more data formats and instrument models.
- Add data visualization features (e.g., charts generated directly in Excel or as standalone image files).
- Provide more configurable analysis parameters.
- Develop a more comprehensive graphical user interface (GUI).

## Contributions

Suggestions for improvements or code contributions are welcome.

---

_This tool aims to improve the efficiency of electrochemical data processing. If you have specific needs or find bugs, please provide feedback._
