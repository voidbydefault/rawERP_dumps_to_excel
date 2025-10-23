# Universal Raw-Dump File to Excel Converter

This Python script is designed for batch-converting various file types—including delimited text, CSV, and simple HTML files into Microsoft Excel format (`.xlsx` or `.xlsb`).

## Features

*   **Batch Conversion:** Processes all files found within a specified source folder.
*   **Format Support:** Converts to standard `.xlsx` or the optimized binary `.xlsb` format.
*   **Intelligent Parsing:**
    *   Automatic encoding detection using `chardet`.
    *   Attempts to detect common delimiters (comma, semicolon, tab, pipe) for text/CSV files.
    *   Custom parser to extract and merge simple HTML tables.
*   **Progress Display:** Uses `tqdm` to show progress during conversion.
*   **Clean Output:** Converted Excel files are saved without index or header rows.

## Prerequisites

*   Python 3.x
*   **For `.xlsb` Output:** A Windows operating system with Microsoft Excel installed is required, as this format conversion relies on the Excel COM interface (`win32com.client`).

## Installation (Using PyCharm IDE)

1. Download code  
   ```bash
   git clone https://github.com/voidbydefault/rawERP_dumps_to_excel
   ```

3. **Open the Project in PyCharm:** Open the repository folder as a project in PyCharm.

4.  **Install Core Dependencies:**
    Use PyCharm's built-in features to install `requirements.txt` or use **Terminal** (located at the bottom of the IDE) to install all core dependencies listed in `requirements.txt`:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Click on play icon in PyCharm IDE to run the script or use terminal:

    ```bash
    python converter.py
    ```

2.  Follow the interactive prompts:
    *   **Enter path to source folder:** Provide the directory containing the files you wish to convert.
    *   **Convert to format (xlsx/xlsb):** Specify the desired output format.

### Output Location

The converted Excel files will be saved in a new subfolder named `converted`, which is created inside the original source folder.

## Limitations

*   **XLSB Conversion:** As noted, `.xlsb` output is strictly limited to **Windows systems** with MS Excel installed due to the requirement for the `win32com.client` dependency.
*   **HTML Parsing:** The script uses a basic, custom regular expression-based parser for HTML tables. It is robust for simple structures but may fail to accurately extract data from highly complex, nested, or poorly formed HTML documents.