import os
import pandas as pd
import chardet
import win32com.client as win32
import sys
from tqdm import tqdm
import re
import html
from io import StringIO


def create_output_folder(source_folder):
    """Create 'converted' subfolder if it doesn't exist"""
    output_folder = os.path.join(source_folder, "converted")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    return output_folder


def detect_encoding(file_path):
    """Detect file encoding with chardet"""
    try:
        with open(file_path, 'rb') as f:
            rawdata = f.read(50000)
        result = chardet.detect(rawdata)
        return result['encoding'] if result['confidence'] > 0.7 else 'utf-8'
    except Exception:
        return 'utf-8'


def is_html_content(file_path, encoding):
    """Check if content contains HTML/XML tags"""
    try:
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            content = f.read(1000)
        return any(tag in content.lower() for tag in ('<table', '<tr', '<td', '<html', '<?xml'))
    except:
        return False


def simple_html_table_to_df(html_content):
    """Convert HTML tables to DataFrame without using pandas' HTML parser"""
    # Basic HTML unescape
    html_content = html.unescape(html_content)

    # Find all tables in the HTML
    table_matches = re.findall(r'<table.*?>(.*?)</table>', html_content, re.DOTALL | re.IGNORECASE)
    all_tables = []

    for table_html in table_matches:
        rows = []
        # Find all rows in the table
        row_matches = re.findall(r'<tr.*?>(.*?)</tr>', table_html, re.DOTALL | re.IGNORECASE)

        for row_html in row_matches:
            # Find all cells in the row
            cell_matches = re.findall(r'<t[dh].*?>(.*?)</t[dh]>', row_html, re.DOTALL | re.IGNORECASE)
            row_data = []

            for cell in cell_matches:
                # Clean cell content - remove HTML tags
                cell_clean = re.sub(r'<.*?>', '', cell)
                # Replace HTML entities and whitespace
                cell_clean = re.sub(r'\s+', ' ', cell_clean).strip()
                row_data.append(cell_clean)

            if row_data:
                rows.append(row_data)

        if rows:
            # Create DataFrame from rows
            try:
                # Find max columns for consistent shape
                max_cols = max(len(row) for row in rows)
                padded_rows = [row + [''] * (max_cols - len(row)) for row in rows]
                all_tables.append(pd.DataFrame(padded_rows))
            except:
                continue

    if all_tables:
        return pd.concat(all_tables, ignore_index=True)
    return None


def convert_file(file_path, output_folder, output_format):
    """Convert file to specified Excel format"""
    filename = os.path.basename(file_path)
    base_name, _ = os.path.splitext(filename)
    output_path = os.path.join(output_folder, f"{base_name}.{output_format}")

    try:
        # Detect encoding
        encoding = detect_encoding(file_path)

        # Determine content type
        is_html = is_html_content(file_path, encoding)

        # Process based on content type
        if is_html:
            # Read entire HTML content
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                html_content = f.read()

            # Use custom HTML table parser
            df = simple_html_table_to_df(html_content)

            if df is None or df.empty:
                return f"No valid tables found in HTML: {filename}"
        else:
            # Text/CSV Processing
            try:
                # Read entire file as text
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    content = f.read()

                # Try to detect delimiter
                lines = content.split('\n')[:100]
                delimiter_counts = {',': 0, ';': 0, '\t': 0, '|': 0}
                for line in lines:
                    for delim in delimiter_counts:
                        delimiter_counts[delim] += line.count(delim)

                # Use most common delimiter
                delimiter = max(delimiter_counts, key=delimiter_counts.get)

                # Read as CSV with pandas
                df = pd.read_csv(StringIO(content), delimiter=delimiter,
                                 header=None, engine='python', on_bad_lines='skip')

                if df.empty:
                    return f"No data found in {filename}"

            except Exception as e:
                return f"Text parsing failed for {filename}: {str(e)}"

        # Save to Excel format without index or headers
        if output_format == 'xlsx':
            df.to_excel(output_path, index=False, header=False, engine='openpyxl')
        else:  # XLSB
            temp_xlsx = output_path.replace('.xlsb', '_temp.xlsx')
            df.to_excel(temp_xlsx, index=False, header=False, engine='openpyxl')
            convert_xlsx_to_xlsb(temp_xlsx, output_path)
            os.remove(temp_xlsx)

        return f"Successfully converted {df.shape[0]} rows to {output_format.upper()}"

    except Exception as e:
        return f"Unexpected error processing {filename}: {str(e)}"


def convert_xlsx_to_xlsb(xlsx_path, xlsb_path):
    """Convert XLSX to XLSB using Excel COM interface (Windows only)"""
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        wb.SaveAs(os.path.abspath(xlsb_path), FileFormat=50)  # 50 = xlExcel12 (XLSB)
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        raise RuntimeError(f"XLSB conversion failed: {str(e)}")


def main():
    print("File Conversion Tool")
    print("--------------------\n")

    # Get source folder
    source_folder = input("Enter path to source folder: ").strip()
    if not os.path.isdir(source_folder):
        print("\nError: Invalid folder path")
        return

    # Create output folder
    output_folder = create_output_folder(source_folder)

    # Get output format
    output_format = input("Convert to format (xlsx/xlsb): ").lower().strip()
    if output_format not in ('xlsx', 'xlsb'):
        print("\nError: Invalid format. Please choose 'xlsx' or 'xlsb'")
        return

    # Check for Windows if XLSB selected
    if output_format == 'xlsb' and not sys.platform.startswith('win'):
        print("\nError: XLSB conversion requires Windows")
        return

    # Process files
    print("\nStarting conversion...")
    processed_count = 0
    error_count = 0

    files = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]

    # Initialize tqdm progress bar
    pbar = tqdm(files, desc="Processing files")
    for filename in pbar:
        file_path = os.path.join(source_folder, filename)
        pbar.set_postfix(file=filename[:20])

        result = convert_file(file_path, output_folder, output_format)
        pbar.write(f"{filename}: {result}")

        if "Successfully" in result:
            processed_count += 1
        else:
            error_count += 1

    print(f"\nConversion complete! {processed_count} files converted, {error_count} errors")
    print(f"Converted files saved to: {output_folder}")


if __name__ == "__main__":
    main()