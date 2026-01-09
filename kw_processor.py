"""
DDR Keyword Risk Processor

Scans Daily Drilling Reports (DDR) Excel files for risk-related keywords
in the "Details of Operations in Sequence and Remarks" column and outputs
matches to a new Excel sheet.
"""

from pathlib import Path
import re
import pandas as pd
from openpyxl import load_workbook


# Risk keywords to search for
RISK_KEYWORDS = [
    "High TRQ",
    "Torque",
    "Stuck",
    "String vibration",
    "Mud loss",
    "Mud gain",
    "Drag",
    "String installation",
    "String running",
    "Washout",
    "Hole problems",
    "Differential sticking",
    "Pack-off",
    "Kick",
    "Well control",
]


def extract_report_number(filepath: Path) -> str:
    """Extract report number from filename or return filename as identifier."""
    filename = filepath.stem
    # Try to extract number after "DDR" or "#"
    match = re.search(r'(?:DDR\s*#?\s*|#\s*)(\d+)', filename, re.IGNORECASE)
    if match:
        return match.group(1)
    return filename


def find_operations_column(df: pd.DataFrame) -> int:
    """Find the column index containing 'DETAILS OF OPERATIONS IN SEQUENCE AND REMARKS'."""
    for row_idx in range(min(100, len(df))):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[row_idx, col_idx]
            if isinstance(cell_value, str) and 'DETAILS OF OPERATIONS' in cell_value.upper():
                return col_idx
    return -1


def find_operations_data(df: pd.DataFrame, ops_col: int) -> list:
    """
    Extract operations data rows with their time information.
    Returns list of dicts with 'time_from', 'time_to', 'details'.
    """
    results = []

    # Find the header row for operations
    header_row = -1
    for row_idx in range(min(100, len(df))):
        cell_value = df.iloc[row_idx, ops_col]
        if isinstance(cell_value, str) and 'DETAILS OF OPERATIONS' in cell_value.upper():
            header_row = row_idx
            break

    if header_row == -1:
        return results

    # Data starts after the header rows (typically 2 rows: header + FROM/TO/DURATION row)
    data_start = header_row + 2

    for row_idx in range(data_start, len(df)):
        # Get the details text from operations column
        details = df.iloc[row_idx, ops_col]

        # Skip empty rows or rows that are section headers
        if pd.isna(details) or not isinstance(details, str) or details.strip() == '':
            # Check if this is a section break (like TOTAL row or new section)
            col0_val = df.iloc[row_idx, 0]
            if isinstance(col0_val, str) and ('TOTAL' in col0_val.upper() or 'SUMMARY' in col0_val.upper()):
                continue
            continue

        # Get time information (columns 0 and 1 typically have FROM and TO times)
        time_from = df.iloc[row_idx, 0]
        time_to = df.iloc[row_idx, 1]

        # Format time/date
        time_str = ""
        if pd.notna(time_from):
            if isinstance(time_from, pd.Timestamp):
                time_str = time_from.strftime('%Y-%m-%d %H:%M')
            else:
                time_str = str(time_from)
        if pd.notna(time_to):
            if isinstance(time_to, pd.Timestamp):
                time_str += f" to {time_to.strftime('%H:%M')}"
            elif time_str:
                time_str += f" to {time_to}"

        results.append({
            'time_date': time_str,
            'details': details
        })

    return results


def search_keywords(text: str, keywords: list) -> list:
    """Search for keywords in text (case-insensitive). Returns list of matched keywords."""
    matched = []
    text_lower = text.lower()
    for kw in keywords:
        if kw.lower() in text_lower:
            matched.append(kw)
    return matched


def process_ddr_file(filepath: Path, keywords: list) -> list:
    """
    Process a single DDR Excel file and return list of risk matches.
    Each match contains: report_number, time_date, risks
    """
    results = []
    report_number = extract_report_number(filepath)

    # Read Excel file
    xl = pd.ExcelFile(filepath)

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)

        # Find the operations column
        ops_col = find_operations_column(df)
        if ops_col == -1:
            continue

        # Extract operations data
        operations = find_operations_data(df, ops_col)

        for op in operations:
            matched_keywords = search_keywords(op['details'], keywords)
            if matched_keywords:
                results.append({
                    'Report Number': report_number,
                    'Time/Date': op['time_date'],
                    'Risks': ', '.join(matched_keywords)
                })

    return results


def process_all_files(input_folder: Path, keywords: list) -> pd.DataFrame:
    """Process all Excel files in the input folder."""
    all_results = []

    # Find all Excel files
    excel_files = list(input_folder.glob('*.xlsx')) + list(input_folder.glob('*.xls'))

    for filepath in excel_files:
        print(f"Processing: {filepath.name}")
        try:
            results = process_ddr_file(filepath, keywords)
            all_results.extend(results)
            print(f"  Found {len(results)} risk entries")
        except Exception as e:
            print(f"  Error processing file: {e}")

    return pd.DataFrame(all_results)


def save_results(df: pd.DataFrame, output_path: Path, sheet_name: str = 'Risk Analysis'):
    """Save results to a new Excel file or append to existing."""
    if df.empty:
        print("No results to save.")
        return

    if output_path.exists():
        # Append to existing file
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Create new file
        df.to_excel(output_path, sheet_name=sheet_name, index=False)

    print(f"Results saved to: {output_path}")


def main():
    # Paths
    script_dir = Path(__file__).parent
    input_folder = script_dir / 'input-sheets'
    output_file = script_dir / 'output' / 'risk_analysis.xlsx'

    # Create output directory if it doesn't exist
    output_file.parent.mkdir(parents=True, exist_ok=True)

    print("DDR Risk Keyword Processor")
    print("=" * 40)
    print(f"Input folder: {input_folder}")
    print(f"Output file: {output_file}")
    print(f"Keywords: {len(RISK_KEYWORDS)} defined")
    print("=" * 40)

    if not input_folder.exists():
        print(f"Error: Input folder not found: {input_folder}")
        return

    # Process all files
    results_df = process_all_files(input_folder, RISK_KEYWORDS)

    print("=" * 40)
    print(f"Total risk entries found: {len(results_df)}")

    # Save results
    if not results_df.empty:
        save_results(results_df, output_file)
        print("\nResults preview:")
        print(results_df.to_string())
    else:
        print("No risk keywords found in any files.")


if __name__ == "__main__":
    main()
