import pandas as pd
import re
from pathlib import Path

# File paths - can include prior year data from next year's file
FILES = {
    2020: "downloads/FinancialStatement-2021-Tahunan-MEDC.xlsx",  # 2020 data in 2021 file
    2021: "downloads/FinancialStatement-2021-Tahunan-MEDC.xlsx",
    2022: "downloads/FinancialStatement-2022-Tahunan-MEDC.xlsx",
    2023: "downloads/FinancialStatement-2023-Tahunan-MEDC.xlsx",
    2024: "downloads/FinancialStatement-2024-Tahunan-MEDC.xlsx",
}

# Sheet codes to look for
SHEETS = {
    'balance_sheet': ['1210000', '121'],
    'income_statement': ['1311000', '131'],
    'cashflow': ['1510000', '151']
}

def find_sheet(excel_file, codes):
    """Find sheet that starts with any of the given codes."""
    for sheet_name in excel_file.sheet_names:
        for code in codes:
            if str(sheet_name).startswith(code):
                return sheet_name
    return None

def parse_number(value):
    """Convert Indonesian number format to float."""
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s or s == 'nan':
        return None
    
    # Handle negative (parentheses)
    is_neg = s.startswith('(') and s.endswith(')')
    if is_neg:
        s = s[1:-1]
    
    # Remove periods (thousands), replace comma with period (decimal)
    s = s.replace('.', '').replace(',', '.')
    
    try:
        num = float(s)
        return -num if is_neg else num
    except:
        return None

def find_year_column(df, year):
    """Find column containing the target year."""
    # Search first 15 rows for year
    for row_idx in range(min(15, len(df))):
        for col_idx in range(len(df.columns)):
            cell = str(df.iat[row_idx, col_idx])
            if str(year) in cell:
                return col_idx
    
    # If year not found, try to find by position
    # Typically: Column 1 (B) = current year, Column 2 (C) = prior year
    # For 2021 file requesting 2020, we want column 2 (C)
    # For other years, default to column 1 (B)
    return 1  # Default to second column

def extract_sheet_data(file_path, sheet_codes, year):
    """Extract account names and values from a sheet."""
    xl = pd.ExcelFile(file_path)
    sheet_name = find_sheet(xl, sheet_codes)
    
    if not sheet_name:
        print(f"  ⚠ Sheet not found for {year}")
        return pd.DataFrame({'Account': [], str(year): []})
    
    print(f"  ✓ Found sheet: {sheet_name}")
    
    # Read sheet without any parsing
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
    
    # Find the value column
    value_col = find_year_column(df, year)
    
    # Account column is typically column 0
    account_col = 0
    
    # Extract as-is
    result = pd.DataFrame({
        'Account': df.iloc[:, account_col].astype(str),
        str(year): df.iloc[:, value_col].apply(parse_number)
    })
    
    # Remove empty rows and obvious headers
    result = result[result['Account'].str.strip() != '']
    result = result[result['Account'] != 'nan']
    
    # Remove common header patterns (case insensitive, must match the whole text)
    account_lower = result['Account'].str.lower().str.strip()
    result = result[~account_lower.isin([
        'laporan posisi keuangan',
        'statement of financial position',
        'laporan laba rugi',
        'income statement',
        'profit or loss',
        'laporan arus kas',
        'cash flow statement',
        'statement of cash flows'
    ])]
    
    print(f"  → Extracted {len(result)} rows")
    return result

def consolidate_statement(statement_type):
    """Consolidate one statement type across all years."""
    sheet_codes = SHEETS[statement_type]
    print(f"\nProcessing {statement_type}...")
    
    all_data = {}
    
    # Extract data for each year
    for year, file_path in sorted(FILES.items()):
        if not Path(file_path).exists():
            print(f"  ⚠ File not found: {year}")
            continue
        
        print(f"  {year}:")
        df = extract_sheet_data(file_path, sheet_codes, year)
        all_data[year] = df
    
    # Use first year as base
    if not all_data:
        return pd.DataFrame()
    
    first_year = min(all_data.keys())
    result = all_data[first_year][['Account']].copy()
    
    # Add each year's data
    for year in sorted(all_data.keys()):
        result = result.merge(
            all_data[year],
            on='Account',
            how='left'
        )
    
    return result

# Main execution
print("="*60)
print("MEDC Financial Statement Consolidation")
print("="*60)

balance_sheet = consolidate_statement('balance_sheet')
income_statement = consolidate_statement('income_statement')
cashflow = consolidate_statement('cashflow')

# Create output file with separate sheets
output_file = "MEDC_FS_Consolidated.xlsx"
print(f"\nCreating {output_file}...")

with pd.ExcelWriter(output_file, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
    workbook = writer.book
    
    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'bg_color': '#366092',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    number_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })
    
    account_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter'
    })
    
    # Define statements to write
    statements = [
        (balance_sheet, 'Balance Sheet'),
        (income_statement, 'Income Statement'),
        (cashflow, 'Cash Flow')
    ]
    
    # Write each statement to its own sheet
    for df, sheet_name in statements:
        if df.empty:
            print(f"  ⚠ Skipping {sheet_name} - no data")
            continue
        
        # Write dataframe to sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
        worksheet = writer.sheets[sheet_name]
        
        # Set column widths
        worksheet.set_column(0, 0, 50)  # Account column
        worksheet.set_column(1, len(df.columns)-1, 15)  # Year columns
        
        # Format header row
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_format)
        
        # Format data cells
        for row_num in range(1, len(df) + 1):
            # Account column
            worksheet.write(row_num, 0, df.iloc[row_num-1, 0], account_format)
            
            # Number columns
            for col_num in range(1, len(df.columns)):
                cell_value = df.iloc[row_num-1, col_num]
                # Handle NaN/None values
                if pd.isna(cell_value):
                    worksheet.write_blank(row_num, col_num, None, number_format)
                else:
                    worksheet.write(row_num, col_num, cell_value, number_format)
        
        # Freeze panes (first row and first column)
        worksheet.freeze_panes(1, 1)
        
        print(f"  ✓ Created sheet: {sheet_name} ({len(df)} rows)")

print(f"\n✓ Complete! Saved to {output_file}")
print(f"\nOutput structure:")
print(f"  • Balance Sheet     - Separate sheet")
print(f"  • Income Statement  - Separate sheet")
print(f"  • Cash Flow         - Separate sheet")
print("="*60)