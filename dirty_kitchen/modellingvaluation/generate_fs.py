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

# Create output file
output_file = "MEDC_FS_Consolidated.xlsx"
print(f"\nCreating {output_file}...")

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet('Consolidated')
    
    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'bg_color': '#366092',
        'font_color': 'white'
    })
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    
    row = 0
    
    # Balance Sheet
    worksheet.write(row, 0, 'BALANCE SHEET', header_format)
    row += 1
    balance_sheet.to_excel(writer, sheet_name='Consolidated', startrow=row, index=False)
    worksheet.set_column(0, 0, 60)
    for col in range(1, len(balance_sheet.columns)):
        worksheet.set_column(col, col, 15, number_format)
    row += len(balance_sheet) + 2
    
    # Income Statement
    worksheet.write(row, 0, 'INCOME STATEMENT', header_format)
    row += 1
    income_statement.to_excel(writer, sheet_name='Consolidated', startrow=row, index=False)
    for col in range(1, len(income_statement.columns)):
        worksheet.set_column(col, col, 15, number_format)
    row += len(income_statement) + 2
    
    # Cash Flow
    worksheet.write(row, 0, 'CASH FLOW STATEMENT', header_format)
    row += 1
    cashflow.to_excel(writer, sheet_name='Consolidated', startrow=row, index=False)
    for col in range(1, len(cashflow.columns)):
        worksheet.set_column(col, col, 15, number_format)

print(f"✓ Complete! Saved to {output_file}")
print("="*60)