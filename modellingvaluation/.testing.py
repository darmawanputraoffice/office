import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

class FinancialModel:
    """
    Comprehensive Financial Modeling Tool for analyzing financial statements
    Supports Balance Sheet, Income Statement, and Cash Flow Statement analysis
    """
    
    def __init__(self, excel_file_path, forecast_years=3):
        """
        Initialize the financial model with an Excel file containing financial statements
        
        Parameters:
        excel_file_path (str): Path to the Excel file with financial data
        forecast_years (int): Number of years to forecast (default: 3)
        """
        self.file_path = excel_file_path
        self.forecast_years = forecast_years
        self.balance_sheet = None
        self.income_statement = None
        self.cash_flow = None
        self.years = []
        self.forecast_data = {}
        
    def load_data(self):
        """Load financial statements from Excel file"""
        try:
            # Try different header rows
            for header_row in [0, 1]:
                try:
                    self.balance_sheet = pd.read_excel(self.file_path, sheet_name='Balance Sheet', header=header_row)
                    self.income_statement = pd.read_excel(self.file_path, sheet_name='Income Statement', header=header_row)
                    self.cash_flow = pd.read_excel(self.file_path, sheet_name='Cash Flow', header=header_row)
                    
                    # Clean column names
                    self.balance_sheet.columns = self.balance_sheet.columns.astype(str)
                    self.income_statement.columns = self.income_statement.columns.astype(str)
                    self.cash_flow.columns = self.cash_flow.columns.astype(str)
                    
                    print(f"Columns found: {list(self.balance_sheet.columns)[:10]}")
                    
                    # Extract years from column headers - try different patterns
                    # Pattern 1: Pure digits (2020, 2021, etc.)
                    self.years = [col for col in self.balance_sheet.columns if col.isdigit() and len(col) == 4]
                    
                    # Pattern 2: Check if columns contain 4-digit years
                    if not self.years:
                        self.years = [col for col in self.balance_sheet.columns 
                                     if any(year in str(col) for year in ['2020', '2021', '2022', '2023', '2024', '2025'])]
                    
                    # Pattern 3: Try to extract numeric columns
                    if not self.years:
                        for col in self.balance_sheet.columns:
                            try:
                                year_int = int(float(col))
                                if 2000 <= year_int <= 2030:
                                    self.years.append(str(year_int))
                            except:
                                pass
                    
                    if self.years:
                        print(f"✓ Data loaded successfully with header row {header_row}")
                        print(f"✓ Years available: {self.years}")
                        print(f"✓ Forecast period: {self.forecast_years} years")
                        return True
                        
                except Exception as e:
                    print(f"Trying header row {header_row} failed: {e}")
                    continue
            
            print("✗ Could not find year columns in the data")
            print("Please check that your Excel file has year columns (2020, 2021, etc.)")
            return False
            
        except Exception as e:
            print(f"✗ Error loading data: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def calculate_growth_rates(self, values):
        """Calculate year-over-year growth rates"""
        growth_rates = []
        for i in range(1, len(values)):
            if values[i-1] != 0 and not pd.isna(values[i-1]) and not pd.isna(values[i]):
                growth = ((values[i] - values[i-1]) / abs(values[i-1]))
                growth_rates.append(growth)
            else:
                growth_rates.append(0)
        return growth_rates
    
    def _get_value(self, df, keyword, year_col):
        """Helper function to extract value from dataframe"""
        try:
            # Get the first column name (Account column)
            account_col = df.columns[0]
            
            # Search for the keyword
            mask = df[account_col].astype(str).str.contains(keyword, case=False, na=False)
            row = df[mask]
            
            if not row.empty and year_col in df.columns:
                value = row[year_col].values[0]
                return float(value) if not pd.isna(value) and value != '' else 0
            return 0
        except Exception as e:
            return 0
    
    def forecast_income_statement(self):
        """Forecast income statement items"""
        print("\n" + "="*60)
        print("FORECASTING INCOME STATEMENT")
        print("="*60)
        
        # Get account column name
        account_col = self.income_statement.columns[0]
        
        # Create forecast dataframe
        forecast_df = self.income_statement.copy()
        
        # Generate forecast years
        last_year = int(self.years[-1])
        forecast_years = [str(last_year + i) for i in range(1, self.forecast_years + 1)]
        
        # Add forecast year columns
        for year in forecast_years:
            forecast_df[year] = 0.0
        
        # Get historical revenue to calculate growth
        revenue_keywords = ['Penjualan', 'pendapatan usaha']
        revenue_row = None
        
        for keyword in revenue_keywords:
            mask = self.income_statement[account_col].astype(str).str.contains(keyword, case=False, na=False)
            if mask.any():
                revenue_row = mask.idxmax()
                break
        
        if revenue_row is not None:
            # Get historical revenue values
            historical_revenue = []
            for year in self.years:
                val = self._get_value(self.income_statement, 'Penjualan', year)
                historical_revenue.append(val)
            
            # Calculate average growth rate
            growth_rates = self.calculate_growth_rates(historical_revenue)
            avg_growth = np.mean([g for g in growth_rates if g != 0]) if growth_rates else 0.05
            
            print(f"Historical Revenue Growth Rate: {avg_growth*100:.2f}%")
            print(f"Using growth rate for forecasting: {avg_growth*100:.2f}%")
            
            # Forecast each line item
            for idx, row in forecast_df.iterrows():
                account_name = str(row[account_col])
                
                if pd.isna(account_name) or account_name.strip() == '':
                    continue
                
                # Get historical values
                historical_values = []
                for year in self.years:
                    val = row[year] if year in row and not pd.isna(row[year]) else 0
                    historical_values.append(float(val) if val != '' else 0)
                
                # Skip if all historical values are zero
                if all(v == 0 for v in historical_values):
                    continue
                
                # Calculate item-specific growth or use revenue growth
                item_growth_rates = self.calculate_growth_rates(historical_values)
                if item_growth_rates and any(g != 0 for g in item_growth_rates):
                    item_growth = np.mean([g for g in item_growth_rates if g != 0])
                else:
                    item_growth = avg_growth
                
                # Forecast forward
                last_value = historical_values[-1]
                for i, year in enumerate(forecast_years):
                    forecast_value = last_value * (1 + item_growth) ** (i + 1)
                    forecast_df.at[idx, year] = forecast_value
        
        self.forecast_data['income_statement'] = forecast_df
        print("✓ Income statement forecasted")
        return forecast_df
    
    def forecast_balance_sheet(self):
        """Forecast balance sheet items"""
        print("\n" + "="*60)
        print("FORECASTING BALANCE SHEET")
        print("="*60)
        
        account_col = self.balance_sheet.columns[0]
        forecast_df = self.balance_sheet.copy()
        
        last_year = int(self.years[-1])
        forecast_years = [str(last_year + i) for i in range(1, self.forecast_years + 1)]
        
        for year in forecast_years:
            forecast_df[year] = 0.0
        
        # Get revenue from forecasted income statement
        if 'income_statement' in self.forecast_data:
            revenue_growth = 0.05  # Default
            
            # Calculate average asset growth
            asset_values = []
            for year in self.years:
                val = self._get_value(self.balance_sheet, 'Aset', year)
                if val > 0:
                    asset_values.append(val)
            
            if len(asset_values) > 1:
                asset_growth_rates = self.calculate_growth_rates(asset_values)
                revenue_growth = np.mean([g for g in asset_growth_rates if g != 0]) if asset_growth_rates else 0.05
            
            print(f"Using growth rate for balance sheet: {revenue_growth*100:.2f}%")
            
            # Forecast each line item
            for idx, row in forecast_df.iterrows():
                account_name = str(row[account_col])
                
                if pd.isna(account_name) or account_name.strip() == '':
                    continue
                
                historical_values = []
                for year in self.years:
                    val = row[year] if year in row and not pd.isna(row[year]) else 0
                    historical_values.append(float(val) if val != '' else 0)
                
                if all(v == 0 for v in historical_values):
                    continue
                
                # Use same growth rate for simplicity
                last_value = historical_values[-1]
                for i, year in enumerate(forecast_years):
                    forecast_value = last_value * (1 + revenue_growth) ** (i + 1)
                    forecast_df.at[idx, year] = forecast_value
        
        self.forecast_data['balance_sheet'] = forecast_df
        print("✓ Balance sheet forecasted")
        return forecast_df
    
    def forecast_cash_flow(self):
        """Forecast cash flow items"""
        print("\n" + "="*60)
        print("FORECASTING CASH FLOW")
        print("="*60)
        
        account_col = self.cash_flow.columns[0]
        forecast_df = self.cash_flow.copy()
        
        last_year = int(self.years[-1])
        forecast_years = [str(last_year + i) for i in range(1, self.forecast_years + 1)]
        
        for year in forecast_years:
            forecast_df[year] = 0.0
        
        # Calculate average operating cash flow growth
        ocf_values = []
        for year in self.years:
            val = self._get_value(self.cash_flow, 'Penerimaan dari pelanggan', year)
            if val > 0:
                ocf_values.append(val)
        
        ocf_growth = 0.05
        if len(ocf_values) > 1:
            ocf_growth_rates = self.calculate_growth_rates(ocf_values)
            ocf_growth = np.mean([g for g in ocf_growth_rates if g != 0]) if ocf_growth_rates else 0.05
        
        print(f"Using growth rate for cash flow: {ocf_growth*100:.2f}%")
        
        # Forecast each line item
        for idx, row in forecast_df.iterrows():
            account_name = str(row[account_col])
            
            if pd.isna(account_name) or account_name.strip() == '':
                continue
            
            historical_values = []
            for year in self.years:
                val = row[year] if year in row and not pd.isna(row[year]) else 0
                historical_values.append(float(val) if val != '' else 0)
            
            if all(v == 0 for v in historical_values):
                continue
            
            last_value = historical_values[-1]
            for i, year in enumerate(forecast_years):
                forecast_value = last_value * (1 + ocf_growth) ** (i + 1)
                forecast_df.at[idx, year] = forecast_value
        
        self.forecast_data['cash_flow'] = forecast_df
        print("✓ Cash flow forecasted")
        return forecast_df
    
    def calculate_financial_ratios(self):
        """Calculate key financial ratios for all years including forecast"""
        print("\n" + "="*60)
        print("CALCULATING FINANCIAL RATIOS")
        print("="*60)
        
        last_year = int(self.years[-1])
        forecast_years = [str(last_year + i) for i in range(1, self.forecast_years + 1)]
        all_years = self.years + forecast_years
        
        ratios_data = {
            'Metric': [],
            'Category': []
        }
        
        for year in all_years:
            ratios_data[year] = []
        
        # Use forecasted data if available
        bs_df = self.forecast_data.get('balance_sheet', self.balance_sheet)
        is_df = self.forecast_data.get('income_statement', self.income_statement)
        cf_df = self.forecast_data.get('cash_flow', self.cash_flow)
        
        # Define metrics to calculate
        metrics = [
            ('Revenue', 'Income Statement', 'Penjualan'),
            ('Gross Profit', 'Income Statement', 'laba bruto'),
            ('Operating Income', 'Income Statement', 'laba (rugi) dari operasi'),
            ('Net Income', 'Income Statement', 'Jumlah laba \\(rugi\\)'),
            ('Total Assets', 'Balance Sheet', 'Aset'),
            ('Operating Cash Flow', 'Cash Flow', 'Kas diperoleh dari'),
        ]
        
        for metric_name, category, keyword in metrics:
            ratios_data['Metric'].append(metric_name)
            ratios_data['Category'].append(category)
            
            for year in all_years:
                if category == 'Income Statement':
                    value = self._get_value(is_df, keyword, year)
                elif category == 'Balance Sheet':
                    value = self._get_value(bs_df, keyword, year)
                else:
                    value = self._get_value(cf_df, keyword, year)
                
                ratios_data[year].append(value)
        
        # Calculate ratios
        ratio_metrics = [
            ('Gross Margin (%)', 'Profitability'),
            ('Operating Margin (%)', 'Profitability'),
            ('Net Margin (%)', 'Profitability'),
            ('ROA (%)', 'Profitability'),
        ]
        
        for ratio_name, category in ratio_metrics:
            ratios_data['Metric'].append(ratio_name)
            ratios_data['Category'].append(category)
            
            for year in all_years:
                revenue_idx = 0  # Revenue is first metric
                net_income_idx = 3
                total_assets_idx = 4
                
                revenue = ratios_data[year][revenue_idx]
                
                if ratio_name == 'Gross Margin (%)':
                    gross_profit = ratios_data[year][1]
                    value = (gross_profit / revenue * 100) if revenue != 0 else 0
                elif ratio_name == 'Operating Margin (%)':
                    operating_income = ratios_data[year][2]
                    value = (operating_income / revenue * 100) if revenue != 0 else 0
                elif ratio_name == 'Net Margin (%)':
                    net_income = ratios_data[year][net_income_idx]
                    value = (net_income / revenue * 100) if revenue != 0 else 0
                elif ratio_name == 'ROA (%)':
                    net_income = ratios_data[year][net_income_idx]
                    total_assets = ratios_data[year][total_assets_idx]
                    value = (net_income / total_assets * 100) if total_assets != 0 else 0
                else:
                    value = 0
                
                ratios_data[year].append(value)
        
        ratios_df = pd.DataFrame(ratios_data)
        self.forecast_data['ratios'] = ratios_df
        print("✓ Financial ratios calculated")
        return ratios_df
    
    def export_to_excel(self, output_file='Financial_Model_Output.xlsx'):
        """Export all forecasts and analysis to Excel"""
        print("\n" + "="*60)
        print("EXPORTING TO EXCEL")
        print("="*60)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write forecasted statements
            if 'income_statement' in self.forecast_data:
                self.forecast_data['income_statement'].to_excel(
                    writer, sheet_name='Income Statement Forecast', index=False
                )
                print("✓ Income Statement Forecast exported")
            
            if 'balance_sheet' in self.forecast_data:
                self.forecast_data['balance_sheet'].to_excel(
                    writer, sheet_name='Balance Sheet Forecast', index=False
                )
                print("✓ Balance Sheet Forecast exported")
            
            if 'cash_flow' in self.forecast_data:
                self.forecast_data['cash_flow'].to_excel(
                    writer, sheet_name='Cash Flow Forecast', index=False
                )
                print("✓ Cash Flow Forecast exported")
            
            if 'ratios' in self.forecast_data:
                self.forecast_data['ratios'].to_excel(
                    writer, sheet_name='Financial Ratios', index=False
                )
                print("✓ Financial Ratios exported")
        
        print(f"\n✓ All data exported successfully to: {output_file}")
        return output_file
    
    def run_full_model(self, output_file='Financial_Model_Output.xlsx'):
        """Run the complete financial model"""
        print("\n" + "="*70)
        print(" "*20 + "FINANCIAL MODEL")
        print("="*70)
        
        # Load data
        if not self.load_data():
            return False
        
        # Run forecasts
        self.forecast_income_statement()
        self.forecast_balance_sheet()
        self.forecast_cash_flow()
        self.calculate_financial_ratios()
        
        # Export results
        self.export_to_excel(output_file)
        
        print("\n" + "="*70)
        print(" "*20 + "MODEL COMPLETE")
        print("="*70)
        
        return True


# Main execution
if __name__ == "__main__":
    # Initialize the financial model
    model = FinancialModel("MEDC_FS_Consolidated.xlsx", forecast_years=3)
    
    # Run the full model
    model.run_full_model(output_file='Financial_Model_Output.xlsx')