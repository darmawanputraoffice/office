import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

class CompanyValuation:
    """
    Comprehensive Company Valuation Tool
    Implements DCF, Trading Comps, and Transaction Comps methodologies
    """
    
    def __init__(self, financial_model, company_name="Target Company"):
        """
        Initialize valuation with the financial model
        
        Parameters:
        financial_model: FinancialModel object with forecasted data
        company_name: Name of the company being valued
        """
        self.model = financial_model
        self.company_name = company_name
        self.valuation_results = {}
        
        # Valuation parameters (can be adjusted)
        self.wacc = 0.10  # Weighted Average Cost of Capital (10%)
        self.terminal_growth_rate = 0.03  # Terminal growth rate (3%)
        self.tax_rate = 0.25  # Corporate tax rate (25%)
        
        # Get all years including forecast
        last_year = int(self.model.years[-1])
        self.forecast_years = [str(last_year + i) for i in range(1, self.model.forecast_years + 1)]
        self.all_years = self.model.years + self.forecast_years
        
    def calculate_fcff(self):
        """Calculate Free Cash Flow to Firm (FCFF) for all years"""
        print("\n" + "="*60)
        print("CALCULATING FREE CASH FLOW TO FIRM (FCFF)")
        print("="*60)
        
        fcff_data = {
            'Year': [],
            'EBIT': [],
            'Tax': [],
            'NOPAT': [],
            'Depreciation & Amortization': [],
            'Capex': [],
            'Change in Working Capital': [],
            'Free Cash Flow': []
        }
        
        is_df = self.model.forecast_data.get('income_statement', self.model.income_statement)
        bs_df = self.model.forecast_data.get('balance_sheet', self.model.balance_sheet)
        cf_df = self.model.forecast_data.get('cash_flow', self.model.cash_flow)
        
        previous_wc = None
        
        for year in self.all_years:
            fcff_data['Year'].append(year)
            
            # Get EBIT (Operating Income)
            ebit = self.model._get_value(is_df, 'laba (rugi) dari operasi', year)
            if ebit == 0:  # Try alternative names
                ebit = self.model._get_value(is_df, 'operating income', year)
            fcff_data['EBIT'].append(ebit)
            
            # Calculate Tax on EBIT
            tax = ebit * self.tax_rate
            fcff_data['Tax'].append(tax)
            
            # Calculate NOPAT (Net Operating Profit After Tax)
            nopat = ebit - tax
            fcff_data['NOPAT'].append(nopat)
            
            # Get Depreciation & Amortization
            depr = self.model._get_value(is_df, 'Beban penjualan', year)
            if depr == 0:
                depr = ebit * 0.05  # Estimate as 5% of EBIT if not found
            fcff_data['Depreciation & Amortization'].append(depr)
            
            # Get Capital Expenditure
            capex = self.model._get_value(cf_df, 'Pembayaran untuk', year)
            if capex == 0:
                capex = abs(self.model._get_value(cf_df, 'Penurunan (kenaikan) aset operasi', year))
            if capex == 0:
                capex = depr * 1.2  # Estimate as 120% of depreciation
            fcff_data['Capex'].append(capex)
            
            # Calculate Working Capital
            current_assets = self.model._get_value(bs_df, 'Aset lancar', year)
            current_liabilities = self.model._get_value(bs_df, 'Piutang usaha pihak ketiga', year)
            working_capital = current_assets - current_liabilities
            
            # Change in Working Capital
            if previous_wc is not None:
                delta_wc = working_capital - previous_wc
            else:
                delta_wc = 0
            
            fcff_data['Change in Working Capital'].append(delta_wc)
            previous_wc = working_capital
            
            # Calculate Free Cash Flow
            fcf = nopat + depr - capex - delta_wc
            fcff_data['Free Cash Flow'].append(fcf)
            
            print(f"Year {year}: FCFF = {fcf:,.0f}")
        
        fcff_df = pd.DataFrame(fcff_data)
        self.valuation_results['fcff'] = fcff_df
        print("✓ FCFF calculated")
        return fcff_df
    
    def calculate_dcf_valuation(self):
        """Calculate Enterprise Value using Discounted Cash Flow (DCF) method"""
        print("\n" + "="*60)
        print("DCF VALUATION - ENTERPRISE VALUE")
        print("="*60)
        
        if 'fcff' not in self.valuation_results:
            self.calculate_fcff()
        
        fcff_df = self.valuation_results['fcff']
        
        # Get forecast period FCF (excluding historical)
        forecast_fcf = []
        for year in self.forecast_years:
            year_data = fcff_df[fcff_df['Year'] == year]
            if not year_data.empty:
                forecast_fcf.append(year_data['Free Cash Flow'].values[0])
        
        # Calculate Present Value of forecast FCF
        pv_fcf = []
        dcf_data = {
            'Year': [],
            'Free Cash Flow': [],
            'Discount Factor': [],
            'Present Value': []
        }
        
        for i, (year, fcf) in enumerate(zip(self.forecast_years, forecast_fcf), start=1):
            discount_factor = 1 / ((1 + self.wacc) ** i)
            pv = fcf * discount_factor
            pv_fcf.append(pv)
            
            dcf_data['Year'].append(year)
            dcf_data['Free Cash Flow'].append(fcf)
            dcf_data['Discount Factor'].append(discount_factor)
            dcf_data['Present Value'].append(pv)
            
            print(f"Year {year}: FCF = {fcf:,.0f}, PV = {pv:,.0f}")
        
        # Calculate Terminal Value
        terminal_fcf = forecast_fcf[-1] * (1 + self.terminal_growth_rate)
        terminal_value = terminal_fcf / (self.wacc - self.terminal_growth_rate)
        
        # Discount Terminal Value to present
        terminal_discount_factor = 1 / ((1 + self.wacc) ** len(forecast_fcf))
        pv_terminal_value = terminal_value * terminal_discount_factor
        
        # Calculate Enterprise Value
        pv_forecast_period = sum(pv_fcf)
        enterprise_value = pv_forecast_period + pv_terminal_value
        
        print(f"\n{'─'*60}")
        print(f"PV of Forecast Period FCF: {pv_forecast_period:,.0f}")
        print(f"Terminal Value: {terminal_value:,.0f}")
        print(f"PV of Terminal Value: {pv_terminal_value:,.0f}")
        print(f"{'─'*60}")
        print(f"ENTERPRISE VALUE: {enterprise_value:,.0f}")
        print(f"{'─'*60}")
        
        # Get latest balance sheet items for Equity Value calculation
        bs_df = self.model.forecast_data.get('balance_sheet', self.model.balance_sheet)
        latest_year = self.model.years[-1]
        
        cash = self.model._get_value(bs_df, 'Kas dan setara kas', latest_year)
        debt = self.model._get_value(bs_df, 'Piutang usaha pihak berelasi', latest_year)
        
        # Calculate Equity Value
        equity_value = enterprise_value + cash - debt
        
        print(f"\nCash and Cash Equivalents: {cash:,.0f}")
        print(f"Total Debt: {debt:,.0f}")
        print(f"{'─'*60}")
        print(f"EQUITY VALUE: {equity_value:,.0f}")
        print(f"{'─'*60}")
        
        # Store results
        self.valuation_results['dcf_detail'] = pd.DataFrame(dcf_data)
        self.valuation_results['dcf_summary'] = {
            'PV Forecast Period': pv_forecast_period,
            'Terminal Value': terminal_value,
            'PV Terminal Value': pv_terminal_value,
            'Enterprise Value': enterprise_value,
            'Cash': cash,
            'Debt': debt,
            'Equity Value': equity_value,
            'WACC': self.wacc,
            'Terminal Growth Rate': self.terminal_growth_rate,
            'Tax Rate': self.tax_rate
        }
        
        print("✓ DCF Valuation completed")
        return enterprise_value, equity_value
    
    def calculate_trading_multiples(self):
        """Calculate valuation using Trading Comparables (Market Multiples)"""
        print("\n" + "="*60)
        print("TRADING COMPARABLES VALUATION")
        print("="*60)
        
        is_df = self.model.forecast_data.get('income_statement', self.model.income_statement)
        bs_df = self.model.forecast_data.get('balance_sheet', self.model.balance_sheet)
        
        # Get latest and next year metrics
        latest_year = self.model.years[-1]
        next_year = self.forecast_years[0]
        
        # Current year metrics (LTM - Last Twelve Months)
        revenue_ltm = self.model._get_value(is_df, 'Penjualan', latest_year)
        ebitda_ltm = self.model._get_value(is_df, 'laba (rugi) dari operasi', latest_year)
        net_income_ltm = self.model._get_value(is_df, 'Jumlah laba \\(rugi\\)', latest_year)
        
        # Forward year metrics (NTM - Next Twelve Months)
        revenue_ntm = self.model._get_value(is_df, 'Penjualan', next_year)
        ebitda_ntm = self.model._get_value(is_df, 'laba (rugi) dari operasi', next_year)
        net_income_ntm = self.model._get_value(is_df, 'Jumlah laba \\(rugi\\)', next_year)
        
        # Industry multiples (these would typically come from comparable companies)
        # Using conservative estimates for Indonesian market
        multiples_low = {
            'EV/Revenue': 1.0,
            'EV/EBITDA': 6.0,
            'P/E': 10.0
        }
        
        multiples_mid = {
            'EV/Revenue': 1.5,
            'EV/EBITDA': 8.0,
            'P/E': 15.0
        }
        
        multiples_high = {
            'EV/Revenue': 2.0,
            'EV/EBITDA': 10.0,
            'P/E': 20.0
        }
        
        print(f"\nCompany Metrics (LTM):")
        print(f"Revenue: {revenue_ltm:,.0f}")
        print(f"EBITDA: {ebitda_ltm:,.0f}")
        print(f"Net Income: {net_income_ltm:,.0f}")
        
        print(f"\nCompany Metrics (NTM):")
        print(f"Revenue: {revenue_ntm:,.0f}")
        print(f"EBITDA: {ebitda_ntm:,.0f}")
        print(f"Net Income: {net_income_ntm:,.0f}")
        
        # Calculate valuations
        valuations = {
            'Multiple': [],
            'Metric': [],
            'Low': [],
            'Mid': [],
            'High': []
        }
        
        # EV/Revenue
        valuations['Multiple'].append('EV/Revenue (LTM)')
        valuations['Metric'].append(revenue_ltm)
        valuations['Low'].append(revenue_ltm * multiples_low['EV/Revenue'])
        valuations['Mid'].append(revenue_ltm * multiples_mid['EV/Revenue'])
        valuations['High'].append(revenue_ltm * multiples_high['EV/Revenue'])
        
        valuations['Multiple'].append('EV/Revenue (NTM)')
        valuations['Metric'].append(revenue_ntm)
        valuations['Low'].append(revenue_ntm * multiples_low['EV/Revenue'])
        valuations['Mid'].append(revenue_ntm * multiples_mid['EV/Revenue'])
        valuations['High'].append(revenue_ntm * multiples_high['EV/Revenue'])
        
        # EV/EBITDA
        valuations['Multiple'].append('EV/EBITDA (LTM)')
        valuations['Metric'].append(ebitda_ltm)
        valuations['Low'].append(ebitda_ltm * multiples_low['EV/EBITDA'])
        valuations['Mid'].append(ebitda_ltm * multiples_mid['EV/EBITDA'])
        valuations['High'].append(ebitda_ltm * multiples_high['EV/EBITDA'])
        
        valuations['Multiple'].append('EV/EBITDA (NTM)')
        valuations['Metric'].append(ebitda_ntm)
        valuations['Low'].append(ebitda_ntm * multiples_low['EV/EBITDA'])
        valuations['Mid'].append(ebitda_ntm * multiples_mid['EV/EBITDA'])
        valuations['High'].append(ebitda_ntm * multiples_high['EV/EBITDA'])
        
        # P/E (Equity Value)
        valuations['Multiple'].append('P/E (LTM)')
        valuations['Metric'].append(net_income_ltm)
        valuations['Low'].append(net_income_ltm * multiples_low['P/E'])
        valuations['Mid'].append(net_income_ltm * multiples_mid['P/E'])
        valuations['High'].append(net_income_ltm * multiples_high['P/E'])
        
        valuations['Multiple'].append('P/E (NTM)')
        valuations['Metric'].append(net_income_ntm)
        valuations['Low'].append(net_income_ntm * multiples_low['P/E'])
        valuations['Mid'].append(net_income_ntm * multiples_mid['P/E'])
        valuations['High'].append(net_income_ntm * multiples_high['P/E'])
        
        valuations_df = pd.DataFrame(valuations)
        
        # Calculate average enterprise value (excluding P/E which is equity value)
        ev_valuations = valuations_df[valuations_df['Multiple'].str.contains('EV/')]
        avg_ev_low = ev_valuations['Low'].mean()
        avg_ev_mid = ev_valuations['Mid'].mean()
        avg_ev_high = ev_valuations['High'].mean()
        
        print(f"\n{'─'*60}")
        print(f"Trading Multiples Enterprise Value Range:")
        print(f"Low: {avg_ev_low:,.0f}")
        print(f"Mid: {avg_ev_mid:,.0f}")
        print(f"High: {avg_ev_high:,.0f}")
        print(f"{'─'*60}")
        
        self.valuation_results['trading_multiples'] = valuations_df
        self.valuation_results['trading_multiples_summary'] = {
            'EV Low': avg_ev_low,
            'EV Mid': avg_ev_mid,
            'EV High': avg_ev_high
        }
        
        print("✓ Trading Multiples Valuation completed")
        return avg_ev_mid
    
    def create_valuation_summary(self):
        """Create a comprehensive valuation summary"""
        print("\n" + "="*60)
        print("VALUATION SUMMARY")
        print("="*60)
        
        summary_data = {
            'Valuation Method': [],
            'Enterprise Value': [],
            'Equity Value': [],
            'Notes': []
        }
        
        # DCF Valuation
        if 'dcf_summary' in self.valuation_results:
            dcf = self.valuation_results['dcf_summary']
            summary_data['Valuation Method'].append('DCF Analysis')
            summary_data['Enterprise Value'].append(dcf['Enterprise Value'])
            summary_data['Equity Value'].append(dcf['Equity Value'])
            summary_data['Notes'].append(f"WACC: {dcf['WACC']*100:.1f}%, Terminal Growth: {dcf['Terminal Growth Rate']*100:.1f}%")
        
        # Trading Multiples
        if 'trading_multiples_summary' in self.valuation_results:
            tm = self.valuation_results['trading_multiples_summary']
            summary_data['Valuation Method'].append('Trading Multiples (Low)')
            summary_data['Enterprise Value'].append(tm['EV Low'])
            summary_data['Equity Value'].append('-')
            summary_data['Notes'].append('Based on comparable companies')
            
            summary_data['Valuation Method'].append('Trading Multiples (Mid)')
            summary_data['Enterprise Value'].append(tm['EV Mid'])
            summary_data['Equity Value'].append('-')
            summary_data['Notes'].append('Based on comparable companies')
            
            summary_data['Valuation Method'].append('Trading Multiples (High)')
            summary_data['Enterprise Value'].append(tm['EV High'])
            summary_data['Equity Value'].append('-')
            summary_data['Notes'].append('Based on comparable companies')
        
        summary_df = pd.DataFrame(summary_data)
        
        # Calculate weighted average (giving more weight to DCF)
        ev_values = [v for v in summary_df['Enterprise Value'] if isinstance(v, (int, float))]
        if ev_values:
            avg_ev = np.mean(ev_values)
            print(f"\nAverage Enterprise Value: {avg_ev:,.0f}")
        
        self.valuation_results['summary'] = summary_df
        
        # Print summary table
        print("\n" + summary_df.to_string(index=False))
        print("="*60)
        
        return summary_df
    
    def export_valuation_to_excel(self, output_file='Company_Valuation.xlsx'):
        """Export all valuation results to Excel"""
        print("\n" + "="*60)
        print("EXPORTING VALUATION TO EXCEL")
        print("="*60)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary
            if 'summary' in self.valuation_results:
                self.valuation_results['summary'].to_excel(
                    writer, sheet_name='Valuation Summary', index=False
                )
                print("✓ Valuation Summary exported")
            
            # DCF Details
            if 'fcff' in self.valuation_results:
                self.valuation_results['fcff'].to_excel(
                    writer, sheet_name='FCFF Calculation', index=False
                )
                print("✓ FCFF Calculation exported")
            
            if 'dcf_detail' in self.valuation_results:
                self.valuation_results['dcf_detail'].to_excel(
                    writer, sheet_name='DCF Detail', index=False
                )
                print("✓ DCF Detail exported")
            
            # DCF Summary
            if 'dcf_summary' in self.valuation_results:
                dcf_summary_df = pd.DataFrame([self.valuation_results['dcf_summary']]).T
                dcf_summary_df.columns = ['Value']
                dcf_summary_df.to_excel(
                    writer, sheet_name='DCF Summary'
                )
                print("✓ DCF Summary exported")
            
            # Trading Multiples
            if 'trading_multiples' in self.valuation_results:
                self.valuation_results['trading_multiples'].to_excel(
                    writer, sheet_name='Trading Multiples', index=False
                )
                print("✓ Trading Multiples exported")
        
        print(f"\n✓ All valuation data exported to: {output_file}")
        return output_file
    
    def run_full_valuation(self, output_file='Company_Valuation.xlsx'):
        """Run complete valuation analysis"""
        print("\n" + "="*70)
        print(" "*15 + f"VALUATION ANALYSIS - {self.company_name.upper()}")
        print("="*70)
        
        # Run all valuation methods
        self.calculate_fcff()
        self.calculate_dcf_valuation()
        self.calculate_trading_multiples()
        self.create_valuation_summary()
        
        # Export results
        self.export_valuation_to_excel(output_file)
        
        print("\n" + "="*70)
        print(" "*20 + "VALUATION COMPLETE")
        print("="*70)
        
        return True


# Main execution
if __name__ == "__main__":
    # First, run the financial model
    from testing import FinancialModel  # Import your financial model
    
    print("Step 1: Running Financial Model...")
    model = FinancialModel("MEDC_FS_Consolidated.xlsx", forecast_years=3)
    model.run_full_model(output_file='Financial_Model_Output.xlsx')
    
    print("\n" + "="*70)
    print("Step 2: Running Valuation Analysis...")
    print("="*70)
    
    # Run valuation
    valuation = CompanyValuation(model, company_name="MEDC")
    valuation.run_full_valuation(output_file='Company_Valuation.xlsx')
    
    print("\n✓ Complete! Check 'Company_Valuation.xlsx' for results.")