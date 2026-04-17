import pandas as pd


class InputHandler:
    def __init__(self, filepath):
        self.filepath = filepath
        self.proj_years = ['FY2025E', 'FY2026E', 'FY2027E', 'FY2028E', 'FY2029E']
        self.hist_years = ['FY2022A', 'FY2023A', 'FY2024A']
        self._extra_years = ['FY2028E', 'FY2029E']  # years not in template — extrapolated

    def _safe_read(self, sheet_name):
        """Safely read a sheet, return empty DataFrame on failure."""
        try:
            df = pd.read_excel(self.filepath, sheet_name=sheet_name, header=None)
            return df
        except Exception as e:
            print(f"Warning: Could not read sheet '{sheet_name}': {e}")
            return pd.DataFrame()

    def get_assumptions(self):
        """Read Assumptions sheet — returns dict of {label: {year: value}}."""
        df = self._safe_read('Assumptions')
        if df.empty:
            return {}

        # Find header row (contains FY2022A etc.)
        header_row = None
        for i, row in df.iterrows():
            if any('FY' in str(v) for v in row.values if v is not None):
                header_row = i
                break

        if header_row is None:
            print("Warning: Could not find header row in Assumptions.")
            return {}

        headers = list(df.iloc[header_row])
        result = {}

        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            label = str(row.iloc[0]).strip() if row.iloc[0] is not None else ''
            # Skip blanks, section headers (no values), and description rows (start with spaces)
            if not label or label == 'nan' or label.startswith('  ') or label.isupper():
                continue
            row_data = {}
            for j, header in enumerate(headers):
                if str(header).startswith('FY') and j < len(row):
                    val = row.iloc[j]
                    if val is not None and str(val) != 'nan':
                        try:
                            row_data[str(header)] = float(val)
                        except (ValueError, TypeError):
                            row_data[str(header)] = val
            if row_data:
                result[label] = row_data

        # Extrapolate FY2028E and FY2029E using FY2027E values (template only goes to FY2027E)
        for label, year_data in result.items():
            last_val = year_data.get('FY2027E')
            if last_val is not None:
                year_data.setdefault('FY2028E', last_val)
                year_data.setdefault('FY2029E', last_val)

        # Extend loan balances using the growth rate
        if 'Total Loans (Ending, $mm)' in result and 'Loan Growth Rate (%)' in result:
            loans  = result['Total Loans (Ending, $mm)']
            growth = result['Loan Growth Rate (%)'].get('FY2027E', 0.07)
            loans['FY2028E'] = loans.get('FY2027E', 0) * (1 + growth)
            loans['FY2029E'] = loans.get('FY2028E', 0) * (1 + growth)

        # Extend avg earning assets by same YoY growth rate
        if 'Avg. Earning Assets ($mm)' in result:
            ea = result['Avg. Earning Assets ($mm)']
            base_26 = ea.get('FY2026E', 1)
            base_27 = ea.get('FY2027E', 0)
            growth_rate = (base_27 / base_26 - 1) if base_26 else 0.06
            ea['FY2028E'] = base_27 * (1 + growth_rate)
            ea['FY2029E'] = ea['FY2028E'] * (1 + growth_rate)

        # Extend non-interest income lines by same YoY increment
        for label in ['Fee & Commission Income ($mm)', 'Trading Income ($mm)', 'Other Non-Interest Income ($mm)']:
            if label in result:
                d = result[label]
                increment = d.get('FY2027E', 0) - d.get('FY2026E', 0)
                d['FY2028E'] = d.get('FY2027E', 0) + increment
                d['FY2029E'] = d['FY2028E'] + increment

        return result

    def _read_statement(self, sheet_name):
        """Read a financial statement sheet — returns dict of {label: {year: value}}."""
        df = self._safe_read(sheet_name)
        if df.empty:
            return {}

        # Find header row
        header_row = None
        for i, row in df.iterrows():
            if any('FY' in str(v) for v in row.values if v is not None):
                header_row = i
                break

        if header_row is None:
            return {}

        headers = list(df.iloc[header_row])
        result = {}

        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            label = str(row.iloc[0]).strip() if row.iloc[0] is not None else ''
            if not label or label == 'nan':
                continue
            row_data = {}
            for j, header in enumerate(headers):
                if str(header).startswith('FY') and j < len(row):
                    val = row.iloc[j]
                    if val is not None and str(val) != 'nan':
                        try:
                            row_data[str(header)] = float(val)
                        except (ValueError, TypeError):
                            pass
            if row_data:
                result[label] = row_data

        return result

    def get_income_statement(self):
        return self._read_statement('Income Statement')

    def get_balance_sheet(self):
        return self._read_statement('Balance Sheet')

    def get_capital_adequacy(self):
        return self._read_statement('Capital Adequacy')

    def get_cash_flow(self):
        return self._read_statement('Cash Flow')
