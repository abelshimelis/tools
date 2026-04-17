import pandas as pd
import openpyxl


class InputHandler:
    def __init__(self, filepath):
        self.filepath = filepath

    def get_statements(self):
        income = pd.read_excel(self.filepath, sheet_name="Income Statement", header=1, index_col=0)
        print("Income columns:", income.columns.tolist())
        income = income.dropna(axis=1, how='all')
        print("After dropna:", income.columns.tolist())
        income = income[[c for c in income.columns if str(c).endswith('A')]]
        print("After filtering 'A':", income.columns.tolist())

        balance = pd.read_excel(self.filepath, sheet_name="Balance Sheet", header=1, index_col=0)
        balance = balance.dropna(axis=1, how='all')
        balance = balance[[c for c in balance.columns if str(c).endswith('A')]]

        cashflow = pd.read_excel(self.filepath, sheet_name="Cash Flow Statement", header=1, index_col=0)
        cashflow = cashflow.dropna(axis=1, how='all')
        cashflow = cashflow[[c for c in cashflow.columns if str(c).endswith('A')]]

        return {"income": income, "balance": balance, "cashflow": cashflow}

    def get_assumptions(self):
        df = pd.read_excel(self.filepath, sheet_name="Assumptions", header=None)
        # Remove blank rows
        df = df[df[0].notna()]
        # Remove description rows (they start with 2 spaces)
        df = df[~df[0].astype(str).str.startswith('  ')]
        # Remove section headers (no values in year columns)
        df = df[df[[2, 3, 4, 5, 6]].notna().any(axis=1)]
        df = df.set_index(0)
        df = df[[2, 3, 4, 5, 6]]
        df.columns = ['2025P', '2026P', '2027P', '2028P', '2029P']
        # Remove duplicate labels — keep first occurrence
        df = df[~df.index.duplicated(keep='first')]
        return df

    def get_owc(self):
        df = pd.read_excel(self.filepath, sheet_name="OWC Schedule", header=1, index_col=0)
        return df.dropna(axis=1, how='all')

    def get_depreciation_schedule(self):
        df = pd.read_excel(self.filepath, sheet_name="Depreciation Schedule", header=1, index_col=0)
        return df.dropna(axis=1, how='all')

    def get_debt_schedule(self):
        df = pd.read_excel(self.filepath, sheet_name="Debt Schedule", header=1, index_col=0)
        return df.dropna(axis=1, how='all')