import pandas as pd


class FinancialModel:
    def __init__(self, statements, assumptions, owc, depreciation, debt):
        self.income_hist  = statements["income"]
        self.balance_hist = statements["balance"]
        self.assumptions  = assumptions
        self.years = ['2025P', '2026P', '2027P', '2028P', '2029P']

        # Friendly built-in defaults so projections run even if template inputs are incomplete
        self.assumption_defaults = {
            'Net sales growth %': 0.10,
            'Membership income growth %': 0.05,
            'COGS % of revenue': 0.40,
            'SG&A % of revenue': 0.20,
            'Shares outstanding (mm)': 100.0,
            'Effective Tax Rate (ETR) %': 0.25,
            'PP&E useful life — Straight-Line (yrs)': 20.0,
            'CAPEX useful life — Straight-Line (yrs)': 10.0,
            'CAPEX % of revenue': 0.05,
            'Mandatory ST debt change ($mm)': 0.0,
            'Mandatory LT debt change ($mm)': 0.0,
            'Short-term interest rate %': 0.05,
            'Long-term interest rate %': 0.06,
            'Capital lease interest ($mm)': 0.0,
            'Cash interest rate %': 0.02,
            'Deferred Tax ($mm) — non-cash timing difference': 0.0,
            'Days receivables (DSO)': 60.0,
            'Inventory turnover days (DIO)': 90.0,
            'Days prepaid': 30.0,
            'Days payable — AP (DPO, on COGS)': 45.0,
            'Days payable — Accrued liabilities': 30.0,
            'Days payable — Accrued tax': 30.0,
            'Other operating activities ($mm)': 0.0,
            'Investments & acquisitions ($mm)': 0.0,
            'Dividend payout ratio %': 0.20,
            'Share issuance / (buyback) ($mm)': 0.0,
        }

    def a(self, label, year):
        """Get assumption value for a label and projection year."""
        try:
            val = self.assumptions.loc[label, year]
            if isinstance(val, pd.Series):
                val = val.iloc[0]
            if pd.isna(val):
                raise KeyError
            return float(val)
        except (KeyError, TypeError):
            if label in self.assumption_defaults:
                return float(self.assumption_defaults[label])
            raise KeyError(f"Missing assumption '{label}' in template and no fallback default is set")

    def _hist(self, label, stmt='income'):
        """Get the 2024A (last historical) value from income or balance sheet."""
        df = self.income_hist if stmt == 'income' else self.balance_hist
        try:
            val = df.loc[label, '2024A']
            return float(val) if pd.notna(val) else 0.0
        except (KeyError, TypeError):
            return 0.0

    @staticmethod
    def _safe_div(numerator, denominator):
        if denominator is None or denominator == 0 or pd.isna(denominator):
            return 0.0
        return numerator / denominator

    def project(self):
        results = []

        # ── Base year values from 2024A ────────────────────────────────
        prev_net_sales   = self._hist('Net sales')
        prev_membership  = self._hist('Membership and other income')
        prev_cash        = self._hist('Cash and cash equivalents', 'balance')
        prev_ppe         = self._hist('Property, plant & equipment, net', 'balance')
        prev_st_debt     = self._hist('Short-term borrowings', 'balance')
        prev_lt_debt     = self._hist('Long-term debt', 'balance')
        prev_retained    = self._hist('Retained earnings', 'balance')
        prev_ar          = self._hist('Receivables, net', 'balance')
        prev_inv         = self._hist('Inventories', 'balance')
        prev_prepaid     = self._hist('Prepaid expenses and other', 'balance')
        prev_ap          = self._hist('Accounts payable', 'balance')
        prev_al          = self._hist('Accrued liabilities', 'balance')
        prev_at          = self._hist('Accrued income taxes', 'balance')
        share_capital    = self._hist('Share capital & premium', 'balance')
        goodwill         = self._hist('Goodwill', 'balance')
        other_assets     = self._hist('Other assets & deferred charges', 'balance')

        for year in self.years:
            r = {'Year': year}

            # ── INCOME STATEMENT ──────────────────────────────────────────
            r['Net Sales']         = prev_net_sales * (1 + self.a('Net sales growth %', year))
            r['Membership Income'] = prev_membership * (1 + self.a('Membership income growth %', year))
            r['Total Revenue']     = r['Net Sales'] + r['Membership Income']

            r['COGS']              = r['Total Revenue'] * self.a('COGS % of revenue', year)
            r['Gross Profit']      = r['Total Revenue'] - r['COGS']
            r['SGA']               = r['Total Revenue'] * self.a('SG&A % of revenue', year)
            r['EBITDA']            = r['Gross Profit'] - r['SGA']
            r['EBITDA Margin']     = self._safe_div(r['EBITDA'], r['Total Revenue'])

            # D&A — straight-line on opening PP&E + new CAPEX
            r['CAPEX']             = r['Total Revenue'] * self.a('CAPEX % of revenue', year)
            ppe_life               = self.a('PP&E useful life — Straight-Line (yrs)', year)
            capex_life             = self.a('CAPEX useful life — Straight-Line (yrs)', year)
            r['DA']                = (prev_ppe / ppe_life) + (r['CAPEX'] / capex_life)

            r['EBIT']              = r['EBITDA'] - r['DA']
            r['EBIT Margin']       = self._safe_div(r['EBIT'], r['Total Revenue'])

            # Debt balances
            r['ST Debt End']       = prev_st_debt + self.a('Mandatory ST debt change ($mm)', year)
            r['LT Debt End']       = prev_lt_debt + self.a('Mandatory LT debt change ($mm)', year)

            # Interest expense — average of beginning and ending balances
            r['Interest Exp ST']   = ((prev_st_debt + r['ST Debt End']) / 2) * self.a('Short-term interest rate %', year)
            r['Interest Exp LT']   = ((prev_lt_debt + r['LT Debt End']) / 2) * self.a('Long-term interest rate %', year)
            r['Capital Lease Int'] = self.a('Capital lease interest ($mm)', year)
            r['Interest Income']   = prev_cash * self.a('Cash interest rate %', year)
            r['Net Interest']      = (r['Interest Exp ST'] + r['Interest Exp LT'] +
                                      r['Capital Lease Int'] - r['Interest Income'])

            r['EBT']               = r['EBIT'] - r['Net Interest']
            r['EBT Margin']        = self._safe_div(r['EBT'], r['Total Revenue'])

            # ✅ Tax using ETR — NOT statutory rate (corrected)
            r['Tax Expense']       = r['EBT'] * self.a('Effective Tax Rate (ETR) %', year)
            r['Net Income']        = r['EBT'] - r['Tax Expense']
            r['Net Margin']        = self._safe_div(r['Net Income'], r['Total Revenue'])

            r['Shares']            = self.a('Shares outstanding (mm)', year)
            r['EPS']               = self._safe_div(r['Net Income'], r['Shares'])
            r['DPS']               = self._safe_div(r['Net Income'] * self.a('Dividend payout ratio %', year), r['Shares'])

            # ── CASH FLOW STATEMENT ───────────────────────────────────────
            r['Deferred Tax']      = self.a('Deferred Tax ($mm) — non-cash timing difference', year)

            # OWC balances — calculated from assumption ratios
            r['AR']                = (self.a('Days receivables (DSO)', year) / 365) * r['Total Revenue']
            r['Inventory']         = (self.a('Inventory turnover days (DIO)', year) / 365) * r['COGS']
            r['Prepaid']           = (self.a('Days prepaid', year) / 365) * r['Total Revenue']
            r['AP']                = (self.a('Days payable — AP (DPO, on COGS)', year) / 365) * r['COGS']

            # ✅ Accrued liabilities use Total Revenue — NOT SG&A (corrected)
            r['Accrued Liab']      = (self.a('Days payable — Accrued liabilities', year) / 365) * r['Total Revenue']
            r['Accrued Tax']       = (self.a('Days payable — Accrued tax', year) / 365) * r['Tax Expense']

            # Changes in OWC (negative = cash outflow for rising assets, positive for rising liabilities)
            r['Chg AR']            = -(r['AR'] - prev_ar)
            r['Chg Inv']           = -(r['Inventory'] - prev_inv)
            r['Chg Prepaid']       = -(r['Prepaid'] - prev_prepaid)
            r['Chg AP']            = r['AP'] - prev_ap
            r['Chg Accrued Liab']  = r['Accrued Liab'] - prev_al
            r['Chg Accrued Tax']   = r['Accrued Tax'] - prev_at
            r['Net OWC Change']    = (r['Chg AR'] + r['Chg Inv'] + r['Chg Prepaid'] +
                                      r['Chg AP'] + r['Chg Accrued Liab'] + r['Chg Accrued Tax'])

            r['Other Oper']        = self.a('Other operating activities ($mm)', year)
            r['CFO']               = (r['Net Income'] + r['DA'] + r['Deferred Tax'] +
                                      r['Other Oper'] + r['Net OWC Change'])

            # ✅ CAPEX negative = cash outflow (corrected)
            r['CAPEX CF']          = -r['CAPEX']
            r['Investments']       = -self.a('Investments & acquisitions ($mm)', year)
            r['CFI']               = r['CAPEX CF'] + r['Investments']

            r['ST Debt Chg']       = r['ST Debt End'] - prev_st_debt
            r['LT Debt Chg']       = r['LT Debt End'] - prev_lt_debt

            # ✅ Dividends negative = cash outflow (corrected)
            r['Dividends']         = -(r['Net Income'] * self.a('Dividend payout ratio %', year))
            r['Share Issuance']    = self.a('Share issuance / (buyback) ($mm)', year)
            r['CFF']               = r['ST Debt Chg'] + r['LT Debt Chg'] + r['Dividends'] + r['Share Issuance']

            r['Net Chg Cash']      = r['CFO'] + r['CFI'] + r['CFF']
            r['Ending Cash']       = prev_cash + r['Net Chg Cash']
            r['Free Cash Flow']    = r['CFO'] + r['CAPEX CF']

            # ── BALANCE SHEET ─────────────────────────────────────────────
            r['PPE Net']           = prev_ppe + r['CAPEX'] - r['DA']
            r['Total Curr Assets'] = r['Ending Cash'] + r['AR'] + r['Inventory'] + r['Prepaid']
            r['Goodwill']          = goodwill
            r['Other Assets']      = other_assets
            r['Total Assets']      = r['Total Curr Assets'] + r['PPE Net'] + r['Goodwill'] + r['Other Assets']

            r['Total Curr Liab']   = r['ST Debt End'] + r['AP'] + r['Accrued Liab'] + r['Accrued Tax']
            r['LT Debt']           = r['LT Debt End']
            r['Total Liab']        = r['Total Curr Liab'] + r['LT Debt']

            r['Share Capital']     = share_capital
            r['Retained Earnings'] = prev_retained + r['Net Income'] - abs(r['Dividends'])
            r['Total Equity']      = r['Share Capital'] + r['Retained Earnings']
            r['Total Liab Equity'] = r['Total Liab'] + r['Total Equity']

            # Balance sheet check — must equal 0
            r['BS Check']          = round(r['Total Assets'] - r['Total Liab Equity'], 2)

            results.append(r)

            # ── Update prev values for next year ──────────────────────────
            prev_net_sales  = r['Net Sales']
            prev_membership = r['Membership Income']
            prev_cash       = r['Ending Cash']
            prev_ppe        = r['PPE Net']
            prev_st_debt    = r['ST Debt End']
            prev_lt_debt    = r['LT Debt End']
            prev_retained   = r['Retained Earnings']
            prev_ar         = r['AR']
            prev_inv        = r['Inventory']
            prev_prepaid    = r['Prepaid']
            prev_ap         = r['AP']
            prev_al         = r['Accrued Liab']
            prev_at         = r['Accrued Tax']

        self.projections = pd.DataFrame(results).set_index('Year')
        return self.projections