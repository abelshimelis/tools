import pandas as pd


class FinancialModel:
    def __init__(self, assumptions, income_hist, balance_hist, capital_hist, cashflow_hist):
        self.assumptions   = assumptions
        self.income_hist   = income_hist
        self.balance_hist  = balance_hist
        self.capital_hist  = capital_hist
        self.cashflow_hist = cashflow_hist
        self.proj_years    = ['FY2025E', 'FY2026E', 'FY2027E', 'FY2028E', 'FY2029E']
        self.projections   = {}

    def _a(self, label, year, default=0.0):
        """Safely get an assumption value."""
        try:
            return float(self.assumptions.get(label, {}).get(year, default))
        except (TypeError, ValueError):
            return default

    def _h(self, data, label, year='FY2024A', default=0.0):
        """Safely get a historical value."""
        try:
            return float(data.get(label, {}).get(year, default))
        except (TypeError, ValueError):
            return default

    def project(self):
        self.projections['income']   = self._project_income()
        self.projections['balance']  = self._project_balance()
        self.projections['capital']  = self._project_capital()
        self.projections['cashflow'] = self._project_cashflow()
        return self.projections

    # ── INCOME STATEMENT ──────────────────────────────────────────────
    def _project_income(self):
        rows = []
        prev_fee = self._h(self.income_hist, 'Fee & Commission Income', self.base_year)
        prev_trading = self._h(self.income_hist, 'Trading Income', self.base_year)
        prev_other = self._h(self.income_hist, 'Other Operating Income', self.base_year)
        prev_gross_loans = self._h(self.balance_hist, 'Gross Loans & Advances', self.base_year)
        prev_total_deposits = self._h(self.balance_hist, 'Total Customer Deposits', self.base_year)
        prev_casa = self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year)
        prev_term = self._h(self.balance_hist, 'Term / Fixed Deposits', self.base_year)
        prev_interbank_placements = self._h(self.balance_hist, 'Interbank Placements', self.base_year)
        prev_investment_securities = self._h(self.balance_hist, 'Investment Securities', self.base_year)
        prev_interbank_borrowings = self._h(self.balance_hist, 'Interbank Borrowings', self.base_year)
        prev_debt_securities = self._h(self.balance_hist, 'Debt Securities Issued', self.base_year)
        prev_subordinated = self._h(self.balance_hist, 'Subordinated Debt', self.base_year)
        prev_total_assets = self._h(self.balance_hist, 'TOTAL ASSETS', self.base_year)
        prev_equity = self._h(self.balance_hist, 'TOTAL EQUITY', self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth = self._a('Loan Growth (YoY %)', year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['Gross Loans'] = prev_gross_loans * (1 + loan_growth)
            r['Total Deposits'] = prev_total_deposits * (1 + deposit_growth)
            casa_share = prev_casa / prev_total_deposits if prev_total_deposits else 0.55
            r['CASA Deposits'] = r['Total Deposits'] * casa_share
            r['Term Deposits'] = r['Total Deposits'] - r['CASA Deposits']

            r['Interbank Placements'] = prev_interbank_placements * (1 + deposit_growth)
            r['Investment Securities'] = prev_investment_securities * (1 + deposit_growth)
            r['Interbank Borrowings'] = prev_interbank_borrowings * (1 + deposit_growth)
            r['Debt Securities'] = prev_debt_securities * (1 + deposit_growth)
            r['Subordinated Debt'] = prev_subordinated

            r['Interest Income'] = (
                r['Gross Loans'] * self._a('Avg. Loan Yield (%)', year) +
                r['Investment Securities'] * self._a('Avg. Investment Yield (%)', year) +
                r['Interbank Placements'] * self._a('Avg. Investment Yield (%)', year)
            )
            r['Interest Expense'] = (
                (r['CASA Deposits'] + r['Term Deposits']) * self._a('Avg. Cost of Deposits (%)', year) +
                r['Interbank Borrowings'] * self._a('Avg. Cost of Borrowings (%)', year) +
                r['Debt Securities'] * self._a('Avg. Cost of Borrowings (%)', year) +
                r['Subordinated Debt'] * self._a('Avg. Cost of Borrowings (%)', year)
            )
            r['Net Interest Income'] = r['Interest Income'] - r['Interest Expense']
            r['NIM (%)'] = self._a('Net Interest Margin (NIM)', year)

            r['Fee & Commission'] = prev_fee * (1 + self._a('Fee Income Growth (%)', year))
            r['Trading Income'] = self._a('Trading Income ($mm)', year)
            r['Other Non-Interest'] = self._a('Other Income ($mm)', year)
            r['Total Non-Interest'] = r['Fee & Commission'] + r['Trading Income'] + r['Other Non-Interest']

            r['Total Operating Income'] = r['Net Interest Income'] + r['Total Non-Interest']
            r['Cost to Income Ratio'] = self._a('Cost-to-Income Ratio (%)', year)
            r['Total Opex'] = r['Total Operating Income'] * r['Cost to Income Ratio']
            r['Staff Costs'] = r['Total Opex'] * 0.55
            r['Tech & Infra'] = r['Total Opex'] * 0.25
            r['Occupancy & Other'] = r['Total Opex'] * 0.20
            r['Depreciation & Amortisation'] = self._a('D&A ($mm)', year)

            r['Pre-Provision Profit'] = r['Total Operating Income'] - r['Total Opex']
            r['Loan Loss Provisions'] = r['Gross Loans'] * self._a('Loan Loss Provision / Loans', year)
            r['Other Impairments'] = 0.0
            r['Total Provisions'] = r['Loan Loss Provisions'] + r['Other Impairments']

            r['Pre-Tax Profit'] = r['Pre-Provision Profit'] - r['Total Provisions']
            r['Tax Expense'] = r['Pre-Tax Profit'] * self._a('Effective Tax Rate (%)', year)
            r['Net Profit After Tax'] = r['Pre-Tax Profit'] - r['Tax Expense']

            r['Cost-to-Income (%)'] = self._a('Cost-to-Income Ratio (%)', year)
            r['LLP / Loans (%)'] = self._a('Loan Loss Provision / Loans', year)
            r['ROE'] = r['Net Profit After Tax'] / prev_equity if prev_equity else 0.0
            r['ROA'] = r['Net Profit After Tax'] / prev_total_assets if prev_total_assets else 0.0
            r['NPL Ratio'] = self._a('NPL Ratio (%)', year)

            rows.append(r)

            prev_gross_loans = r['Gross Loans']
            prev_total_deposits = r['Total Deposits']
            prev_casa = r['CASA Deposits']
            prev_term = r['Term Deposits']
            prev_interbank_placements = r['Interbank Placements']
            prev_investment_securities = r['Investment Securities']
            prev_interbank_borrowings = r['Interbank Borrowings']
            prev_debt_securities = r['Debt Securities']
            prev_fee = r['Fee & Commission']
            prev_trading = r['Trading Income']
            prev_other = r['Other Non-Interest']

        return pd.DataFrame(rows).set_index('Year')

    # ── BALANCE SHEET ─────────────────────────────────────────────────
    def _project_balance(self):
        rows = []

        prev_cash = self._h(self.balance_hist, 'Cash & Central Bank Reserves', self.base_year)
        prev_interbank_placements = self._h(self.balance_hist, 'Interbank Placements', self.base_year)
        prev_investment_securities = self._h(self.balance_hist, 'Investment Securities', self.base_year)
        prev_gross_loans = self._h(self.balance_hist, 'Gross Loans & Advances', self.base_year)
        prev_loan_loss_reserves = self._h(self.balance_hist, 'Less: Loan Loss Reserves', self.base_year)
        prev_fixed_assets = self._h(self.balance_hist, 'Fixed Assets & ROU', self.base_year)
        prev_goodwill = self._h(self.balance_hist, 'Goodwill & Intangibles', self.base_year)
        prev_other_assets = self._h(self.balance_hist, 'Other Assets (Deferred Tax, Derivatives)', self.base_year)
        prev_total_deposits = self._h(self.balance_hist, 'Total Customer Deposits', self.base_year)
        prev_casa = self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year)
        prev_term = self._h(self.balance_hist, 'Term / Fixed Deposits', self.base_year)
        prev_interbank_borrowings = self._h(self.balance_hist, 'Interbank Borrowings', self.base_year)
        prev_debt_securities = self._h(self.balance_hist, 'Debt Securities Issued', self.base_year)
        prev_subordinated = self._h(self.balance_hist, 'Subordinated Debt', self.base_year)
        prev_other_liabilities = self._h(self.balance_hist, 'Other Liabilities (Derivatives, Accruals)', self.base_year)
        prev_retained = self._h(self.balance_hist, 'Retained Earnings', self.base_year)
        prev_oci = self._h(self.balance_hist, 'Other Comprehensive Income (OCI)', self.base_year)
        prev_at1 = self._h(self.balance_hist, 'AT1 Capital Instruments', self.base_year)
        prev_share_capital = self._h(self.balance_hist, 'Share Capital & Premium', self.base_year)

        cash_delta = prev_cash - self._h(self.balance_hist, 'Cash & Central Bank Reserves', 'FY2023A', prev_cash)
        fixed_assets_delta = prev_fixed_assets - self._h(self.balance_hist, 'Fixed Assets & ROU', 'FY2023A', prev_fixed_assets)
        goodwill_delta = prev_goodwill - self._h(self.balance_hist, 'Goodwill & Intangibles', 'FY2023A', prev_goodwill)
        other_assets_delta = prev_other_assets - self._h(self.balance_hist, 'Other Assets (Deferred Tax, Derivatives)', 'FY2023A', prev_other_assets)
        interbank_delta = prev_interbank_borrowings - self._h(self.balance_hist, 'Interbank Borrowings', 'FY2023A', prev_interbank_borrowings)
        debt_sec_delta = prev_debt_securities - self._h(self.balance_hist, 'Debt Securities Issued', 'FY2023A', prev_debt_securities)
        other_liabilities_delta = prev_other_liabilities - self._h(self.balance_hist, 'Other Liabilities (Derivatives, Accruals)', 'FY2023A', prev_other_liabilities)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth = self._a('Loan Growth (YoY %)', year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['Gross Loans'] = prev_gross_loans * (1 + loan_growth)
            r['Loan Loss Reserves'] = -(r['Gross Loans'] * self._a('NPL Ratio (%)', year)
                                        * self._a('Coverage Ratio (%)', year))
            r['Net Loans'] = r['Gross Loans'] + r['Loan Loss Reserves']

            r['Total Deposits'] = prev_total_deposits * (1 + deposit_growth)
            casa_share = prev_casa / prev_total_deposits if prev_total_deposits else 0.55
            r['CASA Deposits'] = r['Total Deposits'] * casa_share
            r['Term Deposits'] = r['Total Deposits'] - r['CASA Deposits']

            r['Interbank Borrowings'] = prev_interbank_borrowings + interbank_delta
            r['Debt Securities'] = prev_debt_securities + debt_sec_delta
            r['Subordinated Debt'] = prev_subordinated

            r['Interbank Placements'] = prev_interbank_placements * (1 + deposit_growth)
            r['Investment Securities'] = prev_investment_securities * (1 + deposit_growth)
            r['Fixed Assets'] = prev_fixed_assets + fixed_assets_delta
            r['Goodwill'] = prev_goodwill + goodwill_delta
            r['Other Assets'] = prev_other_assets + other_assets_delta
            r['Cash & Reserves'] = prev_cash + cash_delta

            r['Total Assets'] = (r['Cash & Reserves'] + r['Interbank Placements'] +
                                 r['Investment Securities'] + r['Net Loans'] +
                                 r['Fixed Assets'] + r['Goodwill'] + r['Other Assets'])

            net_profit = self.projections['income'].loc[year, 'Net Profit After Tax'] if 'income' in self.projections else 0
            dividend_payout = self._a('Dividend Payout Ratio (%)', year)
            r['Share Capital'] = prev_share_capital
            r['Retained Earnings'] = prev_retained + net_profit - (net_profit * dividend_payout)
            r['OCI Reserves'] = prev_oci
            r['AT1 Capital'] = prev_at1
            r['Total Equity'] = (r['Share Capital'] + r['Retained Earnings'] +
                                 r['OCI Reserves'] + r['AT1 Capital'])

            # Balance to the residual liability line
            r['Other Liabilities'] = (r['Total Assets'] -
                                      (r['CASA Deposits'] + r['Term Deposits'] +
                                       r['Interbank Borrowings'] + r['Debt Securities'] +
                                       r['Subordinated Debt'] + r['Total Equity']))
            r['Total Liabilities'] = (r['CASA Deposits'] + r['Term Deposits'] +
                                      r['Interbank Borrowings'] + r['Debt Securities'] +
                                      r['Subordinated Debt'] + r['Other Liabilities'])
            r['Total Liab & Equity'] = r['Total Liabilities'] + r['Total Equity']
            r['BS Check'] = round(r['Total Assets'] - r['Total Liab & Equity'], 2)

            prev_cash = r['Cash & Reserves']
            prev_interbank_placements = r['Interbank Placements']
            prev_investment_securities = r['Investment Securities']
            prev_gross_loans = r['Gross Loans']
            prev_loan_loss_reserves = r['Loan Loss Reserves']
            prev_fixed_assets = r['Fixed Assets']
            prev_goodwill = r['Goodwill']
            prev_other_assets = r['Other Assets']
            prev_total_deposits = r['Total Deposits']
            prev_casa = r['CASA Deposits']
            prev_term = r['Term Deposits']
            prev_interbank_borrowings = r['Interbank Borrowings']
            prev_debt_securities = r['Debt Securities']
            prev_retained = r['Retained Earnings']

            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')

    # ── CAPITAL ADEQUACY ──────────────────────────────────────────────
    def _project_capital(self):
        rows = []

        prev_cet1 = self._h(self.capital_hist, 'Common Equity Tier 1 (CET1)', self.base_year)
        prev_at1 = self._h(self.capital_hist, 'Additional Tier 1 (AT1)', self.base_year)
        prev_tier2 = self._h(self.capital_hist, 'Tier 2 Capital (Sub Debt + Reserves)', self.base_year)
        prev_credit_rwa = self._h(self.capital_hist, 'Credit Risk RWA ($mm)', self.base_year)
        prev_market_rwa = self._h(self.capital_hist, 'Market Risk RWA ($mm)', self.base_year)
        prev_operational_rwa = self._h(self.capital_hist, 'Operational Risk RWA ($mm)', self.base_year)
        prev_total_deposits = self._h(self.balance_hist, 'Total Customer Deposits', self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth = self._a('Loan Growth (YoY %)', year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['CET1 Capital'] = self._h(self.capital_hist, 'Common Equity Tier 1 (CET1)', year, prev_cet1)
            r['AT1 Capital'] = self._h(self.capital_hist, 'Additional Tier 1 (AT1)', year, prev_at1)
            r['Tier1 Capital'] = r['CET1 Capital'] + r['AT1 Capital']
            r['Tier2 Capital'] = self._h(self.capital_hist, 'Tier 2 Capital (Sub Debt + Reserves)', year, prev_tier2)
            r['Total Reg Capital'] = r['Tier1 Capital'] + r['Tier2 Capital']

            if year == 'FY2025E':
                r['Credit Risk RWA'] = self._h(self.capital_hist, 'Credit Risk RWA ($mm)', year, prev_credit_rwa * (1 + loan_growth))
                r['Market Risk RWA'] = self._h(self.capital_hist, 'Market Risk RWA ($mm)', year, prev_market_rwa * (1 + deposit_growth))
                r['Operational Risk RWA'] = self._h(self.capital_hist, 'Operational Risk RWA ($mm)', year, prev_operational_rwa * 1.03)
            else:
                prev_rwa_growth = self._a('Loan Growth (YoY %)', year)
                r['Credit Risk RWA'] = prev_credit_rwa * (1 + prev_rwa_growth)
                r['Market Risk RWA'] = prev_market_rwa * (1 + deposit_growth)
                r['Operational Risk RWA'] = prev_operational_rwa * 1.03

            r['Total RWA'] = r['Credit Risk RWA'] + r['Market Risk RWA'] + r['Operational Risk RWA']
            if r['Total RWA'] > 0:
                r['CET1 Ratio'] = r['CET1 Capital'] / r['Total RWA']
                r['Tier1 Ratio'] = r['Tier1 Capital'] / r['Total RWA']
                r['CAR'] = r['Total Reg Capital'] / r['Total RWA']
            else:
                r['CET1 Ratio'] = r['Tier1 Ratio'] = r['CAR'] = 0.0

            total_loans = self._h(self.balance_hist, 'Gross Loans & Advances', year, self._h(self.balance_hist, 'Gross Loans & Advances', self.base_year))
            total_deposits = self._h(self.balance_hist, 'Total Customer Deposits', year, prev_total_deposits * (1 + deposit_growth))
            r['LDR'] = total_loans / total_deposits if total_deposits else 0.0

            prev_credit_rwa = r['Credit Risk RWA']
            prev_market_rwa = r['Market Risk RWA']
            prev_operational_rwa = r['Operational Risk RWA']
            prev_cet1 = r['CET1 Capital']
            prev_at1 = r['AT1 Capital']
            prev_tier2 = r['Tier2 Capital']
            prev_total_deposits = total_deposits

            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')

    # ── CASH FLOW STATEMENT ───────────────────────────────────────────
    def _project_cashflow(self):
        rows = []
        prev_loans = self._h(self.balance_hist, 'Gross Loans & Advances', self.base_year)
        prev_deposits = (self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year) +
                         self._h(self.balance_hist, 'Term / Fixed Deposits', self.base_year))
        prev_cash = self._h(self.cashflow_hist, 'Closing Cash Balance', self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            net_profit = self.projections['income'].loc[year, 'Net Profit After Tax'] if 'income' in self.projections else 0
            provisions = self.projections['income'].loc[year, 'Loan Loss Provisions'] if 'income' in self.projections else 0

            r['Opening Cash Balance'] = prev_cash
            r['Net Income'] = net_profit
            r['Depreciation & Amortisation'] = self._a('D&A ($mm)', year)
            r['Add: Loan Loss Provisions'] = provisions
            r['Change in Loans (net)'] = -(self.projections['balance'].loc[year, 'Gross Loans'] - prev_loans)
            r['Change in Deposits (net)'] = (self.projections['balance'].loc[year, 'Total Deposits'] - prev_deposits)
            r['Change in Other Working Capital'] = self._h(self.cashflow_hist, 'Change in Other Working Capital', year, self._h(self.cashflow_hist, 'Change in Other Working Capital', self.base_year, -105))
            r['Net Cash from Operating Activities'] = (r['Net Income'] + r['Depreciation & Amortisation'] +
                                                      r['Add: Loan Loss Provisions'] + r['Change in Loans (net)'] +
                                                      r['Change in Deposits (net)'] + r['Change in Other Working Capital'])

            r['Purchase of Investment Securities'] = self._h(self.cashflow_hist, 'Purchase of Investment Securities', year, self._h(self.cashflow_hist, 'Purchase of Investment Securities', self.base_year, -500))
            r['Proceeds from Securities Maturity / Sale'] = self._h(self.cashflow_hist, 'Proceeds from Securities Maturity / Sale', year, self._h(self.cashflow_hist, 'Proceeds from Securities Maturity / Sale', self.base_year, 400))
            r['Capital Expenditure (Capex)'] = self._h(self.cashflow_hist, 'Capital Expenditure (Capex)', year, self._h(self.cashflow_hist, 'Capital Expenditure (Capex)', self.base_year, -105))
            r['Acquisitions / Disposals (net)'] = self._h(self.cashflow_hist, 'Acquisitions / Disposals (net)', year, self._h(self.cashflow_hist, 'Acquisitions / Disposals (net)', self.base_year, 0))
            r['Net Cash from Investing Activities'] = (r['Purchase of Investment Securities'] +
                                                      r['Proceeds from Securities Maturity / Sale'] +
                                                      r['Capital Expenditure (Capex)'] +
                                                      r['Acquisitions / Disposals (net)'])

            dividends = -(net_profit * self._a('Dividend Payout Ratio (%)', year))
            r['Dividends Paid'] = dividends
            r['Issuance of Debt Securities'] = self._h(self.cashflow_hist, 'Issuance of Debt Securities', year, self._h(self.cashflow_hist, 'Issuance of Debt Securities', self.base_year, 150))
            r['Repayment of Borrowings'] = self._h(self.cashflow_hist, 'Repayment of Borrowings', year, self._h(self.cashflow_hist, 'Repayment of Borrowings', self.base_year, -200))
            r['Share Capital Issued / (Bought Back)'] = self._h(self.cashflow_hist, 'Share Capital Issued / (Bought Back)', year, self._h(self.cashflow_hist, 'Share Capital Issued / (Bought Back)', self.base_year, 0))
            r['Net Cash from Financing Activities'] = (r['Issuance of Debt Securities'] +
                                                      r['Repayment of Borrowings'] +
                                                      r['Share Capital Issued / (Bought Back)'] +
                                                      r['Dividends Paid'])

            r['NET CHANGE IN CASH'] = (r['Net Cash from Operating Activities'] +
                                      r['Net Cash from Investing Activities'] +
                                      r['Net Cash from Financing Activities'])
            r['Closing Cash Balance'] = r['Opening Cash Balance'] + r['NET CHANGE IN CASH']

            prev_loans = self.projections['balance'].loc[year, 'Gross Loans']
            prev_deposits = self.projections['balance'].loc[year, 'Total Deposits']
            prev_cash = r['Closing Cash Balance']
            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')
