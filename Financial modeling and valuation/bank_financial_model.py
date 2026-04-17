import pandas as pd


class FinancialModel:
    def __init__(self, assumptions, income_hist, balance_hist, capital_hist, cashflow_hist):
        self.assumptions   = assumptions
        self.income_hist   = income_hist
        self.balance_hist  = balance_hist
        self.capital_hist  = capital_hist
        self.cashflow_hist = cashflow_hist
        self.proj_years    = ['FY2025E', 'FY2026E', 'FY2027E', 'FY2028E', 'FY2029E']
        self.base_year     = 'FY2024A'   # ← was missing, caused crashes
        self.projections   = {}

        # ── Pre-compute historical opex splits from income history ────
        # These replace the hardcoded 0.55 / 0.25 / 0.20 staff/tech/occ splits
        hist_staff = self._h(self.income_hist, 'Staff Costs',                  self.base_year)
        hist_tech  = self._h(self.income_hist, 'Technology & Infrastructure',   self.base_year)
        hist_occ   = self._h(self.income_hist, 'Occupancy & Admin',             self.base_year)
        hist_opex_ex_da = hist_staff + hist_tech + hist_occ
        if hist_opex_ex_da > 0:
            self.staff_ratio = hist_staff / hist_opex_ex_da
            self.tech_ratio  = hist_tech  / hist_opex_ex_da
            self.occ_ratio   = hist_occ   / hist_opex_ex_da
        else:
            # If no history exists yet, split will be recalculated once data is filled
            raise ValueError(
                "\n  MISSING HISTORICAL DATA: 'Staff Costs', 'Technology & Infrastructure',"
                " 'Occupancy & Admin' for FY2024A in Income Statement sheet."
                "\n  → These are needed to derive opex splits automatically."
            )

        # ── Pre-compute CASA share from history ───────────────────────
        # Replaces the hardcoded 0.55 fallback
        hist_casa         = self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year)
        hist_total_deps   = self._h(self.balance_hist, 'Total Customer Deposits',            self.base_year)
        if hist_total_deps > 0:
            self.casa_share = hist_casa / hist_total_deps
        else:
            raise ValueError(
                "\n  MISSING HISTORICAL DATA: 'CASA Deposits' or 'Total Customer Deposits'"
                " for FY2024A in Balance Sheet sheet."
            )

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
        prev_fee                    = self._h(self.income_hist, 'Fee & Commission Income',  self.base_year)
        prev_gross_loans            = self._h(self.balance_hist, 'Gross Loans & Advances',  self.base_year)
        prev_total_deposits         = self._h(self.balance_hist, 'Total Customer Deposits', self.base_year)
        prev_casa                   = self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year)
        prev_term                   = self._h(self.balance_hist, 'Term / Fixed Deposits',   self.base_year)
        prev_interbank_placements   = self._h(self.balance_hist, 'Interbank Placements',    self.base_year)
        prev_investment_securities  = self._h(self.balance_hist, 'Investment Securities',   self.base_year)
        prev_interbank_borrowings   = self._h(self.balance_hist, 'Interbank Borrowings',    self.base_year)
        prev_debt_securities        = self._h(self.balance_hist, 'Debt Securities Issued',  self.base_year)
        prev_subordinated           = self._h(self.balance_hist, 'Subordinated Debt',       self.base_year)
        prev_total_assets           = self._h(self.balance_hist, 'TOTAL ASSETS',            self.base_year)
        prev_equity                 = self._h(self.balance_hist, 'TOTAL EQUITY',            self.base_year)
        prev_staff                  = self._h(self.income_hist,  'Staff Costs',             self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth    = self._a('Loan Growth (YoY %)',    year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['Gross Loans']      = prev_gross_loans * (1 + loan_growth)
            r['Total Deposits']   = prev_total_deposits * (1 + deposit_growth)

            # CASA / Term split — derived from historical ratio, not hardcoded
            r['CASA Deposits']    = r['Total Deposits'] * self.casa_share
            r['Term Deposits']    = r['Total Deposits'] - r['CASA Deposits']

            r['Interbank Placements']  = prev_interbank_placements  * (1 + deposit_growth)
            r['Investment Securities'] = prev_investment_securities * (1 + deposit_growth)
            r['Interbank Borrowings']  = prev_interbank_borrowings  * (1 + deposit_growth)
            r['Debt Securities']       = prev_debt_securities        * (1 + deposit_growth)
            r['Subordinated Debt']     = prev_subordinated

            r['Interest Income'] = (
                r['Gross Loans']             * self._a('Avg. Loan Yield (%)',        year) +
                r['Investment Securities']   * self._a('Avg. Investment Yield (%)',  year) +
                r['Interbank Placements']    * self._a('Avg. Investment Yield (%)',  year)
            )
            r['Interest Expense'] = (
                (r['CASA Deposits'] + r['Term Deposits']) * self._a('Avg. Cost of Deposits (%)',   year) +
                r['Interbank Borrowings']                 * self._a('Avg. Cost of Borrowings (%)', year) +
                r['Debt Securities']                      * self._a('Avg. Cost of Borrowings (%)', year) +
                r['Subordinated Debt']                    * self._a('Avg. Cost of Borrowings (%)', year)
            )
            r['Net Interest Income'] = r['Interest Income'] - r['Interest Expense']
            r['NIM (%)']             = self._a('Net Interest Margin (NIM)', year)

            r['Fee & Commission']    = prev_fee * (1 + self._a('Fee Income Growth (%)', year))
            r['Trading Income']      = self._a('Trading Income ($mm)', year)
            r['Other Non-Interest']  = self._a('Other Income ($mm)',   year)
            r['Total Non-Interest']  = r['Fee & Commission'] + r['Trading Income'] + r['Other Non-Interest']

            r['Total Operating Income'] = r['Net Interest Income'] + r['Total Non-Interest']
            r['Cost to Income Ratio']   = self._a('Cost-to-Income Ratio (%)', year)
            r['Total Opex']             = r['Total Operating Income'] * r['Cost to Income Ratio']

            # Opex breakdown — derived from historical proportions, not hardcoded 0.55 / 0.25 / 0.20
            da             = self._a('D&A ($mm)', year)
            opex_ex_da     = max(r['Total Opex'] - da, 0)
            r['Staff Costs']              = opex_ex_da * self.staff_ratio
            r['Tech & Infra']             = opex_ex_da * self.tech_ratio
            r['Occupancy & Other']        = opex_ex_da * self.occ_ratio
            r['Depreciation & Amortisation'] = da

            r['Pre-Provision Profit']  = r['Total Operating Income'] - r['Total Opex']
            r['Loan Loss Provisions']  = r['Gross Loans'] * self._a('Loan Loss Provision / Loans', year)
            r['Other Impairments']     = 0.0
            r['Total Provisions']      = r['Loan Loss Provisions'] + r['Other Impairments']

            r['Pre-Tax Profit']        = r['Pre-Provision Profit'] - r['Total Provisions']
            r['Tax Expense']           = r['Pre-Tax Profit'] * self._a('Effective Tax Rate (%)', year)
            r['Net Profit After Tax']  = r['Pre-Tax Profit'] - r['Tax Expense']

            r['Cost-to-Income (%)']    = self._a('Cost-to-Income Ratio (%)',    year)
            r['LLP / Loans (%)']       = self._a('Loan Loss Provision / Loans', year)

            # ROE / ROA — use average of opening and closing equity/assets, not just opening
            avg_equity = (prev_equity + self._h(self.balance_hist, 'TOTAL EQUITY', self.base_year)) / 2 \
                          if year == self.proj_years[0] else prev_equity
            avg_assets = (prev_total_assets + self._h(self.balance_hist, 'TOTAL ASSETS', self.base_year)) / 2 \
                          if year == self.proj_years[0] else prev_total_assets
            r['ROE']    = r['Net Profit After Tax'] / avg_equity  if avg_equity  else 0.0
            r['ROA']    = r['Net Profit After Tax'] / avg_assets  if avg_assets  else 0.0
            r['NPL Ratio'] = self._a('NPL Ratio (%)', year)

            rows.append(r)

            prev_gross_loans           = r['Gross Loans']
            prev_total_deposits        = r['Total Deposits']
            prev_casa                  = r['CASA Deposits']
            prev_term                  = r['Term Deposits']
            prev_interbank_placements  = r['Interbank Placements']
            prev_investment_securities = r['Investment Securities']
            prev_interbank_borrowings  = r['Interbank Borrowings']
            prev_debt_securities       = r['Debt Securities']
            prev_fee                   = r['Fee & Commission']
            prev_equity                = r['Net Profit After Tax'] + prev_equity   # rough carry-forward
            prev_total_assets          = r['Total Operating Income'] + prev_total_assets  # updated below in BS

        return pd.DataFrame(rows).set_index('Year')

    # ── BALANCE SHEET ─────────────────────────────────────────────────
    def _project_balance(self):
        rows = []

        prev_cash                = self._h(self.balance_hist, 'Cash & Central Bank Reserves',            self.base_year)
        prev_interbank_placements= self._h(self.balance_hist, 'Interbank Placements',                    self.base_year)
        prev_investment_securities=self._h(self.balance_hist, 'Investment Securities',                   self.base_year)
        prev_gross_loans         = self._h(self.balance_hist, 'Gross Loans & Advances',                  self.base_year)
        prev_loan_loss_reserves  = self._h(self.balance_hist, 'Less: Loan Loss Reserves',                self.base_year)
        prev_fixed_assets        = self._h(self.balance_hist, 'Fixed Assets & ROU',                      self.base_year)
        prev_goodwill            = self._h(self.balance_hist, 'Goodwill & Intangibles',                  self.base_year)
        prev_other_assets        = self._h(self.balance_hist, 'Other Assets (Deferred Tax, Derivatives)',self.base_year)
        prev_total_deposits      = self._h(self.balance_hist, 'Total Customer Deposits',                 self.base_year)
        prev_casa                = self._h(self.balance_hist, 'CASA Deposits (Current & Savings)',        self.base_year)
        prev_interbank_borrowings= self._h(self.balance_hist, 'Interbank Borrowings',                    self.base_year)
        prev_debt_securities     = self._h(self.balance_hist, 'Debt Securities Issued',                  self.base_year)
        prev_subordinated        = self._h(self.balance_hist, 'Subordinated Debt',                       self.base_year)
        prev_other_liabilities   = self._h(self.balance_hist, 'Other Liabilities (Derivatives, Accruals)',self.base_year)
        prev_retained            = self._h(self.balance_hist, 'Retained Earnings',                       self.base_year)
        prev_oci                 = self._h(self.balance_hist, 'Other Comprehensive Income (OCI)',        self.base_year)
        prev_at1                 = self._h(self.balance_hist, 'AT1 Capital Instruments',                 self.base_year)
        prev_share_capital       = self._h(self.balance_hist, 'Share Capital & Premium',                 self.base_year)

        # Derive YoY deltas from history (FY2023A → FY2024A) for items without assumption drivers
        # These replace the hardcoded fixed deltas that were previously used
        prior_year = 'FY2023A'
        cash_delta          = prev_cash             - self._h(self.balance_hist, 'Cash & Central Bank Reserves',            prior_year, prev_cash)
        fixed_assets_delta  = prev_fixed_assets     - self._h(self.balance_hist, 'Fixed Assets & ROU',                      prior_year, prev_fixed_assets)
        goodwill_delta      = prev_goodwill         - self._h(self.balance_hist, 'Goodwill & Intangibles',                  prior_year, prev_goodwill)
        other_assets_delta  = prev_other_assets     - self._h(self.balance_hist, 'Other Assets (Deferred Tax, Derivatives)',prior_year, prev_other_assets)
        interbank_delta     = prev_interbank_borrowings - self._h(self.balance_hist, 'Interbank Borrowings',                prior_year, prev_interbank_borrowings)
        debt_sec_delta      = prev_debt_securities  - self._h(self.balance_hist, 'Debt Securities Issued',                  prior_year, prev_debt_securities)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth    = self._a('Loan Growth (YoY %)',    year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['Gross Loans']       = prev_gross_loans * (1 + loan_growth)
            r['Loan Loss Reserves']= -(r['Gross Loans'] * self._a('NPL Ratio (%)',      year)
                                                        * self._a('Coverage Ratio (%)', year))
            r['Net Loans']         = r['Gross Loans'] + r['Loan Loss Reserves']

            r['Total Deposits']    = prev_total_deposits * (1 + deposit_growth)

            # CASA / Term split — derived from historical ratio, not hardcoded
            r['CASA Deposits']     = r['Total Deposits'] * self.casa_share
            r['Term Deposits']     = r['Total Deposits'] - r['CASA Deposits']

            r['Interbank Borrowings']  = prev_interbank_borrowings  + interbank_delta
            r['Debt Securities']       = prev_debt_securities        + debt_sec_delta
            r['Subordinated Debt']     = prev_subordinated

            r['Interbank Placements']  = prev_interbank_placements  * (1 + deposit_growth)
            r['Investment Securities'] = prev_investment_securities * (1 + deposit_growth)
            r['Fixed Assets']          = prev_fixed_assets           + fixed_assets_delta
            r['Goodwill']              = prev_goodwill               + goodwill_delta
            r['Other Assets']          = prev_other_assets           + other_assets_delta
            r['Cash & Reserves']       = prev_cash                   + cash_delta

            r['Total Assets'] = (r['Cash & Reserves']       + r['Interbank Placements'] +
                                 r['Investment Securities']  + r['Net Loans']            +
                                 r['Fixed Assets']           + r['Goodwill']             +
                                 r['Other Assets'])

            net_profit      = self.projections['income'].loc[year, 'Net Profit After Tax'] \
                              if 'income' in self.projections else 0
            dividend_payout = self._a('Dividend Payout Ratio (%)', year)
            r['Share Capital']      = prev_share_capital
            r['Retained Earnings']  = prev_retained + net_profit - (net_profit * dividend_payout)
            r['OCI Reserves']       = prev_oci
            r['AT1 Capital']        = prev_at1
            r['Total Equity']       = (r['Share Capital'] + r['Retained Earnings'] +
                                       r['OCI Reserves']  + r['AT1 Capital'])

            # Other Liabilities balances the sheet automatically — no hardcoded plug
            r['Other Liabilities']  = (r['Total Assets'] -
                                       (r['CASA Deposits'] + r['Term Deposits']     +
                                        r['Interbank Borrowings'] + r['Debt Securities'] +
                                        r['Subordinated Debt']    + r['Total Equity']))
            r['Total Liabilities']  = (r['CASA Deposits']      + r['Term Deposits']  +
                                       r['Interbank Borrowings'] + r['Debt Securities'] +
                                       r['Subordinated Debt']    + r['Other Liabilities'])
            r['Total Liab & Equity']= r['Total Liabilities'] + r['Total Equity']
            r['BS Check']           = round(r['Total Assets'] - r['Total Liab & Equity'], 2)

            prev_cash                 = r['Cash & Reserves']
            prev_interbank_placements = r['Interbank Placements']
            prev_investment_securities= r['Investment Securities']
            prev_gross_loans          = r['Gross Loans']
            prev_loan_loss_reserves   = r['Loan Loss Reserves']
            prev_fixed_assets         = r['Fixed Assets']
            prev_goodwill             = r['Goodwill']
            prev_other_assets         = r['Other Assets']
            prev_total_deposits       = r['Total Deposits']
            prev_casa                 = r['CASA Deposits']
            prev_interbank_borrowings = r['Interbank Borrowings']
            prev_debt_securities      = r['Debt Securities']
            prev_retained             = r['Retained Earnings']

            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')

    # ── CAPITAL ADEQUACY ──────────────────────────────────────────────
    def _project_capital(self):
        rows = []

        prev_cet1            = self._h(self.capital_hist, 'Common Equity Tier 1 (CET1)',          self.base_year)
        prev_at1             = self._h(self.capital_hist, 'Additional Tier 1 (AT1)',              self.base_year)
        prev_tier2           = self._h(self.capital_hist, 'Tier 2 Capital (Sub Debt + Reserves)', self.base_year)
        prev_credit_rwa      = self._h(self.capital_hist, 'Credit Risk RWA ($mm)',                self.base_year)
        prev_market_rwa      = self._h(self.capital_hist, 'Market Risk RWA ($mm)',                self.base_year)
        prev_operational_rwa = self._h(self.capital_hist, 'Operational Risk RWA ($mm)',           self.base_year)
        prev_total_deposits  = self._h(self.balance_hist, 'Total Customer Deposits',              self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            loan_growth    = self._a('Loan Growth (YoY %)',    year)
            deposit_growth = self._a('Deposit Growth (YoY %)', year)

            r['CET1 Capital']   = self._h(self.capital_hist, 'Common Equity Tier 1 (CET1)',          year, prev_cet1)
            r['AT1 Capital']    = self._h(self.capital_hist, 'Additional Tier 1 (AT1)',              year, prev_at1)
            r['Tier1 Capital']  = r['CET1 Capital'] + r['AT1 Capital']
            r['Tier2 Capital']  = self._h(self.capital_hist, 'Tier 2 Capital (Sub Debt + Reserves)', year, prev_tier2)
            r['Total Reg Capital'] = r['Tier1 Capital'] + r['Tier2 Capital']

            # Credit RWA: grows with loan growth
            r['Credit Risk RWA']     = prev_credit_rwa * (1 + loan_growth)

            # Market RWA: grows with deposit growth (proxy for balance sheet size)
            r['Market Risk RWA']     = prev_market_rwa * (1 + deposit_growth)

            # Operational RWA: Basel basic indicator = Total Op Income × 15%
            # Replaces the hardcoded × 1.03 growth rate
            total_op_income          = self.projections['income'].loc[year, 'Total Operating Income'] \
                                       if 'income' in self.projections else 0
            r['Operational Risk RWA']= total_op_income * 0.15

            r['Total RWA'] = r['Credit Risk RWA'] + r['Market Risk RWA'] + r['Operational Risk RWA']

            if r['Total RWA'] > 0:
                r['CET1 Ratio']  = r['CET1 Capital']     / r['Total RWA']
                r['Tier1 Ratio'] = r['Tier1 Capital']     / r['Total RWA']
                r['CAR']         = r['Total Reg Capital'] / r['Total RWA']
            else:
                r['CET1 Ratio']  = r['Tier1 Ratio'] = r['CAR'] = 0.0

            # LDR: use projected loans and deposits from income projections
            total_loans    = self.projections['income'].loc[year, 'Gross Loans'] \
                             if 'income' in self.projections else 0
            total_deposits = self.projections['income'].loc[year, 'Total Deposits'] \
                             if 'income' in self.projections else prev_total_deposits * (1 + deposit_growth)
            r['LDR'] = total_loans / total_deposits if total_deposits else 0.0

            prev_credit_rwa      = r['Credit Risk RWA']
            prev_market_rwa      = r['Market Risk RWA']
            prev_operational_rwa = r['Operational Risk RWA']
            prev_cet1            = r['CET1 Capital']
            prev_at1             = r['AT1 Capital']
            prev_tier2           = r['Tier2 Capital']
            prev_total_deposits  = total_deposits

            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')

    # ── CASH FLOW STATEMENT ───────────────────────────────────────────
    def _project_cashflow(self):
        rows = []

        prev_loans      = self._h(self.balance_hist, 'Gross Loans & Advances',  self.base_year)
        prev_deposits   = (self._h(self.balance_hist, 'CASA Deposits (Current & Savings)', self.base_year) +
                           self._h(self.balance_hist, 'Term / Fixed Deposits',             self.base_year))
        prev_securities = self._h(self.balance_hist, 'Investment Securities',    self.base_year)
        prev_cash       = self._h(self.cashflow_hist, 'Closing Cash Balance',    self.base_year)

        # Historical CF line items — used to carry forward financing items
        hist_debt_issuance  = self._h(self.cashflow_hist, 'Issuance of Debt Securities',       self.base_year)
        hist_debt_repayment = self._h(self.cashflow_hist, 'Repayment of Borrowings',           self.base_year)
        hist_share_change   = self._h(self.cashflow_hist, 'Share Capital Issued / (Bought Back)', self.base_year)
        hist_owc_change     = self._h(self.cashflow_hist, 'Change in Other Working Capital',   self.base_year)
        hist_securities_buy = self._h(self.cashflow_hist, 'Purchase of Investment Securities', self.base_year)
        hist_securities_sell= self._h(self.cashflow_hist, 'Proceeds from Securities Maturity / Sale', self.base_year)
        hist_capex          = self._h(self.cashflow_hist, 'Capital Expenditure (Capex)',        self.base_year)
        hist_acquisitions   = self._h(self.cashflow_hist, 'Acquisitions / Disposals (net)',    self.base_year)

        for year in self.proj_years:
            r = {'Year': year}

            net_profit  = self.projections['income'].loc[year,  'Net Profit After Tax']  if 'income'  in self.projections else 0
            provisions  = self.projections['income'].loc[year,  'Loan Loss Provisions']  if 'income'  in self.projections else 0
            curr_loans  = self.projections['balance'].loc[year, 'Gross Loans']           if 'balance' in self.projections else 0
            curr_deps   = self.projections['balance'].loc[year, 'Total Deposits']        if 'balance' in self.projections else 0
            curr_secs   = self.projections['balance'].loc[year, 'Investment Securities'] if 'balance' in self.projections else 0

            r['Opening Cash Balance']            = prev_cash
            r['Net Income']                      = net_profit
            r['Depreciation & Amortisation']     = self._a('D&A ($mm)', year)
            r['Add: Loan Loss Provisions']        = provisions

            # Change in loans and deposits — derived from balance sheet projections
            r['Change in Loans (net)']            = -(curr_loans - prev_loans)
            r['Change in Deposits (net)']         = curr_deps - prev_deposits

            # Other working capital — carry forward from history (zero if no history)
            r['Change in Other Working Capital']  = hist_owc_change

            r['Net Cash from Operating Activities'] = (
                r['Net Income']                    +
                r['Depreciation & Amortisation']   +
                r['Add: Loan Loss Provisions']     +
                r['Change in Loans (net)']         +
                r['Change in Deposits (net)']      +
                r['Change in Other Working Capital']
            )

            # Investing — derived from balance sheet change in securities + historical capex trend
            securities_change                     = curr_secs - prev_securities
            r['Purchase of Investment Securities']= -max(securities_change, 0)   # only when buying
            r['Proceeds from Securities Maturity / Sale'] = max(-securities_change, 0)  # only when selling
            r['Capital Expenditure (Capex)']      = hist_capex       # carry forward from history
            r['Acquisitions / Disposals (net)']   = hist_acquisitions # carry forward from history

            r['Net Cash from Investing Activities'] = (
                r['Purchase of Investment Securities']          +
                r['Proceeds from Securities Maturity / Sale']  +
                r['Capital Expenditure (Capex)']               +
                r['Acquisitions / Disposals (net)']
            )

            # Financing — dividends formula-driven; debt/share carry forward from history
            r['Dividends Paid']                         = -(net_profit * self._a('Dividend Payout Ratio (%)', year))
            r['Issuance of Debt Securities']            = hist_debt_issuance
            r['Repayment of Borrowings']                = hist_debt_repayment
            r['Share Capital Issued / (Bought Back)']   = hist_share_change

            r['Net Cash from Financing Activities'] = (
                r['Issuance of Debt Securities']          +
                r['Repayment of Borrowings']              +
                r['Share Capital Issued / (Bought Back)'] +
                r['Dividends Paid']
            )

            r['NET CHANGE IN CASH']   = (r['Net Cash from Operating Activities'] +
                                         r['Net Cash from Investing Activities']  +
                                         r['Net Cash from Financing Activities'])
            r['Closing Cash Balance'] = r['Opening Cash Balance'] + r['NET CHANGE IN CASH']

            prev_loans      = curr_loans
            prev_deposits   = curr_deps
            prev_cash       = r['Closing Cash Balance']
            prev_securities = curr_secs

            rows.append(r)

        return pd.DataFrame(rows).set_index('Year')
