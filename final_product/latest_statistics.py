# A module for calculating the latest statistics, including those in categories valuation and financial health

# A class that deals with calculations such as quick and current ratios
class FinancialHealth:
    # Constructor
    def __init__(self, balance_sheet, income_statement):
        self.balance_sheet = balance_sheet
        self.income_statement = income_statement

    # Calculates the quick ratio
    def quick_ratio(self):
        try:
            current_assets = int(self.balance_sheet.at[0, 'totalCurrentAssets'])
            inventory = int(self.balance_sheet.at[0, 'inventory'])
            current_liabilities = int(self.balance_sheet.at[0, 'totalCurrentLiabilities'])

            return (current_assets - inventory) / current_liabilities
        except:
            return None

    # Calculates the current ratio
    def current_ratio(self):
        try:
            current_assets = int(self.balance_sheet.at[0, 'totalCurrentAssets'])
            current_liabilities = int(self.balance_sheet.at[0, 'totalCurrentLiabilities'])

            return current_assets / current_liabilities
        except:
            return None

    # Calculates the debt-to-equity ratio
    def debt_to_equity(self):
        try:
            liabilities = int(self.balance_sheet.at[0, 'totalLiabilities'])
            shareholders_equity = int(self.balance_sheet.at[0, 'totalShareholderEquity'])

            return liabilities / shareholders_equity
        except:
            return None


# Growth statistics, including revenue decrease/ increase
class Growth:
    # Constructor
    def __init__(self, income_statement):
        self.income_statement = income_statement

    # Calculate the revenue growth
    def revenue_growth(self):
        try:
            current_total_revenue = int(self.income_statement.at[0, 'totalRevenue'])
            current_cost_of_revenue = int(self.income_statement.at[0, 'costOfRevenue'])
            previous_total_revenue = int(self.income_statement.at[1, 'totalRevenue'])
            previous_cost_of_revenue = int(self.income_statement.at[1, 'costOfRevenue'])

            current_revenue = current_total_revenue - current_cost_of_revenue
            previous_revenue = previous_total_revenue - previous_cost_of_revenue

            #return ((current_revenue - previous_revenue) / previous_revenue) * 100
            return ((current_total_revenue - previous_total_revenue) / previous_total_revenue) * 100
        except:
            return None

    # Calcuate the operating income growth
    def operating_income_growth(self):
        try:
            current_operating_income = int(self.income_statement.at[0, 'operatingIncome'])
            previous_operating_income = int(self.income_statement.at[1, 'operatingIncome'])

            return ((current_operating_income - previous_operating_income) / previous_operating_income) * 100
        except:
            return None

    # Calcuate the net income growth
    def net_income_growth(self):
        try:
            current_net_income = int(self.income_statement.at[0, 'netIncome'])
            previous_net_income = int(self.income_statement.at[1, 'netIncome'])

            return ((current_net_income - previous_net_income) / previous_net_income) * 100
        except:
            return None