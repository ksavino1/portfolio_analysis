# Name: Kyle Savino

# You may not import any additional libraries for this challenge besides the following
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt
import yfinance as yf


class PortfolioAnalysis:
    def __init__(self, input_file):
        self.input_file = input_file
        self.output_file = "cleaned_data.xlsx"
        self.sheets = len(pd.ExcelFile(input_file).sheet_names)
        self.clean_data()
        self.asset_values = self.asset_values()
        self.unrealized_pnl = self.unrealized_returns()

    def convertFloat(self, value):
        '''Used for cleaning values in the dummy_data which have quotes'''
        return float(value.strip('"'))


    def clean_data(self):
            """This function iterates through each page of the dummy_data, and goes through each row and column
            to check if the values are clean (ints). If the unitcost is NA, it is approximated by first checking to see if we
            can steal the value from a neighboring month (if no transactions were made in that time period), if not, then
            we approximate it by taking the average closing value of the stock in that month.
            More explanation is provided in comments in the function.
            I used 28 as the "end" of the month to account for shorter months (february). This way of doing it sacrifices
            precision for ease of scaling if this function were to be used on larger datasets with more variety in months.
            There is a more precise way to calculate the price by just hard coding the days in the month - that is how
            I implemented the clean_data for the Market Price column. This methodology is more precise but messier,
            although with the Calendar library, we could make it cleaner.
            I did both implementations to demonstrate the pros/cons of each way. In a real life situation, this is
            something I would ask my teammates/boss to see what they prefer, or if there is a better way"""
            '''Note: I could not discern any sort of pattern for how many significant digits are displayed in the excel 
            sheet, so there is no specific formatting for the cleaned_data sheet.'''
            cleaned_data = {}
            for sheet in range(self.sheets):
                dataframe = pd.read_excel(self.input_file, sheet_name=sheet)
                for row in range(len(dataframe)):
                    for col in range(len(dataframe.iloc[row]))[2:]:  # start from the third column, makes it faster
                        # since we know that the 0 and 1 columns aren't needed here
                        if type(dataframe.iat[row, col]) != int:
                            value = (str(dataframe.iat[row, col])).strip()
                            # converting to str and stripping to catch any whitespace problems
                            if col == 2:
                                if pd.isna(dataframe.iat[row, col]):
                                    '''This code checks to see we can replace any nan values in the unitcost
                                     column with their values in the neighboring month IF and ONLY if the quantity
                                     of that stock in the portfolio has not changed over this period. Since I cannot see
                                      the actual dates of any transactions, I am making the assumption that a constant
                                      quantity means that no trades were made. This isn't necessarily true though, for instance
                                      if we sold 20 of a stock at 8-03-2023, and then bought 20 on 8-05-2023, we would still
                                      be at net 0 change in quantity, but the unitcost would be different. However, given that this
                                      exact scenario is unlikely to happen, and even if it did, the values for unitcost would likely 
                                      still be somewhat similar, I believe this is the best option we have to approximate unitcost 
                                      without having access to dates when transactions were made.'''
                                    secondSheet = pd.read_excel(self.input_file, sheet_name = 1)
                                    for i in range(len(secondSheet)):
                                        if secondSheet.iat[i, 0] == dataframe.iat[row, 0] and secondSheet.iat[i, 1] == dataframe.iat[row, 1]:
                                            dataframe.iat[row, col] = secondSheet.iat[i, 2]
                                            break
                                        '''If the stock in question's quantity was changed in the nearest month (whether 1 share was sold/bought
                                        or it was sold entirely or anything inbetween) then we do the next best option to approximate unit cost
                                        which is by setting it equal to the average closing price of that stock in the given month'''
                                    if pd.isna(dataframe.iat[row, col]):
                                        start = dt.datetime(2023, 7 + sheet, 1)
                                        end = dt.datetime(2023, 7 + sheet, 28)
                                        stock = yf.download(dataframe.iat[row, 0], start=start, end=end)
                                        dailyPrice = stock["Adj Close"]
                                        monthlyAveragePrice = dailyPrice.resample("M").mean()
                                        monthlyAveragePrice = monthlyAveragePrice.values[0]
                                        dataframe.iat[row, col] = monthlyAveragePrice
                                else:
                                    dataframe.iat[row, col] = (self.convertFloat(value))
                            if col == 3:
                                '''Sometimes the market is closed at the end of the month, like if its a weekend
                                or holiday. To account for that, this code repeatedly goes back days 1 by 1 until it has found
                                the last day of the month for that which the market was open, in order to get the adjusted
                                close per share. If there is no adjusted close, it defaults to close'''
                                if pd.isna(dataframe.iat[row, col]):
                                    if sheet == 0:
                                        data = yf.download(str(dataframe.iat[row, 0]), start="2023-07-01",end="2023-07-31")
                                        for date in ["2023-07-31", "2023-07-30", "2023-07-29", "2023-07-28"]:
                                            try:
                                                stockPrice = data.loc[date]["Adj Close"]
                                                break
                                            except (KeyError, IndexError):
                                                try:
                                                    stockPrice = data.loc[date]["Close"]
                                                    break
                                                except (KeyError, IndexError):
                                                    continue
                                    if sheet == 1:
                                        data = yf.download(str(dataframe.iat[row, 0]), start="2023-08-01",end="2023-08-31")
                                        for date in ["2023-08-31", "2023-08-30", "2023-08-29", "2023-08-28"]:
                                            try:
                                                stockPrice = data.loc[date]["Adj Close"]
                                                break
                                            except (KeyError, IndexError):
                                                try:
                                                    stockPrice = data.loc[date]["Close"]
                                                    break
                                                except (KeyError, IndexError):
                                                    continue
                                    if sheet == 2:
                                        data = yf.download(str(dataframe.iat[row, 0]), start="2023-09-01", end="2023-09-30")
                                        for date in ["2023-09-30", "2023-09-29", "2023-09-28", "2023-09-27"]:
                                            try:
                                                stockPrice = data.loc[date]["Adj Close"]
                                                break
                                            except (KeyError, IndexError):
                                                try:
                                                    stockPrice = data.loc[date]["Close"]
                                                    break
                                                except (KeyError, IndexError):
                                                    continue
                                    dataframe.iat[row, col] = stockPrice
                                else:
                                    dataframe.iat[row, col] = (self.convertFloat(value))
                cleaned_data[sheet] = dataframe
            with pd.ExcelWriter(self.output_file) as writer:
                for sheet, dataframe in cleaned_data.items():
                    sheet_name = f"2023-{sheet + 7:02d}-31"
                    if sheet == 2:
                        sheet_name = "2023-09-30"
                    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    def asset_values(self):
        '''First gets the tickers (exlcuding duplicates) from each page
        Then uses a for each loop to scan the sheets for the desired ticker, and add that value
        to the outputted dataframe.'''
        answer = []
        tickers = []
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name=sheet)
            for row in range(len(dataframe)):
                if dataframe.iat[row, 0] not in tickers:
                    tickers.append(dataframe.iat[row,0])
            for row in range(len(dataframe)):
                answer.append([dataframe.iat[row, 0], (dataframe.iat[row, 3] * dataframe.iat[row, 1])])
        tickers.append("NAV")
        theDataFrame = pd.DataFrame(columns = ['2023-07-31', '2023-08-31', '2023-09-30'], index = tickers)
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name=sheet)
            for i in range(len(theDataFrame)):
                for j in range(len(dataframe)):
                    if theDataFrame.index[i] == dataframe.iat[j, 0]:
                        theDataFrame.iat[i, sheet] = dataframe.iat[j, 3] * dataframe.iat[j, 1]
            if sheet == 0:
                sum = theDataFrame['2023-07-31'].sum()
            if sheet == 1:
                sum = theDataFrame['2023-08-31'].sum()
            if sheet == 2:
                sum = theDataFrame['2023-09-30'].sum()
            theDataFrame.iat[len(theDataFrame) -1, sheet] = sum

        return theDataFrame.fillna(0)


    def unrealized_returns(self):
        '''Uses the same logic as asset_values for the layout of the dataframe, the main difference
        is that for this function, the values are determined with the equation:
        (marketprice - unitcost) * quantity = unrealized return'''
        answer = []
        tickers = []
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name=sheet)
            for row in range(len(dataframe)):
                if dataframe.iat[row, 0] not in tickers:
                    tickers.append(dataframe.iat[row,0])
            for row in range(len(dataframe)):
                answer.append([dataframe.iat[row, 0], (dataframe.iat[row, 3] * dataframe.iat[row, 1])])
        theDataFrame = pd.DataFrame(columns = ['2023-07-31', '2023-08-31', '2023-09-30'], index = tickers)
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name=sheet)
            for i in range(len(theDataFrame)):
                for j in range(len(dataframe)):
                    if theDataFrame.index[i] == dataframe.iat[j, 0]:
                        theDataFrame.iat[i, sheet] = (dataframe.iat[j, 3] - dataframe.iat[j, 2]) * dataframe.iat[j, 1]
        return theDataFrame.fillna(0)


    def calculatePortfolioValues(self):
        '''Returns the total value of the portfolio at each month, used for the following
        graphing functions'''
        portfolio_values = []
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name= sheet)
            portfolio_value = (dataframe['Quantity'] * dataframe['MarketPrice']).sum()
            portfolio_values.append(portfolio_value)
        return portfolio_values


    def plot_portfolio(self):
        '''Creates a plot of the portfolio value over the course of the months'''
        portfolio_values = self.calculatePortfolioValues()
        portfolio_series = pd.Series(portfolio_values, index= range(self.sheets))
        plt.figure(figsize=(11, 7))
        plt.plot(portfolio_series.index, portfolio_series.values, label='Portfolio Value', color='blue', marker='o')
        plt.title('Portfolio Value From July 31st to September 31st')
        plt.xlabel('Months since 7-31')
        plt.ylabel('Value (USD)')
        plt.grid(True)
        plt.ylim(bottom=150000, top=max(portfolio_series.values) * 1.05)
        plt.show()


    def plot_liquidity(self):
        '''Creates a plot of the liquidity (percent of assets in cash) of the portfolio over the months'''
        portfolioValues = self.calculatePortfolioValues()
        values = []
        for sheet in range(self.sheets):
            dataframe = pd.read_excel(self.output_file, sheet_name=sheet)
            values.append(dataframe.iloc[-1, 1])
        liquidity = []
        for cash, portfolio in zip(values, portfolioValues):
            if portfolio != 0:
                ratio = cash / portfolio
            else:
                ratio = 0
            liquidity.append(ratio)
        liquiditySeries = pd.Series(liquidity, index= range(self.sheets))
        plt.figure(figsize=(10, 6))
        plt.plot(liquiditySeries.index, liquiditySeries.values, label='Liquidity Ratio', color='blue', marker='o')
        plt.title('Liquidity of Portfolio Over Time')
        plt.xlabel('Months since 7-31')
        plt.ylabel('Percent of portfolio in cash')
        plt.grid(True)
        plt.legend()
        plt.ylim(bottom=0, top=0.1)
        plt.show()


if __name__ == "__main__":  # Do not change anything here - this is how we will test your class as well.
    fake_port = PortfolioAnalysis("dummy_data.xlsx")
    print(fake_port.asset_values)
    print(fake_port.unrealized_pnl)
    fake_port.plot_portfolio()
    fake_port.plot_liquidity()
