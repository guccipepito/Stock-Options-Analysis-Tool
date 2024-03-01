# -*- coding: utf-8 -*-
"""

________  ___  ___  ________  ________  ___  ________  _______   ________  ___  _________  ________
|\   ____\|\  \|\  \|\   ____\|\   ____\|\  \|\   __  \|\  ___ \ |\   __  \|\  \|\___   ___\\   __  \
\ \  \___|\ \  \\\  \ \  \___|\ \  \___|\ \  \ \  \|\  \ \   __/|\ \  \|\  \ \  \|___ \  \_\ \  \|\  \
 \ \  \  __\ \  \\\  \ \  \    \ \  \    \ \  \ \   ____\ \  \_|/_\ \   ____\ \  \   \ \  \ \ \  \\\  \
  \ \  \|\  \ \  \\\  \ \  \____\ \  \____\ \  \ \  \___|\ \  \_|\ \ \  \___|\ \  \   \ \  \ \ \  \\\  \
   \ \_______\ \_______\ \_______\ \_______\ \__\ \__\    \ \_______\ \__\    \ \__\   \ \__\ \ \_______\
    \|_______|\|_______|\|_______|\|_______|\|__|\|__|     \|_______|\|__|     \|__|    \|__|  \|_______|


"""


import subprocess
import os  # Importing the os module for filesystem operations

# Install necessary libraries
subprocess.run(['pip', 'install', 'yfinance', 'numpy', 'scipy'])

import numpy as np  # Importing the NumPy library for numerical calculations
from scipy.stats import norm  # Importing the norm function from the SciPy library for statistics
from datetime import datetime  # Importing the datetime class for date manipulation
import yfinance as yf  # Importing the yfinance library to retrieve stock data
import pandas as pd  # Importing the Pandas library for tabular data manipulation
from openpyxl import Workbook  # Importing the Workbook class from the openpyxl library to create Excel files
from openpyxl.styles import Alignment, Font  # Importing certain classes for Excel file formatting
from google.colab import files  # Importing the files function from google.colab to download files

def get_option_expiration_dates(symbol):
    """
    Gets the expiration dates of options for a given symbol.

    Args:
    symbol: Stock symbol

    Returns:
    List of option expiration dates
    """
    ticker = yf.Ticker(symbol)  # Retrieving ticker data
    expirations = ticker.options  # Retrieving option expiration dates
    return expirations

def calculate_implied_volatility(S, K, T, r, price, option_type):
    """
    Calculates the implied volatility using the Black-Scholes model.

    Args:
    S: Current stock price
    K: Option strike price
    T: Time to expiration in years
    r: Risk-free interest rate
    price: Option price
    option_type: 'call' or 'put'

    Returns:
    Implied volatility
    """
    tol = 0.0001  # Tolerance for implied volatility convergence
    max_iter = 100  # Maximum number of iterations

    def black_scholes(option_type, S, K, T, r, sigma):
        """
        Calculates the option price using the Black-Scholes model.
        """
        d1 = (np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T))
        d2 = d1 - sigma * np.sqrt(T)

        if option_type == 'call':
            option_price = S * norm.cdf(d1) - K * np.exp(-r * T) * norm.cdf(d2)
        else:
            option_price = K * np.exp(-r * T) * norm.cdf(-d2) - S * norm.cdf(-d1)

        return option_price

    sigma = 0.5  # Initial guess for implied volatility
    price_est = black_scholes(option_type, S, K, T, r, sigma)  # Initial estimate of option price
    diff = price_est - price  # Difference between estimated price and observed price
    iter_count = 0  # Iteration counter initialization

    while abs(diff) > tol and iter_count < max_iter:
        vega = S * np.sqrt(T) * norm.pdf((np.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * np.sqrt(T)))
        price_est = black_scholes(option_type, S, K, T, r, sigma)  # Estimate of option price
        diff = price_est - price  # New difference between estimated price and observed price

        if abs(diff) < tol:  # Checking for convergence
            break

        sigma = sigma - (diff / vega)  # Updating implied volatility
        iter_count += 1  # Incrementing iteration counter

    return sigma

def calculate_intrinsic_value(S, K, option_type):
    """
    Calculates the intrinsic value of an option.

    Args:
    S: Current stock price
    K: Option strike price
    option_type: 'call' or 'put'

    Returns:
    Intrinsic value
    """
    if option_type == 'call':
        return max(0, S - K)
    else:
        return max(0, K - S)

def calculate_time_value(price, intrinsic_value):
    """
    Calculates the time value of an option.

    Args:
    price: Option price
    intrinsic_value: Intrinsic value of the option

    Returns:
    Time value
    """
    return max(0, price - intrinsic_value)

def calculate_historical_volatility(symbol, start_date, end_date):
    """
    Calculates historical volatility using historical price data.

    Args:
    symbol: Stock symbol
    start_date: Start date for historical data
    end_date: End date for historical data

    Returns:
    Historical volatility
    """
    stock = yf.Ticker(symbol)  # Retrieving stock data
    historical_data = stock.history(start=start_date, end=end_date)  # Retrieving historical data
    returns = np.log(historical_data['Close'] / historical_data['Close'].shift(1))  # Calculating log returns
    volatility = returns.std() * np.sqrt(252)  # Calculating historical volatility (252 trading days in a year)
    return volatility

def fetch_option_info(ticker, expiry_date):
    """
    Fetches option information for a given ticker and expiry date.

    Args:
    ticker: Stock ticker
    expiry_date: Option expiration date

    Returns:
    Option information
    """
    stock = yf.Ticker(ticker)  # Retrieving stock data
    trade_date = datetime.now()  # Current date
    trade_date = trade_date.replace(tzinfo=None)  # Converting to datetime without timezone

    expiry_date = datetime.strptime(expiry_date, '%Y-%m-%d').replace(tzinfo=None)  # Converting expiration date to datetime object

    if expiry_date < trade_date:
        raise ValueError("The specified expiration date has already passed.")

    option_chain = stock.option_chain(expiry_date.strftime('%Y-%m-%d'))  # Retrieving option chain data
    if len(option_chain.calls) > 0:
        option = option_chain.calls.iloc[0]  # Assuming you're interested in the first available call option
        underlying_price = stock.history(period='1d').iloc[-1]['Close']  # Underlying stock price
        strike_price = option['strike']  # Option strike price
        risk_free_rate = 0.05  # Assumed risk-free interest rate

        days_to_expiry = (expiry_date - trade_date).days / 365  # Time to expiration in years

        # Calculating implied volatility
        implied_volatility = calculate_implied_volatility(underlying_price, strike_price, days_to_expiry, risk_free_rate, option['lastPrice'], 'call')

        # Calculating intrinsic value for call option
        call_intrinsic_value = calculate_intrinsic_value(underlying_price, strike_price, 'call')

        # Calculating intrinsic value for put option
        put_intrinsic_value = calculate_intrinsic_value(underlying_price, strike_price, 'put')

        # Calculating time value
        time_value = calculate_time_value(option['lastPrice'], call_intrinsic_value)

        return underlying_price, strike_price, days_to_expiry, risk_free_rate, implied_volatility, call_intrinsic_value, put_intrinsic_value, time_value
    else:
        raise ValueError("No call option available for the given ticker and expiry date.")

def main():
    symbol = input("Enter the stock symbol: ")
    expiration_dates = get_option_expiration_dates(symbol)

    if expiration_dates:
        option_data_list = []
        for exp_date in expiration_dates:
            try:
                option_data = fetch_option_info(symbol, exp_date)
                # Calculating historical volatility
                start_date = (datetime.now() - pd.Timedelta(days=365)).strftime('%Y-%m-%d')  # One year ago
                end_date = datetime.now().strftime('%Y-%m-%d')
                historical_volatility = calculate_historical_volatility(symbol, start_date, end_date)
                option_data_list.append([exp_date] + list(option_data) + [historical_volatility])
            except ValueError as e:
                print(e)

        # Creating a DataFrame with option data
        columns = ['Expiration Date', 'Underlying Price', 'Strike Price', 'Days to Expiry', 'Risk-Free Rate', 'Implied Volatility', 'Call Intrinsic Value', 'Put Intrinsic Value', 'Time Value', 'Historical Volatility']
        df = pd.DataFrame(option_data_list, columns=columns)

        # Counting the number of existing files for the same ticker
        num_files = sum(f.startswith(f"{symbol}_option_data") for f in os.listdir('.'))
        num_files += 1

        # Exporting data to an Excel file with increasing numbering
        excel_file = f"{symbol}_option_data_{num_files}.xlsx"
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Option Data')

            # Applying style to the Excel file
            workbook = writer.book
            worksheet = writer.sheets['Option Data']
            header_font = Font(bold=True)
            align_center = Alignment(horizontal='center')
            for cell in worksheet["1:1"]:
                cell.font = header_font
                cell.alignment = align_center

            # Adding a title
            title_cell = worksheet.cell(row=1, column=1)
            title_cell.value = "Option Data"
            title_cell.font = Font(size=14, bold=True)
            title_cell.alignment = Alignment(horizontal='center')

            # Automatically adjusting column widths
            for column_cells in worksheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                worksheet.column_dimensions[column].width = adjusted_width

        print(f"Option data has been saved to {excel_file}")

        # Downloading the Excel file
        files.download(excel_file)
    else:
        print("No option expiration dates found for {}".format(symbol))

if __name__ == "__main__":
    main()
