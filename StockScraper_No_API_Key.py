import requests
import xlwt
import xlrd
from xlutils.copy import copy
import datetime

# Alpha Vantage API key
api_key = "YOUR_API_KEY"

# List of stock symbols
stock_symbols = ["GOOG", "MSFT", "INTC", "CSCO", "AAPL"]

# File name for the Excel file
file_name = "stock_data.xls"

# Current date is setup to be 7 days behind, because Alpha Vantage requires premium access to get daily data, and sometimes doesn't update every day.
today = datetime.date.today()
lastweek = today - datetime.timedelta(days=7)
current_date = lastweek.strftime("%Y-%m-%d")

# Keep track of the last row number written to the sheet
last_row = 0

# Check if the file already exists
try:
    # Open the existing file in read mode
    workbook = xlrd.open_workbook(file_name)

    # Make a copy of the existing file
    new_workbook = copy(workbook)
    sheet_names = workbook.sheet_names()

except FileNotFoundError:
    # Create a new file if it doesn't exist
    new_workbook = xlwt.Workbook()
    sheet_names = []

# Iterate through the list of stock symbols
for symbol in stock_symbols:
    # Make a request to the Alpha Vantage API
    url = f"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol={symbol}&apikey={api_key}"
    response = requests.get(url)
    data = response.json()

    # Get the sheet name
    sheet_name = symbol

    if sheet_name in sheet_names:
        sheet = new_workbook.get_sheet(sheet_name)
    else:
        sheet = new_workbook.add_sheet(sheet_name)

    # Write the headers to the sheet
    if last_row == 0:
        sheet.write(0, 0, "Date")
        sheet.write(0, 1, "Open")
        sheet.write(0, 2, "High")
        sheet.write(0, 3, "Low")
        sheet.write(0, 4, "Close")
        last_row += 1

    # Get the data for the current date
    date_data = data["Time Series (Daily)"][current_date]

    # Get the next empty row in the sheet
    current_row = last_row
    next_row = current_row

    # Write the date, open, high, low, close, and volume to the sheet
    sheet.write(next_row, 0, current_date)
    sheet.write(next_row, 1, date_data["1. open"])
    sheet.write(next_row, 2, date_data["2. high"])
    sheet.write(next_row, 3, date_data["3. low"])
    sheet.write(next_row, 4, date_data["4. close"])
    last_row += 1

# Save the new workbook
new_workbook.save(file_name)

#Print success message
print("Data has been appended successfully to the sheets!")

