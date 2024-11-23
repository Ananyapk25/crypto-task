import requests
import pandas as pd
import xlwings as xw
import time

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data from CoinGecko API:", response.status_code)
        return []

# Function to process data into a DataFrame
def process_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    df.columns = ["Name", "Symbol", "Price (USD)", "Market Cap", "24h Volume", "24h Change (%)"]
    return df

# Function to analyze data
def analyze_data(df):
    top_5_by_market_cap = df.nlargest(5, "Market Cap")[["Name", "Market Cap"]]
    avg_price = df["Price (USD)"].mean()
    highest_change = df.nlargest(1, "24h Change (%)")[["Name", "24h Change (%)"]]
    lowest_change = df.nsmallest(1, "24h Change (%)")[["Name", "24h Change (%)"]]

    analysis = {
        "Top 5 by Market Cap": top_5_by_market_cap,
        "Average Price": avg_price,
        "Highest 24h Change": highest_change,
        "Lowest 24h Change": lowest_change,
    }
    return analysis

# Function to update Excel sheet
def update_excel(df, analysis):
    # Start or connect to an active Excel instance
    try:
        app = xw.App(visible=True) if not xw.apps else xw.apps.active
    except:
        raise Exception("Could not start or connect to Excel. Ensure Excel is installed and try again.")

    # Create a new workbook or open the specified workbook if it exists
    if "Crypto_Live_Data.xlsx" in [wb.name for wb in xw.books]:
        wb = xw.Book("Crypto_Live_Data.xlsx")
    else:
        wb = xw.Book()
        wb.save("Crypto_Live_Data.xlsx")  # Save the new workbook

    sheet = wb.sheets[0]

    # Write data to the workbook
    sheet.range("A1").value = "Live Cryptocurrency Data"
    sheet.range("A3").value = df

    # Write analysis results
    sheet.range(f"A{len(df) + 5}").value = "Analysis"
    sheet.range(f"A{len(df) + 6}").value = "Top 5 by Market Cap"
    sheet.range(f"A{len(df) + 7}").value = analysis["Top 5 by Market Cap"]
    sheet.range(f"A{len(df) + 12}").value = f"Average Price: {analysis['Average Price']:.2f} USD"
    sheet.range(f"A{len(df) + 13}").value = "Highest 24h Change"
    sheet.range(f"A{len(df) + 14}").value = analysis["Highest 24h Change"]
    sheet.range(f"A{len(df) + 16}").value = "Lowest 24h Change"
    sheet.range(f"A{len(df) + 17}").value = analysis["Lowest 24h Change"]

    wb.save("Crypto_Live_Data.xlsx")

# Main loop for updating
def main():
    while True:
        data = fetch_crypto_data()
        if data:
            df = process_data(data)
            analysis = analyze_data(df)
            update_excel(df, analysis)
        else:
            print("No data fetched. Retrying in 5 minutes...")
        time.sleep(60)  # Update every 5 minutes

if __name__ == "__main__":
    main()