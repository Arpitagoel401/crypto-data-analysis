import requests
import pandas as pd
import time
import schedule
from openpyxl import Workbook

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
        print("Error fetching data")
        return []

def analyze_data(data):
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])
    
    # Identify top 5 cryptocurrencies by market cap
    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]
    
    # Calculate the average price of the top 50 cryptocurrencies
    avg_price = df["current_price"].mean()
    
    # Find highest and lowest 24-hour percentage price change
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax(), ["name", "price_change_percentage_24h"]]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin(), ["name", "price_change_percentage_24h"]]
    
    return df, top_5, avg_price, highest_change, lowest_change

def update_excel():
    print("Updating Excel with latest cryptocurrency data...")
    data = fetch_crypto_data()
    if not data:
        return
    
    df, top_5, avg_price, highest_change, lowest_change = analyze_data(data)
    
    # Create or update Excel file
    with pd.ExcelWriter("crypto_data.xlsx", engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Top 50 Cryptos", index=False)
        top_5.to_excel(writer, sheet_name="Top 5 by Market Cap", index=False)
        
        summary_data = pd.DataFrame({
            "Metric": ["Average Price", "Highest Change", "Lowest Change"],
            "Value": [avg_price, f"{highest_change[0]}: {highest_change[1]:.2f}%", f"{lowest_change[0]}: {lowest_change[1]:.2f}%"]
        })
        summary_data.to_excel(writer, sheet_name="Summary", index=False)
    
    print("Excel updated successfully.")

# Schedule updates every 5 minutes
schedule.every(5).minutes.do(update_excel)

if __name__ == "__main__":
    update_excel()  # Run once before scheduling
    while True:
        schedule.run_pending()
        time.sleep(1)
