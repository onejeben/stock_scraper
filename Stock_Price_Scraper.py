import yfinance as yf
import time
import csv
import pandas as pd
from datetime import datetime
import keyboard
import os

# ✅ ==== SETTINGS ====
stock_symbols = ["AAPL", "TSLA", "MSFT"]
interval_minutes = 1

# ✅ ==== FOLDER SETUP ====
base_folder = r"C:\Users\capta\Dawson\Desktop\test\Stock_Scraper"
os.makedirs(base_folder, exist_ok=True)

# Auto-name files with start time
start_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
csv_file = os.path.join(base_folder, f"stocks_{start_time}.csv")
excel_file = os.path.join(base_folder, f"stocks_{start_time}.xlsx")

# ✅ ==== YFINANCE FUNCTION ====
def get_stock_price(symbol):
    try:
        ticker = yf.Ticker(symbol)
        price = ticker.fast_info.last_price  # Fast and lightweight
        return float(price) if price is not None else None
    except Exception as e:
        print(f"[ERROR] {symbol} fetch failed: {e}")
        return None

# ✅ ==== PREPARE CSV HEADER ====
with open(csv_file, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerow(["Timestamp", "Stock", "Price"])

print(f"Tracking: {', '.join(stock_symbols)}")
print(f"Logging every {interval_minutes} minute(s).")
print(f"Saving files in: {base_folder}")
print("Press 'q' to stop.")

# ✅ ==== MAIN LOOP ====
try:
    while True:
        if keyboard.is_pressed("q"):
            raise KeyboardInterrupt

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(csv_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for symbol in stock_symbols:
                price = get_stock_price(symbol)
                if price is not None:
                    price = round(price, 3)  # ✅ round to 3 decimals
                    writer.writerow([timestamp, symbol, price])
                    print(f"[{timestamp}] {symbol} @ {price}")
                else:
                    writer.writerow([timestamp, symbol, "N/A"])
                    print(f"[{timestamp}] {symbol} @ N/A")
                time.sleep(0.2)  # Small delay to be safe

        # Faster stop response
        total_wait = interval_minutes * 60
        waited = 0
        while waited < total_wait:
            if keyboard.is_pressed("q"):
                raise KeyboardInterrupt
            time.sleep(0.1)
            waited += 0.1

except KeyboardInterrupt:
    print("\nStopped logging.")
    print(f"Data saved in CSV: {csv_file}")

    # ✅ ==== CONVERT TO EXCEL WITH CHART + SUMMARY ====
    df = pd.read_csv(csv_file)

    with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Stock Data")

        workbook = writer.book
        worksheet = writer.sheets["Stock Data"]

        # Header formatting
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Price formatting
        price_format = workbook.add_format({"num_format": "$#,##0.000"})  # ✅ 3 decimal places
        worksheet.set_column("A:A", 20)
        worksheet.set_column("B:B", 10)
        worksheet.set_column("C:C", 12, price_format)

        # ✅ Trend Graph
        chart = workbook.add_chart({"type": "line"})
        for symbol in stock_symbols:
            stock_df = df[df["Stock"] == symbol]
            if not stock_df.empty:
                start_row = stock_df.index[0] + 2
                end_row = stock_df.index[-1] + 2
                chart.add_series({
                    "name": symbol,
                    "categories": f"='Stock Data'!$A${start_row}:$A${end_row}",
                    "values": f"='Stock Data'!$C${start_row}:$C${end_row}",
                })
        chart.set_title({"name": "Stock Prices Over Time"})
        chart.set_x_axis({"name": "Time"})
        chart.set_y_axis({"name": "Price (USD)"})
        worksheet.insert_chart("E2", chart)

        # ✅ Summary Table
        summary_start = len(df) + 4
        worksheet.write(summary_start, 0, "Summary", workbook.add_format({"bold": True, "bg_color": "#BDD7EE"}))
        worksheet.write_row(summary_start + 1, 0, ["Stock", "Highest", "Lowest", "Average"],
                            workbook.add_format({"bold": True, "bg_color": "#D9E1F2"}))

        row = summary_start + 2
        for symbol in stock_symbols:
            stock_df = df[df["Stock"] == symbol]
            if not stock_df.empty:
                highest = stock_df["Price"].max()
                lowest = stock_df["Price"].min()
                average = stock_df["Price"].mean()
                worksheet.write_row(row, 0, [symbol, highest, lowest, average], price_format)
                row += 1

    print(f"Excel file saved with chart + summary: {excel_file}")
