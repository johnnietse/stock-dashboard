# Stock Portfolio Dashboard (Excel)

## Overview
This project provides an Excel-based stock portfolio dashboard that dynamically updates stock prices using the **Stock History function** and **Stock Data Types**. The dashboard requires an **Office 365 or Microsoft 365** subscription to function properly.

## Features
- **Live Stock Price Updates**: Automatically fetches real-time stock data.
- **Portfolio Summary**: Displays total market value, today's change, and overall gains/losses.
- **Stock Breakdown**: Lists individual stocks, their industries, units, prices, and market values.
- **Historical Price Trends**: Uses **Stock History function** to show 12-month price trends.
- **Industry Distribution**: Summarizes holdings by industry with visual representation.
- **Watchlist**: Tracks potential investments with key metrics like price, P/E ratio, and 52-week high.

## Setup Instructions
1. **Ensure Office 365 or Microsoft 365 Subscription**
   - The dashboard relies on Stock Data Types, which are only available in these versions.
2. **Download the Excel File**
   - Clone the repository or download the file from GitHub.
3. **Enable Data Connections**
   - Open the file and ensure external data sources are allowed.
4. **Refresh Data**
   - Click **"Refresh All"** in the **Data** tab to update stock prices.

## How It Works
- **Stock Data Type**: Converts stock names into live stock data linked to online sources.
- **Dynamic Arrays**: Uses formulas like `SORT`, `UNIQUE`, and `FILTER` to organize stock information.
- **Stock History Function**: Retrieves historical stock prices for trend analysis.
- **Conditional Formatting & Sparklines**: Provides visual insights into stock performance.

## Notes
- **Hidden Rows for Portfolio Growth**: Allows easy expansion of portfolio tracking.
- **Manual Copying of Some Formulas**: Due to Excel limitations, additional stocks may require copying formulas manually.
- **Pie Chart vs. Bar Chart**: A pie chart is used for industry allocation, but a bar chart is recommended for larger datasets.
