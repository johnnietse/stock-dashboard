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

## Screenshot
### Personal Stock Dashboard: A comprehensive view of portfolio performance, trends, and key financial insights—all in one place!
![Stock_dashboard_excel](https://github.com/user-attachments/assets/fd1e53c0-cc4e-462b-8a89-8a7adc0fd008)

# Updates
# Stock Portfolio Dashboard (Excel & Power BI)

## Overview
This project provides a stock portfolio dashboard that dynamically updates stock prices and displays key financial insights. The dashboard was initially built using Excel but has now been fully implemented in Power BI. The data for the Power BI dashboard is imported from an Excel spreadsheet (`My_stock_dashboard_data.xlsx`).

## Features
- **Live Stock Price Updates**: Automatically fetches real-time stock data.
- **Portfolio Summary**: Displays total market value, today's change, and overall gains/losses.
- **Stock Breakdown**: Lists individual stocks, their industries, units, prices, and market values.
- **Historical Price Trends**: Uses historical stock data to show 12-month price trends.
- **Industry Distribution**: Summarizes holdings by industry with visual representation.

## Setup Instructions
### For Power BI:
1. Ensure you have Power BI installed.
2. Import the `My_stock_dashboard_data.xlsx` file as the data source.
3. Refresh the data to load the latest stock prices.

### For Excel:
1. **Ensure Office 365 or Microsoft 365 Subscription**  
   The dashboard relies on Stock Data Types, which are only available in these versions.
2. **Download the Excel File**  
   Clone the repository or download the file from GitHub.
3. **Enable Data Connections**  
   Open the file and ensure external data sources are allowed.
4. **Refresh Data**  
   Click "Refresh All" in the Data tab to update stock prices.

## How It Works
- **Stock Data Type**: Converts stock names into live stock data linked to online sources.
- **Dynamic Arrays**: Uses formulas like SORT, UNIQUE, and FILTER to organize stock information.
- **Stock History Function**: Retrieves historical stock prices for trend analysis.
- **Conditional Formatting & Visualizations**: Provides insights into stock performance.

## Notes
- The Power BI dashboard provides enhanced visualization and interactive analysis.
- The Excel version serves as the data source and can also be used as a standalone version.
- Manual copying of some formulas may be required for additional stocks in Excel.

## Screenshot
**Personal Stock Dashboard (Microsoft Power BI Version)**: A comprehensive view of portfolio performance, trends, and key financial insights—all in one place!
![Screenshot (514)](https://github.com/user-attachments/assets/3a7b9e20-e21b-4b2d-a02b-fb243e93c0dd)



## About
This project integrates both Excel and Power BI to track and analyze stock investments efficiently.


