# VBA Stock Market Analysis

## Project Overview

This project utilizes VBA scripting to efficiently analyze stock market data across multiple quarters and years. The VBA script loops through all stock data for each quarter and calculates essential metrics such as the ticker symbol, quarterly changes, percentage changes, and total stock volume. Additionally, the script identifies stocks with the "Greatest % increase," "Greatest % decrease," and "Greatest total volume." The analysis results are visualized with conditional formatting and organized into quarterly snapshots.

## Background

This project was undertaken to develop programming skills in VBA and automate data analysis for large datasets. The stock market data analysis is performed across multiple years, allowing for insights into how stocks perform over time. By automating repetitive tasks, the project saves time and ensures consistent results.

## Key Features

### Core Features
- **Ticker Symbol Identification:** The script records ticker symbols for each stock.
- **Quarterly Change Calculation:** Calculates the price change from the opening price to the closing price of each stock for each quarter.
- **Percentage Change Calculation:** Computes the percentage change in stock prices for each quarter.
- **Total Stock Volume Calculation:** Totals the volume of stocks traded for each ticker.

### Bonus Features
- **Greatest % Increase:** Identifies the stock with the highest percentage increase over the quarter.
- **Greatest % Decrease:** Identifies the stock with the largest percentage decrease over the quarter.
- **Greatest Total Volume:** Highlights the stock with the highest total volume traded.

### Additional Functionalities
- **Sheet-Wide Analysis:** The script runs across multiple sheets to analyze data for different quarters and years.
- **Conditional Formatting:** Green is used to highlight positive price changes, and red is used for negative changes, making trends easily visible.
- **Yearly and Quarterly Performance:** Results are displayed per year, with visualizations showing changes and trends across different years.

## Files Included
- **VBA Script:** The script used to perform the stock market analysis.
- **Screenshots:** Screenshots of the output for different years and quarters, showing the analysis results.
- **README:** This file, detailing the project and its functionality.

## How to Use

1. **Clone the Repository:** Clone the repository to your local machine.
2. **Open the Excel File:** Open the Excel file containing stock data for multiple years.
3. **Run the VBA Script:** In Excel, open the VBA editor, load the script, and run it. The script will automatically loop through all worksheets and generate the required analysis.
4. **View the Results:** Results, including quarterly changes and top-performing stocks, will be displayed with conditional formatting, and screenshots of the output for different years are included for reference.

## Screenshots Reference

The analysis includes screenshots that showcase the performance of stocks across different years. These screenshots can be found in the repository, illustrating:
- The stock with the greatest percentage increase, decrease, and total volume per year.
- The conditional formatting applied to each stock based on price changes.
  ### example 
![2019 screenshot](https://github.com/user-attachments/assets/627d20bf-df55-4358-84c0-b289c3e2ff2d)


## Performance Considerations

The script was developed using a smaller dataset (`alphabetical_testing.xlsx`) to allow faster testing and debugging. Once finalized, it efficiently runs on larger datasets across multiple quarters and years, providing quick results.
