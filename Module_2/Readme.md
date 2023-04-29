#Description 

This project folder contains a VBA script for analyzing stock market data. 

The script loops through each worksheet in an Excel workbook, reads the stock ticker symbol, opening price, closing price, and volume of shares traded for each day in a given year, and calculates the yearly change and percentage change in stock price. 

The script also identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume for each year.

#Instruction

The script starts by declaring some variables that will be used throughout the code. These variables are used to store the ticker symbol, opening price, closing price, and total stock volume for each stock.

Next, the script loops through each row in the worksheet and retrieves the necessary information for each stock.

For each stock, the script calculates the yearly change and percentage change from the opening price to the closing price.

The script also keeps track of the total stock volume for each stock.

After looping through all the rows, the script determines the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

Finally, the script outputs all the information for each stock, as well as the information for the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The output is formatted using conditional formatting to highlight positive changes in green and negative changes in red.

The script is designed to be flexible and can be easily modified to work with other worksheets or data sets.
