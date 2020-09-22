# 02-VBA-Challenge

In this assignment, we were given two Microsoft Excel files, one to test or VBA code on and the other to complete the assignment, detailing a plethora of stock ticker symbols, dates, open/close prices, and 52-week highs/lows.

Using VBA, I crated a table that pulled out the following data:
    - Every ticker symbol
    - The yearly change for each stock
    - The percent change from open price to close price for each stock
    - The total volume of each stock
    - Conditionally formatted the cells to show green/red depending if a stock had a positive or negative change over the course of the year

To find each ticker, I created a conditional to run until it finds the ticker that does not match the previous stock ticker.

By defining the start of each ticker, I pulled the first open value into a variable and used the same conditional as above to find the very last close price.  From here, I simply subtracted the open price from the close price to find our yearly change.

Using the open price and close price we defined above, I created a conditional to first sort out any 0's that would affect our upcoming formula.  Once our code disregards "0" prices, I divided the yearly change by the open price to find our percent change.

My total stock volume was found using the same conditional as finding each stock ticker but pulled out every single value in the volume column and added them together until the stock ticker changed.

Using these functions and conditionals, we can use the table to sort through our table of stocks to find information such as greatest increase, greatest decrease, most volatile, and more.

