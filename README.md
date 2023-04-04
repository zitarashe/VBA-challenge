# VBA-challenge

I had to use VBA scripting to analyse generated stock market data.

Steps

Created a script that loops through all the stocks for one year and outputs the following information:
* The ticker symbol
* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

The script met the following results:

Retrieval of Data 
* The script loops through one year of stock data and reads/stores all of the following values from each row:
    * ticker symbol 
    * volume of stock 
    * open price 
    * close price 

Column Creation 
* On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
    * ticker symbol 
    * total stock volume 
    * yearly change ($) 
    * percent change

Conditional Formatting 
* Conditional formatting is applied correctly and appropriately to the yearly change column 
* Conditional formatting is applied correctly and appropriately to the percent change column 

Calculated Values
* All three of the following values are calculated correctly and displayed in the output:
    * Greatest % Increase 
    * Greatest % Decrease 
    * Greatest Total Volume 

Looping Across Worksheet 
* The VBA script can run on all sheets successfully.

