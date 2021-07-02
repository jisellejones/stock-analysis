# Stock Analysis

## Overview
The program under analysis was originally written to analyze the daily volume and return on a small amount of stock data. The code was refactored to reduce run time, so it could more efficiently analyze thousands of lines of stock data. The purpose of this analysis is to determine whether or not the refactored code is running at a significantly reduced rate to make computing more efficient.

## Results
The original code was written to loop through 12 different types of stock data called "Tickers". Both programs were coded the same setting up the input box, the format for the table in the spreadsheet, the timer, and declaring an array for the different Tickers. Variables for starting and ending price were also declared. At this point, the code differs. The refactored code declares a variable and several arrays: Dim tickerIndex As Integer, Dim tickerVolumes(12) As Long, Dim tickerStartingPrices(12) As Single, Dim tickerEndingPrices(12) As Single. The array hold onto the data for the report as the program moves through each ticker. The tickerIndex moves to the next ticker with each iteration for the initial For loop.
    

The run times were around 0.83 seconds for each year of stock analyzed.
 
Once refactored, the run time for the same same stock data decreased by about 0.69 seconds *(Figure 1 and Figure 2)*.

*Figure 1- 2017 Refactored Run Time*

![2017_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

*Figure 2 - 2017 Refactored Run Time*

![2018_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary


### Advantages & Disadvantages
