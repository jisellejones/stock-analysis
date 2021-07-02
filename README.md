# Stock Analysis

## Overview
The program under analysis was originally written to analyze the daily volume and return on a small amount of stock data. The code was refactored to reduce run time, so it could more efficiently analyze thousands of lines of stock data. The purpose of this analysis is to determine whether or not the refactored code is running at a significantly reduced rate to make computing more efficient.

## Results
The original code was written to loop through 12 different types of stock data called "Tickers". Both programs were coded the same setting up the input box, the format for the table in the spreadsheet, the timer, and declaring an array for the different Tickers. Variables for starting and ending price were also declared. At this point, the code differs. The refactored code declares a variable and several arrays: Dim tickerIndex As Integer, Dim tickerVolumes(12) As Long, Dim tickerStartingPrices(12) As Single, Dim tickerEndingPrices(12) As Single. 

The arrays hold onto the data for the report as the program moves through each ticker. The tickerIndex moves to the next ticker with each iteration for the initial For loop. The for loop and conditional statements are similar in setup; however, the difference between the code is how the numbers are being stored as the computer iterates through the program. 

*Original program code:*

 If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
  startingPrice = Cells(j, 6).Value

In the original code, the computer is holding on to the data like a person might when using their short term memory. It is there, then it is used immediately and (as happens to the best of us) is forgotten. Where the refactored code is sotring the information within the arrays to hold onto until the iterations are complete on one spreadsheet before it has to switch to the next spreadsheet. Because of this difference, the computer can iterate through the data more efficiently. Once iterating through the data, the computer can then activate the next worksheet and output the data. It's like the difference between multi-tasking and focusing on one task at a time. The work will get done, but multi-tasking will take longer because it is less efficient.

The run times were around 0.83 seconds for each year of stock analyzed.
 
Once refactored, the run time for the same same stock data decreased by about 0.69 seconds *(Figure 1 and Figure 2)*.

*Figure 1- 2017 Refactored Run Time*

![2017_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

*Figure 2 - 2017 Refactored Run Time*

![2018_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
The refactored code increased the speed of the run times by about 0.69 seconds. This will make a significant difference when the code is run over a larger volume of data with possibly thousands of tickers.

### Advantages & Disadvantages
Refactoring code in this way means that the client can run this program on large amounts of stock data while comprising less computer processing memory to do so. This is a big advantage if there are vast amounts of data for the program to run through. The original data for 2017 analyzed 12 types of stock which was about 3,000 rows. This took about 1 second (.83 rounded up). Imagine if this same program were to analyze 3,000 stock types, at the rate of 12 tickers per second, this would take about 4 minutes for the program to run through all the types of stock. It could also take up a lot of processing memory. If I were to run the same 3,000 stock types at a rate of 12 tickers per 0.2 seconds, I could run all 3,000 types of stock in about 50 seconds.

However, if the program is only going to be used on a smaller amount of data, say around 100 different types of stock, it may only take 8 seconds, these 8 seconds may not be worth the amount of time put into refactoring the data.
