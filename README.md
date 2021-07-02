# Stock Analysis

## Overview
The program under analysis was originally written to analyze the daily volume and return on a small amount of stock data. The code was refactored to reduce run time, so it could more efficiently analyze thousands of lines of stock data. The purpose of this analysis is to determine whether or not the refactored code is running at a significantly reduced rate to make computing more efficient.

## Results
The original code was written to loop through 12 different types of stock data called "tickers". Both programs were coded the same in the beginning by setting up the input box, the format for the table in the spreadsheet, the timer, and declaring an array for the different tickers. Variables for starting and ending price were also declared. At this point, the code differs. The refactored code declares a variable and several arrays: 

&emsp;Dim tickerIndex As Integer

&emsp;Dim tickerVolumes(12) As Long

&emsp;Dim tickerStartingPrices(12) As Single

&emsp;Dim tickerEndingPrices(12) As Single. 

The arrays hold onto the data for the report as the program moves through each ticker. The tickerIndex moves to the next ticker with each iteration for the initial For loop. The for loop and conditional statements are similar in setup; however, the difference between the code is how the numbers are being stored as the computer iterates through the program. 

*Original program code determining starting price and holding as a variable*

&emsp;If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

&emsp;&emsp;startingPrice = Cells(j, 6).Value

*Refactored program code determining starting price and holding within an array*

&emsp;If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1) <> tickers(tickerIndex) Then

&emsp;&emsp;tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

In the original code, the computer is holding onto the data like a person might when using their short term memory. It is only meant to be held for a short period of time and used immediately then (as happens to the best of us) is forgotten. This output is nested within the for loop holding all the if statements, so the computer runs through all the rows associated with one ticker then outputs the data before moving to the next ticker. This means the computer must activate a separate worksheet each time it iterates through a ticker.

*Original program code activating worksheet and outputting data within the for loop*

&emsp;For i...

*The program calls the ticker, sets the volume variable to 0, and runs through all if statements.*

*Then the program activates the output worksheet.*

&emsp;Worksheets("All Stocks Analysis").Activate

*Then outputs the data.*

&emsp;Cells(4 + i, 1).Value = ticker

&emsp;Cells(4 + i, 2).Value = totalVolume

&emsp;Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

*And ends the first iteration*

&emsp;Next i

*The program then begins the next iteration.*

&emsp;For i = 0 To 11

&emsp;&emsp;ticker = tickers(i)

&emsp;&emsp;totalVolume = 0
            
*And reactivates the worksheet where the data is stored.*

&emsp;Worksheets(yearValue).Activate

The refactored code stores the information within the arrays as it iterates through all the tickers. Because of this difference, the computer can iterate through the data more efficiently. Once iterating through all the data in one worksheet, the computer can then activate the next worksheet and output the data. It's like the difference between multi-tasking and focusing on one task at a time. The work will get done, but multi-tasking will take longer because it is less efficient when you are switching between two different tasks.

*The program has already iterated through all the ticker data in the data worksheet and now activates the output worksheet.*

&emsp;Worksheets("All Stocks Analysis").Activate

*Then it loops through each array to output each ticker type, volume value, and return value.*

&emsp;For k = 0 To 11
        
&emsp;&emsp;Cells(4 + k, 1).Value = tickers(k)
&emsp;&emsp;Cells(4 + k, 2).Value = tickerVolumes(k)
&emsp;&emsp;Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1

&emsp;Next k


For the original program, the run times were around 0.83 seconds for each year of stock analyzed.
 
Once refactored, the run times for the same same data decreased by about 0.69 seconds *(Figure 1 and Figure 2)*.

*Figure 1- 2017 Refactored Run Time*

![2017_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

*Figure 2 - 2017 Refactored Run Time*

![2018_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
The refactored code increased the speed of the run times by about 0.69 seconds. This will make a significant difference when the code is run over a larger volume of data with possibly thousands of tickers.

### Advantages & Disadvantages
Refactoring code in this way means that the client can run this program on large amounts of stock data while comprising less computer processing memory to do so. This is a big advantage if there are vast amounts of data for the program to run through. The original data for 2017 analyzed 12 types of stock which was about 3,000 rows. This took about 1 second (.83 rounded up). Imagine if this same program were to analyze 3,000 stock types, at the rate of 12 tickers per second, this would take about 4 minutes for the program to run through all the types of stock. It could also take up a lot of processing memory. If I were to run the same 3,000 stock types at a rate of 12 tickers per 0.14 seconds, I could run all 3,000 types of stock in about 35 seconds.

However, if the program is only going to be used on a smaller amount of data, say around 100 different types of stock, it may only take 8 seconds, these 8 seconds may not be worth the amount of time put into refactoring the data.
