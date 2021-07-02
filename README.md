# Stock Analysis

## Overview
The program under analysis was originally written to analyze the daily volume and return on a small amount of stock data. The code was refactored to reduce run time, so it could more efficiently analyze thousands of lines of stock data. The purpose of this analysis is to determine whether or not the refactored code is running at a significantly reduced rate to make computing more efficient.

## Results
The original code was written to loop through data for 12 different stocks. The original code was outputting the For the original code, run times were around 0.83 seconds for each year of stock analyzed.
 
Once refactored, the run time for the same same stock data decreased by about 0.69 seconds *(Figure 3 and Figure 4)*

*Figure 3 2017 Refactored Run Time*

![2017_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/2017_Refactored_Run_Time.png)

*Figure 4 2017 Refactored Run Time*

![2018_Refactored_Run_Time](https://github.com/jisellejones/stock-analysis/blob/main/Resources/2018_Refactored_Run_Time.png)


## Summary


### Advantages & Disadvantages
