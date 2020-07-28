# Stock Analysis using VBA

## Overview of Project

### Purpose
Using Excel and VBA, we analyzed the data for 12 different stocks from two 2017 and 2018 in order to find which stocks had the best yearly return. VBA is a developer tool used to create Macros (often called subroutines), which allows developers to automate task which otherwise would take longer times to permform using Excel alone. For our Module two, we had created a subroutine called AllStockAnalysis() which found the total volume and yearly return for a dozen stocks. Using our AllStockAnalysis() subroutine, we refactored the code in order to be able to use our script to analyze worksheets with thousands of stocks. Using our subroutine called AllStocksAnalysisRefactored() we were able to automate finding the total yearly volume as well as finding the yearly return for the 12 different stocks for both years.

## Results

### Results from Analysis
From out results, we see that there is a better Yearly Return in 2017 for the majority of the stocks, with the exception of the "TERP" stock. Below is the results from our AllStocksAnalysisRefactored() subroutine. Notice the time to obtain these results took a fraction of a second (0.089 sec).

![](https://github.com/evflores001/UCB-Homework/blob/master/Challenge2/Resources/VBA_Challenge_2017.png) 
 
The subsequent year we notice that the majority of the stocks did not have a positive Yearly Return. The only stocks that had consecutive positive returns were the "ENPH" and "RUN" stocks. Below is the spreadsheet displaying the results for the 2018 year. The elapsed time to obtain these results was 0.094 seconds.

![](https://github.com/evflores001/UCB-Homework/blob/master/Challenge2/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code
In general, many developers use other developers code to complete task or to better understand concepts. Advantages to using refactored code is that it allows developers find solutions to problems that other developers have already faced. Developers might actually use their time at work looking at, and working with code from others.
Disadvatages from refactoring code is that some code might be outdated and therefore more difficult to understand and use. 

### Advantages and Disadvantages of the Original and Refactored VBA script
Using the original VBA script served as a great point of reference when writing our refactored code, however it was important to notice that not everything could have been copied and pasted over to our refactored VBA script as some variables had to be changed to comply with out new refactored variables.
From our analysis, we found that an advantage of using our refactored code gave us faster processing time, as well as better functionality when it comes to working with much larger sets than using our original VBA script.
A disadvantage for using the original code was that it would not be effecient in working with large data sets as it would take significantly longer to output our results.


