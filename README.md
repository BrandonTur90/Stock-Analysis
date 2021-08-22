# Green Stocks Analysis

## Overview
  The first goal of this project was to run an analysis on 12 green energy stocks, using data from from 2017 and 2018,
  to make sound predictions on which would be good investement choices. The second goal was to then refactor the code
  that was originally written, so as to make it run smoother and faster.
  


## Results


### 2017

![2017 Stocks at a Glance](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/2017%20Stocks%20at%20a%20Glance.png?raw=true)



### 2018

![2018 Stocks at a Glance](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/2018%20Stocks%20at%20a%20Glance.png?raw=true)



  In 2017 it would have seemed very safe to invest in almost any of the analyzed stocks. Of the 12, only 1 (TERP) had a negative
return, and 6 had a return of 50% or higher. The highest being DQ with a massive 199.4% return. The market however, changes, and
in 2018 you can see most of these stocks are now in the negatives. If you had invested both years into ENPH however, you would
still be quite happy with their 81.9% return in 2018, even if it did fall short of their return in 2017 of 129.5%. The only stock
to lose money both years was TERP.


## The Code

Originally there were two separate macros that needed to be run for the analysis. One to actually do the analysis and another to
better format, and display the analysis with color coding in Excel.

![Sub All Stocks](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/Sub%20All%20Stocks.png?raw=true)

This was the begining of the original analysis code


![Sub Format](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/Sub%20Format.png?raw=true)

And this was needed to run after the first, so as to format and display the results with color coding.



Obviously this was not a smooth code to run or use, and initial times to run them both was about 5-8 seconds give or take,
even with easily placeable buttons. Instead it was needed to run them both and make sure they could work together. As for
the code running the analysis, here's a quick breakdown.


1. The first part of the code was to set the output worksheet, start a timer, create basic headers, and initialize a 
   list of the stocks to be analyzed. This was mostly the same in the original and refactored code, so the
   Sub AllStocksAnalysis image above is what was used.
   
2. Next part of the code was to create an index using the above stock tickers called tickerIndex, along with 3 new
   output arrays. The active worksheet was reset, just to be sure it would run correctly, and set tickerVolumes to 0.
   
   ![Refactored 1](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/Refactored%201.png?raw=true)

3. With the setup out of the way we get to the main part of the code that's doing the analysis. Using a for loop
   the code knows to scan all the rows of data in the spreadsheet. It then checks if the ticker of the stock it's
   checking is the same that's in the tickerIndex, and if it is, it adds up the volume of the stock to get its
   totalVolume. It also checks to find the first and last instance of a stock ticker, with the first being indicative
   of a stock's starting price, and the last being its ending price.
   
   ![Refactored 2](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/Refactored%202.png?raw=true)
   
4. The last part of the refactored code outputs the stocks in the header tables from the beginning.
   
   ![Refactored 3](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/Refactored%203.png?raw=true)
   
But wait, there's more! There are timers to measure how fast this process takes.

![VBA_Challenge_2017](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)


![VBA_Challenge_2018](https://github.com/HexHaunter/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)


As previously said, two macros had to run and it took around 5-8 seconds to run them both. Now the whole code runs
in just over 1 second. A much needed improvement.


## Summary

Refactoring code has several pros and cons. On the pro side, you can clean up a large block of code to make
something that runs faster, and with hopefully less code, which usually means less room for something to go 
wrong. But on the con side, you have to start small and work up with refactoring. Making large sweeping
changes because it *could* run faster, can often lead to more time debugging, than you'll actually save
by refactoring.

In regards to this specific instance, going from two macros that take 5-8 seconds, to one that does it in 
just over 1 second is faster, but more importantly it's easier. The code that needed to be refactored was
overall pretty simple to do, and arguably worth doing, leading to few debugging issues in the process. If the
code was longer or more complex it might not have been worth the refactoring to save a few seconds.
