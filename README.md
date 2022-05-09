# stocks-analysis
## Overview of Project
#### Background
Steve wants to help his parents to analyze the entire stock markets over the last few years and find stocks to invest. 

#### Purpose
This project aims to refactor the code to execute the stock data for 2017 and 2018 in a much more efficient way and quicker execution time.

## Results: 
To evaluate the selected stocks (tickers) performance for 2017 and 2018, we are using the yearValue to allow the user to enter the year of data they want by using yearValue = InputBox("What year would you like to run the analysis on?").  We created a tickerIndex = 0, looping over all rows through stock data with the code For j = 2 To RowCount.  Next, we increase the volume of the current tickerVolume by using the tickerIndex variable as index, use the code: tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value.  After we collected and stored all of the values for the ticker, total volume, starting price and ending price, we calculated the percentage of return based on the starting and ending price with the code: Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1.  Then we output all data into a Stock Analysis worksheet and format the return cells, red and green with the code Cells(k, 3).Interior.Color = vbGreen.  It allows the user to quickly see if the stock has positive or negative returns.  Finally, adding in timer in VBA with the pop-up message shows the elapsed runtime, allowing us to do the time comparison.  

![image](https://user-images.githubusercontent.com/103588178/167333275-79e05c72-be16-4493-b01d-ee463d497ea6.png)

The selected solar and energy stocks show a positive return except TERP in 2017.  The manufacturer and energy provider for stock ticker DQ, ENPH, FSLR and SEDG have raised over 100% return.  The other companies focused on solar power also increased over 20%.   Also, the total daily volume for each stock is all above 35 millions and most of them above 100 millions, which means they have plenty of liquidity. 

![image](https://user-images.githubusercontent.com/103588178/167333367-462d7951-ccfc-4ac3-b67c-f2d37513cc32.png)

In 2018, most selected stocks show negative returns except for a few (ENPH and RUN).  However, the total daily volumes are higher than in 2017.  Therefore, there is no issue with liquidity.  It is the slowdown of the global economy and the oversupply of energy supply.  

## Summary: 
The refactored codes run time for 2017 is 0.87 and 2018 is 0.98, however, the original script ran for 2017 is 0.88 and 2018 is 0.87, the refactored codes have shown 10% more efficient run time.  
#### Pros and Cons of refactoring codes
The pros to refactoring the original VBA scripts to add in index and array to make the codes easier to understand and easier to scale when I continue working with this.  However,  when adding more for loops, it gets confused on the code and increases the risk of error.  
