# Stock Analysis Using Excel VBA

## Project Overview

### Background:
An initial analysis of a dozen stock returns by year was performed using Excel VBA code to determine which stocks appeared best to invest in. The original code proved satisfactory for the initial goal. However, to expand the analysis to thousands of different stocks, the initial code would be slower and take up more computing power than necessary. More efficient code can be used to create a more efficient analysis that is beneficial when working with a much larger dataset.

### Purpose:
To refactor the original code in order to improve the efficiency or calculation time needed to generate the returns for thousands of stocks.

## Results

### Analysis
#### Process:
To create the new macro, AllStocksAnalysisRefactored(), I first pasted in the initial code from VBA_Challenge VB Script file. The initial code creating the header rows in the output sheet, input message box for year input, and ticker array was already complete. Likewise for ending code setting up the formatting of the output data. These two sections of code were similiar to the code from our Module 2 macro AllStocksAnalysis(). In the middle section starting at comment 1a.) the refactored code was built using the instructions provided. The creation of a tickerIndex variable and output arrays for ticker volumes and starting and ending prices allowed for a quicker analysis that still produced correct results. The refactored code is shown below. The non-refactored code has NOT been included below to save space.

Refactored VBA code:
```
'1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
     
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0

    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Checking if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        '3c) Checking if the current row is the last row with the selected ticker. If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3d Increasing the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        End If
                    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
       
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i
```
#### Outcomes:
The refactored code worked as designed and returned stock returns for both 2017 and 2018 that matched the original macro. The stock returns show a clear downturn in 2018 for most stocks compared to 2017. Of the twelve listed, most had a poor performing 2018 after seeing growth in 2017. Using these two years as a guide, it seems advisable to invest in Enphase Energy (ENPH) and Sunrun Inc (RUN) due to back-to-back years of strong returns for both stocks.



## Summary

### Advantages and Disadvantages of Refactoring Code:
Refactoring code can often times make things more organized and "cleaner." This will help others more easily read or follow through at later dates and can make any debugging easier as well. More efficient code may also lead to faster processing/run times, which is very important when running analyses in large datasets. Yet, refactoring can bring disadvantages at times. New bugs may be created in the refactoring process leading to the breaking of something that was working. Also, the time spent refactoring might not be worth the investment in time, especially if the end-user never notices an improvement.

### Advantages and Disadvantages of Refactoring Stock Analysis Code:

The benefit of quicker processing time was achieved in the stock analysis refactor macro. Run time using the original macro created in Module 2 was around .60 seconds for both 2017 and 2018. The refactored macro decreased run times to a little under .10 seconds for both years. If the number of stocks being analyzed increased from twelve to tens of thousands, this improved processing time would be quit beneficial. However, for this project, the drop from .6 seconds to .10 seconds is inmaterial to the user when just twelve stocks are used. The original code already worked and completed in under a second. In real life, if the user knew they would be only ever need to analyze a couple dozen different stocks, then the effort undertaken to refactor the orignal code would not have been worth the effort.

![VBA_Challenge_2017](https://github.com/bfox87/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018](https://github.com/bfox87/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)
