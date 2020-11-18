# Stock-analysis with Excel VBA
## Overview of Project
Refer to [VBA_Challenge.xlsm](../main/VBA_Challenge.xlsm).

### Purpose
Refactor the Module 2 solution code to loop through all the data one time in order to collect the same information that we did in this module.
Then, we’ll determine whether refactoring our code successfully made the VBA script run faster.

### Analysis and Challenges
- There are 12 different stocks. We have to retrieve the ticker, the total daily volume, and the return on each stock.
- Create the input box, chart headers, ticker array, and activate the worksheet, then write the following the script:
   
      '1a) Create a ticker Index
       Dim tickerIndex As Byte
       tickerIndex = 0

      '1b) Create three output arrays
       Dim tickerVolumes(12) As Long
       Dim tickerStartingPrices(12) As Single
       Dim tickerEndingPrices(12) As Single
    
      '2a) Create a for loop to initialize the tickerVolumes to zero.
       For i = 0 To 11
         tickerVolumes(i) = 0
         tickerStartingPrices(i) = 0
         tickerEndingPrices(i) = 0
       Next i
        
      '2b) Loop over all the rows in the spreadsheet.
       For i = 2 To RowCount
    
      '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
      '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         End If
               
      '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
      '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
         End If
             
       Next i

### Summary
#### The advantages and disadvantages of refactoring code in general:
