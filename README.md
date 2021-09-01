# All Stocks Analysis

## Overview of Project
The purpose of this analysis was to refactor the original code in order for improved effeciency and increased functionality.  Our client, Steve, is happy with the orginal code that was written, but he is wanting to expand his analysis to the entire stock market.  The original code is too clunky and slow to run an analysis on that many stocks so the code needs to be refactored to perform the same analysis, but running as quickly as possible.  This optimization will allow Steve to loop through all the tickers in the stock market in a timely manner. 

## Results
The refactored code is as follows:

Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex.
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
     

    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

This subroutine starts with creating an array of all of the known tickers within the spreadsheet.  Arrays for the ticker volumes, starting prices, and ending prices are also created.  A ticker index is created in order to loop through the tickers with increased function.  This ticker index is set to zero.  The bulk of the code consists of three different for loops.  The first loop runs through all the tickers and sets the ticker volumes equal to zero.  The second loop runs through the entire datset and stores all of the data (Ticker Volumes, Starting Prices, Ending Prices) using the ticker index while only looping through the whole dataset once.  This makes the program far more efficient and brings the run time down significantly.  The final loop runs through each array and prints the values into one of the columns in the "All Stocks Analysis" spreadsheet.  The remainder of the code just adds conditional formatting to add increased visual awareness to the finished analysis.  



## Summary
The refactored code offers a couple benefits over the original code.  Since the original code contained a nested for loop, it was much slower.  This is because it had to loop over the entire dataset 11 times for each of the 11 tickers.  With the new refactored code, the program only loops through the entire dataset once so it is far more efficient.  This is very important if the dataset is to be expanded to contain the entire stock market, because with the original code it would have to loop over the dataset for information on each ticker.  So as more tickers are added, the slower the program would be.  The main disadvantage with the refactored code is that there are more variables and lots of opportunity for syntax errors.  The original code is also much simpler to understand than the refactored code.  Other than that, the refactored code is far superior and is optimized for much more functionality when working with larger datasets.






