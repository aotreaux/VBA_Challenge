Stock Data Analysis

1.  Purpose
The purpose was to utlizie refactoring on a prewritten VBA code in order to analyse various stocks' data from the years 2017 and 2018. Once refactored we can determine if the process was slower or faster and thus helpful to Steve and his parents from an efficiency standpoint.

2.  Results
The first necessary step was to copy the code and create all of the arrays and other formatting specifics to get the VBA_Challenge.xlms functional. The coding was tehn writeen out under each numbered question to maintain a sequential flow of operations.  Below is the code used for refactoring.

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
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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

For 2017 the code ran for .484375 seconds and for 2018 the code ran for .3125 seconds.

3.  Summary
In general, the main advanatges of refactoring is that it can find bugs within the program, makes the overal code easier to understand, and improves the design. The main disadvantage of refactoring is that can be time consuming and you may not have all the information necessary to follow through the refactoring process to the end. With regards to this case, the overall benefit to refactoring the original code was that we saved .5-.7 seconds of time. The original scrpt did not lend as much user funcationality as the refactored script as well.