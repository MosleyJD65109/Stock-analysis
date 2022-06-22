# Stock Analysis
## Overview of Project:
### Purpose
The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.
##  Analysis and Challenges:

### Analysis of Total Daily Volume and Return

<img width="250" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104540261/174865066-455ffde2-cd7a-4607-ab23-ec29175d395d.png">

<img width="188" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104540261/174865124-44c203e3-bb2d-4d52-94a7-172eebcf5ffb.png">

The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.
### Challenges and Difficulties Encountered
To setup the new refactored code I had to copy and paste the starter code. This gave me the framework I needed to create the ticker array, chart headers, and the input box. Each of the steps were detailed in the comment sections of the code as you can see listed below.

[All Stocks Analysis Refactored.txt](https://github.com/MosleyJD65109/stock-analysis/files/8951483/All.Stocks.Analysis.Refactored.txt)



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
    
     Dim tickerIndex As Single
     tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row's ticker doesn't match, increase the tickerIndex
    
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
        'If Then
        
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If
        
        
        '3c) Check if the current row is the last row with the selected ticker
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
    
     Cells(i, 3).Interior.Color = xlNone
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub



## Results:
<img width="205" alt="VBA_Challenge_2017_runtime" src="https://user-images.githubusercontent.com/104540261/174933132-916730c0-27b6-471a-852f-f3cd79f9397d.png">


<img width="194" alt="VBA_Challenge_2018_runtime" src="https://user-images.githubusercontent.com/104540261/174933181-82d581ef-1c7d-40e9-8949-c1409fca46a4.png">

