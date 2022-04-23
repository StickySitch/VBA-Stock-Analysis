Attribute VB_Name = "Module2"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    'Getting user input to select which year to analyze
        yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Starting timer
        startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
    
    'Header that lets the user know what years analysis they are looking at.
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Creating row headers
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
        Dim tickers(12) As String
        
    'Assigning ticker values to "tickers" array
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
    
    'Activate data worksheet for the specified year
        Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Creating ticker Index
        Dim tickerIndex As Integer
    

    'Creating three output arrays
        Dim tickerVolumes() As Long
        Dim tickerStartingPrice(12) As Single
        Dim tickerEndingPrice(12) As Single
    'Initializing "tickerIndex" to 0
        tickerIndex = 0
    
    
    'Looping through "tickerVolumes" array and initializing array variables to 0
    For i = 0 To 11
    
        'Resizing "tickerVolumes" array to the correct size
            ReDim tickerVolumes(i)
            
        'Initializing "tickerVolumes" to 0
            tickerVolumes(i) = 0
        
    Next i
    
     
    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'Increase volume traded for current ticker in "tickers" array
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                'Initializes "tickerStartingPrice" to the tickers starting price for the year
                    tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
                    
            End If
            
        
        
        'Checks if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                'Initializes "tickerEndingPrice" to the tickers ending price for the year
                    tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            
                'Increases the tickerIndex
                    tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    'Loops through your arrays to output the Ticker, Total Daily Trade Volume, and Return.
    For i = 0 To 11
        'Activating the correct output worksheet
            Worksheets("All Stocks Analysis").Activate
            tickerIndex = i
        'Populating cells with ticker symbols, Ticker Daily Trade Volumes, and yearly return
            Cells(4 + i, 1).Value = tickers(tickerIndex)
            Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
            Cells(4 + i, 3).Value = tickerEndingPrice(tickerIndex) / tickerStartingPrice(tickerIndex) - 1
        
    Next i
    
    
    Worksheets("All Stocks Analysis").Activate
    'Formatting columns
    'Making column headers bold
    Range("A3:C3").Font.FontStyle = "Bold"
    'Underlining the headers
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Formatting Total Daily Volume
    Range("B4:B15").NumberFormat = "#,##0"
    'Formatting return percentage
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    
    'Changing return color format
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
        
        'Makes cell green if the return percentage is positive
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
        'Makes cell green if the return percentage is positive
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
