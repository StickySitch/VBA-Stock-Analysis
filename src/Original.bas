Attribute VB_Name = "Module1"
Sub MacroCheck()
Dim testMessage As String
testMessage = "Hello World"
MsgBox (testMessage)

End Sub

Sub DQAnalysis()
    'Referencing the "DQ Analysis" worksheet
    Worksheets("DQ Analysis").Activate
    
    
    
    'Row headers created
    Range("A1").Value = "DAQO (Ticker: DQ)"
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Referencing "2018" worksheet
    Worksheets("2018").Activate
    
    toalVolume = 0
    rowStart = 2
    'rowEnd code taken from: https://software-solutions-online.com/excel-vba-count-rows-with-data/
    rowEnd = ActiveSheet.UsedRange.Rows.Count
    
    'Instantiating Starting and Closing Price for the year
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Looping through all populated rows
    For i = rowStart To rowEnd
        'Add up the total volume traded for DQ ticker
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
            
        End If
        'Getting the years STARTING price for the DQ ticker
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            startingPrice = Cells(i, 6).Value
            
        End If
        'Getting the years ENDING price for the DQ ticker
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            endingPrice = Cells(i, 6).Value
            
        End If
        
    Next i
    
    
    Worksheets("DQ Analysis").Activate
    'populate cells with 2018 total trade volume for DQ ticker
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    
    'Calculate and asign 2018's return percetage of the DQ ticker
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
End Sub

Sub AllStocksAnalysis()
    
'Instantiate timer objects
    Dim startTime As Single
    Dim endTime As Single
    
'User inputs what year they want to analyze
    yearValue = InputBox("What year would you like to run the analysis on?")
    
'Start timer
    startTime = Timer
        
'Activating correct worksheet
    Worksheets("All Stocks Analysis").Activate

'Formatting the output sheet on "All Stocks Analysis" worksheet
    Range("A1").Value = "All Stocks (" + yearValue + ")."

'Creating row headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'Instantiating ticker array and assigning ticker values
    Dim tickers(11) As String
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
    
    'Instantiating starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
        
        
    'Activate data worksheet
        Worksheets(yearValue).Activate
    
    'Finding number of populated rows
        rowEnd = ActiveSheet.UsedRange.Rows.Count
        
        
    'Looping through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
   
                
        
        'Activating correct worksheet
            Worksheets(yearValue).Activate
        
        'Looping through rows in the data
            For j = 2 To rowEnd
            
            'Getting total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
            
            'Getting starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
            
            'Getting ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
        
            Next j
        'Outputting data for current ticker
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
            
        
        Next i
        
'Activating correct worksheet
    Worksheets("All Stocks Analysis").Activate

'Header formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
'Row formatting
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    Columns("C").AutoFit

'Initializing rowStart and rowEnd values
    rowStart = 4
    rowEnd = ActiveSheet.UsedRange.Rows.Count


'Looping through each row
For i = rowStart To rowEnd
    
    'Changing positive return cell colors to green
        If Cells(i, 3).Value > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        
    'Changing negative return cell colors to red
        ElseIf Cells(i, 3).Value < 0 Then
        
            Cells(i, 3).Interior.Color = vbRed
            
    'Changing no return cell colors to green
        Else
            Cells(i, 3).Interior.Color = xlNone
        
        End If
    

Next i
            
        'End timer
            endTime = Timer
        'Displays how long the analysis took
            MsgBox "Analysis complete in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub


Sub ClearCells()
Cells.Clear
End Sub


