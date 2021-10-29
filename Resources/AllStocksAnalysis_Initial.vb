Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?", , , 50, 50)
    
        startTime = Timer

    'Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize an array of all tickers.
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
    
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Activate the data worksheet.
    Sheets(yearValue).Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through the tickers.
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through rows in the data.
        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
            'Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
            'Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                startingPrice = Cells(j, 6).Value
            
            End If
           
            'Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                endingPrice = Cells(j, 6).Value
            
            End If
            
        Next j
        
     'Output the data for the current ticker
     
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
                
    Next i
       'formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A1:C3").Font.Name = "Arial Rounded MT"
    Range("A1:C3").Font.Bold = True
    Range("A1:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A1:C3").Font.Size = 16
    Range("A1:C3").Font.ColorIndex = 18
    Range("C4:C15").NumberFormat = "0.00%"
    Range("B4:B15").NumberFormat = "$0,000.00"
    Columns("B").AutoFit
    
    'conditional formatting return cells
    
    dataRowStart = 4
    dataRowEnd = 15
    
    
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            'color cells green
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            'color cells red
            Cells(i, 3).Interior.Color = vbRed
        
        Else
        'clear the cell color
            Cells(i, 3).Interior.Color = xlNone
            
        End If
    
    Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
End Sub