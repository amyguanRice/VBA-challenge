Sub Stock_loops()
    Dim Stock_Count As Double
    Dim Stock_Opening As Double
    Dim Stock_Closing As Double
    Dim Stock_Change As Double
    Dim Stock_PctChange As Double
    Dim Stock_TotVol As Double
    Dim LastRow As Double
    Dim Ticker_MaxPctInc As String
    Dim Ticker_MaxPctDec As String
    Dim Ticker_MaxTotVol As String
    Dim Value_MaxPctInc As Double
    Dim Value_MaxPctDec As Double
    Dim Value_MaxTotVol As Double
      
    For Each ws In Worksheets
        
        ' write the header line
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Volume"
        
        ' Count the number of rows in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the Stock Info
        Stock_Count = 0
        Stock_Opening = ws.Cells(2, 3).Value
        Stock_Ticker = ws.Cells(2, 1).Value
        Stock_TotVol = ws.Cells(2, 7).Value
        
        ' write the initial info into first row
        ws.Cells(2, 9).Value = Stock_Ticker
        
        ' Now loop through all rows
        For irow = 3 To Int(LastRow)
            
            Stock_Ticker_temp = ws.Cells(irow, 1).Value
            Stock_TotVol = Stock_TotVol + ws.Cells(irow, 7).Value
            
            If StrComp(Stock_Ticker_temp, Stock_Ticker) Then
                
                Stock_Closing = ws.Cells(irow - 1, 6).Value
                
                Stock_Change = Stock_Closing - Stock_Opening
                
                If Stock_Opening = 0 Then
                    Stock_PctChange = 0
                Else
                    Stock_PctChange = Stock_Change / Stock_Opening
                End If
                
                ' now write the Ticker, Yearly Change, % Change and TotVol
                
                ws.Cells(Stock_Count + 2, 9) = Stock_Ticker
                ws.Cells(Stock_Count + 2, 10) = Stock_Change
                ws.Cells(Stock_Count + 2, 11) = Stock_PctChange
                ws.Cells(Stock_Count + 2, 11).NumberFormat = "0.00%"
                ws.Cells(Stock_Count + 2, 12) = Stock_TotVol
                
                'ws.Cells(Stock_Count + 2, 13) = Stock_Opening
                'ws.Cells(Stock_Count + 2, 14) = Stock_Closing
                
                ' mark positive change in green and negative in red
                
                If Stock_Change > 0 Then
                    ws.Cells(Stock_Count + 2, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Stock_Count + 2, 10).Interior.ColorIndex = 3
                End If
                                         
                Stock_Count = Stock_Count + 1
                
                Stock_Ticker = ws.Cells(irow, 1).Value
                
                Stock_Opening = ws.Cells(irow, 3).Value
                
                Stock_TotVol = ws.Cells(irow, 7).Value
                           
            ' Close the If/Else Statement
            End If
        
        Next irow
        
        ' add the last stock to the summary table
        ws.Cells(Stock_Count + 2, 9) = Stock_Ticker
        
        Stock_Closing = ws.Cells(LastRow, 6).Value
        Stock_Change = Stock_Closing - Stock_Opening
        Stock_PctChange = Stock_Change / Stock_Opening
        
        ws.Cells(Stock_Count + 2, 10) = Stock_Change
        ws.Cells(Stock_Count + 2, 11) = Stock_PctChange
        ws.Cells(Stock_Count + 2, 11).NumberFormat = "0.00%"
        ws.Cells(Stock_Count + 2, 12) = Stock_TotVol
        
        'ws.Cells(Stock_Count + 2, 13) = Stock_Opening
        'ws.Cells(Stock_Count + 2, 14) = Stock_Closing
               
        If Stock_Change > 0 Then
            ws.Cells(Stock_Count + 2, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(Stock_Count + 2, 10).Interior.ColorIndex = 3
        End If
        
        ' ------------------------------------------------------------
        ' Now loop through the summary table to output the stocks with max %inc, max %dec and max TotVol
        
        LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' initialize the array
        
        Ticker_MaxPctInc = ws.Cells(2, 9).Value
        Ticker_MaxPctDec = ws.Cells(2, 9).Value
        Ticker_MaxTotVol = ws.Cells(2, 9).Value
        Value_MaxPctInc = ws.Cells(2, 11).Value
        Value_MaxPctDec = ws.Cells(2, 11).Value
        Value_MaxTotVol = ws.Cells(2, 12).Value
        
        ' loop through the summary table
        
        For jrow = 3 To Int(LastSummaryRow)
            
            If ws.Cells(jrow, 11) > Value_MaxPctInc Then
                Ticker_MaxPctInc = ws.Cells(jrow, 9)
                Value_MaxPctInc = ws.Cells(jrow, 11)
            End If
            
            If ws.Cells(jrow, 11) < Value_MaxPctDec Then
                Ticker_MaxPctDec = ws.Cells(jrow, 9)
                Value_MaxPctDec = ws.Cells(jrow, 11)
            End If
            
            If ws.Cells(jrow, 12) > Value_MaxTotVol Then
                Value_MaxTotVol = ws.Cells(jrow, 12)
            End If
            
        Next jrow
        
        ' output the finalized three stocks
        
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(2, 16) = Ticker_MaxPctInc
        ws.Cells(3, 16) = Ticker_MaxPctDec
        ws.Cells(4, 16) = Ticker_MaxTotVol
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 17) = Value_MaxPctInc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17) = Value_MaxPctDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17) = Value_MaxTotVol
        
    Next ws
        
End Sub

