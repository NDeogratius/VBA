Sub FindGreatestValues()
' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxVolume As Double: maxVolume = 0
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    Dim i As Long
    
    ' declare and instantiate maxIncrease and maxIncrease variables
    Dim maxIncrease As Double: maxIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    Dim maxDecrease As Double: maxDecrease = Application.WorksheetFunction.Min(Range("K:K"))
    
    
    
    ' Loop through all worksheets
        For Each ws In ThisWorkbook.Worksheets
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            
            For i = 2 To lastRow
                If ws.Cells(i, 11).Value = maxIncrease Then
                    maxIncrease = ws.Cells(i, 11).Value
                    tickerIncrease = ws.Cells(i, 9).Value
                End If
                
                If ws.Cells(i, 11).Value = maxDecrease Then
                    maxDecrease = ws.Cells(i, 11).Value
                    
                    tickerDecrease = ws.Cells(i, 9).Value
                End If
                
                If ws.Cells(i, 12).Value > maxVolume Then
                    maxVolume = ws.Cells(i, 12).Value
                    tickerVolume = ws.Cells(i, 9).Value
                    
                End If
            Next i
            
            ' create the row and column titles for the calculated values
        
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
        
            ' Output the results
            ws.Cells(2, 16).Value = tickerIncrease
            ws.Cells(2, 17).Value = maxIncrease
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ws.Cells(3, 16).Value = tickerDecrease
            ws.Cells(3, 17).Value = maxDecrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
            
            ws.Cells(4, 16).Value = tickerVolume
            ws.Cells(4, 17).Value = maxVolume
            ws.Cells(4, 17).NumberFormat = "#,##0"
    Next ws
End Sub
