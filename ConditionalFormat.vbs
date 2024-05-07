Sub ConditionalFormat():
' declare variables
    Dim ws As Worksheet
    Dim cell As Range
    
' loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    ' iterate through each cell in column J changing the iterior color based on the cell value
        For Each cell In ws.Range("J2:J" & ws.Cells(Rows.Count, 2).End(xlUp).Row)
            If cell.Value < 0 Then
                cell.Interior.ColorIndex = 3
            Else
                cell.Interior.ColorIndex = 4
            End If
        Next cell

    ' iterate through each cell in column K changing the iterior color based on the cell value
        For Each cell In ws.Range("K2:K" & ws.Cells(Rows.Count, 2).End(xlUp).Row)
            If cell.Value < 0 Then
                cell.Interior.ColorIndex = 3
            Else
                cell.Interior.ColorIndex = 4
            End If
        Next cell
    Next ws
End Sub
