Sub ticker_summary():
 
' Uses the <date> column to identify where the ticker symbol starts and ends

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentTicker As String
    Dim startDate As Date
    Dim endDate As Date
    Dim openPrice As Variant
    Dim closePrice As Variant
    Dim quarter As Integer
    Dim outputRow As Long
    Dim quarterlyChange As Variant
    Dim percentageChange As Variant
    Dim quarterlySum As Double

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Find the last row of data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize variables
        currentTicker = ws.Cells(2, 1).Value
        startDate = ws.Cells(2, 2).Value
        endDate = ws.Cells(2, 2).Value
        openPrice = ws.Cells(2, 3).Value
        closePrice = ws.Cells(2, 6).Value
        quarterlySum = ws.Cells(2, 7).Value
        quarter = WorksheetFunction.RoundDown((Month(ws.Cells(2, 2).Value) - 1) / 3, 0)
        outputRow = 2

        ' Loop through each row in the current worksheet
        For i = 3 To lastRow
            ' Check if loop is still within the same ticker and quarter
            If ws.Cells(i, 1).Value = currentTicker And _
               quarter = WorksheetFunction.RoundDown((Month(ws.Cells(i, 2).Value) - 1) / 3, 0) Then
                ' Add the value in column G to the quarterly sum
                quarterlySum = quarterlySum + ws.Cells(i, 7).Value
                ' Check if the date is less than the current start date for the ticker
                If ws.Cells(i, 2).Value < startDate Then
                    startDate = ws.Cells(i, 2).Value
                    openPrice = ws.Cells(i, 3).Value
                End If
                ' Check if the date is more than the current maximum date for the ticker
                If ws.Cells(i, 2).Value > endDate Then
                    endDate = ws.Cells(i, 2).Value
                    closePrice = ws.Cells(i, 6).Value
                End If
            Else
                ' Output the ticker and the values for the previous ticker
                ws.Cells(outputRow, 9).Value = currentTicker
                quartelyChange = closePrice - openPrice
                percentageChange = ((closePrice - openPrice) / openPrice)
                ws.Cells(outputRow, 10).Value = Round(quartelyChange, 2)
                ws.Cells(outputRow, 10).NumberFormat = "0.00"
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
               
                ws.Cells(outputRow, 12).Value = quarterlySum
                outputRow = outputRow + 1

                ' Reset variables for the new ticker or quarter
                currentTicker = ws.Cells(i, 1).Value
                startDate = ws.Cells(i, 2).Value
                endDate = ws.Cells(i, 2).Value
                openPrice = ws.Cells(i, 3).Value
                closePrice = ws.Cells(i, 6).Value
                quarter = WorksheetFunction.RoundDown((Month(ws.Cells(i, 2).Value) - 1) / 3, 0)
                quarterlySum = ws.Cells(i, 7).Value
            End If

        Next i

        ' Output the ticker and the values for the minimum and maximum date of the last ticker
               ws.Cells(outputRow, 9).Value = currentTicker
                quartelyChange = closePrice - openPrice
                percentageChange = ((closePrice - openPrice) / openPrice)
                
                ws.Cells(outputRow, 10).Value = Round(quartelyChange, 2)
                ws.Cells(outputRow, 10).NumberFormat = "0.00"
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                
                ws.Cells(outputRow, 12).Value = quarterlySum
                outputRow = outputRow + 1

        ' Column headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    Next ws
    'FindGreatestValues
    'ConditionalFormat

End Sub

