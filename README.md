# VBA Challenge
 **Table of content:**
 - [Subroutine: ticker_summary](#item-one)
 - [Subroutine: ConditionalFormat](#item-two)
 - [Subroutine: FindGreatestValues](#item-three)
 
This challenge is guided by the following instructions
- Create a script that loops through all the stocks for each quarter and outputs the following information:
- The ticker symbol
- Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
- The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
- The total stock volume of the stock.

The output of the VBA script(s) should look similar to this immage
![alt text](image.png)

The workbook consists of 4 worksheets with similar data and layout as shown in the screenshot below
![alt text](image-1.png)


The Task is broken down into three VBA scripts as explained below
## Subroutine: ticker_summary
```vb
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

```

This subroutine performs a summary of stock data by ticker symbol and quarter.

### Variables

- `ws`: Represents the current worksheet.
- `lastRow`: Stores the last row of data in the current worksheet.
- `i`: Loop counter.
- `currentTicker`: Stores the current ticker symbol.
- `startDate`, `endDate`: Store the start and end dates for the current ticker.
- `openPrice`, `closePrice`: Store the opening and closing prices for the current ticker.
- `quarter`: Stores the current quarter.
- `outputRow`: Stores the row number for output.
- `quarterlyChange`: Stores the change in price over the quarter for the current ticker.
- `percentageChange`: Stores the percentage change in price over the quarter for the current ticker.
- `quarterlySum`: Stores the sum of the volume of the current ticker for the quarter.

### Process

1. The subroutine loops through each worksheet in the workbook.
2. For each worksheet, it finds the last row of data and initializes the variables.
3. It then loops through each row in the worksheet.
4. If the ticker and quarter are the same as the current ticker and quarter, it updates the quarterly sum and checks if the date is less or more than the current start and end dates for the ticker.
5. If the ticker or quarter changes, it outputs the ticker and the values for the previous ticker and resets the variables for the new ticker or quarter.
6. Finally, it outputs the ticker and the values for the minimum and maximum date of the last ticker and sets the column headers for the summary table.

### Output

Once the script runs, each worksheet will appear as shown in the following screenshot
![alt text](image-2.png)


## Subroutine: ConditionalFormat
```vb
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

```
This subroutine applies conditional formatting to the summary table. It changes the cell color based on whether the value is less than 0.

### Output
The output of the conditionalFormat subroutine is shown in the screenshoot below
![alt text](image-3.png)


## Subroutine: FindGreatestValues

```vb
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
```
This subroutine finds the greatest percentage increase, greatest percentage decrease, and greatest total volume among all tickers and outputs the results.

### Output
When tSubroutine: FindGreatestValues runs successfully each worksheet appears as below

![alt text](image-4.png)
