Sub multiple_year_stock_data()
    ' Loop through all sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
 
        ' Declare variables
        Dim lastRow As Long
        Dim Ticker As String
        Dim stockVolume As Double
        Dim yearChange As Double
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim percentChange As Double
        Dim summary_table_row As Integer
 
        ' Initialize variables
        summary_table_row = 2
        stockVolume = 0
 
        ' Find last row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
        ' Set opening price
        openingPrice = ws.Cells(2, 3).Value
 
        ' Add headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
 
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we're still within the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker
                Ticker = ws.Cells(i, 1).Value
 
                ' Set closing price
                closingPrice = ws.Cells(i, 6).Value
 
                ' Calculate yearly change
                yearChange = closingPrice - openingPrice
 
                ' Calculate percent change
                If openingPrice <> 0 Then
                    percentChange = yearChange / openingPrice
                Else
                    percentChange = 0
                End If
 
                ' Add to the stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
 
                ' Print the ticker in the summary table
                ws.Range("I" & summary_table_row).Value = Ticker
 
                ' Print the yearly change to the summary table
                ws.Range("J" & summary_table_row).Value = yearChange
 
                ' Color formatting
                If yearChange < 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3 ' Red
                ElseIf yearChange > 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 0 ' No color
                End If
 
                ' Print the percent change to the summary table
                ws.Range("K" & summary_table_row).Value = percentChange
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
 
                ' Print the stock volume to the summary table
                ws.Range("L" & summary_table_row).Value = stockVolume
 
                ' Reset the stock volume
                stockVolume = 0
 
                ' Move to the next row in the summary table
                summary_table_row = summary_table_row + 1
 
                ' Reset the opening price
                openingPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Add to the stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
        Next i
 
        ' Bonus section
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
 
        ' Find last row of the summary table
        Dim lastRow2 As Long
        lastRow2 = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
 
        ' Find greatest increase, decrease, and total volume
        For rr = 2 To lastRow2
            If ws.Cells(rr, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow2)) Then
                ws.Range("P2").Value = ws.Cells(rr, 11).Value
                ws.Range("P2").NumberFormat = "0.00%"
                ws.Range("O2").Value = ws.Cells(rr, 9).Value
            ElseIf ws.Cells(rr, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow2)) Then
                ws.Range("P3").Value = ws.Cells(rr, 11).Value
                ws.Range("P3").NumberFormat = "0.00%"
                ws.Range("O3").Value = ws.Cells(rr, 9).Value
            ElseIf ws.Cells(rr, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow2)) Then
                ws.Range("P4").Value = ws.Cells(rr, 12).Value
                ws.Range("O4").Value = ws.Cells(rr, 9).Value
            End If
        Next rr
    Next ws
End Sub

