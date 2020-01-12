Attribute VB_Name = "Module1"
Sub VBAStocks():

Dim wsMySheet As Worksheet
Application.ScreenUpdating = False
For Each wsMySheet In ThisWorkbook.Sheets
wsMySheet.Select
 
    Cells(1, 10) = "Ticker"
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percent Change"
    Cells(1, 13) = "Total Stock Volume"

    Dim row As LongPtr
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim percentChange As Double
    Dim totalStock As LongPtr
    Dim stockRowCount As LongPtr
    Dim formatRow As LongPtr
    Dim resultRow As LongPtr
    Dim volumeRow As LongPtr
    Dim max As Double
    Dim min As Double
    Dim volumeMax As Double
    Dim tickerMax As String
    Dim tickerMin As String
    Dim tickerVolumeMax As String

    totalStock = 0
    openPrice = 0
    closePrice = 0
    yearChange = 0
    percentChange = 0
    stockRowCount = 2

    For row = 2 To Range("A2").End(xlDown).row
    totalStock = totalStock + Cells(row, 7).Value


    If Cells(row - 1, 1).Value <> Cells(row, 1).Value Then
    openPrice = Cells(row, 3).Value
    End If

    If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
    Cells(stockRowCount, 10) = Cells(row, 1).Value
    closePrice = Cells(row, 6).Value
    yearChange = closePrice - openPrice
    Cells(stockRowCount, 11) = yearChange
    If (openPrice = 0 And closePrice = 0) Then
    percentChange = 0
    ElseIf (openPrice = 0 And closePrice <> 0) Then
    percentChange = 1
    Else
    percentChange = yearChange / openPrice
    Cells(stockRowCount, 12).Value = percentChange
    Cells(stockRowCount, 12).NumberFormat = "0.00%"
    End If
    Cells(stockRowCount, 13) = totalStock
    stockRowCount = stockRowCount + 1
    totalStock = 0
    openPrice = 0
    closePrice = 0
    yearChange = 0
    percentChange = 0
    End If
    Next row

    For formatRow = 2 To Range("K2").End(xlDown).row
    If Cells(formatRow, 11).Value >= 0 Then
    Cells(formatRow, 11).Interior.ColorIndex = 4
    Else
    Cells(formatRow, 11).Interior.ColorIndex = 3
    End If
   
    Next formatRow
    
    Cells(1, 17) = "Ticker"
    Cells(1, 18) = "Value"
    Cells(2, 16) = "Greatest% Increase"
    Cells(3, 16) = "Greatest% Decrease"
    Cells(4, 16) = "Greatest Total Volume"

       
    max = Application.WorksheetFunction.max(Columns("L"))
    min = Application.WorksheetFunction.min(Columns("L"))
    volumeMax = Application.WorksheetFunction.max(Columns("M"))

    For resultRow = 2 To Range("L2").End(xlDown).row
    If Cells(resultRow, 12) = max Then
    tickerMax = Cells(resultRow, 10)
    End If

    If Cells(resultRow, 12) = min Then
    tickerMin = Cells(resultRow, 10)
    End If

    Next resultRow

    For volumeRow = 2 To Range("M2").End(xlDown).row
    If Cells(volumeRow, 13) = volumeMax Then
    tickerVolumeMax = Cells(volumeRow, 10)
    End If
    Next volumeRow

    Cells(2, 18).Value = max
    Cells(2, 18).NumberFormat = "0.00%"
    Cells(3, 18).Value = min
    Cells(3, 18).NumberFormat = "0.00%"
    Cells(4, 18).Value = volumeMax
    Cells(2, 17).Value = tickerMax
    Cells(3, 17).Value = tickerMin
    Cells(4, 17).Value = tickerVolumeMax

Next wsMySheet
 
Application.ScreenUpdating = True

End Sub



