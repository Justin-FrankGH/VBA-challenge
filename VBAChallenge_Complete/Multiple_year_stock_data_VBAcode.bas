Attribute VB_Name = "Module1"
Sub StockLoop():

'designate variables
Dim row As Long
Dim totalStockVolume As Currency
Dim totalsRowCounter As Integer
Dim openPrice As Currency
Dim closePrice As Currency
Dim percentMax As Variant
Dim percentMin As Variant
Dim StockMax As Currency

'Loop through all Sheets
Dim i As Integer
Dim ws_num As Integer
Dim starting_ws As Worksheet

Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count
For i = 1 To ws_num
    ThisWorkbook.Worksheets(i).Activate

'Label Headers
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly_Change"
Cells(1, 12).Value = "Percent_Change"
Cells(1, 13).Value = "StockValue_Total"
Cells(2, 15).Value = "Greatest_%_Increase"
Cells(3, 15).Value = "Greatest_%_Decrease"
Cells(4, 15).Value = "Greatest_Total_Volume"

'Change number type formatting
Range("L:L").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"

'Set starting values
totalStockVolume = 0
totalsRowCounter = 2
openPrice = 0
closePrice = 0
percentMax = 0
percentMin = 0
StockMax = 0

'Loop through Columns
For row = 2 To Range("A2").End(xlDown).row

'define totalStockVolume and percentChange for Loop
    totalStockVolume = totalStockVolume + Cells(row, 8).Value
'store Open Price
    If Cells(row, 1).Value <> Cells(row - 1, 1).Value Then
        openPrice = Cells(row, 3).Value
    End If
    If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
    'Store Close Price
        closePrice = Cells(row, 6).Value
        'Fill in totals Columns
        Cells(totalsRowCounter, 10).Value = Cells(row, 1).Value
        Cells(totalsRowCounter, 11).Value = closePrice - openPrice
        If openPrice <> 0 Then
            Cells(totalsRowCounter, 12).Value = (closePrice - openPrice) / openPrice
        Else
            Cells(totalsRowCounter, 12).Value = 0
        End If
        Cells(totalsRowCounter, 13).Value = totalStockVolume
        totalsRowCounter = totalsRowCounter + 1
        'Reset
        openPrice = 0
        closePrice = 0
        totalStockVolume = 0
    End If

Next row

For row = 2 To Range("A2").End(xlDown).row

    'Challenge problems
    'Greatest Percent Change
    If Cells(row, 12).Value > percentMax Then
        percentMax = Cells(row, 12).Value
        Cells(2, 16).Value = Cells(row, 10).Value
        Cells(2, 17).Value = percentMax
    End If
    'Smallest Percent Change
    If Cells(row, 12).Value < percentMin Then
        percentMin = Cells(row, 12).Value
        Cells(3, 16).Value = Cells(row, 10).Value
        Cells(3, 17).Value = percentMin
    End If
    'Greatest Stock Total
    If Cells(row, 13).Value > StockMax Then
        StockMax = Cells(row, 13).Value
        Cells(4, 16).Value = Cells(row, 10).Value
        Cells(4, 17).Value = StockMax
    End If
    
    'Formatting
    If Cells(row, 11).Value > 0 Then
        Cells(row, 11).Interior.ColorIndex = 4
    ElseIf Cells(row, 11).Value < 0 Then
        Cells(row, 11).Interior.ColorIndex = 3
    End If

Next row

Columns("B:Q").EntireColumn.AutoFit

Next
starting_ws.Activate 'activate the worksheet that was originally active

End Sub
