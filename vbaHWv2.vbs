Option Explicit
Sub TickerCalculation()

'Data
'Col 1 - Ticker
'Col 2 - Date
'Col 3 - Open
'Col 4 - High
'Col 5 - low
'Col 6 - Close
'Col 7 - Volume
'Col 8 - Volume in 1000's

'Output
'Col 9 - Ticker
'Col 10 - Yearly Change
'Col 11 - Percent Change
'Col 12 - Total Stock Volume
Dim dataRow As Long
Dim outputRow As Long
Dim sheetNum As Long


'I know this is first ticker, first row
'therefore, save the open price
'Create a counter for total stock volume

Dim openPrice As Double
Dim totalStockVolume As Double
Dim closePrice As Double


For sheetNum = 1 To Worksheets.Count

    Worksheets(sheetNum).Activate

    outputRow = 2

    openPrice = ActiveSheet.Range("C2").Value

    'Start loop at A2
    For dataRow = 2 To Range("A2").End(xlDown).Row
        If ActiveSheet.Cells(dataRow, 1).Value <> ActiveSheet.Cells(dataRow + 1, 1).Value Then
        'Now at the edge
        'add whatever is in Col G to the total stock volume counter
        totalStockVolume = totalStockVolume + Cells(dataRow, 7) / 1000
        'then grab the closing price from Col F
        closePrice = ActiveSheet.Cells(dataRow, 6).Value
        'Now calculate the yearly change as close_price - open_price
        'Calculate yearly percent change as close_price - open_price / open_price
        'Since there might be a division by 0, put in a check so that denominator is not 0
        'Copy over the value in Col A to Col I
        'Then dump the yearly change, percent change and total stock volume to J,K,L

        'Ticker
        Range("I1").Value = "Ticker"
        ActiveSheet.Cells(outputRow, 9).Value = ActiveSheet.Cells(dataRow, 1).Value

        'Yearly Change
        Range("J1").Value = "Yearly Change"
        ActiveSheet.Cells(outputRow, 10).Value = closePrice - openPrice
        
        'Percent Change
        If openPrice = 0 Then
            ActiveSheet.Cells(outputRow, 11).Value = "NaN"
        Else
            ActiveSheet.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
        End If

        'Percent Change Column Formatting
        Range("K1").Value = "Percent Change"
        Range("K1").End(xlDown).NumberFormat = "0.00%"

        'Total Stock Volume
        Range("L1").Value = "Total Stock Volume"
        ActiveSheet.Cells(outputRow, 12).Value = totalStockVolume * 1000

        'Coloring Positive or Negative Change
        If ActiveSheet.Cells(outputRow, 10).Value < 0 Then
            ActiveSheet.Cells(outputRow, 10).Interior.ColorIndex = 3
        Else 
            ActiveSheet.Cells(outputRow, 10).Interior.ColorIndex = 4
        End If

        'Column Fit
        Columns("A:M").AutoFit

        'Add 1 to the row counter for the output table
        outputRow = outputRow + 1

        'Then update the new open price to be the open price of the next row
        totalStockVolume = 0
        openPrice = ActiveSheet.Cells(dataRow + 1, 3).Value
        Else
            'If it's not the edge, then
            'don't change the open value
            'add whatever is in Col G to the total stock volume counter
            totalStockVolume = totalStockVolume + ActiveSheet.Cells(dataRow, 7) / 1000
        End If
    Next dataRow

Next sheetNum

End Sub
