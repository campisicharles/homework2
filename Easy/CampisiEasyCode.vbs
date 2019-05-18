Sub TotalStockVolume():

'specify all worksheets
For Each ws In Worksheets

'last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'variables
Dim ticker As String
Dim Total_Volume As Double
Total_Volume = 0

'tracking location of stock and volume
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'add columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

'Loop through sheet
For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("J" & Summary_Table_Row).Value = Total_Volume
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Volume = 0
    Else
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub
