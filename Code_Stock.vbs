Sub StockMarket()

'Define loop for every sheet
For Each ws In Worksheets

'Define varibles
Dim i As Long
Dim TickerName As String
Dim Price_Open As Double
Price_Open = 0
Dim Ticker_Row As Long
Ticker_Row = 2
Dim Price_Close As Double
Price_Close = 0
Dim Price_Change As Double
Price_Change = 0
Dim Price_Percent As Double
Price_Percent = 0
Dim Total_Volume As Double
Total_Volume = 0
Dim First_Price As Long
First_Price = 2


'Define headers names
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Porcent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Code for summary table
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
Price_Open = ws.Cells(First_Price, 3).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And Price_Open <> 0 Then


TickerName = ws.Cells(i, 1).Value
ws.Range("I" & Ticker_Row).Value = TickerName

Price_Close = ws.Cells(i, 6).Value

Price_Change = Price_Close - Price_Open
ws.Range("J" & Ticker_Row).Value = Price_Change

Price_Percent = (Price_Change / Price_Open)
ws.Range("K" & Ticker_Row).Value = Price_Percent
ws.Range("K" & Ticker_Row).Style = "Percent"
ws.Range("K" & Ticker_Row).NumberFormat = "0.00%"

Total_Volume = Total_Volume + ws.Cells(i, 7).Value
ws.Range("L" & Ticker_Row).Value = Total_Volume

Ticker_Row = Ticker_Row + 1
Price_Percent = 0
Total_Volume = 0
First_Price = i + 1
Price_Open = ws.Cells(First_Price, 3).Value


Else

Total_Volume = Total_Volume + ws.Cells(i, 7).Value

End If

Next i

'Code for conditional formatting
For i = 2 To lastRow

If ws.Cells(i, 10).Value >= 0 Then

ws.Cells(i, 10).Interior.ColorIndex = 4

Else

ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

'Code for min & max
Dim Percent_Last_Row As Long
Percent_Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row

Dim Percent_Max As Double
Percent_Max = 0
Dim Percent_Min As Double
Percent_Min = 0

For i = 2 To Percent_Last_Row



If ws.Cells(i, 11).Value > Percent_Max Then

Percent_Max = ws.Cells(i, 11).Value
ws.Cells(2, 17).Value = Percent_Max
ws.Cells(2, 17).Style = "Percent"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

End If

Next i

For i = 2 To Percent_Last_Row

If ws.Cells(i, 11).Value < Percent_Min Then

Percent_Min = ws.Cells(i, 11).Value
ws.Cells(3, 17).Value = Percent_Min
ws.Cells(3, 17).Style = "Percent"
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

End If

Next i

'Code for greatest total volume
Dim Volume_Last_Row As Long
Volume_Last_Row = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim Volume_Max As Double
Volume_Max = 0

For i = 2 To Volume_Last_Row

If Volume_Max < ws.Cells(i, 12) Then

Volume_Max = ws.Cells(i, 12).Value
ws.Cells(4, 17).Value = Volume_Max
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

End If

Next i

Next ws

End Sub

