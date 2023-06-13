Sub Stocks()

Dim i, rows_in_summary_table As Integer
Dim ticker As String
Dim opening_price, closing_price, yearly_change, percent_change As Double
Dim total_stock_volume, last_row As Long

ticker = Cells(2, 1).Value
opening_price = Cells(2, 3).Value
yearly_change = 0
percent_change = 0
total_stock_volume = 0
last_row = Cells(Rows.Count, 1).End(xlUp).Row

rows_in_summary_table = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To last_row
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    closing_price = Cells(i, 6).Value
    yearly_change = closing_price - opening_price
    percent_change = (yearly_change / opening_price)
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
    Range("I" & rows_in_summary_table).Value = ticker
    Range("J" & rows_in_summary_table).Value = yearly_change
    Range("J" & rows_in_summary_table).Style = "Currency"
    Range("K" & rows_in_summary_table).Value = percent_change
    Range("K" & rows_in_summary_table).Style = "Percent"
    Range("K" & rows_in_summary_table).NumberFormat = "0.00%"
    Range("L" & rows_in_summary_table).Value = total_stock_volume
      
      If closing_price > opening_price Then
         Range("J" & rows_in_summary_table).Interior.Color = RGB(0, 255, 0)
         Range("K" & rows_in_summary_table).Interior.Color = RGB(0, 255, 0)
      Else
         Range("J" & rows_in_summary_table).Interior.Color = RGB(255, 0, 0)
         Range("K" & rows_in_summary_table).Interior.Color = RGB(255, 0, 0)
      End If
      
    rows_in_summary_table = rows_in_summary_table + 1
    total_stock_volume = 0
    opening_price = Cells(i + 1, 3).Value
  Else
    total_stock_volume = total_stock_volume + Cells(i, 7)
  End If
Next i

Columns("I:L").AutoFit

End Sub
