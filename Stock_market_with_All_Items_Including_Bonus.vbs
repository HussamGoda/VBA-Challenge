Sub Stocks()

Dim i, rows_in_summary_table As Integer
Dim ticker As String
Dim opening_price, closing_price, yearly_change, percent_change As Double
Dim total_stock_volume, last_row As Double
Dim max_percent_change, min_percent_change As Double
Dim max_stoch_volume As Double
Dim counter, sheets_number As Integer

sheets_number = Application.Worksheets.count

For counter = 1 To sheets_number
Worksheets(counter).Activate

ticker = ActiveSheet.Cells(2, 1).Value
opening_price = ActiveSheet.Cells(2, 3).Value
yearly_change = 0
percent_change = 0
total_stock_volume = 0
last_row = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row

rows_in_summary_table = 2

ActiveSheet.Cells(1, 9).Value = "Ticker"
ActiveSheet.Cells(1, 10).Value = "Yearly Change"
ActiveSheet.Cells(1, 11).Value = "Percent Change"
ActiveSheet.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To last_row
  If ActiveSheet.Cells(i + 1, 1).Value <> ActiveSheet.Cells(i, 1).Value Then
    ticker = ActiveSheet.Cells(i, 1).Value
    closing_price = ActiveSheet.Cells(i, 6).Value
    yearly_change = closing_price - opening_price
    percent_change = (yearly_change / opening_price)
    total_stock_volume = total_stock_volume + ActiveSheet.Cells(i, 7).Value
    ActiveSheet.Range("I" & rows_in_summary_table).Value = ticker
    ActiveSheet.Range("J" & rows_in_summary_table).Value = yearly_change
    ActiveSheet.Range("J" & rows_in_summary_table).Style = "Currency"
    ActiveSheet.Range("K" & rows_in_summary_table).Value = percent_change
    ActiveSheet.Range("K" & rows_in_summary_table).Style = "Percent"
    ActiveSheet.Range("K" & rows_in_summary_table).NumberFormat = "0.00%"
    ActiveSheet.Range("L" & rows_in_summary_table).Value = total_stock_volume
      
      If closing_price > opening_price Then
         ActiveSheet.Range("J" & rows_in_summary_table).Interior.Color = RGB(0, 255, 0)
         ActiveSheet.Range("K" & rows_in_summary_table).Interior.Color = RGB(0, 255, 0)
      Else
         ActiveSheet.Range("J" & rows_in_summary_table).Interior.Color = RGB(255, 0, 0)
         ActiveSheet.Range("K" & rows_in_summary_table).Interior.Color = RGB(255, 0, 0)
      End If
      
    rows_in_summary_table = rows_in_summary_table + 1
    total_stock_volume = 0
    opening_price = ActiveSheet.Cells(i + 1, 3).Value
  Else
    total_stock_volume = total_stock_volume + ActiveSheet.Cells(i, 7).Value
  End If
Next i

ActiveSheet.Columns("I:L").AutoFit

ActiveSheet.Range("P1").Value = "Ticker"
ActiveSheet.Range("Q1").Value = "Value"
ActiveSheet.Range("O2").Value = "Greatest % Increase"
ActiveSheet.Range("O3").Value = "Greatest % Decrease"
ActiveSheet.Range("O4").Value = "Greatest Total Vollume"

last_row = ActiveSheet.Range("I" & Rows.count).End(xlUp).Row

max_percent_change = 0
For i = 2 To last_row
    If ActiveSheet.Range("K" & i).Value >= max_percent_change Then
       max_percent_change = ActiveSheet.Range("K" & i).Value
       ActiveSheet.Range("P2").Value = ActiveSheet.Range("I" & i).Value
       ActiveSheet.Range("Q2").Value = max_percent_change
    End If
Next i

 ActiveSheet.Range("Q2").Value = max_percent_change
 ActiveSheet.Range("Q2").Style = "Percent"
 ActiveSheet.Range("Q2").NumberFormat = "0.00%"
 
 
min_percent_change = 0
For i = 2 To last_row
    If ActiveSheet.Range("K" & i).Value < min_percent_change Then
       min_percent_change = ActiveSheet.Range("K" & i).Value
       ActiveSheet.Range("P3").Value = ActiveSheet.Range("I" & i).Value
       ActiveSheet.Range("Q3").Value = min_percent_change
    End If
Next i

 ActiveSheet.Range("Q3").Value = min_percent_change
 ActiveSheet.Range("Q3").Style = "Percent"
 ActiveSheet.Range("Q3").NumberFormat = "0.00%"
 

max_stoch_volume = 0
For i = 2 To last_row
    If Range("L" & i).Value >= max_stoch_volume Then
       max_stoch_volume = ActiveSheet.Range("L" & i).Value
       ActiveSheet.Range("P4").Value = ActiveSheet.Range("I" & i).Value
       ActiveSheet.Range("Q4").Value = max_stoch_volume
    End If
Next i
 
 ActiveSheet.Columns("O:Q").AutoFit

Next counter

Worksheets("2018").Activate


 
End Sub
