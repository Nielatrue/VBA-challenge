 Private Sub easy()


 Dim ticker As String
 Dim vol As Double
 Dim Summary_Table_Row As Double
 Dim year_open As Double
 Dim year_close As Double
 
 year_open = Cells(2, 3)
  vol = 0
   Summary_Table_Row = 2
 
 Cells(1, 9).Value = "ticker"
 Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Yearly Percentage"
 Cells(1, 12).Value = "Total Stock Vol"
 
 
 For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
  vol = vol + Cells(i, 7)
  ticker = Cells(i, 1)
  
  
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  Cells(Summary_Table_Row, "I") = ticker
  
  year_close = Cells(i, 6).Value
 
  Cells(Summary_Table_Row, "J") = year_close - year_open
 
 If year_open = 0 Then
 Cells(Summary_Table_Row, "K") = Null
 
 Else
 
 Cells(Summary_Table_Row, "K") = FormatPercent((year_close - year_open) / year_open, 2)
 
 
 
 End If
 
 If Cells(Summary_Table_Row, "J") > 0 Then
 Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
 Else: Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
 
 End If
 
 
 Cells(Summary_Table_Row, "L") = vol
 
 'If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 'year_close = Cells(i, 6).Value
 
 
'yearly_change = year_close - year_open
 'yearl_change = Cells(i, 6) - Cells(i, 3)
 
 
 'ticker = Cells(i, 1).Value
' vol = vol + Cells(i, 7).Value
 'Range("J" & Summary_Table_Row).Value = yearly_change
 'Range("I" & Summary_Table_Row).Value = ticker
 'Range("K" & Summary_Table_Row).Value = year_percent
 'Range("L" & Summary_Table_Row).Value = vol
 
 Summary_Table_Row = Summary_Table_Row + 1
 
 vol = 0
 
 year_open = Cells(i + 1, 3)
 
 'Else
 'vol = vol + Cells(i, 7).Value
 
 End If
 Next i
 




 End Sub

