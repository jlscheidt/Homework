Sub yearlyChange()

  Dim symbol As String

  Dim total_volume As Double
  total_volume = 0

  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  
  Dim ws As Worksheet
  
  Dim startPrice As Double
  startPrice = 0
  Dim endPrice As Double
  endPrice = 0
  Dim priceChange As Double
  priceChange = 0
  
  
  Set ws = ActiveSheet
  
  Cells(1, 9).Value = "Symbol"
  Cells(1, 10).Value = "Total Volume"
  Cells(1, 11).Value = "Yearly Change"
  Cells(1, 12).Value = "Percent Change"
  
  
  lastRow = ws.Cells(rows.Count, 1).End(xlUp).row
  
  startPrice = Cells(2, 3).Value
  
  

  For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      symbol = Cells(i, 1).Value
      
      endPrice = Cells(i, 6).Value
      
      priceChange = endPrice - startPrice
      
      
      Range("K" & Summary_Table_Row).Value = priceChange
      
      Range("L" & Summary_Table_Row).Value = priceChange / startPrice

      
      
      
      
      
      startPrice = Cells(i, 3).Value


      total_volume = total_volume + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = symbol

      Range("J" & Summary_Table_Row).Value = total_volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      total_volume = 0

    Else

      total_volume = total_volume + Cells(i, 7).Value

    End If

  Next i

End Sub

