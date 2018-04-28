Sub greatestChange()

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
  Dim changePercentage
  
  
  Set ws = ActiveSheet
  
  Cells(1, 9).Value = "Symbol"
  Cells(1, 10).Value = "Total Volume"
  Cells(1, 11).Value = "Yearly Change"
  Cells(1, 12).Value = "Percent Change"
  
  
  Cells(1, 16).Value = "Symbol"
  Cells(1, 17).Value = "Value"
  
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  
  
  
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  startPrice = Cells(2, 3).Value
  
  

  For i = 3 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      symbol = Cells(i, 1).Value
      
      endPrice = Cells(i, 6).Value
      
      priceChange = endPrice - startPrice
      
      
      Range("K" & Summary_Table_Row).Value = priceChange
      
      
      If startPrice <> 0 Then
        changePercentage = priceChange / startPrice
      Else
        changePercentage = 0
      End If
      
      If priceChange > 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      Range("L" & Summary_Table_Row).Value = changePercentage

      
      
      
      
      
      startPrice = Cells(i + 1, 3).Value


      total_volume = total_volume + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = symbol

      Range("J" & Summary_Table_Row).Value = total_volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      total_volume = 0

    Else

      total_volume = total_volume + Cells(i, 7).Value

    End If

  Next i
  
  
  Dim currentHighIncrease As Double
  Dim currentHighIncreaseSymbol As String
  
  Dim currentHighDecrease As Double
  Dim currentHighDecreaseSymbol As String
  
  Dim currentHighVolume As Double
  Dim currentHighVolumeSymbol As String
  
  
  
  currentHighIncrease = 0
  currentHighIncreaseSymbol = Cells(2, 9).Value
  
  currentHighDecrease = Cells(2, 12).Value
  currentHighDecreaseSymbol = Cells(2, 1).Value
  
  currentHighVolume = Cells(2, 10).Value
  currentHighVolumeSymbol = Cells(2, 1).Value
  
  
  For j = 2 To Summary_Table_Row
     If Cells(j, 12).Value > currentHighIncrease Then
        currentHighIncrease = Cells(j, 12).Value
        currentHighIncreaseSymbol = Cells(j, 9).Value
    End If
    If Cells(j, 12).Value < currentHighDecrease Then
        currentHighDecrease = Cells(j, 12).Value
        currentHighDecreaseSymbol = Cells(j, 9).Value
    End If
    
    If Cells(j, 10).Value > currentHighVolume Then
        currentHighVolume = Cells(j, 10).Value
        currentHighVolumeSymbol = Cells(j, 9).Value
    End If
  Next j
  
  Cells(2, 16).Value = currentHighIncreaseSymbol
  Cells(2, 17).Value = currentHighIncrease
  
  Cells(3, 16).Value = currentHighDecreaseSymbol
  Cells(3, 17).Value = currentHighDecrease
  
  Cells(4, 16).Value = currentHighVolumeSymbol
  Cells(4, 17).Value = currentHighVolume
  

End Sub



