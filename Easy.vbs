Sub totalVolume()

  Dim symbol As String

  Dim total_volume As Long
  total_volume = 0

  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  Dim ws As Worksheet
  
  Set ws = ActiveSheet
  
  Cells(1, 9).Value = "Symbol"
  Cells(1, 10).Value = "Total Volume"
  
  lastRow = ws.Cells(rows.Count, 1).End(xlUp).row
  
  

  For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      symbol = Cells(i, 1).Value


      total_volume = total_volume + Cells(i, 3).Value

      Range("I" & Summary_Table_Row).Value = symbol

      Range("J" & Summary_Table_Row).Value = total_volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      total_volume = 0

    Else

      total_volume = total_volume + Cells(i, 3).Value

    End If

  Next i

End Sub
