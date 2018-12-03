Sub StockAnalysis()

  Dim ws As Worksheet
  For Each ws In Worksheets
  ws.Activate
  Dim Ticker As String

  Dim Total_Volume As Double
  Total_Volume = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Total Stock Volume"
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker = Cells(i, 1).Value

      Total_Volume = Total_Volume + Cells(i, 7).Value

      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ws.Range("J" & Summary_Table_Row).Value = Total_Volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      Total_Volume = 0

    Else

      Total_Volume = Total_Volume + Cells(i, 7).Value

    End If
      
  Next i

 Next ws

End Sub
