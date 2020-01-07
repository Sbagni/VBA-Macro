Sub stock_data():

Dim ws As Worksheet
For Each ws In Worksheets
Dim Ticker As String
 Dim Total_volume As Double
  Total_volume = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To LastRow
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               Ticker = ws.Cells(i, 1).Value
                         Total_volume = Total_volume + ws.Cells(i, 7).Value
                         ws.Cells(1, 8).Value = "Ticker"
                         ws.Cells(1, 10).Value = "Total volume"
                         ws.Range("I" & Summary_Table_Row).Value = Ticker
                            ws.Range("j" & Summary_Table_Row).Value = Total_volume
             Summary_Table_Row = Summary_Table_Row + 1
            Total_volume = 0
    Else
    Total_volume = Total_volume + ws.Cells(i, 7).Value
  End If
Next i
Next ws

End Sub

