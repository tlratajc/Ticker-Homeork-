Sub stock_ticker()

  ' Set an initial variable for holding the brand name
  Dim ticker As String

  ' Set an initial variable for holding the total ticker
  Dim ticker_total As Double
  ticker_total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim summary_table_row As Integer
  summary_table_row = 2

  ' Loop through all ticker symbols
  For i = 2 To 800000
  

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      ticker_Name = Cells(i, 1).Value

      ' Add to the ticker Total
      ticker_total = ticker_total + Cells(i, 7).Value

      ' Print the ticker name in the Summary Table
      Range("j" & summary_table_row).Value = ticker_Name

      ' Print the ticker Amount to the Summary Table
      Range("k" & summary_table_row).Value = ticker_total

      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Reset the ticker Total
      ticker_total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the ticker Total
      ticker_total = ticker_total + Cells(i, 7).Value

    End If

  Next i

End Sub


