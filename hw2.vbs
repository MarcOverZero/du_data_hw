Sub tickerAnalysis()
NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count

stock_total = 0
ticker_count = 2

For i = 2 To NumRows
  current_row = Cells(i, 1)
  next_row = Cells(i + 1, 1)

  stock_total = stock_total + Cells(i, 7).Value

  If current_row.Value <> next_row.Value Then
    Cells(ticker_count, 9) = Cells(ticker_count, 1).Value
    Cells(ticker_count, 10) = stock_total
    stock_total = 0
    ticker_count = ticker_count + 1
  End If
  Next i

End Sub
