Sub stocktest()

  ' Variables for looping through multiple sheets in one workbook
  Dim iIndex as Integer
  Dim ws as Excel.Worksheet

  ' Set an initial variable for holding the ticker name
  Dim ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim tickerTotal As Double
  tickerTotal = 0

  ' Set an initial variable for the value of ticker at year start
  Dim tickerYearOpen As Double
  tickerYearOpen = 0

  ' Set an initial variable for the value of ticker at year close
  Dim tickerYearClose As Double
  tickerYearClose = 0

  ' Keep track of the location for each ticker in the summary table
  Dim summaryTable As Integer
  summaryTable = 2

For iIndex = 1 to ActiveWorkbook.Worksheets.count
Set ws = Worksheets(iIndex)
ws.Activate

  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"


  ' Loop through all stocks
    For i = 2 To LastRow

    ' Check if we are still within the same stock, if it is not...
    
    ' This statement stores the first closing value of the year for the stock
    If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
        tickerYearOpen = Cells(i, 6).Value
    End If

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set ticker name
      ticker = ws.Cells(i, 1).Value
      
      'Stores the last closing value of the year for the stock
      tickerYearClose = ws.Cells(i, 6).Value

      ' Add to ticker stock total
      tickerTotal = tickerTotal + ws.Cells(i, 7).Value

      ' Print the ticker name in summary table
      ws.Range("I" & summaryTable).Value = ticker

      ' Print the total stock volume in the summary table
      ws.Range("L" & summaryTable).Value = tickerTotal

      ' Print Yearly Change in the summary table
      ws.Range("J" & summaryTable).Value = tickerYearClose - tickerYearOpen

      ' Print Percent Change in the summary table. This also formats the color of the Yearly Change cell
      If tickerYearClose = 0 And tickerYearOpen = 0 Then
        ws.Range("K" & summaryTable).Value = "0%" ' There was no change
        ws.Range("J" & summaryTable).Interior.ColorIndex = 2
      ElseIf tickerYearClose = 0 Then
        ws.Range("K" & summaryTable).Value = "-100%" ' The stock lost all value
        ws.Range("J" & summaryTable).Interior.ColorIndex = 3
      Else
        ws.Range("K" & summaryTable).Value = ((tickerYearClose - tickerYearOpen) / tickerYearClose) * 100 & "%"
        If ws.Range("K" & summaryTable).Value < 0 Then
          ws.Range("J" & summaryTable).Interior.ColorIndex = 3
        Else
        ws.Range("J" & summaryTable).Interior.ColorIndex = 4
        End If

      End If
      
      ' Add one to the summary table row
      summaryTable = summaryTable + 1
      
      ' Reset the stock volume total
      tickerTotal = 0

    ' If the cell immediately following a row is the same stock...
    Else

      ' Add to the stock volume total
      tickerTotal = tickerTotal + ws.Cells(i, 7).Value

    End If

    Next i

summaryTable = 2
Next iIndex

End Sub