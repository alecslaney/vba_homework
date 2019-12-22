Sub vbaChallenge()

  ' Variables for looping through multiple sheets in one workbook
  Dim iIndex As Integer
  Dim ws As Excel.Worksheet

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

  ' Set an initial variable for percent change calculation
  Dim percentChange As Double
  percentChange = 0

  ' Keep track of the location for each ticker in the summary table
  Dim summaryTable As Integer
  summaryTable = 2

  ' Set an initial variable for storing largest stock volume total
  Dim largestTotal As Double
  largestTotal = 0
  
  ' Stores accompanying ticker
  Dim largestTicker As String

  ' Sets initial variables for storing largest positive and negative percent changes
  Dim largestPosChange as Double
  largestPosChange = 0
  Dim largestNegChange as Double
  largestNegChange = 0

  ' Stores accompanying tickers
  Dim largestPosTicker as String
  Dim largestNegTicker as String

' This runs the macro on each worksheet in the workbook
For iIndex = 1 To ActiveWorkbook.Worksheets.Count
Set ws = Worksheets(iIndex)
ws.Activate

  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Creates headers on each sheet
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"

  ' Loop through all stocks
    For i = 2 To LastRow

    ' This statement stores the first opening value of the year for the stock
    If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
        tickerYearOpen = Cells(i, 3).Value
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

      ' Print Percent Change in the summary table.
      If tickerYearClose = 0 And tickerYearOpen = 0 Then
        ws.Range("K" & summaryTable).Value = "0%" ' There was no change: prints 0% accordingly
        ws.Range("J" & summaryTable).Interior.ColorIndex = 2
      ElseIf tickerYearClose = 0 Then
        ws.Range("K" & summaryTable).Value = "-100%" ' The stock lost all value: prints 100% loss
        ws.Range("J" & summaryTable).Interior.ColorIndex = 3
      Else
        percentChange = ((tickerYearClose - tickerYearOpen) / tickerYearClose) * 100
        ws.Range("K" & summaryTable).Value = percentChange & "%" ' Standard percent change calculation
        
        'Determines and stores largest percent increase/decrease in summary table
        If percentChange > 0 Then
          If percentChange > largestPosChange Then
          largestPosChange = percentChange
          largestPosTicker = ticker
          ws.Range("P2").Value = largestPosTicker
          ws.Range("Q2").Value = largestPosChange & "%"
          End If
        End If

        If percentChange < 0 Then
          If percentChange < largestNegChange Then
          largestNegChange = percentChange
          largestNegTicker = ticker
          ws.Range("P3").Value = largestNegTicker
          ws.Range("Q3").Value = largestNegChange & "%"
          End If
        End If

        If ws.Range("K" & summaryTable).Value < 0 Then
          ws.Range("J" & summaryTable).Interior.ColorIndex = 3
        Else
        ws.Range("J" & summaryTable).Interior.ColorIndex = 4
        End If

      End If
      
      ' Add one to the summary table row
      summaryTable = summaryTable + 1
      
      ' Checks if the total stock volume is the largest summed so far. If it is, update the tracker.
      If tickerTotal > largestTotal Then
      largestTotal = tickerTotal
      largestTicker = ticker
      ws.Range("P4").Value = largestTicker
      ws.Range("Q4").Value = largestTotal
      End If
      
      ' Reset the stock volume total
      tickerTotal = 0

    ' If the cell immediately following a row is the same stock:
    Else
      ' Add to the stock volume total
      tickerTotal = tickerTotal + ws.Cells(i, 7).Value

    End If

    Next i

'Resetting storage variables before moving to the next worksheet  
summaryTable = 2
largestTicker = 0
largestTotal = 0
largestPosChange = 0
largestNegChange = 0
Next iIndex

End Sub