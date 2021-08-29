Sub stockinfo():



  Dim lastrow As Long
  Dim tickersymbol As String
  Dim yearopen As Double
  Dim tickerlastrow As Integer
  Dim totalstockvol As Double
  Dim yearclose As Double
  Dim yearchange As Double
  Dim percentchange As Double
  Dim lastrowoutput As Long
  Dim percentchange1 As String



  For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    For i = 2 To lastrow
      'Assign ticker symbole of the current row to this vari.able
      tickersymbol = ws.Cells(i, 1).Value

      'Find the lastrow we output to
      tickerlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

      'Check if the symbol of the current row is different than the previous
      If (ws.Cells(i - 1, 1).Value) <> (tickersymbol) Then
        'if different, assign opening price to year open variable
        yearopen = ws.Cells(i, 3).Value
        totalstockvol = ws.Cells(i, 7).Value

      ElseIf ws.Cells(i + 1, 1).Value <> (tickersymbol) Then
        ws.Cells(tickerlastrow + 1, 9).Value = tickersymbol
        totalstockvol = totalstockvol + ws.Cells(i, 7).Value
        'print totalstockvol to output list
        ws.Cells(tickerlastrow + 1, 12).Value = totalstockvol
        'find yearclose value from previous row
        yearclose = ws.Cells(i - 1, 6).Value
        'Calculate year change and print to output
        yearchange = yearclose - yearopen
        ws.Cells(tickerlastrow + 1, 10).Value = yearchange
        'Calculate percentchange and print to output
        If yearopen = 0 Then
            percentchange = 0
        Else
            percentchange = (yearclose - yearopen) / yearopen
        End If
        ws.Cells(tickerlastrow + 1, 11).Value = percentchange

      Else
      'if the ticker is the same, add the stock volume to a totalstockvol
        totalstockvol = totalstockvol + ws.Cells(i, 7).Value


      End If

    Next i

    lastrowoutput = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrowoutput
      ws.Cells(i, 11).Style = "Percent"

      If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
      Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
      End If

    Next i



  Next ws

End Sub
