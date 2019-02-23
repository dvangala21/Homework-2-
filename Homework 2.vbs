Sub Tickervolume()

For Each ws In Worksheets

'set variables
Dim volume As Double
Dim openpx As Double

Dim ticker As String
Dim LastRow As Double
Dim closepx As Double
Dim SummaryRow As Double

ws.Cells(1, 10).Value = "ticker"
ws.Cells(1, 11).Value = "Volume"
ws.Cells(1, 12).Value = "open"
ws.Cells(1, 13).Value = "close"
ws.Cells(1, 14).Value = "diff"
ws.Cells(1, 15).Value = "pct change"
'set last row variable
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set initial values
SummaryRow = 2
volume = 0
openpx = ws.Range("C2").Value



For i = 2 To LastRow
      ' when on last row of ticker , add final volume and display
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            volume = volume + ws.Range("G" + CStr(i)).Value
            ws.Cells(SummaryRow, 11).Value = volume
            volume = 0
            ticker = ws.Cells(i, 1).Value
            closepx = ws.Cells(i, 6).Value
            ws.Cells(SummaryRow, 10).Value = ticker
            ws.Cells(SummaryRow, 12).Value = openpx
            ws.Cells(SummaryRow, 13).Value = closepx
            ws.Cells(SummaryRow, 14).Value = closepx - openpx
            ws.Cells(SummaryRow, 15).Value = FormatPercent((closepx - openpx) / openpx)
            openpx = ws.Cells(i + 1, 3).Value
            SummaryRow = SummaryRow + 1
       Else
            volume = volume + ws.Cells(i, 7).Value

   
            
            
    End If
    Next i
    
'new loop for coloring


LastSummaryRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
For SummaryRow = 2 To LastSummaryRow
    If ws.Cells(SummaryRow, 14).Value < 0 Then
        ws.Cells(SummaryRow, 14).Interior.ColorIndex = 3
    Else
        ws.Cells(SummaryRow, 14).Interior.ColorIndex = 4
    End If
    SummaryRow = SummaryRow + 1
    Next SummaryRow


' loop through to find max percentage change

        maxpct = ws.Cells(2, 15).Value
For SummaryRow = 2 To LastSummaryRow
    If ws.Cells(SummaryRow, 15).Value > maxpct Then
        maxpct = ws.Cells(SummaryRow, 15).Value
    End If
Next SummaryRow


Dim minpct As Double
Dim maxvol As Double


' loop through to find max volume

        maxvol = ws.Cells(2, 11).Value
For SummaryRow = 2 To LastSummaryRow
    If ws.Cells(SummaryRow, 11).Value > maxvol Then
        maxvol = ws.Cells(SummaryRow, 11).Value
    End If
Next SummaryRow

' loop through to find min percent

  minpct = ws.Cells(2, 15).Value
For SummaryRow = 2 To LastSummaryRow
    If ws.Cells(SummaryRow, 15).Value < minpct Then
        minpct = ws.Cells(SummaryRow, 15).Value
    End If
Next SummaryRow
  
 'print the values in their respective slottzz
 
  ws.Cells(2, 18).Value = maxpct
  ws.Cells(3, 18).Value = minpct
  ws.Cells(4, 18).Value = maxvol
  
    Next ws
  
  End Sub
