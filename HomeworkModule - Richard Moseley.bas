Attribute VB_Name = "Module1"
Sub main()
'Loop through worksheets
For Each ws In Worksheets
        'Establish variables
        Dim lastRow As Long
        Dim totalVol As Double
        totalVol = 0
        Dim writeRow As Long
        writeRow = 2
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim yearChange As Double
        Dim placeHold As Long
        placeHold = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Label columns and rows for assessed data
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'Loop through data
    For i = 2 To lastRow
        'Add to Volume and check next ticker
        totalVol = ws.Cells(i, 7).Value + totalVol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Perform all write in operations when changing ticker symbol
            ws.Cells(writeRow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(writeRow, 12).Value = totalVol
            yearOpen = ws.Cells(placeHold, 3)
            yearClose = ws.Cells(i, 6)
            yearChange = yearClose - yearOpen
            ws.Cells(writeRow, 10).Value = yearChange
            If yearOpen = 0 Then
                ws.Cells(writeRow, 11).Value = 0
            Else
                ws.Cells(writeRow, 11).Value = yearChange / yearOpen
            End If
            'Add appropriate formatting
            ws.Cells(writeRow, 11).NumberFormat = "0.00%"
            If ws.Cells(writeRow, 10).Value <= 0 Then
                ws.Cells(writeRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(writeRow, 10).Interior.ColorIndex = 4
            End If
            'Iterate/Reset variables
            placeHold = i + 1
            writeRow = writeRow + 1
            totalVol = 0
        End If
    Next i
    'Bonus Points
        'Tried this for a long time with odd results. Apparently the second transfer of
        'data has to happen in its own "for loop"? Is this due to sequencing?
    For i = 2 To lastRow
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value <= ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
    Next i
Next ws
End Sub

