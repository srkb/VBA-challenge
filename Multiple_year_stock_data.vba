Sub VBAChallengeStockMarket()
'define worksheet
Dim ws As Worksheet

'define variables
Dim Ticker As String
Dim PriceAtYearOpen As Double
Dim PriceAtYearClose As Double
Dim YearlyChange As Double
Dim TotalStockVolume As Double
Dim PercentageChange As Double
Dim StratingRowForResult As Integer

'Loop for worksheets
For Each ws In Worksheets
    
    'Init startrows
    StratingRowForResult = 2
    StartingRowForData = 2
    TotalStockVolume = 0

    'prepare result headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

        'Loop for rows
        For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                PriceAtYearOpen = ws.Cells(StartingRowForData, 3).Value
                PriceAtYearClose = ws.Cells(i, 6).Value

            If PriceAtYearOpen = 0 Then
                PercentageChange = PriceAtYearClose
            Else
                YearlyChange = PriceAtYearClose - PriceAtYearOpen
                PercentageChange = YearlyChange / PriceAtYearOpen
            End If

            If YearlyChange >=0 Then
                ws.Cells(StratingRowForResult, 10).Interior.Color = vbGreen
            Else
                ws.Cells(StratingRowForResult, 10).Interior.Color = vbRed
            End If    

            ws.Cells(StratingRowForResult, 9).Value = Ticker
            ws.Cells(StratingRowForResult, 10).Value = YearlyChange
            ws.Cells(StratingRowForResult, 11).Value = PercentageChange
            ws.Cells(StratingRowForResult, 11).NumberFormat = "0.00%"
            ws.Cells(StratingRowForResult, 12).Value = TotalStockVolume

            StratingRowForResult = StratingRowForResult + 1

            StartingRowForData = i
            YearlyChange = 0
            PercentageChange = 0
            TotalStockVolume = 0

            End If
        Next i
Next ws
End Sub
