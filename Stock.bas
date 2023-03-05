Attribute VB_Name = "Module1"
Sub stock()
    'Run through all worksheets
    For Each ws In Worksheets
        'Declare variables for Ticker, opening/closing price, stock volume, lastrow
        Dim ticker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim stockVolume As LongLong
        Dim lastRow As Long
        Dim counter As Integer
        Dim greatestInc As Double
        Dim greastestDec As Double
        Dim greatestVol As LongLong
        Dim incTicker As String
        Dim decTicker As String
        Dim volTicker As String
        'Declare values
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        stockVolume = 0
        greatestInc = 0
        greatestDec = 0
        greatestVol = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        counter = 1
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Start for loop
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'This is for a new Ticker
                'Store and print ticker
                ticker = ws.Cells(i, 1).Value
                ws.Cells(counter + 1, 9).Value = ticker
                'Store and print openingPrice
                openingPrice = ws.Cells(i, 3).Value
                'Add stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                'Add to counter
                counter = counter + 1
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'This is for the last row of a ticker before a new ticker
                'Store closingPrice
                closingPrice = ws.Cells(i, 6).Value
                'Computer and print yearly change, and calculate percent change
                yearlyChange = (closingPrice - openingPrice)
                ws.Cells(counter, 10).Value = yearlyChange
                ws.Cells(counter, 11).Value = (yearlyChange / openingPrice)
                'Final addition to stock volume and print stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                ws.Cells(counter, 12).Value = stockVolume
                'Zero out closingPrice, openingPrice, stock volume
                openingPrice = 0
                closingPrice = 0
                stockVolume = 0
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                'This is in case the ticker is the same as prior row
                'Add stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        'Run new loop through summary table to find greatest % increase and decrease, and greatest total volume
        For i = 2 To lastRow
            If ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > greatestInc Then
                'If % change is greater than 0 and greater than previous row
                    'Then store % change as greatestInc and sticker as incTicker
                    greatestInc = ws.Cells(i, 11).Value
                    incTicker = ws.Cells(i, 9).Value
                    'If stock volume greater than previous row, store it as greatestVol and ticker as volTicker
                    If ws.Cells(i, 12).Value > greatestVol Then
                        greatestVol = ws.Cells(i, 12).Value
                        volTicker = ws.Cells(i, 9).Value
                    End If
                ElseIf ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < greatestDec Then
                'Else if % change is less than 0 and less than previous row
                    'Then store % change as greatestDec and decTicker
                    greatestDec = ws.Cells(i, 11).Value
                    decTicker = ws.Cells(i, 9).Value
                    'If stock volume greater than previous row, store it as volTicker
                    If ws.Cells(i, 12).Value > greatestVol Then
                        greatestVol = ws.Cells(i, 12).Value
                        volTicker = ws.Cells(i, 9).Value
                    End If
                ElseIf ws.Cells(i, 12).Value > greatestVol Then
                    greatestVol = ws.Cells(i, 12).Value
                    volTicker = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = incTicker
        ws.Cells(2, 17).Value = greatestInc
        ws.Cells(3, 16).Value = decTicker
        ws.Cells(3, 17).Value = greatestDec
        ws.Cells(4, 16).Value = volTicker
        ws.Cells(4, 17).Value = greatestVol
    'Once everything else is finished, then make sure it runs through every sheet
    Next ws
End Sub
