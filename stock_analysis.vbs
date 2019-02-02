
Sub StockAnalysis():
    Dim lRow As Long    ' Last nonempty row index
    Dim ticker, nextTicker As String    ' Name of current stock ticker
    Dim wRow As Long    'Index of row currently writing to
    Dim volume As Double

    Dim openPrice, closePrice, yearlyChange, percentChange As Double

    ' Keep up with what is greatest % increase, decrease, volume and due to what stocks
    Dim greatestIncrease, greatestDecrease, greatestVolume As Double
    Dim gIncTicker, gDecTicker, gVolTicker As String

    For Each ws In Worksheets
        ' Get last row index
        lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' First, add column headers to the write columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ' Initialize write row
        wRow = 2

        ' Initialize ticker and volume vars
        ticker = ws.Range("A2").Value
        volume = 0

        ' Initialize opening price
        openPrice = ws.Range("C2").Value

        ' Reset greatest record trackers
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        For i = 2 To lRow
            volume = volume + ws.Cells(i, 7)
            nextTicker = ws.Cells(i + 1, 1).Value
            
            If ticker <> nextTicker Then    ' If we are at the end of current ticker
                ' Read in closing price
                closePrice = ws.Cells(i, 6).Value

                ' Compute yearly change
                yearlyChange = closePrice - openPrice

                ' Write out the values
                ws.Cells(wRow, 9).Value = ticker
                ws.Cells(wRow, 10).Value = yearlyChange
                ws.Cells(wRow, 12).Value = volume
                
                ' Compute percent change and write it out
                If openPrice = 0 Then ' Cannot divide by 0
                    ws.Cells(wRow, 11).Value = "#N/A"
                Else
                    percentChange = yearlyChange / openPrice
                    ws.Cells(wRow, 11).Value = percentChange
                    
                    ' Check against the record % trackers
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        gDecTicker = ticker
                    ElseIf percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        gIncTicker = ticker
                    End If
                End If

                ' Compare total volume to record volume
                If volume > greatestVolume Then
                    greatestVolume = volume
                    gVolTicker = ticker
                End If
                

                ' Color the yearlyChange cell red if negative, green if positive
                If yearlyChange > 0 Then
                    ws.Cells(wRow, 10).Interior.Color = RGB(198, 239, 206)
                    ws.Cells(wRow, 10).Font.Color = RGB(0, 97, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(wRow, 10).Interior.Color = RGB(255, 199, 206)
                    ws.Cells(wRow, 10).Font.Color = RGB(156, 0, 6)
                End If

                ' Format percentChange cell as percentage
                ws.Cells(wRow, 11).NumberFormat = "0.00%"

                ticker = nextTicker                 ' Read in the new ticker
                openPrice = ws.Cells(i + 1, 3).Value ' Read the next opening price
                volume = 0                          ' Reset the volume
                wRow = wRow + 1                     ' Go to next write row
            End If
        Next i

        ' Write out the greatest .... stuff
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = gIncTicker
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = gDecTicker
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = gVolTicker
        ws.Range("Q4").Value = greatestVolume

        ' Format the record cells correctly
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "@"

        ' When we are done writing, autofit the write columns
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub

