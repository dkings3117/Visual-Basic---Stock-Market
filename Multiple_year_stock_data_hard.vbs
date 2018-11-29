Sub Hard()
    
    Dim groupRow As Integer
    Dim ticker As String
    Dim yearBeginPrice As Double
    Dim yearEndPrice As Double
    Dim yearlyChange As Double
    Dim yearlyPercentChange As Double
    Dim totalStockVolume As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestVolume As Double
    Dim i As Long
    
    ' Loop through all sheets
    For Each ws In Worksheets

        ' Set up the header row and summary labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        groupRow = 2
        ticker = ws.Range("A2").Value
        yearBeginPrice = ws.Range("C2").Value
        totalStockVolume = 0

        ' Find the last row of each worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' MsgBox (lastRow)

        For i = 2 To lastRow

            totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
            
            ' If last row in the stock group, populate columns I through L
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                yearEndPrice = ws.Range("F" & i).Value
                yearlyChange = yearEndPrice - yearBeginPrice
                If yearBeginPrice <> 0 Then
                    yearlyPercentChange = yearlyChange / yearBeginPrice
                End If

                ws.Range("I" & groupRow).Value = ticker
                ws.Range("J" & groupRow).Value = yearlyChange
                If yearlyChange < 0 Then
                    ws.Range("J" & groupRow).Interior.ColorIndex = 3 ' red
                ElseIf yearlyChange > 0 Then
                    ws.Range("J" & groupRow).Interior.ColorIndex = 4 ' green
                End If
                If yearBeginPrice <> 0 Then
                    ws.Range("K" & groupRow).Value = yearlyPercentChange
                    ws.Range("K" & groupRow).NumberFormat = "0.00%"
                Else
                    ws.Range("K" & groupRow).Value = ""
                End If
                ws.Range("L" & groupRow).Value = totalStockVolume

                ' If first group, then initialize variables
                ' Otherwise compare this row to current min and max values
                If groupRow = 2 Then
                    greatestPercentIncreaseTicker = ticker
                    greatestPercentDecreaseTicker = ticker
                    greatestVolumeTicker = ticker
                    greatestPercentIncrease = yearlyPercentChange
                    greatestPercentDecrease = yearlyPercentChange
                    greatestVolume = totalStockVolume
                Else
                    If yearlyPercentChange > greatestPercentIncrease Then
                        greatestPercentIncreaseTicker = ticker
                        greatestPercentIncrease = yearlyPercentChange
                    ElseIf yearlyPercentChange < greatestPercentDecrease Then
                        greatestPercentDecreaseTicker = ticker
                        greatestPercentDecrease = yearlyPercentChange
                    End If
                    If totalStockVolume > greatestVolume Then
                        greatestVolumeTicker = ticker
                        greatestVolume = totalStockVolume
                    End If
                End If

                ' Initialize next stock group
                groupRow = groupRow + 1
                ticker = ws.Cells(i + 1, 1).Value
                yearBeginPrice = ws.Cells(i + 1, 3)
                totalStockVolume = 0
            End If
        Next i

        ' Display which stocks have the greatest increase & decrease and greatest volume
        ws.Range("P2").Value = greatestPercentIncreaseTicker
        ws.Range("P3").Value = greatestPercentDecreaseTicker
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q2").Value = greatestPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = greatestPercentDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = greatestVolume

        ' Autofit to display data
        ws.Columns("A:Q").AutoFit

    Next ws

End Sub
