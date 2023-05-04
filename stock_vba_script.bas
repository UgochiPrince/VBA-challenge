Attribute VB_Name = "Module1"
Sub stockAnalysis()

    

    'declare variables

    Dim ticker As String

    Dim openingPrice As Double

    Dim closingPrice As Double

    Dim yearlyChange As Double

    Dim percentChange As Double

    Dim totalVolume As Double

    

    'set initial values

    openingPrice = 0

    closingPrice = 0

    yearlyChange = 0

    percentChange = 0

    totalVolume = 0

    

    'set up worksheet

    Dim ws As Worksheet

    

    'loop through all worksheets

    For Each ws In ThisWorkbook.Worksheets

    

        'set up summary table

        ws.Range("I1").Value = "Ticker"

        ws.Range("J1").Value = "Yearly Change"

        ws.Range("K1").Value = "Percent Change"

        ws.Range("L1").Value = "Total Volume"

        

        'set up variables for summary table

        Dim summaryRow As Integer

        summaryRow = 2

        Dim lastRow As Long

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        

        'declare additional variables for finding maximum values

        Dim maxPercentIncreaseTicker As String

        Dim maxPercentDecreaseTicker As String

        Dim maxTotalVolumeTicker As String

        Dim maxPercentIncrease As Double

        Dim maxPercentDecrease As Double

        Dim maxTotalVolume As Double

        maxPercentIncrease = -1

        maxPercentDecrease = 1

        maxTotalVolume = -1

        

        'loop through all stocks

        For i = 2 To lastRow

            

            'check if we've moved on to a new ticker symbol

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                

                'set ticker symbol

                ticker = ws.Cells(i, 1).Value

                

                'set closing price

                closingPrice = ws.Cells(i, 6).Value

                

                'calculate yearly change

                yearlyChange = closingPrice - openingPrice

                

                ' Set the background color of the cell based on the value of yearlyChange

                If yearlyChange < 0 Then

                    ws.Range("J" & summaryRow).Interior.Color = vbRed

                Else

                    ws.Range("J" & summaryRow).Interior.Color = vbGreen

                End If

                        

                'calculate percent change

                If openingPrice <> 0 Then

                    percentChange = yearlyChange / openingPrice

                Else

                    percentChange = 0

                End If

                

                'add to total volume

                totalVolume = totalVolume + ws.Cells(i, 7).Value

                

                'output to summary table

                ws.Range("I" & summaryRow).Value = ticker

                ws.Range("J" & summaryRow).Value = yearlyChange

                ws.Range("K" & summaryRow).Value = percentChange

                ws.Range("K" & summaryRow).NumberFormat = "0.00%"

                ws.Range("L" & summaryRow).Value = totalVolume

                

                'check for maximum values

                If percentChange > maxPercentIncrease Then

                    maxPercentIncrease = percentChange

                    maxPercentIncreaseTicker = ticker

                End If

                

                If percentChange < maxPercentDecrease Then

                    maxPercentDecrease = percentChange

                    maxPercentDecreaseTicker = ticker

                End If

                

                If totalVolume > maxTotalVolume Then

                    maxTotalVolume = totalVolume

                    maxTotalVolumeTicker = ticker

                End If

                            'reset values for next ticker symbol

            openingPrice = 0

            totalVolume = 0

            yearlyChange = 0

            percentChange = 0

            

            'move to next row in summary table

            summaryRow = summaryRow + 1

        

        Else 'if we're still on the same ticker symbol

        

            'set opening price if not already set

            If openingPrice = 0 Then

                openingPrice = ws.Cells(i, 3).Value

            End If

            

            'add to total volume

            totalVolume = totalVolume + ws.Cells(i, 7).Value

            

        End If

        

    Next i

    

    'output maximum values to worksheet

    ws.Range("O2").Value = "Greatest % Increase"

    ws.Range("O3").Value = "Greatest % Decrease"

    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Range("P1").Value = "Ticker"

    ws.Range("Q1").Value = "Value"

    ws.Range("Q2").Value = maxPercentIncrease

    ws.Range("Q2").NumberFormat = "0.00%"

    ws.Range("Q3").Value = maxPercentDecrease

    ws.Range("Q3").NumberFormat = "0.00%"

    ws.Range("Q4").Value = maxTotalVolume

    ws.Range("P2").Value = maxPercentIncreaseTicker

    ws.Range("P3").Value = maxPercentDecreaseTicker

    ws.Range("P4").Value = maxTotalVolumeTicker

    

Next ws

End Sub
