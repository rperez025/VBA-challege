Attribute VB_Name = "Module1"
Sub stockInfo()

        ' declare worksheets for the loop
        Dim ws As Worksheet
       
        For Each ws In Worksheets

            ' print the column headers for summary stock table 1 and 2
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("O2").Value = "Greatest % increase"
            ws.Range("O3").Value = "Greatest % decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
       
            ' use a variable to reference the last row in the sheet
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
            ' declare Ticker as string
            Dim ticker As String
       
            ' declare opening price as Double
            Dim openingPrice As Double
            openingPrice = 0
       
            ' declare closing price as Double
            Dim closingPrice As Double
            closingPrice = 0
       
            ' declare yearly price change as Double and set as 0
             Dim yearlyPriceChange As Double
            yearlyPriceChange = 0
       
            ' declare percent change as Double and set as 0
            Dim percentChange As Double
            percentChange = 0
       
            ' declare total stock volume as Double
            Dim totalVolume As Double
            ' initialize total stock volume to 0
            totalVolume = 0
       
            ' Stock summary output reference
            Dim summaryStockTable1 As Integer
            summaryStockTable1 = 2 ' start at Row 2
                   
            ' set the opening price for the first ticker symbol in the data set since the rest of the tickers'
            ' open price will be initialized within the For Loop below
            openingPrice = ws.Cells(2, 3).Value
       
            ' create a loop starting from row 2 until the end of the data in the sheet (i.e., lastRow)
            For Row = 2 To (lastRow - 1)
            
                ' one of 2 things will happen in column A
                    ' the ticker symbol will change:
                        ' set the ticker symbol variable to be placed in the summary stock table 1
                        ' calculate the yearly price change to be placed in the summary stock table 1
                        ' calculate the percent change of the price to be placed in the summary stock table 1
                        ' add on the total volume one last time before adding to the summary stock table 1
                        ' put the ticker symbol in the column I
                        ' put the yearly price change in column J
                        ' put the percent change of the price in column K
                        ' add 1 to the summary stock table 1 rows
                        ' reset the total volume to 0
            
                ' or the ticker symbol will not change
                    ' if no changes, then add on to the total volume
            
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' if the ticker symbol changes
                ' set the ticker symbol variable
                ticker = ws.Cells(Row, 1).Value
                
                ' calculate the yearly price change and the percent change
                closingPrice = ws.Cells(Row, 6).Value
                yearlyPriceChange = closingPrice - openingPrice
                
                ' calculate the percentage change
                percentChange = (yearlyPriceChange / openingPrice)
                
                ' add on the total volume on the last item
                totalVolume = totalVolume + ws.Cells(Row, 7)
                
                ' input the ticker symbol in Column I of summary stock table 1
                ws.Cells(summaryStockTable1, 9).Value = ticker
                
                ' input the yearly price change in Column J of summary stock table 1
                ws.Cells(summaryStockTable1, 10).Value = yearlyPriceChange
                
                ' apply conditional formatting to price changes - positive (4) and negative (3)
                If (yearlyPriceChange > 0) Then
                    ws.Range("J" & summaryStockTable1).Interior.ColorIndex = 4
                ElseIf (yearlyPriceChange <= 0) Then
                    ws.Range("J" & summaryStockTable1).Interior.ColorIndex = 3
                End If
                
                ' input the percent change in Column K of summary stock table 1
                ws.Cells(summaryStockTable1, 11).Value = percentChange
                
                 ' apply conditional formatting to percent change - positive (4) and negative (3) (NOTE: Rubric was inconsistent with image in Assignment
                 ' which doesn't show the Percent Change column color-formatted
                If (percentChange > 0) Then
                    ws.Range("K" & summaryStockTable1).Interior.ColorIndex = 4
                ElseIf (percentChange <= 0) Then
                    ws.Range("K" & summaryStockTable1).Interior.ColorIndex = 3
                End If
                
                ' put the final total stock volume in Column L of summary stock table 1
                ws.Cells(summaryStockTable1, 12).Value = totalVolume
                
                ' AutoFit the columns I to L
                ws.Columns("I:L").AutoFit
                
                ' add on 1 to the summary stock table 1 rows (moves to the next row in the summary stock table 1)
                summaryStockTable1 = summaryStockTable1 + 1
                
                ' reset the total volume to 0
                totalVolume = 0
                
                ' reset yearly price change
                yearlyPriceChange = 0
                
                ' reset percent change
                percentChange = 0
                
                ' retrieve next ticker symbol's opening price
                openingPrice = ws.Cells(Row + 1, 3).Value
            
            Else
                ' if the ticker symbol does not change
                ' add the stock volume from Column G to the total charges
                totalVolume = totalVolume + ws.Cells(Row, 7).Value
            
            End If
       
        Next Row
       
            ' populate summary stock table 2
       
            ' declare maximum % change for the greatest % increase and input in summary stock table 2
            Dim maxPercent As Double
            maxPercent = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("Q2").Value = maxPercent
       
            ' declare minimum % change for the greatest % decrease and input in summary stock table 2
            Dim minPercent As Double
            minPercent = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("Q3").Value = minPercent
       
            ' declare maximum total volume for the greatest total stock value and input in summary stock table 2
            Dim maxTotalVolume As Double
            maxTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("Q4").Value = maxTotalVolume
       
            ' match the ticker symbol to the greatest % increase and input in summary stock table 2
            Dim matchTickerMax As Integer
            matchTickerMax = WorksheetFunction.Match(maxPercent, ws.Range("K:K"), 0)
            ws.Range("P2").Value = ws.Range("I" & matchTickerMax).Value
       
            ' match the ticker symbol to the greatest % decrease and input in summary stock table 2
            Dim matchTickerMin As Integer
            matchTickerMin = WorksheetFunction.Match(minPercent, ws.Range("K:K"), 0)
            ws.Range("P3").Value = ws.Range("I" & matchTickerMin).Value
       
            ' match the ticker symbol to the greatest total stock value and input in summary stock table 2
            Dim matchTickerVolume As Integer
            matchTickerVolume = WorksheetFunction.Match(maxTotalVolume, ws.Range("L:L"), 0)
            ws.Range("P4").Value = ws.Range("I" & matchTickerVolume).Value
            
            ' apply % format to column K
            ws.Columns("K:K").NumberFormat = "0.00%"
            
            ' apply % formatting to greatest percentages
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
            ' apply AutoFit formatting to greatest percentages and volume column
            ws.Columns("O:Q").AutoFit
                       
     Next ws
        
End Sub

    

