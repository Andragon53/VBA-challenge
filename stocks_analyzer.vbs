Sub StockSort()
    ' Initial variables - counter, ticker code, table length
    Dim i As Long
    Dim ws As Worksheet
    
    ' Loop through all worksheets
    For Each ws In Worksheets
        ' initial variables
        Dim tickerCode As String
        Dim lastRow As Long

        ' Set initial variables for holding the three tracked stats
        Dim endingPrice As Double
        Dim startingPrice As Double
        Dim totalStockVol As Double
        startingPrice = 0
        endingPrice = 0
        totalStockVol = 0
        
        ' Keep track of the location for each ticker code in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Find the last row of the table to set the end of the loop
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' set the first ticker's opening price
        startingPrice = ws.Cells(2, 3).Value

        ' Loop through all stock table codes
        For i = 2 To lastRow
            ' Check if we are still within the same ticker code, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker code
                tickerCode = ws.Cells(i, 1).Value

                ' Add to the stock volume
                totalStockVol = totalStockVol + ws.Cells(i, 7).Value

                ' Get the ending price for the ticker code
                endingPrice = ws.Cells(i, 6).Value

                ' Print the ticker code in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = tickerCode

                ' Print the change in price in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = (endingPrice - startingPrice)

                ' Print the calculated percentage change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = ((endingPrice - startingPrice) / startingPrice)

                ' Print the Total Stock Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = totalStockVol

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Set the next ticker's opening price
                startingPrice = ws.Cells(i + 1, 3).Value

                ' Reset the stock volume
                totalStockVol = 0
            ' If the cell immediately following a row is the same brand...
            Else
                ' Add to the stock volume
                totalStockVol = totalStockVol + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub

Sub StockMax()
    ' Initial variables - counter, ticker code, table length
    Dim i As Long
    Dim ws As Worksheet
    
    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Define and set initial variables for holding the three tracked stats
        Dim increaseTickerCode As String
        Dim decreaseTickerCode As String
        Dim totalTickerCode As String
        Dim highPercent As Double
        Dim lowPercent As Double
        Dim highStock As Double
        highPercent = 0
        lowPercent = 0
        highStock = 0
        
        ' Find the last row of the table to set the end of the loop
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Loop through all stock table codes
        For i = 2 To lastRow
            ' Check if the percentage is greater than the stored greatest. If so...
            If ws.Cells(i, 11).Value > highPercent Then
                ' Set the ticker code to the new code
                increaseTickerCode = ws.Cells(i, 9).Value

                ' Set the max percentage to the new highest
                highPercent = ws.Cells(i, 11).Value
            End If

            ' Check if the percentage is less than the stored lowest. If so...
            If ws.Cells(i, 11).Value < lowPercent Then
                ' Set the ticker code to the new code
                decreaseTickerCode = ws.Cells(i, 9).Value

                ' Set the max percentage to the new lowest
                lowPercent = ws.Cells(i, 11).Value
            End If

            ' Check if the stock total is greater than the stored total. If so...
            If ws.Cells(i, 12).Value > highStock Then
                ' Set the ticker code to the new code
                totalTickerCode = ws.Cells(i, 9).Value

                ' Set the max percentage to the new highest
                highStock = ws.Cells(i, 12).Value
            End If
        Next i

        ' Print out the variables to the secondary summary table
        ws.Range("P2").Value = increaseTickerCode
        ws.Range("P3").Value = decreaseTickerCode
        ws.Range("P4").Value = totalTickerCode
        ws.Range("Q2").Value = highPercent
        ws.Range("Q3").Value = lowPercent
        ws.Range("Q4").Value = highStock
    Next ws
End Sub
