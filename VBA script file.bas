Attribute VB_Name = "Module1"
Sub stockmarket()
    Dim tickersymbol As String
    Dim YearChange As Double
    Dim PercentageChange As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total_stock_volume As Currency
    Dim ticker_row As Long '
    Dim maxPercentageChange As Double
    Dim maxTicker As String
    Dim minPercentageChange As Double
    Dim minTicker As String
    Dim maxVol As Currency
    maxVol = -1
    Dim maxVol_ticker As String
    ' Loop through each worksheet
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        ticker_row = 2 ' Initialize it here, outside of the loop
        ' Making new column names
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percentage Change"
        ws.Range("N1").Value = "Total stock volume"
        ws.Range("S2").Value = "Greatest % Increase"
        ws.Range("S3").Value = "Greatest % Decrease"
        ws.Range("S4").Value = "Greatest Total Volume"
        ws.Range("T1").Value = "Ticker"
        ws.Range("U1").Value = "Value"
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Reset total_stock_volume before starting to process rows for the current worksheet
        total_stock_volume = 0
        ' Loop through the rows of data
        For i = 2 To lastRow
            ' Check if the next row's ticker symbol is different
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Code for different ticker symbol
                tickersymbol = ws.Cells(i, 1).Value
                ' Put ticker symbol into column K
                ws.Cells(ticker_row, 11).Value = tickersymbol
                ' Put total stock volume into column N
                ws.Cells(ticker_row, 14).Value = total_stock_volume
                ' Find the opening row (loop upwards from the current row)
                Dim openingRow As Long
                openingRow = i
                Do While openingRow > 1 And ws.Cells(openingRow, 1).Value = ws.Cells(i, 1).Value
                    openingRow = openingRow - 1
                Loop
                openingRow = openingRow + 1
                ' Get the opening price from the opening row (assuming opening price is in column C)
                Dim openingPrice As Double
                openingPrice = ws.Cells(openingRow, 3).Value
                ' Find the closing price (from the last row of the ticker)
                Dim closingPrice As Double
                closingPrice = ws.Cells(i, 6).Value ' Assuming closing price is in column F
                ' Calculate price difference and put it in column L (Yearly Change)
                YearChange = closingPrice - openingPrice
                ws.Cells(ticker_row, 12).Value = YearChange
                ' Set cell color based on Yearly Change
                If YearChange > 0 Then
                    ws.Cells(ticker_row, 12).Interior.ColorIndex = 4 ' Green color index
                Else
                    ws.Cells(ticker_row, 12).Interior.ColorIndex = 3 ' Red color index
                End If
                ' Calculate percentage change
                If openingPrice <> 0 Then
                    PercentageChange = (YearChange / openingPrice) * 100
                Else
                    PercentageChange = 0 ' Avoid division by zero error
                End If
                ' Put percentage change in column M (Percentage Change)
                ws.Cells(ticker_row, 13).Value = PercentageChange
                ' Check if this percentage change is the maximum so far
                If PercentageChange > maxPercentageChange Then
                    maxPercentageChange = PercentageChange
                    maxTicker = tickersymbol
                End If
                ' Check if this percentage change is the minimum so far
                If PercentageChange < minPercentageChange Then
                    minPercentageChange = PercentageChange
                    minTicker = tickersymbol
                End If
                ' Increment ticker_row
                ticker_row = ticker_row + 1
                ' Reset total_stock_volume for next ticker
                total_stock_volume = 0
            Else
                ' Code for the same ticker symbol
                ' Accumulate total stock volume for the current ticker symbol
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                ' Check if the current total stock volume is greater than the maximum so far
                If total_stock_volume > maxVol Then
                    maxVol = total_stock_volume
                    maxVol_ticker = Cells(ticker_row, 11) ' Update the maxVol_ticker
                End If
            End If
        Next i
        ' Handle the last ticker symbol
        If lastRow > 1 Then
            ' Put total stock volume into column N for the last ticker symbol
            ws.Cells(ticker_row, 14).Value = total_stock_volume
        End If
        ' Output the maximum percentage change and corresponding ticker symbol to cells T2 and U2 respectively
        ws.Range("T2").Value = maxTicker
        ws.Range("U2").Value = maxPercentageChange
        ' Output the minimum percentage change and corresponding ticker symbol to cells T3 and U3 respectively
        ws.Range("T3").Value = minTicker
        ws.Range("U3").Value = minPercentageChange
        ' Output the maximum total volume and corresponding ticker symbol to cells T4 and U4 respectively
        ws.Range("T4").Value = maxVol_ticker
        ws.Range("U4").Value = maxVol
    Next ws
End Sub
