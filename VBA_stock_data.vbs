Attribute VB_Name = "Module1"
Sub StockPriceAnalysis():
    'Objective: Loop through all stocks in spreadsheet for one year and output:
        'The ticker symbol
        'Yearly change from opening price at beginning of year to closing price at the end of that year for that ticker symbol
        'Percentage change from opening price at beginning of year to closing price at the end of that year for that ticker symbol
        'Total stock volume for that ticker symbol
        
    'This code assumes that the data in the spreadsheet is sorted as below:
        'Sorted chronologically from oldest entries to latest entries top to bottom.
        'Grouped by ticker symbol.
        'One year of data per worksheet in workbook.
        
    'Loop through all worksheets in workbook.
    For Each ws In Worksheets
        
        'Declare variables to hold ticker, opening price, closing price, and total stock price.
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalPrice As Double
        
        'Start totalPrice at 0.
        totalPrice = 0
        
        'Declare counter variables for the loops.
        Dim row As Long
        
        'Declare a variable to decide where to populate totals by ticker. There will be one summary row per ticker symbol.
        'Summaries start at row 2 in sheet.
        Dim summaryRow As Integer
        summaryRow = 2
        
        'Write in the summary table headers in columns I through L.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Find the last row in the dataset.
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Start opening price with the very first opening price in the dataset. It'll get changed when the ticker symbol changes.
        openPrice = ws.Cells(2, 3).Value
        
        'Loop through each row in the dataset.
        For row = 2 To lastRow
            
            'Add to the total stock volume for the current ticker.
            totalPrice = totalPrice + ws.Cells(row, 7).Value
            
            'Check to see if the next ticker symbol is different. If so, this is the latest entry for the current ticker.
            If ws.Cells(row + 1, 1) <> ws.Cells(row, 1) Then
                closePrice = ws.Cells(row, 6).Value
                
                'Populate summary table.
                ws.Cells(summaryRow, 9).Value = ws.Cells(row, 1).Value
                ws.Cells(summaryRow, 10).Value = closePrice - openPrice
                ws.Cells(summaryRow, 11).Value = ((closePrice - openPrice) / openPrice)
                ws.Cells(summaryRow, 12).Value = totalPrice
                
                'Adjust color and style.
                ws.Cells(summaryRow, 11).Style = "Percent"
                If ws.Cells(summaryRow, 10).Value > 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(summaryRow, 10).Value < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If
                
                'Reset the stock volume total and opening price.
                'New opening price should be the next row's open price for the first entry of the new ticker.
                totalPrice = 0
                openPrice = ws.Cells(row + 1, 3)
                
                'Increment counters.
                summaryRow = summaryRow + 1
                
            End If
        Next row
    Next ws
End Sub


Sub TickersSummary():
    'Objective: Identify the ticker symbols with the greatest % increase, greatest % decrease, and greatest total volume.
    'Loop through every worksheet in the workbook.
    For Each ws In Worksheets
        
        'Populate new summary table headers with Ticker and Value in columns O and P respectively.
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Declare a loop control variable.
        Dim row As Integer
        
        'Find the last row in the summary table.
            lastRow = ws.Cells(Rows.Count, 9).End(xlUp).row
    
        'Initialize variables for % increase, % decrease, and total volume, and their tickers.
        Dim percentIncrease As Double
        Dim percentDecrease As Double
        Dim totalVolume As Double
        Dim increaseTicker As String
        Dim decreaseTicker As String
        Dim volumeTicker As String
        
        'Start them all at 0.
        percentIncrease = 0
        percentDecrease = 0
        totalVolume = 0
        
        'Loop through the summary table created by StockPriceAnalysis.
        For row = 2 To lastRow
            'If this row's percent change is higher than the current max, it will be the new max.
            'Performed the same way for min percent change and total volume.
            'Update corresponding ticker symbol values as well.
            If ws.Cells(row, 11).Value > percentIncrease Then
                percentIncrease = ws.Cells(row, 11).Value
                increaseTicker = ws.Cells(row, 9).Value
            End If
            If Cells(row, 11).Value < percentDecrease Then
                percentDecrease = ws.Cells(row, 11).Value
                decreaseTicker = ws.Cells(row, 9).Value
            End If
            If Cells(row, 12).Value > totalVolume Then
                totalVolume = ws.Cells(row, 12).Value
                volumeTicker = ws.Cells(row, 9).Value
            End If
        Next row
        
        'Populate new summary table.
        'Table row headers.
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(4, 14).Value = "Greatest total volume"
        
        'Populate with stored greatest increase, decrease, and volume.
        ws.Cells(2, 16).Value = percentIncrease
        ws.Cells(2, 16).Style = "Percent"
        ws.Cells(3, 16).Value = percentDecrease
        ws.Cells(3, 16).Style = "Percent"
        ws.Cells(4, 16).Value = totalVolume
        
        'Populate with stored ticker values.
        ws.Cells(2, 15).Value = increaseTicker
        ws.Cells(3, 15).Value = decreaseTicker
        ws.Cells(4, 15).Value = volumeTicker
    Next ws
    
End Sub
