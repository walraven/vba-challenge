Sub analyzeStocks()
    
    For Each ws In Worksheets

    Dim maxVol As LongLong
    Dim maxVolTick As String
    Dim maxInc As Double
    Dim maxIncTick As String
    Dim maxDec As Double
    Dim maxDecTick As String
    
    maxVol = 0
    maxVolTick = ""
    maxInc = 0
    maxIncTick = ""
    maxDec = 0
    maxDecTick = ""
        
    Dim ticker As String
    Dim tickerCount As Integer
    Dim startPrice As Double
    Dim endPrice As Double
    Dim stockVolume As LongLong
    
    ticker = ""
    tickerCount = 1
    'loop through all the stocks in the sheet
    Dim i As Long
    Dim continue As Boolean
    continue = True
    i = 2
    ticker = ws.Cells(i, 1).Value
    startPrice = ws.Cells(i, 3).Value
    endPrice = 0
    stockVolume = 0
    
    'create header row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'change to for loop?
    'for i = 2 to sheet.usedrows.last?
    Do While continue = True
    
    'append columns that display one row for each
        'ticker symbol
            'yearly change
            'yearly percent change
            'total volume of the stock
        'use conditional formatting to highlight
            'positive change in green
            'negative change in red
            
    'check if ticker matches current ticker
    If ws.Cells(i, 1).Value = ticker Then
        stockVolume = stockVolume + ws.Cells(i, 7).Value
    Else
        Debug.Print (ticker)
        tickerCount = tickerCount + 1
        endPrice = ws.Cells(i - 1, 6).Value
        'create row for ticker display
        ws.Cells(tickerCount, 9).Value = ticker
        ticker = ws.Cells(i, 1).Value
        ws.Cells(tickerCount, 10).Value = endPrice - startPrice
        If (endPrice - startPrice) < 0 Then
            ws.Cells(tickerCount, 11).Value = 1 - (endPrice / startPrice)
                maxDec = ws.Cells(tickerCount, 11).Value
                maxDecTick = ws.Cells(i - 1, 1).Value
            ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
        Else
            If startPrice <> 0 Then 'no division by 0
                ws.Cells(tickerCount, 11).Value = (endPrice / startPrice) - 1
                If ((endPrice / startPrice) - 1) > maxInc Then
                maxInc = ws.Cells(tickerCount, 11).Value
                maxIncTick = ws.Cells(i - 1, 1).Value
                End If
            Else
                ws.Cells(tickerCount, 11).Value = "N/A"
                End If
            
            ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
            End If
        ws.Cells(tickerCount, 12).Value = stockVolume
        'check max
        If maxVol < stockVolume Then
            maxVol = stockVolume
            maxVolTick = ws.Cells(i - 1, 1).Value
            End If
        startPrice = ws.Cells(i, 3).Value
        stockVolume = ws.Cells(i, 7).Value
        End If
    If ws.Cells(i, 1).Value = "" Then
        continue = False
        End If
    i = i + 1
    Loop
                    
                   
    'additional rows that display
        'greatest % increase
        'greatest % decrease
        'greatest total volume

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = maxIncTick
    ws.Cells(2, 17).Value = maxInc
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maxDecTick
    ws.Cells(3, 17).Value = maxDec
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maxVolTick
    ws.Cells(4, 17).Value = maxVol


    Next
    'move to next sheet
    

        
End Sub
