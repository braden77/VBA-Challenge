Attribute VB_Name = "Final"
Option Explicit
Option Base 1
Sub Stock_Analysis()

    'Go through each sheet and display info
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
  
        'Input of the original information in the sheet
        Dim ntickers As Long                'The number of rows for all tickers in column A
        Dim tickers() As String             'All the tickers in column A
        Dim ticker_open() As Single         'Open price in Clolumn C
        Dim ticker_close() As Single        'Close price in column F
        Dim ticker_vol() As Long            'Volumn in column G
        
        'Output of unique tickers
        Dim tickersfound() As String        'Different tickers found
        Dim tickercount() As Long           'The number of different tickers
        Dim yearly_change() As Single       'The yearly change price for each ticker
        Dim percent_change() As Single      'The yearly percent change for each ticker
        Dim total_stock_volume() As Double  'The total volume for each ticker
        
        'Variables uses in finding unique tickers
        Dim isNewTicker As Boolean
        Dim nFound As Long
        
        'Counters
        Dim i As Long
        Dim j As Long
        Dim ticker_beginfrom() As Long
        
        'Find unique tickers and save them into array
        With ws.Range("A1")
            ntickers = Range(.Offset(1, 0), .End(xlDown)).Rows.Count
            ReDim tickers(1 To ntickers)
            ReDim ticker_open(1 To ntickers)
            ReDim ticker_close(1 To ntickers)
            ReDim ticker_vol(1 To ntickers)
            For i = 1 To ntickers
                tickers(i) = .Offset(i, 0).Value
                ticker_open(i) = .Offset(i, 2).Value
                ticker_close(i) = .Offset(i, 5).Value
                ticker_vol(i) = .Offset(i, 6).Value
            Next
        End With
        
        nFound = 0
        'Loop through all the tickers
        For i = 1 To ntickers
            'Set the Boolean isNewTicker to True, and change is to False only if the current product code is already found
            isNewTicker = True
            If nFound > 0 Then
                'loop through all tickers already found and compare them to the current tickers
                For j = 1 To nFound
                    If tickers(i) = tickersfound(j) Then
                        'The current ticker is not a new one, so update the number of rows for this ticker and the total volume for this ticker
                        isNewTicker = False
                        tickercount(j) = tickercount(j) + 1
                        total_stock_volume(j) = total_stock_volume(j) + ticker_vol(i)
                        Exit For
                    End If
                Next
            End If
            
            If isNewTicker Then
                'The current ticker is a new one, so update the list of new ticker found so far, and initialize the number of this ticker found
                ' and the volume of this ticker
                nFound = nFound + 1
                ReDim Preserve tickersfound(1 To nFound)
                ReDim Preserve tickercount(1 To nFound)
                ReDim Preserve total_stock_volume(1 To nFound)
                
                tickersfound(nFound) = tickers(i)
                tickercount(nFound) = 1
                total_stock_volume(nFound) = ticker_vol(i)
            End If
        Next
        
         'Save the row number of a ticker when first appears in column A
         ReDim ticker_beginfrom(1 To nFound)
         ticker_beginfrom(1) = 1
         For j = 2 To nFound
            ticker_beginfrom(j) = tickercount(j - 1) + ticker_beginfrom(j - 1)
         Next
        
    
        'calculate the yearly change and percent change
        ReDim yearly_change(1 To nFound)
        ReDim percent_change(1 To nFound)
        
        For j = 1 To (nFound - 1)
            yearly_change(j) = ticker_close(ticker_beginfrom(j + 1) - 1) - ticker_open(ticker_beginfrom(j))
            If ticker_open(ticker_beginfrom(j)) <> 0 Then
                percent_change(j) = yearly_change(j) / ticker_open(ticker_beginfrom(j))
            'if the openning price for a ticker is 0, set the percent change to 1, otherwise there is error with the percent change
            Else
                percent_change(j) = 1
            End If
        Next
        yearly_change(nFound) = ticker_close(ntickers) - ticker_open(ticker_beginfrom(nFound))
        percent_change(nFound) = yearly_change(j) / ticker_open(ticker_beginfrom(j))
        
        
        'Place the results in the columns
        For j = 1 To nFound
            With ws.Range("I1")
                .Offset(j, 0).Value = tickersfound(j)
                .Offset(j, 1).Value = yearly_change(j)
                .Offset(j, 2).Value = percent_change(j)
                .Offset(j, 3).Value = total_stock_volume(j)
            End With
        Next
        
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Columns("J").NumberFormat = "#,##0.00"
            .Columns("K").NumberFormat = "0.00%"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
        End With
        
        'Change the background colour for yearly change in column J
        For j = 1 To nFound
            If ws.Range("J1").Offset(j, 0).Value > 0 Then
                ws.Range("J1").Offset(j, 0).Interior.Color = vbGreen
            ElseIf ws.Range("J1").Offset(j, 0).Value < 0 Then
                ws.Range("J1").Offset(j, 0).Interior.Color = vbRed
            End If
        Next
        
        'Retrieve the ticker with the greatest increase, decrease and volume
        With ws
            .Range("Q2").Formula = "=max(K:K)"
            .Range("Q3").Formula = "=min(K:K)"
            .Range("Q4").Formula = "=max(L:L)"
            .Range("P2").Formula = "=INDEX(I:I, MATCH(Q2, K:K,0))"
            .Range("P3").Formula = "=INDEX(I:I, MATCH(Q3, K:K,0))"
            .Range("P4").Formula = "=INDEX(I:I, MATCH(Q4, L:L,0))"
        End With
        
        'Format change for all the columns to fit the contents
        ws.Columns("I:Z").AutoFit
    
    Next
                     
End Sub



    
    
 







