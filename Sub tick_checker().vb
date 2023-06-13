Sub tick_checker()

    Dim ws As Worksheet
    
    'loops through workshees
    
    For Each ws In Worksheets
    
        'declaring all needed variables
        
        Dim check_row As Integer
        Dim tick_name As String
        Dim vol As Double
        Dim last_row As Long
        Dim year_open As Double
        Dim year_close As Double
        Dim holder As Integer
        Dim max_per As Double
        Dim min_per As Double
        Dim max_vol As Double
        
        'establishing reference points
        
        check_row = 2
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        holder = 0
        
        'adding column headers
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'looping through the data
        
        For i = check_row To last_row
        
            'if next row is different, adds up all the required data
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                'sets where values are going to be placed
                
                year_close = ws.Cells(i, 6).Value
                tick_name = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                
                'places tick name and year change in appropriate columns
                
                ws.Range("I" & check_row).Value = tick_name
                ws.Range("J" & check_row).Value = year_close - year_open
                
                'colors cells according to yearly change value
                
                If ws.Range("J" & check_row).Value < 0 Then
                    ws.Range("J" & check_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & check_row).Interior.ColorIndex = 4
                End If
                
                'places % change and stock volumn in columns
                
                ws.Range("K" & check_row).Value = FormatPercent((year_close - year_open) / year_open)
                ws.Range("L" & check_row).Value = vol
                
                'moves on to next row for input and resets holder variables
                
                check_row = check_row + 1
                vol = 0
                holder = 0
                
            Else
                
                'if next row is the same, add the volume and check next row
                
                vol = vol + Cells(i, 7)
                
                'grabs the yearly open for a ticker
                
                If holder = 0 Then
                    year_open = ws.Cells(i, 3).Value
                    holder = holder + 1
                End If
                    
            End If
            
        Next i
        
        'searches for greatest increase/decrease and volume then puts them in a chart
        
        ws.Cells(2, 16).Value = FormatPercent(WorksheetFunction.Max(ws.Range("K:K")))
        ws.Cells(3, 16).Value = FormatPercent(WorksheetFunction.Min(ws.Range("K:K")))
        ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))
        
        'sets the above values as variables to match to
        
        max_per = WorksheetFunction.Max(ws.Range("K:K"))
        min_per = WorksheetFunction.Min(ws.Range("K:K"))
        max_vol = WorksheetFunction.Max(ws.Range("L:L"))
        
        'searches for the matching value and returns the ticker name to the chart
        
        ws.Range("O2") = ws.Cells(WorksheetFunction.Match(max_per, ws.Range("K:K"), 0), 9)
        ws.Range("O3") = ws.Cells(WorksheetFunction.Match(min_per, ws.Range("K:K"), 0), 9)
        ws.Range("O4") = ws.Cells(WorksheetFunction.Match(max_vol, ws.Range("L:L"), 0), 9)
            
    Next ws

End Sub

