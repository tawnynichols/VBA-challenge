Attribute VB_Name = "Module1"
Sub stocks()
    
    'loop through all worksheets within the woorkbook
    '-----------------------------------------------
    For Each ws In Worksheets
        
        'create variable to find ticker name
        '-----------------------------------------------
        Dim ticker_name As String
        
        'create varibale to keep track of vol total
        '-----------------------------------------------
        Dim vol_total As Double
        vol_total = 0
        
        'create variable for open and close prices
        '-----------------------------------------------
        Dim open_price As Double
        Dim close_price As Double
        
        'create variable for yearly change & %
        '-----------------------------------------------
        Dim yrly_chg As Double
        Dim Percent As Double
        
        'create and set counters
        '-----------------------------------------------
        Dim ticker_counter As Integer
        ticker_counter = 2
        
        Dim open_counter As Integer
        open_counter = 0
        
        
        'print header for ticker symbol, yearly change, %, and volume
        '-----------------------------------------------
        ws.Cells(1, 10).Value = "Ticker Symbol"

        ws.Cells(1, 11).Value = "Yearly Change"
        
        ws.Cells(1, 12).Value = "Percent Change"
        
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        
        'find last row
        '-----------------------------------------------
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'find last coloumn
        '-----------------------------------------------
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        
        'loop to find all ticker ids to create summary table
        '-----------------------------------------------
        For i = 2 To LastRow
            
            
            'Find last row of ticker name
            '-----------------------------------------------
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'print ticker name in summary table
                '-----------------------------------------------
                ticker_name = ws.Cells(i, 1).Value
                
                'add to vol total
                '-----------------------------------------------
                vol_total = vol_total + ws.Cells(i, 7).Value
                
                
                'print ticker name in summary table
                '-----------------------------------------------
                ws.Cells(ticker_counter, 10).Value = ticker_name
                
                
                'print vol total in summary table
                '-----------------------------------------------
                ws.Cells(ticker_counter, 13).Value = vol_total
                
                
                'find close price
                '-----------------------------------------------
                close_price = ws.Cells(i, 6).Value
                
                
                'print yearly change in summary table
                '-----------------------------------------------
                yrly_chg = (close_price - open_price)
                ws.Cells(ticker_counter, 11).Value = yrly_chg
                
                'check for 0 prices that may cause an error
                '-----------------------------------------------
                If open_price = 0 Then
                    
                    ws.Cells(ticker_counter, 12).Value = "0.00%"
                    
                Else
                    
                    'print % of change
                    '-----------------------------------------------
                    Percent = (yrly_chg / open_price)
                    ws.Cells(ticker_counter, 12).Value = Format(Percent, "0.00%")
                    
                End If
                
                
                'add to the ticker counter
                '-----------------------------------------------
                ticker_counter = ticker_counter + 1
                
                'reset vol total
                '-----------------------------------------------
                vol_total = 0
                
                'reset open counter
                '-----------------------------------------------
                open_counter = 0
                
                
            'If cells are the same as ticker
            '-----------------------------------------------
            Else
                
                'add to vol total
                '-----------------------------------------------
                vol_total = vol_total + ws.Cells(i, 7).Value
                
                'find open price
                '-----------------------------------------------
                open_price = ws.Cells(i - open_counter, 3).Value
                open_counter = open_counter + 1
                
                
            End If
            
            
            Next i
            
            'find last row of summary table
            '-----------------------------------------------
            table_last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            'loop to format yearly change
            '-----------------------------------------------
            For i = 2 To table_last_row
                
                'if positve number
                '-----------------------------------------------
                If ws.Cells(i, 11).Value > 0 Then
                    
                    'then set color to green
                    '-----------------------------------------------
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                    
                Else
                    
                    'everything else set color to red
                    '-----------------------------------------------
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                    
                End If
                
                Next i
                
                'set greatest row and colomn headers
                '-----------------------------------------------
                ws.Cells(2, 15).Value = "Greatest % increase"
                ws.Cells(3, 15).Value = "Greatest % decrease"
                ws.Cells(4, 15).Value = "Greatest total volume"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                
                
                'loop to find Greatest increase, decrease and volume
                '-----------------------------------------------
                For i = 2 To table_last_row
                    
                    'if greatest increase then ...
                    '-----------------------------------------------
                    If ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & table_last_row)) Then
                        
                        'print ticker name and value
                        '-----------------------------------------------
                        ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
                        ws.Cells(2, 17).Value = Format(ws.Cells(i, 12).Value, "0.00%")
                        
                    End If
                    
                    'if greatest decrease then...
                    '-----------------------------------------------
                    If ws.Cells(i, 12).Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & table_last_row)) Then
                        
                        'print ticker name and value
                        '-----------------------------------------------
                        ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
                        ws.Cells(3, 17).Value = Format(ws.Cells(i, 12).Value, "0.00%")
                        
                    End If
                    
                    'if greatest volumn then...
                    '-----------------------------------------------
                    If ws.Cells(i, 13).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & table_last_row)) Then
                        
                        'print ticker name and value
                        '-----------------------------------------------
                        ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
                        ws.Cells(4, 17).Value = ws.Cells(i, 13).Value
                        
                        
                    End If
                    
                    
                    Next i
                    
                    'next worksheet
                    '-----------------------------------------------
                    Next ws
                    
                End Sub
                
                
                


