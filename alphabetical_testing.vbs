Sub stock_analysis()

    'For loop to compute all the worksheets
    For Each ws In Worksheets
        
        'Name the header for Column H, I, J, K
        ws.Range("H1").Value = "Ticker Symbol"
        ws.Range("I1").Value = "Yearly Change"
        ws.Range("J1").Value = "Percent Change"
        ws.Range("K1").Value = "Total Stock Volume"
        
        'Define the variable to calculate last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define the variable for total_volume & row_counter & initial price of the first ticker of the worksheet
        total_stock_volume = ws.Range("G2").Value
        row_counter = 2
        start_price = ws.Range("F2").Value
        
            For r = 2 To LastRow
                
                'Defining ticker_counter to specify stock ticker in column H
                ticker_name = ws.Cells(r, 1).Value
                
                If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                    
                    'Insert ticker name in column H consequently
                    ws.Range("H" & row_counter).Value = ticker_name
                    
                    'Insert total stock volume in column K
                    ws.Range("K" & row_counter).Value = total_stock_volume
                    'Resetting stock volume to zero for next ticker
                    total_stock_volume = 0
                    
                    'variable for year end price of the stock
                    end_price = ws.Range("F" & r).Value
                    'Insert total change in stock price from year end to year start in column I
                    ws.Range("I" & row_counter).Value = end_price - start_price
                        
                        'Conditional formatting the percent change column to spit out red when value below 0 and spit out green when value above 0
                        If ws.Range("I" & row_counter).Value >= 0 Then
                            
                            ws.Range("I" & row_counter).Interior.ColorIndex = 4
                            
                            Else
                            
                            ws.Range("I" & row_counter).Interior.ColorIndex = 3
                            
                        End If
                                        
                        'If condition to make sure the start price is always above 0, as anything divided by 0 is infinity and that will give us runtime error
                        If start_price <= 0 Then
                            
                            ws.Range("J" & row_counter).Value = 0
                        
                            Else
                        
                            'yearly percentage difference of the stock
                            ws.Range("J" & row_counter).Value = ((end_price - start_price) / (start_price))
                            ws.Range("J" & row_counter).NumberFormat = "#.##%"
                            
                        End If
                        
                    'Resetting the start price of the stock for the new ticker
                    start_price = ws.Range("F" & (r + 1)).Value
                
                    'Pushing the new ticker data down a row
                    row_counter = row_counter + 1
                                       
                    
                End If
                
                'Adding stock volume over each row
                total_stock_volume = total_stock_volume + ws.Range("G" & (r + 1)).Value
                
            Next r
            
            'Inserting value headers
            ws.Range("N1").Value = "Ticker"
            ws.Range("O1").Value = "Value"
            ws.Range("O2:O3").NumberFormat = "#.##%"
            ws.Range("M2").Value = "Greatest % Increase"
            ws.Range("M3").Value = "Greatest % Decrease"
            ws.Range("M4").Value = "Greatest Total Volume"
            
            ws.Range("O2").Value = WorksheetFunction.Max(ws.Range("J:J"))
            ws.Range("O3").Value = WorksheetFunction.Min(ws.Range("J:J"))
            ws.Range("O4").Value = WorksheetFunction.Max(ws.Range("K:K"))
            
            LastRow_Min_Max = ws.Cells(Rows.Count, 11).End(xlUp).Row
            'Loop through percent change column to find the Greatest % increase, Greatest % decrease and Greatest total volume
            For r = 2 To LastRow_Min_Max
                
                If ws.Range("J" & r).Value = ws.Range("O2").Value Then
                    
                    ws.Range("N2").Value = ws.Range("H" & r).Value
                
                End If
                            
                If ws.Range("J" & r).Value = ws.Range("O3").Value Then
                    
                    ws.Range("N3").Value = ws.Range("H" & r).Value
                
                End If
                
                If ws.Range("K" & r).Value = ws.Range("O4").Value Then
                    
                    ws.Range("N4").Value = ws.Range("H" & r).Value
                
                End If
            
            Next r
            ws.Columns("H:O").AutoFit
                        
    Next ws
    
End Sub
