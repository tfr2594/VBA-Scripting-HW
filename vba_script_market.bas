Attribute VB_Name = "Module1"
Sub stock_market()

    'go through all the worksheets
    For Each ws In Worksheets

        'Create labels for the top the sheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Setting up Varaibles so I use them later in the code
        Dim lastrow As Long
        Dim ticker As String
        Dim open_year As Double
        Dim close_year As Double
        Dim change_year As Double
        Dim change_percent As Double
        Dim prev_amount As Long
        prev_amount = 2
        Dim sum_row As Long
        sum_row = 2
        Dim tot_tick_vol As Double
        tot_tick_vol = 0
        
        
        'where the rows of info ends for each sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'to ignore the first row headers
        For i = 2 To lastrow

            'code behind adding volume under the same ticker
            tot_tick_vol = tot_tick_vol + ws.Cells(i, 7).Value
            'compare that it's still the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                'find the ticker intials for each sheet
                ticker = ws.Cells(i, 1).Value
                'keeps all the tickers in order, 1 ticker name for ach
                ws.Range("I" & sum_row).Value = ticker
                'adds the sum total for those tickers
                ws.Range("L" & sum_row).Value = tot_tick_vol
                'goes back to 0 to refresh
                tot_tick_vol = 0
                'now figure out the yearly change and the percent change
                'but set which columns goes for eachlabel
                open_year = ws.Range("C" & prev_amount)
                close_year = ws.Range("F" & i)
                'math to figure out year_change
                change_year = close_year - open_year
                ws.Range("J" & sum_row).Value = change_year

                'If clause and divsion section
                If open_year = 0 Then
                    change_percent = 0
                'dealt with if 0 now time for finding percent
                Else
                    open_year = ws.Range("C" & prev_amount)
                    change_percent = change_year / open_year
                End If

                ' make the postive change green and the negative red, check notebook for formula
                If ws.Range("J" & sum_row).Value >= 0 Then
                    ws.Range("J" & sum_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & sum_row).Interior.ColorIndex = 3
                End If
            
                ' for a new row to be addded
                sum_row = sum_row + 1
                prev_amount = i + 1
                End If
                
            Next i
    
    Next ws

End Sub
'please work

