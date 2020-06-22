Attribute VB_Name = "Module1"
Sub stock()

Dim r, k As Integer
Dim volume, close_price, open_price, year_change, perc_change, perc_max, perc_min, max_volume As Double
Dim ticker, max_perc_ticker, min_perc_ticker, max_volume_ticker As String
Dim ws As Worksheet
'Dim max_range, min_range, volume_range, myrange, myrange2 As Range



r = 2
i = 2


'loop through sheets
For Each ws In ThisWorkbook.Worksheets

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    open_price = ws.Cells(2, 3).Value 'sets first open price
    
    Do While IsEmpty(ws.Cells(i, 1).Value) = False 'goes down the row until there are no other entries
        volume = volume + ws.Cells(i, 7) 'adds total volume
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'checks for changes in ticker name
            ws.Cells(r, 9).Value = ws.Cells(i, 1).Value 'sets ticker name in chart
            ws.Cells(r, 12).Value = volume 'sets total volume in chart
            volume = 0 'resets total volume
            close_price = ws.Cells(i, 6).Value 'determines final price
            year_change = close_price - open_price 'calulates yearly change
            perc_change = close_price / open_price 'calulates percentage change
            ws.Cells(r, 11).Value = perc_change 'puts percentage change into table
            ws.Cells(r, 11).NumberFormat = "0.00%" 'converts percentage change into percent
            ws.Cells(r, 10).Value = year_change 'puts yearly change into table
            ws.Cells(i, 3).Value = open_price 'sets opening price for next set
                If ws.Cells(r, 10).Value > 0 Then 'conditional formating for color
                    ws.Cells(r, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(r, 10).Value < 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 3
                End If
                    
            ws.Columns("I:L").AutoFit
            
            
                    
            r = r + 1 'next r for table placement
            
        End If
    i = i + 1 'next i for loop
    Loop
    r = 2 'resets increment
    i = 2 'resets increment
    
    
    'here down was a working attempt at the chellenge portion.
    
    
    'ws.Cells(1, 16).Value = "Value"
    'ws.Cells(1, 15).Value = "Ticker"
    'ws.Cells(2, 14).Value = "Greatest % Increase"
    'ws.Cells(3, 14).Value = "Greatest % Decrease"
    'ws.Cells(4, 14).Value = "Greatest Total Volume"

    
    'Set myrange = Worksheets(ws).Range("J:J")
    'Set myrange2 = Worksheets(ws).Range("L:L")
    'perc_max = Application.WorksheetFunction.Max("J:J")
    'MsgBox (perc_max)
    
    'perc_max = ws.Cells(2, 16).Value
    'perc_min = Application.WorksheetFunction.Min("J2:J10000")
    'max_volume = Application.WorksheetFunction.Max("L2:L10000")
        
    'ws.Cells(2, 16).Value = perc_max
    'ws.Cells(3, 16).Value = perc_min
    'ws.Cells(4, 16).Value = max_volume
    
    
    'Set max_range = Range("J:J").Find(perc_max)
    'Set min_range = Range("J:J").Find(perc_min)
    'Set volume_range = Range("L:L").Find(vol_max)
    
    
        
    'MsgBox (max_range.Address)
    'MsgBox (min_range.Address)
    'MsgBox (volume_range.Address)
    

    'ws.Cells(2, 15).Value = max_perc_ticker
  '  ws.Cells(3, 15).Value = min_perc_ticker
   ' ws.Cells(4, 15).Value = max_volume_ticker
    

       
        
        
        
    
Next

End Sub

