Sub tickerStock()
    
    'Worksheet loop'
    
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        
        ws.Activate
        
    'Find the last row of the table
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'Add headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly_Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim ticker As String
    Dim percent_change As Double
    
    Dim volume As Double
    Dim row As Integer
    Dim column As Integer
    
    volume = 0
    row = 2
    column = 1
    
    'Setting the initial price'
    open_price = Cells(2, column + 2).Value
    
    'Loop through all ticker to check for mismatch'
    For i = 2 To last_row
    
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    
    'Setting ticker name'
    ticker = Cells(i, column).Value
    Cells(row, column + 8).Value = ticker
    
    'Setting close price'
    close_price = Cells(i, column + 5).Value
    
    'Calculate quarterly change
    quarterly_change = close_price - open_price
    Cells(row, column + 9).Value = quarterly_change
    
    'Calculate percent change'
    percent_change = quarterly_change / open_price
    Cells(row, column + 10).Value = percent_change
    Cells(row, column + 10).NumberFormat = "0.00%"
    
    'Calculate total volume per quarter'
    volume = volume + Cells(i, column + 6).Value
    Cells(row, column + 11).Value = volume
    
    'Iterate to the next row'
    row = row + 1
    
    'Reset open_price to next ticker'
    open_price = Cells(i + 1, column + 2)
    
    'Reset volume for next ticker'
    volume = 0
    
    Else
        volume = volume + Cells(i, column + 6).Value
    End If
Next i

    'Find the last row of ticker column'
    quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    'Apply conditional formatting to the Quarterly Change Column'
    With ws.Range("J2:J" & quarterly_change_last_row).FormatConditions
    .Delete 'Clear any existing formatting'
    
     'Apply formatting for positive quarterly change (green)'
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(1).Interior.Color = RGB(0, 255, 0) ' Lime color
            
            ' Apply formatting for negative quarterly change (red)'
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(2).Interior.Color = RGB(255, 0, 0) ' Red Color
            
            ' Apply formatting for zero quarterly change (yellow)'
            .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .Item(3).Interior.Color = RGB(255, 255, 255) ' White color
        End With

    'Apply conditional formatting to the Percent Change Column'
    With ws.Range("K2:K" & quarterly_change_last_row).FormatConditions
    .Delete 'Clear any existing formatting'
    
     'Apply formatting for positive quarterly change (green)'
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(1).Interior.Color = RGB(0, 255, 0) ' Lime color
            
            ' Apply formatting for negative quarterly change (red)'
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(2).Interior.Color = RGB(255, 0, 0) ' Red Color
            
            ' Apply formatting for zero quarterly change (yellow)'
            .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .Item(3).Interior.Color = RGB(255, 255, 255) ' White color
        End With

    'Set Ticker, Value, Greatest % Increase, Greatest % Decrease, & Total Volume heaaders'
    Cells(1, 16).Value = "Ticker"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Find the highest value of each ticker'
    For k = 2 To quarterly_change_last_row
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(2, 16).Value = Cells(k, 9).Value
            Cells(2, 17).Value = Cells(k, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(3, 16).Value = Cells(k, 9).Value
            Cells(3, 17).Value = Cells(k, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row)) Then
            Cells(4, 16).Value = Cells(k, 9).Value
            Cells(4, 17).Value = Cells(k, 12).Value
        End If
    
Next k

Next ws

End Sub

Sub tickerStock()
    
    'Worksheet loop'
    
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        
        ws.Activate
        
    'Find the last row of the table
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'Add headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly_Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim ticker As String
    Dim percent_change As Double
    
    Dim volume As Double
    Dim row As Integer
    Dim column As Integer
    
    volume = 0
    row = 2
    column = 1
    
    'Setting the initial price'
    open_price = Cells(2, column + 2).Value
    
    'Loop through all ticker to check for mismatch'
    For i = 2 To last_row
    
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    
    'Setting ticker name'
    ticker = Cells(i, column).Value
    Cells(row, column + 8).Value = ticker
    
    'Setting close price'
    close_price = Cells(i, column + 5).Value
    
    'Calculate quarterly change
    quarterly_change = close_price - open_price
    Cells(row, column + 9).Value = quarterly_change
    
    'Calculate percent change'
    percent_change = quarterly_change / open_price
    Cells(row, column + 10).Value = percent_change
    Cells(row, column + 10).NumberFormat = "0.00%"
    
    'Calculate total volume per quarter'
    volume = volume + Cells(i, column + 6).Value
    Cells(row, column + 11).Value = volume
    
    'Iterate to the next row'
    row = row + 1
    
    'Reset open_price to next ticker'
    open_price = Cells(i + 1, column + 2)
    
    'Reset volume for next ticker'
    volume = 0
    
    Else
        volume = volume + Cells(i, column + 6).Value
    End If
Next i

    'Find the last row of ticker column'
    quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    'Apply conditional formatting to the Quarterly Change Column'
    With ws.Range("J2:J" & quarterly_change_last_row).FormatConditions
    .Delete 'Clear any existing formatting'
    
     'Apply formatting for positive quarterly change (green)'
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(1).Interior.Color = RGB(0, 255, 0) ' Lime color
            
            ' Apply formatting for negative quarterly change (red)'
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(2).Interior.Color = RGB(255, 0, 0) ' Red Color
            
            ' Apply formatting for zero quarterly change (yellow)'
            .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .Item(3).Interior.Color = RGB(255, 255, 255) ' White color
        End With

    'Apply conditional formatting to the Percent Change Column'
    With ws.Range("K2:K" & quarterly_change_last_row).FormatConditions
    .Delete 'Clear any existing formatting'
    
     'Apply formatting for positive quarterly change (green)'
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(1).Interior.Color = RGB(0, 255, 0) ' Lime color
            
            ' Apply formatting for negative quarterly change (red)'
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(2).Interior.Color = RGB(255, 0, 0) ' Red Color
            
            ' Apply formatting for zero quarterly change (yellow)'
            .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .Item(3).Interior.Color = RGB(255, 255, 255) ' White color
        End With

    'Set Ticker, Value, Greatest % Increase, Greatest % Decrease, & Total Volume heaaders'
    Cells(1, 16).Value = "Ticker"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Find the highest value of each ticker'
    For k = 2 To quarterly_change_last_row
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(2, 16).Value = Cells(k, 9).Value
            Cells(2, 17).Value = Cells(k, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(3, 16).Value = Cells(k, 9).Value
            Cells(3, 17).Value = Cells(k, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row)) Then
            Cells(4, 16).Value = Cells(k, 9).Value
            Cells(4, 17).Value = Cells(k, 12).Value
        End If
    
Next k

Next ws

End Sub

