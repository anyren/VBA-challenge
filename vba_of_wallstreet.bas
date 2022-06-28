Attribute VB_Name = "Module1"
Sub vba_of_wallstreet()

    Dim row As Long
    Dim column_count As Long
    Dim ticker As String
    Dim summary_table_row As Long
    Dim year_close As Double
    Dim year_open As Double
    Dim volume As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim greatest_volume As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double

    For Each ws In Worksheets
        'Set summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'create variable to keep track of where to insert next row in summary table
        summary_table_row = 2
        
        'find number of values in column A
        column_count = ws.Range("A" & Rows.count).End(xlUp).row
        
        'assuming data is sorted by ticker, date - add sort later?
        For row = 2 To column_count
            ticker = ws.Cells(row, 1).Value
            If ticker <> ws.Cells(row + 1, 1).Value Then
            
                'summary calculation
                volume = volume + ws.Cells(row, 7)
                year_close = ws.Cells(row, 6).Value
                year_change = year_close - year_open
                percent_change = year_change / year_open
                
                'update summary table
                ws.Cells(summary_table_row, 9).Value = ticker
                ws.Cells(summary_table_row, 10).Value = year_change
                ws.Cells(summary_table_row, 11).Value = percent_change
                ws.Cells(summary_table_row, 11).Value = FormatPercent(Cells(summary_table_row, 11).Value)
                ws.Cells(summary_table_row, 12).Value = volume
                
                'format the yearly_change column
                If year_change > 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                ElseIf year_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
                
                'increment counters
                summary_table_row = summary_table_row + 1
                volume = 0
            Else
                volume = volume + ws.Cells(row, 7).Value
                If ticker <> ws.Cells(row - 1, 1).Value Then
                    year_open = ws.Cells(row, 3).Value
                End If
            End If
        Next row
        
        'create bonus table
        'set headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'set variables
        column_count = ws.Range("I" & Rows.count).End(xlUp).row
        greatest_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & column_count).Value)
        greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & column_count).Value)
        greatest_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & column_count).Value)
        
        For row = 2 To row_count
            ticker = ws.Cells(row, 9).Value
            If ws.Cells(row, 11).Value = greatest_increase Then
                ws.Range("P2") = ticker
                ws.Range("Q2") = FormatPercent(greatest_increase)
            ElseIf ws.Cells(row, 11).Value = greatest_decrease Then
                ws.Range("P3") = ticker
                ws.Range("Q3") = FormatPercent(greatest_decrease)
            End If
            'separate out conditional for volume since it could also meet one of the above criteria
            If ws.Cells(row, 12).Value = greatest_volume Then
                ws.Range("P4") = ticker
                ws.Range("Q4") = greatest_volume
            End If
        Next row
    
    Next ws

End Sub

