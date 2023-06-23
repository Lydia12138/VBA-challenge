Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In Worksheets
        'Set the variable
        Dim WorksheetName As String
        Dim j As Long
        Dim start As Long
        'Last row column
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim Percent_Change As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percent_Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set the start row
        start = 2
        
        'Set the result table's start row is 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(start, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(start, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(start, 10).Value < 0 Then
                    ws.Cells(start, 10).Interior.ColorIndex = 3
                
                    Else
                    ws.Cells(start, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate percent change in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(start, 11).Value = Format(Percent_Change, "Percent")
                    
                    Else
                    
                    ws.Cells(start, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(start, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                start = start + 1
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I which I just create
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set the Variant with cells' value
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'Loop for summary table
            For i = 2 To LastRowI
            
                'For greatest total volume
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'For greatest decrease
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
            
    Next ws
        
End Sub

