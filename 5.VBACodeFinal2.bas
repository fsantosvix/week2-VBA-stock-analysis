Sub stock():

    Dim ws As Worksheet
    
    'Loop through worksheets
    For Each ws In ActiveWorkbook.Worksheets
    
            
            Dim i As Long
            Dim LastRow As Long
            Dim Ticker As String
            Dim OpenJan As Single
            Dim CloseDec As Single
            Dim YearlyChange As Double
            Dim PercentChange As Double
            Dim PrintRow As Integer
            
            'Declare a variable to position the information exported from the dataset
            PrintRow = 2
            
            ' Find out the last used row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Add headers
            ws.Range("J1").Value = "Ticker"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("M1").Value = "Total Stock Volume"
            
            'Set the counter to 0
            TotalVolume = 0
            
            For i = 2 To LastRow
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Print the ticker symbol to the new table
                    ws.Cells(PrintRow, 10).Value = ws.Cells(i, 1).Value
                    
                    'Add to the Total Stock Volume
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    'Print the Total Stock Volume to the new table
                    ws.Cells(PrintRow, 13).Value = TotalVolume
                    'Set the Total Stock Volume back to zero
                    TotalVolume = 0
                    
                    'Retrieve the stock value for the FIRST day of each year in case rows relate to different stocks
                    If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                        OpenJan = ws.Cells(i, 3).Value
                    'Retrieve the stock value for the LAST day of each year in case rows relate to different stocks
                    ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
                        CloseDec = ws.Cells(i, 6).Value
                    End If
                    
                    'Calculate and Print the Yearly Change
                    YearlyChange = CloseDec - OpenJan
                    ws.Cells(PrintRow, 11).Value = Round(YearlyChange, 2)
                    
                    'Format Yearly Change cells
                    If YearlyChange > 0 Then
                        ws.Cells(PrintRow, 11).Interior.ColorIndex = 4
                    ElseIf YearlyChange < 0 Then
                        ws.Cells(PrintRow, 11).Interior.ColorIndex = 3
                    End If
                    
                    'Calculate and Print the Percent Change
                    PercentChange = YearlyChange / OpenJan
                    Percentual = FormatPercent(PercentChange) 'instructions taken from https://analysistabs.com/vba/functions/formatpercent/
                    ws.Cells(PrintRow, 12).Value = Percentual
                    
                    
                     'Format Percent Change cells
                    If PercentChange > 0 Then
                        ws.Cells(PrintRow, 12).Interior.ColorIndex = 4
                    ElseIf PercentChange < 0 Then
                        ws.Cells(PrintRow, 12).Interior.ColorIndex = 3
                    End If
                    
                    'Add one to the new table row
                    PrintRow = PrintRow + 1
                   
                Else
                    
                    'Add to the Total Stock Volume
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                    'Retrieve the stock value for the FIRST day of each year in case rows relate to the same stock
                    If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                        OpenJan = ws.Cells(i, 3).Value
                    'Retrieve the stock value for the LAST day of each year in case rows relate to the same stock
                    ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
                        CloseDec = ws.Cells(i, 6).Value
                    End If
                            
                End If
                
            Next i
            
            '============
            '   Bonus
            '============
            
            Dim LastRowNew As Long
            Dim j As Long
            Dim MaxIncrease As Double
            Dim MaxDecrease As Double
            
            
            ' Find out the last used row
            LastRowNew = ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            'Add headers
            ws.Range("Q1").Value = "Ticker"
            ws.Range("R1").Value = "Value"
            ws.Range("P2").Value = "Greatest % Increase"
            ws.Range("P3").Value = "Greatest % Decrease"
            ws.Range("P4").Value = "Greatest Total Volume"
            
            MaxIncrease = Range("L2").Value
            MaxDecrease = Range("L2").Value
            GreatTotal = Range("M2").Value
            
            
            'Loop through the new table created to find Greatest figures
            For j = 2 To LastRowNew
            
                'Identify the Greatest Increase and equivalent ticker
                If ws.Range("L" & j).Value > MaxIncrease Then
                    MaxIncrease = ws.Range("L" & j).Value
                    IncTicker = ws.Range("J" & j).Value
                    
                End If
                
                'Identify the Greatest Decrease and ticker
                If ws.Range("L" & j).Value < MaxDecrease Then
                    MaxDecrease = ws.Range("L" & j).Value
                    DecTicker = ws.Range("J" & j).Value
                    
                End If
                
                'Identify the Greatest Total Volume and ticker
                If ws.Range("M" & j).Value > GreatTotal Then
                    GreatTotal = ws.Range("M" & j).Value
                    GreatTotalTicker = ws.Range("J" & j).Value
                    
                End If
                
            Next j
            
            ws.Range("R2").Value = FormatPercent(MaxIncrease)
            ws.Range("Q2").Value = IncTicker
            
            ws.Range("R3").Value = FormatPercent(MaxDecrease)
            ws.Range("Q3").Value = DecTicker
            
            ws.Range("R4").Value = GreatTotal
            ws.Range("Q4").Value = GreatTotalTicker
        
    Next ws
        
End Sub
