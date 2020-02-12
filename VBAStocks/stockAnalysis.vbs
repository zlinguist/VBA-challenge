Sub stockAnalysis()

    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets

        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        Dim row As Long
        Dim current_open As Variant
        Dim current_close As Variant
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
        current_row = 2
        current_open = ws.Cells(2, 3).Value
        current_volume = ws.Cells(2, 7).Value

        ws.Range("K1").EntireColumn.NumberFormat = "0.00%"

        For row = 2 To lastrow
            current_volume = current_volume + ws.Cells(row, 7).Value
            
            If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
                current_close = ws.Cells(row, 6).Value()
                ws.Cells(current_row, 9) = ws.Cells(row, 1).Value
                ws.Cells(current_row, 10) = (current_close - current_open)
                If current_open <> 0 Then
                    ws.Cells(current_row, 11) = (ws.Cells(current_row, 10).Value / current_open)
                End If
                ws.Cells(current_row, 12) = current_volume
                
                If ws.Cells(current_row, 10) >= 0 Then
                    ws.Range("J" & current_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & current_row).Interior.ColorIndex = 3
                End If
                    
                current_row = current_row + 1
                current_open = ws.Cells(row + 1, 3).Value
                current_volume = ws.Cells(row + 1, 7).Value
                
            End If
        Next row

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        lastsummalryrow = ws.Cells(Rows.Count, 9).End(xlUp).row

        max_increase = 0
        min_increase = 0
        max_volume = 0
        max_increase_row = 1
        min_increase_row = 1
        max_volume_row = 1

        For row = 2 To lastsummalryrow

            If max_increase < ws.Cells(row, 11) Then
                max_increase = ws.Cells(row, 11).Value
                max_increase_row = row
            End If

            If min_increase > ws.Cells(row, 11) Then
                min_increase = ws.Cells(row, 11)
                min_increase_row = row
            End If

            If max_volume < ws.Cells(row, 12) Then
                max_volume = ws.Cells(row, 12).Value
                max_volume_row = row
            End If

        Next row


        ws.Range("P2") = ws.Cells(max_increase_row, 9).Value()
        ws.Range("Q2") = max_increase
        ws.Range("P3") = ws.Cells(min_increase_row, 9).Value()
        ws.Range("Q3") = min_increase
        ws.Range("P4") = ws.Cells(max_volume_row, 9).Value()
        ws.Range("Q4") = max_volume

        ws.Range("Q2:Q3").NumberFormat = "0.00%"

    Next ws



'    Range("I1:Q70926") = ""

End Sub



