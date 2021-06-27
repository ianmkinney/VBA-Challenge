Sub challenge()

        summary_table = Cells(Rows.Count, 9).End(xlUp).Row

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

        For i = 2 To summary_table
           
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
                
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
 
            End If
        
        Next i
        
End Sub

