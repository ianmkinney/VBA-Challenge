Sub ticker()
        
        Dim ticker As String
        Dim tickervol As Double
        Dim ticker_sum As Integer
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
        tickervol = 0
        ticker_sum = 2
        open_price = Cells(2, 3).Value

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              ticker = Cells(i, 1).Value

              tickervol = tickervol + Cells(i, 7).Value

              Range("I" & ticker_sum).Value = ticker

              Range("L" & ticker_sum).Value = tickervol

              close_price = Cells(i, 6).Value

               yearly_change = (close_price - open_price)
              
              Range("J" & ticker_sum).Value = yearly_change
              
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

              Range("K" & ticker_sum).Value = percent_change
              Range("K" & ticker_sum).NumberFormat = "0.00%"
   
              ticker_sum = ticker_sum + 1

              tickervol = 0

              open_price = Cells(i + 1, 3)
            
            Else
              
              tickervol = tickervol + Cells(i, 7).Value

            
            End If
        
        Next i
        
        summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

End Sub

