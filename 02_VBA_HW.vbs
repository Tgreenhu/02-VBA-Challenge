Sub Ticker()
    For Each ws In Worksheets
        ws.Activate
        Call Set_Titles
    Next ws
End Sub

Sub Set_Titles()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
    Call Calc_Summary
End Sub

Sub Calc_Summary()
    Dim current_ticker As String
    Dim next_ticker As String
    Dim total_rows As Double
    Dim sum_rows As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    
    total_rows = Cells(Rows.Count, "A").End(xlUp).Row
    sum_rows = 2
    total_volume = 0
    open_price = Cells(2, "C").Value
    For Row = 2 To total_rows
        'TICKER
        current_ticker = Cells(Row, "A").Value
        next_ticker = Cells(Row + 1, "A").Value
        If current_ticker <> next_ticker Then
            Cells(sum_rows, "I").Value = current_ticker
            'YEARLY CHANGE
            close_price = Cells(Row, "F").Value
            yearly_change = (close_price - open_price)
            Cells(sum_rows, "J").Value = yearly_change
            'PERCENT CHANGE
            If (open_price = 0 And close_price = 0) Then
                percent_change = 0
            ElseIf (open_price = 0 And close_price <> 0) Then
                percent_change = 1
            Else
                percent_change = (yearly_change / open_price)
                Cells(sum_rows, "K").Value = percent_change
                Cells(sum_rows, "K").NumberFormat = "0.00%"
            End If
            'TOTAL VOLUME
            total_volume = total_volume + Cells(Row, "G").Value
            Cells(sum_rows, "L").Value = total_volume
            sum_rows = sum_rows + 1
            total_volume = 0
            open_price = Cells(Row + 1, 3).Value
        Else
            total_volume = total_volume + Cells(Row, "G").Value
        End If
    Next Row
    Debug.Print ActiveSheet.Name
End Sub