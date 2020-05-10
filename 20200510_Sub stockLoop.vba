Sub stock()
Dim Ticker As String
Dim TickerOpenPriceYearBeginning As Double
Dim TickerClosePriceYearEnd As Double
Dim TickerVolume As LongLong
Dim TickerVolumeTotal As LongLong
Dim TickerCloseStart_OpenEnd_ChangeYearly As Double
Dim TickerCloseStart_OpenEnd_ChangeYearly_Percentage As Double
Dim LastRow As LongLong
Dim TargetRow As Integer
TargetRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Ticker = Cells(2, 1)
TickerOpenPriceYearBeginning = Cells(2, 3)
TickerVolumeTotal = 0
' Inside For loop with "row" as a counter
For Row = 2 To LastRow
'TickerClosePriceYearEnd = Cells(1, 1) ' ???
'TickerOpenPriceYearBeginning = Cells(1, 1) '???
TickerVolumeTotal = TickerVolumeTotal + Cells(Row, 7).Value

    ' Inside If condition
    ' If ticker name changed then
    If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
        ' targetRow = targetRow + 1
        TargetRow = TargetRow + 1
        Ticker = Cells(Row, 1).Value
        ' Reset all ticker specific parameters
        'Range("K" & TargetRow).Value = Cells(LastRow, 1).Value
        'Range("L" & TargetRow).Value = Cells(LastRow, 1).Value
            'TickerOpenPriceYearBeginning = reset
            'TickerClosePriceYearEnd = reset
            'TickerVolumeTotal = reset
        ' Set header for the ticker
        Cells(1, "L") = "Ticker"
        Cells(1, "M") = "Volume Total"
        'Set tickername in the target row
        Cells(TargetRow - 1, "L") = Ticker
            ' Calculate total volume for the stock
        Cells(TargetRow - 1, "M") = TickerVolumeTotal

    End If 'condition
Next Row
'TickerVolumeTotal = TickerVolumeTotal + Cells(TargetRow, 7).Value

End Sub
