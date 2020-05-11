Attribute VB_Name = "Module4"
Sub stock()
Dim Ticker As String
Dim TickerOpenPriceYearBeginning As Double
Dim TickerClosePriceYearEnd As Double
Dim TickerVolume As LongLong
Dim TickerVolumeTotal As LongLong
Dim TickerCloseStart_OpenEnd_ChangeYearly As Double
Dim TickerCloseStart_OpenEnd_ChangeYearly_Percentage As Double
Dim LastRow As LongLong
Dim TargetRow As LongLong
Dim TickerStart As LongLong
Dim TickerEnd As LongLong

TargetRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Ticker = Cells(2, 1)
TickerOpenPriceYearBeginning = Cells(2, 3)
TickerVolumeTotal = 0
TickerStart = 2
TickerEnd = 2

' Set summary table header
Cells(1, "L") = "Ticker"
Cells(1, "M") = "Volume Total"
Cells(1, "N") = "Delta"
Cells(1, "O") = "Delta%"
Cells(2, "L") = "A"

' Loop through all rows in active worksheet
For Row = 2 To LastRow
    ' Check if ticker name is changed (not equal to the one in previous cell)
    If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
            ' if not the same ticker, need to summarize the results for previous ticker
            ' if not the same ticker, still need to add volume from the current row to the total volume
        TickerVolumeTotal = TickerVolumeTotal + Cells(Row, 7).Value
            ' To calculate the difference between close price at the end of the year and open price at the beginning of the year.
            ' Set new ticker ends and ticker start counters
        TickerEnd = Row ' set ticker end to be equal to current row
        Ticker = Cells(Row, 1).Value ' Set ticker name as in current row, column 1("A")
        Cells(TargetRow, "L") = Ticker ' Note: targetRow is the row inside summary table, where we want to record results for a specific ticker. Set the ticker name into summary table
            ' Set total volume for the ticker into the summary table:
        Cells(TargetRow, "M") = TickerVolumeTotal
            ' 1. Calculate difference between open and close (which are in different columns and in different rows):
            '       a) Use TickerStart and TickerEnd for the rows' counter.
            '       b) Use columns "F" for close price and column "C" for open price.
        TickerCloseStart_OpenEnd_ChangeYearly = Cells(TickerEnd, "F") - Cells(TickerStart, "C")
            '       c) calculate the percentage by dividing by an open price:
        TickerCloseStart_OpenEnd_ChangeYearly_Percentage = (Cells(TickerEnd, "F") - Cells(TickerStart, "C")) / Cells(TickerStart, "C")
            '       d) write results into summary table:
        Cells(TargetRow, "N") = TickerCloseStart_OpenEnd_ChangeYearly
        Cells(TargetRow, "O") = TickerCloseStart_OpenEnd_ChangeYearly_Percentage
            ' When finished writing results for the specific ticker, need to change the following parameters:
        TargetRow = TargetRow + 1 ' in the summary table, next row to fill up will need to be one row below, so existing plus one
        TickerVolumeTotal = 0 ' reset total volume for the specific ticker, so that next ticker calculates from 0
        TickerStart = Row + 1 ' for the close-open calculations, change ticker start counter to the current row + 1
    Else ' if ticker name is the same as ticker name in the cell above:
        TickerVolumeTotal = TickerVolumeTotal + Cells(Row, 7).Value ' add volume from current row to the total ticker volume calculated so far
    End If ' end if condition to avoid compiler issues

Next Row ' continue to the next loop

End Sub
