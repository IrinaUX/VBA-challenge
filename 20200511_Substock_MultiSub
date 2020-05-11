Sub stock()

' Increment the function through each worksheet in the workbook
For Each WS In Worksheets
    Dim Worksheet As String
    WS.Activate
    Dim Ticker As String
    Dim TickerOpenPriceYearBeginning As Double
    Dim TickerClosePriceYearEnd As Double
    Dim TickerVolume As LongLong
    Dim TickerVolumeTotal As LongLong
    Dim TickerCloseStart_OpenEnd_ChangeYearly As Double
    Dim TickerCloseStart_OpenEnd_ChangeYearly_Percentage As Double
    Dim lastrow As LongLong
    Dim TargetRow As LongLong
    Dim TickerStart As LongLong
    Dim TickerEnd As LongLong
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As LongLong
    Dim TargetGreatestRow As Integer
    
    TargetRow = 2
    TargetGreatestRow = 2
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
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
'    ' Set greatest increase table static
'    Cells(2, "R") = "Greatest % Increase"
'    Cells(3, "R") = "Greatest % Decrease"
'    Cells(4, "R") = "Greatest Total Volume"
'    Cells(1, "S") = "Ticker"
'    Cells(1, "T") = "Value"
'
    ' Loop through all rows in active worksheet
    For Row = 2 To lastrow
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
                '           Note: overflow here in sheet "P", likely because cannot divide by zero:
                '           i. try to catch condition, when start price is zero:
            If Cells(TickerStart, "C") = 0 Then
                'MsgBox ("Cell at ticker start in worksheet " & WS.Name & " is 0 at line " & TickerStart)
                Cells(TargetRow, "K") = "Cell at ticker start in worksheet " & WS.Name & " is 0 at line " & TickerStart
            Else '       d) calculate percentage:
                TickerCloseStart_OpenEnd_ChangeYearly_Percentage = (Cells(TickerEnd, "F") - Cells(TickerStart, "C")) / Cells(TickerStart, "C")
'                '          e) highlight red, if change is negative:
'                If TickerCloseStart_OpenEnd_ChangeYearly < 0 Then
'                    Cells(TargetRow, "N").Interior.ColorIndex = 3
'                Else ' highlight gree, if positive
'                    Cells(TargetRow, "N").Interior.ColorIndex = 4
'                End If
'                '          f) highlight red, if change is negative for the percentage column:
'                If TickerCloseStart_OpenEnd_ChangeYearly_Percentage < 0 Then
'                    Cells(TargetRow, "O").Interior.ColorIndex = 3
'                Else ' highlight gree, if positive
'                    Cells(TargetRow, "O").Interior.ColorIndex = 4
'                End If
            End If
            Cells(TargetRow, "N") = TickerCloseStart_OpenEnd_ChangeYearly
            Cells(TargetRow, "O") = TickerCloseStart_OpenEnd_ChangeYearly_Percentage
                ' find last row in the summary table
            Dim LastRowInSummaryTable As Long
            LastRowInSummaryTable = Cells(Rows.Count, "O").End(xlUp).Row
                 Range("O2:O" & LastRowInSummaryTable).NumberFormat = "0.00%"
'            Columns("O").NumberFormat = "0.00%"
    '            If Range("N2:N" & LastRowInSummaryTable).Value < 0 Then
    '                Columns("N").Interior.ColorIndex = 3
    '            Else
    '                Columns("O").Interior.ColorIndex = 4
    '            End If
                ' When finished writing results for the specific ticker, need to change the following parameters:
            TargetRow = TargetRow + 1 ' in the summary table, next row to fill up will need to be one row below, so existing plus one
            TickerVolumeTotal = 0 ' reset total volume for the specific ticker, so that next ticker calculates from 0
            TickerStart = Row + 1 ' for the close-open calculations, change ticker start counter to the current row + 1
        Else ' if ticker name is the same as ticker name in the cell above:
            TickerVolumeTotal = TickerVolumeTotal + Cells(Row, 7).Value ' add volume from current row to the total ticker volume calculated so far
        End If ' end if condition to avoid compiler issues
'        ' 3. Inside For loop, check which ticker has maximum close-open increase in percentage
'        '     a) Initialize first percent increase and percent decrease from the summary table as the Greatest Increase:
'        '     b) Check if cell below is bigger than current cell, if yes, update the greatest increase value
'        If Cells(Row + 1, "O").Value > Cells(Row, "O").Value Then
'            GreatestIncrease = Cells(Row + 1, "O").Value
'            GreatestIncreaseTicker = Cells(Row + 1, "A").Value
'        Else
'            GreatestDecrease = Cells(Row + 1, "O").Value
'            GreatestDecreaseTicker = Cells(Row + 1, "A").Value
'        End If
'        ' Update Greatest Summary table
'        Cells(2, "S") = GreatestIncrease
'        Cells(3, "S") = GreatestDecrease
'        '    c) check greatest total volume
        
    Next Row ' continue to the next loop
Call TableFormatting
Call CopySummaryTableToNewSheet
Call GreatestSummaryTable
Call FindTickerForGreatestIncrease
Call FindTickerForGreatestDecrease
Call FindTickerForGreatestTotalVolume
Next WS
End Sub

Sub TableFormatting()
    'MsgBox ("Ready to format the summary table?")
    Dim lastrow As LongLong
    lastrow = Cells(Rows.Count, "L").End(xlUp).Row
'    WS.Activate
    
'    Dim d As Double
    Dim Range As Range
    Dim Row As LongLong
    Set Range = ActiveSheet.Range("N2:N" & lastrow)
    For Row = 2 To lastrow
    'If Cell.Text <> "" And IsNumeric(Cell.Value) = True Then
    If Cells(Row, "N").Value < 0 Then
        Cells(Row, "N").Interior.ColorIndex = 3
    ElseIf Cells(Row, "N").Value > 0 Then
        Cells(Row, "N").Interior.ColorIndex = 4
    End If
    If Cells(Row, "O").Value < 0 Then
        Cells(Row, "O").Interior.ColorIndex = 3
    ElseIf Cells(Row, "O").Value > 0 Then
        Cells(Row, "O").Interior.ColorIndex = 4
    End If
    'End If
    Next
    
'    Dim r As Range
'    r = Range("N2")
'    r.FormatConditions.Add Type:=xlExpression, Formula1:="=$N2>0"
'    r.FormatConditions(r.FormatConditions.Count).SetFirstPriority
'    With r.FormatConditions(1)
'        '.Interior.PatternColorIndex = xlAutomatic
'        .Interior.ColorIndex = 4
'        '.Font.ColorIndex = 26
'    End With
'    r.FormatConditions(1).Add Type:=xlExpression, Formula1:="=$N2<0"
'    With r.FormatConditions(1)
'        '.Interior.PatternColorIndex = xlAutomatic
'        .Interior.ColorIndex = 3
'        '.Font.ColorIndex = 26
'    End With
'
'    Set r = Nothing
'    Range("N2:N" & LastRow).FormatConditions.QueryInterface
    'MsgBox ("Last row = " & LastRow)
'    With WS.Range(2, LastRow)
'        .FormatConditions.Delete
'    End With
  
End Sub

Sub CopySummaryTableToNewSheet()

    'MsgBox ("Ready to sort?")
    
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "O").End(xlUp).Row
    Range("O1:O" & lastrow).Copy Range("Q1:Q" & lastrow)
    Range("T1:T" & lastrow).EntireColumn.AutoFit
    Range("Q2:Q" & lastrow).Sort key1:=Range("Q2:Q" & lastrow), order1:=xlDescending, Header:=xlNo
    Cells(1, "Q") = "Delta % sorted"
    
    'Range("L1:L" & lastrow).Copy Range("P1:P" & lastrow)
    'Range("P1:P" & lastrow).EntireColumn.AutoFit
    Range("M1:M" & lastrow).Copy Range("R1:R" & lastrow)
    Range("R1:R" & lastrow).EntireColumn.AutoFit
    Range("R2:R" & lastrow).Sort key1:=Range("R2:R" & lastrow), order1:=xlDescending, Header:=xlNo
    
    'Range("P1:P & lastrow).Sort key1:=Range("P2:R" & lastrow), order1:=xlDescending, Header:=xlNo
    
    'Range("O1:O" & lastrow).PasteSpecial
    
    'Range("A1").EntireRow.Copy Range("A5")
    
    'Range("O2:O" & lastrow).Sort key1:=Range("O2:O" & lastrow), order1:=xlDescending, Header:=xlNo
   
'    Dim LastRow As LongLong
'    LastRow = Cells(Rows.Count, "O").End(xlUp).Row
'    Dim objWorkBook As String
'    objWorkBook = WS.Workbooks.Add()
'    Worksheets("Sheet1").Move After:=Worksheets("P")
'
    
    'objWorksheet.Name = "A-New"
    'ActiveSheet.Move Before:=ActiveWorkbook.Sheets(1)
    
' Create Excel object
'Set objExcel = CreateObject("Excel.Application")
' Open the workbook
' Set objWorkbook = objExcel.Workbooks.Open _
'   ("C:\myworkbook.xlsx")
' Set to True or False, whatever you like
    'objExcel.Visible = True

' Select the range on Sheet1 you want to copy
    'objWorkBook.Worksheets(1).Range("A1:A" & LastRow).Copy
' Paste it on Sheet2, starting at A1
    'objWorkBook.Worksheets("Sheet1").Range("A1").PasteSpecial
' Activate Sheet2 so you can see it actually pasted the data
'objWorkBook.Worksheets("Sheet").Activate


    'Dim GreatestIncreasePercentage As Double
'    Dim GreatestDecreasePercentage As Double
'    Dim GreatestTotalVolume As LongLong
'
'
'
'
'
End Sub


Sub GreatestSummaryTable()
    ' Set greatest increase table static
    Cells(2, "T") = "Greatest % Increase"
    Cells(3, "T") = "Greatest % Decrease"
    Cells(4, "T") = "Greatest Total Volume"
    Cells(1, "U") = "Ticker"
    Cells(1, "V") = "Value"
    Range("T1").EntireColumn.AutoFit
   
    Dim lastrowDelta As LongLong
    Dim lastrowVolume As LongLong
    
    lastrowDelta = Cells(Rows.Count, "O").End(xlUp).Row
    lastrowVolume = Cells(Rows.Count, "M").End(xlUp).Row
    
    Cells(2, "V") = Cells(2, "Q")
    Cells(2, "V").NumberFormat = "0.00%"
    Cells(3, "V") = Cells(lastrowDelta, "Q")
    Cells(3, "V").NumberFormat = "0.00%"
    Cells(4, "V") = Cells(2, "R")
End Sub

Sub FindTickerForGreatestIncrease()
    'MsgBox ("Ready to find the ticker?")
    Dim lastrow As LongLong
    lastrow = Cells(Rows.Count, "L").End(xlUp).Row
    For Row = 2 To lastrow - 1
        If Cells(Row, "O") = Cells(2, "V") Then
            Cells(2, "U") = Cells(Row, "L")
            Exit For
        'ElseIf Cells(Row, "O") = Cells(3, "V") Then
        End If
    Next Row
End Sub

Sub FindTickerForGreatestDecrease()
    'MsgBox ("Ready to find the ticker?")
    Dim lastrow As LongLong
    lastrow = Cells(Rows.Count, "L").End(xlUp).Row
    For Row = 2 To lastrow - 1
        If Cells(Row, "O") = Cells(3, "V") Then
            Cells(3, "U") = Cells(Row, "L")
            Exit For
        End If
    Next Row
End Sub

Sub FindTickerForGreatestTotalVolume()
    Dim lastrow As LongLong
    lastrow = Cells(Rows.Count, "M").End(xlUp).Row
    For Row = 2 To lastrow - 1
        If Cells(Row, "M") = Cells(4, "V") Then
            Cells(4, "U") = Cells(Row, "L")
            Exit For
        End If
    Next Row
End Sub
