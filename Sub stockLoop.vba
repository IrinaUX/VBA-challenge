Sub stockLoop():

    ' Setup parameters:
    
        Dim WS As Worksheet
        
'       Loop through each Sheet in the Worksheets:
        For Each WS In ActiveWorkbook.Worksheets
            WS.Activate ' Set current worksheet as active
            Cells(1, "J") = "ticker"
            Cells(1, "K") = "price change"
            Cells(1, "L") = "price change, %"
            Cells(1, "M") = "total"
'           MsgBox (ticker) ' Pop up a message with the ticker text
            
            Dim i As Long
            
'           Set lastRow to know how many loops to run
            
            lastRow = WS.Range("A1").CurrentRegion.Rows.Count
'           MsgBox (lastRow)
'           Loop through the rows in Active worksheet, read ticker and write into Cell (i, "J")
            For i = 2 To lastRow
'                ticker = Worksheets(WS.Name).Cells(i, "A").Value ' Set ticker from the tange of cells in column "A"
'                Cells(i, "J") = ticker
            
'               Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
                Dim priceOpen As Single
                Dim priceClose As Single
                Dim priceChangeOpenClose As Single
                Dim pricePercentChange As Single
                Dim stockVolume As Long
        
                priceOpen = Worksheets(WS.Name).Cells(i, "C").Value
                priceClose = Worksheets(WS.Name).Cells(i, "F").Value
                priceChangeOpenClose = priceClose - priceOpen
                Cells(i, "K") = priceChangeOpenClose
                pricePercentChange = ((priceClose - priceOpen) / priceClose) * 100
                Cells(i, "L") = priceChangeOpenClose
                
                Dim tickerName As String
                tickerName = Cells(2, "A")
                Dim stockVilInitial As Single
                stockVolInitial = Cells(2, "G")
                Cells(2, "M") = stockVolInitial
                
                
'                MsgBox (tickerName & " " & stockVilInitial & " " & stockVolInitial)
                For j = 3 To lastRow
                    If Cells(j, "A") = tickerName Then
                        Dim totalPrevious As Long
                        totalPrevious = Cells((j - 1), "M")
                        'MsgBox ("totalPrevious" & " " & totalPrevious)
                        Dim stock As Long
                        stock = Cells(j, "G")
                        stockVolume = totalPrevious + stock
                        Cells(j, "M") = stockVolume
                        'Exit For
                    Else
                        Cells(j, "A") = tickerName
                            stockVolume = Cells(j, "G") + Cells(j - 1, "G")
                            Cells(j, "M") = stockVolume
                        'Exit For
                    End If
                Next j
                'MsgBox (tickerName)
                
                'stockVolume =
'               MsgBox (priceOpen & " " & priceClose & " " & priceChangeOpenClose & " " & pricePercentChange)
                MsgBox ("totalPrevious" & " = " & totalPrevious & ", total = " & " " & stock)

            
            
            Exit For
            Next i
            
            
            
    
    Next
    

    
    


End Sub


