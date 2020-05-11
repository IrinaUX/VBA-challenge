Attribute VB_Name = "Module2"
'# VBA Homework - The VBA of Wall Street


'## Instructions
'
'* Create a script that will loop through all the stocks for one year and output the following information:
'

Sub stockData()

Dim Ticker As String
Dim tickerCount As Integer

For Each WS In Worksheets ' loop through all existing worksheets

    Dim LastRow As LongLong
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row ' Find the counter range - number of rows
    
    Dim openPrice As Double
    Dim HighPrice As Double
    Dim LowPrice As Double
    Dim ClosePrice As Double
    
    Dim Volume As LongLong
    Dim YearlyChange As Double
    Dim YearlyChangePercentage As Double
    Dim YearlyVolume As LongLong
    
    'Dim tickerNameHeader As String ' set ticker name in the header of summary table
    Cells(1, "J") = "ticker" ' Set ticker name in the summary table
    Cells(1, "K") = "close-open, $"
    Cells(1, "L") = "close-open, %"
    Cells(1, "M") = "volume"
    
    For Row = 2 To LastRow
    
    If Cells(i + 1, 1) <> Cells(i, 1) Then
    
    
    Else
    
    
    
    Ticker = Cells(2, "A")
    Cells(2, "J") = Ticker
    Cells(2, "K") = Cells(lastRow
    



'    For i = 2 To 10 ' lastRow - 1 ' Loop through all the rows
'        Dim tickerName As String ' set the ticker name
'        tickerName = Cells(i, "A") ' get the ticker name from cell A1
'        openPrice = Cells(i, "C")
'        highPrice = Cells(i, "D")
'        lowPrice = Cells(i, "E")
'        closePrice = Cells(i, "F")
'        volume = Cells(i, "G")
'
'        Dim runningTotal As Integer ' Set ticker counter
'        'tickerCounter = 0 'Initialize ticker counter as zero
'        If Cells((i + 1), "A") = tickerName Then  ' Check if ticker counter is the same and keep counting
'
'        Else:
'            tickerCounter = 1
'            'MsgBox ("not the same counter" & ", - " & tickerCounter)
'        End If
'    'Exit For
'    Next i
'Exit For
'
''  * The ticker symbol.
''
''  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
''
''  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
''
''  * The total stock volume of the stock.
''
''* You should also have conditional formatting that will highlight positive change in green and negative change in red.
''
''* The result should look as follows.
''
''![moderate_solution](Images/moderate_solution.png)
''
''### CHALLENGES
''
''1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
''
''![hard_solution](Images/hard_solution.png)
''
''2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
''
''### Other Considerations
''
''* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.
''
''* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.
''
''## Submission
''
''* To submit please upload the following to Github:
''
''  * A screen shot for each year of your results on the Multi Year Stock Data.
''
''  * VBA Scripts as separate files.
''
''* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.
''
''- - -
''
Next
End Sub


