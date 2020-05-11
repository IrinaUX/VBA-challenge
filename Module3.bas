Attribute VB_Name = "Module3"
Sub stock()

  ' Set an initial variable for holding the brand name
  Dim tickerName As String

  ' Set an initial variable for holding the total per credit card brand
  Dim Brand_Total As LongLong
  Brand_Total = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Find last row
  Dim LastRow As LongLong
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all credit card purchases
  For i = 2 To LastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      tickerName = Cells(i, 1).Value

      ' Add to the Brand Total
      Brand_Total = Brand_Total + Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & Summary_Table_Row).Value = tickerName

      ' Print the Brand Amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = Brand_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Brand_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Brand_Total = Brand_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

