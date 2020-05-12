Sub bubblesort()
Dim tempVariable As Integer



For numPasses = 1 To 3
    For i = 2 To 11

        If Cells(i, 1) < Cells(i - 1, 1) Then

            tempVariable = Cells(i - 1, 1)
            Cells(i - 1, 1) = Cells(i, 1)
            Cells(i, 1) = tempVariable
            Cells(i, 2) = ("Iteration #: " & i)
        Else
            Cells(i, 2) = ("already in good order")
        End If

    Next i
Next numPasses
End Sub
