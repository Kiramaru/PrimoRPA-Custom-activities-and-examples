Function Checking_For_Green_Selection(NumberRow As Long) As Boolean
    If Cells(NumberRow + 1, 1).Interior.ColorIndex <> xlNone Then
        Checking_For_Green_Selection = True
    Else
        Checking_For_Green_Selection = False
    End If
End Function