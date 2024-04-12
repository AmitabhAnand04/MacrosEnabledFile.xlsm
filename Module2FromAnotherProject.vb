Option Explicit

Sub TestModule2()
    Dim i As Integer
    Dim j As Integer
    j = 2
    For i = 1 To 100
        Sheet1.Cells(i, j).Value = i
    Next i
    
End Sub