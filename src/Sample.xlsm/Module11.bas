Attribute VB_Name = "Module11"
Option Explicit

Sub �ۑ�10()
    Dim i, j
    j = 1
    For i = 1 To 10
        Worksheets("sheet2").Cells(j, 1) = Worksheets("sheet1").Cells(i, 1) * 10
        j = j + 1
    Next
End Sub
