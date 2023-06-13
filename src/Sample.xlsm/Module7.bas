Attribute VB_Name = "Module7"
Option Explicit

Sub ‰Û‘è6()
    Dim i, j
    j = 1
    For i = 1 To 10
        If Worksheets("sheet1").Cells(i, 1).Value Mod 2 = 0 Then
            Worksheets("sheet1").Cells(i, 1).Copy
            Worksheets("sheet2").Cells(j, 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If
        Do While Cells(j, 1) <> ""
            j = j + 1
        Loop
    Next
    
End Sub
