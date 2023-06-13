Attribute VB_Name = "Module10"
Option Explicit

Sub ‰Û‘è9()
    With Worksheets("sheet1").Range("A1:A2").NumberFormatLocal = "mm/dd"
    End With

    Worksheets("sheet1").Range("A1") = "05/10"
    
    Worksheets("sheet1").Range("A2") = "05/23"
    
    Worksheets("sheet1").Range("A3") = DateDiff("d", Worksheets("sheet1").Range("A1"), Worksheets("sheet1").Range("A2"))
    
End Sub
