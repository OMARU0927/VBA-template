Attribute VB_Name = "Module9"
Option Explicit

Sub ‰Û‘è8()
    Worksheets("sheet2").Range("B1") = Application.WorksheetFunction.Max(Worksheets("sheet1").Range("A1:A10"))
    
    Worksheets("sheet2").Range("B2") = Application.WorksheetFunction.Min(Worksheets("sheet1").Range("A1:A10"))
    
End Sub
