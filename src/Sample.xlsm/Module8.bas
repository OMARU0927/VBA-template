Attribute VB_Name = "Module8"
Option Explicit

Sub 課題7()
    Worksheets("sheet1").Range("A1") = "ジョニー"
    Worksheets("sheet1").Range("A2") = "マイケル"
    Worksheets("sheet1").Range("A3") = "ナンシー"
    Worksheets("sheet1").Range("A4") = "ボブ"
    Worksheets("sheet1").Range("A5") = "エレン"
    
    Dim r        As Range
    Dim i
    Dim j
    Dim iRnd     As Long
    Dim arCell() As String
    Dim s        As String
    
    j = Worksheets("sheet1").Range("A1:A5").Count
    
    ReDim arCell(j - 1)
    
    i = 0
    For Each r In Worksheets("sheet1").Range("A1:A5")
        arCell(i) = r.Value
        
        i = i + 1
    Next
    
    Call Randomize
    
    For i = j - 1 To 1 Step -1
        iRnd = Int((i + 1) * Rnd)
        
        s = arCell(iRnd)
        arCell(iRnd) = arCell(i)
        arCell(i) = s
    Next
    
    Application.ScreenUpdating = False
    
    i = 0
    For Each r In Worksheets("sheet2").Range("A1:A5")
        r.Value = arCell(i)
        i = i + 1
    Next
    
    Application.ScreenUpdating = True
End Sub
