Attribute VB_Name = "Module8"
Option Explicit

Sub �ۑ�7()
    Worksheets("sheet1").Range("A1") = "�W���j�["
    Worksheets("sheet1").Range("A2") = "�}�C�P��"
    Worksheets("sheet1").Range("A3") = "�i���V�["
    Worksheets("sheet1").Range("A4") = "�{�u"
    Worksheets("sheet1").Range("A5") = "�G����"
    
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
