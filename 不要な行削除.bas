Attribute VB_Name = "Module1"
Option Explicit

Sub ïsóvÇ»çsçÌèú()
  Dim Fir As Integer
  Dim Lst As Integer
  Dim i As Integer
  Dim t As Integer
  Dim myRange As Range
  
  Fir = 6
  Lst = 25
  t = 4
  Range("2:12").EntireRow.Delete
  
  For i = 1 To 100
    Range("A" & Fir & ":" & "B" & Lst).EntireRow.Delete
    Fir = Fir + t
    Lst = Lst + t
    i = i + 1
  Next
    
End Sub

