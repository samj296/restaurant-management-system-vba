VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_calculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    
    

Function Calculate(ByRef CalcNum As String, CalcValue As String) As String
    Dim Old As String
    On Error GoTo ErrorHandling
    Old = shOther.Range("A1").Value
    If Not VBA.Left(CalcNum, 1) = "=" Then
        
        CalcNum = "=" & CalcNum
        
    
    End If
        
    If IsNumeric(VBA.Right(CalcNum, 1)) = False Then
        CalcNum = VBA.Left(CalcNum, Len(CalcNum) - 1)
        
    End If
    
    shOther.Range("A1").Value = CalcNum
    
    CalcValue = shOther.Range("A1").Value
    shOther.Range("A1").Value = Old
Exit Function

ErrorHandling:
    Select Case Err.Number
    
    Case Is = 13
    shOther.Range("A1").Value = Old
    CalcNum = ""
    CalcValue = ""
    
    End Select


End Function
