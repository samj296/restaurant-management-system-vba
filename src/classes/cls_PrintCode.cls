VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_PrintCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function PrintCode(SDish As String, pPrintCode As String)
    Dim PCon As New ADODB.Connection
    Dim PRec As New ADODB.Recordset
    Dim PSQL As String
        
        If SDish <> "" Then
            PSQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Cat]='" & SDish & "'"
            Call Connection_Vindhya_Main_File(PCon)
            PRec.Open PSQL, PCon, adOpenKeyset, adLockOptimistic
                With PRec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        pPrintCode = .Fields("Print Name")
                    End If
                End With
                PRec.Close
                PCon.Close
        End If
End Function
