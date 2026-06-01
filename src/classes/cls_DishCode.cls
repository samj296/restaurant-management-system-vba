VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_DishCOde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Dish_DishVariant_Amnt(DishCode As String, dish As String, DishVar As String, Amount As Double, BillType As String)
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long   'For looping
    Dim VarType As String
        
    If IsNumeric(DishCode) = False Then
            
        MsgBox "Invalid dish Code", , "THE VINDHYA CAFE"
        DishCode = ""
        dish = ""
        DishVar = ""
        Exit Function
        
    End If
    
        SQL = "Select * From TBL_Vindhya_Menu WHERE TBL_Vindhya_Menu.[Dish Code]=" & CLng(DishCode)
        Call Connection_Vindhya_Main_File(Con)
            
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .BOF = False Or .EOF = False Then
                .MoveFirst
                
                dish = .Fields("Dish Cat")
                DishVar = .Fields("Variant")
                Amount = .Fields(BillType)
                
                
            End If
            
        End With
        
            
        
        Rec.Close
        Con.Close
    
    
End Function
