VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Function Vindhya_Billing_Detail(ByRef DataCon As ADODB.Connection, DataRec As ADODB.Recordset, SQL As String)
    Set DataCon = New ADODB.Connection
    Set DataRec = New ADODB.Recordset
    
        Call Connection_Vindhya_Billing_Detail(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
End Function


Function Vindhya_CashFlow(ByRef DataCon As ADODB.Connection, DataRec As ADODB.Recordset, SQL As String)
    Set DataCon = New ADODB.Connection
    Set DataRec = New ADODB.Recordset
    
        Call Connection_Vindhya_Cash_Flow(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
End Function


Function Vindhya_Main_File(ByRef DataCon As ADODB.Connection, DataRec As ADODB.Recordset, SQL As String)
    Set DataCon = New ADODB.Connection
    Set DataRec = New ADODB.Recordset
    
        Call Connection_Vindhya_Main_File(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
    
End Function
