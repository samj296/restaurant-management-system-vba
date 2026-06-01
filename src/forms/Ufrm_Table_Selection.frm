VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Table_Selection 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "Ufrm_Table_Selection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Table_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_Table_Selection_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim I As Long
    Dim orderCount As Long
    Dim HDish As cls_Dish_Qty_Amount
    Dim Orderlist As New Collection
    Dim Bprint As New cls_Billing_Completion
    Dim HoldNo As String, PrintType As String, SQL As String
    Dim BillNo As String, BillingType As String, Disc As Long, PaymentMode As String, TableNo As String, GST As Long
    Dim CustName As String, MobileNo As String, Tno As Byte 'for updating the TBL_Table_Status in database
    Dim TAmnt As Double, CalcAmnt As Double
    
    BillingType = Ufrm_Billing.cbox_Billing_Type.Text
    
        For I = 0 To lbox_Free_Table_List.ListCount
            If lbox_Free_Table_List.Selected(I) = True Then
                Tno = lbox_Free_Table_List.list(I, 1)
              TableNo = lbox_Free_Table_List.list(I, 0) & " " & lbox_Free_Table_List.list(I, 1)
            End If
        Next I
            
        CustName = Ufrm_Billing.txt_Custmoner_Name.Text
        MobileNo = Ufrm_Billing.txt_Mobile_Number.Text
            If MobileNo = "" Then
                MobileNo = "0"
            End If
            If CustName = "" Then
                CustName = "No Name"
            End If
            
    PrintType = "Hold"
        orderCount = Ufrm_Billing.Lbox_OrderList.ListCount
        For I = 1 To orderCount - 1
            Set HDish = New cls_Dish_Qty_Amount
            
            HDish.dish = Ufrm_Billing.Lbox_OrderList.list(I, 0)
            HDish.DishVariant = Ufrm_Billing.Lbox_OrderList.list(I, 1)
            HDish.QTY = Ufrm_Billing.Lbox_OrderList.list(I, 2)
            HDish.Amount = Ufrm_Billing.Lbox_OrderList.list(I, 3)
            HDish.PrintCode = Ufrm_Billing.Lbox_OrderList.list(I, 4)
            TAmnt = TAmnt + HDish.Amount
            Orderlist.Add HDish
        Next I
        
        Bprint.Bill_Hold Orderlist, BillingType, HoldNo, BillNo, Disc, PaymentMode, TableNo, CustName, MobileNo, TAmnt
        
         SQL = "Select * from TBL_Table_Status"
        
        Call Connection_Vindhya_Billing_Detail(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .BOF = False Or .EOF = False Then
                .MoveFirst
                Do Until .EOF
                        If .Fields("Table Number") = Tno Then
                            .Fields("Status") = "Occupied"
                            .Fields("Customer Name") = CustName
                            .Fields("KOT Hold Number") = HoldNo
                            .Update
                        End If
                    .MoveNext
                Loop
                
            Else
                
                MsgBox "Unable to update in Table List, Please Contact Sam", , "THE VINDHYA CAFE"
                
            End If
            
            
        
        End With
        
        Rec.Close
        Con.Close
            
            TAmnt = Ufrm_Billing.lbl_TAmnt.Caption
            CalcAmnt = Ufrm_Billing.lbl_FAmnt.Caption
            
            Bprint.Hold_Bill_Print PrintType, Orderlist, HoldNo, BillNo, Disc, GST, TableNo, CustName, MobileNo, TAmnt, CalcAmnt
    
            Ufrm_Billing.lbl_Table_Status.Caption = ""
                        
                 Me.Hide
                 
                        
                            
End Sub


Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte
    Dim R As Byte
    
    lbox_Free_Table_List.Clear
    lbox_Free_Table_List.ColumnCount = 2
    lbox_Free_Table_List.ColumnWidths = "42,10"
    lbox_Engage_Table.Clear
    lbox_Engage_Table.ColumnCount = 2
    lbox_Engage_Table.ColumnWidths = "42,10"
    
    
    I = 0
    R = 0
    SQL = "Select * From TBL_Table_Status"
    
        Call Connection_Vindhya_Billing_Detail(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .BOF = False Or .EOF = False Then
                .MoveFirst
                Do Until .EOF
                    If .Fields("Status") = "Free" Then
                        lbox_Free_Table_List.AddItem
                        lbox_Free_Table_List.list(I, 0) = "Table No -"
                        lbox_Free_Table_List.list(I, 1) = .Fields("Table Number")
                        
                        I = I + 1
                    Else
                        lbox_Engage_Table.AddItem
                        lbox_Engage_Table.list(R, 0) = "Table No -"
                        lbox_Engage_Table.list(R, 1) = .Fields("Table Number")
                        R = R + 1
                    End If
                       
                .MoveNext
                Loop
            
            
            End If
        
        End With
        
        Rec.Close
        Con.Close
    
End Sub
