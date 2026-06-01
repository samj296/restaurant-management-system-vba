VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Billing 
   Caption         =   "THE VINDHYA CAFE (Billing)"
   ClientHeight    =   13220
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   18855
   OleObjectBlob   =   "Ufrm_Billing.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Add_Order_Click()
    On Error GoTo Error_Handler
    Dim Order As cls_Dish_Qty_Amount
    Dim I As Byte
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
        
'        Lbox_OrderList.AddItem
'        Lbox_OrderList.List(0, 0) = "Dish"
'        Lbox_OrderList.List(0, 1) = "Variant"
'        Lbox_OrderList.List(0, 2) = "QTY"
'        Lbox_OrderList.List(0, 3) = "Amount"
'        Lbox_OrderList.List(0, 4) = "Print Code"
    
        
    I = Lbox_OrderList.ListCount
    
    'Listbpx start from 0 the header is zero and count start from 1, no need to subtract 1 now
    
    If cbox_Billing_Type.Text <> "" And cbx_Dish_Variant.Text <> "" And _
    cbox_Dish.Text <> "" Then
        If Not cbox_Dish.Text = "" Or Not cbx_Dish_Variant.Text = "" Then
            Set Order = New cls_Dish_Qty_Amount
            Order.dish = cbox_Dish.Text
            Order.DishVariant = cbx_Dish_Variant.Text
            Order.PrintCode = lbl_Print_Code.Caption
            Order.Amount = lbl_dish_Amount.Caption
            Order.QTY = txt_QTY.Text
            
            SQL = "Select * From TBL_Vindhya_Menu WHERE TBL_Vindhya_Menu.[Dish Cat]='" & Order.dish & "' AND TBL_Vindhya_Menu.[Variant]='" & Order.DishVariant & "'"
            Call Connection_Vindhya_Main_File(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                    Order.DiscStatus = Rec.Fields("Disc Status")
                
            Rec.Close
            Con.Close
            
            Lbox_OrderList.AddItem
            Lbox_OrderList.list(I, 0) = Order.dish
            Lbox_OrderList.list(I, 1) = Order.DishVariant
            Lbox_OrderList.list(I, 2) = Order.QTY
            Lbox_OrderList.list(I, 3) = Order.Amount
            Lbox_OrderList.list(I, 4) = Order.PrintCode
            Lbox_OrderList.list(I, 5) = Order.DiscStatus
            
            
            cbx_Dish_Variant.Text = ""
            cbox_Dish.Text = ""
            lbl_dish_Amount.Caption = ""
            txt_QTY.Text = 1
            cbox_Dish_Code.Text = ""
            lbl_dish_Amount.Caption = ""
            cbox_Dish.Text = ""
            cbx_Dish_Variant.Text = ""
            lbl_FAmnt.Caption = ""
                
                lbl_Tamnt_Header = "Total Amount"
            
            If lbl_TAmnt.Caption <> "" Then
                
                lbl_TAmnt.Caption = CLng(lbl_TAmnt.Caption) + Order.Amount
            
            Else
                            
                lbl_TAmnt.Caption = Order.Amount
                
            End If
               
                
              Call txt_Disc_Change
                
                
        Else
            MsgBox "Please Select the proper Dish and Variant to add the order to the order list", , "THE VINDHYA CAFE"
        End If
        
        
        
        
    Else
        MsgBox "Please enter all the deatil before adding the order", , "THE VINDHYA CAFE"
        Exit Sub
    End If
    
    Exit Sub
    
Error_Handler:
  
  Select Case Err.Number
    
    Case Else
    
        MsgBox "Some error has occured, and the error number is - " & Err.Number & ", and the error discription is - " & Err.Description & ".", , "THE VINDHYA CAFE"
        If Me.Visible = False Then
            Me.Show
        End If
    
  End Select
    
    
End Sub

Private Sub btn_Back_Click()
    Me.Hide
    Ufrm_DashBoard.Show
End Sub



Private Sub btn_Cancel_Bill_Click()
    If shMain.Range("F7").Value = "Admin" Then
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            Dim Reason As String
            Dim Bno As String
            
            Call Connection_Vindhya_Billing_Detail(Con)
                Bno = "Vindhya/" & Excel.Application.InputBox("Please enter the Billno.(Only number)-", "THE VINDHYA CAFE", , , , , , 2)
                SQL = "Select * From TBL_Bill_Total Where TBL_Bill_Total.[Bill Number]='" & Bno & "'"
                Reason = Excel.Application.InputBox("Please enter the reason for cancelling the bill", "THE VINDHYA CAFE", , , , , , 2)
                
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .Fields("Status") = "Cancel"
                        If Reason <> "" Then
                            .Fields("Reference") = Reason
                        End If
                        .Update
                        MsgBox "Bill cancelled successfully.", , "THE VINDHYA CAFE"
                    Else
                        MsgBox "Unable to find the bill detail please enter the bill no. correctly.", , "THE VINDHYA CAFE"
                    End If
                End With
                
                Rec.Close
                Con.Close
    Else
        MsgBox "You are not authorise to cancel the bill.", , "THE VINDHYA CAFE"
        Exit Sub
    End If
End Sub

Private Sub btn_Cancel_Hold_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte
    Dim TableNo As Byte
    Dim HoldNo As String
    Dim HoldStat As Boolean
    Dim Old As Byte
    
    'Table Status need to update
    'Table TBL_Hold_Detail need to update
    'Table TBL_Hold_Total need to update
    HoldStat = False
        For I = 1 To lbox_Hold_Table_List.ListCount
            If lbox_Hold_Table_List.Selected(I) = True Then
                If lbox_Hold_Table_List.list(I, 2) <> "" Then
                    TableNo = CByte(lbox_Hold_Table_List.list(I, 0))
                    HoldNo = lbox_Hold_Table_List.list(I, 2)
                    HoldStat = True
                    Old = I
                End If
            End If
        Next I
        
        If HoldStat = True Then
            Call Connection_Vindhya_Billing_Detail(Con)
            SQL = "Select * From TBL_Table_Status Where TBL_Table_Status.[Table Number]=" & TableNo
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .Fields("Status") = "Free"
                    .Fields("Customer Name") = ""
                    .Fields("KOT Hold Number") = ""
                    .Update
                End If
            End With
            
            Rec.Close
            SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "'"
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                            .Fields("Status") = "Cancelled"
                            .Update
                            .MoveNext
                        Loop
                        
                    End If
                End With
            Rec.Close
            SQL = "Select * From TBL_KOT_HOLD_TOTAL WHERE TBL_KOT_HOLD_TOTAL.[KOT Hold Number]='" & HoldNo & "'"
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        .Fields("Status") = "Cancelled"
                        .Update
                    End If
                End With
            Rec.Close
            Con.Close
        End If
        
        Call UserForm_Initialize
        lbox_Hold_Table_List.Selected(Old) = True
End Sub

Private Sub btn_Duplicate_Bill_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim BillUser As String, Printcodecls As cls_PrintCode
    Dim Orderlist As New Collection, dish As cls_Dish_Qty_Amount, Bill As New cls_Billing_Completion
    Dim PrintType As String, HoldNo As String, BillNo As String, Disc As Long, GST As Long, TableNo As String, _
        CustName As String, MobileNo As String, TAmnt As Double, CalcAmnt
        
        BillUser = Excel.Application.InputBox("Enter the Bill no.(Only No)-", "THE VINDHYA CAFE", , , , , , 2)
        
        PrintType = "Duplicate Bill"
        
        BillNo = "Vindhya/" & BillUser
        TableNo = ""
        SQL = "Select * From TBL_Bill_Total Where TBL_Bill_Total.[Bill Number]='" & BillNo & "'"
        Call Connection_Vindhya_Billing_Detail(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    HoldNo = .Fields("KOT Hold Number")
                    If IsNull(.Fields("Disc Amount")) = False Then
                        Disc = .Fields("Disc Amount")
                    Else
                        Disc = 0
                    End If
                    
                    CalcAmnt = .Fields("Total Amount")
                    TAmnt = .Fields("Total Amount") + Disc
                    MobileNo = .Fields("Mobile No")
                Else
                 MsgBox "No record found.", , "THE VINDHYA CAFE"
                 Exit Sub
                End If
            End With
        Rec.Close
        SQL = "Select * From TBL_Bill_Detailed Where TBL_Bill_Detailed.[Bill Number]='" & BillNo & "'"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    Do Until .EOF
                        Set dish = New cls_Dish_Qty_Amount
                        Set Printcodecls = New cls_PrintCode
                        dish.DishVariant = .Fields("Food Item")
                        dish.Amount = .Fields("Amount")
                        Printcodecls.PrintCode dish.DishVariant, dish.PrintCode
                        Orderlist.Add dish
                        .MoveNext
                    Loop
                End If
                
            End With
            Rec.Close
            Con.Close
            SQL = "Select * From TBL_Customer_Detail Where TBL_Customer_Detail.[Customer Mobile Number]='" & MobileNo & "'"
            Call Connection_Vindhya_Main_File(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        CustName = .Fields("Customer Name")
                    End If
                End With
        
        Rec.Close
        Con.Close
        
        Bill.Hold_Bill_Print PrintType, Orderlist, HoldNo, BillNo, Disc, GST, TableNo, CustName, MobileNo, TAmnt, CalcAmnt
        
End Sub

Private Sub btn_Duplicate_KOT_Print_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim KOT As String
    Dim Orderlist As New Collection, dish As cls_Dish_Qty_Amount, Bill As New cls_Billing_Completion
    Dim PrintType As String, HoldNo As String, BillNo As String, Disc As Long, GST As Long, TableNo As String, _
        CustName As String, MobileNo As String, TAmnt As Double, CalcAmnt
        
        KOT = Excel.Application.InputBox("Enter the HOLD no.(Only No)-", "THE VINDHYA CAFE", , , , , , 2)
        HoldNo = "Vindhya/" & KOT
        SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "'"
        
        Call Connection_Vindhya_Billing_Detail(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    Do Until .EOF
                        If IsNull(.Fields("Table Number")) = False Then
                            TableNo = .Fields("Table Number")
                        End If
                        
                        Set dish = New cls_Dish_Qty_Amount
                        dish.dish = .Fields("Item")
                        dish.QTY = .Fields("QTY")
                        dish.PrintCode = .Fields("Print Code")
                        dish.DishVariant = .Fields("Variant")
                        dish.Amount = .Fields("Amount")
                        Orderlist.Add dish
                        .MoveNext
                    Loop
                Else
                    MsgBox "No record found for the given detail.", , "THE VINDHYA CAFE"
                    Exit Sub
                End If
            End With
            
            Rec.Close
            
            SQL = "Select * From TBL_KOT_HOLD_Total Where TBL_KOT_HOLD_Total.[KOT Hold Number]='" & HoldNo & "'"
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With Rec
                    If .EOF = False Or .BOF = False Then
                        CustName = .Fields("Cust Name")
                    End If
                End With
                
            Rec.Close
            Con.Close
            
        
        PrintType = "Duplicate Hold"
        BillNo = ""
        Disc = 0
        GST = 0
        TAmnt = 0
        CalcAmnt = 0
        
        Bill.Hold_Bill_Print PrintType, Orderlist, HoldNo, BillNo, Disc, GST, TableNo, CustName, MobileNo, TAmnt, CalcAmnt
        
        
    
End Sub

Private Sub btn_Hold_Add_Order_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte
    Dim HoldNo As String
    Dim TableNo As String
    Dim Amnt As Double
                    
'                Lbox_Hold_OrderList.List(0, 0) = "Dish"
'                Lbox_Hold_OrderList.List(0, 1) = "Variant"
'                Lbox_Hold_OrderList.List(0, 2) = "QTY"
'                Lbox_Hold_OrderList.List(0, 3) = "Amount"
'                Lbox_Hold_OrderList.List(0, 4) = "Print Code"
                
                On Error GoTo Error_Handler
        
        If cbox_Hold_Dish.Text <> "" And cbx_Hold_Dish_Variant.Text <> "" Then
            
                    For I = 0 To lbox_Hold_Table_List.ListCount
                        
                        If lbox_Hold_Table_List.Selected(I) = True Then
                            TableNo = lbox_Hold_Table_List.list(I, 0)
                            HoldNo = lbox_Hold_Table_List.list(I, 2)
                            GoTo NextStep
                        End If
                    
                    Next I

            If HoldNo = "" Or TableNo = "" Then
                MsgBox "Dish can be add to occupied table only", , "THE VINDHYA CAFE"
                Exit Sub
            End If
            
            
            
NextStep:
                Amnt = CDbl(lbl_Hold_dish_Amount.Caption)
                    I = Lbox_Hold_OrderList.ListCount + 1
                SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "'"

                 Call Connection_Vindhya_Billing_Detail(Con)

                 Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic

                    With Rec
                        .AddNew
                        .Fields("Order Date") = VBA.Format(Now, "mm/dd/yyyy")
                        .Fields("KOT Hold Number") = HoldNo
                        .Fields("Table Number") = TableNo
                        .Fields("Item") = cbox_Hold_Dish.Text
                        .Fields("Variant") = cbx_Hold_Dish_Variant.Text
                        .Fields("QTY") = txt_Hold_QTY.Text
                        .Fields("Amount") = lbl_Hold_dish_Amount.Caption
                        .Fields("Print Code") = lbl_Hold_Print_Code.Caption
                        .Fields("Status") = "Pending"
                        .Update
                        

                    End With


                 Rec.Close
            
                SQL = "Select * From TBL_KOT_HOLD_Total Where TBL_KOT_HOLD_Total.[KOT Hold Number]='" & HoldNo & "'"
                
                    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                    
                    With Rec
                        If .EOF = False Or .BOF = False Then
                            .MoveFirst
                            Amnt = Amnt + CDbl(.Fields("Total Amount"))
                            .Fields("Total Amount") = Amnt
                            .Update
                        End If
                        
                    End With
                    
                    cbox_Hold_Dish_Code.Text = ""
                    cbox_Hold_Dish.Text = ""
                    cbx_Hold_Dish_Variant.Text = ""
                    lbl_Hold_dish_Amount.Caption = ""
                    txt_Hold_QTY.Text = ""
                    
                    
                 Call lbox_Hold_Table_List_Change
                 
                 Con.Close
           
        End If
Exit Sub

Error_Handler:
  
  Select Case Err.Number
    
    Case Else
    
        MsgBox "Some error has occured, and the error number is - " & Err.Number & ", and the error discription is - " & Err.Description & ".", , "THE VINDHYA CAFE"
        If Me.Visible = False Then
            Me.Show
        End If
    
  End Select

End Sub

Private Sub btn_Hold_Proceed_Click()

        On Error GoTo Error_Handler
Dim Orderlist As Collection, Btype As String, Disc As Long, DiscType As String, _
DiscPercent As String, PaymentMode1 As String, PaymentMode2 As String, Mode1Amnt As Double, Mode2Amnt As Double, _
BillNo As String, HoldNo As String, Ref As String, _
Amnt As Double, CalcAmnt As Double
    
   Dim PrintType As String, GST As Long, TableNo As String, Bill As New cls_Billing_Completion, _
CustName As String, MobileNo As String, I As Byte, Order As cls_Dish_Qty_Amount



    If txt_Hold_Mobile_No.Text <> "" And txt_Hold_Custmoner_Name <> "" And Lbox_Hold_OrderList.ListCount > 1 Then
        
        
        Set Orderlist = New Collection
        
        For I = 1 To Lbox_Hold_OrderList.ListCount - 1
            
            Set Order = New cls_Dish_Qty_Amount
                Order.dish = Lbox_Hold_OrderList.list(I, 0)
                Order.DishVariant = Lbox_Hold_OrderList.list(I, 1)
                Order.QTY = Lbox_Hold_OrderList.list(I, 2)
                Order.Amount = Lbox_Hold_OrderList.list(I, 3)
                Order.PrintCode = Lbox_Hold_OrderList.list(I, 4)
                
                Orderlist.Add Order
            
            
            
        Next I
        
        'For Hold No
        
        For I = 1 To lbox_Hold_Table_List.ListCount - 1
                If lbox_Hold_Table_List.Selected(I) = True And lbox_Hold_Table_List.list(I, 1) = "Occupied" Then
                    TableNo = lbox_Hold_Table_List.list(I, 0)
                    HoldNo = lbox_Hold_Table_List.list(I, 2)
                    GoTo Proceed
                End If
            
        
        Next I
            
        




    Else
    
            MsgBox "The Table Selected has no order.", , "THE VINDHYA CAFE"
        
    End If
        
    Exit Sub
    
Proceed:
        
'billing type
            
            Btype = "Bill"
        
            
            Amnt = lbl_Hold_TAmnt.Caption
            CalcAmnt = lbl_Hold_FAmnt.Caption
            CustName = txt_Hold_Custmoner_Name.Text
            MobileNo = txt_Hold_Mobile_No.Text
            
        If lbl_Hold_GST.Caption <> "" Then
            GST = lbl_Hold_GST
        Else
            GST = 0
        End If
            
            'disc type
        If txt_Hold_Disc.Text <> "" Then
            Disc = lbl_Hold_Disc_Amnt.Caption
            If lbl_Hold_Disc_Type.Caption = shOther.Range("E2").Value Then
                DiscPercent = txt_Hold_Disc.Text
            ElseIf lbl_Disc_Type.Caption = shOther.Range("E1").Value Then
                DiscPercent = (lbl_Hold_FAmnt.Caption - GST) / txt_Hold_Disc.Text
            End If
        Else
            Disc = 0
            DiscPercent = 0
        End If
        
        If lbl_Hold_Disc_Type.Caption = shOther.Range("E1").Value Then
            
            DiscType = "Rs"
            
        ElseIf lbl_Hold_Disc_Type.Caption = shOther.Range("E2").Value Then
        
            DiscType = "%"
            
        End If
        
        If txt_Hold_PaymentMode1.Text <> "" And IsNumeric(txt_Hold_PaymentMode1.Text) = True Then
                Mode1Amnt = CDbl(txt_Hold_PaymentMode1.Text)
                PaymentMode1 = cbox_Hold_PaymentMode1.Text
        ElseIf txt_Hold_PaymentMode1.Text = "" Then
                
                MsgBox "Payment Mode 1 can't be empty", , "THE VINDHYA CAFE"
                Exit Sub
        ElseIf cbox_Hold_PaymentMode1.Text = "" Then
        
                MsgBox "Select the paymen mode"
                Exit Sub
        Else
        
                MsgBox "Invalid amount enter.", , "THE VINDHYA CAFE"
                Exit Sub
        End If
        
        If cbox_Hold_PaymentMode2 <> "" And IsNumeric(txt_Hold_PaymentMode2.Text) = True Then
            Mode2Amnt = CDbl(txt_Hold_PaymentMode2.Text)
            PaymentMode2 = cbox_Hold_PaymentMode2.Text
        ElseIf cbox_Hold_PaymentMode2.Text = "" And IsNumeric(txt_Hold_PaymentMode2.Text) = True Then
            
            MsgBox "Payment mode - 2 is empty, select the payment mode from the drop down list", , "THE VINDHYA CAFE"
            
            Exit Sub
        ElseIf cbox_Hold_PaymentMode2.Text <> "" And txt_Hold_PaymentMode2.Text = "" Then
            MsgBox "Amount for payment mode - 2 is empty.", , "THE VINDHYA CAFE"
            Exit Sub
        ElseIf cbox_Hold_PaymentMode2.Text <> "" And IsNumeric(txt_Hold_PaymentMode2.Text) = False Then
            MsgBox "Amount for payment mode - 2 is invalid.", , "THE VINDHYA CAFE"
        Else
            Mode2Amnt = 0
        End If
        
            
        
        Select Case True
                    
            Case Is = Mode1Amnt + Mode2Amnt > CalcAmnt
                MsgBox "Return change  " & (Mode1Amnt + Mode2Amnt) - CalcAmnt
            Case Is = Mode1Amnt + Mode2Amnt < CalcAmnt
            
                MsgBox "Amount is less than the bill amount, update the payment mode amount to proceed.", , "THE VINDHYA CAFE"
                    Exit Sub
        End Select
        
            Ref = "Mode1 - (" & txt_Hold_Ref1.Text & ") Mode2 - (" & txt_Hold_Ref2.Text & ")"
            
        Bill.Finalise_Billing Orderlist, Btype, Disc, DiscType, DiscPercent, _
        PaymentMode1, PaymentMode2, Mode1Amnt, Mode2Amnt, BillNo, HoldNo, MobileNo, Ref, Amnt, CalcAmnt, GST, CustName
        
        Bill.Hold_Bill_Print Btype, Orderlist, HoldNo, BillNo, Disc, GST, TableNo, CustName, MobileNo, Amnt, CalcAmnt
    
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select * From TBL_Table_Status Where TBL_Table_Status.[KOT Hold Number]='" & HoldNo & "'"
    
    Call Connection_Vindhya_Billing_Detail(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
        With Rec
            If .BOF = False Or .EOF = False Then
                .Fields("KOT Hold Number") = ""
                .Fields("Customer Name") = ""
                .Fields("Status") = "Free"
                .Update
            
            End If
        
        End With
            
            
            Rec.Close
            Close
            
        Call UserForm_Initialize
        Exit Sub
        
Error_Handler:
  
  Select Case Err.Number
  
    
    Case Else
    
        MsgBox "Some error has occured, and the error number is - " & Err.Number & ", and the error discription is - " & Err.Description & ".", , "THE VINDHYA CAFE"
        If Me.Visible = False Then
            Me.Show
        End If
    
  End Select

End Sub

Private Sub btn_Proceed_Click()
    
    Dim Num As String, SQL As String, Con As New ADODB.Connection, Rec As New ADODB.Recordset, I As Long, dish As cls_Dish_Qty_Amount, Bill As New cls_Billing_Completion
    Num = txt_Mobile_Number.Text
    
    Dim Orderlist As Collection, BillingMode As String, _
HoldNo As String, BillNo As String, Disc As Long, DiscPercent As String, _
PaymentMode As String, TableNumber As String, CustName As String, MobileNo As String, _
TAmnt As Double, PrintType As String, GST As Long, CalcAmnt As Double, PaymentMode1 As String, _
Mode1Amnt As Double, PaymentMode2 As String, Mode2Amnt As Double, Ref As String, Btype As String, _
DiscType As String, GSTStat As Boolean, GSTPercent As Byte
    
'   PrintType As String, OrderList As Collection, HoldNo As String, BillNo As String, Disc As Long, GST As Long, TableNo As String, _
'CustName As String, MobileNo As String
    
            On Error GoTo Error_Handler
    If cbox_Billing_Type.Text = "DineIn" Then
        If Lbox_OrderList.ListCount > 0 Then
                GoSub CustAdd
                lbl_Table_Status.Caption = "Pending"
                Ufrm_Table_Selection.Show
                        
                    If lbl_Table_Status.Caption = "" Then
                        
                    
                        Call UserForm_Initialize
                    Else
                        MsgBox "Unable to proceed, either the table is not selected or some error has occurred please contact Sam", , "THE VINDHYA CAFE"
                        Exit Sub
                    End If
                    
                  Exit Sub
        Else
            MsgBox "No Order to proceed", , "THE VINDHYA CAFE"
            
            Exit Sub
        End If
          
    Else
        'For Online and takeaway
          If Lbox_OrderList.ListCount > 0 Then
                If cbox_Billing_Type = "Take Away" Then
                    If txt_PaymentMode1.Text <> "" And IsNumeric(txt_PaymentMode1.Text) = True Then
                        Mode1Amnt = CDbl(txt_PaymentMode1.Text)
                        PaymentMode1 = cbox_PaymentMode1.Text
                    ElseIf IsNumeric(txt_PaymentMode1.Text) = False Then
                        MsgBox "Invalid amnt in Payment mode - 1.", , "THE VINDHYA CAFE"
                        Exit Sub
                    End If
                    
                    If txt_PaymentMode2.Text <> "" And IsNumeric(txt_PaymentMode2.Text) = True Then
                        Mode2Amnt = CDbl(txt_PaymentMode2.Text)
                        PaymentMode2 = cbox_PaymentMode2.Text
                    Else
                        Mode2Amnt = 0
                    End If
                    
                    If cbox_PaymentMode1 = "" Or txt_PaymentMode1.Text = "" Then
                        MsgBox "Paymnet mode - 1 can't be empty", , "THE VINDHYA CAFE"
                        Exit Sub
                    End If
                    
                    If Mode1Amnt + Mode2Amnt > lbl_FAmnt.Caption Then
                        MsgBox "Return - " & CDbl(lbl_FAmnt.Caption) - (Mode1Amnt + Mode2Amnt)
                    ElseIf Mode1Amnt + Mode2Amnt < lbl_FAmnt.Caption Then
                        MsgBox "Payment is less than the bill amount.", , "THE VINDHYA CAFE"
                        Exit Sub
                        
                    End If
                    If txt_Mobile_Number.Text = "" Then
                        MobileNo = "0"
                        CustName = "No Name"
                    End If
                    
                    GoSub CustAdd
                    BillingMode = "Take Away"
                    
                Else
                    BillingMode = "Online"
                    Ref = txt_Custmoner_Name.Text
                End If
                    Set Orderlist = New Collection
                
                For I = 1 To Lbox_OrderList.ListCount - 1
                    Set dish = New cls_Dish_Qty_Amount
                        dish.dish = Lbox_OrderList.list(I, 0)
                        dish.DishVariant = Lbox_OrderList.list(I, 1)
                        dish.QTY = Lbox_OrderList.list(I, 2)
                        dish.Amount = Lbox_OrderList.list(I, 3)
                        dish.PrintCode = Lbox_OrderList.list(I, 4)
                        
                        Orderlist.Add dish
                    
                Next I
                
                If lbl_Disc_Type.Caption = shOther.Range("E1").Value Then
                    DiscType = "Rs"
                ElseIf lbl_Disc_Type.Caption = shOther.Range("E2").Value Then
                    DiscType = "%"
                End If
                
                Disc = txt_Disc.Text
                Btype = cbox_Billing_Type.Text
                TAmnt = lbl_TAmnt.Caption
                CustName = txt_Custmoner_Name.Text
                MobileNo = txt_Mobile_Number.Text
                
                
                Bill.Bill_Hold Orderlist, BillingMode, HoldNo, BillNo, Disc, PaymentMode, TableNumber, CustName, MobileNo, TAmnt
                
                Bill.Disc_GST_Calculation Orderlist, TAmnt, CalcAmnt, Disc, DiscType, DiscPercent, GSTStat, GSTPercent, GST, Btype
                
                PrintType = "Hold"
                Bill.Hold_Bill_Print PrintType, Orderlist, HoldNo, BillNo, Disc, GST, TableNumber, CustName, MobileNo, TAmnt, CalcAmnt
                 
                 Bill.Finalise_Billing Orderlist, BillingMode, Disc, DiscType, DiscPercent, _
                 PaymentMode1, PaymentMode2, Mode1Amnt, Mode2Amnt, BillNo, HoldNo, Num, Ref, TAmnt, CalcAmnt, GST, CustName
                 
                PrintType = "Bill"
                Bill.Hold_Bill_Print PrintType, Orderlist, HoldNo, BillNo, Disc, GST, TableNumber, CustName, MobileNo, TAmnt, CalcAmnt
                
                Call UserForm_Initialize

          End If
            
            
        Exit Sub
    
    End If
    
    
CustAdd:
            
            SQL = "SELECT * FROM TBL_Customer_Detail WHERE TBL_Customer_Detail.[Customer Mobile Number] Like '" & Num & "'"
            
        
            If Len(txt_Mobile_Number.Text) <> 10 And Len(txt_Mobile_Number.Text) <> 0 Then
                Dim UserRes As String
                
                UserRes = MsgBox("The number you have entered is not 10 digit. Do you want to proceed", vbYesNo, "THE VINDHYA CAFE")
                
                If UserRes = 6 Then 'if the user press yes
                    GoTo NextStep
                ElseIf UserRes = 7 Then
                    Exit Sub
                Else
                    Exit Sub
                End If
                
                
            Else
                
NextStep:
                Call Connection_Vindhya_Main_File(Con)
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .BOF = True And .EOF = True Then
                        .AddNew
                        .Fields("Customer Name") = txt_Custmoner_Name.Text
                        .Fields("Customer Mobile Number") = txt_Mobile_Number.Text
                        .Update
                        
                    End If
                    
                End With
                
                Rec.Close
            
            
            End If
            
            Con.Close
            
            Return
            
            Exit Sub
    
Error_Handler:
  
  Select Case Err.Number
    
    Case Else
    
        MsgBox "Some error has occured, and the error number is - " & Err.Number & ", and the error discription is - " & Err.Description & ".", , "THE VINDHYA CAFE"
        If Me.Visible = False Then
            Me.Show
        End If
    
  End Select

    
End Sub

Private Sub btn_Remove_Dish_Click()
    Dim I As Long
    Dim Amnt As Double
        If Lbox_OrderList > 1 Then
            For I = 1 To Lbox_OrderList.ListCount - 1
                If Lbox_OrderList.Selected(I) = True Then
                    Amnt = Lbox_OrderList.list(I, 3)
                    Lbox_OrderList.RemoveItem (I)
                    lbl_TAmnt.Caption = CDbl(lbl_TAmnt.Caption) - Amnt
                End If
            Next I
        End If
        Call txt_Disc_Change
End Sub

Private Sub btn_Remove_Dish_Hold_Click()
    Dim I As Long
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim HoldNo As String
    Dim dish As String, DishVar As String, QTY As Byte
    
        For I = 1 To lbox_Hold_Table_List.ListCount - 1
            If lbox_Hold_Table_List.Selected(I) = True Then
                HoldNo = lbox_Hold_Table_List.list(I, 2)
            End If
        Next I
        
        For I = 1 To Lbox_Hold_OrderList.ListCount - 1
            If Lbox_Hold_OrderList.Selected(I) = True Then
                dish = Lbox_Hold_OrderList.list(I, 0)
                DishVar = Lbox_Hold_OrderList.list(I, 1)
                QTY = Lbox_Hold_OrderList.list(I, 2)
            End If
        Next I
    
        SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "' AND TBL_KOT_HOLD_DETAIL.[Item]='" & _
                dish & "' AND TBL_KOT_HOLD_DETAIL.[Variant]='" & DishVar & "' AND TBL_KOT_HOLD_DETAIL.[QTY]=" & QTY
        Call Connection_Vindhya_Billing_Detail(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                .Delete
                
            End If
        End With
        Rec.Close
        Con.Close
        
        Call lbox_Hold_Table_List_Change
End Sub

Private Sub cbox_Billing_Type_Change()
        
        If cbox_Billing_Type.Text = "Online" Then
            lbl_Customer_List = "Online Orders"
            lbl_CutomerName = "Order From"
            lbl_Mobile_number = "Order ID"
            btn_Proceed.Caption = "Bill Print"
            btn_Proceed.Visible = True
DineIn:
            frame_PaymentMode1.Visible = False
            frame_PaymentMode2.Visible = False
            
            
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            Dim I As Byte   'for Looping inside the table
            I = 0
            
            SQL = "Select * From TBL_Payment_Mode"
            
            cbox_Customer_List.Clear
            cbox_Customer_List.ColumnCount = 1
            cbox_Customer_List.ColumnWidths = "140"
            
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        Do Until .EOF
                            If Not .Fields("OnlineSale") = vbNullString Then
                                cbox_Customer_List.AddItem
                                cbox_Customer_List.list(I, 0) = .Fields("OnlineSale")
                                I = I + 1
                            End If
                            
                            .MoveNext
                        Loop
                        
                    End If
                    
                End With
            
                
            Rec.Close
            Con.Close
            
            
        ElseIf cbox_Billing_Type.Text = "DineIn" Then
                btn_Proceed.Caption = "Hold Order"
                btn_Proceed.Visible = True
                lbl_Customer_List = "Customer List"
                lbl_CutomerName = "Customer Name"
                lbl_Mobile_number = "Mobile Number"
                txt_Custmoner_Name.Text = ""
                txt_Mobile_Number.Text = ""
                cbox_Customer_List.Clear
                GoTo DineIn
        
        ElseIf cbox_Billing_Type.Text = "Take Away" Then
        
            frame_PaymentMode1.Visible = True
            frame_PaymentMode2.Visible = True
        
            lbl_Customer_List = "Customer List"
            lbl_CutomerName = "Customer Name"
            lbl_Mobile_number = "Mobile Number"
            txt_Custmoner_Name.Text = ""
            txt_Mobile_Number.Text = ""
            cbox_Customer_List.Clear
            btn_Proceed.Caption = "Bill Print"
            btn_Proceed.Visible = True
        Else
            Call UserForm_Initialize
        End If
End Sub

Private Sub cbox_Customer_List_Change()
        If cbox_Customer_List.Text <> "" And Not cbox_Billing_Type = "Online" Then
        
            txt_Mobile_Number.Text = cbox_Customer_List.Text
            
            
        ElseIf cbox_Billing_Type = "Online" Then
            
            txt_Custmoner_Name.Text = cbox_Customer_List.Text
            
        
        End If
        
End Sub


Private Sub cbox_Dish_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte
    Dim VarType As String
    
        If cbox_Dish.Text = "" Then
        
            Exit Sub
        
        
        ElseIf cbox_Billing_Type.Text = "" Then
            
                cbox_Dish.Text = ""
                MsgBox "Billing type can't be empty.", , "THE VINDHYA CAFE"
                
                Exit Sub
            
        End If
        
    SQL = "Select * From TBL_Vindhya_Dish_Variant WHERE TBL_Vindhya_Dish_Variant.[Dish Cat] = '" & cbox_Dish.Text & "'"
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                VarType = .Fields("Variant Type")
            End If
            
        End With
        
        
        Rec.Close
        cbx_Dish_Variant.Clear
        cbx_Dish_Variant.ColumnCount = 1
        cbx_Dish_Variant.ColumnWidths = "140"
        I = 0
        
    SQL = "Select * From TBL_Variant_Selection"
    
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    If VarType <> "" Then
                        Do Until .EOF
                            If Not .Fields(VarType) = vbNullString Then
                                cbx_Dish_Variant.AddItem
                                cbx_Dish_Variant.list(I, 0) = .Fields(VarType)
                                I = I + 1
                            End If
                            
                            .MoveNext
                        Loop
                    End If
                End If
            
            End With
            
        
        
        Con.Close
        
        
    
End Sub

Private Sub cbox_Dish_Code_Change()
    Dim DCode As String, dish As String, Dvar As String, Amnt As Double, Btype As String
                            'cls_DishCOde Variables
                            'DishCode As String, Dish As String, DishVar As String, Amount As Double, BillType As String
    
    Dim Frm As New cls_DishCOde
    Dim QTY As Long
                
                DCode = cbox_Dish_Code.Text
                
                If DCode = "" Then
                    cbox_Dish.Text = ""
                    cbx_Dish_Variant.Text = ""
                    Exit Sub
                ElseIf cbox_Billing_Type = "" Then
                    MsgBox "Billing type caan't be empty", , "THE VINDHYA CAFE"
                    cbox_Dish_Code.Text = ""
                    Exit Sub
                End If
        


                    
                        
                        Btype = cbox_Billing_Type.Text
                        
                        Frm.Dish_DishVariant_Amnt DCode, dish, Dvar, Amnt, Btype
                        
                        
                        
                        'MsgBox "Dish - " & Dish & vbNewLine & "Dish Code - " & DCode & vbNewLine & "Variant - " & Dvar & vbNewLine & "Amount - " & Amnt
                        
                        cbox_Dish.Text = dish
                        cbx_Dish_Variant.Text = Dvar
                        
                        lbl_dish_Amount = Amnt * txt_QTY.Text
                        If cbox_Dish.Text = "" Then
                            cbox_Dish_Code.Text = DCode
                        End If
                        

    
        
                
End Sub

Private Sub cbox_Hold_Dish_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte
    Dim VarType As String
    
    SQL = "Select * From TBL_Vindhya_Dish_Variant WHERE TBL_Vindhya_Dish_Variant.[Dish Cat] = '" & cbox_Hold_Dish.Text & "'"
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                VarType = .Fields("Variant Type")
            End If
            
        End With
        
        
        Rec.Close
        cbx_Hold_Dish_Variant.Clear
        cbx_Hold_Dish_Variant.ColumnCount = 1
        cbx_Hold_Dish_Variant.ColumnWidths = "140"
        I = 0
        
    SQL = "Select * From TBL_Variant_Selection"
    
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    If VarType <> "" Then
                        Do Until .EOF
                            If Not .Fields(VarType) = vbNullString Then
                                cbx_Hold_Dish_Variant.AddItem
                                cbx_Hold_Dish_Variant.list(I, 0) = .Fields(VarType)
                                I = I + 1
                            End If
                            
                            .MoveNext
                        Loop
                    End If
                End If
            
            End With
            
        
        
        Con.Close
        
        
End Sub

Private Sub cbox_Hold_Dish_Code_Change()
    Dim DCode As String, dish As String, Dvar As String, Amnt As Double, Btype As String
                            'cls_DishCOde Variables
                            'DishCode As String, Dish As String, DishVar As String, Amount As Double, BillType As String
    
    Dim Frm As New cls_DishCOde
    Dim QTY As Long
                
                DCode = cbox_Hold_Dish_Code.Text
                
                If DCode = "" Then
                    Exit Sub
                    cbox_Hold_Dish.Text = ""
                    cbx_Hold_Dish_Variant.Text = ""
                End If
        
                

                    
                        
                        Btype = "DineIn"
                        
                        Frm.Dish_DishVariant_Amnt DCode, dish, Dvar, Amnt, Btype
                        
                        
                        
                        'MsgBox "Dish - " & Dish & vbNewLine & "Dish Code - " & DCode & vbNewLine & "Variant - " & Dvar & vbNewLine & "Amount - " & Amnt
                        
                        cbox_Hold_Dish.Text = dish
                        cbx_Hold_Dish_Variant.Text = Dvar
                        lbl_Hold_Disc_Amnt = Amnt * txt_Hold_QTY.Text
                        If cbox_Hold_Dish.Text = "" Then
                            cbox_Hold_Dish_Code.Text = DCode
                        End If
                        

    
        
End Sub

Private Sub cbox_Hold_PaymentMode1_Change()
    If cbox_Hold_PaymentMode1.Text = "UPI" Then
        lbl_Hold_Ref1.Visible = True
        txt_Hold_Ref1.Visible = True
    Else
        lbl_Hold_Ref1.Visible = False
        txt_Hold_Ref1.Visible = False
        
    End If
End Sub

Private Sub cbox_Hold_PaymentMode2_Change()
   If cbox_Hold_PaymentMode1.Text = "UPI" Then
        lbl_Hold_Ref2.Visible = True
        txt_Hold_Ref2.Visible = True
    Else
        lbl_Hold_Ref2.Visible = False
        txt_Hold_Ref2.Visible = False
        
    End If
End Sub

Private Sub cbox_PaymentMode1_Change()
    If cbox_PaymentMode1.Text = "UPI" Then
        lbl_Ref1.Visible = True
        txt_Ref1.Visible = True
    Else
        lbl_Ref1.Visible = False
        txt_Ref1.Visible = False
        txt_Ref1.Text = ""
    End If
End Sub

Private Sub cbx_Dish_Variant_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim dish As String
    Dim DishVar As String
    Dim BillType As String
    Dim QTY As Long
    
    If cbox_Billing_Type = "" Or cbox_Dish.Text = "" Then
        cbx_Dish_Variant.Text = ""
        MsgBox "Billing Type or Dish column can't be empty", , "THE VINDHYA CAFE"
        cbox_Dish.Text = ""
        cbx_Dish_Variant.Text = ""
        Exit Sub
    End If
    
    dish = cbox_Dish.Text
    DishVar = cbx_Dish_Variant.Text
    BillType = cbox_Billing_Type.Text
    
    If txt_QTY.Text <> "" Then
        QTY = txt_QTY.Text
    Else
        QTY = 1
    End If
    
    SQL = "Select * From TBL_Vindhya_Menu WHERE TBL_Vindhya_Menu.[Dish Cat] = '" & dish & "' AND TBL_Vindhya_Menu.[Variant] = '" _
    & DishVar & "'"
    
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            If .Fields("Status") = True Then
                lbl_dish_Amount.Caption = .Fields(BillType) * QTY
                lbl_Print_Code.Caption = .Fields("Print Name")
            Else
                cbx_Dish_Variant.Text = ""
                MsgBox "This variant has been turned off by the ADMIN.", , "THE VINDHYA CAFE"
                
            End If
        End If
          
    End With
    
    
    
    
End Sub

Private Sub cbx_Hold_Dish_Variant_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim dish As String
    Dim DishVar As String
    Dim BillType As String
    Dim QTY As Long
    
    
    
    dish = cbox_Hold_Dish.Text
    DishVar = cbx_Hold_Dish_Variant.Text
    BillType = "DineIn"
    
    If txt_Hold_QTY.Text <> "" Then
        QTY = txt_Hold_QTY.Text
    Else
        QTY = 1
    End If
    
    SQL = "Select * From TBL_Vindhya_Menu WHERE TBL_Vindhya_Menu.[Dish Cat] = '" & dish & "' AND TBL_Vindhya_Menu.[Variant] = '" _
    & DishVar & "'"
    
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            lbl_Hold_dish_Amount.Caption = .Fields(BillType) * QTY
            lbl_Hold_Print_Code.Caption = .Fields("Print Name")
            
        End If
          
    End With
    
    
    
End Sub



Private Sub lbl_Dec_QTY_Click()
            
            Do Until IsNumeric(VBA.Right(txt_QTY.Text, 1)) = True
                txt_QTY.Text = VBA.Left(txt_QTY.Text, Len(txt_QTY.Text) - 1)
            Loop
            
        If Not txt_QTY.Text = "" And IsNumeric(txt_QTY.Text) = True And IsNumeric(VBA.Right(txt_QTY.Text, 1)) = True Then
            If CByte(txt_QTY.Text) > 1 Then
                txt_QTY.Text = CByte(txt_QTY.Text) - 1
            ElseIf CByte(txt_QTY.Text) < 1 Or CByte(txt_QTY.Text) > 0 Then
                txt_QTY.Text = 1
            End If
        ElseIf Not txt_QTY.Text = "" And IsNumeric(VBA.Right(txt_QTY.Text, 1)) = False Then
            
        Else
            txt_QTY.Text = 1
        End If
End Sub

Private Sub lbl_Disc_Type_Click()
    If lbl_Disc_Type = shOther.Range("E1").Value Then
        lbl_Disc_Type = shOther.Range("E2").Value
    ElseIf lbl_Disc_Type = shOther.Range("E2").Value Then
        lbl_Disc_Type = shOther.Range("E1").Value
    End If
    Call txt_Disc_Change
End Sub



Private Sub lbl_Hold_Dec_QTY_Click()
            
            Do Until IsNumeric(VBA.Right(txt_Hold_QTY.Text, 1)) = True
                txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
            Loop
            
        If Not txt_Hold_QTY.Text = "" And IsNumeric(txt_Hold_QTY.Text) = True And IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True Then
            If CByte(txt_Hold_QTY.Text) > 1 Then
                txt_Hold_QTY.Text = CByte(txt_Hold_QTY.Text) - 1
            ElseIf CByte(txt_Hold_QTY.Text) < 1 Or CByte(txt_Hold_QTY.Text) > 0 Then
                txt_Hold_QTY.Text = 1
            End If
        ElseIf Not txt_Hold_QTY.Text = "" And IsNumeric(Right(txt_Hold_QTY.Text, 1)) = False Then
            
        Else
            txt_Hold_QTY.Text = 1
        End If
End Sub

Private Sub lbl_Hold_Disc_Type_Click()
    If lbl_Hold_Disc_Type.Caption = shOther.Range("E1").Value Then
        lbl_Hold_Disc_Type.Caption = shOther.Range("E2").Value
    ElseIf lbl_Hold_Disc_Type.Caption = shOther.Range("E2").Value Then
        lbl_Hold_Disc_Type.Caption = shOther.Range("E1").Value
    End If
    
    Call lbox_Hold_Table_List_Change
    
End Sub

Private Sub lbl_Hold_Inc_QTY_Click()
    
        If txt_Hold_QTY.Text = "" Then
            txt_Hold_QTY.Text = 1
        End If
        
       
        Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
                txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
        Loop
        
       
        
    If Not txt_Hold_QTY.Text = "" Then
        txt_Hold_QTY.Text = CByte(txt_Hold_QTY.Text) + 1
    Else
    
        txt_Hold_QTY.Text = 2
        
    End If
    
         Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
                txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
        Loop
    
End Sub

Private Sub lbl_Inc_QTY_Click()
        
        If txt_QTY.Text = "" Then
            txt_QTY.Text = 1
        End If
        
       
        
        
       
        
    If Not txt_QTY.Text = "" Then
        txt_QTY.Text = CByte(txt_QTY.Text) + 1
    Else
    
        txt_QTY.Text = 2
        
    End If
    
End Sub



Private Sub lbox_Hold_Table_List_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Byte   'For Looping insidee the table
    Dim HoldNo As String
    Dim MobileNo  As String
    Dim BillCalc As New cls_Billing_Completion
    Dim OrderOut As cls_Dish_Qty_Amount
    
    
   Dim Orderlist As New Collection, TAmnt As Double, _
   CalcAmnt As Double, Disc As Long, DiscType As String, _
   DiscPercent As String, GSTStat As Boolean, GSTPercent As Byte, GST As Long, Btype As String
   
        Btype = "DineIn"
        
        If lbl_Hold_Disc_Type.Caption = shOther.Range("E1").Value Then   'Disc Type is in rupees
        
            DiscType = "Rs"
            
        ElseIf lbl_Hold_Disc_Type.Caption = shOther.Range("E2").Value Then
        
            DiscType = "%"
            
        End If
    
            If txt_Hold_Disc <> "" Then
                Disc = txt_Hold_Disc.Text
                
            End If
            
   '0 is for header
    
    
        For I = 1 To lbox_Hold_Table_List.ListCount
            If lbox_Hold_Table_List.Selected(I) = True Then
                If Not lbox_Hold_Table_List.list(I, 2) = "" Then
                    HoldNo = lbox_Hold_Table_List.list(I, 2)
                    
                Else
                    Lbox_Hold_OrderList.Clear
                    
                    Lbox_Hold_OrderList.AddItem
                    Lbox_Hold_OrderList.list(0, 0) = "Dish"
                    Lbox_Hold_OrderList.list(0, 1) = "Variant"
                    Lbox_Hold_OrderList.list(0, 2) = "QTY"
                    Lbox_Hold_OrderList.list(0, 3) = "Amount"
                    Lbox_Hold_OrderList.list(0, 4) = "Print Code"
                    txt_Hold_Mobile_No.Text = ""
                    txt_Hold_Custmoner_Name.Text = ""
                        lbl_Hold_FAmnt.Caption = ""
                        lbl_Hold_GST.Caption = ""
                        lbl_Hold_Disc_Amnt.Caption = ""
                        
                        lbl_Hold_TAmnt.Caption = ""
                        lbl_Hold_Tamnt_Header.Caption = "Total Amount"
                        lbl_Hold_Tamnt_Header.Visible = False
                    
                    Exit Sub
                End If
                
            End If
        
        Next I
    
NextStep:
        
        I = 1
        SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "'"
        
        Call Connection_Vindhya_Billing_Detail(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                Lbox_Hold_OrderList.Clear
                
                Lbox_Hold_OrderList.AddItem
                Lbox_Hold_OrderList.list(0, 0) = "Dish"
                Lbox_Hold_OrderList.list(0, 1) = "Variant"
                Lbox_Hold_OrderList.list(0, 2) = "QTY"
                Lbox_Hold_OrderList.list(0, 3) = "Amount"
                Lbox_Hold_OrderList.list(0, 4) = "Print Code"
                
            lbl_Hold_FAmnt.Caption = ""
            lbl_Hold_GST.Caption = ""
            lbl_Hold_Disc_Amnt.Caption = ""
            
            lbl_Hold_TAmnt.Caption = ""
            lbl_Hold_Tamnt_Header.Caption = "Total Amount"
            lbl_Hold_Tamnt_Header.Visible = False
        
        With Rec
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                
                Do Until .EOF
                    Set OrderOut = New cls_Dish_Qty_Amount
                    
                    Lbox_Hold_OrderList.AddItem
                    Lbox_Hold_OrderList.list(I, 0) = .Fields("Item")
                    OrderOut.dish = .Fields("Item")
                    
                    Lbox_Hold_OrderList.list(I, 1) = .Fields("Variant")
                    OrderOut.DishVariant = .Fields("Variant")
                    
                    Lbox_Hold_OrderList.list(I, 2) = .Fields("QTY")
                    OrderOut.QTY = .Fields("QTY")
                    
                    Lbox_Hold_OrderList.list(I, 3) = .Fields("Amount")
                    OrderOut.Amount = .Fields("Amount")
                    
                    Lbox_Hold_OrderList.list(I, 4) = .Fields("Print Code")
                    OrderOut.PrintCode = .Fields("Print Code")
                    
                     Orderlist.Add OrderOut
                    TAmnt = TAmnt + CDbl(.Fields("Amount"))
                    I = I + 1
                .MoveNext
                Loop
                
                
            
            End If
            
        
        End With
        
        Rec.Close
        
        SQL = "Select * From TBL_KOT_HOLD_Total WHERE TBL_KOT_HOLD_Total.[KOT Hold Number]='" & HoldNo & "'"
        
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
        With Rec
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                MobileNo = .Fields("Mobile No")
               
            End If
        
        End With
           
           txt_Hold_Mobile_No.Text = MobileNo
           
            
          BillCalc.Disc_GST_Calculation Orderlist, TAmnt, CalcAmnt, Disc, DiscType, DiscPercent, GSTStat, GSTPercent, GST, Btype
          
          lbl_Hold_FAmnt.Caption = CalcAmnt + GST
          lbl_Hold_GST.Caption = GST
          lbl_Hold_Disc_Amnt.Caption = Disc
            
          lbl_Hold_TAmnt.Caption = TAmnt
          lbl_Hold_Tamnt_Header.Caption = "Total Amount"
          lbl_Hold_Tamnt_Header.Visible = True
          
        Rec.Close
        Con.Close
        
        
        
    
End Sub




Private Sub MultiPage1_Change()
    
    Call UserForm_Initialize

    
End Sub




Private Sub txt_Custmoner_Name_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
        If Not cbox_Billing_Type.Text = "Online" Then
            SQL = "Select TBL_Customer_Detail.[Customer Name],TBL_Customer_Detail.[Customer Mobile Number] From TBL_Customer_Detail WHERE TBL_Customer_Detail.[Customer Name] LIKE '%" & txt_Custmoner_Name.Text & "%'"
            Call Connection_Vindhya_Main_File(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                    I = 0
            
                
              cbox_Customer_List.Clear
              cbox_Customer_List.ColumnCount = 2
              cbox_Customer_List.ColumnWidths = "80,120"
              
'                    Call Connection_Vindhya_Main_File(Con)
'                    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                    
                    With Rec
                        If .EOF = False Or .BOF = False Then
                            .MoveFirst
                            
                            Do Until .EOF
                                cbox_Customer_List.AddItem
                                cbox_Customer_List.list(I, 0) = .Fields("Customer Mobile Number")
                                cbox_Customer_List.list(I, 1) = .Fields("Customer Name")
                                I = I + 1
                                .MoveNext
                            Loop
                            
                        End If
                                            
                        
                
                
                
            Rec.Close
            Con.Close
                    End With
        End If
End Sub

Private Sub txt_Disc_Change()

        If IsNumeric(txt_Disc.Text) = False Then
            If txt_Disc.Text <> "" Then
                MsgBox "Disc Should be in numbers.", , "THE VINDHYA CAFE"
                txt_Disc = 0
                Exit Sub
            End If
            
        End If
        
    Dim I As Byte
    Dim Order As cls_Dish_Qty_Amount
    
    
    Set Order = New cls_Dish_Qty_Amount
    
    Order.dish = cbox_Dish.Text
    Order.DishVariant = cbx_Dish_Variant.Text
    Order.Amount = lbl_dish_Amount.Caption
    Order.QTY = txt_QTY.Text
    Order.PrintCode = lbl_Print_Code.Caption
         
                Dim Bill As New cls_Billing_Completion, Orderlist As Collection, TAmnt As Double, _
CalcAmnt As Double, Disc As Long, DiscType As String, _
DiscPercent As String, GSTStat As Boolean, GSTPercent As Byte, GST As Long, Btype As String
                    Set Orderlist = New Collection
                    
                        'Disc Type for calculation of the discount
                    If lbl_Disc_Type.Caption = shOther.Range("E1").Value Then
                        DiscType = "Rs"
                    ElseIf lbl_Disc_Type.Caption = shOther.Range("E2").Value Then
                        DiscType = "%"
                    Else
                        MsgBox "Unable to identify the discount type, contact Sam for rectify this.", , "THE VINDHYA CAFE"
                        Exit Sub
                    End If
                        
                        'Total Amount
                        If lbl_TAmnt.Caption = "" Then
                            lbl_TAmnt.Caption = Order.Amount
                        End If
                        If lbl_TAmnt.Caption <> "" Then
                            TAmnt = lbl_TAmnt.Caption
                        Else
                            TAmnt = 0
                        End If
                        
                        'Disc amnt or percent figure
                        
                        If txt_Disc.Text = "" Then
                            txt_Disc = 0
                        End If
                        
                        Disc = txt_Disc.Text
                        
                        Btype = cbox_Billing_Type.Text
            
                For I = 1 To Lbox_OrderList.ListCount - 1
                    Set Order = New cls_Dish_Qty_Amount
                    
                    
                    Order.dish = Lbox_OrderList.list(I, 0)
                    Order.DishVariant = Lbox_OrderList.list(I, 1)
                    Order.QTY = Lbox_OrderList.list(I, 2)
                    Order.Amount = Lbox_OrderList.list(I, 3)
                    Order.PrintCode = Lbox_OrderList.list(I, 4)
                    
                    Orderlist.Add Order
                    
                
                
                Next I
                
                Bill.Disc_GST_Calculation Orderlist, TAmnt, CalcAmnt, Disc, DiscType, DiscPercent, GSTStat, GSTPercent, GST, Btype
                
                lbl_Disc_Amnt.Caption = Disc
                lbl_GST.Caption = GST
                lbl_FAmnt.Caption = CalcAmnt
    
    
End Sub

Private Sub txt_Hold_Disc_Change()
    If IsNumeric(txt_Hold_Disc.Text) = False Then
        If txt_Hold_Disc.Text = "" Then
            txt_Hold_Disc.Text = 0
        Else
            MsgBox "Disc should be number only", , "THE VINDHYA CAFE"
            txt_Hold_Disc.Text = 0
            Exit Sub
        End If
    End If
    Call lbox_Hold_Table_List_Change
End Sub

Private Sub txt_Hold_Mobile_No_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    Dim Num As String
        
        
        
         Num = txt_Hold_Mobile_No.Text
                    
            SQL = "Select TBL_Customer_Detail.[Customer Name],TBL_Customer_Detail.[Customer Mobile Number] From TBL_Customer_Detail WHERE TBL_Customer_Detail.[Customer Mobile Number]='" & Num & "'"
            
        
           
                Call Connection_Vindhya_Main_File(Con)
                txt_Custmoner_Name.Text = ""
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        txt_Hold_Custmoner_Name.Text = .Fields("Customer Name")
                        
                    End If
                    
                End With
                
                Rec.Close
            
            
    
            
            Con.Close
            
       
End Sub

Private Sub txt_Hold_QTY_Change()
     Dim QTY As String
                        
            QTY = txt_Hold_QTY.Text
            If Not txt_Hold_QTY.Text = "" Then
                Do Until IsNumeric(VBA.Right(QTY, 1)) = True
                    txt_Hold_QTY.Text = Left(QTY, Len(QTY) - 1)
                    QTY = txt_Hold_QTY.Text
                    
                Loop
            Else
                txt_Hold_QTY.Text = 1
            End If
            
            If cbox_Hold_Dish.Text <> "" And cbx_Hold_Dish_Variant <> "" Then
                Call cbx_Hold_Dish_Variant_Change
            End If
End Sub

Private Sub txt_Hold_QTY_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    Dim QTY As String
        
        QTY = txt_Hold_QTY.Text
   
    If QTY = "" Then
    
        txt_Hold_QTY.Text = 1
        
    End If
              
               
    
    If KeyAscii = 43 Then
        
        Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
            txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
        Loop
        
        Call lbl_Hold_Inc_QTY_Click
            
        Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
            txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
        Loop
        
    ElseIf KeyAscii = 45 Then
        
        Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
            txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
        Loop
        
        Call lbl_Hold_Dec_QTY_Click
            
         Do Until IsNumeric(Right(txt_Hold_QTY.Text, 1)) = True
            txt_Hold_QTY.Text = Left(txt_Hold_QTY.Text, Len(txt_Hold_QTY.Text) - 1)
         Loop
            
    End If
    
       
End Sub

Private Sub txt_Mobile_Number_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    Dim Num As String
        
        If Not cbox_Billing_Type = "Online" Then
        
            Num = txt_Mobile_Number.Text
            SQL = "Select TBL_Customer_Detail.[Customer Name],TBL_Customer_Detail.[Customer Mobile Number] From TBL_Customer_Detail WHERE TBL_Customer_Detail.[Customer Mobile Number] LIKE '%" & Num & "%'"
            I = 0
            
                
              cbox_Customer_List.Clear
              cbox_Customer_List.ColumnCount = 2
              cbox_Customer_List.ColumnWidths = "80,120"
              
                    Call Connection_Vindhya_Main_File(Con)
                    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                    
                    With Rec
                        If .EOF = False Or .BOF = False Then
                            .MoveFirst
                            
                            Do Until .EOF
                                cbox_Customer_List.AddItem
                                cbox_Customer_List.list(I, 0) = .Fields("Customer Mobile Number")
                                cbox_Customer_List.list(I, 1) = .Fields("Customer Name")
                                I = I + 1
                                .MoveNext
                            Loop
                            
                        End If
                                            
                        
                        
                        
                    End With
                    
                    Rec.Close
                    
            SQL = "SELECT * FROM TBL_Customer_Detail WHERE TBL_Customer_Detail.[Customer Mobile Number] Like '" & Num & "'"
            
        
            If Len(txt_Mobile_Number.Text) <> 10 And Len(txt_Mobile_Number.Text) <> 0 Then
                txt_Custmoner_Name.Text = "Not 10 digit number"
            Else
                
                txt_Custmoner_Name.Text = ""
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        txt_Custmoner_Name.Text = .Fields("Customer Name")
                        
                    End If
                    
                End With
                
                Rec.Close
            
            
            End If
            
            Con.Close
            
        End If
End Sub

Private Sub txt_QTY_Change()
    Dim QTY As String
                        
            QTY = txt_QTY.Text
            If Not txt_QTY.Text = "" Then
                Do Until IsNumeric(VBA.Right(QTY, 1)) = True
                    txt_QTY.Text = Left(QTY, Len(QTY) - 1)
                    QTY = txt_QTY.Text
                    
                Loop
            Else
                txt_QTY.Text = 1
            End If
            
            If cbox_Billing_Type.Text <> "" And cbox_Dish.Text <> "" And cbx_Dish_Variant <> "" Then
                Call cbx_Dish_Variant_Change
            End If
    
End Sub

Private Sub txt_QTY_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim QTY As String
        QTY = txt_QTY.Text
   
    If QTY = "" Then
        txt_QTY.Text = 1
    End If
              
               
    
    If KeyAscii = 43 Then
        
        
        Call lbl_Inc_QTY_Click
            
        
    ElseIf KeyAscii = 45 Then

        
        Call lbl_Dec_QTY_Click
            
    End If
    
        
        
End Sub

Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub



Private Sub UserForm_Initialize()

   
    
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            Dim I As Long   'For Looing
            
    If MultiPage1.Value = 0 Then
            
            SQL = "Select * From TBL_Billing_Type"
            
            Call Connection_Vindhya_Main_File(Con)
            
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                lbl_Disc_Type = shOther.Range("E1").Value
            
                'Adding list in Billing Type (Online,TakeAway, DineIn) in billing type combo box
                
                cbox_Billing_Type.Clear
                cbox_Billing_Type.ColumnCount = 1
                cbox_Billing_Type.ColumnWidths = "140"
                I = 0
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            
                            If .Fields("Billing Type") <> "" Then
                                cbox_Billing_Type.AddItem
                                cbox_Billing_Type.list(I, 0) = .Fields("Billing Type")
                                I = I + 1
                            End If
                            .MoveNext
                        Loop
                    End If
                
                End With
            
            Rec.Close
            Con.Close
            
            'Adding payment mode list to combox Payment 1 and paymnet 2
            
            SQL = "Select * From TBL_Payment_Mode"
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
                cbox_PaymentMode1.Clear
                cbox_PaymentMode2.Clear
                cbox_PaymentMode1.ColumnCount = 1
                cbox_PaymentMode2.ColumnCount = 1
                cbox_PaymentMode1.ColumnWidths = "140"
                cbox_PaymentMode2.ColumnWidths = "140"
                I = 0
                
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            If Not .Fields("CustPayment") = vbNullString Then
                                cbox_PaymentMode1.AddItem
                                cbox_PaymentMode1.list(I, 0) = .Fields("CustPayment")
                                cbox_PaymentMode2.AddItem
                                cbox_PaymentMode2.list(I, 0) = .Fields("CustPayment")
                                
                                I = I + 1
                                
                                
                            End If
                            
                            .MoveNext
                        Loop
                    
                    End If
                    
                
                End With
            
            Rec.Close
            Con.Close
            
            
            'Adding Dish code in the Dish code combo box
            
            Call Connection_Vindhya_Main_File(Con)
                
            SQL = "Select * From TBL_Vindhya_Menu Order By TBL_Vindhya_Menu.[Dish Code] ASC"
                
                
                cbox_Dish_Code.Clear
                cbox_Dish_Code.ColumnCount = 2
                cbox_Dish_Code.ColumnWidths = "50,120"
                
                I = 0
            
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            If .Fields("Status") = True Then
                                cbox_Dish_Code.AddItem
                                cbox_Dish_Code.list(I, 0) = .Fields("Dish Code")
                                cbox_Dish_Code.list(I, 1) = .Fields("Print Name")
                                I = I + 1
                                
                            End If
                            .MoveNext
                        Loop
                        
                    End If
                    
                End With
                
            Rec.Close
            
            SQL = "Select * From TBL_Vindhya_Dish_Variant"
            
                cbox_Dish.Clear
                cbox_Dish.ColumnCount = 1
                cbox_Dish.ColumnWidths = "140"
                I = 0
                
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            cbox_Dish.AddItem
                            cbox_Dish.list(I, 0) = .Fields("Dish Cat")
                            I = I + 1
                            .MoveNext
                        Loop
                    End If
                
                End With
            
            Rec.Close
            Con.Close
            
            'Hiding Button proceed, it will show only after selecting the billing type
                
                
            
                btn_Proceed.Visible = False
                lbl_Print_Code.Visible = False
                lbl_Ref1.Visible = False
                txt_Ref1.Visible = False
                lbl_Ref2.Visible = False
                txt_Ref2.Visible = False
                
                'Header for the Orderlist Listbox
                
        '        Order.Dish = cbox_Dish.Text
        '        Order.DishVariant = cbx_Dish_Variant.Text
        '        Order.PrintCode = lbl_Print_Code.Caption
        '        Order.Amount = lbl_dish_Amount.Caption
        '        Order.QTY = txt_QTY.Text
                
                Lbox_OrderList.Clear
                Lbox_OrderList.ColumnCount = 6
                Lbox_OrderList.ColumnWidths = "100,80,50,80,80,80"
                lbl_Table_Status.Visible = False
                lbl_Table_Status.Caption = ""
                
                Lbox_OrderList.AddItem
                Lbox_OrderList.list(0, 0) = "Dish"
                Lbox_OrderList.list(0, 1) = "Variant"
                Lbox_OrderList.list(0, 2) = "QTY"
                Lbox_OrderList.list(0, 3) = "Amount"
                Lbox_OrderList.list(0, 4) = "Print Code"
                Lbox_OrderList.list(0, 5) = "Disc Status"
                
                txt_QTY.Text = 1
                lbl_TAmnt = ""
                cbox_Dish.Text = ""
                cbx_Dish_Variant.Text = ""
                lbl_dish_Amount.Caption = ""
                txt_Mobile_Number.Text = ""
                cbox_Dish_Code.Text = ""
                lbl_FAmnt.Caption = ""
                txt_Disc.Text = ""
                lbl_Disc_Amnt = ""
                lbl_GST.Caption = ""
                txt_PaymentMode1.Text = ""
                txt_PaymentMode2.Text = ""
        
    ElseIf MultiPage1.Value = 1 Then
                    
            'Adding payment mode list to combox Payment 1 and paymnet 2
            
            lbl_Hold_Disc_Type = shOther.Range("E1").Value
            SQL = "Select * From TBL_Payment_Mode"
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                txt_Hold_PaymentMode1.Text = ""
                txt_Hold_PaymentMode2.Text = ""
                cbox_Hold_PaymentMode1.Clear
                cbox_Hold_PaymentMode2.Clear
                cbox_Hold_PaymentMode1.ColumnCount = 1
                cbox_Hold_PaymentMode2.ColumnCount = 1
                cbox_Hold_PaymentMode1.ColumnWidths = "140"
                cbox_Hold_PaymentMode2.ColumnWidths = "140"
                I = 0
                
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            If Not .Fields("CustPayment") = vbNullString Then
                                cbox_Hold_PaymentMode1.AddItem
                                cbox_Hold_PaymentMode1.list(I, 0) = .Fields("CustPayment")
                                cbox_Hold_PaymentMode2.AddItem
                                cbox_Hold_PaymentMode2.list(I, 0) = .Fields("CustPayment")
                                
                                I = I + 1
                                
                                
                            End If
                            
                            .MoveNext
                        Loop
                    
                    End If
                    
                
                End With
            
            Rec.Close
            Con.Close
            
            
        'Adding Dish code in the Dish code combo box
            
            Call Connection_Vindhya_Main_File(Con)

            SQL = "Select * From TBL_Vindhya_Menu Order by TBL_Vindhya_Menu.[Dish Code] ASC"


                cbox_Hold_Dish_Code.Clear
                cbox_Hold_Dish_Code.ColumnCount = 2
                cbox_Hold_Dish_Code.ColumnWidths = "50,120"

                I = 0

            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic

                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst

                        Do Until .EOF
                            cbox_Hold_Dish_Code.AddItem
                            cbox_Hold_Dish_Code.list(I, 0) = .Fields("Dish Code")
                            cbox_Hold_Dish_Code.list(I, 1) = .Fields("Print Name")
                            I = I + 1
                            .MoveNext
                        Loop

                    End If

                End With

            Rec.Close
            
            
        'Adding Dish code in the Dish code combo box
            
            Call Connection_Vindhya_Main_File(Con)
                
'            SQL = "Select * From TBL_Vindhya_Menu"
'
'
'                cbox_Hold_Dish_Code.Clear
'                cbox_Hold_Dish_Code.ColumnCount = 2
'                cbox_Hold_Dish_Code.ColumnWidths = "20,120"
'
'                I = 0
'
'            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
'
'                With Rec
'                    If .BOF = False Or .EOF = False Then
'                        .MoveFirst
'
'                        Do Until .EOF
'                            cbox_Hold_Dish_Code.AddItem
'                            cbox_Hold_Dish_Code.List(I, 0) = .Fields("Dish Code")
'                            cbox_Hold_Dish_Code.List(I, 1) = .Fields("Print Name")
'                            I = I + 1
'                            .MoveNext
'                        Loop
'
'                    End If
'
'                End With
'
'            Rec.Close
            
            SQL = "Select * From TBL_Vindhya_Dish_Variant"
            
                cbox_Hold_Dish.Clear
                cbox_Hold_Dish.ColumnCount = 1
                cbox_Hold_Dish.ColumnWidths = "140"
                I = 0
                
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
                With Rec
                    If .BOF = False Or .EOF = False Then
                        .MoveFirst
                        
                        Do Until .EOF
                            cbox_Hold_Dish.AddItem
                            cbox_Hold_Dish.list(I, 0) = .Fields("Dish Cat")
                            I = I + 1
                            .MoveNext
                        Loop
                    End If
                
                End With
            
            Rec.Close
            Con.Close
        
        'Hiding items which are of not use right now
        
   
                lbl_Hold_Print_Code.Visible = False
                lbl_Hold_Ref1.Visible = False
                txt_Hold_Ref1.Visible = False
                lbl_Hold_Ref2.Visible = False
                txt_Hold_Ref2.Visible = False
                
                'Header for the Orderlist Listbox
                
        '        Order.Dish = cbox_Dish.Text
        '        Order.DishVariant = cbx_Dish_Variant.Text
        '        Order.PrintCode = lbl_Print_Code.Caption
        '        Order.Amount = lbl_dish_Amount.Caption
        '        Order.QTY = txt_QTY.Text
                
                txt_Hold_Disc.Text = 0
                Lbox_Hold_OrderList.Clear
                Lbox_Hold_OrderList.ColumnCount = 5
                Lbox_Hold_OrderList.ColumnWidths = "100,80,50,80,80"
                
                
                Lbox_Hold_OrderList.AddItem
                Lbox_Hold_OrderList.list(0, 0) = "Dish"
                Lbox_Hold_OrderList.list(0, 1) = "Variant"
                Lbox_Hold_OrderList.list(0, 2) = "QTY"
                Lbox_Hold_OrderList.list(0, 3) = "Amount"
                Lbox_Hold_OrderList.list(0, 4) = "Print Code"
                
                lbox_Hold_Table_List.Clear
                lbox_Hold_Table_List.ColumnCount = 3
                lbox_Hold_Table_List.ColumnWidths = "40,60,80"
                
                
                lbox_Hold_Table_List.AddItem
                lbox_Hold_Table_List.list(0, 0) = "Table No"
                lbox_Hold_Table_List.list(0, 1) = "Status"
                lbox_Hold_Table_List.list(0, 2) = "Hold NO"
                
                I = 1
                
                txt_Hold_QTY.Text = 1
                
                txt_Hold_QTY.Text = 1
                lbl_Hold_TAmnt = ""
                cbox_Hold_Dish.Text = ""
                cbx_Hold_Dish_Variant.Text = ""
                lbl_dish_Amount.Caption = ""
                txt_Hold_Mobile_No.Text = ""
                lbl_Hold_FAmnt.Caption = ""
        'For Adding  the table list is
        

        SQL = "Select * From TBL_Table_Status"
        
        Call Connection_Vindhya_Billing_Detail(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                
                If .BOF = False Or .EOF = False Then
                    .MoveFirst
                    
                    Do Until .EOF
                        lbox_Hold_Table_List.AddItem
                        lbox_Hold_Table_List.list(I, 0) = .Fields("Table Number")
                        lbox_Hold_Table_List.list(I, 1) = .Fields("Status")
                        If IsNull(.Fields("KOT Hold Number")) = False Then
                            lbox_Hold_Table_List.list(I, 2) = .Fields("KOT Hold Number")
                        
                        End If
                        
                        I = I + 1
                        .MoveNext
                    Loop
                    
                                   
                End If
                
                
            End With
        
        
        
        Rec.Close
        Con.Close
        
        
    End If
        
        
        
    
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
   Cancel = 1
End Sub
