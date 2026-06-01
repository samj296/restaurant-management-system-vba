VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Billing_Completion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Bill_Hold(Orderlist As Collection, BillingMode As String, _
HoldNo As String, BillNo As String, Disc As Long, _
PaymentMode As String, TableNumber As String, CustName As String, MobileNo As String, TAmnt As Double)



    Dim I As Long 'For looping nside the list
    Dim SQL As String
    Dim OrderOut As cls_Dish_Qty_Amount
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim Rec As cls_connection
    Dim Amnt As Double
    Dim Lrow As Long
    Dim printer As String
    Dim Bno As Long
    
    
    
        
         SQL = "Select * From TBL_Bill_Total"
        Call Connection_Vindhya_Billing_Detail(DataCon)
        
        
        DataRec.Open SQL, DataCon, adOpenForwardOnly, adLockOptimistic
        
            With DataRec
                If .RecordCount >= 1 Then
                .MoveLast
                Bno = .Fields("Bill Count for Bill Number") + 1
            Else
                Bno = 1001
            End If
            
                BillNo = "Vindhya/" & Bno
                
                
            End With
                
                'Feeding Data in KOT Hold Detail
                
                SQL = "Select * From TBL_KOT_HOLD_Total"
                Set Rec = New cls_connection
                
                Rec.Vindhya_Billing_Detail DataCon, DataRec, SQL
                   
                HoldNo = "Vindhya/" & DataRec.RecordCount + 1001
                    
                                        
                    For I = 1 To Orderlist.Count
                        Set OrderOut = New cls_Dish_Qty_Amount
                        
                           Set OrderOut = Orderlist.Item(I)
                            
                        
                        
                    
                    Next I
                                    
                            With DataRec
                                .AddNew
                                .Fields("KOT Hold Number") = HoldNo
                                .Fields("Table Number") = TableNumber
                                .Fields("Disc") = Disc
                                .Fields("Total Amount") = Amnt
                                .Fields("Status") = "Pending"
                                .Fields("UserId") = shMain.Range("F6").Value
                                .Fields("Cust Name") = CustName
                                .Fields("Mobile No") = MobileNo
                                .Update
                            
                            End With
                
                
                
                DataRec.Close
                DataCon.Close
                
                SQL = "Select * From TBL_KOT_HOLD_DETAIL"
                
                
                Rec.Vindhya_Billing_Detail DataCon, DataRec, SQL
                    
                    'Default Row height is 15, & column width is 8.43
                    
               
                
                
                
                Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                
                
                For I = 1 To Orderlist.Count
                   Set OrderOut = New cls_Dish_Qty_Amount
                   Set OrderOut = Orderlist.Item(I)
                    With DataRec
                        .AddNew
                        .Fields("Order Date") = VBA.Format(Now, "dd-mm-yyyy")
                        .Fields("KOT Hold Number") = HoldNo
                        .Fields("Table Number") = TableNumber
                        .Fields("Item") = OrderOut.dish
                        .Fields("Variant") = OrderOut.DishVariant
                        .Fields("QTY") = OrderOut.QTY
                        .Fields("Amount") = OrderOut.Amount
                        .Fields("Print Code") = OrderOut.PrintCode
                        .Fields("Status") = "Pending"
                        .Update
                        
                        
                    End With
                        
                Next I
                        
                        DataRec.Close
                        DataCon.Close
                
                
                
    
    
End Function

Function Hold_Bill_Print(PrintType As String, Orderlist As Collection, HoldNo As String, BillNo As String, Disc As Long, GST As Long, TableNo As String, _
CustName As String, MobileNo As String, TAmnt As Double, CalcAmnt)
    Dim OrderOut As cls_Dish_Qty_Amount
    Dim I As Long
    Dim Lrow As Long
    Dim printer As String
    
        If PrintType = "Hold" Or PrintType = "Duplicate Hold" Then
            
                    
                    'Default Row height is 15, & column width is 8.43
                    
                shPrint.Range("A1:B1").Merge
                shPrint.Range("A:B").Font.Size = 8
                shPrint.Range("A1").ColumnWidth = 18
                shPrint.Range("B1").ColumnWidth = 3
                shPrint.Range("A1").RowHeight = 18
                shPrint.Range("A1:B1").Font.Size = 14
                shPrint.Range("A1:B1").Value = "THE VINDHYA CAFE"
                shPrint.Range("A2:B2").Merge
                shPrint.Range("A2:B2").Value = "KOT NO. - " & HoldNo
                shPrint.Range("A3:B3").Merge
                shPrint.Range("A3:B3").Value = CustName
                shPrint.Range("A4:B4").Merge
                shPrint.Range("A4:B4").Value = VBA.Format(VBA.Now, "DD-MM-YY - H:MM AM/PM")
                shPrint.Range("A5:B5").Merge
                shPrint.Range("A5:B5").Value = "ORDER ID - " & VBA.Right(HoldNo, 3)
                If PrintType = "Duplicate Hold" Then
                    shPrint.Range("A6:B6").Merge
                    shPrint.Range("A6:B6").Value = "Duplicate"
                End If
                    Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                    
                
                 For I = 1 To Orderlist.Count
                   Set OrderOut = New cls_Dish_Qty_Amount
                   Set OrderOut = Orderlist.Item(I)
                    
                        shPrint.Range("A" & Lrow + I).Value = OrderOut.PrintCode
                        shPrint.Range("B" & Lrow + I).Value = OrderOut.QTY
                        With shPrint.Range("A" & Lrow + I, "B" & Lrow + I)
                            .HorizontalAlignment = xlCenter
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Borders(xlEdgeTop).ThemeColor = 2
                            .Borders(xlEdgeTop).TintAndShade = 0
                            .Borders(xlEdgeTop).Weight = xlThin
                            
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Borders(xlEdgeBottom).ThemeColor = 2
                            .Borders(xlEdgeBottom).TintAndShade = 0
                            .Borders(xlEdgeBottom).Weight = xlThin
                            
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeLeft).ThemeColor = 2
                            .Borders(xlEdgeLeft).TintAndShade = 0
                            .Borders(xlEdgeLeft).Weight = xlThin
                            
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).ThemeColor = 2
                            .Borders(xlEdgeRight).TintAndShade = 0
                            .Borders(xlEdgeRight).Weight = xlThin
                            
                        End With
                        
                Next I
                
                        'Printin the orderlist for Kitchen
                        
                 Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row
                printer = shMain.Range("F3").Value
                shPrint.Range("A1", "B" & Lrow).PrintOut ActivePrinter:=printer
                    
                        'Deleting Data from the Print Sheet
                    shPrint.Range("A1", "B1").UnMerge
                    shPrint.Range("A:A").RowHeight = 15
                    shPrint.Range("1:1").ColumnWidth = 8.43      'Default Row height is 15, & column width is 8.43
                    shPrint.Range("A1", "B" & Lrow).Delete xlUp
                    
                
        ElseIf PrintType = "Bill" Or PrintType = "Duplicate Bill" Then
                shPrint.Range("A:C").Font.Size = 8
                shPrint.Range("A1:C1").Merge
                shPrint.Range("A1").ColumnWidth = 12
                shPrint.Range("B1").ColumnWidth = 3
                shPrint.Range("C1").ColumnWidth = 6
                shPrint.Range("A1").RowHeight = 18
                shPrint.Range("A1:C1").Font.Size = 14
                shPrint.Range("A1:C1").Value = "THE VINDHYA CAFE"
                shPrint.Range("A2").Value = "Date"
                shPrint.Range("B2:C2").Merge
                shPrint.Range("B2:C2").Value = VBA.Format(VBA.Date, "DD-MM-YY")
                shPrint.Range("A3:C3").Merge
                shPrint.Range("A3:C3").Value = "Bill No:- " & BillNo
                shPrint.Range("A3:C3").HorizontalAlignment = xlLeft
                shPrint.Range("A4:C4").Merge
                shPrint.Range("A4:C4").Value = CustName & " - " & MobileNo
                shPrint.Range("A5:A5").Merge
                shPrint.Range("A5:A5").Value = "Order No-" & VBA.Right(HoldNo, 3)
                If PrintType = "Duplicate Bill" Then
                    shPrint.Range("A6:B6").Merge
                    shPrint.Range("A6:B6").Value = "Duplicate"
                End If
                
                            
                    Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
            
                    For I = 1 To Orderlist.Count
                        
                        Set OrderOut = New cls_Dish_Qty_Amount
                        Set OrderOut = Orderlist.Item(I)
                        
                        
                        shPrint.Range("A" & Lrow + I).Value = OrderOut.PrintCode
                        shPrint.Range("B" & Lrow + I).Value = OrderOut.QTY
                        shPrint.Range("C" & Lrow + I).Value = OrderOut.Amount
                            
                        
                    
                    Next I
                
                Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row
                         shPrint.Range("A" & Lrow + 1).Value = "Amnt bef calc"
                         shPrint.Range("C" & Lrow + 1).Value = TAmnt
                         
                Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                
                    If Not Disc = 0 Then
                        shPrint.Range("A" & Lrow + 1).Value = "Disc"
                        shPrint.Range("C" & Lrow + 1).Value = Disc
                        Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row
                    End If
                    
                    If Not GST = 0 Then
                        shPrint.Range("A" & Lrow + 1).Value = "CGST"
                        shPrint.Range("C" & Lrow + 1).Value = GST / 2
                        shPrint.Range("A" & Lrow + 2).Value = "SGST"
                        shPrint.Range("C" & Lrow + 2).Value = GST / 2
                        Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row
                    Else
                        shPrint.Range("A" & Lrow + 1).Value = "CGST"
                        shPrint.Range("C" & Lrow + 1).Value = 0
                        shPrint.Range("A" & Lrow + 2).Value = "SGST"
                        shPrint.Range("C" & Lrow + 2).Value = 0
                    End If
                        
                        shPrint.Range("A" & Lrow + 3).Value = "FinalAmnt"
                        shPrint.Range("c" & Lrow + 3).Value = CalcAmnt
                      
                      Lrow = shPrint.Range("A" & Rows.Count).End(xlUp).Row
                      
                printer = shMain.Range("F4").Value
                shPrint.Range("A1", "C" & Lrow).PrintOut ActivePrinter:=printer
                
               
                
                     'Default Row height is 15, & column width is 8.43
                     
                shPrint.Range("A1", "C1").UnMerge
                shPrint.Range("A:A").RowHeight = 15
                shPrint.Range("1:1").ColumnWidth = 8.43
                shPrint.Range("A1", "C" & Lrow).Delete
                          
        
        End If
    
End Function

Function Disc_GST_Calculation(Orderlist As Collection, TAmnt As Double, _
CalcAmnt As Double, Disc As Long, DiscType As String, _
DiscPercent As String, GSTStat As Boolean, GSTPercent As Byte, GST As Long, Btype As String)
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim OrderOut As cls_Dish_Qty_Amount
    

    
        SQL = "Select TBL_Billing_Type.Query, TBL_Billing_Type.UserInput From TBL_Billing_Type"
        
        Call Connection_Vindhya_Main_File(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        
            
                With Rec
                    .MoveFirst
                    Do Until .EOF
                            If .Fields("UserInput") = DiscType Then
                                DiscType = .Fields("Query")
                            End If
                            
                            If .Fields("Query") = "GST Status" And .Fields("UserInput") = "Y" And Btype <> "Online" Then
                                GSTStat = True
                               
                            End If
                            If .Fields("Query") = "GST %" Then
                                GSTPercent = CByte(.Fields("UserInput"))
                            End If
                            
                            .MoveNext
                    
                    Loop
                    
                    
                End With
            
            
            Rec.Close
            Con.Close
        
        
                    
               'Calculation of Disc
               
                If Not Disc = 0 Then
                    If DiscType = "DiscType 1" Then 'Disc Type in rupees
                        CalcAmnt = TAmnt - Disc
                        DiscPercent = (Disc / TAmnt) * 100
                    ElseIf DiscType = "DiscType 2" Then 'Disc type in Percent
                        CalcAmnt = TAmnt - (TAmnt * (Disc / 100))
                        DiscPercent = Disc
                        Disc = TAmnt * (DiscPercent / 100)
                    End If
                    
                    
                Else
                        CalcAmnt = TAmnt
                        DiscPercent = 0
                        Disc = 0
                
                End If
                
                If GSTStat = True Then
                        GST = CalcAmnt * GSTPercent / 100
                Else
                        GST = 0
                    
                End If
                
                

End Function

Function Finalise_Billing(Orderlist As Collection, Btype As String, Disc As Long, DiscType As String, _
DiscPercent As String, PaymentMode1 As String, PaymentMode2 As String, Mode1Amnt As Double, Mode2Amnt As Double, _
BillNo As String, HoldNo As String, Mobile As String, Ref As String, _
Amnt As Double, CalcAmnt As Double, GST As Long, CustName As String)
    
    
    
    
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim OrderOut As cls_Dish_Qty_Amount
'    Dim Amnt As Double
'    Dim CalcAmnt As Double 'After disc and GST amount
    Dim GSTStat As Boolean
    Dim GSTPercent As Byte

    Dim Lrow As Long
    
       
        
        
        'Recording Data in the Bill Total Table
        
        
        SQL = "Select * From TBL_Bill_Total"
        Call Connection_Vindhya_Billing_Detail(Con)
        
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            
        
            
            With Rec
                Dim Bno As Double
            
            If .RecordCount >= 1 Then
                .MoveLast
                Bno = .Fields("Bill Count for Bill Number") + 1
            Else
                Bno = 1001
            End If
            
                BillNo = "Vindhya/" & Bno
                
                For I = 1 To Orderlist.Count
                    Set OrderOut = New cls_Dish_Qty_Amount
                    
                    Set OrderOut = Orderlist.Item(I)
                    
                    
                Next I
                    
         
                                   
                    
                    .AddNew
                    .Fields("Billing Date") = VBA.Format(Now, "DD/MM/YYYY")
                    .Fields("Bill Number") = BillNo
                    .Fields("Bill Count for Bill Number") = Bno
                    .Fields("KOT Hold Number") = HoldNo
                    If Not Disc = 0 Then
                        .Fields("Disc Amount") = Amnt - CalcAmnt
                    End If
                    If Not DiscPercent = "" Or Not DiscPercent = 0 Then
                        .Fields("Disc %") = DiscPercent
                    End If
                    .Fields("GST") = GST
                    .Fields("Total Amount") = CalcAmnt
                    .Fields("UserId") = shMain.Range("F6").Value
                    .Fields("Mobile No") = Mobile
                    .Fields("Reference") = Ref
                    
                    
                If Not Btype = "Online" Then
                    Select Case PaymentMode1
                        
                        Case Is = ""
                        MsgBox "Payment Mode1 can't be empty", , "THE VINDHYA CAFE"
                        
                        Case Is = "Cash"
                        If Mode1Amnt = CalcAmnt Or Mode1Amnt < CalcAmnt Then
                            .Fields("Cash") = Mode1Amnt
                        Else
                            .Fields("Cash") = CalcAmnt
                        End If
                        Case Is = "PayTm"
                        
                        If Mode1Amnt = CalcAmnt Or Mode1Amnt < CalcAmnt Then
                            .Fields("PayTm") = Mode1Amnt
                        Else
                            .Fields("PayTm") = CalcAmnt
                        End If
                        
                        Case Is = "Card"
                        If Mode1Amnt = CalcAmnt Or Mode1Amnt < CalcAmnt Then
                            .Fields("Card") = Mode1Amnt
                        Else
                            .Fields("Card") = CalcAmnt
                        End If
                        
                        Case Is = "UPI"
                        If Mode1Amnt = CalcAmnt Or Mode1Amnt < CalcAmnt Then
                            .Fields("UPI") = Mode1Amnt
                        Else
                            .Fields("UPI") = CalcAmnt
                        End If
                        
                    End Select
                    
                    Select Case PaymentMode2
                        
                        Case Is = ""
                            .Update
                            GoTo NextStep
                        
                        Case Is = "Cash"
                        If CalcAmnt - Mode1Amnt > 0 Then
                            .Fields("Cash") = CalcAmnt - Mode1Amnt
                        Else
                            .Fields("Cash") = 0
                        End If
                        Case Is = "PayTm"
                        If CalcAmnt - Mode1Amnt > 0 Then
                            .Fields("PayTm") = CalcAmnt - Mode1Amnt
                        Else
                            .Fields("PayTm") = 0
                        End If
                        
                        Case Is = "Card"
                        If CalcAmnt - Mode1Amnt > 0 Then
                            .Fields("Card") = CalcAmnt - Mode1Amnt
                        Else
                            .Fields("Card") = 0
                        End If
                        
                        Case Is = "UPI"
                        If CalcAmnt - Mode1Amnt > 0 Then
                            .Fields("UPI") = CalcAmnt - Mode1Amnt
                        Else
                            .Fields("UPI") = 0
                        End If
                    
                    End Select
                Else
                    .Fields("Online") = CalcAmnt
                    .Fields("Reference") = CustName
                End If
                        .Update
                        
            End With
    
NextStep:
                Rec.Close
            'Recording Data in Bill Detail Table
                
                
                
            
            
              
            SQL = "Select * From TBL_Bill_Detailed"
            
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
                With Rec
                    For I = 1 To Orderlist.Count
                        
                        Set OrderOut = New cls_Dish_Qty_Amount
                        Set OrderOut = Orderlist.Item(I)
                        
                        .AddNew
                        .Fields("Billing Date") = VBA.Format(Now, "dd-mm-yyyy")
                        .Fields("Food Item") = OrderOut.dish
                        .Fields("Amount") = OrderOut.Amount
                        .Fields("Bill Number") = BillNo
                        .Update
                        
                        
                        
                        
                    
                    Next I
                
                End With
                
                Rec.Close
'to update the same on hold list in database
                SQL = "Select * From TBL_KOT_HOLD_DETAIL WHERE TBL_KOT_HOLD_DETAIL.[KOT Hold Number]='" & HoldNo & "'"
                
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                    With Rec
                        If .BOF = False Or .EOF = False Then
                            .MoveFirst
                            Do Until .EOF
                                .Fields("Status") = "Completed"
                                .Update
                                .MoveNext
                            Loop
                        Else
                            MsgBox "Unable to update the status in KOT Hold list, please update the list manually from the database", , _
                            "THE VINDHYA CAFE"
                        End If
                    
                    End With
                    
                    Rec.Close
                SQL = "Select * From TBL_KOT_HOLD_Total WHERE TBL_KOT_Hold_Total.[KOT Hold Number]='" & HoldNo & "'"
                
'to update the hold list total in database

                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                    
                    With Rec
                        If .BOF = False Or .EOF = False Then
                            
                            
                                .Fields("Status") = "Completed"
                            
                            
                        End If
                    
                    End With

              'Rec.Close
              Con.Close
                             
              
    
End Function
