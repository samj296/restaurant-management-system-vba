VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Expense 
   Caption         =   "THE VINDHYA CAFE (Expense)"
   ClientHeight    =   13320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "Ufrm_Expense.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Expense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Back_Click()
    Me.Hide
    Ufrm_DashBoard.Show
    
End Sub

Private Sub btn_Category_Add_List_Click()
    If shMain.Range("F7").Value = "Admin" And txt_Category_Add_List <> "" Then
        Dim DataCon As New ADODB.Connection
        Dim DataRec As New ADODB.Recordset
        Dim SQL As String
        
        SQL = "Select * From TBL_Ledger WHERE TBL_Ledger.[Ledger Cat]='" & txt_Category_Add_List.Text & "'"
        Call Connection_Vindhya_Cash_Flow(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            With DataRec
                If .EOF = False Or .BOF = False Then
                    MsgBox "The category already in the list.", , "THE VINDHYA CAFE"
                    Exit Sub
                End If
                
            End With
        DataRec.Close
        
        SQL = "Select * From TBL_Ledger"
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        With DataRec
            .AddNew
            .Fields("Ledger Cat") = txt_Category_Add_List.Text
            .Update
            
        End With
        
        DataRec.Close
        DataCon.Close
        
        Call btn_Clear_Click
        Call UserForm_Initialize
    
    ElseIf shMain.Range("F7").Value <> "Admin" Then
        
        MsgBox "You are not authorise to add Category list.", , "THE VINDHYA CAFE"
        
        txt_Category_Add_List.Text = ""
        
    ElseIf txt_Category_Add_List = "" Then
            
            MsgBox "Please enter the text to add the Category", , "THE VINDHYA CAFE"
               
        
    End If
    
End Sub

Private Sub btn_Clear_Click()
    txt_Date.Text = VBA.Format(Now, "dd/mm/yyyy")
    txt_Expense.Text = ""
    Cbox_Category.Value = ""
    Cbox_Payment_Mode.Value = ""
    txt_Category_Add_List.Text = ""
    txt_Amount.Text = ""
    txt_Adding_Amount.Text = ""
    txt_Narration.Text = ""
End Sub



Private Sub btn_Export_Data_Click()
    Dim WB As Workbook
    Dim Sh As Worksheet
    Dim I As Double
    Dim FPath As String
    Dim FileName As String
    Dim Lrow As Double

        Set WB = Workbooks.Add
        Set Sh = WB.Worksheets(1)
        
        If lbox_Expense_List.ListCount > 1 Then
            For I = 0 To lbox_Expense_List.ListCount - 1
                Sh.Cells(I + 1, 1).Value = lbox_Expense_List.list(I, 0)
                Sh.Cells(I + 1, 2).Value = lbox_Expense_List.list(I, 1)
                Sh.Cells(I + 1, 3).Value = lbox_Expense_List.list(I, 2)
                Sh.Cells(I + 1, 4).Value = lbox_Expense_List.list(I, 3)
                Sh.Cells(I + 1, 5).Value = lbox_Expense_List.list(I, 4)
                Sh.Cells(I + 1, 6).Value = lbox_Expense_List.list(I, 5)
                
            Next I
                
                Sh.Range("A:F").HorizontalAlignment = xlLeft
                Lrow = Sh.Range("B" & Rows.Count).End(xlUp).Row
                
                Sh.Range("A1", "F" & Lrow).Borders.LineStyle = xlContinuous
                Sh.Range("A1", "F1").Font.Size = 12
                Sh.Range("A1", "F1").Font.Bold = True
                Sh.Range("A1", "F1").Interior.Color = RGB(0, 0, 255)
                Sh.Range("A1", "F1").Font.Color = RGB(255, 255, 255)
                Sh.Range("A1", "F1").HorizontalAlignment = xlCenter
                Sh.Columns("A:F").AutoFit
                
            With Application.FileDialog(msoFileDialogFolderPicker)
                .ButtonName = "Pick Folder"
                .Title = "THE VINDHYA CAFE"
                .Show
                FPath = .SelectedItems(1)
                
            End With
            Application.DisplayAlerts = False
            Application.DisplayAlerts = True
            WB.SaveAs FPath & "\Expense_" & VBA.Format(Now, "DD_MM_YY") & ".xlsx"
            WB.Close
            Set WB = Nothing
        End If
    
    
    
End Sub

Private Sub btn_Update_Click()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select * From TBl_Expense"
    Call Connection_Vindhya_Cash_Flow(DataCon)
    
    DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
    
    With DataRec
        .AddNew
        .Fields("ExpDate") = txt_Date.Text
        .Fields("Particular") = txt_Expense.Text
        .Fields("Ledger Cat") = Cbox_Category.Value
        .Fields("Payment Mode") = Cbox_Payment_Mode.Value
        .Fields("Amount") = txt_Amount.Text
        .Fields("Narration") = txt_Narration.Text
        .Update
        
    End With
    
    DataRec.Close
    DataCon.Close
    
    Call btn_Clear_Click
    
    
End Sub

Private Sub txt_Adding_Amount_Change()
    Dim Amnt As New cls_calculation
    Dim Num As String
    Dim Res As String
    
    Num = txt_Adding_Amount.Text
    
    Amnt.Calculate Num, Res
    
    txt_Amount.Text = Res

End Sub






Private Sub txt_Date_From_AfterUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    
    lbox_Expense_List.Clear
    lbox_Expense_List.ColumnCount = 6
    lbox_Expense_List.ColumnWidths = "80,80,80,80,80,100"
    lbox_Expense_List.AddItem
    lbox_Expense_List.list(0, 0) = "Date"
    lbox_Expense_List.list(0, 1) = "Particular"
    lbox_Expense_List.list(0, 2) = "Category"
    lbox_Expense_List.list(0, 3) = "Mode"
    lbox_Expense_List.list(0, 4) = "Amount"
    lbox_Expense_List.list(0, 5) = "Narration"
        
                
        If txt_Date_to = "" Then
            If IsDate(VBA.Format(txt_Date_From.Text, "DD-MM-YYYY")) = True Then
                SQL = "Select * From TBL_Expense WHERE TBL_Expense.[ExpDate] = #" & VBA.Format(txt_Date_From.Text, "MM/DD/YYYY") & "# ORDER BY TBL_Expense.[ExpDate] ASC"
                Call Connection_Vindhya_Cash_Flow(Con)
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                   I = 1
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                        
                        lbox_Expense_List.AddItem
                        lbox_Expense_List.list(I, 0) = VBA.Format(.Fields("ExpDate"), "DD-MM-YYYY")
                        lbox_Expense_List.list(I, 1) = .Fields("Particular")
                        lbox_Expense_List.list(I, 2) = .Fields("Ledger Cat")
                        lbox_Expense_List.list(I, 3) = .Fields("Payment Mode")
                        lbox_Expense_List.list(I, 4) = .Fields("Amount")
                        If .Fields("Narration") <> "" Then
                            lbox_Expense_List.list(I, 5) = .Fields("Narration")
                        End If
                            I = I + 1
                        .MoveNext
                        Loop
                        
                    
                    End If
                
                End With
                Rec.Close
                Con.Close
            End If
        ElseIf txt_Date_to.Text <> "" And txt_Date_From.Text <> "" And _
        IsDate(VBA.Format(txt_Date_From.Text, "DD-MM-YYYY")) = True And _
        IsDate(VBA.Format(txt_Date_to.Text, "DD-MM-YYYY")) = True Then
            Call txt_Date_to_AfterUpdate
        End If
    
End Sub

Private Sub txt_Date_to_AfterUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
        If txt_Date_to.Text <> "" And txt_Date_From.Text <> "" And _
        IsDate(VBA.Format(txt_Date_From.Text, "DD-MM-YYYY")) = True And _
        IsDate(VBA.Format(txt_Date_to.Text, "DD-MM-YYYY")) = True Then
            SQL = "Select * From TBL_Expense Where TBL_Expense.[ExpDate] BETWEEN #" & VBA.Format(txt_Date_From.Text, "MM-DD-YYYY") & "# AND #" & VBA.Format(txt_Date_to.Text, "MM-DD-YYYY") & "# ORDER BY TBL_Expense.[ExpDate] ASC"
    
    lbox_Expense_List.Clear
    lbox_Expense_List.ColumnCount = 6
    lbox_Expense_List.ColumnWidths = "80,80,80,80,80,100"
    lbox_Expense_List.AddItem
    lbox_Expense_List.list(0, 0) = "Date"
    lbox_Expense_List.list(0, 1) = "Particular"
    lbox_Expense_List.list(0, 2) = "Category"
    lbox_Expense_List.list(0, 3) = "Mode"
    lbox_Expense_List.list(0, 4) = "Amount"
    lbox_Expense_List.list(0, 5) = "Narration"
    I = 1
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    Do Until .EOF
                       lbox_Expense_List.AddItem
                       lbox_Expense_List.list(I, 0) = VBA.Format(.Fields("ExpDate"), "DD-MM-YYYY")
                       lbox_Expense_List.list(I, 1) = .Fields("Particular")
                       lbox_Expense_List.list(I, 2) = .Fields("Ledger Cat")
                       lbox_Expense_List.list(I, 3) = .Fields("Payment Mode")
                       lbox_Expense_List.list(I, 4) = .Fields("Amount")
                            If .Fields("Narration") <> "" Then
                             lbox_Expense_List.list(I, 5) = .Fields("Narration")
                            End If
                        
                        I = I + 1
                    .MoveNext
                    Loop
                End If
                
                
            End With
            
            
            Rec.Close
            Con.Close
        ElseIf txt_Date_to.Text = "" And IsDate(txt_Date_From.Text) = True Then
                Call txt_Date_From_AfterUpdate
        End If

End Sub

Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Deactivate()
    Me.Hide
    Ufrm_DashBoard.Show
End Sub

Private Sub UserForm_Initialize()
    Dim Con As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long   'For looping inside the database
    
    txt_Date.Text = VBA.Format(VBA.Date, "DD/MM/YYYY")
    
    SQL = "Select * From TBL_Ledger"
    I = 0
     Call Connection_Vindhya_Cash_Flow(Con)
     
        'For adding list to the category combobox
    
        DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        Cbox_Category.Clear
        Cbox_Category.ColumnCount = 1
        Cbox_Category.ColumnWidths = "120"
        
        With DataRec
            If .BOF = True And .EOF = True Then
                
                GoTo Mode
                
            End If
            .MoveFirst
            
            Do Until .EOF
                If IsNull(.Fields("Ledger Cat")) = False Then
                    Cbox_Category.AddItem
                    Cbox_Category.list(I, 0) = .Fields("Ledger Cat")
                    I = I + 1
                End If
                .MoveNext
            Loop
        
        End With
        
Mode:
            'For adding list to the Mode of payment combobox
            
        DataRec.Close
        
        SQL = "Select * From TBL_Payment_Mode"
        
        DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        Cbox_Payment_Mode.Clear
        Cbox_Payment_Mode.ColumnCount = 1
        Cbox_Payment_Mode.ColumnWidths = "120"
        I = 0
        
        With DataRec
            If .EOF = True And .BOF = True Then
                GoTo Final
            End If
            .MoveFirst
            
            Do Until .EOF
                If IsNull(.Fields("Received")) = False Then
                  Cbox_Payment_Mode.AddItem
                  Cbox_Payment_Mode.list(I, 0) = .Fields("Received")
                  I = I + 1
                End If
            
                .MoveNext
            Loop
            
            
        
        End With
        
    
Final:
    
    DataRec.Close
    Con.Close
   If shMain.Range("F7").Value = "Admin" Then
        lbox_Expense_List.Clear
        lbox_Expense_List.ColumnCount = 6
        lbox_Expense_List.ColumnWidths = "80,80,80,80,80,100"
        lbox_Expense_List.AddItem
        lbox_Expense_List.list(0, 0) = "Date"
        lbox_Expense_List.list(0, 1) = "Particular"
        lbox_Expense_List.list(0, 2) = "Category"
        lbox_Expense_List.list(0, 3) = "Mode"
        lbox_Expense_List.list(0, 4) = "Amount"
        lbox_Expense_List.list(0, 5) = "Narration"
        Frm_ExpenseList.Visible = True
    Else
        Frm_ExpenseList.Visible = False
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
   Cancel = 1
End Sub
