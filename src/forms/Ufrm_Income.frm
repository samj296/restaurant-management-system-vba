VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Income 
   Caption         =   "THE VINDHYA CAFE (Income)"
   ClientHeight    =   13210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18855
   OleObjectBlob   =   "Ufrm_Income.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Add_Income_Click()
    If shMain.Range("F7").Value = "Admin" Then
            Dim Con As New ADODB.Connection
            Dim DataRec As New ADODB.Recordset
            Dim SQL As String
            
                SQL = "Select * From TBL_Income_Mode WHERE TBL_Income_Mode.[Source]='" & txt_Income_Add_List & "'"
                
                Call Connection_Vindhya_Cash_Flow(Con)
                
                DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With DataRec
                    If .EOF = False Or .BOF = False Then
                        MsgBox "The Income mode already in the list", , "THE VINDHYA CAFE"
                        Exit Sub
                    End If
                End With
                
                DataRec.Close
                SQL = "Select * From TBL_Income_Mode"
                
                DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With DataRec
                    .AddNew
                    .Fields("Source") = txt_Income_Add_List.Text
                    .Update
                    
                End With
                
                txt_Income_Add_List.Text = ""
                
                
                DataRec.Close
                Con.Close
                
                    Call UserForm_Activate
    Else
        
        MsgBox "You are not authorise to add income source", , "THE VINDHYA CAFE"
        
        Call btn_Clear_Click
        
    End If
        
        

    
End Sub



Private Sub btn_Back_Click()
    
    Me.Hide
    
    Ufrm_DashBoard.Show
    
End Sub

Private Sub btn_Clear_Click()
    txt_Date.Text = VBA.Format(VBA.Date, "DD/MM/YY")
    txt_Received.Text = ""
    Cbox_Payment_Mode.Value = ""
    Cbox_IncomeSource.Value = ""
    txt_Amount.Text = ""
    txt_Adding_Amount.Text = ""
    txt_Narration.Text = ""
    txt_Income_Add_List.Text = ""
    
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
        
        If lbox_Income_List.ListCount > 1 Then
            For I = 0 To lbox_Income_List.ListCount - 1
                Sh.Cells(I + 1, 1).Value = lbox_Income_List.list(I, 0)
                Sh.Cells(I + 1, 2).Value = lbox_Income_List.list(I, 1)
                Sh.Cells(I + 1, 3).Value = lbox_Income_List.list(I, 2)
                Sh.Cells(I + 1, 4).Value = lbox_Income_List.list(I, 3)
                Sh.Cells(I + 1, 5).Value = lbox_Income_List.list(I, 4)
                Sh.Cells(I + 1, 6).Value = lbox_Income_List.list(I, 5)
                
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
            WB.SaveAs FPath & "\Income_" & VBA.Format(Now, "DD_MM_YY") & ".xlsx"
            WB.Close
            Set WB = Nothing
        End If
    
End Sub

Private Sub btn_Update_Click()
    Dim Con As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    
    If Not txt_Date.Text = "" And Not txt_Received.Text = "" And Not Cbox_Payment_Mode.Value = "" _
        And Not Cbox_IncomeSource.Value = "" And Not txt_Amount.Text = "" Then
             SQL = "Select * From TBL_Income"
             
             Call Connection_Vindhya_Cash_Flow(Con)
             DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
             
             With DataRec
                 .AddNew
                 .Fields("IncDate") = txt_Date.Text
                 .Fields("Particular") = txt_Received.Text
                 .Fields("Bank/Cash") = Cbox_Payment_Mode.Value
                 .Fields("Rec From") = Cbox_IncomeSource.Value
                 .Fields("Amount") = txt_Amount.Text
                 .Fields("Narration") = txt_Narration.Text
                 .Update
            
             End With
             
             
             DataRec.Close
             Con.Close
             Call btn_Clear_Click
    Else
        
        MsgBox "Please fill all the details, to upadte", , "THE VINDHYA CAFE"
        
    End If
        
End Sub




Private Sub txt_Adding_Amount_Change()
    
        Dim Amnt As New cls_calculation
        Dim Result As String
        Dim Num As String
    
            Num = txt_Adding_Amount.Text
            Amnt.Calculate Num, Result
        
        txt_Amount.Text = Result
    
Exit Sub


End Sub








Private Sub txt_Date_From_AfterUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
        
                lbox_Income_List.Clear
                lbox_Income_List.ColumnCount = 6
                lbox_Income_List.ColumnWidths = "80,80,80,80,80,100"
                lbox_Income_List.AddItem
                lbox_Income_List.list(0, 0) = "Date"
                lbox_Income_List.list(0, 1) = "Particular"
                lbox_Income_List.list(0, 2) = "Source"
                lbox_Income_List.list(0, 3) = "Mode"
                lbox_Income_List.list(0, 4) = "Amount"
                lbox_Income_List.list(0, 5) = "Narration"
                
        
        If txt_Date_to = "" Then
            If IsDate(VBA.Format(txt_Date_From.Text, "DD-MM-YYYY")) = True Then
                SQL = "Select * From TBL_Income WHERE TBL_Income.[IncDate] = #" & VBA.Format(txt_Date_From.Text, "MM/DD/YYYY") & "# ORDER BY TBL_Income.[IncDate] ASC"
                Call Connection_Vindhya_Cash_Flow(Con)
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                   I = 1
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                        
                        lbox_Income_List.AddItem
                        lbox_Income_List.list(I, 0) = VBA.Format(.Fields("IncDate"), "DD-MM-YYYY")
                        lbox_Income_List.list(I, 1) = .Fields("Particular")
                        lbox_Income_List.list(I, 2) = .Fields("Rec From")
                        lbox_Income_List.list(I, 3) = .Fields("Bank/Cash")
                        lbox_Income_List.list(I, 4) = .Fields("Amount")
                        If .Fields("Narration") <> "" Then
                            lbox_Income_List.list(I, 5) = .Fields("Narration")
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

'SELECT *
'FROM TBL_Expense
'WHERE TBL_Expense.[ExpDate] between #8/1/2021# and #3/4/2022#;


End Sub





Private Sub txt_Date_to_AfterUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
        If txt_Date_to.Text <> "" And txt_Date_From.Text <> "" And _
        IsDate(VBA.Format(txt_Date_From.Text, "DD-MM-YYYY")) = True And _
        IsDate(VBA.Format(txt_Date_to.Text, "DD-MM-YYYY")) = True Then
            SQL = "Select * From TBL_Income Where TBL_Income.IncDate BETWEEN #" & VBA.Format(txt_Date_From.Text, "MM-DD-YYYY") & "# AND #" & VBA.Format(txt_Date_to.Text, "MM-DD-YYYY") & "# ORDER BY TBL_Income.[IncDate] ASC"
    
    lbox_Income_List.Clear
    lbox_Income_List.ColumnCount = 6
    lbox_Income_List.ColumnWidths = "80,80,80,80,80,100"
    lbox_Income_List.AddItem
    lbox_Income_List.list(0, 0) = "Date"
    lbox_Income_List.list(0, 1) = "Particular"
    lbox_Income_List.list(0, 2) = "Source"
    lbox_Income_List.list(0, 3) = "Mode"
    lbox_Income_List.list(0, 4) = "Amount"
    lbox_Income_List.list(0, 5) = "Narration"
    I = 1
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    Do Until .EOF
                       lbox_Income_List.AddItem
                       lbox_Income_List.list(I, 0) = VBA.Format(.Fields("IncDate"), "DD-MM-YYYY")
                       lbox_Income_List.list(I, 1) = .Fields("Particular")
                       lbox_Income_List.list(I, 2) = .Fields("Bank/Cash")
                       lbox_Income_List.list(I, 3) = .Fields("Rec From")
                       lbox_Income_List.list(I, 4) = .Fields("Amount")
                            If .Fields("Narration") <> "" Then
                             lbox_Income_List.list(I, 5) = .Fields("Narration")
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
    Dim I As Byte
    
        SQL = "Select * From TBL_Payment_Mode"
        I = 0
    'Calling connection and opening recordset for updating in the combobox in userform
    
    Call Connection_Vindhya_Cash_Flow(Con)
    DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    txt_Date.Text = VBA.Format(VBA.Date, "DD/MM/YYYY")
    
        Cbox_Payment_Mode.Clear
        Cbox_Payment_Mode.ColumnCount = 1
        Cbox_Payment_Mode.ColumnWidths = "120"
    
        'Updating the payment mode list in the userform
        
        
    With DataRec
        If .BOF = True And .EOF = True Then
            GoTo IncomeSource
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
    
IncomeSource:

    DataRec.Close
    


    SQL = "Select * from TBL_Income_Mode"
    
    DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    'updating the Income source in the Userform
    
        Cbox_IncomeSource.Clear
        Cbox_IncomeSource.ColumnCount = 1
        Cbox_IncomeSource.ColumnWidths = "120"
            I = 0
        With DataRec
            If .BOF = True And .EOF = True Then
                GoTo Final
            End If
            .MoveFirst
            
            
            Do Until .EOF
                If IsNull(.Fields("Source")) = False Then
                    Cbox_IncomeSource.AddItem
                    Cbox_IncomeSource.list(I, 0) = .Fields("Source")
                    I = I + 1
                End If
                
            .MoveNext
            Loop
            
            
            
        End With
        
Final:
    DataRec.Close
    Con.Close
    If shMain.Range("F7").Value = "Admin" Then
    
        lbox_Income_List.Clear
        lbox_Income_List.ColumnCount = 6
        lbox_Income_List.ColumnWidths = "80,80,80,80,80,100"
        lbox_Income_List.AddItem
        lbox_Income_List.list(0, 0) = "Date"
        lbox_Income_List.list(0, 1) = "Particular"
        lbox_Income_List.list(0, 2) = "Source"
        lbox_Income_List.list(0, 3) = "Mode"
        lbox_Income_List.list(0, 4) = "Amount"
        lbox_Income_List.list(0, 5) = "Narration"
        Frm_IncomeList.Visible = True
    Else
        
        Frm_IncomeList.Visible = False
    
    End If
    
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
   Cancel = 1
End Sub
