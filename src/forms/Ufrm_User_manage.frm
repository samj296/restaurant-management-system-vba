VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_User_manage 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   13215
   ClientLeft      =   510
   ClientTop       =   855
   ClientWidth     =   18855
   OleObjectBlob   =   "Ufrm_User_manage.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_User_manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Add_Staff_Click()
    
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    
    Dim Post As String
    Dim STaff As New cls_Staff_User
    
    
        Call Connection_Vindhya_Cash_Flow(Con)
                'Checking in the Staff list table whether it has the same name
        SQL = "Select * From TBL_Staff_Detail WHERE TBL_Staff_Detail.[Staff Name]='" & txt_Staff_Full_Name.Text & "'"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                MsgBox "The staff with the same name is already in the database please change the name to proceed.", , "THE VINDHYA CAFE"
                txt_Staff_Full_Name.Text = ""
                Exit Sub
            End If
        End With
        
        Rec.Close
                'Checking if the Ledger whether it has the same name
        SQL = "Select * From TBL_Ledger WHERE TBL_Ledger.[Ledger Cat]='" & txt_Staff_Full_Name.Text & "'"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                MsgBox "Please enter different name as the this name is already in Expense Category", , "THE VINDHYA CAFE"
                txt_Staff_Full_Name.Text = ""
                Exit Sub
            End If
        
        
        End With
        
        Rec.Close
    

    
        SQL = "Select * From TBL_Staff_Detail ORDER BY TBL_Staff_Detail.[ID] ASC"
                    'updating in Staff Detail
        
            
            If txt_Staff_Full_Name.Text = "" Then
                MsgBox "Enter the staff name to update the staff detail.", , "THE VINDHYA CAFE"
                Exit Sub
            End If
        
        STaff.StaffName = txt_Staff_Full_Name.Text
        
            If IsNumeric(txt_Staff_Salary.Text) = False Then
                MsgBox "Please enter proper amount to proceed", , "THE VINDHYA CAFE"
                txt_Staff_Salary.Text = ""
                Exit Sub
            End If
            
        STaff.Salary = txt_Staff_Salary.Text
        
        If VBA.Format(VBA.Date, "MM") < 4 Then

                    STaff.FYear = CLng(VBA.Format(VBA.Date, "YYYY")) - 1

            ElseIf VBA.Format(VBA.Date, "MM") > 3 Then
                    STaff.FYear = VBA.Format(VBA.Date, "YYYY")

            End If
            If IsDate(txt_Joining_Date.Text) = True Then
                STaff.JoinDate = VBA.Format(txt_Joining_Date.Text, "DD-MM-YYYY")
                STaff.JoInMonth = CByte(VBA.Format(STaff.JoinDate, "MM"))
                STaff.JoinDay = CByte(VBA.Format(STaff.JoinDate, "dd")) - 1
            Else
                MsgBox "Please enter the correct Joining Date to proceed.", , "THE VINDHYA CAFE"
            End If
        
        If txt_Post.Text = "" Then
            MsgBox "Enter the post to update the staff detail.", , "THE VINDHYA CAFE"
            Exit Sub
        End If
        Post = txt_Post.Text
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveLast
                    STaff.UserId = .Fields("ID") + 1
                Else
                    STaff.UserId = 1
                End If
                
                If txt_Staff_Full_Name.Text <> "" And txt_Joining_Date.Text <> "" And txt_Staff_Salary.Text <> "" And txt_Post.Text <> "" Then
                    .AddNew
                    .Fields("Staff Name") = STaff.StaffName
                    .Fields("Post") = Post
                    .Fields("ID") = STaff.UserId
                    .Fields("Joining Date") = STaff.JoinDate
                    .Update
                End If
            End With
            
            
            
            
        Rec.Close
        

            
            
        Con.Close
        
        STaff.SalaryUpdate

            txt_Staff_Full_Name.Text = ""
            txt_Staff_Salary.Text = ""
            txt_Joining_Date.Text = ""
            txt_Post.Text = ""
            
        Call UserForm_Initialize
End Sub

Private Sub btn_Add_User_Click()
    If txt_User_Full_Name.Text <> "" And txt_Password.Text <> "" And cbox_Admin_rights <> "" Then
        Dim Con As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim SQL As String
            
            If cbox_Admin_rights = "Admin" Or cbox_Admin_rights = "Restricted" Or cbox_Admin_rights = "Moderate" Then
                    SQL = "Select * From TBL_User"
                    
                    Call Connection_Vindhya_Main_File(Con)
                    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                    
                    With Rec
                        .AddNew
                        .Fields("Full Name") = txt_User_Full_Name.Text
                        .Fields("User Name") = txt_UserID.Text
                        .Fields("Password") = txt_Password.Text
                        .Fields("Admin Rights") = cbox_Admin_rights.Text
                        .Update
                    
                    End With
            Else
                
                MsgBox "Please Select the proper Admin Rights to add the User.", , "THE VINDHYA CAFE"
            
            End If
    
    End If
    
                Rec.Close
                Con.Close
                
    txt_User_Full_Name.Text = ""
    txt_UserID.Text = ""
    txt_Password.Text = ""
    cbox_Admin_rights.Text = ""
    
    Call UserForm_Initialize
End Sub

Private Sub btn_Back_Click()
    Me.Hide
    Ufrm_DashBoard.Show
End Sub

Private Sub btn_Delete_Staff_Click()
    Dim M As Long
    Dim STaff As New cls_Staff_User
    
    If IsDate(txt_UserInput_Staff.Text) = False Then
        MsgBox "Input Valid date the staff left in the Userinput box.", , "THE VINDHYA CAFE"
        Exit Sub
    End If
    
    For M = 1 To lbox_Staff_List.ListCount - 1
        If lbox_Staff_List.Selected(M) = True Then
            
            STaff.UserId = lbox_Staff_List.list(M, 0)
            STaff.StaffName = lbox_Staff_List.list(M, 1)
            STaff.Salary = lbox_Staff_List.list(M, 3)
            STaff.JoinDate = lbox_Staff_List.list(M, 4)
            STaff.ResignDate = Format(txt_UserInput_Staff.Text, "DD-MM-YYYY")
            Exit For
        End If
    Next M
    
    STaff.ResignProceed
    txt_UserInput_Staff.Text = ""
    MsgBox "Updated data successfully.", , "THE VINDHYA CAFE"
    Call UserForm_Initialize
    
End Sub

Private Sub btn_Delete_User_Click()
    Dim I As Long
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim IDNO As Long
        IDNO = 0
        For I = 1 To lbox_User_List.ListCount - 1
            If lbox_User_List.Selected(I) = True Then
                IDNO = lbox_User_List.list(I, 3)
                
            End If
            
        Next I
        
        If IDNO = 0 Then
            MsgBox "Please Select the Proper user to Delete the User.", , "THE VINDHYA CAFE"
            Exit Sub
        End If
    
    SQL = "Select * From TBL_User WHERE TBL_User.[UserID_NO]= " & IDNO
        
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            .Delete
            .Update
            
        End If
        
    
    End With
    
    Rec.Close
    Con.Close
    
    Call UserForm_Initialize
    
End Sub

Private Sub btn_Edit_Staff_Click()

'        cbox_Staff_Edit_List.AddItem
'        cbox_Staff_Edit_List.list(0) = "Salary"
'        cbox_Staff_Edit_List.AddItem
'        cbox_Staff_Edit_List.list(1) = "Post"
'        cbox_Staff_Edit_List.AddItem
'        cbox_Staff_Edit_List.list(2) = "Joining Date"
        
        
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim UserInput As String
    Dim User As New cls_Staff_User
    Dim I As Long
    Dim FYear As String
    Dim MonthStat As Boolean
    Dim UserStat As Boolean
    Dim oldSalary As Double
    
    
        MonthStat = False
        UserStat = False
     If VBA.Format(VBA.Date, "MM") < 4 Then
        
            FYear = CLng(VBA.Format(VBA.Date, "YYYY")) - 1
            
    ElseIf VBA.Format(VBA.Date, "MM") > 3 Then
            FYear = VBA.Format(VBA.Date, "YYYY")
            
    End If
        
        For I = 1 To lbox_Staff_List.ListCount - 1
            If lbox_Staff_List.Selected(I) = True Then
                User.UserId = lbox_Staff_List.list(I, 0)
                User.StaffName = lbox_Staff_List.list(I, 1)
                oldSalary = lbox_Staff_List.list(I, 3)
                UserStat = True
                Exit For
            End If
        Next I
        
            If UserStat = False Then
                MsgBox "No user Selected please select the user to continue"
                Exit Sub
            End If
    
    
   Select Case cbox_Staff_Edit_List.Text
        Case Is = "Salary"
            If IsNumeric(txt_UserInput_Staff.Text) = False Then
                MsgBox "Salary should be in number only", , "THE VINDHYA CAFE"
                Exit Sub
            Else
                User.Salary = txt_UserInput_Staff.Text
            End If
            
            For I = 1 To 12
                If cbox_month_Staff.Text = VBA.MonthName(I) Then
                    User.JoInMonth = I
                    MonthStat = True
                    Exit For
                End If
                
            Next I
                If MonthStat = False Then
                    MsgBox "Month is not valid", , "THE VINDHYA CAFE"
                    Exit Sub
                End If
                User.UserType = "Old"
              User.SalaryUpdate
              cbox_Staff_Edit_List.Text = ""
              txt_UserInput_Staff.Text = ""
              Call UserForm_Initialize
        Case Is = "Post"
            SQL = "Select * From TBL_Staff_Detail WHERE TBL_Staff_Detail.[ID]=" & User.UserId
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        .Fields("Post") = txt_UserInput_Staff.Text
                        .Update
                    End If
                End With
                
            Rec.Close
            Con.Close
            
        Case Is = "Joining Date"
            'deleting the old salary
            
            User.UserType = "Old"
            User.Salary = oldSalary
            
            If IsDate(txt_UserInput_Staff.Text) = False Then
                MsgBox "Invalid joining date. Please enter the joining date again.", , "THE VINDHYA CAFE"
                txt_UserInput_Staff.Text = ""
                Exit Sub
            End If
            
            User.JoinDate = txt_UserInput_Staff.Text
            
            
            For I = 1 To 12
                If cbox_month_Staff.Text = VBA.MonthName(I) Then
                    User.JoInMonth = I
                    MonthStat = True
                    Exit For
                End If
                
            Next I
                If MonthStat = False Then
                    MsgBox "Month is not valid", , "THE VINDHYA CAFE"
                    Exit Sub
                End If
            
            
            
            SQL = "Select * From TBL_Staff_Salary_Detail Where TBL_Staff_Salary_Detail.[Staff Id]=" & User.UserId & " AND TBL_Salary_Detail.[Salary (Year)]=#" & FYear & "#"
            Call Connection_Vindhya_Cash_Flow(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                For I = 1 To 12
                    .Fields(MonthName(I)) = 0
                Next I
                .Update
            End With
                  
            Rec.Close
            Con.Close
            
            
            User.SalaryUpdate
            
   End Select
   
   


End Sub

Private Sub btn_Edit_User_Click()
        If cbox_User_Edit_List.Text <> "" Then
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            Dim I As Long
            Dim IDNO As Long
            Dim EditItem As String
                EditItem = cbox_User_Edit_List.Text
                
                IDNO = 0
                For I = 1 To lbox_User_List.ListCount - 1
                    If lbox_User_List.Selected(I) = True Then
                        IDNO = lbox_User_List.list(I, 3)
                    End If
                    
                Next I
                If IDNO = 0 Then
                    MsgBox "Please select the proper User to Edit"
                    Exit Sub
                End If
                
                If EditItem = "Full Name" Or EditItem = "User Name" Or EditItem = "Password" Or EditItem = "Admin Rights" Then
        

                        
                        SQL = "Select * FROM TBL_User WHERE TBL_User.[UserID_NO]=" & IDNO
                        
                        Call Connection_Vindhya_Main_File(Con)
                        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                        
                            With Rec
                                If .EOF = False Or .BOF = False Then
                                    If EditItem = "Admin Rights" And cbox_UserInput.Text <> "" Then
                                        .Fields("Admin Rights") = cbox_UserInput.Text
                                        .Update
                                    ElseIf txt_User_Input.Text <> "" And (EditItem = "Full Name" Or EditItem = "User Name" Or EditItem = "Password") Then
                                        .Fields(EditItem) = txt_User_Input.Text
                                        .Update
                                    End If
                                
                                End If
                                
                            
                            End With
                        
                        cbox_User_Edit_List.Text = ""
                        cbox_UserInput.Text = ""
                        txt_User_Input.Text = ""
                        
                        Rec.Close
                        Con.Close
                        
                        Call UserForm_Initialize
                        
                End If
                
        ElseIf cbox_User_Edit_List.Text = "" Then
                    
             MsgBox "Select the option from the dropdown the item that need to be edited.", , "THE VINDHYA CAFE"
             Exit Sub
            
        End If
End Sub

Private Sub cbox_Staff_Edit_List_Change()
    If cbox_Staff_Edit_List.Text = "Salary" Then
        cbox_month_Staff.Visible = True
        lbl_month_Staff.Visible = True
    Else
        cbox_month_Staff.Visible = False
        lbl_month_Staff.Visible = False
    End If
End Sub

Private Sub cbox_User_Edit_List_Change()
    If cbox_User_Edit_List.Text = "Password" Then
        txt_User_Input.PasswordChar = "*"
        cbox_UserInput.Visible = False
        txt_User_Input.Visible = True
    ElseIf cbox_User_Edit_List = "Admin Rights" Then
        txt_User_Input.Visible = False
        cbox_UserInput.Text = ""
        cbox_UserInput.Visible = True
    Else
        txt_User_Input.PasswordChar = ""
        txt_User_Input.Visible = True
        cbox_UserInput.Visible = False
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim CurMonth As String
    Dim UserName As cls_Staff_User
    
        SQL = "Select * From TBL_User"
        
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        CurMonth = VBA.Format(Now, "MMMM")
        
        cbox_UserInput.Visible = False
        
        lbox_User_List.Clear
        lbox_User_List.ColumnCount = 4
        lbox_User_List.ColumnWidths = "100,120,120,60"
        
        lbox_User_List.AddItem
        lbox_User_List.list(0, 0) = "Full Name"
        lbox_User_List.list(0, 1) = "User Name"
        lbox_User_List.list(0, 2) = "Admin Rights"
        lbox_User_List.list(0, 3) = "ID No"
        
        cbox_month_Staff.Clear
        cbox_month_Staff.ColumnCount = 1
        cbox_month_Staff.ColumnWidths = "100"
        For I = 1 To 12
            cbox_month_Staff.AddItem
            cbox_month_Staff.list(I - 1) = VBA.MonthName(I)
        Next I
        lbl_month_Staff.Visible = False
        cbox_month_Staff.Visible = False
        
        cbox_User_Edit_List.Clear
        cbox_User_Edit_List.ColumnCount = 1
        cbox_User_Edit_List.ColumnWidths = "100"
        cbox_User_Edit_List.AddItem
        cbox_User_Edit_List.list(0) = "Full Name"
        cbox_User_Edit_List.AddItem
        cbox_User_Edit_List.list(1) = "User Name"
        cbox_User_Edit_List.AddItem
        cbox_User_Edit_List.list(2) = "Password"
        cbox_User_Edit_List.AddItem
        cbox_User_Edit_List.list(3) = "Admin Rights"
        
        cbox_Staff_Edit_List.Clear
        cbox_Staff_Edit_List.ColumnCount = 1
        cbox_Staff_Edit_List.ColumnWidths = "100"
        
        cbox_Staff_Edit_List.AddItem
        cbox_Staff_Edit_List.list(0) = "Salary"
        cbox_Staff_Edit_List.AddItem
        cbox_Staff_Edit_List.list(1) = "Post"
        cbox_Staff_Edit_List.AddItem
        cbox_Staff_Edit_List.list(2) = "Joining Date"
        
        cbox_Admin_rights.Clear
        cbox_Admin_rights.ColumnCount = 1
        cbox_Admin_rights.ColumnWidths = "100"
        cbox_Admin_rights.AddItem
        cbox_Admin_rights.list(0) = "Admin"
        cbox_Admin_rights.AddItem
        cbox_Admin_rights.list(1) = "Restricted"
        cbox_Admin_rights.AddItem
        cbox_Admin_rights.list(2) = "Moderate"
        
        cbox_UserInput.Clear
        cbox_UserInput.ColumnCount = 1
        cbox_UserInput.ColumnWidths = "100"
        cbox_UserInput.AddItem
        cbox_UserInput.list(0) = "Admin"
        cbox_UserInput.AddItem
        cbox_UserInput.list(1) = "Restricted"
        cbox_UserInput.AddItem
        cbox_UserInput.list(2) = "Moderate"
        
        lbox_Staff_List.Clear
        lbox_Staff_List.ColumnCount = 6
        lbox_Staff_List.ColumnWidths = "50,80,50,80,80,80"
        
        lbox_Staff_List.AddItem
        lbox_Staff_List.list(0, 0) = "ID No"
        lbox_Staff_List.list(0, 1) = "Name"
        lbox_Staff_List.list(0, 2) = "Post"
        lbox_Staff_List.list(0, 3) = CurMonth & " - Salary"
        lbox_Staff_List.list(0, 4) = "Join On"
        lbox_Staff_List.list(0, 5) = "Left On"
        
        I = 1
        
        With Rec
            
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                Do Until .EOF
                    lbox_User_List.AddItem
                    lbox_User_List.list(I, 0) = .Fields("Full Name")
                    lbox_User_List.list(I, 1) = .Fields("User Name")
                    lbox_User_List.list(I, 2) = .Fields("Admin Rights")
                    lbox_User_List.list(I, 3) = .Fields("UserID_NO")
                    
                    I = I + 1
                    
                    
                .MoveNext
                Loop
                
            End If
            
        End With
    
    Rec.Close
    Con.Close
    
    I = 1
    
    Call Connection_Vindhya_Cash_Flow(Con)
    
    SQL = "Select * From TBL_Staff_Detail"
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                Set UserName = New cls_Staff_User
                UserName.StaffName = .Fields("Staff Name")
                lbox_Staff_List.AddItem
                lbox_Staff_List.list(I, 0) = .Fields("ID")
                lbox_Staff_List.list(I, 1) = UserName.StaffName
                lbox_Staff_List.list(I, 2) = .Fields("Post")
                lbox_Staff_List.list(I, 4) = VBA.Format(.Fields("Joining Date"), "DD-MM-YYYY")
                
                If IsNull(.Fields("Reisgn On")) = False Then
                    lbox_Staff_List.list(I, 5) = .Fields("Reisgn On")
                End If
     
                I = I + 1
            .MoveNext
            Loop
            
        End If
    
    End With
    
        
    Rec.Close
    
    SQL = "Select * From TBL_Staff_Salary_Detail"
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        For I = 1 To lbox_Staff_List.ListCount - 1
            Set UserName = New cls_Staff_User
            UserName.StaffName = lbox_Staff_List.list(I, 1)
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                Do Until .EOF
                    If .Fields("Staff name") = UserName.StaffName Then
                        lbox_Staff_List.list(I, 3) = .Fields(CurMonth)
                    End If
                    .MoveNext
                Loop
            End If
        Next I
    
    
    End With
    
    Rec.Close
    Con.Close
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
   Cancel = 1

End Sub
