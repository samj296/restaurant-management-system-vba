VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Salary_Detail 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   13210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18855
   OleObjectBlob   =   "Ufrm_Salary_Detail.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Salary_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btn_Back_Click()
       Me.Hide
    
    Ufrm_DashBoard.Show
End Sub

Private Sub btn_Export_Click()
    Dim WB As Workbook
    Dim I As Double
    Dim Sh As Worksheet
    Dim FPath As String
    Dim IDNO As Long
    Dim STaff As String
    
        For I = 1 To lbox_Staff_List.ListCount - 1
            If lbox_Staff_List.Selected(I) = True Then
                IDNO = lbox_Staff_List.list(I, 0)
                STaff = lbox_Staff_List.list(I, 1)
            End If
        Next I
    
    Set WB = Workbooks.Add
    Set Sh = WB.Worksheets(1)
        Sh.Name = "Paid Amount"
        
    If lbox_Paid_List.ListCount > 1 Then
        For I = 0 To lbox_Paid_List.ListCount - 1
            Sh.Range("A" & I + 1).Value = lbox_Paid_List.list(I, 0)
            Sh.Range("B" & I + 1).Value = lbox_Paid_List.list(I, 1)
            Sh.Range("C" & I + 1).Value = lbox_Paid_List.list(I, 2)
            Sh.Range("D" & I + 1).Value = lbox_Paid_List.list(I, 3)
            Sh.Range("E" & I + 1).Value = lbox_Paid_List.list(I, 4)
            
        Next I
    End If
        Sh.Range("A1:E1").Interior.Color = RGB(0, 153, 255)
        Sh.Range("A1:E1").Font.Color = RGB(255, 255, 255)
        Sh.Columns("A:E").AutoFit
        
        Set Sh = WB.Worksheets.Add
        Sh.Name = "Salary Detail"
        
    If lbox_Salary_Detail.ListCount > 1 Then
        For I = 0 To lbox_Salary_Detail.ListCount - 1
            Sh.Range("A" & I + 1).Value = lbox_Salary_Detail.list(I, 0)
            Sh.Range("B" & I + 1).Value = lbox_Salary_Detail.list(I, 1)
            Sh.Range("C" & I + 1).Value = lbox_Salary_Detail.list(I, 2)
            Sh.Range("D" & I + 1).Value = lbox_Salary_Detail.list(I, 3)
            Sh.Range("E" & I + 1).Value = lbox_Salary_Detail.list(I, 4)
            Sh.Range("F" & I + 1).Value = lbox_Salary_Detail.list(I, 5)
            Sh.Range("G" & I + 1).Value = lbox_Salary_Detail.list(I, 6)
            Sh.Range("H" & I + 1).Value = lbox_Salary_Detail.list(I, 7)
            
        Next I
    End If
        Sh.Range("A1:H1").Interior.Color = RGB(0, 153, 255)
        Sh.Range("A1:H1").Font.Color = RGB(255, 255, 255)
        Sh.Columns("A:H").AutoFit
        
       With Application.FileDialog(msoFileDialogFolderPicker)
            .ButtonName = "Select the folder to export"
            .Title = "THE VINDHYA CAFE"
            .Show
            FPath = .SelectedItems(1)
       End With
    
    WB.SaveAs FileName:=FPath & "\" & STaff
    WB.Close
    
End Sub

Private Sub btn_Update_Click()
    Dim STaff As String
    Dim IDNO As Long
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    Dim M As Boolean
    Dim D_G As Boolean
    
    
    
    
    For I = 1 To lbox_Staff_List.ListCount - 1
        If lbox_Staff_List.Selected(I) = True Then
            IDNO = lbox_Staff_List.list(I, 0)
            STaff = lbox_Staff_List.list(I, 1)
            GoTo NextStep
        End If
    Next I
        
        If IDNO = 0 Then
            MsgBox "Please select the user from the list.", , "THE VINDHYA CAFE"
        End If
    
    Exit Sub

NextStep:
        M = False
    For I = 1 To 12
        If cbox_Month.Text = MonthName(I) Then
            M = True
            GoTo FinalUpdate
        End If
        
    Next I
    
        If M = False Then
            MsgBox "Please select the proper month from the dropdown list.", , "THE VINDHYA CAFE"
            Exit Sub
        End If
        
    
FinalUpdate:

    If cbox_Cat = "Gift/Bonus" Then
        SQL = "Select * From TBL_Staff_Gift_Bonus WHERE TBL_Staff_Gift_Bonus.[Id no]=" & IDNO & " AND TBL_Staff_Gift_Bonus.[Bonus Year]='" & txt_Year.Text & "'"
        D_G = True
    ElseIf cbox_Cat = "Deduction" Then
        SQL = "Select * From TBL_Staff_Deduction WHERE TBL_Staff_Deduction.[Id no]=" & IDNO & " AND TBL_Staff_Deduction.[Deduction Year]='" & txt_Year.Text & "'"
        D_G = False
    Else
        MsgBox "Select the appropriate category from the dropdown list.", , "THE VINDHYA CAFE"
        Exit Sub
    End If
    
    If IsNumeric(txt_Amount.Text) = False Then
        
        MsgBox "Please enter the correct amount.", , "THE VINDHYA CAFE"
        Exit Sub
    End If
    
    
    


    Call Connection_Vindhya_Cash_Flow(Con)
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
           .Fields(cbox_Month.Text) = txt_Amount.Text
           .Update
        End If
    End With
    
    
    Rec.Close
    
    SQL = "Select * From TBL_Staff_Salary_Detail WHERE TBL_Staff_Salary_Detail.[Staff Id]=" & IDNO & " AND TBL_Staff_Salary_Detail.[Salary (Year)]='" & txt_Year.Text & "'"
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                If D_G = True Then
                    I = .Fields("Final Amnt(Deduction and Gift)")
                    .Fields("Final Amnt(Deduction and Gift)") = I + CDbl(txt_Amount.Text)
                    .Update
                Else
                End If
                
            End If
        End With
    Rec.Close
    Con.Close
    
    Call lbox_Staff_List_Change
    
End Sub

Private Sub lbox_Staff_List_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    Dim STaff As String
    Dim IDNO As Long
    Dim TAmnt As Double
    Dim ForFooter As Double
    Dim CMonth As Byte
    Dim DueCmonth As Double
    Dim M As Byte
    Dim FYear As String
    Dim CalYear As String
    Dim Salary, Gift, Deduction, Paid, R As Byte, C As Byte
    
    M = 4
    
   FYear = txt_Year.Text

    ReDim Salary(1 To 12)
    ReDim Gift(1 To 12)
    ReDim Deduction(1 To 12)
    ReDim Paid(1 To 12)
    
    
    
    
    lbox_Salary_Detail.Clear
    lbox_Salary_Detail.ColumnCount = 8
    lbox_Salary_Detail.ColumnWidths = "80,80,100,100,100,100,100,100"
    lbox_Salary_Detail.AddItem
    lbox_Salary_Detail.list(0, 0) = "Due On"
    lbox_Salary_Detail.list(0, 1) = "Month"
    lbox_Salary_Detail.list(0, 2) = "Deduction"
    lbox_Salary_Detail.list(0, 3) = "Gift/Bonus"
    lbox_Salary_Detail.list(0, 4) = "Salary Amount"
    lbox_Salary_Detail.list(0, 5) = "Final Amnt"
    lbox_Salary_Detail.list(0, 6) = "Amount Paid"
    lbox_Salary_Detail.list(0, 7) = "Due Amount"
    lbox_Paid_List.Clear
    lbox_Paid_List.ColumnCount = 5
    lbox_Paid_List.ColumnWidths = "80,120,80,100,120"
    lbox_Paid_List.AddItem
    lbox_Paid_List.list(0, 0) = "Date"
    lbox_Paid_List.list(0, 1) = "Particular"
    lbox_Paid_List.list(0, 2) = "Mode"
    lbox_Paid_List.list(0, 3) = "Amount"
    lbox_Paid_List.list(0, 4) = "Narration"
    
    lbl_Total_Amnt_Paid.Caption = ""
    
        For I = 1 To lbox_Staff_List.ListCount - 1
            If lbox_Staff_List.Selected(I) = True Then
                IDNO = lbox_Staff_List.list(I, 0)
                STaff = lbox_Staff_List.list(I, 1)
                GoTo Proceed
            End If
            
        Next I
        
Exit Sub
    
Proceed:
    SQL = "Select SUM(TBL_Expense.[Amount]) AS TOTAL From TBL_Expense WHERE TBL_Expense.[Ledger Cat] = '" & STaff & "'"
    
    Call Connection_Vindhya_Cash_Flow(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .BOF = False Or .EOF = False Then
            .MoveFirst
            If IsNull(.Fields("Total")) = False Then
                TAmnt = .Fields("TOTAL")
            Else
                TAmnt = 0
            End If
        Else
            TAmnt = 0
        End If
        lbl_Total_Amnt_Paid.Caption = TAmnt
        ForFooter = TAmnt
    End With
    
    Rec.Close
    
    
    SQL = "Select * From TBL_Staff_Deduction WHERE TBL_Staff_Deduction.[Id no]=" & IDNO & " AND TBL_Staff_Deduction.[Deduction Year]='" & FYear & "'"
        
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        With Rec
            
                For R = 1 To 12
                    If .EOF = False Or .BOF = False Then
                        Deduction(R) = .Fields(MonthName(R))
                    Else
                        Deduction(R) = 0
                    End If
                Next R
                
           
        
        End With
            
            
        
        Rec.Close
        
        SQL = "Select * From TBL_Staff_Gift_Bonus WHERE TBL_Staff_Gift_Bonus.[Id no]=" & IDNO & " AND TBL_Staff_Gift_Bonus.[Bonus Year]='" & FYear & "'"
            
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                
                    For R = 1 To 12
                    If .EOF = False Or .BOF = False Then
                        Gift(R) = .Fields(MonthName(R))
                    Else
                        Gift(R) = 0
                    End If
                    Next R
                
                
            
            End With
        Rec.Close
        
        SQL = "Select * From TBL_Staff_Salary_Detail WHERE TBL_Staff_Salary_Detail.[Staff Id]=" & IDNO & " AND TBL_Staff_Salary_Detail.[Salary (Year)]='" & FYear & "'"
        
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                
                    For R = 1 To 12
                     If .EOF = False Or .BOF = False Then
                        Salary(R) = .Fields(MonthName(R))
                     Else
                        Salary(R) = 0
                     End If
                    Next R
                    
                
            
            End With
        Rec.Close
        Con.Close
        
        
    CMonth = VBA.Format(VBA.Date, "MM")
        
    For I = 1 To 12
        lbox_Salary_Detail.AddItem
        If M < 4 And M > 1 Then
            CalYear = CLng(FYear) + 1
        Else
        CalYear = FYear
        End If
        If M = 12 Then
            lbox_Salary_Detail.list(I, 0) = "05-" & MonthName(1) & "-" & CalYear
        Else
            lbox_Salary_Detail.list(I, 0) = "05-" & MonthName(M + 1) & "-" & CalYear
        End If

        lbox_Salary_Detail.list(I, 1) = MonthName(M)
        lbox_Salary_Detail.list(I, 2) = Deduction(M)
        lbox_Salary_Detail.list(I, 3) = Gift(M)
        lbox_Salary_Detail.list(I, 4) = Salary(M)
        lbox_Salary_Detail.list(I, 5) = (CDbl(lbox_Salary_Detail.list(I, 4)) + CDbl(lbox_Salary_Detail.list(I, 3))) - CDbl(lbox_Salary_Detail.list(I, 2))
        
        If TAmnt > CDbl(lbox_Salary_Detail.list(I, 5)) Or TAmnt = CDbl(lbox_Salary_Detail.list(I, 5)) Then
            lbox_Salary_Detail.list(I, 6) = lbox_Salary_Detail.list(I, 5)
            
        ElseIf TAmnt < CDbl(lbox_Salary_Detail.list(I, 5)) Then
            lbox_Salary_Detail.list(I, 6) = TAmnt
            TAmnt = 0
        End If
        
        If CDbl(lbox_Salary_Detail.list(I, 6)) < lbox_Salary_Detail.list(I, 5) Then
            lbox_Salary_Detail.list(I, 7) = CDbl(lbox_Salary_Detail.list(I, 5)) - CDbl(lbox_Salary_Detail.list(I, 6))
        Else
            lbox_Salary_Detail.list(I, 7) = 0
        End If
        
      If VBA.Format(VBA.Date, "DD-MM-YYYY") >= VBA.Format(lbox_Salary_Detail.list(I, 0), "DD-MM-YYYY") Then
        DueCmonth = DueCmonth + CDbl(lbox_Salary_Detail.list(I, 5))
      End If
        
        If M = 12 Then
            M = 1
        Else
            M = M + 1
        End If
    
    Next I
        
        lbox_Salary_Detail.AddItem
        lbox_Salary_Detail.list(13, 0) = "--------------------------"
        lbox_Salary_Detail.list(13, 1) = "--------------------------"
        lbox_Salary_Detail.list(13, 2) = "--------------------------"
        lbox_Salary_Detail.list(13, 3) = "--------------------------"
        lbox_Salary_Detail.list(13, 4) = "--------------------------"
        lbox_Salary_Detail.list(13, 5) = "--------------------------"
        lbox_Salary_Detail.list(13, 6) = "--------------------------"
        lbox_Salary_Detail.list(13, 7) = "--------------------------"
            
        lbox_Salary_Detail.AddItem
        lbox_Salary_Detail.list(14, 0) = "Total due till date"
        lbox_Salary_Detail.list(14, 1) = DueCmonth
        
        
        
        Dim oldSalary As Double
        Dim OldDeduction As Double
        Dim OldGift As Double
        Dim OldFAmnt As Double
        
        SQL = "Select * From TBL_Staff_Salary_Detail WHERE TBL_Staff_Salary_Detail.[Staff Id]=" & IDNO
        
        Call Connection_Vindhya_Cash_Flow(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        oldSalary = 0
        With Rec
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                Do Until .EOF
                    If .Fields("Salary (Year)") < FYear Then
                        
                        oldSalary = oldSalary + .Fields("Final Amnt(Deduction and Gift)")
                        
                    End If
                .MoveNext
                Loop
                
            
            End If
            
        End With
        
        Rec.Close
        
        SQL = "Select * From TBl_Expense WHERE TBl_Expense.[Ledger Cat]='" & STaff & "'"
        I = 1
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                Do Until .EOF
                    lbox_Paid_List.AddItem
                    lbox_Paid_List.list(I, 0) = .Fields("ExpDate")
                    lbox_Paid_List.list(I, 1) = .Fields("Particular")
                    lbox_Paid_List.list(I, 2) = .Fields("Payment Mode")
                    lbox_Paid_List.list(I, 3) = .Fields("Amount")
                    If IsNull(.Fields("Narration")) = False Then
                        lbox_Paid_List.list(I, 4) = .Fields("Narration")
                    Else
                        lbox_Paid_List.list(I, 4) = ""
                    End If
                    I = I + 1
                    .MoveNext
                Loop
            End If
        
        End With
        
        
        Rec.Close
        Con.Close
        
        If CDbl(DueCmonth) <= CDbl(ForFooter) Then
            lbox_Salary_Detail.AddItem
            lbox_Salary_Detail.list(15, 0) = "Advance Money"
            lbox_Salary_Detail.list(15, 1) = (CDbl(ForFooter) + oldSalary) - CDbl(DueCmonth)
        Else
            lbox_Salary_Detail.AddItem
            lbox_Salary_Detail.list(15, 0) = "Due Amount"
            lbox_Salary_Detail.list(15, 1) = CDbl(DueCmonth) + oldSalary - CDbl(ForFooter)
        End If
        
        
    
End Sub



Private Sub UserForm_Initialize()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    
    lbox_Staff_List.Clear
    lbox_Staff_List.ColumnCount = 2
    lbox_Staff_List.ColumnWidths = "80,120"
    lbox_Staff_List.AddItem
    lbox_Staff_List.list(0, 0) = "ID NO"
    lbox_Staff_List.list(0, 1) = "Name"
    lbox_Salary_Detail.Clear
    lbox_Salary_Detail.ColumnCount = 8
    lbox_Salary_Detail.ColumnWidths = "80,80,100,100,100,100,100,100"
    lbox_Salary_Detail.AddItem
    lbox_Salary_Detail.list(0, 0) = "Due On"
    lbox_Salary_Detail.list(0, 1) = "Month"
    lbox_Salary_Detail.list(0, 2) = "Deduction"
    lbox_Salary_Detail.list(0, 3) = "Gift/Bonus"
    lbox_Salary_Detail.list(0, 4) = "Salary Amount"
    lbox_Salary_Detail.list(0, 5) = "Final Amnt"
    lbox_Salary_Detail.list(0, 6) = "Amount Paid"
    lbox_Salary_Detail.list(0, 7) = "Due Amount"
    
    lbox_Paid_List.Clear
    lbox_Paid_List.ColumnCount = 5
    lbox_Paid_List.ColumnWidths = "80,120,80,100,120"
    lbox_Paid_List.AddItem
    lbox_Paid_List.list(0, 0) = "Date"
    lbox_Paid_List.list(0, 1) = "Particular"
    lbox_Paid_List.list(0, 2) = "Mode"
    lbox_Paid_List.list(0, 3) = "Amount"
    lbox_Paid_List.list(0, 4) = "Narration"
    
    cbox_Month.Clear
    cbox_Month.ColumnCount = 1
    cbox_Month.ColumnWidths = "100"
    
    cbox_Cat.Clear
    cbox_Cat.ColumnCount = 1
    cbox_Cat.ColumnWidths = "100"
    cbox_Cat.AddItem "Gift/Bonus"
    cbox_Cat.AddItem "Deduction"
    
    For I = 1 To 12
        cbox_Month.AddItem MonthName(I)
        
    Next I
    
    lbl_Total_Amnt_Paid.Caption = ""
    
    
    SQL = "SELECT * FROM TBL_Staff_Detail"
    Call Connection_Vindhya_Cash_Flow(Con)
    
    I = 1
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                lbox_Staff_List.AddItem
                lbox_Staff_List.list(I, 0) = .Fields("ID")
                lbox_Staff_List.list(I, 1) = .Fields("Staff Name")
                I = I + 1
                .MoveNext
            Loop
            
        End If
        
    End With
    
    Rec.Close
    Con.Close
    Dim FYear As String
    
     If VBA.Format(VBA.Date, "MM") < 4 Then
        
            FYear = CLng(VBA.Format(VBA.Date, "YYYY")) - 1
            
    ElseIf VBA.Format(VBA.Date, "MM") > 3 Then
            FYear = VBA.Format(VBA.Date, "YYYY")
            
    End If
    
    txt_Year.Text = FYear
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
   Cancel = 1
End Sub
