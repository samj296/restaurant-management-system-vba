VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Staff_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StaffName As String
Public Password As String
Public Salary As Double
Public UserId As Long
Public Deduct As Double
Public Gift As Double
Public JoinDate As Date
Public ResignDate As String
Public FYear As String
Public JoinDay As Byte
Public JoInMonth As Byte
Public UserType As String


Function SalaryUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim M As Byte
    
   
    
    SQL = "Select * From TBL_Staff_Salary_Detail"
        Call Connection_Vindhya_Cash_Flow(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If UserType = "Old" Then
                    .MoveFirst
                    Do Until .EOF
                        If .Fields("Staff name") = StaffName And .Fields("Salary (Year)") And .Fields("Staff Id") = UserId Then
                            GoTo Proceed
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                Else
                    .AddNew
                End If



                .Fields("Staff name") = StaffName
                .Fields("Staff Id") = UserId
                .Fields("Salary (Year)") = FYear
Proceed:
                For M = 1 To 12
                    Select Case JoInMonth
                        Case Is = 1
                            If VBA.Format(JoinDate, "YYYY") > FYear Then
                                If M = 1 Or M = 2 Or M = 3 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") <= FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 2
                            If VBA.Format(JoinDate, "YYYY") > FYear Then
                                If M = 2 Or M = 3 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") <= FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 3
                            If VBA.Format(JoinDate, "YYYY") > FYear Then
                                If M = 3 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") <= FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 4
                            
                                .Fields(MonthName(M)) = Salary
                            
                        Case Is = 5
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 4 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                        Case Is = 6
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 5 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 7
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 6 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 8
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 7 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 9
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 8 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 10
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 9 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 11
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M > 10 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                        Case Is = 12
                            If VBA.Format(JoinDate, "YYYY") >= FYear Then
                                If M = 1 Or M = 2 Or M = 3 Or M - 12 Then
                                    .Fields(MonthName(M)) = Salary
                                End If
                            ElseIf VBA.Format(JoinDate, "YYYY") < FYear Then
                                .Fields(MonthName(M)) = Salary
                            End If
                            
                    End Select
                    
                Next M
                    .Update
                If UserType = "Old" Then
                    
                    
                    GoTo FinalStep
                    
                End If
            End With
        
        Rec.Close
        
            'updating in Gift/Bonus Table
        SQL = "Select * From TBL_Staff_Gift_Bonus"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                .AddNew
                .Fields("Id no") = UserId
                .Fields("Bonus Year") = FYear
                For M = 1 To 12
                    .Fields(MonthName(M)) = 0
                Next M
                .Update
            End With
        
        Rec.Close
        
            'updating in Deduction Table
        SQL = "Select * From TBL_Staff_Deduction"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                .AddNew
                .Fields("Id no") = UserId
                .Fields("Deduction Year") = FYear
                For M = 1 To 12
                    Select Case M
                        Case Is = JoInMonth
                            If JoinDay > 1 Then
                                Deduct = (Salary / 30) * (JoinDay)
                                .Fields(MonthName(M)) = Deduct
                            End If
                        Case Else
                            .Fields(MonthName(M)) = 0
                    End Select
                Next M
                .Update
            End With
        
        Rec.Close
            'updating in Ledger Table
        SQL = "Select * From TBL_Ledger"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                .AddNew
                .Fields("Ledger Cat") = StaffName
                .Fields("ID NO") = UserId
                .Update
            End With
FinalStep:
        Rec.Close
        Con.Close
End Function

Function ResignProceed()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim M As Byte 'For looping Month
    Dim OldSal As Double
    Dim NewSal As Double
            'joinmonth is using just for declaring the month of the staff left the job
        JoInMonth = CByte(VBA.Format(ResignDate, "MM"))
        JoinDay = CByte(VBA.Format(ResignDate, "DD"))
    
    If VBA.Format(VBA.Date, "MM") < 4 Then
        
            FYear = CLng(VBA.Format(VBA.Date, "YYYY")) - 1
            
    ElseIf VBA.Format(VBA.Date, "MM") > 3 Then
            FYear = VBA.Format(VBA.Date, "YYYY")
            
    End If
            'Updating in Staff Deatil Table
        SQL = "Select * FROM TBL_Staff_Detail WHERE TBL_Staff_Detail.[ID] =" & UserId
        
        Call Connection_Vindhya_Cash_Flow(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
            With Rec
                .MoveFirst
                .Fields("Reisgn On") = ResignDate
                .Update
                
            End With
        Rec.Close
            'updating in Salary Detail Table
        SQL = "Select * From TBL_Staff_Salary_Detail WHERE TBL_Staff_Salary_Detail.[Salary (Year)] = '" & FYear _
        & "' AND TBL_Staff_Salary_Detail.[Staff Id]=" & UserId
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
            With Rec
                .MoveFirst
                For M = 1 To 12
                        'Finding the correct month the staff left
                     Select Case JoInMonth
                         Case Is = 1
                             
                             If FYear = (VBA.Format(ResignDate, "YYYY") - 1) Then
                                 If M = 1 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                
                                 If M = 2 Or M = 3 Then
                                    .Fields(MonthName(M)) = 0
                                 End If
                             End If
                             
                         Case Is = 2
                            
                            If FYear = (VBA.Format(ResignDate, "YYYY") - 1) Then
                                If M = 2 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                
                                If M = 3 Then
                                    .Fields(MonthName(M)) = 0
                                    .Update
                                    Exit For
                                End If
                            End If
                            
                         Case Is = 3
                                If M = 3 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                .Update
                            Exit For
                            
                         Case Is = 4
                         
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 4 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M <> 4 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                            
                         Case Is = 5
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 5 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M <> 4 And M <> 5 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 6
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 6 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 6 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                            
                         Case Is = 7
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 7 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 7 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 8
                         
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 8 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 8 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 9
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                
                                If M = 9 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 9 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 10
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 10 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 10 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 11
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 11 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Or M > 11 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                         Case Is = 12
                            
                            If FYear = VBA.Format(ResignDate, "YYYY") Then
                                If M = 12 Then
                                    OldSal = .Fields(MonthName(M))
                                    NewSal = ((OldSal / 30) * JoinDay)
                                    .Fields(MonthName(M)) = NewSal
                                 End If
                                 
                                If M < 4 Then
                                    .Fields(MonthName(M)) = 0
                                End If
                            End If
                         
                     End Select
                        
                        .Update
                
                Next M
                
            End With
        
            

        Rec.Close
        Con.Close
        
End Function

