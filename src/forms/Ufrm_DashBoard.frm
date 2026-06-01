VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_DashBoard 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   13210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "Ufrm_DashBoard.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_DashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Billing_Click()
    Me.Hide
    If Ufrm_Billing.Visible = False Then
        Ufrm_Billing.Show
    End If
    
End Sub





Private Sub btn_Billing_History_Click()
   
    
End Sub

Private Sub btn_ExpenseFrm_Click()
    If shMain.Range("F7").Value = "Admin" Or shMain.Range("F7").Value = "Moderate" Then
    
        Me.Hide
        Ufrm_Expense.Show
    Else
        
        MsgBox "You are not authorise to go to expense section.", , "THE VINDHYA CAFE"
        
    End If
    
End Sub

Private Sub Btn_IncomeFrm_Click()
    If shMain.Range("F7").Value = "Admin" Or shMain.Range("F7").Value = "Moderate" Then
        
        Me.Hide
        
        Load Ufrm_Income
        Ufrm_Income.Show
        
    Else
        
        MsgBox "You are not authorise to go to income section.", , "THE VINDHYA CAFE"
        
    End If
    
    
End Sub

Private Sub btn_Manage_User_Click()
    If shMain.Range("F7").Value = "Admin" Then
        Me.Hide
        Load Ufrm_User_manage
        Ufrm_User_manage.Show
    Else
        MsgBox "You are not authorise to Manage User", , "THE VINDHYA CAFE"
        Exit Sub
    End If
End Sub



Private Sub btn_Menu_Detail_Click()
    If shMain.Range("F7").Value = "Admin" Then
    
        Me.Hide
        Ufrm_Menu_Management.Show
    Else
        MsgBox "You are not authorise to change the Menu.", , "THE VINDHYA CAFE"
    End If
    
End Sub

Private Sub btn_Salary_Detail_Click()
    If shMain.Range("F7").Value = "Admin" Then
        Me.Hide
        Ufrm_Salary_Detail.Show
    Else
        MsgBox "You are not authorise to view the Salary detail.", , "THE VINDHYA CAFE"
    End If
End Sub

Private Sub btn_Show_WorkBook_Click()
     If shMain.Range("F7").Value = "Admin" Then
            Call Excel_Mode
            
            If Excel.Application.Visible = True Then
                Me.Hide
            End If
    Else
        MsgBox "You are not authorise to open the Excel File.", , "THE VINDHYA CAFE"
    End If
End Sub

Private Sub btn_SwitchUser_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    
        SQL = "Select * From TBL_Login_Detail WHERE TBL_Login_Detail.[UserID]=" & shMain.Range("F6").Value & " AND TBL_Login_Detail.[LogOut Date] Is NULL"
    
    Call Connection_Vindhya_Main_File(Con)
                    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
        With Rec
            If .EOF = False Or .BOF = False Then
                .Fields("LogOut Date") = Now
                .Update
            
            End If
        
        
        End With
                    
                    shMain.Visible = xlSheetVeryHidden
                    shOther.Visible = xlSheetVeryHidden
                    shPrint.Visible = xlSheetVisible
                    
                    shMain.Range("F6", "F7").Value = ""
                    
                    
                    ThisWorkbook.Save
                    Me.Hide
                    Ufrm_Login.Show
    
End Sub

Private Sub optbtn_OFF_Click()
    If shMain.Range("F7").Value = "Admin" Then
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            
                SQL = "Select * From TBL_Billing_Type Where TBL_Billing_Type.[Query] Like '%GST%'"
                
                Call Connection_Vindhya_Main_File(Con)
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                            If .Fields("Query") = "GST Status" And .Fields("UserInput") = "Y" Then
                                .Fields("UserInput") = "N"
                                .Update
                            End If
                            
                            .MoveNext
                        Loop
                    
                    End If
                    
                End With
                Rec.Close
                Con.Close
     Else
        MsgBox "You are not authorise to change the GST status.", , "THE VINDHYA CAFE"
        Exit Sub
     End If
End Sub

Private Sub optbtn_ON_Click()
    If shMain.Range("F7").Value = "Admin" Then
            Dim Con As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim SQL As String
            Dim Old As Byte
                SQL = "Select * From TBL_Billing_Type Where TBL_Billing_Type.[Query] Like '%GST%' Order By TBL_Billing_Type.[Query] Desc"
                
                Call Connection_Vindhya_Main_File(Con)
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                            If .Fields("Query") = "GST Status" And .Fields("UserInput") = "N" Then
                                .Fields("UserInput") = "Y"
                                .Update
                                .MoveNext
                                If .Fields("Query") = "GST %" Then
                                    Old = .Fields("UserInput")
                                    .Fields("UserInput") = Excel.Application.InputBox("Please enter the percentage of the GST as per the GOVT Law. Last Percentage updated was " & .Fields("UserInput") & "%", "THE VINDHYA CAFE", , , , , , 1)
                                    .Update
                                End If
                            End If
                            
                            
                            .MoveNext
                        Loop
                    
                    End If
                    
                End With
            
            Rec.Close
            Con.Close
     Else
        MsgBox "You are not authorise to cgange the GST status.", , "THE VINDHYA CAFE"
        Exit Sub
     End If
     
End Sub

Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub



Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call UserForm_Initialize
    
End Sub

Private Sub UserForm_Initialize()
    Dim cash As Double, Card As Double, UPI As Double, PayTm As Double, Online As Double
    Call Daily_Report
    
    
    
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim MonMode As String
    Dim I As Byte
    For I = 1 To 5
        Select Case I
            Case Is = 1
                MonMode = "Cash"
                GoSub AmntPro
                cash = CDbl(MonMode)
            Case Is = 2
                MonMode = "Card"
                GoSub AmntPro
                Card = CDbl(MonMode)
            Case Is = 3
                MonMode = "UPI"
                GoSub AmntPro
                UPI = CDbl(MonMode)
            Case Is = 4
                MonMode = "PayTm"
                GoSub AmntPro
                PayTm = CDbl(MonMode)
            Case Is = 5
                MonMode = "Online"
                GoSub AmntPro
                Online = CDbl(MonMode)
            
            
        End Select
    Next I

    lbl_Cash_Amount.Caption = cash
    lbl_Card_Amount.Caption = Card
    lbl_UPI_Amount.Caption = UPI
    lbl_PayTm_Amount.Caption = PayTm
    lbl_Online_Amount.Caption = Online

        GoTo NextStep
AmntPro:
        
    SQL = "SELECT SUM(TBL_Bill_Total.[" & MonMode & "]) AS Total FROM TBL_Bill_Total WHERE TBL_Bill_Total.[Billing Date] = #" & VBA.Format(VBA.Date, "MM-DD-YY") & "# AND TBL_Bill_Total.[Status] IS NULL or TBL_Bill_Total.[Billing Date]=#" & VBA.Format(VBA.Date, "MM-DD-YY") & "# AND TBL_Bill_Total.[Status]='Completed'"


    Call Connection_Vindhya_Billing_Detail(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic

        With Rec

            If .EOF = False Or .BOF = False Then
                .MoveFirst
                If IsNull(.Fields("Total")) = False Then
                    MonMode = .Fields(0)
                Else
                    MonMode = 0
                End If

            End If

        End With


    Rec.Close
    Con.Close

    Return
    
NextStep:
    Dim Chrt As Chart
    
    Set Chrt = shMain.ChartObjects("Sale_Chart").Chart
    
    Chrt.Export VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chart.jpg", "jpg" 'Filename:=VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chart.jpg", Filtername:="JPG"
    
    img_SaleChart.Picture = LoadPicture(VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chart.jpg") 'VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chart.jpg")
    
    Kill (VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chart.jpg")
    
        If shMain.Range("F7").Value = "Admin" Then
            FRM_GST.Visible = True
        Else
            FRM_GST.Visible = False
            
        End If
    

    
    SQL = "Select * From TBL_Billing_Type Where TBL_Billing_Type.[Query] Like '%GST%' Order By TBL_Billing_Type.[Query] Desc"
    
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                If .Fields("Query") = "GST Status" And .Fields("UserInput") = "Y" Then
                    optbtn_ON = True
                ElseIf .Fields("Query") = "GST Status" And .Fields("UserInput") = "N" Then
                    optbtn_OFF = True
                    
                End If
                
            .MoveNext
            Loop
        End If
        
    End With
    
    Rec.Close
    Con.Close
    
End Sub



Private Sub UserForm_Terminate()
    Call LogOut
End Sub
