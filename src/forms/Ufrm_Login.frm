VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Login 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840.001
   OleObjectBlob   =   "Ufrm_Login.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_LogIn_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    
        SQL = "Select * From TBL_User Where strcomp(TBL_User.[User Name],'" & txt_UserID.Text & "',0)= 0" & " And strcomp(TBL_User.[Password],'" & txt_Password.Text & "',0)= 0"
        
        Call Connection_Vindhya_Main_File(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            
            If .BOF = True And .EOF = True Then
                MsgBox "Login failed, invalid credential.", , "THE VINDHYA CAFE"
                
                txt_UserID.Text = ""
                txt_Password.Text = ""
                txt_UserID.SetFocus
                Rec.Close
                Con.Close
                Exit Sub
            ElseIf .BOF = False Or .EOF = False Then
                
                shMain.Range("F6").Value = .Fields("UserID_NO")
                shMain.Range("F7").Value = .Fields("Admin Rights")
                shMain.Visible = xlSheetVeryHidden
                shOther.Visible = xlSheetVeryHidden
                shPrint.Visible = xlSheetVisible
                
                
                
            End If
            
            
        
        End With
        
        Rec.Close
        
        SQL = "Select * From TBL_Login_Detail"
          
          Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
          
        With Rec
            .AddNew
            .Fields("Login Date") = Now
            .Fields("UserID") = shMain.Range("F6").Value
            .Update
        
        End With
        
        
        
        Rec.Close
        Con.Close
       txt_UserID.Text = ""
       txt_Password.Text = ""
       
       Me.Hide
       
        Ufrm_DashBoard.Show
    
    
End Sub

Private Sub btn_LogIn_Enter()
     Call btn_LogIn_Click
End Sub




Private Sub UserForm_Terminate()
    Excel.Application.DisplayAlerts = False
    Excel.Application.Quit
    
End Sub
