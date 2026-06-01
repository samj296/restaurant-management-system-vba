Attribute VB_Name = "ExtraWork"
Option Explicit

Sub LogOut()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim WB As Workbook
    
    SQL = "Select * From TBL_Login_Detail WHERE TBL_Login_Detail.[UserID]=" & shMain.Range("F6").Value & " AND TBL_Login_Detail.[LogOut Date] Is NULL"
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
       If .EOF = False Or .BOF = False Then
           .Fields("LogOut Date") = Now
            .Update
            shMain.Visible = xlSheetVeryHidden
            shOther.Visible = xlSheetVeryHidden
            shPrint.Visible = xlSheetVisible
        End If
        
    End With
    
  
        
            Rec.Close
            Con.Close
    
    
    Dim Sfile As Object
    Dim FPath As String
    Dim DPath As String
    
    
            On Error GoTo ErrorHandler
        
Reprocess:

        If shMain.Range("F12").Value = Null Or shMain.Range("F12").Value = "" Then

            MsgBox "File Location has changed, please select the folder where the database is stored", _
                vbInformation, "THE VINDHYA CAFE"

                
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .Title = "THE VINDHYA CAFE"
                    .ButtonName = "Pick the folder"
                    If .Show = 0 Then
                        MsgBox "Unable to proceed, as no folder is selected"
                        GoTo Reprocess
                    Else
                        
                        FPath = .SelectedItems(1)
                        shMain.Range("F12").Value = FPath
                        GoTo Recheck
                    End If
                End With
            
            
        ElseIf shMain.Range("F12").Value <> "" Then
            
            Dim Fcheck As String 'To check the the file path is correct
Recheck:
            FPath = shMain.Range("F12").Value & "\"
            
            Fcheck = Dir(FPath)
            
            If Fcheck = VBA.Constants.vbNullString Then
                shMain.Range("F12").Value = ""
                
                GoTo Reprocess
            End If
            
    
        End If
            
            ThisWorkbook.Save
        
            FPath = shMain.Range("F9").Value
            DPath = shMain.Range("F12").Value
            
            Set Sfile = CreateObject("scripting.filesystemobject")
            
            
            Sfile.copyfolder Source:=FPath, Destination:=DPath
            
            
            
            Excel.Application.DisplayAlerts = True
            Excel.Application.Quit
   
            
    
        Exit Sub

    'On error
ErrorHandler:
    Select Case Err.Number
        Case Is = 52
    
        
            MsgBox "Please Select the correct folder to process the data or the server where the data is kept is not responding," _
            & "pleas select the correct folder or check the connection to the server", , "THE VINDHYA CAFE"
            shMain.Range("F9").Value = ""
            GoTo Reprocess
        Case Is = -2147467259

            MsgBox "Please Select the correct folder to process the data", , "THE VANDHYA CAFE"
            shMain.Range("F9").Value = ""
            GoTo Reprocess
        Case Is = -2147217843
            MsgBox "DataBase Password has been changed Please set the old password to proceed or contact Admin to change the Password." _
            , , "THE VINDHYA CAFE"
            Exit Sub
        Case Else
            MsgBox "Some error has occured and the error number is -" & Err.Number _
            & ", and the error description is -" & Err.Description, , "THE VINDHYA CAFE"
          
    
    End Select
    
    
    
            
        
    
    
End Sub

Sub APP_Mode()
    
    shMain.Visible = xlSheetVeryHidden
    shOther.Visible = xlSheetVeryHidden
    Excel.Application.Visible = False
    Ufrm_DashBoard.Show
    
    
End Sub

Sub Excel_Mode()
    
    
    If shMain.Range("F7").Value = "Admin" Then
        Excel.Application.Visible = True
        shMain.Visible = xlSheetVisible
        shOther.Visible = xlSheetVisible
        shPrint.Visible = xlSheetVisible
        shMain.Activate
        Excel.Application.DisplayAlerts = True
        Application.WindowState = xlMaximized
        
    Else
        MsgBox "You are not authorise to open view the excel file.", , "THE VINDHYA CAFE"
        Exit Sub
    End If
    
End Sub
