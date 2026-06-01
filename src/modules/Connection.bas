Attribute VB_Name = "Connection"
Option Explicit

Sub Connection_Vindhya_Main_File(Main_Con As ADODB.Connection)
    

    'File name is  "\Vindhya_Main_File.accdb"
        Set Main_Con = New ADODB.Connection
        
        Dim FPath As String
    
        
        On Error GoTo ErrorHandler
        
Reprocess:

        If shMain.Range("F9").Value = Null Or shMain.Range("F9").Value = "" Then

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
                        shMain.Range("F9").Value = FPath
                        GoTo Recheck
                    End If
                End With
            
            
        ElseIf shMain.Range("F9").Value <> "" Then
            
            Dim Fcheck As String 'To check the the file path is correct
Recheck:
            FPath = shMain.Range("F9").Value & "\Vindhya_Main_File.accdb"
            
            Fcheck = Dir(FPath)
            
            If Fcheck = VBA.Constants.vbNullString Then
                shMain.Range("F9").Value = ""
                
                GoTo Reprocess
            End If
            
    
        End If
        
        
    Main_Con.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & FPath & ";JET OLEDB:DataBase Password =Sam@Vindhya1902"
            
    
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



Sub Connection_Vindhya_Cash_Flow(Main_Con As ADODB.Connection)

    'File name is  "\Vindhya_Cash_Flow.accdb"
    
        Set Main_Con = New ADODB.Connection
        Dim FPath As String
    
        
        On Error GoTo ErrorHandler
        
Reprocess:

        If shMain.Range("F9").Value = Null Or shMain.Range("F9").Value = "" Then

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
                        shMain.Range("F9").Value = FPath
                        GoTo Recheck
                    End If
                End With
            
            
        ElseIf shMain.Range("F9").Value <> "" Then
            
            Dim Fcheck As String 'To check the the file path is correct
Recheck:
            FPath = shMain.Range("F9").Value & "\Vindhya_Cash_Flow.accdb"
            
            Fcheck = Dir(FPath)
            
            If Fcheck = VBA.Constants.vbNullString Then
                shMain.Range("F9").Value = ""
                
                GoTo Reprocess
            End If
            
    
        End If
        
        
    Main_Con.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & FPath & ";JET OLEDB:DataBase Password =Sam@Vindhya1902"
            
    
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


Sub Connection_Vindhya_Billing_Detail(Main_Con As ADODB.Connection)
    
    'File name is /Vindhya_Billing_Detail.accdb
    
        Set Main_Con = New ADODB.Connection
        Dim FPath As String
    
        
        On Error GoTo ErrorHandler
        
Reprocess:

        If shMain.Range("F9").Value = Null Or shMain.Range("F9").Value = "" Then

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
                        shMain.Range("F9").Value = FPath
                        GoTo Recheck
                    End If
                End With
            
            
        ElseIf shMain.Range("F9").Value <> "" Then
            
            Dim Fcheck As String 'To check the the file path is correct
Recheck:
            FPath = shMain.Range("F9").Value & "\Vindhya_Billing_Detail.accdb"
            
            Fcheck = Dir(FPath)
            
            If Fcheck = VBA.Constants.vbNullString Then
                shMain.Range("F9").Value = ""
                
                GoTo Reprocess
            End If
            
    
        End If
        
        
    Main_Con.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & FPath & ";JET OLEDB:DataBase Password =Sam@Vindhya1902"
            
    
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
