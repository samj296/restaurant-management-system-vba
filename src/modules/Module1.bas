Attribute VB_Name = "Module1"
Option Explicit

Sub Inbox_Msgol()
    Dim OL As Outlook.Application
    Dim Ns As Outlook.Namespace
    Dim OM As Outlook.Mailitem
    Dim Fol As Outlook.Folder
    Dim I As Object
    Dim L As Byte
    
        L = 0
        Set OL = Outlook.Application
        Set Ns = OL.GetNamespace("MAPI")
        Set Fol = Ns.GetDefaultFolder(olFolderInbox)
        
            For Each I In Fol.Items
                If I.Class = olMail Then
                    Set OM = I
                        If OM.SenderName = "Keseya" Then
                            Debug.Print OM.SenderName, OM.SenderEmailAddress
                        End If
'                    L = L + 1
'                    If L = 101 Then
'                        Exit For
'                    End If
                End If
                
            Next I
    
End Sub
