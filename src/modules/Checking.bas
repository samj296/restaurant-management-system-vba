Attribute VB_Name = "Checking"
Option Explicit
Sub Field_Value_Check()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
        SQL = "Select * From TBL_Vindhya_Menu"
        
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            .MoveFirst
            MsgBox .Fields("Status"), , "THE VINDHYA CAFE"
            
            
        End With
         
        Rec.Close
        Con.Close
        
    
End Sub
Sub Add_Column()
                
      Dim Con As New ADODB.Connection
      Dim Rec As Recordset
      Dim SQL As String
      Dim Cmnd As New ADODB.Command
      
      SQL = "Alter Table TBL_Variant_Selection ADD COLUMN Test TEXT(225);"
      
      Call Connection_Vindhya_Main_File(Con)
      
      
      
      Cmnd.CommandText = SQL
      Cmnd.ActiveConnection = Con
      Set Rec = Cmnd.Execute
      
      
      
      Con.Close
      
  
    

End Sub

Sub Remove_column()
    Dim Con As New ADODB.Connection
    Dim Rec As New Recordset
    Dim SQL As String
    Dim Cmnd As New ADODB.Command

    SQL = "ALTER TABLE TBL_Variant_Selection DROP COLUMN Test;"
    
    Call Connection_Vindhya_Main_File(Con)
    
    Cmnd.CommandText = SQL
    Cmnd.ActiveConnection = Con
    Set Rec = Cmnd.Execute
    
    Con.Close
    
End Sub

Sub GST()

    Dim SQL As String
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim DiscType As String
    Dim GSTStat As String
    Dim GSTPercent As Byte
    
            SQL = "SELECT TBL_Billing_Type.[Query], TBL_Billing_Type.[UserInput] FROM TBL_Billing_Type"

            DiscType = shMain.Range("D4").Value
            
            MsgBox SQL
            MsgBox DiscType
        Call Connection_Vindhya_Main_File(Con)
                
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            
            With Rec
               If .EOF = False And .BOF = False Then
                    
                .MoveFirst
                
                Do Until .EOF
                    If .Fields("UserInput") = DiscType Then
                        DiscType = .Fields("Query")
                    End If
                    
                    If .Fields("Query") = "GST" And .Fields("UserInput") = "Y" Then
                        GSTStat = "Yes"
                    End If
                    If .Fields("Query") = "GST %" Then
                        GSTPercent = CByte(.Fields("UserInput"))
                    End If
                    
                    .MoveNext
                Loop
                
                Else
                    MsgBox "No Record found", , "THE VINDHYA CAFE"
                    
                End If
                
            End With
                    MsgBox "Discount Type-" & DiscType & vbNewLine & "GST-" & GSTStat & vbNewLine & "GST Percent-" & GSTPercent
                    
        Rec.Close
        Con.Close
End Sub

Sub Data_Expo()
    
    Dim WB As Workbook
    Dim Sh As Worksheet
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim RecTab As ListObject
    Dim LRange As Range
    
    Set WB = Workbooks("Alarasi expense 2020-21.xlsx")
    Set Sh = WB.Worksheets("Paid")
    
    SQL = "Select * From TBL_Expense"
    
    Call Connection_Vindhya_Cash_Flow(Con)
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    Set RecTab = Sh.ListObjects("Table3")
    With Rec
        For Each LRange In RecTab.ListColumns("Date").DataBodyRange
                        .AddNew
                    If LRange.Value <> "" Then
                        
                        .Fields("ExpDate") = LRange.Value
                    End If
                    
                    If LRange.Offset(, 1).Value <> "" Then
                        .Fields("Particular") = LRange.Offset(, 1).Value
                    End If
                    
                    If LRange.Offset(, 3).Value <> "" Then
                        .Fields("Payment Mode") = LRange.Offset(, 3).Value
                    End If
                    
                    If LRange.Offset(, 2).Value <> "" Then
                        .Fields("Ledger Cat") = LRange.Offset(, 2).Value
                    End If
                    
                    If LRange.Offset(, 4).Value <> "" Then
                        .Fields("Amount") = LRange.Offset(, 4).Value
                    End If
                    
                    If LRange.Offset(, 5).Value <> "" Then
                        .Fields("Narration") = LRange.Offset(, 5).Value
                    End If
                    .Update
            
        Next LRange
    End With
    
    Rec.Close
    Con.Close
    
    Set Sh = Nothing
    Set WB = Nothing
    
    

End Sub

Sub Salary()
   Dim Con As New ADODB.Connection
   Dim Rec As New ADODB.Recordset
   Dim SQL As String
   Dim list As New Collection
   Dim User As cls_Staff_User
   Dim I As Byte
   
   
   SQL = "Select * From TBL_Staff_Detail"
        Call Connection_Vindhya_Cash_Flow(Con)
    
   Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
   
    With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                Set User = New cls_Staff_User
                    User.StaffName = .Fields("Staff Name")
                    User.UserId = .Fields("ID")
                    
                    list.Add User
                .MoveNext
            Loop
        End If
    End With
   Rec.Close
   
   SQL = "Select * From TBL_Staff_Deduction"
   
   Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
   
   With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                For Each User In list
                    If User.UserId = .Fields("Id no") Then
                        For I = 1 To 12
                            User.Deduct = User.Deduct + .Fields(MonthName(I))
                        Next I
                        
                    End If
                Next User
                .MoveNext
            Loop
        End If
   End With
   Rec.Close
   
   SQL = "Select * From TBL_Staff_Gift_Bonus"
   
   Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
   
   With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                For Each User In list
                    For I = 1 To 12
                        User.Gift = User.Gift + .Fields(MonthName(I))
                    Next I
                Next User
              .MoveNext
            Loop
        End If
        
   End With
   Rec.Close
   
   SQL = "Select * From TBL_Staff_Salary_Detail"
   Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
   
   With Rec
        If .EOF = False Or .BOF = False Then
            .MoveFirst
            Do Until .EOF
                For Each User In list
                    For I = 1 To 12
                        User.Salary = User.Salary + .Fields(MonthName(I))
                    Next I
                    
                    .Fields("Final Amnt(Deduction and Gift)") = (User.Salary + User.Gift) - User.Deduct
                Next User
            .MoveNext
            Loop
        End If
   End With
   
   Rec.Close
   Con.Close
End Sub

Sub asdkjb()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
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
        
        
            FPath = shMain.Range("F9").Value
            DPath = shMain.Range("F12").Value
            
            Set Sfile = CreateObject("scripting.filesystemobject")
            
            
            Sfile.copyfolder Source:=FPath, Destination:=DPath
            
   
   
            
    
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

Sub fieldName()
Dim FDate As String
Dim TDate As String


FdateEntry:
        FDate = Excel.Application.InputBox("Please enter the from date, in the format of DD-MM-YY", "THE VINDHYA CAFE", , , , , , 2)
        If IsDate(FDate) = False Then
            MsgBox "Not a valid date.", , "THE VINDHYA CAFE"
                GoTo FdateEntry
        End If
TdateEntry:
        TDate = Excel.Application.InputBox("Pleaseenter the to date, in the format of DD-MM-YY", "THE VINDHYA CAFE", , , , , , 2)
        If IsDate(TDate) = False And TDate <> "" Then
            MsgBox "not a valid date.", , "THE VINDHYA CAFE"
                GoTo TdateEntry
        End If
                MsgBox "From Date =" & FDate & " To date =" & TDate
End Sub
