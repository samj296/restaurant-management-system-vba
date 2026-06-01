VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Menu_Management 
   Caption         =   "THE VINDHYA CAFE"
   ClientHeight    =   13215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18855
   OleObjectBlob   =   "Ufrm_Menu_Management.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Menu_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btn_Back_Click()
    Me.Hide
    Ufrm_DashBoard.Show
End Sub

Private Sub btn_Delete_Dish_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As ADODB.Recordset
    Dim Cmd As New ADODB.Command
    Dim SQL As String
    Dim I As Long
    Dim dish As String
    Dim UserInput As String
        
        For I = 1 To lbox_Dish.ListCount - 1
            If lbox_Dish.Selected(I) = True Then
                dish = lbox_Dish.list(I, 1)
            End If
        Next I
             
            If dish <> "" Then
                UserInput = VBA.MsgBox("Are you Sure you want to delete the Dish and its Variant?", vbYesNo, "THE VINDHYA CAFE")
            Else
                MsgBox "Select the Proper dish to Delete.", , "THE VINDHYA CAFE"
                Exit Sub
            End If
             
             If UserInput = vbNo Then
                Exit Sub
             End If
             
            SQL = "Alter Table TBL_Variant_Selection Drop " & dish
            Call Connection_Vindhya_Main_File(Con)
            
            Cmd.CommandText = SQL
            Cmd.ActiveConnection = Con
            Set Rec = Cmd.Execute
            Con.Close
            
            Set Rec = New ADODB.Recordset
            Call Connection_Vindhya_Main_File(Con)
            SQL = "Select * From TBL_Vindhya_Dish_Variant Where TBL_Vindhya_Dish_Variant.[Dish Cat]='" & dish & "'"
            
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    .Delete
                    .Update
                End If
            End With
            
            Rec.Close
            SQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Cat]='" & dish & "'"
            Rec.Open
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    Do Until .EOF
                        .Delete
                        .Update
                        .MoveNext
                    Loop
                End If
            End With
            
            
            
            Rec.Close
            
            SQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Cat]='" & dish & "'"
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveFirst
                        Do Until .EOF
                            .Delete
                            .MoveNext
                        Loop
                    End If
                End With
                
            Rec.Close
            Con.Close
            
            Call UserForm_Initialize
        
    
End Sub

Private Sub btn_Delete_Variant_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim dish As String
    Dim DishVar As String
    Dim I As Long
        
        For I = 1 To lbox_Dish.ListCount - 1
            If lbox_Dish.Selected(I) = True Then
                dish = lbox_Dish.list(I, 1)
            End If
            
        Next I
        
        For I = 1 To Lbox_Variant.ListCount - 1
            If Lbox_Variant.Selected(I) = True Then
                DishVar = Lbox_Variant.list(I, 1)
            End If
            
        Next I
        
        
    
    SQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Cat]='" & dish & "' AND TBL_Vindhya_Menu.[Variant]='" & DishVar & "'"
    
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveLast
                    .Delete
                    .Update
                    
                End If
                
            End With
            
        Rec.Close
    SQL = "Select * From TBL_Variant_Selection Where TBL_Variant_Selection.[" & dish & "]='" & DishVar & "'"
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If .EOF = False Or .BOF = False Then
                .Fields(dish) = ""
                .Update
                
            End If
        End With
        
    
        Rec.Close
        Con.Close
    
        Call lbox_Dish_Change
End Sub

Private Sub BTN_Dish_Add_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As ADODB.Recordset
    Dim Cmd As New ADODB.Command
    Dim SQL As String
    Dim DishNo As Long
    
        If txt_Dish_Add <> "" Then
            SQL = "Alter Table TBL_Variant_Selection ADD COLUMN " & txt_Dish_Add.Text & " TEXT(225);"
            Call Connection_Vindhya_Main_File(Con)
            
            Cmd.CommandText = SQL
            Cmd.ActiveConnection = Con
            Set Rec = Cmd.Execute
            Con.Close
            Set Rec = New ADODB.Recordset
            Call Connection_Vindhya_Main_File(Con)
            SQL = "Select * From TBL_Vindhya_Dish_Variant ORDER BY TBL_Vindhya_Dish_Variant.[Dish No] ASC"
                Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    If .EOF = False Or .BOF = False Then
                        .MoveLast
                        DishNo = .Fields("Dish No") + 1
                        .AddNew
                        .Fields("Dish No") = DishNo
                        .Fields("Dish Cat") = txt_Dish_Add.Text
                        .Fields("Variant Type") = txt_Dish_Add.Text
                        .Update
                        txt_Dish_Add.Text = ""
                    End If
                End With
            
            Rec.Close
            Con.Close
        End If
    Call UserForm_Initialize
End Sub

Private Sub btn_Export_Dish_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim C As Byte
    Dim R As Long
    
        
        
        SQL = "Select * From TBL_Vindhya_Menu Order By TBL_Vindhya_Menu.[Dish Code] ASC"
        
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        R = 2
        With Rec
            'For Header
            If .EOF = False Or .BOF = False Then
            
                For C = 1 To 7
                    shPrint.Cells(1, C).Value = .Fields(C - 1).Name
                Next C
            
            End If
            .MoveFirst
            Do Until .EOF
                shPrint.Cells(R, 1).Value = .Fields(0)
                shPrint.Cells(R, 2).Value = .Fields(1)
                shPrint.Cells(R, 3).Value = .Fields(2)
                shPrint.Cells(R, 4).Value = .Fields(3)
                shPrint.Cells(R, 5).Value = .Fields(4)
                shPrint.Cells(R, 6).Value = .Fields(5)
                shPrint.Cells(R, 7).Value = .Fields(6)
                R = R + 1
                .MoveNext
            Loop
        End With
        
            shPrint.Cells(1, 1).CurrentRegion.Columns.AutoFit
            shPrint.Cells(1, 1).CurrentRegion.PrintOut ActivePrinter:="Microsoft Print to PDF"
            shPrint.Cells(1, 1).CurrentRegion.ColumnWidth = 8.43
            shPrint.Cells(1, 1).CurrentRegion.Delete
            
  Rec.Close
  Con.Close
End Sub

Private Sub btn_Update_Amount_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim DishVar As String
    Dim VarStat As Boolean
    Dim dish As String
    Dim DishStat As Boolean
    Dim I As Long
    Dim Lbox As Long
    Dim DishCode As Long
    Dim NewVar As String
    
        DishStat = False
        VarStat = False
            
           
        
        Select Case cbox_Amnt_Type.Text
            Case Is = "Online", "Take Away", "DineIn"
                If IsNumeric(txt_Amount.Text) = False Then
                    MsgBox "Please enter the amount to update the list.", , "THE VINDHYA CAFE"
                    txt_Amount.Text = ""
                    Exit Sub
                End If
            
        End Select
        
        For I = 1 To lbox_Dish.ListCount - 1
            If lbox_Dish.Selected(I) = True Then
                dish = lbox_Dish.list(I, 1)
'
                DishStat = True
                
            End If
        Next I
        
        For I = 1 To Lbox_Variant.ListCount - 1
            If Lbox_Variant.Selected(I) = True Then
                DishVar = Lbox_Variant.list(I, 1)
                DishCode = Lbox_Variant.list(I, 0)
                VarStat = True
                Lbox = I
            End If
        Next I
        
        If DishStat = False Or VarStat = False Then
            MsgBox "Please Selct the proper Variant to update", , "THE VINDHYA CAFE"
            Exit Sub
        End If
        
        SQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Code]=" & DishCode & " AND TBL_Vindhya_Menu.[Dish Cat]='" & dish & "' AND TBL_Vindhya_Menu.[Variant]='" & DishVar & "'"
        
        Call Connection_Vindhya_Main_File(Con)
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveFirst
                    If cbox_Amnt_Type.Text = "Online" Or cbox_Amnt_Type.Text = "Take Away" Or _
                    cbox_Amnt_Type.Text = "DineIn" Or cbox_Amnt_Type.Text = "Print Name" Or cbox_Amnt_Type.Text = "Variant" Then
                        .Fields(cbox_Amnt_Type.Text) = txt_Amount.Text
                        .Update
                        
                    End If
                End If
            End With
        
        
        If cbox_Amnt_Type.Text = "Variant" Then
            Rec.Close
            SQL = "Select * From TBL_Variant_Selection WHERE TBL_Variant_Selection.[" & dish & "]='" & DishVar & "'"
            NewVar = txt_Amount.Text
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                .Fields(dish) = NewVar
                .Update
            End With
            
        End If
        
        Rec.Close
        Con.Close
        txt_Amount.Text = ""
        cbox_Amnt_Type.Text = ""
        Call lbox_Dish_Change
        Lbox_Variant.Selected(Lbox) = True
End Sub

Private Sub btn_Variant_Add_Click()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim DishCode As Long
    Dim dish As String
    Dim I As Long
            dish = ""
            For I = 1 To lbox_Dish.ListCount - 1
                If lbox_Dish.Selected(I) = True Then
                    dish = lbox_Dish.list(I, 1)
                End If
            Next I
            
            If dish = "" Then
                MsgBox "Please select the proper Dish from the list.", , "THE VINDHYA CAFE"
                Exit Sub
            End If
            
        
        
        If txt_Variant_Add.Text <> "" Then
            SQL = "Select * From TBL_Vindhya_Menu Order By TBL_Vindhya_Menu.[Dish Code] ASC"
            Call Connection_Vindhya_Main_File(Con)
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
            
            With Rec
                If .EOF = False Or .BOF = False Then
                    .MoveLast
                    DishCode = CLng(.Fields("Dish Code")) + 1
                    .AddNew
                    .Fields("Dish Code") = DishCode
                    .Fields("Dish Cat") = dish
                    .Fields("Variant") = txt_Variant_Add.Text
                    .Fields("Online") = 0
                    .Fields("Take Away") = 0
                    .Fields("DineIn") = 0
                    .Fields("Print Name") = txt_Variant_Add.Text
                    .Fields("Status") = True
                    .Update
                End If
            End With
            Rec.Close
            SQL = "Select * From TBL_Variant_Selection"
            Dim UpStat As Boolean
                UpStat = False
            Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
                With Rec
                    .MoveFirst
                    Do Until .EOF
                        If IsNull(.Fields(dish)) = True Then
                            .Fields(dish) = txt_Variant_Add.Text
                            txt_Variant_Add.Text = ""
                            .Update
                            UpStat = True
                            Exit Do
                        End If
                        
                       
                        .MoveNext
                    Loop
                     If UpStat = False Then
                            .AddNew
                            .Fields(dish) = txt_Variant_Add.Text
                            txt_Variant_Add.Text = ""
                            .Update
                            UpStat = True
                        End If
                End With
            
Proce:
            Rec.Close
            Con.Close
            Call lbox_Dish_Change
            
        End If
End Sub

Private Sub lbox_Dish_Change()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim dish As String
    
        For I = 1 To lbox_Dish.ListCount - 1
            If lbox_Dish.Selected(I) = True Then
                dish = lbox_Dish.list(I, 1)
            End If
        Next I
    
    SQL = "Select * From TBL_Vindhya_Menu Where TBL_Vindhya_Menu.[Dish Cat]='" & dish & "' Order By TBL_Vindhya_Menu.[Dish Code] Asc"
    
    Call Connection_Vindhya_Main_File(Con)
    
    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With Rec
        If .EOF = False Or .BOF = False Then
                Lbox_Variant.Clear
                Lbox_Variant.ColumnCount = 6
                Lbox_Variant.ColumnWidths = "50,120,80,50,50,50"
                Lbox_Variant.Height = 384
                Lbox_Variant.Width = 520
                Lbox_Variant.Left = 251
                Lbox_Variant.Top = 36
                Lbox_Variant.AddItem
                Lbox_Variant.list(0, 0) = "Dish Code"
                Lbox_Variant.list(0, 1) = "Variant"
                Lbox_Variant.list(0, 2) = "Print Name"
                Lbox_Variant.list(0, 3) = "Online"
                Lbox_Variant.list(0, 4) = "Take Away"
                Lbox_Variant.list(0, 5) = "Dine In"
                .MoveFirst
                I = 1
            Do Until .EOF
                Lbox_Variant.AddItem
                Lbox_Variant.list(I, 0) = .Fields("Dish Code")
                Lbox_Variant.list(I, 1) = .Fields("Variant")
                Lbox_Variant.list(I, 2) = .Fields("Print Name")
                Lbox_Variant.list(I, 3) = .Fields("Online")
                Lbox_Variant.list(I, 4) = .Fields("Take Away")
                Lbox_Variant.list(I, 5) = .Fields("DineIn")
                
                I = I + 1
                .MoveNext
                
            Loop
        Else
                Lbox_Variant.Clear
                Lbox_Variant.ColumnCount = 6
                Lbox_Variant.ColumnWidths = "50,120,80,50,50,50"
                Lbox_Variant.Height = 384
                Lbox_Variant.Width = 520
                Lbox_Variant.Left = 251
                Lbox_Variant.Top = 36
                Lbox_Variant.AddItem
                Lbox_Variant.list(0, 0) = "Dish Code"
                Lbox_Variant.list(0, 1) = "Variant"
                Lbox_Variant.list(0, 2) = "Print Name"
                Lbox_Variant.list(0, 3) = "Online"
                Lbox_Variant.list(0, 4) = "Take Away"
                Lbox_Variant.list(0, 5) = "Dine In"
        End If
    End With
        
    Rec.Close
    Con.Close
End Sub

Private Sub UserForm_Initialize()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Long
    Dim dish As String
    
        lbox_Dish.Clear
        lbox_Dish.ColumnCount = 2
        lbox_Dish.ColumnWidths = "50,120"
        lbox_Dish.AddItem
        lbox_Dish.list(0, 0) = "S. NO."
        lbox_Dish.list(0, 1) = "Dish"
        
        Lbox_Variant.Clear
        Lbox_Variant.ColumnCount = 6
        Lbox_Variant.ColumnWidths = "50,120,80,50,50,50"
        Lbox_Variant.AddItem
        Lbox_Variant.list(0, 0) = "Dish Code"
        Lbox_Variant.list(0, 1) = "Variant"
        Lbox_Variant.list(0, 2) = "Print Name"
        Lbox_Variant.list(0, 3) = "Online"
        Lbox_Variant.list(0, 4) = "Take Away"
        Lbox_Variant.list(0, 5) = "Dine In"
        
        
    Call Connection_Vindhya_Main_File(Con)
    
    SQL = "Select * From TBL_Vindhya_Dish_Variant ORDER BY TBL_Vindhya_Dish_Variant.[Dish No] ASC"
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        I = 1
        With Rec
            If .EOF = False Or .BOF = False Then
                .MoveFirst
                Do Until .EOF
                    lbox_Dish.AddItem
                        lbox_Dish.list(I, 0) = .Fields("Dish No")
                        lbox_Dish.list(I, 1) = .Fields("Dish Cat")
                        I = I + 1
                    .MoveNext
                Loop
                
            End If
            
        End With
        
        Rec.Close
    
    cbox_Amnt_Type.Clear
    cbox_Amnt_Type.ColumnCount = 1
    cbox_Amnt_Type.ColumnWidths = "100"
    cbox_Amnt_Type.AddItem "Online"
    cbox_Amnt_Type.AddItem "Take Away"
    cbox_Amnt_Type.AddItem "DineIn"
    cbox_Amnt_Type.AddItem "Print Name"
    cbox_Amnt_Type.AddItem "Variant"
    
    
    
'    Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
'
'    With Rec
'        If .EOF = False Or .BOF = False Then
'            .MoveFirst
'            For i = 0 To .Fields.Count - 1
'                Dish = .Fields(i).Name
'             lbox_Dish.AddItem
'             lbox_Dish.list(i, 0) = Dish
'            Next i
'        End If
'    End With
'
'    Rec.Close
    Con.Close
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MsgBox "Press Back Button to show the Dash Board.", vbInformation, "THE VINDHYA CAFE"
    Cancel = 1
End Sub
