VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ufrm_Billing_History 
   Caption         =   "THE VINDHYA CAFE(Billing History)"
   ClientHeight    =   13320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "Ufrm_Billing_History.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ufrm_Billing_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txt_Date_From_AfterUpdate()
    Dim Con As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim SQL As String
    Dim I As Double
    Dim FDate As String
    Dim TDate As String
FdateEntry:
        FDate = txt_Date_From.Text
        If IsDate(FDate) = False And FDate <> "" Then
            MsgBox "Not a valid date.", , "THE VINDHYA CAFE"
                GoTo FdateEntry
        End If
TdateEntry:
        TDate = txt_Date_to.Text
        If IsDate(TDate) = False And TDate <> "" Then
            MsgBox "Not a valid date.", , "THE VINDHYA CAFE"
                GoTo TdateEntry
        End If
            If FDate <> "" And TDate <> "" Then
                SQL = "Select * From TBL_Bill_Total Where TBL_Bill_Total.[Billing Date] Between #" & FDate & "# And #" & TDate & "#"
            ElseIf FDate <> "" And TDate = "" Then
                SQL = "Select * From TBL_Bill_Total Where TBL_Bill_Total.[Billing Date]=#" & FDate & "#"
            ElseIf FDate = "" And TDate = "" Then
                Exit Sub
            End If
            
            
            
End Sub

Private Sub UserForm_Initialize()
    lbox_Bill_number.Clear
    lbox_Bill_number.ColumnCount = 3
    lbox_Bill_number.ColumnWidths = "80,50,30"
    lbox_Bill_number.AddItem
    lbox_Bill_number.list(0, 0) = "Bill No."
    lbox_Bill_number.list(0, 1) = "Total Amnt"
    lbox_Bill_number.list(0, 2) = "Status"
    
End Sub
