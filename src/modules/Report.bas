Attribute VB_Name = "Report"
Option Explicit

Sub Daily_Report()
    
    Dim Data As New cls_connection
    Dim SQL As String
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim cash As Double
    Dim Card As Double
    Dim UPI As Double
    Dim PayTm As Double
    Dim Zomato As Double
    Dim Swiggy As Double
    Dim WB As Workbook
    Dim RowCount As Long
    Dim shRe As Worksheet
    Dim T As ListObject
    Dim I As Long
    Dim FirstRow As Long
    Dim FPath As String
    
         
    
        
        On Error GoTo ErrorHandler
        
Reprocess:

        If shMain.Range("F11").Value = Null Or shMain.Range("F11").Value = "" Then

            MsgBox "File Location has changed, please select the folder where the report is Saved", _
                vbInformation, "THE VINDHYA CAFE"

                
                With Application.FileDialog(msoFileDialogFolderPicker)
                    .Title = "THE VINDHYA CAFE"
                    .ButtonName = "Pick the folder"
                    If .Show = 0 Then
                        MsgBox "Unable to proceed, as no folder is selected"
                        GoTo Reprocess
                    Else
                        
                        FPath = .SelectedItems(1)
                        shMain.Range("F11").Value = FPath
                        GoTo Recheck
                    End If
                End With
            
            
        ElseIf shMain.Range("F11").Value <> "" Then
            
            Dim Fcheck As String 'To check the the file path is correct
Recheck:
            FPath = shMain.Range("F11").Value & "\Daily_Report.xlsx"
            
            Fcheck = Dir(FPath)
            
            If Fcheck = VBA.Constants.vbNullString Then
                shMain.Range("F11").Value = ""
                
                GoTo Reprocess
            End If
            
    
        End If
        
        
  
            
    
  

    
        
        Workbooks.Open (FPath)
            Set WB = Workbooks("Daily_Report.xlsx")
        
        WB.Application.Visible = False
            Set shRe = WB.Worksheets("Report")
                
            Set T = shRe.ListObjects("Report_TAmnt")
            RowCount = T.ListRows.Count
            FirstRow = 1
            
            
            
    cash = 0
    Card = 0
    UPI = 0
    PayTm = 0
    Zomato = 0
    Swiggy = 0
    
        'for access query the date is in MM/DD/YYYY format
    
        SQL = "SELECT * FROM TBL_Bill_Total WHERE TBL_Bill_Total.[Billing Date] like #" & VBA.Format(DateValue(Now), "mm/dd/yyyy") & "#"
    

        
    Data.Vindhya_Billing_Detail DataCon, DataRec, SQL
        
    With DataRec
        If .BOF = True And .EOF = True Then
            
            shMain.Range("B4:B9").Value = 0
            
             For I = RowCount To FirstRow Step -1
                T.ListRows(I).Delete
            Next I
            
            GoTo FinalStep
        Else
        RowCount = shRe.Range("A" & Rows.Count).End(xlUp).Row
            
            .MoveFirst
                
     
            Do Until .EOF
                If IsNull(.Fields("Status")) = True Then
                        shRe.Range("A" & RowCount).Value = VBA.Format(Now, "DD/MM/YYYY")
                        shRe.Range("B" & RowCount).Value = .Fields("Bill Number")
                        shRe.Range("C" & RowCount).Value = .Fields("Cash")
                        shRe.Range("D" & RowCount).Value = .Fields("PayTm")
                        shRe.Range("E" & RowCount).Value = .Fields("Card")
                        shRe.Range("F" & RowCount).Value = .Fields("UPI")
                        
                        If .Fields("Reference") = "Zomato" Then
                            shRe.Range("H" & RowCount).Value = .Fields("Online")
                        ElseIf .Fields("Reference") = "Swiggy" Then
                            shRe.Range("G" & RowCount).Value = .Fields("Online")
                        ElseIf .Fields("Reference") = "PayTm" Then
                            shRe.Range("I" & RowCount).Value = .Fields("Online")
                        End If
                        
                        
                        RowCount = RowCount + 1
                        
                        'Cash amount
                        If IsNull(.Fields("Cash")) = False Then
                            cash = cash + .Fields("Cash")
                        End If
                        'Paytm amount
                        If IsNull("PayTm") = False Then
                            PayTm = PayTm + .Fields("PayTm")
                        End If
                        'Card amount
                        If IsNull("Card") = False Then
                            Card = Card + .Fields("Card")
                        End If
                        'UPI amount
                        If IsNull(.Fields("UPI")) = False Then
                            UPI = UPI + .Fields("UPI")
                        End If
                        'Swiggy amount
                        If .Fields("Reference") = "Swiggy" Then
                            Swiggy = Swiggy + .Fields("Online")
                        End If
                        'Zomato
                        If .Fields("Reference") = "Zomato" Then
                            Zomato = Zomato + .Fields("Online")
                        End If
                End If
            .MoveNext
            Loop

            shMain.Range("B4").Value = cash
            shMain.Range("B5").Value = PayTm
            shMain.Range("B6").Value = Card
            shMain.Range("B7").Value = UPI
            shMain.Range("B8").Value = Swiggy
            shMain.Range("B9").Value = Zomato
            
        End If
        
        
    End With
    
FinalStep:
    DataRec.Close
    DataCon.Close
    
    WB.Save
    WB.Close
    
      Exit Sub

    'On error
ErrorHandler:
    Select Case Err.Number

        Case Else
            MsgBox "Some error has occured and the error number is -" & Err.Number _
            & ", and the error description is -" & Err.Description, , "THE VINDHYA CAFE"
          
    
    End Select

End Sub
