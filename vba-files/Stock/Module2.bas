Attribute VB_Name = "Module2"
Sub GotoSheet()
    Dim Houses As Workbook
    Dim xRet As Variant
    Dim xSht As Worksheet
    '
    
    'Application.ScreenUpdating = False
    'Set Houses = Workbooks.Open("S:\Co-operate Affairs\Population Registry\Birth Certificate\old_scans\Books\Houses.xlsm")
    'Houses.Activate
    xRet = Application.InputBox("Go to this sheet", "Kudadobe")
    On Error Resume Next
    If xRet = False Then Exit Sub
    On Error GoTo 0
    On Error Resume Next
    Set xSht = Sheets(xRet)
    If xSht Is Nothing Then Set xSht = Sheets(Val(xRet))
    If xSht Is Nothing Then
        MsgBox "This sheet does not exist", , "Kudadonbe"
    Else
        xSht.Activate
    End If
End Sub

