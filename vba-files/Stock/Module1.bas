Attribute VB_Name = "Module1"

Sub createNewItem()
    Dim pCode As Integer
    Dim ItemCode As String
    
    Dim lastItem As Integer
    Dim pName As String
    Dim pageLink As String
    
    Dim SourceSheet As Worksheet
    Dim LastItemSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim ContentSheet As Worksheet
    Dim StockWorkbook As Workbook
    Dim SelectedCell As Range
    
    Set StockWorkbook = ActiveWorkbook
    'MsgBox StockWorkbook.Name
    ' check if StockWorkbook is Stock Workbook or not
    ' if StockWorkbook name is not "stock_update_v6_2022" then exit sub
    If (StockWorkbook.Name <> "stock_update_v6_2022.xlsx") Then
        MsgBox "This is not a Stock Workbook"
        Exit Sub
    End If
    
    Set SelectedCell = Selection
    
    pName = SelectedCell.Value
    pCode = SelectedCell.Offset(0, -1).Value
    pCodeAddress = SelectedCell.Offset(0, -1).address
    pCodeAddress = Replace(pCodeAddress, "$", "")
    balance = SelectedCell.Offset(0, 1).address
    lastItem = pCode - 1
    'MsgBox "Item Code: " & pCode & vbNewLine & _
        "Item Name: " & pName & vbNewLine & _
        "Last Item: " & lastItem

    Set ContentSheet = StockWorkbook.Sheets("Content")
    'MsgBox ContentSheet.Name
    
    Set SourceSheet = StockWorkbook.Sheets("SampleItemSheet")
    'MsgBox SourceSheet.Name
    
    Set LastItemSheet = StockWorkbook.Sheets(lastItem)
    'MsgBox LastItemSheet.Name
    
    'Set NewSheet = StockWorkbook.Sheets.Add(After:=LastItemSheet)
    SourceSheet.Copy After:=LastItemSheet
    Set NewSheet = ActiveSheet
    NewSheet.Name = pCode
    NewSheet.Range("B5").Value = pName
    ContentSheet.Range(balance).Formula = "=" & pCode & "!H5"
    pageLink = "'" & pCode & "'!A1"
    ItemCode = CStr(pCode)
    ContentSheet.Activate
    ContentSheet.Hyperlinks.Add Anchor:=Range(pCodeAddress), address:="", SubAddress:=pageLink, textToDisplay:=ItemCode
    
End Sub

Sub stokIn()

Dim pCode As Integer
Dim recDate As String
Dim poNo As String
Dim recUser As String
Dim recQty As Integer
Dim pr As Workbook
Dim upDatingRow As Integer
Dim SelectedCell As Range


Set pr = ActiveWorkbook
'check if pr is Purchasing Requisition or not
'if first two letters of pr is not "PR" then exit sub
If (Left(pr.Name, 2) <> "PR") Then
    MsgBox "This is not a Purchasing Requisition"
    Exit Sub
End If

upDatingRow = ActiveCell.Row

Set SelectedCell = Selection

pCode = SelectedCell.Value
recQty = SelectedCell.Offset(0, 1).Value
poNo = SelectedCell.Offset(0, 2).Value
recDate = SelectedCell.Offset(0, 3).Value

recUser = SelectedCell.Offset(0, 4).Value


Dim stockBook As Workbook
'Application.ScreenUpdating = False
Set stockBook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
Set stockSheet = stockBook.Worksheets(pCode)
stockSheet.Activate

Dim lastRow As Long
lastRow = Range("C" & Rows.Count).End(xlUp).Row
newRow = lastRow + 1

stockSheet.Range("A" & newRow).Value = recDate
stockSheet.Range("C" & newRow).Value = poNo
stockSheet.Range("D" & newRow).Value = recUser
stockSheet.Range("F" & newRow).Value = recQty
'stockSheet.Range("F" & newRow).Select
'stockSheet.Range("F" & newRow).Activate



stockBook.Save
'stockBook.Close


pr.Activate
Set prSheet = pr.Worksheets("Sheet1")
SelectedCell.Offset(0, 5).Value = ChrW(10003)
pr.Save

End Sub

Sub stockOut()

    Dim pCode As Integer
    Dim reDate As String
    Dim reqNo As String
    Dim reqUser As String
    Dim reqQty As Integer
    Dim setReqUser As Boolean

    Dim grf As Workbook
    Dim upDatingRow As Integer

    Set grf = ActiveWorkbook

'check if grf is Goods Requisition Form or not
'if first two letters of grf is not "GRF" then exit sub
    'GRF-2023-001
    If (Left(grf.Name, 3) <> "GRF") Then
        MsgBox "This is not a Goods Requisition Form"
        Exit Sub
    End If

    upDatingRow = ActiveCell.Row

    reDate = Range("A33").Value
    reqNo = Range("A5").Text
    reqUser = Range("E33").Text
    'MsgBox reqUser

    If (reqUser = "") Then
        While (setReqUser = False)
            setReqUser = MsgBox("Have you selected Requested Person?", vbYesNo) = vbYes
        WEnd
        reqUser = ActiveCell.Text
    End If

    'Exit Sub

    pCode = Range("H" & upDatingRow).Value
    reqQty = Range("I" & upDatingRow).Value



    Dim stockBook As Workbook
    'Application.ScreenUpdating = False
    Set stockBook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
    'MsgBox pCode
    Set stockSheet = stockBook.Worksheets(pCode)

    stockSheet.Activate

    Dim lastRow As Long
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
    newRow = lastRow + 1
    stockSheet.Range("A" & newRow).Value = reDate
    stockSheet.Range("C" & newRow).Value = reqNo
    stockSheet.Range("D" & newRow).Value = reqUser
    stockSheet.Range("G" & newRow).Value = reqQty
    'stockSheet.Range("G" & newRow).Select
    'stockSheet.Range("G" & newRow).Activate

    'MsgBox "LastRow " & lastRow
    stockBook.Save
    'stockBook.Close

    grf.Activate
    Set grfSheet = grf.Worksheets("Goods Requisition")
    grfSheet.Range("L" & upDatingRow).Value = ChrW(10003)
    grf.Save



End Sub
    ' make a function to get next PR number
        



