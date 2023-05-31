
Public prItemCodeArray(1 to 50) As Integer
Public prItemQtyArray(1 to 50) As Integer
Public nextItemCode As Integer

Public Function nextPRNo(prDate) As String
    ' search all the files in pr folder
    ' get the last PR number
    ' increment the last PR number
    ' return the new PR number
    ' PR-2023-001
    Dim prFolder As String
    Dim prNo As String
    Dim lastPRNo As Integer
    Dim fileName As String

    prFolder = "\\server\sections\Co-operate Affairs\Procurement\PR Form\2023\"
    ' get all file names in pr folder
    fileName = Dir(prFolder & "*.xlsm")
    Do While fileName <> ""

        MsgBox fileName
        prNo = Val(Mid(fileName, 9, Len(fileName) - 12))
        ' Update the last PR number if a larger number is found
        If (prNo > lastPRNo) Then
            lastPRNo = prNo
        End If

        ' Move to next file name
        fileName = Dir

    Loop
    newPRNo = "PR-" & Year(prDate) & "-" & Format(lastPRNo + 1, "000")
    nextPRNo = newPRNo

End Function

Public Function addItemToPr(prSheet, itemCode, itemCodeRow, grfNumber, stockSheet, grfSheet, requestedAmount)
    ' set stock sheet
    Dim lastRowInStock As Integer

    Dim i As Integer
    Dim itemName As String
    Dim stockBalance As Integer
    Dim requestedDate As String
    Dim requestedSection As String
    Dim lastRowInPR As Integer
    Dim newRow As Integer




    lastRowInStock = stockSheet.Range("B" & Rows.Count).End(xlUp).Row

    For i = 1 To lastRowInStock
        If (stockSheet.Range("B" & i).Value = itemCode) Then
            ' if item found then add to PR
            ' get item name
            itemName = stockSheet.Range("C" & i).Value
            MsgBox itemName
            ' get stock balance
            stockBalance = stockSheet.Range("D" & i).Value
            ' get requested amount
            ' get requested date
            requestedDate = grfSheet.Range("A6").Value
            ' get requested section
            requestedSection = grfSheet.Range("A4").Value
            ' get received qty
            lastRowInPR = prSheet.Range("A" & Rows.Count).End(xlUp).Row
            newRow = lastRowInPR + 1
            prSheet.Range("A" & newRow).Value = requestedAmount
            prSheet.Range("B" & newRow).Value = "pcs"
            prSheet.Range("C" & newRow).Value = itemName
            prSheet.Range("E" & newRow).Value = stockBalance
            prSheet.Range("F" & newRow).Value = requestedDate
            prSheet.Range("G" & newRow).Value = requestedSection
            prSheet.Range("H" & newRow).Value = grfNumber
            prSheet.Range("J" & newRow).Value = itemCode

        End If
    Next i


End Function

' sub to generate new PR
Sub newPR()

    Dim grf As Workbook
    Dim grfSheet As Worksheet
    Dim grfNumber As String
    Dim itemCode As Integer

    Dim itemCodeRow As Integer

    Dim pr As Workbook
 

    Dim backDate As Boolean
    Dim prDate As String
    Dim prSheet As Worksheet
    Dim prNo As String


    ' set grf to active workbook
    Set grf = ActiveWorkbook
    ' check if open workbook is GRF or not
    ' if open workbook is not GRF then exit sub

    If (Left(grf.Name, 3) <> "GRF") Then
        MsgBox Left(grf.Name, 3)
        MsgBox "This is not a Goods Requisition Form"
        Exit Sub
    End If

    Set grfSheet = grf.Worksheets("Goods Requisition")
    ' get activeRow
    itemCodeRow = ActiveCell.Row

    ' ask if pr need to be a backdated
    backDate = MsgBox("Do you want to backdate this PR?", vbYesNo) = vbYes

    ' if yes get the date else proceed with current date
    If (backDate) Then
        prDate = InputBox("Enter PR Date (dd/mm/yyyy)")
    Else
        prDate = Date
    End If

    ' get PR template
    Set pr = Workbooks.Open("\\server\sections\Co-operate Affairs\Procurement\PR Form\templates\PR_template.xltm")
    ' get PR template sheet
    Set prSheet = pr.Worksheets("Sheet1")
    ' set date to PR
    prSheet.Range("H7").Value = prDate
    ' format pr number
    'example PR-2023-001

    prNo = nextPRNo(prDate)

    ' MsgBox prNo
    ' set pr number to PR
    prSheet.Range("A7").Value = prNo

    ' get selected item code form grf
    ' itemCode = grfSheet.Range("H" & itemCodeRow).Value
    grfNumber = grfSheet.Range("A5").Value


    Dim stockBook As Workbook
    Dim stockSheet As Worksheet
    Set stockBook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
    Set stockSheet = stockBook.Worksheets("Content")

    ' get itemcode and store in prItemCodeArray

    ' loop through prItemCodeArray
    ' run addItemToPr for each item code
    Dim numberItesmNeedToBeAdded As Integer
    Dim moreItems As Boolean
    moreItems = True
    numberItesmNeedToBeAdded = nextItemCode
    prSheet.Activate
    For i = 1 To numberItesmNeedToBeAdded
        itemCode = prItemCodeArray(i)
        requestedAmount = prItemQtyArray(i)
        addItemToPr prSheet, itemCode, itemCodeRow, grfNumber, stockSheet, grfSheet, requestedAmount

    Next i

    ' save PR with PR number
    ' pr.SaveAs "\\server\sections\Co-operate Affairs\Procurement\PR Form\2023\" & prNo & ".xlsm"

End Sub

' need a seperate Sub to add item to prItemCodeArray
Public Sub addItemToPrItemCodeArray()
    Dim grfSheet As Worksheet
    Set grfSheet = ActiveWorkbook.Worksheets("Goods Requisition")
    nextItemCode = nextItemCode + 1
    If (nextItemCode > 50) Then
        msgBox "You can only add 50 items at a time, Make new PR for more items"
        Exit Sub
    End If
    prItemCodeArray(nextItemCode) = grfSheet.Range("H" & ActiveCell.Row).Value
    prItemQtyArray(nextItemCode) = grfSheet.Range("J" & ActiveCell.Row).Value

    With ActiveCell.Font
        .Color = RGB(0, 100, 0) ' Dark green color (adjust the RGB values as needed)
    End With

End Sub


