Attribute VB_Name = "StockCounts"
Public  Function LoadStockCountData()
    Dim stockCountSheet As Worksheet
    Dim stockSheet As Worksheet
End Function

Public  Sub stockCount()
    ' set workbook & worksheet
    Dim stockBook As Workbook
    Dim stockCountBook As Workbook
    Dim stockCountSheet As Worksheet
    Dim stockSheet As Worksheet
    Dim stockCountData As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim stockCountDataRange As Range
    
    
    Set stockCountBook = ActiveWorkbook
    ' if stockCountBook name is not StockCount then exit sub
    If (stockCountBook.Name <> "StockCount.xlsx") Then
        MsgBox "This is not a Stock Count Workbook"
        Exit Sub
    End If

    Set stockCountSheet = stockCountBook.Worksheets("Data")
    Set stockCountDataRange = stockCountSheet.UsedRange
    
    lastRow = stockCountDataRange.Rows.Count
    lastCol = stockCountDataRange.Columns.Count

    ' stockCountData = stockCountDataRange.Resize(lastRow - 2, lastCol).Offset(2, 0).Value


    Set stockBook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
    ' MsgBox "Counting_OK"
    ' get item number
    Dim itemNo As integer
    Dim itemQty As integer
    Dim itemBalance As integer
    Dim itemDifference As integer
    Dim updatedDate As Date
    Dim updatedBy As String
    Dim comment(1 to 2 As String
    Dim itemName(1 to 2) As String
    Dim itemBrand(1 to 2 As String
    Dim itemModel As String
    Dim itemSerialNo As String
    Dim itemSize As String
    Dim itemWeight As String
    Dim itemStatus(1 to 2 As String
    Dim itemLocation(1 to 2 As String
    Dim itemRecievedDate As Date
    Dim itemPrice As Double
    Dim itemSupplier(1 to 2 As String
    Dim itemOtherInfo(1 to 2 As String

    ' item received date
    ' go to item sheet
    ' enter count
    ' calculate difference
    ' adjust stock balance by adding difference to 
    ' to adjust stock balance add to stockIn col is its a positive number
    ' to adjust stock balance add to stockOut col if its a negative number
    ' add date
    ' add user
    ' add comment



End Sub