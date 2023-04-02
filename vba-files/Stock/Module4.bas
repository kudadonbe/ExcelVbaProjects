Attribute VB_Name = "Module4"
'birth Certificate
Sub genDocName()

Dim no, nid, doc_name, full_name, full_address As String

Dim address() As String

Dim selectedRow As Integer

Dim BirthRecords As Workbook

Set BirthRecords = ActiveWorkbook

selectedRow = ActiveCell.Row

no = Range("A" & selectedRow).Text
no = Replace(no, "/", "_")
nid = Range("D" & selectedRow).Text

full_name = Range("E" & selectedRow).Text
full_name = Replace(full_name, " ", "_")

full_address = Range("I" & selectedRow).Text

address = Split(full_address, ",")

doc_name = no + "_" + nid + "_" + full_name + "_" + address(0)

Range("B" & selectedRow).Value = doc_name

End Sub

Sub SaveAsPDF()

    Dim BirthRecords As Workbook
    
    Set BirthRecords = ActiveWorkbook
    
    'selectedRow = ActiveCell.Row

    Dim fileName, doc_name As String
    doc_name = Range("P9").Text
    fileName = "S:\Co-operate Affairs\Population Registry\Birth Certificate\Houses\" & doc_name
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End Sub
