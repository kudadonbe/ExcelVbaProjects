'safeReportWeekly
Attribute VB_Name = "safeReportWeekly"

Public Sub wrSafe()
    ' On Error GoTo ErrorHandler
    Application.ScreenUpdating = False ' Disable screen updating
    Application.DisplayAlerts = False ' Disable alerts to prevent confirmation prompts


    Dim SafeData As Workbook
    Dim safeDataSheet As Worksheet
    Dim ActiveReportSheet As Worksheet
    'Dim ws As Worksheet

   
    ' Set selectedRange = Selection ' Get the selected range
    ' selectedRow = selectedRange.Row


    Set SafeData = Workbooks.Open("S:\Co-operate Affairs\Safe\2023\Safe_2023.xlsx")
    
    'Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set safeDataSheet = SafeData.Worksheets("Data")
    Set ActiveReportSheet = SafeData.Worksheets(ActiveSheet.Name)

  
    Dim startDate As Date, endDate As Date


    Dim recDateArr(10 to 10000) As String
    Dim detailsArr(10 to 10000) As String
    Dim recNumArr(10 to 10000) As String
    Dim totalArr(10 to 10000) As String
    Dim isCanceledArr(10 to 10000) As String

    Dim glCodeArr(1 to 10000) As String
    Dim incomeCodeArr(1 to 10000) As String
    Dim activityArr(1 to 10000) As String
    Dim totalSum(1 to 10000) As Double

    
    
    
    
    Dim recDate As String
    Dim details As String
    Dim recNum As String
    Dim isCanceled As String
    Dim isWeekly As Boolean
    

    Dim glCode As String
    Dim incomeCode As String
    Dim activity As String
    Dim total As Double
    
    Dim lastRow As Long
    Dim i As Long
    Dim indexOfFoundIncomeCode As Long
    Dim totalIncome As Double
    Dim newRecord As Long
    Dim nextFifty As Integer
    Dim nextRecord As Integer
    Dim nextNagudhuRecord As Integer
    Dim avlRowsForTransMonthly As Integer
    Dim avlRowsForTransWeekly As Integer
    Dim avlRowsForCatDetails As Integer

    Dim payType As String


    Dim totalRowsOfData As Integer
    Dim insertNewRows As Integer
    Dim mNaguthuStartsAt As Integer
    Dim wNaguthuStartsAt As Integer

    
    

    avlRowsForTransMonthly = 23
    avlRowsForTransWeekly = 7
    avlRowsForCatDetails = 7
    mNaguthuStartsAt = 36
    wNaguthuStartsAt = 20

    payType = "Cash"
    
    
    startDate = ActiveReportSheet.Range("B9")

    If (ActiveReportSheet.Name = "MonthlyReport") Then 
    
        endDate = DateSerial(Year(startDate), Month(startDate) + 1, 0) ' Set end date here
        nextNagudhuRecord = (mNaguthuStartsAt - 1)
        ' MsgBox endDate
    Else
        isWeekly = True
        endDate = DateAdd("d", 6, startDate) ' Set end date here
        nextNagudhuRecord = (wNaguthuStartsAt - 1)
    
    End If

    newRecord = 0 
    nextRecord = 10 ' Starting Row number of Record in report
 
    ' Find the last row in the worksheet
    lastRow = safeDataSheet.Cells(safeDataSheet.Rows.Count, 1).End(xlUp).Row
    
    
    ' Loop through the rows in the worksheet and extract data between the date range
    For i = 2 To lastRow ' Assuming data starts from row 2
        
        If safeDataSheet.Cells(i, "D") >= startDate And safeDataSheet.Cells(i, "D") <= endDate And safeDataSheet.Cells(i, "Y") = payType Then
            
            recDate = safeDataSheet.Cells(i, "D")
            recNum = safeDataSheet.Cells(i, "E")
            details = safeDataSheet.Cells(i, "O")
            total = safeDataSheet.Cells(i, "U")
            isCanceled = safeDataSheet.Cells(i, "W")

            glCode = safeDataSheet.Cells(i, "J")
            incomeCode = safeDataSheet.Cells(i, "G")
            activity = safeDataSheet.Cells(i, "H")


            recDateArr(nextRecord) = recDate
            recNumArr(nextRecord) = recNum 
            detailsArr(nextRecord) = details
            totalArr(nextRecord) = total
            isCanceledArr(nextRecord) = isCanceled
            
            nextRecord = nextRecord + 1

            For indexOfFoundIncomeCode = 1 To UBound(incomeCodeArr)
                ' Debug.Print "looking for " + incomeCode
                If incomeCodeArr(indexOfFoundIncomeCode) = incomeCode Then
                    ' Found the income code, exit the loop and return the row index
                    Exit For
                End If
                
            Next indexOfFoundIncomeCode

            If indexOfFoundIncomeCode <= UBound(incomeCodeArr) Then
                ' Found the income code, do something with it (e.g. print the row data)
                'Debug.Print IncomeDetails(indexOfFoundIncomeCode, totalCol)
                totalSum(indexOfFoundIncomeCode) = totalSum(indexOfFoundIncomeCode) + total
            Else
                ' Debug.Print "New Income"
                newRecord = newRecord + 1
                glCodeArr(newRecord) = glCode
                incomeCodeArr(newRecord) = incomeCode
                activityArr(newRecord) = activity
                totalSum(newRecord) = total
                'Debug.Print "index of " & CStr(incomeCodeArr(newRecord)) & " is " & newRecord
                '' Income code not found
                
            End If

        End If


    Next i

    If (isWeekly) Then 

        If (nextRecord - 9 > avlRowsForTransWeekly) Then
            totalRowsOfData = nextRecord - 9
            insertNewRows = totalRowsOfData - avlRowsForTransWeekly 
            nextNagudhuRecord = nextNagudhuRecord + insertNewRows
             MsgBox "Inserting for Weekly Tans " & insertNewRows 
            ActiveReportSheet.Rows("15:" & 14 + insertNewRows).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove     
        End If

        If (newRecord > avlRowsForCatDetails) Then
            totalRowsOfData = newRecord
            insertNewRows = totalRowsOfData - avlRowsForCatDetails 
           ' nextNagudhuRecord = nextNagudhuRecord + insertNewRows
             MsgBox "Inserting for Weekly Cat " & insertNewRows 
            ActiveReportSheet.Rows(nextNagudhuRecord + 3 & ":" & nextNagudhuRecord + insertNewRows + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 
            
        End If
        
    Else
        If ((nextRecord - 9) > avlRowsForTransMonthly) Then
            
            totalRowsOfData = nextRecord - 9
            insertNewRows = totalRowsOfData - avlRowsForTransMonthly 
            nextNagudhuRecord = nextNagudhuRecord + insertNewRows
             MsgBox "Inserting for Monthly Trans " & insertNewRows 
            ActiveReportSheet.Rows("31:" & 30 + insertNewRows).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 
      
        End If
        
        If (newRecord > avlRowsForCatDetails) Then
            
            totalRowsOfData = newRecord
            insertNewRows = totalRowsOfData - avlRowsForCatDetails 
            'nextNagudhuRecord = nextNagudhuRecord + insertNewRows
             MsgBox "Inserting for Monthly Cat " & insertNewRows 
            ActiveReportSheet.Rows(nextNagudhuRecord + 3 & ":" & nextNagudhuRecord + insertNewRows + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 
            
        End If
    End If


    

    For z = 10 To nextRecord

            ActiveReportSheet.Range("B" & z).Value = recDateArr(z)
            ActiveReportSheet.Range("C" & z).Value = recNumArr(z)
            ActiveReportSheet.Range("D" & z).Value = detailsArr(z)
            ActiveReportSheet.Range("I" & z).Value = totalArr(z)

            If (isCanceledArr(z) = "Yes") Then 
                ActiveReportSheet.Range("B"& z & ":I"& z).Font.Strikethrough = True 'add strikethrough
                ActiveReportSheet.Range("B"& z & ":I"& z).Font.Color = vbRed 'change font color to red
            Else
                ActiveReportSheet.Range("B"& z & ":I"& z).Font.Strikethrough = False 'add strikethrough
                ActiveReportSheet.Range("B"& z & ":I"& z).Font.Color = vbBlack 'change font color to red
            End If
    Next z

    
    For y = 1 To newRecord
        'Debug.Print CStr(glCodeArr(y)) & " | " & CStr(incomeCodeArr(y)) & " | " & CStr(activityArr(y)) & " | " & CStr(totalSum(y))
        ActiveReportSheet.Range("B" & nextNagudhuRecord + y).Value = glCodeArr(y)
        ActiveReportSheet.Range("C" & nextNagudhuRecord + y).Value = incomeCodeArr(y)
        ActiveReportSheet.Range("D" & nextNagudhuRecord + y).Value = activityArr(y)
        ActiveReportSheet.Range("K" & nextNagudhuRecord + y).Value = totalSum(y)
    Next y

    'SafeData.Close SaveChanges:=False
    'SafeData.Close SaveChanges:=True
    Application.DisplayAlerts = True ' Enable alerts
    Application.ScreenUpdating = True ' Enable screen updating

    

    'ErrorHandler:
    ' Error handling code here
        ' SafeData.Close SaveChanges:=False
        ' MsgBox "An error occurred: " & Err.Description
        ' Exit Sub

End Sub
