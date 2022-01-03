Attribute VB_Name = "Module1"
Sub ShowForm()
    'Show form for entering changes
    frmChange.Show
End Sub

Sub RunReport()

'----------------------------------------------------
'This is the load files button marco that clear table
'and refresh all power query


'Turn off Screen Updates when running macro
Application.ScreenUpdating = False

'Check if there is file path for the loading
If ActiveSheet.Range("C6") = Empty Or ActiveSheet.Range("C7") = Empty Then
    'Sents message if no file path
    MsgBox "No Files Selected." & vbNewLine & vbNewLine & "Please select files."
    Exit Sub
End If

'Clear Error Report Table
Call ClearTable

'Refresh all power query
If RefreshAll(False) Then

'Turns on Screen Update
Application.ScreenUpdating = True

MsgBox ("Files have been loaded")

End If
End Sub

Sub ErrorReport(tbl As ListObject)

'Turn off Screen Updates when running macro
Application.ScreenUpdating = False

Dim myArr As Variant, errArr() As Variant, errTbl As ListObject
Dim equipCode As Variant, posCode As Variant, StartDate As Double, EndDate As Double
Dim equip As Variant, pos As Variant, newrow As ListRow

'Set Worksheet variables
Set wsDate = Worksheets("Input")
Set wsParm = Worksheets("Parameters")

'Checks the table name and then set the row number
If tbl.Name = "Pilot" Then
    Row = 22
Else
    Row = 3
End If

dtEndDate = wsDate.Range("F3")

'Set the Bid month date range and month from input tab
StartDate = Format(wsDate.Range("E3"), "yyyymmdd")
EndDate = Format(wsDate.Range("F3"), "yyyymmdd")
mo = wsDate.Range("C3")

'Load array
myArr = tbl.DataBodyRange.Value

'Set variable to run logic
equipCode = wsParm.ListObjects("EquipCode").DataBodyRange.Value
posCode = wsParm.ListObjects("PosCode").DataBodyRange.Columns(2).Value
Set errTbl = Worksheets("Error Report").ListObjects("error")

'Get last row of data set
LastRow = tbl.DataBodyRange.Rows.Count
r = 1

'Loop thru dataset to see if there are errors. See errors below
For i = 1 To LastRow
  If myArr(i, 3) <> "7PO" Then
    If myArr(i, 2) < StartDate Then
        reason = "1-Prior to Bid Month"
        GoTo Load
    ElseIf myArr(i, 2) > EndDate Then
        reason = "1-After Bid Month"
        GoTo Load
    End If
    If myArr(i, 17) > 75 Then
        reason = "2-Over 75 hours"
        GoTo Load
    End If
    If IsInArray(myArr(i, 13), equipCode) = False Then
        reason = "3-Invaild Equip Code"
        GoTo Load
    End If
    If IsInArray(myArr(i, 12), posCode) = False Then
        reason = "4-Invaild Position Code"
        GoTo Load
    End If
    If myArr(i, 20) = "T" Then
        reason = "5-Employee Termed"
        GoTo Load
    End If
    If myArr(i, 18) = "" Then
        reason = "6-No Earning Code " + "(" + myArr(i, 3) + ")"
        GoTo Load
    End If
    'This is used to remove FL9 from dataset
    If myArr(i, 3) = "FL9" Then
        reason = "7-FL9 UTA"
        GoTo Load
    End If
    If myArr(i, 20) = "L" Then
        reason = "8-Employee on Leave"
        GoTo Load
    End If
    If myArr(i, 21) = "ALPAC" And myArr(i, 27) < dtEndDate Then
        reason = "9-Pilot Surfer"
        GoTo Load
    End If
    If myArr(i, 3) = "LLP" Then
        reason = "7-LLP UTA"
        GoTo Load
    End If
  End If
Load:
    'Check the reason
    If IsEmpty(reason) = False Then
    
    'Adds a new row
    Set newrow = errTbl.ListRows.Add
        'Sets new data
        With newrow
            .Range(1) = mo
            .Range(2) = myArr(i, 15)
            .Range(3) = myArr(i, 1)
            .Range(4) = myArr(i, 19)
            .Range(6) = myArr(i, 12)
            .Range(7) = myArr(i, 2)
            .Range(8) = myArr(i, Row)
            .Range(9) = myArr(i, 13)
            .Range(10) = myArr(i, 17)
            .Range(11) = reason
            If reason = "7-FL9 UTA" Or reason = "7-LLP UTA" Then
            'Add to exclude from main dataset
                .Range(12) = "X"
            End If
        End With
        r = r + 1
    End If
    
    'Clear out variable for reuse
    reason = Empty
Next

'Sort the table.
With errTbl.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("Error[Crew Type]"), SortOn:=xlSortOnValues
    .SortFields.Add Key:=Range("Error[Reason]"), SortOn:=xlSortOnValues
    .Header = xlYes
    .Apply
End With

'Clear out theh array variable
myArr = Empty

Application.ScreenUpdating = True
End Sub

Private Function IsInArray(valToBeFound As Variant, arr As Variant)
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Sub ClearTable()

With Worksheets("Error Report").ListObjects(1)
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

End Sub

Sub WriteFile(fileName As String, ws As String, obname As String)
'----------------------------------------------------------
'needs to set the fileName, worksheet to export, object name

'Set the file path
filePath = Worksheets("Input").Range("C12").Value & "\"

'Set the month and yr for filename
monthText = Worksheets("Input").Range("C3")
yr = Worksheets("Input").Range("C2")

'Convert month to month number
monthNum = Month(DateValue("01 " & monthText & " " & yr))

'Exclude column to check
If ws = "Pilot" Then
    vExcludeCol = 24
Else
    vExcludeCol = 23
End If

Set tbl = ThisWorkbook.Worksheets(ws).ListObjects(obname)

'Load the array
myArr = tbl.DataBodyRange.Value

'Set Full path
myFile = filePath & fileName & "_" & Format(monthNum, "00") & Right(yr, 2) & ".txt"

'Set last row for dataset
LastRow = tbl.DataBodyRange.Rows.Count
r = 1

'Create file to write to
Open myFile For Output As #1

'Loop thru dataset
For i = r To LastRow

'Check the exclude column
If myArr(i, vExcludeCol) <> "X" Then
    'Set the text line
    LineStr = Format(myArr(i, 1), "@@@@@@@@@") _
               + Format(myArr(i, 2), "@@@@@@@@") _
               + Format(myArr(i, 3), "@@@") _
               + Format(myArr(i, 4), "@") _
               + Format(myArr(i, 5), "000") _
               + Format(myArr(i, 6), "00") _
               + Format(myArr(i, 7), "@@@@") _
               + Format(myArr(i, 8), "@@@@@@@@") _
               + Format(myArr(i, 9), "@@@") _
               + Format(myArr(i, 10), "@@@") _
               + Format(myArr(i, 11), "@@@@@") _
               + Format(myArr(i, 12), "@@") _
               + Format(myArr(i, 13), "@@@") _
               + Format(myArr(i, 14), "@@@@@")
     Print #1, LineStr
End If
Next

'Close the text file
Close #1



End Sub
Sub RunErrorReport()

Dim tblFA As ListObject, tblPilot As ListObject

Set tblFA = ThisWorkbook.Worksheets("FA").ListObjects("FA")
Set tblPilot = ThisWorkbook.Worksheets("Pilot").ListObjects("Pilot")

Set errTbl = Worksheets("Error Report").ListObjects("error")

If errTbl.DataBodyRange Is Nothing Then
    Call ErrorReport(tblFA)
    Call ErrorReport(tblPilot)
End If

MsgBox ("Error Report is ready for review")

End Sub

Sub ExportFiles()

Call WriteFile("FA_Final_Pay", "FA", "FA")
Call WriteFile("Pilot_Final_Pay", "Pilot", "Pilot")

MsgBox ("Files have been created")

End Sub

Function RefreshAll(Optional ByVal bOn As Boolean = True) As Boolean


On Error GoTo Error
For Each con In ThisWorkbook.Connections
    With con.OLEDBConnection
                .BackgroundQuery = False
                .Refresh
    End With
Next con


If bOn Then
    MsgBox "Refresh Complete"
    
End If

RefreshAll = True

Exit Function
Error:
MsgBox "Error in Refreshing " & con
RefreshAll = False

End Function

Private Function GetFile() As String
Dim file As FileDialog
Dim sItem As String
Set file = Application.FileDialog(msoFileDialogFilePicker)
With file
    .Title = "Select a File"
    .AllowMultiSelect = False
    '.InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFile = sItem
Set file = Nothing
End Function

Sub FASelectFile()

ActiveSheet.Range("C6") = GetFile()

End Sub

Sub PilotSelectFile()

ActiveSheet.Range("C7") = GetFile()

End Sub
Sub VacSickSelectFile()

ActiveSheet.Range("C8") = GetFile()

End Sub
Private Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    '.InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function

Sub ExportFldr()

ActiveSheet.Range("C12") = GetFolder()
End Sub

Sub ExportFldrER()

ActiveSheet.Range("B4") = GetFolder()
End Sub
Sub ExportErrorReport()

Call ExportErrRptParam("FA")
Call ExportErrRptParam("Pilot")

End Sub
Sub ExportErrRptParam(ByVal grp As String)

Dim wb As Workbook, wsNew As Worksheet
Dim ws As Worksheet
Dim arrData As Variant, arrHeader As Variant

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set ws = ActiveSheet
Set errTbl = Worksheets("Error Report").ListObjects("error")
Set redTbl = Worksheets("Redline").ListObjects("Redline")

If Not errTbl.DataBodyRange Is Nothing Then

    arrHeader = errTbl.HeaderRowRange
    arrData = errTbl.DataBodyRange
    
    flpath = ws.Cells.Find("Export Path").Offset(0, 1)
    
    If flpath = Empty Then
        MsgBox "Enter a Folder Path"
        Call ExportFldrER
        Exit Sub
    End If
    
    fileName = Worksheets("Input").Range("C2") & "_" & Worksheets("Input").Range("C3") & "_" & grp & "ErrorReport"
    r = 2
    fullPath = flpath & "\" & fileName & ".xlsx"
    Set wb = Workbooks.Add
    wb.Worksheets("Sheet1").Name = "Error Report"
    Set wsNew = wb.Worksheets("Error Report")
    For c = 1 To UBound(arrHeader, 2)
        wsNew.Cells(1, c) = arrHeader(1, c)
    Next
    For i = 1 To UBound(arrData)
        If arrData(i, 5) = grp Then
            For c = 1 To UBound(arrData, 2)
                wsNew.Cells(r, c) = arrData(i, c)
            Next
            r = r + 1
        End If
    Next
    wsNew.Cells.EntireColumn.AutoFit

    r = 2
    arrHeader = redTbl.HeaderRowRange
    arrData = redTbl.DataBodyRange
    
    Set wsNew = wb.Worksheets.Add
    wsNew.Name = "Redline"
    For c = 1 To UBound(arrHeader, 2)
        wsNew.Cells(1, c) = arrHeader(1, c)
    Next
    For i = 1 To UBound(arrData)
        If arrData(i, 5) = grp And arrData(i, UBound(arrData, 2)) = "Y" Then
            For c = 1 To UBound(arrData, 2)
                wsNew.Cells(r, c) = arrData(i, c)
            Next
            r = r + 1
        End If
    Next
    wsNew.Cells.EntireColumn.AutoFit
    
    wb.SaveAs (fullPath)
    wb.Close

    MsgBox grp & "Report Export Completed"

Else
    MsgBox "Please Run Error Report"

End If



Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub
