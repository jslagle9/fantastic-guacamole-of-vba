Attribute VB_Name = "Module1"
Public Sub TheRanger()

''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro sets up named ranges
' for reference in INDEX MATCH formulas
'
' Process:
' Pulls existing names in workbook into AllNames
' string for conflict check
' Prompts user to use sheet or define manually
' (split with inputbox or msgbox)
' Prompts user to name the full table range
' checks for conflicts against AllNames str
'' if conflict fail, reprompt with error msg
'' if conflict pass, eval RangeName against SpecChar
' function, change as needed
' assign RangeName, add RangeName to AllNames
'
' Headers level:
' string NamedRange.name from RangeName
' checks for conflicts against AllNames str
' same as above
'
' Column level:
' For Each col in range
' string NamedRange.name from cell.value
'' checks for conflicts against AllNames str
' same as above
''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WorkBk As Workbook
Dim sh, sh2 As Worksheet
Dim nm As Name
Dim rng, rng2, rng3, rng4, rng5, temprng As Range
Dim TableRng, HeaderRng As Range
Dim FolderPath, tdy, FileName, SheetStr, CellVal As String
Dim PassString As String
Dim AllNames As String
Dim shName, CoStr, RStr, RowStr, TableStr, HStr, HeaderStr, ColumnStr, RangeName, AllShtNames As String
Dim i, c, iVal, LastRow, LastCol, ShtNum As Integer
Dim diaFolder, fDialog As FileDialog


Result1 = MsgBox("This macro prompts you to" & _
vbCrLf & "select a file, prompts to select a " & _
vbCrLf & "sheet or a specific range, name that" & _
vbCrLf & "range, and auto-magically creates" & _
vbCrLf & "INDEX-MATCH ready named ranges for" & _
vbCrLf & "the data and its headers and columns." & _
vbCrLf & _
vbCrLf & "Continue?" & vbCrLf, vbYesNo + vbExclamation)
If Result1 = vbNo Then
GoTo CancelMac
Else: GoTo RunTheSub
End If

RunTheSub:
'Pick the workbook
' Open file dialog to select file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Select a file"
    fDialog.Show
    
    'Debug breakpoint goes here, otherwise file dialog will run the rest
    If (fDialog.SelectedItems.Count = 0) Then GoTo GracefulExit1
    FileName = fDialog.SelectedItems(1)

'Open the workbook, activate
Set WorkBk = Workbooks.Open(FileName)

' Set wkbk and sh variables
c = WorkBk.Names.Count

'Build existing list of names
For i = 1 To c
AllNames = AllNames & WorkBk.Names(i).Name & " "
'i = i + 1
Next
'reset counter
i = 0
Debug.Print AllNames

Result1 = MsgBox("This macro defaults to using " & _
vbCrLf & "an entire worksheet for the " & _
vbCrLf & "named range. Click Yes to use " & _
vbCrLf & "an entire worksheet, and click " & _
vbCrLf & "No to select a range with InputBox. " & vbCrLf, vbYesNo + vbExclamation)
If Result1 = vbNo Then
GoTo VariantUseSelectedRange
Else: GoTo DefaultUseWholeSheet
End If

VariantUseSelectedRange:
' variant prompts user to select a range with inputbox
Set TableRng = Application.InputBox(prompt:="Click and drag to select Table Range to be named", Type:=8)
'name table manually
Set sh = TableRng.Worksheet
LastRow = TableRng.Rows.Count
LastCol = TableRng.Columns.Count
GoTo NameTheTable

DefaultUseWholeSheet:
'pick a sheet in workbook
' Set wkbk and sh variables
c = WorkBk.Worksheets.Count
'Build existing list of names
For i = 1 To c
AllShtNames = AllShtNames & WorkBk.Worksheets(i).Name & " - " & WorkBk.Worksheets(i).Index & ", " & vbCrLf
'i = i + 1
Next
Debug.Print AllShtNames
'reset counter
i = 0
ShtNum = Application.InputBox(prompt:="Enter the Index Number of the worksheet as listed below:" & vbCrLf & AllShtNames, Type:=1)
Set sh = Worksheets(ShtNum)
'Define ranges in sht
' find last row
LastRow = sh.UsedRange.Rows.Count
' find last column
LastCol = sh.UsedRange.Columns.Count
'name table manually
GoTo NameTheTable

' take inputbox name
NameTheTable:
TableStr = Application.InputBox("Enter full table range name")
' name clean, conflictcheck
PassString = TableStr
Debug.Print PassString
Call NameCleaner(PassString)
Debug.Print PassString
Call ConflictCheck(PassString, AllNames)
Debug.Print PassString

' define string var for name ref sh
SheetStr = "='" & sh.Name & "'!"
CoStr = "_C"
HStr = "_H"
HeaderStr = "_Headers"

''''''''''''''''''''''''''''''''''''''''''''''''''
'! MAIN PROCESS
With WorkBk
'! define full table range for INDEX MATCH
'build name string
PassString = TableStr & "_Table"
' pass name string to NameCleaner function
Call NameCleaner(PassString)
Debug.Print PassString
' conflict check
Call ConflictCheck(PassString, AllNames)
Debug.Print PassString
' commit name
.Names.Add Name:=PassString, _
RefersToR1C1:=SheetStr & "R1C1:R" & LastRow & "C" & LastCol
' commit to AllNames
AllNames = AllNames & " " & PassString
'reset counter, reset passstr
i = 0
PassString = ""

''''''''''''''''''''''''

'! define header range for INDEX MATCH
'build name string
PassString = TableStr & HeaderStr
' pass name string to NameCleaner function
Call NameCleaner(PassString)
' conflict check
Call ConflictCheck(PassString, AllNames)
' commit name
.Names.Add Name:=TableStr & HeaderStr, _
RefersToR1C1:=SheetStr & "R1C1:R1" & "C" & LastCol
' commit to AllNames
AllNames = AllNames & " " & PassString
'reset counter, reset passstr
i = 0
PassString = ""

''''''''''''''''''''''''
'! Define column headers for INDEX MATCH
For i = 1 To LastCol
' get header value for column
CellVal = sh.Range(Cells(1, i).Address()).Value
Debug.Print i
Debug.Print CellVal
' build name string
PassString = TableStr & CellVal & HStr
' pass name string to NameCleaner function
Call NameCleaner(PassString)
Debug.Print PassString
' conflict checker
Call ConflictCheck(PassString, AllNames)
Debug.Print PassString
' commit name
.Names.Add Name:=PassString, _
RefersToR1C1:=SheetStr & "R1C" & i
' commit to AllNames
AllNames = AllNames & " " & PassString
Debug.Print AllNames
' iterate
'i = i + 1
Next
'reset counter, reset passstr
i = 0
PassString = ""

''''''''''''''''''''''''
'! Define columns for INDEX MATCH
For i = 1 To LastCol
' get header value for column
CellVal = sh.Range(Cells(1, i).Address()).Value
' build name string
PassString = TableStr & CellVal & CoStr
' pass name string to NameCleaner function
Call NameCleaner(PassString)
Debug.Print PassString
' conflict checker
Call ConflictCheck(PassString, AllNames)
Debug.Print PassString
' commit name
.Names.Add Name:=PassString, _
RefersToR1C1:=SheetStr & "R1C" & i & ":R" & LastRow & "C" & i
' commit to AllNames
AllNames = AllNames & " " & PassString
' iterate
'i = i + 1
Next
' End With ActiveWorkbook
End With
GoTo Finished

CancelMac:
MsgBox "Macro did not run.", vbCritical, "Notice!"
Exit Sub

GracefulExit1:
MsgBox "Macro did not run. Run macro again and select a file when prompted.", vbCritical, "Notice!"
Exit Sub

Finished:
MsgBox "Macro complete!", vbInformation, "Ding!"
'WorkBk.Close savechanges:=True

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
Function NameCleaner(PassString As String) As String
Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?,:,;,-,/"  'modify as needed
Dim newString As String
Dim char As Variant
newString = PassString
For Each char In Split(SpecialCharacters, ",")
    newString = Replace(newString, char, "")
Next
' replace spaces with underscores
newString = Replace(newString, " ", "_")
Debug.Print newString
PassString = newString
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''
Function ConflictCheck(PassString As String, AllNames As String) As String
Dim iVal, i As Integer
i = 1
iVal = InStr(1, AllNames, PassString, 1)

Do Until iVal = 0
iVal = InStr(1, AllNames, PassString, 1)
PassString = PassString & "_" & i
'i = i + 1
Debug.Print
Loop
i = 0
End Function
Public Sub TheAntiRanger()
' Deletes all named ranges with prompt
Dim xName As Name
Dim AllNames, FileName As String
Dim i, c As Long
Dim diaFolder, fDialog As FileDialog
Dim WorkBk As Workbook

Result1 = MsgBox("This macro prompts you to" & _
vbCrLf & "select a file, confirm if you want " & _
vbCrLf & "to delete all named ranges." & _
vbCrLf & _
vbCrLf & "Continue?" & vbCrLf, vbYesNo + vbExclamation)
If Result1 = vbNo Then
GoTo CancelMac
Else: GoTo RunTheSub
End If

RunTheSub:
'Pick the workbook
' Open file dialog to select file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Select a file"
    fDialog.Show
    
    'Debug breakpoint goes here, otherwise file dialog will run the rest
    If (fDialog.SelectedItems.Count = 0) Then GoTo GracefulExit1
    FileName = fDialog.SelectedItems(1)

'Open the workbook, activate
Set WorkBk = Workbooks.Open(FileName)

  ' count names
c = WorkBk.Names.Count
Debug.Print c
  ' if 0 ranges then exit with msgbox
If c = 0 Then GoTo NoRanges

'Build list of names
For i = 1 To c
AllNames = AllNames & WorkBk.Names(i).Name & ", " & vbCrLf
'i = i + 1
Debug.Print AllNames
Next

'msgbox prompt
  Result = MsgBox("Delete all " & c & " named range(s)? " & vbCrLf & "Associated formulae may not work!" & vbCrLf & "List: " & vbCrLf & AllNames, vbYesNo + vbCritical)
' delete all names on Yes button
If Result = vbYes Then
For Each Name In WorkBk.Names
Name.Delete
Next
MsgBox c & " ranges deleted!", vbExclamation
Exit Sub
' exit sub on No button
Else
MsgBox "No ranges were deleted! " & c & " named range(s) unchanged.", vbExclamation
Exit Sub
End If

CancelMac:
MsgBox "Macro did not run!", vbExclamation
Exit Sub

NoRanges:
MsgBox "No ranges found!", vbExclamation
Exit Sub

GracefulExit1:
MsgBox "Macro did not run. Run macro again and select a file when prompted.", vbCritical, "Notice!"
WorkBk.Close savechanges:=False
Exit Sub

End Sub
