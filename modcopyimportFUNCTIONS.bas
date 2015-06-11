Attribute VB_Name = "modcopyimportFUNCTIONS"
Option Compare Database
Option Explicit

Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Public Function basCopyFromXLS(SourceFile$, sImportTable$, bClear As Boolean, Optional sSheet$) As Boolean
'created & first used 01/18/06
'copy xls data; append to Access temp table
'workaround to this TransferSpreadsheet issue:
' Each field in the spreadsheet must be of the same data type as the corresponding field in Access.

Dim xlApp As Object
Dim xlWorkbook As Object
Dim xlWorksheet As Object
Dim strSQL$

basCopyFromXLS = True
On Error GoTo err_Import

DoCmd.SetWarnings False
If bClear Then
strSQL = "Delete * FROM " & sImportTable
CurrentDb.Execute strSQL
End If

Set xlApp = CreateObject("Excel.Application")
Set xlWorkbook = xlApp.Workbooks.Open(fileName:=SourceFile)
'If IsNothing(sSheet) Then
If sSheet = "" Then
    Set xlWorksheet = xlWorkbook.Worksheets(1)
Else
    Set xlWorksheet = xlWorkbook.Worksheets(sSheet)
End If
xlWorksheet.Activate
xlWorksheet.UsedRange.Select
xlWorksheet.UsedRange.Copy

DoCmd.OpenTable sImportTable
DoCmd.RunCommand acCmdPasteAppend
DoCmd.Close acTable, sImportTable

exit_import:
ClearClipboard
xlApp.Quit
Set xlWorkbook = Nothing
Set xlWorksheet = Nothing
Set xlApp = Nothing
Exit Function

err_Import:
MsgBox Error$ & "; Import Incomplete"
basCopyFromXLS = False
GoTo exit_import
End Function


'**********************************************************
'* Function: Returns true if argument evaluates to nothing
'* 1. IsNothing(Nothing) -> True
'* 2. IsNothing(NonObjectVariableOrLiteral) -> False
'* 3. IsNothing(ObjectVariable) -> True if instantiated,
'*                                 otherwise False
'**********************************************************
Public Function IsNothing(pvarToTest As Variant) As Boolean
    On Error Resume Next
    IsNothing = (pvarToTest Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function 'IsNothing




Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function





