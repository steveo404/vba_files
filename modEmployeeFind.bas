Attribute VB_Name = "modEmployeeFind"
Option Compare Database

Public Static Sub EmployeeFind()

Dim db As Database
Dim strSQL As String
Dim i As Integer
Dim strTableName As String
Dim strField As String
Dim strCriteria As String

i = 1
strTableName = "EmpChain_5Levels"
strField = "Level" + Trim(Str(i))
strCriteria = strField + " = '34202'"



'Set db = CurrentDb()

'strSQL = "ALTER TABLE TestChain ADD COLUMN ID INT"
'strSQL = strSQL + "[2014 Timecard Data_07092014].Employee, "

'DoCmd.SetWarnings False
'db.Execute (strSQL)

If IsNull(DLookup("ChainID", strTableName, strCriteria)) Then
    'MsgBox ("Not Found")
End If

'i = DLookup("ChainID", "EmpChain_5Levels", "Lvl3 = '13307'")


Do Until i = 6
    If IsNull(DLookup("ChainID", strTableName, strCriteria)) Then
        i = i + 1
        strField = "Level" + Trim(Str(i))
        strCriteria = strField + " = '16600'"
    Else
        MsgBox ("Found It")
    End If
Loop



End Sub
