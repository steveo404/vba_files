Attribute VB_Name = "modTimeCard"
Option Compare Database
Option Explicit

Public Static Sub TimeCard_Review()

Dim db As Database
Dim strSQL As String
Dim i As Integer
Dim strTableName As String


Set db = CurrentDb()

strTableName = "tblTimeCard_Step1"
If TableExists(strTableName) Then
    DoCmd.DeleteObject acTable, strTableName
End If

strSQL = "SELECT [2014 Timecard Data_07092014].Office, "
strSQL = strSQL + "[2014 Timecard Data_07092014].Employee, "
strSQL = strSQL + "[2014 Timecard Data_07092014].[Last Name], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[First Name], "
strSQL = strSQL + "[2014 Timecard Data_07092014].Week, "
strSQL = strSQL + "[2014 Timecard Data_07092014].[Audit Number], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approval Level], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approval Level Desc], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approval Date], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approver Employee], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approver Job Class], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[EMP Approver Job Title], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approval Level], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approval Level Desc], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approval Date], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approver Employee], "
strSQL = strSQL + "[Firmwide Listing_07092014].[First Name] AS [SUP Approver First Name], "
strSQL = strSQL + "[Firmwide Listing_07092014].[Last Name] AS [SUP Approver Last Name], "
strSQL = strSQL + "[Firmwide Listing_07092014].Email AS [SUP Approver Email], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approver Job Class], "
strSQL = strSQL + "[2014 Timecard Data_07092014].[SUP Approver Job Title] "
strSQL = strSQL + "INTO tblTimeCard_Step1 "
strSQL = strSQL + "FROM [2014 Timecard Data_07092014] "
strSQL = strSQL + "LEFT JOIN [Firmwide Listing_07092014] ON [2014 Timecard Data_07092014].[SUP Approver Employee] = [Firmwide Listing_07092014].[Employee Number]; "

DoCmd.SetWarnings False
db.Execute (strSQL)

Application.RefreshDatabaseWindow


End Sub


