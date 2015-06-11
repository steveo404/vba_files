Attribute VB_Name = "modImportXlsxfiles"
Option Compare Database
Option Explicit

Sub ImportXlsx()

    Dim strPath As String
    Dim strFile As String
    Dim strTblNm As String
    Dim fileDate As String
    Dim TMPNM As String
    Dim Count As Integer
    
    strPath = "C:\Users\soneal\Documents\IDEA_Projects\Common Input Files\"
    strFile = Dir(strPath & "*.xlsx")
    fileDate = TodayDate()
    
    While strFile <> ""
        strTblNm = Mid(strFile, 1, (Len(strFile) - 5)) + "_" + fileDate
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, strTblNm, strPath & strFile, True
        strFile = Dir()
    Wend
    
    strPath = "C:\Users\soneal\Documents\IDEA_Projects\PC-6\Input Files\"
    strFile = Dir(strPath & "*.xlsx")
    strTblNm = Mid(strFile, 1, (Len(strFile) - 5)) + "_" + fileDate
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, strTblNm, strPath & strFile, True
    
    Application.RefreshDatabaseWindow
    
End Sub


