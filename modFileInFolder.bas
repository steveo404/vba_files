Attribute VB_Name = "modFileInFolder"
Option Compare Database

Sub Read_Files_In_Folder()
    Dim rs As DAO.Recordset
    strSQL = "DELETE * FROM List_Of_Files"
    CurrentDb.Execute strSQL
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fldr = fs.GetFolder("I:\Continuous Auditing\HR Data\InputFiles\")
    Set fls = fldr.Files
    Set rs = CurrentDb.OpenRecordset("List_Of_Files")
    For Each fl In fls
        rs.AddNew
        rs.Fields(0) = fl.Name
        rs.Fields(1) = fl.DateCreated
        rs.Fields(2) = fl.DateLastModified
        rs.Fields(3) = fl.Size \ 1024 + 1
        rs.Update
        Next fl
    rs.Close
    Set rs = Nothing
    
End Sub


