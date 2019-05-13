Attribute VB_Name = "mainModule"
Public cnStr As String

Sub Main()
    'cnStr = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=DBA;user id=pmsdev;password=3xp1r3d;DATA SOURCE=COMMDCPROJECT"
    Dim TextFileData As String, MyArray() As String, i As Long
    Open App.Path & "\config.txt" For Binary As #1
    TextFileData = Space$(LOF(1))
    Get #1, , TextFileData
    Close #1
    MyArray() = Split(TextFileData, vbCrLf)
    cnStr = Replace$(MyArray(0), "cnStr:", "")
    
    frmClientView.Show
End Sub
