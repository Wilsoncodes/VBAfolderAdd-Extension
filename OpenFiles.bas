Attribute VB_Name = "OpenFiles"
Sub recordFiles(filePath)
    clearData.cleanRun
    Dim strFileName As String
    strFileName = Dir(filePath)
    Dim a As Integer
    a = 2
    Do While Len(strFileName) > 0
        Sheets("Sheet1").Range("A" & a).Value = strFileName
        strFileName = Dir
        a = a + 1
    Loop
End Sub
