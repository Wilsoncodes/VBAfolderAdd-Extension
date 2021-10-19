Attribute VB_Name = "OpenFiles"
Sub test()
    recordFiles ("c:\tmp\")
End Sub

Sub recordFiles(filePath)
    Dim strFileName As String
    strFileName = Dir(filePath)
    Do While Len(strFileName) > 0
        Debug.Print strFileName
        strFileName = Dir
    Loop
End Sub
