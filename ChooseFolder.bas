Attribute VB_Name = "ChooseFolder"
Sub chooseFolder()
Dim strFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            strFolder = .SelectedItems(1)
        End If
    End With
    
    If strFolder <> "" Then
        Sheets("Sheet1").Range("G1").Value = strFolder
        OpenFiles.recordFiles (strFolder + "\") 'as stated in OpenFiles module
    End If
End Sub
