Attribute VB_Name = "clearData"
Sub cleanRun()
Attribute cleanRun.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Sheet1").Range("A2:C2").Select
    Sheets("Sheet1").Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Sheet1").Range("A2").Select
End Sub
