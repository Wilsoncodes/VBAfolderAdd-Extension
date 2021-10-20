Attribute VB_Name = "ManipulateSave"
Sub splitCell()
    Dim lRow As Long
    lRow = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lRow
        Results = Split(Sheets("Sheet1").Range("A" & i).Text, ".")
        ArrayLen = UBound(Results) - LBound(Results) + 1
        a = 0
        Do While a < ArrayLen - 1
            Sheets("Sheet1").Range("C" & i).Value = Sheets("Sheet1").Range("C" & i).Value & Results(a)
            a = a + 1
        Loop
            If ArrayLen - 1 = 0 Then
                Sheets("Sheet1").Range("C" & i).Value = Results(a)
            Else
                Sheets("Sheet1").Range("B" & i).Value = Results(a)
            End If
            Sheets("Sheet1").Range("C" & i).Value = Sheets("Sheet1").Range("C" & i).Value & "." & Sheets("Sheet1").Range("G2").Value
    Next i
End Sub

Sub saveRename()

    Dim lRow As Long
    lRow = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
    Dim direct As String
    direct = Sheets("Sheet1").Range("G1").Value & "\"
    For i = 2 To lRow
        Name direct & Sheets("Sheet1").Range("A" & i).Value As direct & Sheets("Sheet1").Range("C" & i).Value
    Next i
End Sub

