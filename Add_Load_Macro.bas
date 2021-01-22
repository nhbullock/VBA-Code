Attribute VB_Name = "Module11"
Sub Add_Load()
Attribute Add_Load.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_Load Macro
'
Dim i As Integer, Row As Integer

Application.ScreenUpdating = False

For i = 2 To 1000
    If Worksheets("Protein Schedule").Cells(i, 2).Value = CLng((Mid(Worksheets("ETA").Cells(2, 22), 22, 9))) Then
        If Worksheets("Protein Schedule").Cells(i, 3).Value = CLng(Mid(Worksheets("ETA").Cells(2, 22), 32)) Then
            Row = i
            i = 1000
        End If
    End If
Next


Worksheets("Protein Schedule").Cells(Row, 11).Value = CLng(Right(Worksheets("ETA").Cells(11, 22).Value, 5))
Worksheets("Protein Schedule").Cells(Row, 12).Value = TimeValue(Right(Worksheets("ETA").Cells(15, 22).Value, 5)) & DateValue(Mid(Worksheets("ETA").Cells(15, 22).Value, 22, 10))
Worksheets("Protein Schedule").Cells(Row, 13).Value = TimeValue(Right(Worksheets("ETA").Cells(16, 22).Value, 5)) & DateValue(Mid(Worksheets("ETA").Cells(16, 22).Value, 22, 10))
Worksheets("Protein Schedule").Cells(Row, 15).Value = DateValue(Mid(Worksheets("ETA").Cells(16, 22).Value, 22, 10))

Application.ScreenUpdating = True
'
End Sub
