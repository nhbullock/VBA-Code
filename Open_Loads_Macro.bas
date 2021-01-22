Attribute VB_Name = "Module15"
Sub Open_Loads()
Attribute Open_Loads.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Open_Loads Macro
'
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim t As Integer
t = 3
Dim Loads As Integer
Loads = 0

Worksheets("Open").Select
For i = 2 To 6000
    If IsEmpty(Worksheets("Protein Schedule").Cells(i, 1)) = False Then
        If IsEmpty(Worksheets("Protein Schedule").Cells(i, 9)) = True Then
        Loads = Loads + 1
        End If
    End If
Next

O = 3
For i = 3 To 6000
    If IsEmpty(Worksheets("Open").Cells(i, 1)) = False Then
        O = O + 1
        End If
Next
    
Cells(O, 1).Select
Range(ActiveCell, "I3").Select
Selection.ClearContents

For i = 2 To 6000
        If IsEmpty(Worksheets("Protein Schedule").Cells(i, 1)) = False Then
            If IsEmpty(Worksheets("Protein Schedule").Cells(i, 9)) = True Then
                Worksheets("Open").Range("A:H").Rows(t).Value = Worksheets("Protein Schedule").Range("A:H").Rows(i).Value
                Worksheets("Open").Cells(t, 9).FormulaArray = "=INDEX('Protein Schedule'!I:I,MATCH(INDEX(B:B,ROW())&INDEX(C:C,ROW()),'Protein Schedule'!B:B&'Protein Schedule'!C:C,0))"
                t = t + 1
            End If
        End If
Next

Application.ScreenUpdating = True

'
End Sub
Sub Today()
Attribute Today.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Today Macro
'

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim Today As Integer
Today = 0
Dim t As Integer
t = 3
Dim Carryover As Range
    With Worksheets("ETA").Cells
        Set Carryover = .Find("Carryovers -", After:=.Range("A2"), LookIn:=xlValues)
            If Not Carryover Is Nothing Then
                Carryover.Select
            End If
        End With
        
For i = 2 To 1000
    If Worksheets("Protein Schedule").Cells(i, 7).Value = Date Then
        Today = Today + 1
    End If
Next

For i = 1 To (Today - ActiveCell.Row) + 3
    ActiveCell.EntireRow.Insert Shift:=xlShiftDown
Next
        
ActiveCell.Offset(-1, 0).Select
Range(ActiveCell, "J3").Select
Selection.ClearContents
      
For i = 2 To 6000
    If IsEmpty(Worksheets("Protein Schedule").Cells(i, 1)) = False Then
        If Worksheets("Protein Schedule").Cells(i, 7).Value = Date Then
            Worksheets("ETA").Range("A:I").Rows(t).Value = Worksheets("Protein Schedule").Range("A:I").Rows(i).Value
            Worksheets("ETA").Cells(t, 10).FormulaArray = "=IF(INDEX('Protein Schedule'!$K:$K,MATCH(INDEX(B:B,ROW())&INDEX(C:C,ROW()),'Protein Schedule'!$B:$B&'Protein Schedule'!$C:$C,0))>0,""Loaded"","""")"
                t = t + 1
            End If
        End If
Next
              
Carryover.Offset(-1, 0).Select
Range(ActiveCell, "I2").Select
        ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("D:D"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("F:F"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("B:B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ETA").Sort
        .SetRange Selection
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Worksheets("ETA").Cells(1, 15).Value = Date

Application.ScreenUpdating = True
'
End Sub
Sub Carryover()
Attribute Carryover.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Carryover Macro
'
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim t As Integer

       Dim Carryover As Range
        With Worksheets("ETA").Cells
            Set Carryover = .Find("Carryovers -", After:=.Range("A2"), LookIn:=xlValues)
            If Not Carryover Is Nothing Then
                Carryover.Select
            End If
        End With
        
    ActiveCell.Offset(2, 0).Select
    Range("A" & ActiveCell.Row & ":J" & ActiveCell.Row).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

t = Carryover.Offset(2, 0).Row
For i = 2 To 6000
    If IsEmpty(Worksheets("Protein Schedule").Cells(i, 1)) = False Then
        If Worksheets("Protein Schedule").Cells(i, 7).Value < Date Then
            If IsEmpty(Worksheets("Protein Schedule").Cells(i, 11)) = True Then
                Worksheets("ETA").Range("A:I").Rows(t).Value = Worksheets("Protein Schedule").Range("A:I").Rows(i).Value
                Worksheets("ETA").Cells(t, 10).FormulaArray = "=IF(INDEX('Protein Schedule'!$K:$K,MATCH(INDEX(B:B,ROW())&INDEX(C:C,ROW()),'Protein Schedule'!$B:$B&'Protein Schedule'!$C:$C,0))>0,""Loaded"","""")"
                t = t + 1
            End If
        End If
    End If
Next

    Carryover.Offset(1, 0).Select
    Range("A" & ActiveCell.Row & ":I" & ActiveCell.Row).Select
    Range(Selection, Selection.End(xlDown)).Select

        ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("D:D"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("F:F"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ETA").Sort.SortFields.Add Key:=Range("B:B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ETA").Sort
        .SetRange Selection
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Application.ScreenUpdating = True

'
End Sub
