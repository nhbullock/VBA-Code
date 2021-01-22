Attribute VB_Name = "Module16"
Sub Post_Loads()
Attribute Post_Loads.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Post_Loads Macro
'
Dim Numloads As Integer, loads As Range, count As Long, q As String, LR As Long, LC As Long, r As Long, c As Long, pos As Long
Dim v As Variant

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Numloads = Selection.Rows.count
Set loads = Selection

For i = 1 To Numloads
q = ""
Worksheets("Protein Schedule").Cells(2, 1).EntireRow.Insert Shift:=xlShiftDown
Worksheets("Protein Schedule").Cells(2, 2).Value = loads((Numloads - i + 1), 1).Value
Worksheets("Protein Schedule").Cells(2, 1).Value = Cells(loads(Numloads - i + 1).Row, 1).Value
Worksheets("Protein Schedule").Cells(2, 7).Value = Cells(loads(Numloads - i + 1).Row, 10).Value
count = Len(Worksheets("Protein Schedule").Cells(2, 7).Value) - Len(Replace(Worksheets("Protein Schedule").Cells(2, 7).Value, ",", "")) + 1
    For t = 1 To count
        q = q & t & ","
    Next
q = Left(q, Len(q) - 1)
Worksheets("Protein Schedule").Cells(2, 3).Value = q
Next


Worksheets("Protein Schedule").Select
LR = Cells(Rows.count, 1).End(xlUp).Row
LC = Cells(1, Columns.count).End(xlToLeft).Column
r = 2
Do While r <= LR
    For c = 1 To LC
        v = Cells(r, c).Value
        If InStr(v, ",") Then Exit For ' we need to split
    Next
    If c <= LC Then ' We need to split
        Rows(r).EntireRow.Insert
        LR = LR + 1
        For c = 1 To LC
            v = Cells(r + 1, c).Value
            pos = InStr(v, ",")
            If pos Then
                Cells(r, c).Value = Left(v, pos - 1)
                Cells(r + 1, c).Value = Trim(Mid(v, pos + 1))
            Else
                Cells(r, c).Value = v
            End If
        Next
    End If
    r = r + 1
Loop

For i = 2 To LR
    If Cells(i, 7).Value = "m" Then Cells(i, 7).Value = Cells(i, 1).Value
    If Cells(i, 7).Value = "t" Then Cells(i, 7).Value = Cells(i, 1).Value + 1
    If Cells(i, 7).Value = "w" Then Cells(i, 7).Value = Cells(i, 1).Value + 2
    If Cells(i, 7).Value = "th" Then Cells(i, 7).Value = Cells(i, 1).Value + 3
    If Cells(i, 7).Value = "f" Then Cells(i, 7).Value = Cells(i, 1).Value + 4
Next

For i = 2 To 100
    If IsEmpty(Cells(i, 4)) Then Cells(i, 4).Formula = "=IFERROR(INDEX(Protein_Loads,MATCH('Protein Schedule'!B" & i & ",Contract_Range,0),4),"""")"
    If IsEmpty(Cells(i, 5)) Then Cells(i, 5).Formula = "=IFERROR(INDEX(Protein_Loads,MATCH('Protein Schedule'!B" & i & ",Contract_Range,0),5),"""")"
    If IsEmpty(Cells(i, 6)) Then Cells(i, 6).Formula = "=IFERROR(INDEX(Protein_Loads,MATCH('Protein Schedule'!B" & i & ",Contract_Range,0),6),"""")"
    If IsEmpty(Cells(i, 8)) Then Cells(i, 8).Formula = "=IFERROR(INDEX(Protein_Loads,MATCH('Protein Schedule'!B" & i & ",Contract_Range,0),11),"""")"
    If IsEmpty(Cells(i, 9)) Then Cells(i, 9).Formula = "=IFERROR(INDEX(Protein_Loads,MATCH('Protein Schedule'!B" & i & ",Contract_Range,0),12),"""")"
    If IsEmpty(Cells(i, 10)) Then Cells(i, 10).FormulaArray = "=INDEX('Protein Rates'!$E$4:$AA$35,MATCH(D" & i & "&H" & i & ",'Protein Rates'!$A$4:$A$35&'Protein Rates'!$B$4:$B$35,0),MATCH(I" & i & ",'Protein Rates'!$E$3:$AA$3,0))"
    If IsEmpty(Cells(i, 14)) Then Cells(i, 14).Formula = "=IFERROR(IF(K" & i & ">1,IF(O" & i & "<=G" & i & ",""ON TIME"",""LATE""),IF(K" & i & "=1,""CANCELLED"",IF(G" & i & "<TODAY(),""CARRYOVER"",""YES""))),"""")"
    If IsEmpty(Cells(i, 21)) Then Cells(i, 21).Formula = "=ROUND(ABS(((T" & i & "-S" & i & ")-INT((T" & i & "-S" & i & ")))*24),2)"
    If IsEmpty(Cells(i, 22)) Then Cells(i, 22).Formula = "=IFERROR(MAX(0,(U" & i & "-INDEX('Prot. Carriers'!$K:$K,MATCH(INDEX(Carriers,ROW()),'Prot. Carriers'!$B:$B,0)))*INDEX('Prot. Carriers'!$J:$J,MATCH(INDEX(Carriers,ROW()),'Prot. Carriers'!$B:$B,0))),0)"
    If IsEmpty(Cells(i, 23)) Then Cells(i, 23).Formula = "=ROUND(ABS(((M" & i & "-L" & i & "))*24),2)"
    Cells(i, 23).NumberFormat = "0.00;;"
    If IsEmpty(Cells(i, 24)) Then Cells(i, 24).Formula = "=IFERROR(MAX(0,(W" & i & "-INDEX('Prot. Carriers'!$K:$K,MATCH(INDEX(Carriers,ROW()),'Prot. Carriers'!$B:$B,0)))*INDEX('Prot. Carriers'!$J:$J,MATCH(INDEX(Carriers,ROW()),'Prot. Carriers'!$B:$B,0))),0)"
    Cells(i, 24).NumberFormat = "$#,##0.00;;"
Next

'
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
