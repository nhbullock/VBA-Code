Attribute VB_Name = "Module2"
Sub Splt()
    Dim LR As Long, LC As Long, r As Long, c As Long, pos As Long
    Dim v As Variant

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    LC = Cells(1, Columns.Count).End(xlToLeft).Column
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
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
