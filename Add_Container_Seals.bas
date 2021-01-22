Attribute VB_Name = "Module2"
Sub Add_Container_Seals()
Attribute Add_Container_Seals.VB_ProcData.VB_Invoke_Func = " \n14"

' Add_Container_Seals Macro

Dim I As Integer, row As Integer

Application.ScreenUpdating = False

'Finds SO number in Shipping Details'
For I = 3 To 1000
    If Worksheets("Shipping Details").Cells(I, 1).Value = Worksheets("Container Sheet").Cells(2, 3).Value Then
    row = I
    I = 1000
    End If
Next

'Checks Container sheet for container in Column C and copies to Shipping Details'
For s = 5 To 39
    If IsEmpty(Worksheets("Container Sheet").Cells(s, 3).Value) = False Then
        Worksheets("Shipping Details").Cells(row, 2 * (s - 4) + 56).Value = Worksheets("Container Sheet").Cells(s, 2).Value & " SEAL " & Worksheets("Container Sheet").Cells(s, 4).Value
        Worksheets("Shipping Details").Cells(row, 2 * (s - 4) + 57).Value = Worksheets("Container Sheet").Cells(s, 3).Value
    End If
Next
Application.ScreenUpdating = True

End Sub

