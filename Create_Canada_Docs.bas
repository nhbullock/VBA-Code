Attribute VB_Name = "Module6"
Sub Create_Canada_Docs()
Attribute Create_Canada_Docs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_Canada_Docs Macro
'
Dim I As Integer, row As Integer
Application.ScreenUpdating = False

For I = 3 To 1000
    If Worksheets("Shipping Details").Cells(I, 1).Value = Worksheets("Forms").Cells(4, 7).Value Then
    row = I
    I = 1000
    End If
Next

'Consignee - find in Tables tab
For I = 3 To 1000
    If Worksheets("TABLES").Cells(I, 1).Value = Worksheets("Shipping Details").Cells(row, 10).Value Then
    consignee = I
    I = 1000
    End If
Next

'Notify Party - find in Tables tab
For I = 3 To 1000
    If Worksheets("TABLES").Cells(I, 1).Value = Worksheets("Shipping Details").Cells(row, 11).Value Then
    notify = I
    I = 1000
    End If
Next

'Buyer - find in Tables tab
For I = 3 To 1000
    If Worksheets("TABLES").Cells(I, 1).Value = Worksheets("Shipping Details").Cells(row, 9).Value Then
    buyer = I
    I = 1000
    End If
Next

'Shipper - find in Tables tab
For I = 3 To 1000
    If Worksheets("TABLES").Cells(I, 1).Value = Worksheets("Shipping Details").Cells(row, 12).Value Then
    shipper = I
    I = 1000
    End If
Next

Worksheets("Forms").Cells(7, 7).Value = Worksheets("Shipping Details").Cells(row, 25).Value 'Number of containers'

Worksheets("CANADA CO").Cells(5, 3).Value = Worksheets("Shipping Details").Cells(row, 12).Value 'Shipper'

Worksheets("BL INSTRUCTIONS").Cells(4, 6).Value = Worksheets("Shipping Details").Cells(row, 24).Value 'Booking number'
Worksheets("BL INSTRUCTIONS").Cells(6, 6).Value = Worksheets("Shipping Details").Cells(row, 1).Value 'Shipping Reference'
Worksheets("BL INSTRUCTIONS").Cells(24, 4).Value = Worksheets("Shipping Details").Cells(row, 29).Value 'Place of receipt'
Worksheets("BL INSTRUCTIONS").Cells(26, 1).Value = Worksheets("Shipping Details").Cells(row, 23).Value 'Vessel'
Worksheets("BL INSTRUCTIONS").Cells(26, 3).Value = Worksheets("Shipping Details").Cells(row, 26).Value 'Voyage'
Worksheets("BL INSTRUCTIONS").Cells(26, 4).Value = Worksheets("Shipping Details").Cells(row, 30).Value 'Port of Loading'
Worksheets("BL INSTRUCTIONS").Cells(28, 1).Value = Worksheets("Shipping Details").Cells(row, 33).Value 'Port of Discharge'
Worksheets("BL INSTRUCTIONS").Cells(28, 4).Value = Worksheets("Shipping Details").Cells(row, 33).Value 'Place of Delivery'
Worksheets("BL INSTRUCTIONS").Cells(14, 6).Value = "CANADA" 'Country of Origin'
Worksheets("BL INSTRUCTIONS").Cells(31, 5).Value = Worksheets("Shipping Details").Cells(row, 126).Value 'CAED vs AES'

'Consignee
Worksheets("BL INSTRUCTIONS").Cells(11, 1).Value = Worksheets("TABLES").Cells(consignee, 2).Value
Worksheets("BL INSTRUCTIONS").Cells(12, 1).Value = Worksheets("TABLES").Cells(consignee, 3).Value
Worksheets("BL INSTRUCTIONS").Cells(13, 1).Value = Worksheets("TABLES").Cells(consignee, 4).Value
Worksheets("BL INSTRUCTIONS").Cells(14, 1).Value = Worksheets("TABLES").Cells(consignee, 5).Value
Worksheets("BL INSTRUCTIONS").Cells(15, 1).Value = Worksheets("TABLES").Cells(consignee, 6).Value

'Notify Party
Worksheets("BL INSTRUCTIONS").Cells(17, 1).Value = Worksheets("TABLES").Cells(notify, 2).Value
Worksheets("BL INSTRUCTIONS").Cells(18, 1).Value = Worksheets("TABLES").Cells(notify, 3).Value
Worksheets("BL INSTRUCTIONS").Cells(19, 1).Value = Worksheets("TABLES").Cells(notify, 4).Value
Worksheets("BL INSTRUCTIONS").Cells(20, 1).Value = Worksheets("TABLES").Cells(notify, 5).Value
Worksheets("BL INSTRUCTIONS").Cells(21, 1).Value = Worksheets("TABLES").Cells(notify, 6).Value

Worksheets("BL INSTRUCTIONS").Range("A32:C62").Value = "" 'Clears contents'

'Container number and Weights'
For s = 32 To 62
    If IsEmpty(Worksheets("Shipping Details").Cells(row, 2 * (s - 31) + 56).Value) = False Then
        Worksheets("BL INSTRUCTIONS").Cells(s, 1).Value = Worksheets("Shipping Details").Cells(row, 2 * (s - 31) + 56).Value
        Worksheets("BL INSTRUCTIONS").Cells(s, 3).Value = Format(Worksheets("Shipping Details").Cells(row, 2 * (s - 31) + 57).Value / 1000, "#,##0.000")
    End If
Next

Worksheets("BL INSTRUCTIONS").Cells(32, 4).Value = Worksheets("Shipping Details").Cells(row, 25).Value & "  x  " & Worksheets("Shipping Details").Cells(row, 44).Value & " Containers  " '# of containers'
Worksheets("BL INSTRUCTIONS").Cells(35, 4).Value = Worksheets("Shipping Details").Cells(row, 7).Value 'Material'
Worksheets("BL INSTRUCTIONS").Cells(8, 12).Value = Worksheets("Shipping Details").Cells(row, 1) 'PO number'

Worksheets("CI").Cells(39, 12).Value = ""
Worksheets("CI").Cells(39, 13).Value = ""
Worksheets("CI").Cells(39, 10).Value = ""
Worksheets("CI").Cells(39, 6).Value = ""
Worksheets("CI").Cells(39, 14).Value = ""
Worksheets("CI").Cells(41, 12).Value = ""
Worksheets("CI").Cells(41, 13).Value = ""
Worksheets("CI").Cells(41, 10).Value = ""
Worksheets("CI").Cells(41, 6).Value = ""
Worksheets("CI").Cells(41, 14).Value = ""

Worksheets("CI").Cells(19, 10).Value = "CANADA" 'Country of Origin of Goods'
Worksheets("CI").Cells(19, 12).Value = "CA" 'Country Code'
Worksheets("CI").Cells(5, 10).Value = Worksheets("Shipping Details").Cells(row, 49).Value 'Export Invoice #'
Worksheets("CI").Cells(3, 11).Value = Worksheets("Shipping Details").Cells(row, 40).Value 'Bill of Lading # - Needs Booking # on FOB shipments'
Worksheets("CI").Cells(7, 10).Value = Worksheets("Shipping Details").Cells(row, 8).Value 'Buyers Reference'
Worksheets("CI").Cells(19, 13).Value = Worksheets("Shipping Details").Cells(row, 34).Value 'Country of Final Destination'
Worksheets("CI").Cells(22, 3).Value = Worksheets("Shipping Details").Cells(row, 23).Value 'Vessel'
Worksheets("CI").Cells(24, 3).Value = Worksheets("Shipping Details").Cells(row, 31).Value 'Port of Loading'
Worksheets("CI").Cells(26, 3).Value = Worksheets("Shipping Details").Cells(row, 33).Value 'Port of Discharge'
Worksheets("CI").Cells(28, 3).Value = Worksheets("Shipping Details").Cells(row, 33).Value 'Destination'
Worksheets("CI").Cells(22, 8).Value = Worksheets("Shipping Details").Cells(row, 26).Value 'Voyage Number'
Worksheets("CI").Cells(24, 8).Value = Worksheets("Shipping Details").Cells(row, 36).Value 'Departure Date'
Worksheets("CI").Cells(26, 8).Value = Worksheets("Shipping Details").Cells(row, 37).Value 'Arrival Date'
Worksheets("CI").Cells(28, 8).Value = Worksheets("Shipping Details").Cells(row, 37).Value 'Arrival Date'
Worksheets("CI").Cells(37, 6).Value = Worksheets("Shipping Details").Cells(row, 7).Value 'Description of Goods'
Worksheets("CI").Cells(70, 3).Value = "OF CANADA ORIGIN"

'Pastes CFR/FOB and Destination/Loading Port'
If Worksheets("Shipping Details").Cells(row, 18).Value = "FOB" Then
    Worksheets("CI").Cells(70, 11).Value = Worksheets("Shipping Details").Cells(row, 18).Value & "  " & Worksheets("Shipping Details").Cells(row, 31).Value
    Worksheets("CI").Cells(2, 11).Value = "Booking No"
    Else
    Worksheets("CI").Cells(70, 11).Value = Worksheets("Shipping Details").Cells(row, 18).Value & "  " & Worksheets("Shipping Details").Cells(row, 33).Value
    Worksheets("CI").Cells(2, 11).Value = "Bill of Lading No"
End If

Worksheets("CI").Cells(72, 12).Value = Worksheets("Shipping Details").Cells(row, 36).Value 'Date of Issue'
Worksheets("CI").Cells(37, 2).Value = Worksheets("Shipping Details").Cells(row, 25).Value '# of containers'
Worksheets("CI").Cells(5, 13).Value = Worksheets("Shipping Details").Cells(row, 1).Value & "   " & Worksheets("Shipping Details").Cells(row, 2).Value 'Exporter's reference'
Worksheets("CI").Cells(5, 12).Value = Worksheets("Shipping Details").Cells(row, 36).Value 'Export Invoice Date'

Worksheets("CI").Cells(32, 12).Value = Format(Worksheets("Shipping Details").Cells(row, 129).Value, "#,##0.000") 'Gross weight'
Worksheets("CI").Cells(37, 10).Value = Format(Worksheets("Shipping Details").Cells(row, 129).Value, "#,##0.000") 'Net weight'
Worksheets("CI").Cells(59, 11).Value = Format(Worksheets("Shipping Details").Cells(row, 129).Value, "#,##0.000") 'Total package weight'
Worksheets("CI").Cells(61, 11).Value = Format(Worksheets("Shipping Details").Cells(row, 129).Value, "#,##0.000") 'Total this cargo'

'Consignee
Worksheets("ci").Cells(10, 3).Value = Worksheets("TABLES").Cells(consignee, 2).Value 'Formula for Company'
Worksheets("ci").Cells(11, 3).Value = Worksheets("TABLES").Cells(consignee, 3).Value 'Formula for Address'
Worksheets("ci").Cells(12, 3).Value = Worksheets("TABLES").Cells(consignee, 4).Value 'Formula for Address'
Worksheets("ci").Cells(13, 3).Value = Worksheets("TABLES").Cells(consignee, 5).Value 'Formula for Address'
Worksheets("ci").Cells(14, 3).Value = Worksheets("TABLES").Cells(consignee, 6).Value 'Formula for Address'
'Notify Party
Worksheets("ci").Cells(16, 3).Value = Worksheets("TABLES").Cells(notify, 2).Value 'Formula for Company'
Worksheets("ci").Cells(17, 3).Value = Worksheets("TABLES").Cells(notify, 3).Value 'Formula for Address'
Worksheets("ci").Cells(18, 3).Value = Worksheets("TABLES").Cells(notify, 4).Value 'Formula for Address'
Worksheets("ci").Cells(19, 3).Value = Worksheets("TABLES").Cells(notify, 5).Value 'Formula for Address'
Worksheets("ci").Cells(20, 3).Value = Worksheets("TABLES").Cells(notify, 6).Value 'Formula for Address'
'Buyer
Worksheets("ci").Cells(10, 11).Value = Worksheets("TABLES").Cells(buyer, 2).Value 'Formula for Buyer'
Worksheets("ci").Cells(11, 11).Value = Worksheets("TABLES").Cells(buyer, 3).Value 'Formula for Address'
Worksheets("ci").Cells(12, 11).Value = Worksheets("TABLES").Cells(buyer, 4).Value 'Formula for Address'
Worksheets("ci").Cells(13, 11).Value = Worksheets("TABLES").Cells(buyer, 5).Value 'Formula for Address'
Worksheets("ci").Cells(14, 11).Value = Worksheets("TABLES").Cells(buyer, 6).Value 'Formula for Address'
Worksheets("CI").Cells(37, 12).Value = Worksheets("Shipping Details").Cells(row, 4).Value 'Unit price'
Worksheets("CI").Cells(31, 12).Value = "Gross Weight (MT)"
Worksheets("CI").Cells(35, 10).Value = "MT"
Worksheets("CI").Cells(37, 13).Value = "USD/MT"
Worksheets("CI").Cells(37, 12).Value = Worksheets("Shipping Details").Cells(row, 4).Value 'Contract price'
Worksheets("CI").Cells(18, 21).Value = Worksheets("Shipping Details").Cells(row, 1).Value 'Contract number'

'Consignee
Worksheets("co").Cells(18, 3).Value = Worksheets("TABLES").Cells(consignee, 2).Value 'Formula for Consignee'
Worksheets("co").Cells(19, 3).Value = Worksheets("TABLES").Cells(consignee, 3).Value 'Formula for Address'
Worksheets("co").Cells(20, 3).Value = Worksheets("TABLES").Cells(consignee, 4).Value 'Formula for Address'
Worksheets("co").Cells(21, 3).Value = Worksheets("TABLES").Cells(consignee, 5).Value 'Formula for Address'
Worksheets("co").Cells(22, 3).Value = Worksheets("TABLES").Cells(consignee, 6).Value 'Formula for Address'

Worksheets("co").Cells(25, 3).Value = Worksheets("Shipping Details").Cells(row, 41).Value 'Booking number'
Worksheets("co").Cells(26, 3).Value = Worksheets("Shipping Details").Cells(row, 136).Value 'Export Certificate'
Worksheets("co").Cells(29, 3).Value = Worksheets("Shipping Details").Cells(row, 7).Value 'Description of Goods'
Worksheets("co").Cells(9, 3).Value = Worksheets("Shipping Details").Cells(row, 36).Value 'Date'

Worksheets("CO").Range("B34:H53").Value = ""
Worksheets("CO").Range("B95:H114").Value = ""

'Container numbers and weights'
For s = 34 To 53 'Page 1'
    If IsEmpty(Worksheets("Shipping Details").Cells(row, 2 * (s - 33) + 56).Value) = False Then
        Worksheets("CO").Cells(s, 2).Value = Worksheets("Shipping Details").Cells(row, 2 * (s - 33) + 56).Value
        Worksheets("CO").Cells(s, 4).Value = Format(Worksheets("Shipping Details").Cells(row, 2 * (s - 33) + 57).Value / 1000, "#,##0.000")
        Worksheets("CO").Cells(s, 8).Value = "MT"
    End If
Next
For s = 95 To 109 'Page 2'
    If IsEmpty(Worksheets("Shipping Details").Cells(row, 2 * (s - 94) + 56).Value) = False Then
        Worksheets("CO").Cells(s, 2).Value = Worksheets("Shipping Details").Cells(row, 2 * (s - 94) + 96).Value
        Worksheets("CO").Cells(s, 4).Value = Format(Worksheets("Shipping Details").Cells(row, 2 * (s - 94) + 97).Value / 1000, "#,##0.000")
        Worksheets("CO").Cells(s, 8).Value = "MT"
    End If
Next
 
Worksheets("co").Cells(8, 14).Value = Worksheets("Shipping Details").Cells(row, 1).Value 'PO Number'
Worksheets("co").Cells(124, 13).Value = Worksheets("Shipping Details").Cells(row, 44).Value 'Container size - posted on side'

Worksheets("PACKING LIST").Cells(5, 2).Value = Worksheets("TABLES").Cells(consignee, 2).Value 'Consignee'
Worksheets("PACKING LIST").Cells(6, 2).Value = Worksheets("TABLES").Cells(consignee, 3).Value 'Address'
Worksheets("PACKING LIST").Cells(7, 2).Value = Worksheets("TABLES").Cells(consignee, 4).Value 'Address'
Worksheets("PACKING LIST").Cells(8, 2).Value = Worksheets("TABLES").Cells(consignee, 5).Value 'Address'
Worksheets("PACKING LIST").Cells(9, 2).Value = Worksheets("TABLES").Cells(consignee, 6).Value 'Address - pastes tax ID'
Worksheets("PACKING LIST").Cells(4, 7).Value = Worksheets("co").Cells(9, 3).Value 'Date'
Worksheets("PACKING LIST").Cells(85, 12).Value = Worksheets("Shipping Details").Cells(row, 44).Value 'Container size'
Worksheets("PACKING LIST").Cells(12, 4).Value = Worksheets("Shipping Details").Cells(row, 7).Value 'Description of Goods'

'Copies Description of Goods & Containers from CO to Packing List and WC'
Worksheets("PACKING LIST").Range("B12:H36").Value = Worksheets("CO").Range("B29:H53").Value
Worksheets("WC").Range("B12:H36").Value = Worksheets("CO").Range("B29:H53").Value

'Copies page 2'
Worksheets("PACKING LIST").Range("B61:H81").Value = Worksheets("CO").Range("B94:H114").Value
Worksheets("WC").Range("B61:H81").Value = Worksheets("CO").Range("B94:H114").Value

Worksheets("WC").Cells(5, 2).Value = Worksheets("TABLES").Cells(consignee, 2).Value 'Consignee'
Worksheets("WC").Cells(6, 2).Value = Worksheets("TABLES").Cells(consignee, 3).Value 'Address'
Worksheets("WC").Cells(7, 2).Value = Worksheets("TABLES").Cells(consignee, 4).Value 'Address'
Worksheets("WC").Cells(8, 2).Value = Worksheets("TABLES").Cells(consignee, 5).Value 'Address'
Worksheets("WC").Cells(9, 2).Value = Worksheets("TABLES").Cells(consignee, 6).Value 'Address - pastes tax ID'
Worksheets("WC").Cells(4, 7).Value = Worksheets("co").Cells(9, 3).Value 'Date'
Worksheets("WC").Cells(4, 8).Value = Worksheets("Shipping Details").Cells(row, 1) 'PO number'
Worksheets("WC").Cells(86, 12).Value = Worksheets("Shipping Details").Cells(row, 44).Value 'Container size'


Worksheets("PD").Cells(9, 1).Value = "Ship Name: " & " " & Worksheets("Shipping Details").Cells(row, 23).Value & " " & Worksheets("Shipping Details").Cells(row, 26).Value 'Ship Name'
Worksheets("PD").Cells(11, 3).Value = Worksheets("Shipping Details").Cells(row, 40).Value 'Bill of Lading number'
Worksheets("PD").Cells(46, 2).Value = "GREELEY, CO " & Worksheets("Shipping Details").Cells(row, 36).Value 'Date' 'Greeley and Date from CO'
Worksheets("PD").Cells(44, 7).Value = Format(Worksheets("Shipping Details").Cells(row, 129).Value, "#,##0.000") & "  MT" 'Net weight from BL instructions'
Worksheets("PD").Cells(8, 14).Value = Worksheets("Shipping Details").Cells(row, 1) 'Sales Order number'

Worksheets("CO").Cells(18, 2).Value = "CONSIGNEE:"
Worksheets("WC").Cells(4, 2).Value = "Consignee"
Worksheets("Packing List").Cells(4, 2).Value = "Consignee"

Worksheets("CANADA CO").Cells(26, 3).Value = "BULK IN " & Worksheets("Shipping Details").Cells(row, 44).Value & " FCL CONTAINERS"
Worksheets("CANADA CO").Cells(27, 3).Value = Worksheets("Shipping Details").Cells(row, 25).Value & " FCL"

Worksheets("CANADA CO").Cells(5, 3).Value = Worksheets("TABLES").Cells(shipper, 2).Value 'Shipper'
Worksheets("CANADA CO").Cells(6, 3).Value = Worksheets("TABLES").Cells(shipper, 3).Value 'Address'
Worksheets("CANADA CO").Cells(7, 3).Value = Worksheets("TABLES").Cells(shipper, 4).Value 'Address'
Worksheets("CANADA CO").Cells(8, 3).Value = Worksheets("TABLES").Cells(shipper, 5).Value 'Address'

Worksheets("CANADA CO").Cells(11, 3).Value = Worksheets("TABLES").Cells(consignee, 2).Value 'Consignee'
Worksheets("CANADA CO").Cells(12, 3).Value = Worksheets("TABLES").Cells(consignee, 3).Value 'Address'
Worksheets("CANADA CO").Cells(13, 3).Value = Worksheets("TABLES").Cells(consignee, 4).Value 'Address'
Worksheets("CANADA CO").Cells(14, 3).Value = Worksheets("TABLES").Cells(consignee, 5).Value 'Address'
Worksheets("CANADA CO").Cells(15, 3).Value = Worksheets("TABLES").Cells(consignee, 6).Value 'Address - pastes tax ID'

Worksheets("CANADA CO").Cells(23, 3).Value = Worksheets("Shipping Details").Cells(row, 23).Value & "  " & Worksheets("Shipping Details").Cells(row, 26).Value 'Vessel'
Worksheets("CANADA CO").Cells(28, 3).Value = Worksheets("Shipping Details").Cells(row, 7).Value 'Material'
Worksheets("CANADA CO").Cells(26, 5).Value = Worksheets("CANADA CO").Cells(27, 3).Value 'Quantity'
Worksheets("CANADA CO").Cells(26, 6).Value = Worksheets("Shipping Details").Cells(row, 129).Value 'Total container weight'

'Signatures'
Worksheets("CI").Cells(76, 11).Value = Worksheets("Forms").Cells(4, 17).Value
Worksheets("WC").Cells(38, 2).Value = "Name of Authorised Signatory: " & Worksheets("Forms").Cells(4, 17).Value
Worksheets("WC").Cells(84, 2).Value = "Name of Authorised Signatory: " & Worksheets("Forms").Cells(4, 17).Value
Worksheets("PACKING LIST").Cells(38, 2).Value = "Name of Authorised Signatory:  " & Worksheets("Forms").Cells(4, 17).Value
Worksheets("PACKING LIST").Cells(83, 2).Value = "Name of Authorised Signatory:  " & Worksheets("Forms").Cells(4, 17).Value
Worksheets("PD").Cells(46, 6).Value = Worksheets("Forms").Cells(4, 17).Value

'Set print area based on # of pages'
If IsEmpty(Worksheets("Shipping Details").Cells(row, 98).Value) = False Then
    Worksheets("WC").PageSetup.PrintArea = "$A$1:$I$89"
    Worksheets("PACKING LIST").PageSetup.PrintArea = "$A$1:$I$88"
    Worksheets("CO").PageSetup.PrintArea = "$A$1:$I$130"
    Else
    Worksheets("WC").PageSetup.PrintArea = "$A$1:$I$45"
    Worksheets("PACKING LIST").PageSetup.PrintArea = "$A$1:$I$45"
    Worksheets("CO").PageSetup.PrintArea = "$A$1:$I$63"
End If

'Set print area based on # of pages'
If IsEmpty(Worksheets("Shipping Details").Cells(row, 98).Value) = False Then
    Worksheets("WC").PageSetup.PrintArea = "$A$1:$I$89"
    Worksheets("PACKING LIST").PageSetup.PrintArea = "$A$1:$I$88"
    Worksheets("CO").PageSetup.PrintArea = "$A$3:$I$61,A65:I121"
    Else
    Worksheets("WC").PageSetup.PrintArea = "$A$1:$I$45"
    Worksheets("PACKING LIST").PageSetup.PrintArea = "$A$1:$I$45"
    Worksheets("CO").PageSetup.PrintArea = "$A$3:$I$61"
End If

Sheets("Forms").Select
Application.ScreenUpdating = True
        
End Sub
