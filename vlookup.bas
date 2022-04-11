Attribute VB_Name = "Module1"
'Import data to table for Par Preview
Sub importparview()
    On Error GoTo Oops
    
    Dim openImport As Variant
    Dim LRow As Long, thisRow As Long
    Dim myCell As Range, thisCell As Range, i As Long
    ConditionUpload = False
    
    i = 0
    thisRow = 7
    Set myCell = ThisWorkbook.Sheets(1).Range("B" & thisRow & ":K" & thisRow)
    Do While i < 10
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisRow = thisRow + 1
        Set myCell = ThisWorkbook.Sheets(1).Range("B" & thisRow & ":K" & thisRow)
    Loop
    Set myCell = ThisWorkbook.Sheets(1).Range("B" & thisRow - 1 & ":K" & thisRow - 1)
    openImport = Application.GetOpenFilename(Title:="Pilih file yang akan diimport", fileFilter:="Excel Files (*.xls*),")
    If openImport <> False Then
        'cek apakah sudah dibuka untuk filenya
        Do While Application.ProtectedViewWindows.Count > 0
            Application.ProtectedViewWindows(1).Edit
        Loop
        Set openImport = Application.Workbooks.Open(openImport)
        
        LRow = openImport.Sheets(1).Cells(openImport.Sheets(1).Rows.Count, 2).End(xlUp).Row
        openImport.Sheets(1).Range("A2:k" & LRow).copy
        myCell.PasteSpecial
        openImport.Close
        
        'Create table line
        Dim tblExists As Boolean
        Dim rangeTable As Long
        
        rangeTable = ThisWorkbook.Sheets(1).Range("B" & Rows.Count).End(xlUp).Row
        tblExists = False
        
        Dim o As ListObject
        For Each o In Sheets(1).ListObjects
          If o.Name = "myTable1" Then tblExists = True
        Next o
        
        If tblExists Then
          Sheets(1).ListObjects("myTable1").Unlist
        End If
        ThisWorkbook.Sheets(1).Unprotect Password:="pass"
        ThisWorkbook.Sheets(1).Cells.Locked = False
        ThisWorkbook.Sheets(1).Range("A6:K" & thisRow - 1).Locked = False
        Sheet1.ListObjects.Add(xlSrcRange, Range("B6:L" & rangeTable), , xlYes).Name = "myTable1"
        
        Set myCell = ThisWorkbook.Sheets(1).Range("C6:K" & rangeTable)
        Dim cek As String
        cek = myCell.Address
        rangeTable = ThisWorkbook.Sheets(1).Range("B" & Rows.Count).End(xlUp).Row
                
        If IsEmpty(ThisWorkbook.Sheets(1).Range("T7")) Then
            ConditionTable = False
        Else
            Dim s As String
            s = ThisWorkbook.Sheets(1).Range("T7").Value
            ConditionTable = True
        End If
        
        'Autonumber

        Sheets(1).Range("A7").Formula = "=Row(A1)"
        Sheets(1).Range("A7").AutoFill Range("A7:A" & rangeTable)
        Sheets(1).Range("A7:A" & rangeTable).copy
        Sheets(1).Range("A7:A" & rangeTable).PasteSpecial Paste:=xlPasteValues
        
        ConditionUpload = True
    End If
Oops:
    'handle error here
End Sub
Sub importDataparvi()
    On Error GoTo Oops
    
    Dim openImport As Variant
    Dim LRow As Long, thisRow As Long
    Dim myCell As Range, thisCell As Range, i As Long
    ConditionUpload = False
    
    i = 0
    thisRow = 2
    Set myCell = ThisWorkbook.Sheets(2).Range("B" & thisRow & ":L" & thisRow)
    Do While i < 10
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisRow = thisRow + 1
        Set myCell = ThisWorkbook.Sheets(2).Range("B" & thisRow & ":L" & thisRow)
    Loop
    Set myCell = ThisWorkbook.Sheets(2).Range("B" & thisRow - 1 & ":L" & thisRow - 1)
    openImport = Application.GetOpenFilename(Title:="Pilih file yang akan diimport", fileFilter:="Excel Files (*.xls*),")
    If openImport <> False Then
        'cek apakah sudah dibuka untuk filenya
        Do While Application.ProtectedViewWindows.Count > 0
            Application.ProtectedViewWindows(1).Edit
        Loop
        Set openImport = Application.Workbooks.Open(openImport)
        
        LRow = openImport.Sheets(1).Cells(openImport.Sheets(1).Rows.Count, 2).End(xlUp).Row
        openImport.Sheets(1).Range("A2:M" & LRow).copy
        myCell.PasteSpecial
        openImport.Close
        
        'Create table line
        Dim tblExists As Boolean
        Dim rangeTable As Long
        
        rangeTable = ThisWorkbook.Sheets(2).Range("B" & Rows.Count).End(xlUp).Row
        tblExists = False
        
        Dim o As ListObject
        For Each o In Sheets(2).ListObjects
          If o.Name = "myTable1" Then tblExists = True
        Next o
        
        If tblExists Then
          Sheets(2).ListObjects("myTable1").Unlist
        End If
        ThisWorkbook.Sheets(2).Unprotect Password:="pass"
        ThisWorkbook.Sheets(2).Cells.Locked = False
        ThisWorkbook.Sheets(2).Range("A1:M" & thisRow - 1).Locked = False
        Sheet2.ListObjects.Add(xlSrcRange, Range("B1:M" & rangeTable), , xlYes).Name = "myTable1"
        
        Set myCell = ThisWorkbook.Sheets(2).Range("C1:M" & rangeTable)
        Dim cek As String
        cek = myCell.Address
        rangeTable = ThisWorkbook.Sheets(2).Range("B" & Rows.Count).End(xlUp).Row
                
        'Autonumber
        Sheets(2).Range("A2").Formula = "=Row(A1)"
        Sheets(2).Range("A2").AutoFill Range("A2:A" & rangeTable)
        Sheets(2).Range("A2:A" & rangeTable).copy
        Sheets(2).Range("A2:A" & rangeTable).PasteSpecial Paste:=xlPasteValues
        Sheets(2).Range("A1").Select
        
        ConditionUpload = True
    End If
Oops:
    'handle error here
End Sub


Sub importDataemail()
    On Error GoTo Oops
    
    Dim openImport As Variant
    Dim LRow As Long, thisRow As Long
    Dim myCell As Range, thisCell As Range, i As Long
    ConditionUpload = False
    
    i = 0
    thisRow = 2
    Set myCell = ThisWorkbook.Sheets(3).Range("B" & thisRow & ":K" & thisRow)
    Do While i < 10
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisRow = thisRow + 1
        Set myCell = ThisWorkbook.Sheets(3).Range("B" & thisRow & ":J" & thisRow)
    Loop
    Set myCell = ThisWorkbook.Sheets(3).Range("B" & thisRow - 1 & ":J" & thisRow - 1)
    openImport = Application.GetOpenFilename(Title:="Pilih file yang akan diimport", fileFilter:="Excel Files (*.xls*),")
    If openImport <> False Then
        'cek apakah sudah dibuka untuk filenya
        Do While Application.ProtectedViewWindows.Count > 0
            Application.ProtectedViewWindows(1).Edit
        Loop
        Set openImport = Application.Workbooks.Open(openImport)
        
        LRow = openImport.Sheets(1).Cells(openImport.Sheets(1).Rows.Count, 2).End(xlUp).Row
        openImport.Sheets(1).Range("A2:J" & LRow).copy
        myCell.PasteSpecial
        openImport.Close
        
        'Create table line
        Dim tblExists As Boolean
        Dim rangeTable As Long
        
        rangeTable = ThisWorkbook.Sheets(3).Range("B" & Rows.Count).End(xlUp).Row
        tblExists = False
        
        Dim o As ListObject
        For Each o In Sheets(3).ListObjects
          If o.Name = "myTable1" Then tblExists = True
        Next o
        
        If tblExists Then
          Sheets(3).ListObjects("myTable1").Unlist
        End If
        ThisWorkbook.Sheets(3).Unprotect Password:="pass"
        ThisWorkbook.Sheets(3).Cells.Locked = False
        ThisWorkbook.Sheets(3).Range("A1:L" & thisRow - 1).Locked = False
        Sheet3.ListObjects.Add(xlSrcRange, Range("B1:K" & rangeTable), , xlYes).Name = "myTable1"
        
        Set myCell = ThisWorkbook.Sheets(3).Range("C1:K" & rangeTable)
        Dim cek As String
        cek = myCell.Address
        rangeTable = ThisWorkbook.Sheets(3).Range("B" & Rows.Count).End(xlUp).Row
                
        'Autonumber
        Sheets(3).Range("A2").Formula = "=Row(A1)"
        Sheets(3).Range("A2").AutoFill Range("A2:A" & rangeTable)
        Sheets(3).Range("A2:A" & rangeTable).copy
        Sheets(3).Range("A2:A" & rangeTable).PasteSpecial Paste:=xlPasteValues
        Sheets(3).Range("A1").Select
        
        ConditionUpload = True
    End If
Oops:
    'handle error here
End Sub

Sub clearsheet1()
On Error GoTo Oops
    Dim rangeTable As Long
    Dim LRow As Long, DupliRow As Long, tblExists As Boolean
    rangeTable = ThisWorkbook.Sheets(1).Range("B" & Rows.Count).End(xlUp).Row
    Dim o As ListObject
    For Each o In Sheets(1).ListObjects
      If o.Name = "myTable1" Then tblExists = True
    Next o
    
    If tblExists Then
      Sheets(1).ListObjects("myTable1").Unlist
    End If
    Sheets(1).Range("A7:N" & rangeTable).delete
    Sheet1.ListObjects.Add(xlSrcRange, Range("B6:N6"), , xlYes).Name = "myTable1"
Oops:

End Sub

Sub clearsheet2()
On Error GoTo Oops
    Dim rangeTable As Long
    Dim LRow As Long, DupliRow As Long, tblExists As Boolean
    rangeTable = ThisWorkbook.Sheets(2).Range("B" & Rows.Count).End(xlUp).Row
    Dim o As ListObject
    For Each o In Sheets(2).ListObjects
      If o.Name = "myTable1" Then tblExists = True
    Next o
    
    If tblExists Then
      Sheets(2).ListObjects("myTable1").Unlist
    End If
    Sheets(2).Range("A2:M" & rangeTable).delete
    Sheet2.ListObjects.Add(xlSrcRange, Range("A1:M1"), , xlYes).Name = "myTable1"
Oops:

End Sub

Sub clearsheet3()
On Error GoTo Oops
    Dim rangeTable As Long
    Dim LRow As Long, DupliRow As Long, tblExists As Boolean
    rangeTable = ThisWorkbook.Sheets(3).Range("B" & Rows.Count).End(xlUp).Row
    Dim o As ListObject
    For Each o In Sheets(3).ListObjects
      If o.Name = "myTable1" Then tblExists = True
    Next o
    
    If tblExists Then
      Sheets(3).ListObjects("myTable1").Unlist
    End If
    Sheets(3).Range("A2:K" & rangeTable).delete
    Sheet3.ListObjects.Add(xlSrcRange, Range("B1:K1"), , xlYes).Name = "myTable1"
Oops:

End Sub

Sub clearsheet4()
On Error GoTo Oops
    Dim rangeTable As Long
    Dim LRow As Long, DupliRow As Long, tblExists As Boolean
    rangeTable = ThisWorkbook.Sheets(4).Range("B" & Rows.Count).End(xlUp).Row
    Dim o As ListObject
    For Each o In Sheets(4).ListObjects
      If o.Name = "myTable1" Then tblExists = True
    Next o
    
    If tblExists Then
      Sheets(4).ListObjects("myTable1").Unlist
    End If
    Sheets(4).Range("A2:N" & rangeTable).delete
    Sheet1.ListObjects.Add(xlSrcRange, Range("A1:N1"), , xlYes).Name = "myTable1"
Oops:

End Sub

Sub chekdatasheet2()
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    Dim count_row As Long
    Dim count_col As Long
    Dim y As Range
    Dim z As Range
    
    
    Set a = ThisWorkbook.Worksheets("Par-VI")
    Set b = ThisWorkbook.Worksheets("PREVIEW")
    
    c = a.Range("B" & Rows.Count).End(xlUp).Row
    d = b.Range("B" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("B2:J" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("F" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 5, False)
    a.Range("G" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 6, False)
    a.Range("J" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 9, False)
    a.Range("M" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 1, False)
    Next x
    
    Sheets("Par-VI").Activate
    count_col = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlToRight)))
    count_row = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))
    
    Set y = Range(Cells(1, 13), Cells(count_row, 13))
    For Each z In y.Cells
    
        If z = "" Then
        z = "Open KCP"
        End If
    Next z
    adjustBlank
End Sub


Sub adjustBlank()
    thisRow = ThisWorkbook.Sheets("Par-VI").Cells(ThisWorkbook.Sheets("Par-VI").Rows.Count, 3).End(xlUp).Row
    For i = 1 To thisRow
        If Sheets("Par-VI").Range("M" & i).Value = "Open KCP" Then

            a = a
            ThisWorkbook.Sheets("Par-VI").Range("B" & i & ":M" & i).copy
            previewrow = ThisWorkbook.Sheets("PREVIEW").Cells(ThisWorkbook.Sheets("PREVIEW").Rows.Count, 3).End(xlUp).Row + 1
            ThisWorkbook.Sheets("PREVIEW").Range("B" & previewrow & ":M" & previewrow).PasteSpecial
            Dim num As Integer
            num = ThisWorkbook.Sheets("PREVIEW").Range("A" & previewrow - 1).Value
            ThisWorkbook.Sheets("PREVIEW").Range("A" & previewrow).Value = num + 1
        End If
    Next i
End Sub

Sub chekdataVal()
'copy data
    Dim wsSource, wsExtract As Worksheet
    
    Set wsSource = ThisWorkbook.Sheets("PREVIEW")
    Set wsExtract = ThisWorkbook.Sheets("Data_Val")

    wsSource.Range("A6:N5000").SpecialCells(xlCellTypeVisible).copy
    wsExtract.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
'chek validasi
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("DATABI")
    
    c = a.Range("C" & Rows.Count).End(xlUp).Row
    d = b.Range("C" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("C2:K" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("N" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 1, False)
    Next x
    
    Dim ws As Worksheet
    Set ws = Worksheets("Data_Val")

    For Each z In ws.Range("N1:N5000")
        If z.Value = "" Then z.Value = "Delete"
    Next
'delete data select

End Sub

Sub validasi_update()
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range

    
    Columns("x:x").Select
    Selection.NumberFormat = "General"
    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("DATABI")
    
    c = a.Range("C" & Rows.Count).End(xlUp).Row
    d = b.Range("C" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("C2:K" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("M" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 8, False)
    a.Range("N" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 9, False)
    Next x

End Sub


Sub Del_final()
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    
    Set a = ThisWorkbook.Worksheets("PREVIEW")
    Set b = ThisWorkbook.Worksheets("Data_Val")
    
    c = a.Range("B" & Rows.Count).End(xlUp).Row
    d = b.Range("B" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("B2:J" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("N" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 1, False)
    Next x

    Dim ws As Worksheet
    Set ws = Worksheets("PREVIEW")
    Dim z As Range

    For Each z In ws.Range("N7:N10000")
        If z.Value <> "" Then z.Value = "Delete"
    Next
    Dim i As Long
    For i = 2 To Sheets(1).UsedRange.Rows.Count
        If Cells(i, 14).Value = "Delete" Then
            Rows(i).EntireRow.delete
        End If
    Next i
    
End Sub

Sub MoveParview()
    thisRow = ThisWorkbook.Sheets("Data_Val").Cells(ThisWorkbook.Sheets("Data_Val").Rows.Count, 3).End(xlUp).Row
    For i = 2 To thisRow
        If Sheets("Data_Val").Range("N" & i).Value <> "" Then

            a = a
            ThisWorkbook.Sheets("Data_Val").Range("B" & i & ":M" & i).copy
            previewrow = ThisWorkbook.Sheets("PREVIEW").Cells(ThisWorkbook.Sheets("PREVIEW").Rows.Count, 3).End(xlUp).Row + 1
            ThisWorkbook.Sheets("PREVIEW").Range("B" & previewrow & ":M" & previewrow).PasteSpecial
            Dim num As Integer
            num = ThisWorkbook.Sheets("PREVIEW").Range("A" & previewrow - 1).Value
            ThisWorkbook.Sheets("PREVIEW").Range("A" & previewrow).Value = num + 1
        End If
    Next i
End Sub

Sub vlookupBI()
Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("DATABI")
    
    c = a.Range("C" & Rows.Count).End(xlUp).Row
    d = b.Range("C" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("C2:K" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("X" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 7, False)
    a.Range("Y" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 4, False)
    a.Range("AN" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 3, False)
    a.Range("AB" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 1, False)
    a.Range("AC" & x).Value = Application.WorksheetFunction.vlookup(a.Range("B" & x).Value, datarange, 4, False)
    Next x
    
End Sub


Sub vlookuppenutupan1()
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("PREVIEW")
    
    c = a.Range("Y" & Rows.Count).End(xlUp).Row
    d = b.Range("C" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("C7:M" & d)
    
    For x = 2 To c
    On Error Resume Next
    'a.Range("AE" & x).Value = Application.WorksheetFunction.VLookup(a.Range("AD" & x).Value, datarange, 3, False)
    a.Range("Z" & x).Value = Application.WorksheetFunction.vlookup(a.Range("Y" & x).Value, datarange, 3, False)
    Next x
    
End Sub

Sub vlookuppenutupan2()
    Dim a As Worksheet, b As Worksheet, e As Worksheet
    Dim c As Long, d As Long, x As Long, f As Long
    Dim datarange As Range, datarange2 As Range

    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("PREVIEW")

    
    c = a.Range("Z" & Rows.Count).End(xlUp).Row
    d = b.Range("E" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("E7:M" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("AA" & x).Value = Application.WorksheetFunction.vlookup(a.Range("Z" & x).Value, datarange, 8, False)
    Next x
    
End Sub

Sub vlookuprelokasi()
    Dim a As Worksheet, b As Worksheet
    Dim c As Long, d As Long, x As Long
    Dim datarange As Range
    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("Dati")
    
    c = a.Range("AN" & Rows.Count).End(xlUp).Row
    d = b.Range("B" & Rows.Count).End(xlUp).Row
    
    Set datarange = b.Range("B2:M" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("AO" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AN" & x).Value, datarange, 2, False)
    Next x
    
End Sub

Sub del_val()
Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("Data_Val")
  ws.Activate
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0
  ws.Range("A2:XFD5000").AutoFilter Field:=14, Criteria1:="Delete"
  Application.DisplayAlerts = False
    ws.Range("A2:XFD5000").SpecialCells(xlCellTypeVisible).delete
  Application.DisplayAlerts = True

  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0
End Sub

Sub val_update()
    Dim ws As Worksheet
    Dim y As Long
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Sheets("Data_Val")
    
    For y = 2 To ws.Cells(Rows.Count, "N").End(xlUp).Row
    If StrConv(ws.Range("N" & y), vbProperCase) = "Relokasi" Then
        ws.Range("K" & y).Value = ws.Range("AO" & y).Value
    ElseIf StrConv(ws.Range("N" & y), vbProperCase) = "Penutupan" Then
        ws.Range("D" & y).Value = ws.Range("X" & y).Value
        ws.Range("E" & y).Value = ws.Range("Z" & y).Value
        ws.Range("L" & y).Value = ws.Range("AA" & y).Value
    ElseIf StrConv(ws.Range("N" & y), vbProperCase) = "Pembukaan" Then
        ws.Range("D" & y).Value = ws.Range("AB" & y).Value
        ws.Range("E" & y).Value = ws.Range("AD" & y).Value
        ws.Range("F" & y).Value = ws.Range("AE" & y).Value
        ws.Range("G" & y).Value = ws.Range("AF" & y).Value
        ws.Range("H" & y).Value = ws.Range("AG" & y).Value
        ws.Range("I" & y).Value = ws.Range("AH" & y).Value
        ws.Range("J" & y).Value = ws.Range("AI" & y).Value
        ws.Range("K" & y).Value = ws.Range("AJ" & y).Value
        ws.Range("L" & y).Value = ws.Range("AK" & y).Value
    End If
    Next y
    Application.ScreenUpdating = True
        
End Sub

Sub vlookuppembukaan1()
    Dim a As Worksheet, b As Worksheet, e As Worksheet
    Dim c As Long, d As Long, x As Long, f As Long
    Dim datarange As Range, datarange2 As Range

    
    Set a = ThisWorkbook.Worksheets("Data_Val")
    Set b = ThisWorkbook.Worksheets("PREVIEW")

    
    c = a.Range("AC" & Rows.Count).End(xlUp).Row
    d = b.Range("C" & Rows.Count).End(xlUp).Row

    Set datarange = b.Range("C7:N" & d)
    
    For x = 2 To c
    On Error Resume Next
    a.Range("AD" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 3, False)
    a.Range("AE" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 4, False)
    a.Range("AF" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 5, False)
    a.Range("AG" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 6, False)
    a.Range("AH" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 7, False)
    a.Range("AI" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 8, False)
    a.Range("AJ" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 9, False)
    a.Range("AK" & x).Value = Application.WorksheetFunction.vlookup(a.Range("AC" & x).Value, datarange, 10, False)
    Next x
    
End Sub
Sub vlookup()
vlookupBI
vlookuppembukaan1
vlookuprelokasi
vlookuppenutupan1
vlookuppenutupan2
End Sub

Sub prosses()
clearsheet4
chekdatasheet2
chekdataVal
validasi_update
vlookup
del_val
val_update
Del_final
MoveParview
MsgBox ("Validasi Done...1")

End Sub




