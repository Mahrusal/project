Attribute VB_Name = "HOLDER"

Sub clear_Rencana_Penarikan_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Rencana Penarikan - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Rencana Penarikan - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub

Sub clear_Data_Pokok_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Data Pokok - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Data Pokok - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub

Sub clear_Rencana_Pembayaran_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Rencana Pembayaran - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Rencana Pembayaran - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub

Sub clear_Realisasi_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Realisasi - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Realisasi - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub

Sub clear_Posisi_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Posisi - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Posisi - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub


Sub clear_Standstill_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Standstill - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Standstill - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub


Sub clear_Pengarsipan_BI()

    row_number = 0
    thisrow = 2
    i = 0
    Set myCell = ThisWorkbook.Sheets("Pengarsipan - BI").Range("A" & thisrow & ":BE" & thisrow)
    Do While i < 57
        i = 0
        For Each thisCell In myCell
            If IsEmpty(thisCell) Then
                i = i + 1
            End If
        Next thisCell
        thisrow = thisrow + 1
        Set myCell = ThisWorkbook.Sheets("Pengarsipan - BI").Range("A" & thisrow & ":BE" & thisrow)
    Loop
    
    Range("A2:BE" & thisrow).delete
End Sub


Sub Calc()
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
'delete
'Sheets("agan01").Cells.clear

    '---- Open filedialog
    With fDialog
        .AllowMultiSelect = True
        .Title = "Please select the files"
        .Filters.Clear
        .Filters.Add "All supported files", "*.txt"
        .Filters.Add "Text Files", "*.txt"

        If .Show = True Then
            Dim fPath As Variant
            Dim thisrow1 As Integer
            Dim array_num As Integer
            Dim gettindex As Integer
            i = 0
            
            Dim myCell As Range
            array_num = 0
            getindex = 0
            ReDim Month_arr(0) As Variant
            ReDim Count_arr(0) As Variant
            ReDim Sum_arr(0) As Variant
            
            Dim Month_tmp As String
            
            For Each fPath In .SelectedItems
                Open fPath For Input As #1
                
                thisrow1 = 2
                row_number = 0
                
                '----- Dapetin baris terakhir dari table
                Set myCell = ThisWorkbook.Sheets("pbl01").Range("A" & thisrow1 & ":C" & thisrow1)
                Do While i = 0
                    For Each thisCell In myCell
                        If IsEmpty(thisCell) Then
                            i = i + 1
                        End If
                    Next thisCell
                    thisrow1 = thisrow1 + 1
                    Set myCell = ThisWorkbook.Sheets("pbl01").Range("A" & thisrow1 & ":C" & thisrow1)
                Loop
                
                Do Until EOF(1)
                    If row_number = 0 Then
                        'Baris pertama
                        row_number = 1
                        thisrow1 = thisrow1 - 1
                        Line Input #1, LineFromFile
                        lineitems = Split(LineFromFile, "|")
                        
                        Month_tmp = Left(lineitems(2), 7)
                        array_num = WorksheetFunction.CountA(Month_arr)
                        
                        ReDim Preserve Month_arr(array_num)
                        ReDim Preserve Count_arr(array_num)
                        ReDim Preserve Sum_arr(array_num)
                        
                        getindex = array_num
                        Month_arr(array_num) = Left(lineitems(2), 7)
                        Count_arr(array_num) = 1
                        Sum_arr(array_num) = CDbl(lineitems(26))
                    ElseIf row_number = 1 Then
                        Line Input #1, LineFromFile
                        lineitems = Split(LineFromFile, "|")
                        
                        'Cek apakah line yang berisi bulan nya sama seperti line sebelumnya
                        If Not Left(lineitems(2), 7) = Month_tmp Then
                            Month_tmp = Left(lineitems(2), 7)
                            'cek apakah month yang di line ini sudah ada dalam array - Grouping by year & month
                            If Not IsInArray(CStr(Left(lineitems(2), 7)), Month_arr) Then
                                array_num = WorksheetFunction.CountA(Month_arr)
                                
                                ReDim Preserve Month_arr(array_num)
                                ReDim Preserve Count_arr(array_num)
                                ReDim Preserve Sum_arr(array_num)
                                
                                getindex = array_num
                                Month_arr(array_num) = Left(lineitems(2), 7)
                                Count_arr(array_num) = 1
                                Sum_arr(array_num) = CDbl(lineitems(26))
                            Else
                                'jika month di line ini sudah ada dalam array
                                For i = LBound(Month_arr) To UBound(Month_arr)
                                    If Month_tmp = Month_arr(i) Then
                                    getindex = i
                                    Exit For
                                    End If
                                Next i
                                
                                Count_arr(getindex) = Count_arr(getindex) + 1
                                Sum_arr(getindex) = CDbl(Sum_arr(getindex)) + CDbl(lineitems(12))
                            End If
                        Else
                            Month_tmp = Left(lineitems(2), 7)
                            Count_arr(getindex) = Count_arr(getindex) + 1
                            Sum_arr(getindex) = CDbl(Sum_arr(getindex)) + CDbl(lineitems(27))
                        End If
                    End If
                Loop
                Close #1
            Next
            
            'print to sheet4
            For i = 1 To array_num
                ThisWorkbook.Sheets("pbl01").Range("A" & thisrow1 + (i - 1)).Value = Month_arr(i)
                ThisWorkbook.Sheets("pbl01").Range("B" & thisrow1 + (i - 1)).Value = Count_arr(i)
                ThisWorkbook.Sheets("pbl01").Range("C" & thisrow1 + (i - 1)).Value = Sum_arr(i)
            Next i
            
        End If
    End With
Sheets("pbl01").Select
'UserForm2.Show
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
End Function

Sub Import_()
Dim myfiles
Dim i As Integer
myfiles = Application.GetOpenFilename(Filefilter:="DHIB, TXT, CSV Files (*.), *.dhib", MultiSelect:=True)
Sheets("pbl01").Select

'DELL
 'Sheets("CDMK-0402").Cells.clear

If Not IsEmpty(myfiles) Then
    For i = LBound(myfiles) To UBound(myfiles)
         With ActiveSheet.QueryTables.Add(Connection:= _
            "TEXT;" & myfiles(i), Destination:=Range("A" & Rows.Count).End(xlUp).Offset(1, 0))
            .Name = "Sample"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 4, 1, 1)
        .TextFileTrailingMinusNumbers = True
'       .Refresh BackgroundQuery:=False
        On Error Resume Next
        End With
    Next i
    
    
Else
    MsgBox "No File Selected"
End If
'Sheet10.Select
'UserForm2.Show
End Sub

Sub import_excel()

Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    '---- Open filedialog
    With fDialog
        .AllowMultiSelect = True
        .Title = "Please select the files"
        .Filters.Clear
        .Filters.Add "All supported files", "*.xlsx"
        .Filters.Add "Excel Files", "*.xlsx"

        If .Show = True Then
            Dim fPath As Variant
            For Each fPath In .SelectedItems
                Open fPath For Input As #1
                Application.ScreenUpdating = False
                    
                Dim src As Workbook         ' The source workbook.
                Set src = Workbooks.Open(fPath, True, True)
                
                Dim iTotalRows As Integer   ' Get the total Used Range rows in the source file.
                iTotalRows = src.Worksheets("Sheet1").UsedRange.Rows.Count
                
                Dim iTotalCols As Integer   ' Get the total Columns in the source file.
                iTotalCols = src.Worksheets("Sheet1").UsedRange.Columns.Count
        
                Dim iRows, iCols As Integer
                ReDim urutan_arr(57) As Variant
                
                row_number = 0
                thisrow = 2
                i = 1
                Set myCell = ThisWorkbook.Sheets("Pengarsipan - BI").Range("A" & thisrow & ":BE" & thisrow)
                Do While Not i = 0
                    i = 0
                    For Each thisCell In myCell
                        If Not IsEmpty(thisCell) Then
                            i = i + 1
                        End If
                    Next thisCell
                    thisrow = thisrow + 1
                    Set myCell = ThisWorkbook.Sheets("Pengarsipan - BI").Range("A" & thisrow & ":BE" & thisrow)
                Loop
                
                ' Now, read the source and copy data to the master file.
                For iRows = 1 To iTotalRows
                    For iCols = 1 To iTotalCols
                        If iRows = 1 Then
                            cek_header = src.Worksheets("Sheet1").Cells(iRows, iCols)
                            j = 0
                            Set myCell = ThisWorkbook.Sheets("Pengarsipan - BI").Range("A1:BE1")
                            For Each thisCell In myCell
                                If thisCell = cek_header Then
                                    i = j + 1
                                End If
                                j = j + 1
                            Next thisCell
                            urutan_arr(iCols - 1) = i
                        Else
                            ThisWorkbook.Sheets("Pengarsipan - BI").Cells(thisrow - 2, urutan_arr(iCols - 1)) = src.Worksheets("Sheet1").Cells(iRows, iCols)
                        End If
                    Next iCols
                    thisrow = thisrow + 1
                Next iRows
                
                iRows = 0
            
                src.Close False
                Set src = Nothing
                Close #1
            Next
        End If
    End With
End Sub

Sub Import_all_sheet()
    'Check the target sheet name has already exists
    Dim ws As Worksheet
    Dim flag
    
    For Each ws In Worksheets
        If ws.Name = "Rencana Penarikan - BI" Then flag = True
    Next ws
    
    
    'Get the target folder name
    Dim folderPath
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            folderPath = .SelectedItems(1)
        End If
    End With
    
    'Add the new sheet to the last
   
    Dim file As String
    Dim extensionName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim acceptedArr
    acceptedArr = Array("xls", "xlsx", "xlsm", "csv", "tsv")
    Dim filePath
    Dim mergeSheetRows As Integer: mergeSheetRows = 0
    
    file = Dir(folderPath & "/*")
    
    Application.ScreenUpdating = False
    Do While file <> ""
        filePath = folderPath & "/" & file
        extensionName = fso.GetExtensionName(filePath)
        validFileFlag = filter(acceptedArr, extensionName)
        
        If UBound(validFileFlag) <> -1 Then
            Set wb = Workbooks.Open(FileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
            
            'Merge for each sheet in each file
            Dim sheet
            Dim sheetRows

            
            For Each sheet In wb.Worksheets
                sheetRows = sheet.Cells(Rows.Count, "A").End(xlUp).Row
                sheet.Rows("2:" & sheetRows).Copy ThisWorkbook.Worksheets("Rencana Penarikan - BI").Range("A" & mergeSheetRows + 1)
            Next
            mergeSheetRows = ThisWorkbook.Worksheets("Rencana Penarikan - BI").Cells(Rows.Count, "A").End(xlUp).Row
            
            wb.Close SaveChanges:=False
        End If
        
        'Proceed to the next file
        file = Dir()
    Loop
    Application.ScreenUpdating = True
    'UserForm2.Show
End Sub

Sub Import_txtnamefile()
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    '---- Open filedialog
    With fDialog
        .AllowMultiSelect = True
        .Title = "Please select the files"
        .Filters.Clear
        .Filters.Add "All supported files RKS", ""
        .Filters.Add "Text Files", ""


        If .Show = True Then
            Dim fPath As Variant
            Set fso = CreateObject("Scripting.FileSystemObject")

            For Each fPath In .SelectedItems
                Open fPath For Input As #1

                row_number = 0
                i = 0
                thisrow = 1
                '----- Dapetin baris terakhir dari table
                Set myCell = ThisWorkbook.Sheets("rks").Range("A" & thisrow & ":B" & thisrow)
                Do While i = 0
                    For Each thisCell In myCell
                        If IsEmpty(thisCell) Then
                            i = i + 1
                        End If
                    Next thisCell
                    thisrow = thisrow + 1
                    Set myCell = ThisWorkbook.Sheets("pbl01").Range("A" & thisrow & ":B" & thisrow)
                Loop

                count_line = 0
                Do Until EOF(1)
                    Line Input #1, LineFromFile
                    lineitems = Split(LineFromFile, "|")
                    count_line = count_line + 1
                Loop
                'Kalo pgn pake extension ini
                'ThisWorkbook.Sheets(5).Range("A" & thisrow - 1).Value = fso.getfilename(fPath)
                'Kalo ga pgn pake extension yg ini
                ThisWorkbook.Sheets("rks").Range("A" & thisrow - 1).Value = fso.getbasename(fPath)
                ThisWorkbook.Sheets("rks").Range("B" & thisrow - 1).Value = count_line
                Close #1
            Next

        End If
    End With
    '-----parsing

End Sub
Sub import_excel_with_name_file()
Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    '---- Open filedialog
    With fDialog
        .AllowMultiSelect = True
        .Title = "Please select the files"
        .Filters.Clear
        .Filters.Add "All supported files", "*.xls"
        .Filters.Add "Excel Files", "*.xls"

        If .Show = True Then
            Dim fPath As Variant
            For Each fPath In .SelectedItems
                Open fPath For Input As #1
                Application.ScreenUpdating = False
                    
                Dim src As Workbook         ' The source workbook.
                Set src = Workbooks.Open(fPath, True, True)
                
                Dim iTotalRows As Integer   ' Get the total Used Range rows in the source file.
                iTotalRows = src.Worksheets("sheet1").UsedRange.Rows.Count
                
                Dim iTotalCols As Integer   ' Get the total Columns in the source file.
                iTotalCols = src.Worksheets("sheet1").UsedRange.Columns.Count
        
                Dim iRows, iCols As Integer
                ReDim urutan_arr(58) As Variant
                
                row_number = 0
                thisrow = 2
                i = 1
                k = 1
                Set myCell = ThisWorkbook.Sheets("Sheet1").Range("A" & thisrow & ":BF" & thisrow)
                Do While Not i = 0
                    i = 0
                    For Each thisCell In myCell
                        If Not IsEmpty(thisCell) Then
                            i = i + 1
                        End If
                    Next thisCell
                    thisrow = thisrow + 1
                    Set myCell = ThisWorkbook.Sheets(1).Range("A" & thisrow & ":BF" & thisrow)
                Loop
                
                ' Now, read the source and copy data to the master file.
                For iRows = 1 To iTotalRows
                    For iCols = 1 To iTotalCols
                        If iRows = 1 Then
                            cek_header = src.Worksheets("Sheet1").Cells(iRows, iCols)
                            j = 0
                            Set myCell = ThisWorkbook.Sheets(1).Range("A1:BF1")
                            For Each thisCell In myCell
                                If thisCell = cek_header Then
                                    i = j + 1
                                End If
                                j = j + 1
                            Next thisCell
                            urutan_arr(iCols - 1) = i
                        Else
                            Set myCell = ThisWorkbook.Sheets(1).Range("A1:A" & thisrow)
                            If k = 1 Then
                                i = 0
                                j = 0
                                For Each thisCell In myCell
                                    If thisCell = src.Name And i = 0 Then
                                        i = k
                                    ElseIf thisCell = src.Name And i <> 0 Then
                                        j = k
                                    End If
                                    k = k + 1
                                Next thisCell
                                If i <> 0 Then
                                    ThisWorkbook.Sheets(1).Range("A" & i & ":BF" & k).delete
                                    thisrow = thisrow - (j - i) - 1
                                End If
                            End If
                            ThisWorkbook.Sheets(1).Cells(thisrow - 2, 1) = src.Name
                            ThisWorkbook.Sheets(1).Cells(thisrow - 2, urutan_arr(iCols - 1)) = src.Worksheets("Sheet1").Cells(iRows, iCols)
                        End If
                    Next iCols
                    thisrow = thisrow + 1
                Next iRows
                
                iRows = 0
            
                src.Close False
                Set src = Nothing
                Close #1
            Next
        End If
    End With
End Sub

Sub delete()
 Dim n As Integer
 For n = 1 To 10
    a = Worksheets("LCS").Cells(Rows.Count, 2).End(xlUp).Row
    For i = 35 To a
        If Worksheets("LCS").Cells(i, 2).Value <> "LCS" Then
            Rows(i).delete
        End If
    Next
 Next n
End Sub



