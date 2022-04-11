Attribute VB_Name = "modFManager"
Public fPath As String
Public IsSubFolder As Boolean
Public iRow As Long
Public FSO As Scripting.FileSystemObject
Public SourceFolder As Scripting.folder, SubFolder As Scripting.folder
Public FileItem As Scripting.File
Public IsFileTypeExists As Boolean
Public Sub ListFilesInFolder(SourceFolder As Scripting.folder, IncludeSubfolders As Boolean)
    On Error Resume Next
    For Each FileItem In SourceFolder.Files
' display file properties
        Cells(iRow, 2).Formula = iRow - 13
        Cells(iRow, 3).Formula = FileItem.Name
        Cells(iRow, 4).Formula = FileItem.Path
        Cells(iRow, 5).Formula = Int(FileItem.Size / 1024)
        Cells(iRow, 6).Formula = FileItem.Type
        Cells(iRow, 7).Formula = FileItem.DateLastModified
        Cells(iRow, 8).Select
        Selection.Hyperlinks.Add Anchor:=Selection, Address:= _
        FileItem.Path, TextToDisplay:="Click Here to Open"
'Cells(iRow, 8).Formula = "=HYPERLINK(""" & FileItem.Path & """,""" & "Click Here to Open" & """)"
        iRow = iRow + 1 ' next row number
        Next FileItem
        If IncludeSubfolders Then
            For Each SubFolder In SourceFolder.SubFolders
                ListFilesInFolder SubFolder, True
                Next SubFolder
            End If
            
            Set FileItem = Nothing
            Set SourceFolder = Nothing
            Set FSO = Nothing
        End Sub

        Public Sub ListFilesInFolderXtn(SourceFolder As Scripting.folder, IncludeSubfolders As Boolean)
            On Error Resume Next
            Dim FileArray As Variant
            FileArray = Get_File_Type_Array
            For Each FileItem In SourceFolder.Files
                Call ReturnFileType(FileItem.Type, FileArray)
                If IsFileTypeExists = True Then
                    Cells(iRow, 2).Formula = iRow - 13
                    Cells(iRow, 3).Formula = FileItem.Name
                    Cells(iRow, 4).Formula = FileItem.Path
                    Cells(iRow, 5).Formula = Int(FileItem.Size / 1024)
                    Cells(iRow, 6).Formula = FileItem.Type
                    Cells(iRow, 7).Formula = FileItem.DateLastModified
                    Cells(iRow, 8).Select
                    Selection.Hyperlinks.Add Anchor:=Selection, Address:= _
                    FileItem.Path, TextToDisplay:="Click Here to Open"
'Cells(iRow, 8).Formula = "=HYPERLINK(""" & FileItem.Path & """,""" & "Click Here to Open" & """)"
                    iRow = iRow + 1 ' next row number
                End If
                Next FileItem
                If IncludeSubfolders Then
                    For Each SubFolder In SourceFolder.SubFolders
                        ListFilesInFolderXtn SubFolder, True
                        Next SubFolder
                    End If
                    Set FileItem = Nothing
                    Set SourceFolder = Nothing
                    Set FSO = Nothing
                End Sub
                
                Sub ClearResult()
                    If Range("B14") <> "" Then
                        Range("B14").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Range(Selection, Selection.End(xlToRight)).Select
                        Range(Selection.Address).ClearContents
                    End If
                End Sub
                Public Function ReturnFileType(fileType As String, FileArray As Variant) As Boolean
                    Dim I As Integer
                    IsFileTypeExists = False
                    For I = 1 To UBound(FileArray) + 1
                        If FileArray(I - 1) = fileType Then
                            IsFileTypeExists = True
                            Exit For
                        Else
                            IsFileTypeExists = False
                        End If
                    Next
                End Function
            Public Function Get_File_Type_Array() As Variant
                    Dim I, j, TotalSelected As Integer
                    Dim arrList() As String
                    TotalSelected = 0
                    For I = 0 To Sheet1.ListBoxFileTypes.ListCount - 1
                        If Sheet1.ListBoxFileTypes.Selected(I) = True Then
                            TotalSelected = TotalSelected + 1
                        End If
                    Next
                    ReDim arrList(0 To TotalSelected - 1) As String
                    j = 0
                    I = 0
                    For I = 0 To Sheet1.ListBoxFileTypes.ListCount - 1
                        If Sheet1.ListBoxFileTypes.Selected(I) = True Then
                            arrList(j) = Left(Sheet1.ListBoxFileTypes.List(I), InStr(1, Sheet1.ListBoxFileTypes.List(I), "(") - 1)
                            j = j + 1
                        End If
                    Next
                    Get_File_Type_Array = arrList
                End Function
Function Timestamp(Reference As Range)
    If Reference.Value <> "" Then
        Timestamp = Format(Now, "dd-mm-yyyy hh:mm:ss")
    Else
        Timestamp = ""
    End If
End Function

