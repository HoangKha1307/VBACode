' The ListFilesInFolder subroutine allows the user to select a folder and generates a new worksheet listing all files within that folder. 
' Each file is listed with its name and a clickable hyperlink that leads to the file's location.
Sub ListFilesInFolder()
    Dim folderPath As String
    Dim fs As Object
    Dim f As Object
    Dim i As Long
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim wsExist As Boolean

    ' Open a folder selection dialog for the user
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)  ' If a folder is selected, store its path
        Else
            MsgBox "You did not select a folder. The process is cancelled.", vbExclamation
            Exit Sub  ' Exit if no folder is selected
        End If
    End With

    ' Check if the worksheet "List Files" already exists
    On Error Resume Next
    Set wsNew = ThisWorkbook.Worksheets("List Files")
    On Error GoTo 0

    If wsNew Is Nothing Then
        ' Create a new worksheet if it does not exist
        Set wsNew = ThisWorkbook.Worksheets.Add
        wsNew.Name = "List Files"
    Else
        ' If it exists, clear the existing content
        wsNew.Cells.Clear
    End If

    ' Create a FileSystemObject to work with files
    Set fs = CreateObject("Scripting.FileSystemObject")

    ' Write column headers to the worksheet
    With wsNew
        .Cells(1, 1).Value = "File Name"
        .Cells(1, 2).Value = "Link"
        newRow = 2  ' Start populating data from row 2

        ' Loop through all the files in the selected folder
        For Each f In fs.GetFolder(folderPath).Files
            .Cells(newRow, 1).Value = f.Name  ' Write file name to column 1
            .Hyperlinks.Add Anchor:=.Cells(newRow, 2), Address:=f.Path  ' Create hyperlink in column 2
            newRow = newRow + 1  ' Move to the next row
        Next f
    End With

    ' Release the FileSystemObject
    Set fs = Nothing

    ' Display a message box when the process is complete
    MsgBox "Process completed. The files have been listed in the new worksheet.", vbInformation
End Sub
