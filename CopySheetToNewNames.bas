Sub CopySheetToNewNames()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim cell As Range
    Dim newSheetName As String
    Dim selectedCell As Range
    Dim sheetName As String
    Dim selectedRange As Range
    Dim sheetExists As Boolean
    
    ' Display a dialog for the user to select any cell in the source sheet
    On Error Resume Next ' Ignore errors if the user doesn't select a cell
    Set selectedCell = Application.InputBox("Select any cell in the source sheet", Type:=8)
    On Error GoTo 0 ' Reset error handling
    
    ' Check if the user didn't select a cell
    If selectedCell Is Nothing Then
        MsgBox "You didn't select a cell. The process has been canceled.", vbExclamation
        Exit Sub
    End If
    
    ' Get the name of the sheet of the selected cell
    sheetName = selectedCell.Worksheet.Name
    
    ' Set the source worksheet
    Set wsSource = ThisWorkbook.Sheets(sheetName)
    
    ' Display a dialog for the user to select the range with new sheet names
    On Error Resume Next ' Ignore errors if the user doesn't select a range
    Set selectedRange = Application.InputBox("Select the range of new sheet names (e.g., B16:B27)", Type:=8)
    On Error GoTo 0 ' Reset error handling
    
    ' Check if the user didn't select a range
    If selectedRange Is Nothing Then
        MsgBox "You didn't select a range for new sheet names. The process has been canceled.", vbExclamation
        Exit Sub
    End If
    
    ' Determine the first and last row in the selected range
    firstRow = selectedRange.Row
    lastRowInRange = selectedRange.Rows(selectedRange.Rows.Count).Row
    
    ' Loop through each cell in the selected range to create new sheets
    For Each cell In selectedRange
        If cell.Value <> "" Then
            newSheetName = cell.Value
            
            ' Check if the sheet already exists
            sheetExists = False
            On Error Resume Next
            Set wsNew = ThisWorkbook.Sheets(newSheetName)
            On Error GoTo 0
            
            ' If the sheet doesn't exist, copy and rename the new sheet
            If wsNew Is Nothing Then
                ' Copy the source sheet
                wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                
                ' Set the newly copied sheet
                Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                wsNew.Name = newSheetName
                
                ' Delete unnecessary rows
                For i = lastRowInRange To firstRow Step -1
                    If i <> cell.Row Then
                        wsNew.Rows(i).Delete
                    End If
                Next i
            End If
            
            ' Reset wsNew to continue with the next cell
            Set wsNew = Nothing
        End If
    Next cell
End Sub
