' The ConsolidateWorksheets subroutine is designed to combine data from multiple worksheets across various Excel files into a single workbook. 
' The subroutine allows users to select a folder containing Excel files, and it then creates a new workbook that consolidates all the worksheets from those files.
Sub ConsolidateWorksheets()
    Dim wbDest As Workbook
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim destRow As Long
    
    ' Prompt the user to enter the folder path containing the Excel files
    folderPath = InputBox("Enter the folder path containing the Excel files:")
    
    ' Ensure the folder path ends with a backslash (\) for consistency
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Create a new workbook for consolidating the data
    Set wbDest = Workbooks.Add
    
    ' Loop through all Excel files in the specified folder
    fileName = Dir(folderPath & "*.xls*") ' Get the first Excel file
    Do While fileName <> "" ' Continue looping until no more files are found
        Set wbSource = Workbooks.Open(folderPath & fileName) ' Open the source workbook
        
        ' Loop through each worksheet in the source workbook
        For Each wsSource In wbSource.Worksheets
            ' Create a new worksheet in the destination workbook
            Set wsDest = wbDest.Worksheets.Add
            wsDest.Name = wbSource.Name & "_" & wsSource.Name ' Name the worksheet using source file and sheet names
            
            ' Find the last used row and column in the source worksheet
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
            
            ' Copy the data from the source worksheet to the destination worksheet
            wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol)).Copy _
                Destination:=wsDest.Cells(1, 1)
        Next wsSource
        
        ' Close the source workbook without saving any changes
        wbSource.Close SaveChanges:=False
        
        ' Get the next file in the folder
        fileName = Dir
    Loop
    
    ' Save the consolidated workbook with a specific name in the same folder
    wbDest.SaveAs folderPath & "Consolidated_Workbook.xlsx"
    wbDest.Close ' Close the destination workbook
    
    ' Notify the user that the consolidation is complete
    MsgBox "Consolidation of worksheets is complete!"
End Sub
