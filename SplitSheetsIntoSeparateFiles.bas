' The SplitSheetsIntoSeparateFiles subroutine is designed to automate the process of saving each sheet in an Excel workbook as a separate file. 
' This is useful in scenarios where you want to break up a large workbook into smaller, individual files for sharing, backup, or further processing.
Sub SplitSheetsIntoSeparateFiles()
    ' Declare variables for the current workbook, new workbook, and file paths
    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim filePath As String
    Dim sheetName As String
    Dim newFilePath As String
    
    ' Save the path of the current workbook
    Set currentWorkbook = ThisWorkbook
    filePath = currentWorkbook.Path
    
    ' Check if the workbook has been saved (if there's no file path, it hasn't been saved)
    If filePath = "" Then
        MsgBox "Please save the workbook first to a valid location.", vbExclamation
        Exit Sub
    End If
    
    ' Turn off screen updating, alerts, and automatic calculations to speed up the process
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual ' Disable automatic calculations
    
    ' Loop through each sheet in the current workbook
    For Each ws In currentWorkbook.Sheets
        sheetName = ws.Name
        newFilePath = filePath & "\" & sheetName & ".xlsx" ' Define the file path for the new workbook
        
        ' Create a new workbook and copy the sheet into the new workbook
        Set newWorkbook = Workbooks.Add
        ws.Copy After:=newWorkbook.Sheets(1)
        
        ' Check if the default sheet ("Sheet1") exists in the new workbook and delete it
        If newWorkbook.Sheets(1).Name = "Sheet1" Then
            Application.DisplayAlerts = False
            newWorkbook.Sheets(1).Delete
            Application.DisplayAlerts = True
        End If
        
        ' Save the new workbook with the sheet name as its file name
        On Error GoTo ErrorHandler ' Add error handling to manage saving issues
        newWorkbook.SaveAs FileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
        newWorkbook.Close SaveChanges:=False
        On Error GoTo 0 ' Reset error handling
        
    Next ws
    
    ' Re-enable screen updating, alerts, and automatic calculations after the process completes
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Display a message to indicate the operation was successful
    MsgBox "Sheets have been split into separate files.", vbInformation
    Exit Sub

ErrorHandler:
    ' If an error occurs, display an error message and handle cleanup
    MsgBox "Error occurred while saving the file: " & newFilePath, vbCritical
    newWorkbook.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
