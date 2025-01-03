' The MergeDataIntoSample function automates the process of merging data from a "data" worksheet into a pre-existing "Sample" worksheet and creates a new workbook containing multiple sheets. 
' Each new sheet is based on a row in the "data" worksheet, with placeholders in the "Sample" worksheet replaced by corresponding values from the "data" worksheet. 
' The function also handles images by inserting them into the new sheets if the path is provided. The new workbook is saved with a timestamped filename.
Sub MergeDataIntoSample()
    ' Declare worksheets for data and sample, and variables for new workbook, ranges, and other data
    Dim wsData As Worksheet
    Dim wsSample As Worksheet
    Dim wsNew As Worksheet
    Dim newWorkbook As Workbook
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim cell As Range
    Dim columnHeaders As Collection
    Dim header As String
    Dim newSheetName As String
    Dim newContent As String
    Dim fileName As String
    Dim picPath As String
    Dim pic As Object
    Dim targetRange As Range
    Dim aspectRatio As Double
    Dim picWidth As Double, picHeight As Double
    
    ' Set references to "data" and "Sample" worksheets
    Set wsData = ThisWorkbook.Sheets("data")
    Set wsSample = ThisWorkbook.Sheets("Sample")
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Get the last row and last column of the "data" worksheet
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    ' Collect column headers from the first row of the "data" worksheet
    Set columnHeaders = New Collection
    For i = 1 To lastCol
        columnHeaders.Add wsData.Cells(1, i).Value
    Next i
    
    ' Loop through each row in the "data" worksheet (skipping the header row)
    For i = 2 To lastRow
        ' Create a new sheet by copying the "Sample" worksheet
        wsSample.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
        Set wsNew = newWorkbook.Sheets(newWorkbook.Sheets.Count)
        
        ' Set the new sheet's name based on the content in the third column (index 3)
        newSheetName = wsData.Cells(i, 3).Value
        wsNew.Name = newSheetName
        
        ' Replace placeholders in the new worksheet with actual data from the "data" sheet
        For Each cell In wsNew.UsedRange
            If InStr(cell.Value, "[") > 0 And InStr(cell.Value, "]") > 0 Then
                ' Loop through all column headers to replace placeholders with data
                For j = 1 To columnHeaders.Count
                    header = "[" & columnHeaders(j) & "]"
                    If InStr(cell.Value, header) > 0 Then
                        ' If the placeholder is for an image
                        If columnHeaders(j) = "hinhanh" Then
                            picPath = wsData.Cells(i, j).Value
                            On Error Resume Next
                            If Dir(picPath) <> "" Then
                                ' Insert image into the new sheet if the path exists
                                If cell.MergeCells Then
                                    Set targetRange = cell.MergeArea
                                Else
                                    Set targetRange = cell
                                End If
                                ' Add the picture and maintain aspect ratio
                                Set pic = wsNew.Shapes.AddPicture(fileName:=picPath, _
                                    LinkToFile:=msoFalse, SaveWithDocument:=msoCTrue, _
                                    Left:=0, Top:=0, Width:=-1, Height:=-1)
                                aspectRatio = pic.Width / pic.Height
                                picWidth = targetRange.Width
                                picHeight = targetRange.Height
                                ' Adjust the image to fit within the target range
                                If picWidth / aspectRatio <= targetRange.Height Then
                                    pic.Height = picWidth / aspectRatio
                                    pic.Width = picWidth
                                Else
                                    pic.Width = picHeight * aspectRatio
                                    pic.Height = picHeight
                                End If
                                pic.Left = targetRange.Left
                                pic.Top = targetRange.Top
                            End If
                            On Error GoTo 0
                        Else
                            ' Replace the placeholder with the actual value from the "data" sheet
                            newContent = wsData.Cells(i, j).Value
                            cell.Value = Replace(cell.Value, header, newContent)
                        End If
                    End If
                Next j
            End If
        Next cell
    Next i
    
    ' Delete the default sheet created in the new workbook
    Application.DisplayAlerts = False
    newWorkbook.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    ' Save the new workbook with a timestamped filename
    fileName = "KetQua " & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".xlsx"
    newWorkbook.SaveAs ThisWorkbook.Path & "\" & fileName
    
    ' Display a message when the process is complete
    MsgBox "Tron du lieu hoan tat. Tep da duoc luu voi ten: " & fileName, vbInformation
End Sub
