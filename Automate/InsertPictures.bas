' This subroutine, InsertPictures, inserts images into cells of an Excel worksheet based on a list of image file paths provided by the user. 
' It handles both merged and non-merged cells and ensures the images fit into the target cells, maintaining their aspect ratio.
Sub InsertPictures()
    Dim ws As Worksheet
    Dim imgPaths As Range
    Dim targetCells As Range
    Dim cell As Range
    Dim pic As Picture
    Dim targetCell As Range
    Dim imgPath As String
    Dim aspectRatio As Single
    Dim mergedAreasCount As Integer

    ' Set the current worksheet (ActiveSheet) where images will be inserted
    Set ws = ActiveSheet

    ' Prompt the user to select the range containing image paths
    On Error Resume Next
    Set imgPaths = Application.InputBox("Select the range containing image paths:", Type:=8)
    On Error GoTo 0
    If imgPaths Is Nothing Then Exit Sub ' Exit if no range is selected

    ' Prompt the user to select the range containing target cells where images will be inserted
    On Error Resume Next
    Set targetCells = Application.InputBox("Select the range containing target cells:", Type:=8)
    On Error GoTo 0
    If targetCells Is Nothing Then Exit Sub ' Exit if no range is selected

    ' Count the number of merged areas or cells in the target range
    mergedAreasCount = 0
    For Each cell In targetCells
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                mergedAreasCount = mergedAreasCount + 1
            End If
        Else
            mergedAreasCount = mergedAreasCount + 1
        End If
    Next cell

    ' Ensure the two ranges have the same number of cells or merged areas
    If imgPaths.Cells.Count <> mergedAreasCount Then
        MsgBox "Both ranges must have the same number of cells or merged areas.", vbExclamation
        Exit Sub
    End If

    Dim imgIndex As Integer
    imgIndex = 1

    ' Loop through each cell in the range containing image paths
    For Each cell In imgPaths
        imgPath = cell.Value ' Get the image path from the current cell
        If imgPath <> "" Then
            ' Find the corresponding target cell
            Dim targetIndex As Integer
            targetIndex = 1

            ' Loop through the target cells to match the image with the target cell
            For Each targetCell In targetCells
                If targetCell.MergeCells Then
                    If targetCell.Address = targetCell.MergeArea.Cells(1, 1).Address Then
                        If targetIndex = imgIndex Then
                            Exit For
                        End If
                        targetIndex = targetIndex + 1
                    End If
                Else
                    If targetIndex = imgIndex Then
                        Exit For
                    End If
                    targetIndex = targetIndex + 1
                End If
            Next targetCell

            imgIndex = imgIndex + 1

            ' Delete any existing pictures in the target cell before inserting a new one
            Dim shp As Shape
            For Each shp In ws.Shapes
                If Not Intersect(shp.TopLeftCell, targetCell) Is Nothing Then
                    shp.Delete
                End If
            Next shp

            ' Insert the picture from the image path
            Set pic = ws.Pictures.Insert(imgPath)

            ' Calculate the aspect ratio of the image (width / height)
            aspectRatio = pic.Width / pic.Height

            ' Adjust the size of the image to fit the target cell
            With targetCell
                If .MergeCells Then
                    ' If the target cell is merged, adjust the image to fit the merged area
                    With .MergeArea
                        If .Width / aspectRatio > .Height Then
                            pic.Height = .Height
                            pic.Width = .Height * aspectRatio
                        Else
                            pic.Width = .Width
                            pic.Height = .Width / aspectRatio
                        End If
                        ' Position the image in the center of the merged area
                        pic.Left = .Left + (.Width - pic.Width) / 2
                        pic.Top = .Top + (.Height - pic.Height) / 2
                    End With
                Else
                    ' If the target cell is not merged, adjust the image to fit the cell
                    If .Width / aspectRatio > .Height Then
                        pic.Height = .Height
                        pic.Width = .Height * aspectRatio
                    Else
                        pic.Width = .Width
                        pic.Height = .Width / aspectRatio
                    End If
                    ' Position the image in the center of the cell
                    pic.Left = .Left + (.Width - pic.Width) / 2
                    pic.Top = .Top + (.Height - pic.Height) / 2
                End If
            End With
        End If
    Next cell
End Sub
