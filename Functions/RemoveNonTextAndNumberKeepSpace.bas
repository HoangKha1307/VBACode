Function RemoveNonTextAndNumberKeepSpace(inputString As String) As String
    ' Declare a variable to store the result string
    Dim result As String
    ' Declare a variable for looping through each character in the input string
    Dim i As Integer
    ' Declare a variable to track the last character added to the result
    Dim lastChar As String

    ' Initialize the result string as empty and lastChar as an empty string
    result = ""
    lastChar = ""

    ' Loop through each character in the inputString
    For i = 1 To Len(inputString)
        ' Check if the current character is a letter (A-Z, a-z) or a digit (0-9)
        If Mid(inputString, i, 1) Like "[A-Za-z0-9]" Then
            ' If it's a letter or a digit, append it to the result string
            result = result & Mid(inputString, i, 1)
            ' Update the lastChar variable to track the current character
            lastChar = Mid(inputString, i, 1)
        ' Check if the current character is a space
        ElseIf Mid(inputString, i, 1) = " " Then
            ' If the last character was not a space, add a space to the result string
            If lastChar <> " " Then
                result = result & " "
                ' Update lastChar to indicate the current space
                lastChar = " "
            End If
        End If
    Next i

    ' Remove any leading or trailing spaces from the result string and return the result
    RemoveNonTextAndNumberKeepSpace = Trim(result)
End Function
