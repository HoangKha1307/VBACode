Function RemoveNonTextAndNumber(inputString As String) As String
    ' Declare a variable to store the result string
    Dim result As String
    ' Declare a variable for looping through each character in the input string
    Dim i As Integer

    ' Initialize the result string as empty
    result = ""
    
    ' Loop through each character in the inputString
    For i = 1 To Len(inputString)
        ' Check if the current character is either a letter (A-Z, a-z) or a digit (0-9)
        If Mid(inputString, i, 1) Like "[A-Za-z0-9]" Then
            ' If it's a letter or a digit, append it to the result string
            result = result & Mid(inputString, i, 1)
        End If
    Next i

    ' Return the result string containing only alphabetic characters and digits
    RemoveNonTextAndNumber = result
End Function
