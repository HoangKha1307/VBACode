Function RemoveNonText(inputString As String) As String
    ' Declare a variable to store the result string
    Dim result As String
    ' Declare a variable for looping through each character in the input string
    Dim i As Integer

    ' Initialize the result string as empty
    result = ""
    
    ' Loop through each character in the inputString
    For i = 1 To Len(inputString)
        ' Check if the current character is a letter (either uppercase or lowercase)
        If Mid(inputString, i, 1) Like "[A-Za-z]" Then
            ' If it's a letter, append it to the result string
            result = result & Mid(inputString, i, 1)
        End If
    Next i

    ' Return the result string containing only alphabetic characters (A-Z, a-z)
    RemoveNonText = result
End Function
