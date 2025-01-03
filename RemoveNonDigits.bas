Function RemoveNonDigits(inputString As String) As String
    ' Declare a variable to hold the result string
    Dim result As String
    ' Declare a variable for looping through each character in the input string
    Dim i As Integer

    ' Initialize the result as an empty string
    result = ""
    
    ' Loop through each character in the inputString
    For i = 1 To Len(inputString)
        ' Check if the current character is a digit using the Like "#" pattern
        If Mid(inputString, i, 1) Like "#" Then
            ' If it's a digit, append it to the result string
            result = result & Mid(inputString, i, 1)
        End If
    Next i

    ' Return the result string containing only digits
    RemoveNonDigits = result
End Function
