Function splitText(fullText As String, separator As String, position As Integer)
    Dim Vector As Variant

    Vector = Split(fullText, separator)
    If position > UBound(Vector) Then
        splitText = "ERROR - Position greater than division quantity"
    Else
        splitText = Vector(position)
    End If

End Function
