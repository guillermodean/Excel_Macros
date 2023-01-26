Public Function checkValue(ByVal valor As String) As String
    If Len(valor) = 0 Then
        checkValue = " "
    Else
        checkValue = valor
    End If
End Function

Public Function checkValueInt(ByVal valor As String) As Double
    If Len(valor) = 0 Then
        valor = 0
        checkValueInt = valor
    Else
End Function
