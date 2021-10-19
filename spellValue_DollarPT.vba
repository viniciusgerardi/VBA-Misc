' Adapted form https://support.microsoft.com/en-us/office/convert-numbers-into-words-a0d166fb-e1ea-4090-95c8-69442cd55d98

Function SpellCurrencyDollarBR(ByVal MyNumber)
    If MyNumber > 1E+15 Then
        SpellCurrencyDollarBR = "Error: Number too big"
    Else
    
        Dim Dollars, Cents, Temp
        Dim DecimalPlace, Count
        ReDim Place(9) As String
        ReDim Places(9) As String
        
        
        Place(2) = " mil "
        Place(3) = " milhão "
        Place(4) = " bilhão "
        Place(5) = " trilhão "
        Places(2) = " mil "
        Places(3) = " milhões "
        Places(4) = " bilhões "
        Places(5) = " trilhões "
        
        
        
        
        ' String representation of amount.
        MyNumber = Trim(Str(MyNumber))
        ' Position of decimal place, 0 if none.
        DecimalPlace = InStr(MyNumber, ".")
        ' Convert cents and set MyNumber to dollar amount.
        If DecimalPlace > 0 Then
            Cents = GetTensBR(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If
        '

        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundredsBR(Right(MyNumber, 3))
            If Temp <> "" Then
                If MyNumber > 1 Then
                    Dollars = Temp & Places(Count) & Dollars
                Else
                    Dollars = Temp & Place(Count) & Dollars
                End If
            End If
            
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
        Select Case Dollars
            Case ""
                Dollars = "zero Dólares"
            Case "um"
                Dollars = "um Dólar"
             Case Else
                Dollars = Dollars & " Dólares"
        End Select
        Select Case Cents
            Case ""
                Cents = " e nenhum centavo"
            Case "One"
                Cents = " e um centavo"
                  Case Else
                Cents = " e " & Cents & " centavos"
        End Select
        SpellCurrencyDollarBR = Dollars & Cents
    'End If
        
    End If
End Function
      
' Converts a number from 100-999 into text
Function GetHundredsBR(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Select Case MyNumber
            Case 100: Result = "cem"
            Case 101 To 199: Result = "cento e "
            Case 200: Result = "duzentos"
            Case 201 To 299: Result = "duzentos e "
            Case 300: Result = "trezentos"
            Case 301 To 399: Result = "trezentos e "
            Case 400: Result = "quatrocentos"
            Case 401 To 499: Result = "quatrocentos e "
            Case 500: Result = "quinhentos"
            Case 501 To 599: Result = "quinhentos e "
            Case 600: Result = "seiscentos"
            Case 601 To 699: Result = "seiscentos e "
            Case 700: Result = "setecentos"
            Case 701 To 799: Result = "setecentos e "
            Case 800: Result = "oitocentos"
            Case 801 To 899: Result = "oitocentos e "
            Case 900: Result = "novecentos"
            Case 901 To 999: Result = "novecentos e "
            Case Else
        End Select
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTensBR(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigitBR(Mid(MyNumber, 3))
    End If
    GetHundredsBR = Result
    
    
End Function
      
' Converts a number from 10 to 99 into text.
Function GetTensBR(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "dez"
            Case 11: Result = "onze"
            Case 12: Result = "doze"
            Case 13: Result = "treze"
            Case 14: Result = "catorze"
            Case 15: Result = "quinze"
            Case 16: Result = "dezesseis"
            Case 17: Result = "dezeste"
            Case 18: Result = "dezoito"
            Case 19: Result = "dezenove"
            Case Else
        End Select
    Else                     ' If value between 20-99...
        Select Case Val(Left(TensText, 2))
            Case 20: Result = "vinte "
            Case 21 To 29: Result = "vinte e "
            Case 30: Result = "trinta "
            Case 31 To 39: Result = "trinta e "
            Case 40: Result = "quarenta "
            Case 41 To 49: Result = "quarenta e "
            Case 50: Result = "cinquenta "
            Case 51 To 59: Result = "cinquenta e "
            Case 60: Result = "sessenta "
            Case 61 To 69: Result = "sessenta e "
            Case 70: Result = "setenta "
            Case 71 To 79: Result = "setenta e "
            Case 80: Result = "oitenta "
            Case 81 To 89: Result = "oitenta e "
            Case 90: Result = "noventa "
            Case 91 To 99: Result = "noventa e "
            Case Else
        End Select
        Result = Result & GetDigitBR _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTensBR = Result
End Function
     
' Converts a number from 1 to 9 into text.
Function GetDigitBR(Digit)
    Select Case Val(Digit)
        Case 1: GetDigitBR = "um"
        Case 2: GetDigitBR = "dois"
        Case 3: GetDigitBR = "três"
        Case 4: GetDigitBR = "quatro"
        Case 5: GetDigitBR = "cinco"
        Case 6: GetDigitBR = "seis"
        Case 7: GetDigitBR = "sete"
        Case 8: GetDigitBR = "oito"
        Case 9: GetDigitBR = "nove"
        Case Else: GetDigitBR = ""
    End Select
End Function
