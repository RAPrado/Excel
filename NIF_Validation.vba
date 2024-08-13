'Valida se o número informado é um NIF português válido.
'Valid if the number is a portugues NIF valid.
'Parameters :
'Contribuinte : Recieve the NIF to be validate.
'Return : True if is valid or False if is not.

Function Valida_NIF(ByVal Contribuinte As String) As Boolean
    Dim I As Integer
    Dim Digito As Integer
        
    Valida_NIF = False
    Digito = 0
        
    If Len(Contribuinte) = 9 And IsNumeric(Contribuinte) = True Then
        If (Mid(Contribuinte, 1, (1)) = 1 Or Mid(Contribuinte, 1, (1)) = 2 Or Mid(Contribuinte, 1, (1)) = 3 Or Mid(Contribuinte, 1, (1)) = 4 Or Mid(Contribuinte, 1, (1)) = 5 Or Mid(Contribuinte, 1, (1)) = 6 Or Mid(Contribuinte, 1, (1)) = 8 Or Mid(Contribuinte, 1, (1)) = 9) Then
            For I = 1 To 8
                Digito = Digito + (Mid(Contribuinte, I, (1)) * (10 - I))
            Next
            
            Digito = 11 - (Digito Mod 11)
            
            If Digito >= 10 Then
                Digito = 0
            End If
            
            If (Digito = Mid(Contribuinte, 9, (1))) Then
                Valida_NIF = True
            End If
        End If
    End If
End Function
