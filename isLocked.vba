'Function to know if a cell is locked for edtion
'Função para saber se uma célula está bloqueada para edição. Salve esse código dentro de uma macro e depois só chamar a funcção.

Function isLocked(cell As Object) As Boolean
    If cell.Locked Then
        isLocked = True
    Else
        isLocked = False
    End If
End Function
