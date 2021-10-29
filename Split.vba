Sub Criar_Lista(Celulas As String)
    Dim Lista_Celulas() As String
    Dim i As Integer
    
    Lista_Celulas = Split(Celulas, ",") 'Cria array com as celulas informadas.
            
    For i = 0 To UBound(Lista_Celulas)        
        Range(Lista_Celulas(i)).Font.Color = vbWhite
    Next

End Sub
