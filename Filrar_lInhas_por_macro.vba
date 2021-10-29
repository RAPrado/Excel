'Código a ser inserido na sheet
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Celula As String
    Dim Linha_A, Linha_B As String
    
    Linha_A = "D3,E3,F3,G3,D5,E5,F5,G5" 'Celulas com os status a serem filtrados.
    Linha_B = "D4,E4,F4,G4,D6,E6,F6,G6" 'Celulas que ao clicarem irão retirar o filtro.

    
    If Selection.Count = 1 Then
        Celula = Replace(ActiveCell.Address, "$", "")
    
        If InStr(1, Linha_A, Celula) > 0 Then
            Loop_Muda_Cor (Linha_B)          'Muda para branco a cor de todos os status.
            Muda_Cor_Fonte Celula, -16776961 'Muda para vermelho o status selecionado.
                       
            Filtro_Status (Range(Celula).Value)
            
        ElseIf InStr(1, Linha_B, Celula) > 0 Then
            Loop_Muda_Cor (Linha_B) 'Muda para branco a cor de todos os status.
            Filtro_Status ("")
        End If
    End If
End Sub

'Código a ser inserido no módulo
Sub Filtro_Status(Filtro As String)
    Application.CutCopyMode = False
    
    Range("Q4").Value = Filtro
    Range("B8:K39").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("Q3:Q4"), Unique:=False
    'Onde :
    'Range("B8:K39").AdvancedFilter Action:=xlFilterInPlace = Linhas a serem filtradas
    'CriteriaRange:=Range("Q3:Q4") = Celulas que contêm os filtros, sendo a Q3 o título, e Q4 o conteúdo a ser filtrado. O título na Q3 deve estar escrito exatamente igual ao título na Range("B8:K39"), se tiver espaço a esquerda ou direita do conteúdo, deve estar assim no filtro também.
End Sub

Sub Muda_Cor_Fonte(Celula As String, Cor As Variant)
    Range(Celula).Font.Color = Cor
End Sub

Sub Loop_Muda_Cor(Celulas As String)
    Dim Lista_Celulas() As String
    Dim i As Integer
    
    Lista_Celulas = Split(Celulas, ",") 'Cria array com as celulas informadas.
            
    For i = 0 To UBound(Lista_Celulas)
        Muda_Cor_Fonte Left(Lista_Celulas(i), 1) & Right(Lista_Celulas(i), 1) - 1, vbWhite 'Subtrai 1 da posição da célula, pois nesse caso deseja mudar a cor da linha superior à que foi clicada.
    Next

End Sub
