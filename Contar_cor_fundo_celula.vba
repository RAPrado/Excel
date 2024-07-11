'Referencia : https://learn.microsoft.com/pt-pt/previous-versions/office/troubleshoot/office-developer/count-cells-number-with-color-using-vba
'
'Parâmetros:
'range_data = recebe o range de células a ser contado.
'criteria = recebe a cor de fundo a ser contada. Ao invés de passar o nome da cor, basta passar a célula que tem a cor de fundo que se deseja contar.
Function CountCcolor(range_data As Range, criteria As Range) As Long
    Dim datax As Range
    Dim xcolor As Long
    xcolor = criteria.Interior.ColorIndex
    For Each datax In range_data
        If datax.Interior.ColorIndex = xcolor Then
            CountCcolor = CountCcolor + 1
        End If
    Next datax
End Function
