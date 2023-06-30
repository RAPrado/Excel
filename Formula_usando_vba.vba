'Exemplo de como criar fórmula usando vba.
Sub Filtrar()
  'B1 = 10

  'Em vba na fórmula, o "ponto e vírgula" deve ser substituído por "vígula".  

  'Exempolo 1 : Contará tudo que for igual ao valor da célula B1
  range("A1").Formula = "=countifs(c1:c5,b1)"

  'Exempolo 2 : Contará tudo que for maior ou igual ao valor da célula B1
  'Em vba na fórmula, o aspas (") deve ser substituído por duas aspas ("")., logo ">=", ficará como "">="".
  range("A1").Formula = "=countifs(c1:c5,""">"" & b1)"
End Sub
