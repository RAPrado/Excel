# Procura um valor em uma matriz mas com a coluna fixa.
**=VLOOKUP(C6;plan1!A1:C26;3;falso)**<p>
Onde : =VLOOKUP(valor procurado;matriz a ser procurada;número da coluna a ser retornada;tipo da procura, sendo falso o valor exato)

# Procura um valor em uma matriz mas com a coluna dinâmica.
**=INDEX(plan1!$A$1:$O$26;MATCH($C6;plan1!$B$1:$B$26;0);MATCH(D$4;plan1!$A$1:$O$1;0))**<p>
Onde : =INDEX(Matriz a ser procurada;linha;coluna;sendo sendo zero o valor exato))<p>
Usar o Match para retornar a linha e coluna de forma dinâmica, por exemplo : MATCH(Valor procurado;Matriz a ser procurada;0)
