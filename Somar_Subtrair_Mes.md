# Somar ou subtrair o mês em uma data.

Célula C1: 23/06/2023<br>
**=EDATE(DATEVALUE("01/"& MONTH(C1)&"/"&YEAR(C1));-1)**<p>
  
  **Onde:**<br>  
  DATEVALUE=Converte uma data em texto para data em número. Sendo que "01/" é usado para deixar a data com o dia 1.<br>
  EDATE=Soma ou subtrai o valor do mês em uma data. Sendo que "-1" indica que deseja subtrair 1 do mês. Para somar, deixar sem sinal negativo.
