'Código para impedir a impressão de uma ou todas as abas de um arquivo em Excel.
'Code to block the print of one or all sheets of Excel file.

'1-No gerenciador de projetos no VBA do Excel, vá em VBAProject (nome do arquivo excel.xlsx).
'2-Selecione ThisWorkBook.
'3-Na janela de código, selecione Workbook, e em eventos vá em Workbook_BeforePrint.
'4-Entre com o código abaixo.

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim Nome_Aba As String
    Nome_Aba = "Sheet1"
    
    For Each Objeto In Application.ActiveWorkbook.Windows(1).SelectedSheets
        If Objeto.Name = Nome_Aba Then 'Se quiser impedir a impressão para uma aba específica, do contrário deixar apenas a linha 15 e 16.
            MsgBox ("Pedrão, achaste que podias imprimir???? Só que não...")
            Cancel = True
        End If
    Next
End Sub
