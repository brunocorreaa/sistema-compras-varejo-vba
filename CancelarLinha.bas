Attribute VB_Name = "CancelarLinha"
Sub bCancelarLinha()
    ' Desativa a atualização de tela para melhorar a performance e evitar "piscar"
    Application.ScreenUpdating = False

    ' --- DECLARAÇÕES E ATRIBUIÇÕES ---
    Dim db As Worksheet, ap As Worksheet, mn As Worksheet
    Dim Controle_Macro As Worksheet, Controle_Erro As Worksheet
    Dim lastRow As Long, countNonEmpty As Long, last_row_macro As Long
    Dim resposta As Integer, usuario As String, horaAtual As String
    Dim cell As Range, hoje As Date
    
    Set db = ThisWorkbook.Sheets("LINHAS_COLECAO")
    Set ap = ThisWorkbook.Sheets("Apoio")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
    Set Controle_Erro = ThisWorkbook.Sheets("Controle-Erro")
    
    ' --- VALIDAÇÃO DE DADOS ---
    ' Determina quantas linhas estão marcadas para cancelamento na coluna "C" (IrCadastro)
    lastRow = db.Cells(db.Rows.Count, "C").End(xlUp).Row
    countNonEmpty = 0
    
    For Each cell In db.Range("C1:C" & lastRow)
        If Len(Trim(cell.Value)) > 0 Then countNonEmpty = countNonEmpty + 1
    Next cell
    
    ' --- DUPLA CONFIRMAÇÃO DE SEGURANÇA ---
    ' 1ª Confirmação: Intenção geral
    resposta = MsgBox("Você realmente quer executar o botão CANCELAR LINHA?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de uso")
    If resposta <> vbYes Then GoTo SairSemErro

    ' 2ª Confirmação: Quantidade de itens (ajusta o cabeçalho -2)
    resposta = MsgBox("Você realmente quer cancelar " & countNonEmpty - 2 & " pedido(s)?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de cancelamento")
    If resposta <> vbYes Then GoTo SairSemErro

    ' --- LOG DE INÍCIO ---
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1

    With Controle_Macro
        .Range("A" & last_row_macro).Value = "Cancelar Linha"
        .Range("B" & last_row_macro).Value = hoje
        .Range("C" & last_row_macro).Value = horaAtual
        .Range("D" & last_row_macro).Value = usuario
        .Range("E" & last_row_macro).Value = "Iniciada"
    End With

    ' --- EXECUÇÃO ---
    Call bDesbloqueio ' Remove proteção para permitir alterações
    
    ' Chama a sub-rotina de processamento passando o parâmetro de contexto
    ' Esta é a rotina que analisamos anteriormente que gera o Excel de cancelamento
    Call Ir_Cadastro_1("CancelarLinha")
    
    Call bBloqueio ' Restaura a segurança da planilha

    ' --- LOG DE FINALIZAÇÃO ---
    last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row_macro).Value = "Cancelar Linha"
        .Range("B" & last_row_macro).Value = hoje
        .Range("C" & last_row_macro).Value = horaAtual
        .Range("D" & last_row_macro).Value = usuario
        .Range("E" & last_row_macro).Value = "Finalizada"
    End With
    
    MsgBox "Processo de cancelamento concluído com sucesso!", vbInformation, "Pronto"

SairSemErro:
    Application.ScreenUpdating = True
End Sub

