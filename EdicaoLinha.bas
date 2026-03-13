Attribute VB_Name = "EdicaoLinha"
Sub EdicaoLinha()

    Application.ScreenUpdating = False
    
        ' Declarações
        Dim db As Worksheet
        Dim mn As Worksheet
        Dim Controle_Macro As Worksheet
        Dim countNonEmpty As Long
        Dim cell As Range
        Dim lastRow As Long
        Dim resposta As Integer
    
        ' Atribuições (Nomes de abas anonimizados)
        Set db = ThisWorkbook.Sheets("BASE_DADOS")
        Set mn = ThisWorkbook.Sheets("Menu")
        Set Controle_Macro = ThisWorkbook.Sheets("LOG_SISTEMA")

        ' Determina a última linha da coluna de controle
        lastRow = db.Cells(db.Rows.Count, "C").End(xlUp).Row
    
        ' Inicializa o contador
        countNonEmpty = 0
    
        ' Loop através das células para contagem de registros selecionados
        For Each cell In db.Range("C1:C" & lastRow)
            If Len(Trim(cell.Value)) > 0 Then
                countNonEmpty = countNonEmpty + 1
            End If
        Next cell
    
        ' Validação de execução
        resposta = MsgBox("Você realmente quer executar a operação: EDITAR REGISTRO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação")
        
        If resposta <> vbYes Then
            Exit Sub
        End If
    
        ' Validação de quantidade
        resposta = MsgBox("Você realmente quer editar " & countNonEmpty - 2 & " item(s)?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Alteração")
        
        If resposta <> vbYes Then
            Exit Sub
        End If
        
        ' Auditoria
        usuario = Environ("Username")
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")

        ' Registro de Log
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Edicao Registro"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Iniciada"
        End With
        
        ' Processos de suporte
        Call Validacoes("EdicaoRegistro")
        Call bDesbloqueio

        ' Loop de processamento de linhas
        last_row = db.Cells(Rows.Count, "B").End(xlUp).Row
        For j = 3 To last_row
        
            checkAcao = db.Cells(j, 3).Value
            If checkAcao <> "" Then
                
                ' Identificação da linha para o formulário
                Valor_Linha = db.Cells(j, 2).Row
                Valor_ID = db.Cells(j, 2).Value
                
                ' Passagem de parâmetros para a aba auxiliar
                mn.Cells(3, 2).Value = Valor_Linha
                mn.Cells(3, 3).Value = Valor_ID
                
                ' Chamada da interface (Nome do form anonimizado)
                form_InterfaceEdicao.Show
                
            End If
        
        Next j
        
        ' Execução da rotina de atualização
        Call Rotina_Processamento_Interno("EdicaoRegistro")

        ' Segurança e Finalização
        Call bBloqueio
    
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Edicao Registro"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Finalizada"
        End With
    
    MsgBox "Processo concluído!"
    
    Application.ScreenUpdating = True

End Sub
