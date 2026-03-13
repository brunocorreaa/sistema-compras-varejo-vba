Attribute VB_Name = "AdicionarLinha"
Sub bAcaoPrincipal()

    Application.ScreenUpdating = False

        ' Declarações
        Dim ws_Dados As Worksheet
        Dim ws_Param As Worksheet
        Dim ws_Registro As Worksheet
        Dim uLinha As Long
        Dim uLinhaLog As Long
        
        ' Atribuições (Nomes de abas anonimizados)
        Set ws_Dados = ThisWorkbook.Sheets("BASE_PRINCIPAL")
        Set ws_Param = ThisWorkbook.Sheets("PARAMETROS")
        Set ws_Registro = ThisWorkbook.Sheets("LOG_SISTEMA")
        
        ' Variavel booleana
        CancelamentoSolicitado = False
        
        ' Confirmar execução
        resposta = MsgBox("Deseja executar a ação: ADICIONAR NOVO REGISTRO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação")
        
        If resposta <> vbYes Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        ' Dados de auditoria anonimizados
        usuarioID = Environ("Username")
        dataExec = Date
        horaExec = Format(Time, "hh:mm:ss")

        ' Localizar última linha de log
        uLinhaLog = ws_Registro.Cells(ws_Registro.Rows.Count, "B").End(xlUp).Row + 1
        
        ' Registro da inicialização
        With ws_Registro
            .Range("A" & uLinhaLog).Value = "Ação_Novo_Item"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Iniciada"
        End With
        
        ' Chamadas de subsistema
        Call Rotina_Validar("")
        Call Rotina_Acesso_Livre
                
        '<<< MAPEAMENTO DE COLUNAS >>>'
    
        nomesColunas = Array("Info_01", "Info_02", "Info_03", "Info_04", "Info_05", "Info_06", "Info_07", "Info_08", "Info_09", "Info_10", "Info_11")
        
       For Each nome In nomesColunas
            For Each Celula In ws_Dados.Rows(2).Cells
                If Celula.Value = nome Then
                    Select Case nome
                        Case "Info_01": Col_01 = Celula.Column
                        Case "Info_02": Col_02 = Celula.Column
                        Case "Info_03": Col_03 = Celula.Column
                        Case "Info_04": Col_04 = Celula.Column
                        Case "Info_05": Col_05 = Celula.Column
                        Case "Info_06": Col_06 = Celula.Column
                        Case "Info_07": Col_07 = Celula.Column
                        Case "Info_08": Col_08 = Celula.Column
                        Case "Info_09": Col_09 = Celula.Column
                        Case "Info_10": Col_10 = Celula.Column
                        Case "Info_11": Col_11 = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next nome
    
        ' Ultima Linha preenchida
        uLinha = ws_Dados.Cells(ws_Dados.Rows.Count, "B").End(xlUp).Row
        
        ' Duplicação da estrutura
        ws_Dados.Rows(uLinha).AutoFill Destination:=ws_Dados.Rows(uLinha & ":" & uLinha + 1), Type:=xlFillCopy
        
        ' Gerar novo ID sequencial
        Set rngID = ws_Dados.Range("B3:B" & uLinha)
        novoID = WorksheetFunction.Max(rngID)
        ws_Dados.Range("B" & uLinha + 1) = novoID + 1
    
        ' Limpar campos de entrada (conforme marcadores na linha 1)
        uColuna = ws_Dados.Range("A1").End(xlToRight).Column
        
        For i = 1 To uColuna
            marcador = ws_Dados.Cells(1, i).Value
            If (marcador = "VAR1" Or marcador = "VAR2" Or marcador = "VAR3" Or marcador = "VAR4") And ws_Dados.Cells(2, i) <> "ID_REF" Then
                ws_Dados.Cells(uLinha + 1, i).ClearContents
            End If
        Next i
        
        ' Interface de entrada
        UserForm_Entrada_Dados.Show
        
        ' Transferência de valores anonimizada
        Cols_Array = Array(Col_02, Col_03, Col_04, Col_05, Col_06, Col_07, Col_08, Col_09, Col_10, Col_11)
        
        For i = LBound(Cols_Array) To UBound(Cols_Array)
            ws_Dados.Cells(uLinha + 1, Cols_Array(i)) = ws_Param.Cells(2, i + 3).Value
        Next i
        
        ' Status de origem
        ws_Dados.Cells(uLinha + 1, Col_01).Value = "Processado"
        
        ' Formatação visual do novo registro
        ws_Dados.Cells(uLinha + 1, 2).Interior.Color = RGB(200, 200, 200)
        
        ' Tratamento de cancelamento
        If CancelamentoSolicitado = True Then
            ws_Dados.Rows(uLinha + 1).Delete Shift:=xlUp
        End If
        
        ' Restrição de acesso
        Call Rotina_Acesso_Restrito

        ' Log de conclusão
        uLinhaLog = ws_Registro.Cells(ws_Registro.Rows.Count, "B").End(xlUp).Row + 1
        With ws_Registro
            .Range("A" & uLinhaLog).Value = "Ação_Novo_Item"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Finalizada"
        End With
    
    ' Feedback ao usuário
    If CancelamentoSolicitado = True Then
        MsgBox "Operação interrompida.", vbInformation
    Else
        MsgBox "Concluído com sucesso.", vbInformation
    End If
    
    Application.ScreenUpdating = True

End Sub
