Attribute VB_Name = "AdicionarLinhaClone"
Public CancelamentoSolicitado As Boolean

Sub bExecutarProcessoClone()

    Application.ScreenUpdating = False

        ' Declarações
        Dim ws_Base As Worksheet
        Dim ws_Config As Worksheet
        Dim ws_Log As Worksheet
        Dim uLinha As Long
        Dim uLinhaLog As Long
        Dim rngID As Range
        Dim maiorID As Double
        Dim Colunas_Mapeadas As Variant
        Dim String_IDs As String
        
        ' Atribuições (Nomes de abas anonimizados)
        Set ws_Base = ThisWorkbook.Sheets("DADOS_SISTEMA")
        Set ws_Config = ThisWorkbook.Sheets("PARAMETROS")
        Set ws_Log = ThisWorkbook.Sheets("HISTORICO_ACOES")
        
        ' Variavel booleana
        CancelamentoSolicitado = False
        
        ' Confirmação de execução anonimizada
        resposta = MsgBox("Deseja executar a ação: CLONAR REGISTROS SELECIONADOS?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Operação")
        
        If resposta <> vbYes Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        ' Auditoria
        usuarioID = Environ("Username")
        dataExec = Date
        horaExec = Format(Time, "hh:mm:ss")

        ' Localizar linha para registro de log
        uLinhaLog = ws_Log.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        ' Registro da inicialização da macro
        With ws_Log
            .Range("A" & uLinhaLog).Value = "Ação Clone"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Iniciada"
        End With

        ' Chamadas de subsistemas
        Call Rotina_Validar("bExecutarProcessoClone")
        Call Rotina_Desbloquear
        
        '<<< MAPEAMENTO DINÂMICO DE COLUNAS >>>'
    
        itensMapeamento = Array("Atributo_01", "Atributo_02", "Atributo_03", "Atributo_04", "Atributo_05", "Atributo_06", "Atributo_07", "Status_Ref", "Valor_01", "Valor_02", "Valor_03", "Valor_04", "Valor_05", "Referencia_Ano", "Referencia_Periodo")
        
        For Each Item In itensMapeamento
            For Each Celula In ws_Base.Rows(2).Cells
                If Celula.Value = Item Then
                    Select Case Item
                        Case "Atributo_01": Col_01 = Celula.Column
                        Case "Atributo_02": Col_02 = Celula.Column
                        Case "Atributo_03": Col_03 = Celula.Column
                        Case "Atributo_04": Col_04 = Celula.Column
                        Case "Atributo_05": Col_05 = Celula.Column
                        Case "Atributo_06": Col_06 = Celula.Column
                        Case "Atributo_07": Col_07 = Celula.Column
                        Case "Status_Ref": Col_Status = Celula.Column
                        Case "Valor_01": Col_V1 = Celula.Column
                        Case "Valor_02": Col_V2 = Celula.Column
                        Case "Valor_03": Col_V3 = Celula.Column
                        Case "Valor_04": Col_V4 = Celula.Column
                        Case "Valor_05": Col_V5 = Celula.Column
                        Case "Referencia_Ano": Col_Ano = Celula.Column
                        Case "Referencia_Periodo": Col_Periodo = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next Item
        
        ' Compilar chaves únicas em um dicionário
        Dim objDict As Object
        Set objDict = CreateObject("Scripting.Dictionary")
        
        uLinha = ws_Base.Cells(ws_Base.Rows.Count, "B").End(xlUp).Row
        
        For j = 3 To uLinha
            If ws_Base.Cells(j, "C").Value <> "" Then
                If Not objDict.Exists(ws_Base.Cells(j, "B").Value) Then
                    objDict.Add ws_Base.Cells(j, "B").Value, True
                End If
            End If
        Next j
        
        ' Consolidação dos IDs para processamento
        String_IDs = Join(objDict.Keys, ";")
        qtd_registros = Len(String_IDs) - Len(Replace(String_IDs, ";", "")) + 1
        
        If qtd_registros = 1 Then ' Fluxo para registro individual
            
            uLinha = ws_Base.Cells(Rows.Count, "B").End(xlUp).Row
            
            For i = 3 To uLinha
                If ws_Base.Range("B" & i).Value = String_IDs Then
                    linhaAlvo = ws_Base.Range("B" & i).Row
                    Exit For
                End If
            Next i
                
            ws_Base.Rows(linhaAlvo + 1 & ":" & linhaAlvo + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            ws_Base.Rows(linhaAlvo & ":" & linhaAlvo).Copy
            ws_Base.Rows(linhaAlvo + 1 & ":" & linhaAlvo + 1).PasteSpecial Paste:=xlPasteAll
            
            uLinha = ws_Base.Range("B1").End(xlDown).Row
            Set rngID = ws_Base.Range("B3:B" & uLinha)
            maiorID = WorksheetFunction.Max(rngID)
            
            ws_Base.Range("B" & linhaAlvo + 1) = maiorID + 1
            ws_Base.Cells(linhaAlvo + 1, 2).Interior.Color = RGB(200, 200, 200)
            
            ' Armazenar referência temporária
            ws_Config.Cells(2, 2).Value = linhaAlvo
            
            ' Chamar Formulário de Edição Clone
            UserForm_Clone_Edicao.Show
            
            ' Mapear colunas para atualização via Config
            Colunas_Mapeadas = Array(Col_01, Col_02, Col_03, Col_04, Col_05, Col_06, Col_07, Col_Ano, Col_Periodo)
            
            For i = LBound(Colunas_Mapeadas) To UBound(Colunas_Mapeadas)
                ws_Base.Cells(linhaAlvo + 1, Colunas_Mapeadas(i)) = ws_Config.Cells(2, i + 3).Value
            Next i
            
            ' Limpeza seletiva de campos de valores/indicadores
            uColunaBase = ws_Base.Range("A1").End(xlToRight).Column
            For i = 1 To uColunaBase
                Nome_Col = ws_Base.Cells(2, i).Value
                If Nome_Col = "Campo_Limp_01" Or Nome_Col = "Campo_Limp_02" Or _
                   Nome_Col = "Campo_Limp_03" Or Nome_Col = "Campo_Limp_04" Or Nome_Col = "Campo_Limp_05" Then
                    ws_Base.Cells(linhaAlvo + 1, i).ClearContents
                End If
            Next i
            
            If CancelamentoSolicitado = True Then
                ws_Base.Rows(linhaAlvo + 1).Delete Shift:=xlUp
                MsgBox "Operação Interrompida"
                Exit Sub
            End If
            
        Else ' Fluxo para clonagem em lote
            
            respostaLote = MsgBox("Confirmar clonagem integral de " & qtd_registros & " itens?", vbQuestion + vbYesNo + vbDefaultButton2, "Ação em Lote")
            
            If respostaLote <> vbYes Then
                With ws_Log
                    .Range("A" & uLinhaLog).Value = "Ação Clone"
                    .Range("B" & uLinhaLog).Value = dataExec
                    .Range("C" & uLinhaLog).Value = horaExec
                    .Range("D" & uLinhaLog).Value = usuarioID
                    .Range("E" & uLinhaLog).Value = "Cancelada pelo Usuário"
                End With
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            uLinha = ws_Base.Cells(Rows.Count, "B").End(xlUp).Row
            array_IDs = Split(String_IDs, ";")
            Total_IDs = UBound(array_IDs) + 1
                                    
            For j = 0 To Total_IDs - 1
                ID_Atual = array_IDs(j)
                uLinha = ws_Base.Cells(Rows.Count, "B").End(xlUp).Row
                
                For i = 3 To uLinha
                    If ws_Base.Range("B" & i).Text = ID_Atual Then
                        linhaAlvo = ws_Base.Range("B" & i).Row
                        Exit For
                    End If
                Next i
                
                ws_Base.Rows(linhaAlvo + 1 & ":" & linhaAlvo + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                ws_Base.Rows(linhaAlvo & ":" & linhaAlvo).Copy
                ws_Base.Rows(linhaAlvo + 1 & ":" & linhaAlvo + 1).PasteSpecial Paste:=xlPasteAll
                
                uLinha = ws_Base.Range("B1").End(xlDown).Row
                Set rngID = ws_Base.Range("B3:B" & uLinha)
                maiorID = WorksheetFunction.Max(rngID)
                
                ws_Base.Range("B" & linhaAlvo + 1) = maiorID + 1
                ws_Base.Cells(linhaAlvo + 1, Col_Status).Value = "Processado"
                ws_Base.Cells(linhaAlvo + 1, 2).Interior.Color = RGB(200, 200, 200)
                
                ws_Config.Cells(2, 2).Value = linhaAlvo
                ws_Config.Cells(3, 2).Value = array_IDs(j)
                
                UserForm_Clone_Edicao.Show

                Colunas_Mapeadas = Array(Col_01, Col_02, Col_03, Col_04, Col_05, Col_06, Col_07, Col_Ano, Col_Periodo)
                For i = LBound(Colunas_Mapeadas) To UBound(Colunas_Mapeadas)
                    ws_Base.Cells(linhaAlvo + 1, Colunas_Mapeadas(i)) = ws_Config.Cells(2, i + 3).Value
                Next i

                uColunaBase = ws_Base.Range("A1").End(xlToRight).Column
                For i = 1 To uColunaBase
                    Nome_Col = ws_Base.Cells(2, i).Value
                    If Nome_Col = "Campo_Limp_01" Or Nome_Col = "Campo_Limp_02" Or _
                       Nome_Col = "Campo_Limp_03" Or Nome_Col = "Campo_Limp_04" Or Nome_Col = "Campo_Limp_05" Then
                        ws_Base.Cells(linhaAlvo + 1, i).ClearContents
                    End If
                Next i

                ws_Base.Cells(linhaAlvo + 1, Col_Status).Value = "Processado"
                
                If CancelamentoSolicitado = True Then
                    ws_Base.Rows(linhaAlvo + 1).Delete Shift:=xlUp
                    MsgBox "Operação Interrompida"
                    Exit Sub
                End If
            Next j
        End If

        ' Finalização do processo
        ws_Base.Cells(linhaAlvo + 1, Col_Status).Value = "Processado"
        Call Rotina_Bloquear

        uLinhaLog = ws_Log.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With ws_Log
            .Range("A" & uLinhaLog).Value = "Ação Clone"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Finalizada"
        End With

    MsgBox "Processo concluído!"
    Application.ScreenUpdating = True

End Sub

