Attribute VB_Name = "AdicionarLinhaReposicao"
Sub bExecutarProcessoReposicao()

    Application.ScreenUpdating = False

        ' --- DECLARAÇÕES ---
        Dim ws_Base As Worksheet
        Dim ws_Param As Worksheet
        Dim ws_Log As Worksheet
        Dim ws_Erro As Worksheet
        Dim uLinha As Long
        Dim rngID As Range
        Dim maiorID As Double
        Dim Total_Registros As Integer
        Dim array_IDs() As String
        Dim Calc_Validacao As Double
        Dim ID_Selecionado As String
        Dim objDict As Object
        Dim listaErros As Object
        Dim usuarioID As String
        Dim dataExec As Date
        Dim horaExec As String
        Dim itemMapeado As Variant
        Dim Celula As Range
        Dim uLinhaLog As Long
        Dim uLinhaErro As Long
        Dim linhaAlvo As Long
        Dim i As Long
        Dim j As Long
        Dim nomeBusca As String
        Dim Val_Ref_01 As Double
        Dim Val_Ref_02 As Double
        Dim Val_Ref_03 As Double
        Dim Val_Ref_04 As Double
        Dim Val_Ref_05 As Double
        Dim custo_calc As Double
        Dim msgFinalErro As String
        Dim registroErro As Variant
        Dim PermitirProcesso As Boolean
    
        ' --- DECLARAÇÕES DE COLUNAS (ANONIMIZADAS) ---
        Dim Col_Index As Integer
        Dim Col_Val_01 As Integer
        Dim Col_Val_02 As Integer
        Dim Col_Total_Liq As Integer
        Dim Col_Status_Origem As Integer
        Dim Col_Val_Comp As Integer
        Dim Col_Fator_A As Integer
        Dim Col_Fator_B As Integer
        Dim Col_Categoria As Integer
        Dim Col_Periodo_Ref As Integer
        Dim Col_Data_Ref As Integer
        Dim Col_Alvo_Ref As Integer
        
        ' --- ATRIBUIÇÕES DE ABAS ---
        Set ws_Base = ThisWorkbook.Sheets("BASE_REGISTROS")
        Set ws_Param = ThisWorkbook.Sheets("CONFIGURACOES")
        Set ws_Log = ThisWorkbook.Sheets("LOG_EXECUCAO")
        Set ws_Erro = ThisWorkbook.Sheets("LOG_ERROS")
        Set listaErros = CreateObject("System.Collections.ArrayList")
        
        ' --- CONFIRMAÇÃO ---
        If MsgBox("Deseja executar a ação: ADICIONAR REGISTRO DE REPOSIÇÃO?", vbQuestion + vbYesNo + vbDefaultButton2, "Sistema - Confirmação") <> vbYes Then
            Exit Sub
        End If
    
        ' Auditoria
        usuarioID = Environ("Username")
        dataExec = Date
        horaExec = Format(Time, "hh:mm:ss")
    
        ' Registro de Log Inicial
        uLinhaLog = ws_Log.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With ws_Log
            .Range("A" & uLinhaLog).Value = "Ação Reposição"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Iniciada"
        End With
    
        ' Chamadas de Rotinas do Sistema
        Call Rotina_Validar("")
        Call Rotina_Desbloquear
    
        ' --- MAPEAMENTO DE COLUNAS ---
        itensParaMapear = Array("ID_REF", "VAL_PLAN_01", "VAL_PLAN_02", "TOTAL_LIQUIDO", "ORIGEM_REG", "VAL_EFETIVO", "PERIODO", "FATOR_PACK", "FATOR_UNID", "TIPO_REG", "VAL_ALVO_REF", "DATA_REF", "TARGET_REF")
        
        For Each itemMapeado In itensParaMapear
            nomeBusca = itemMapeado
            For Each Celula In ws_Base.Rows(2).Cells
                If Celula.Value = nomeBusca Then
                    Select Case nomeBusca
                        Case "ID_REF": Col_Index = Celula.Column
                        Case "VAL_PLAN_01": Col_Val_01 = Celula.Column
                        Case "VAL_PLAN_02": Col_Val_02 = Celula.Column
                        Case "TOTAL_LIQUIDO": Col_Total_Liq = Celula.Column
                        Case "ORIGEM_REG": Col_Status_Origem = Celula.Column
                        Case "VAL_EFETIVO": Col_Val_Comp = Celula.Column
                        Case "FATOR_PACK": Col_Fator_A = Celula.Column
                        Case "FATOR_UNID": Col_Fator_B = Celula.Column
                        Case "TIPO_REG": Col_Categoria = Celula.Column
                        Case "PERIODO": Col_Periodo_Ref = Celula.Column
                        Case "DATA_REF": Col_Data_Ref = Celula.Column
                        Case "TARGET_REF": Col_Alvo_Ref = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next itemMapeado
        
        ' Coleta de IDs para processamento
        Set objDict = CreateObject("Scripting.Dictionary")
        uLinha = ws_Base.Cells(ws_Base.Rows.Count, "B").End(xlUp).Row
        uLinhaErro = ws_Erro.Cells(ws_Erro.Rows.Count, "B").End(xlUp).Row
        
        For j = 3 To uLinha
            If ws_Base.Cells(j, "C").Value <> "" Then
                If Not objDict.Exists(ws_Base.Cells(j, "B").Value) Then
                    objDict.Add ws_Base.Cells(j, "B").Value, True
                End If
            End If
        Next j
        
        array_IDs = Split(Join(objDict.Keys, ";"), ";")
        Total_Registros = UBound(array_IDs) + 1
        
        ' --- LOOP DE PROCESSAMENTO ---
        For j = 0 To Total_Registros - 1
            
            linhaAlvo = 0
            For i = 3 To uLinha
                If CStr(ws_Base.Cells(i, "B").Value) = CStr(array_IDs(j)) Then
                    linhaAlvo = i
                    Exit For
                End If
            Next i
            
            PermitirProcesso = True
            
            ' Validação Técnica (Numérica)
            If PermitirProcesso And Not (IsNumeric(ws_Base.Cells(linhaAlvo, Col_Val_Comp).Value) And IsNumeric(ws_Base.Cells(linhaAlvo, Col_Fator_B).Value) And ws_Base.Cells(linhaAlvo, Col_Fator_A).Value <> 0) Then
                With ws_Erro
                    .Range("A" & uLinhaErro).Value = "Erro Processamento - Valores Inválidos ou Nulos"
                    .Range("B" & uLinhaErro).Value = dataExec
                    .Range("C" & uLinhaErro).Value = horaExec
                    .Range("D" & uLinhaErro).Value = usuarioID
                End With
                listaErros.Add "Linha Ref: " & ws_Base.Cells(linhaAlvo, Col_Index).Value & " - Inconsistência de dados numéricos."
                PermitirProcesso = False
            End If
    
            ' Validação de Regra de Negócio (Mínimo Aceitável)
            If PermitirProcesso Then
                Calc_Validacao = ws_Base.Cells(linhaAlvo, Col_Val_Comp).Value / ws_Base.Cells(linhaAlvo, Col_Fator_A).Value
                If Calc_Validacao <= ws_Base.Cells(linhaAlvo, Col_Fator_B).Value Then
                    With ws_Erro
                        .Range("A" & uLinhaErro).Value = "Erro Processamento - Mínimo não atingido"
                        .Range("B" & uLinhaErro).Value = dataExec
                        .Range("C" & uLinhaErro).Value = horaExec
                        .Range("D" & uLinhaErro).Value = usuarioID
                    End With
                    listaErros.Add "Atenção: Volume na linha " & ws_Base.Cells(linhaAlvo, Col_Index).Value & " abaixo do limite operacional."
                    PermitirProcesso = False
                End If
            End If
            
            ' Execução da Clonagem e Redistribuição
            If PermitirProcesso Then
            
                ws_Base.Rows(linhaAlvo + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                ws_Base.Rows(linhaAlvo).Copy
                ws_Base.Rows(linhaAlvo + 1).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                
                ' Gerar Novo Index
                uLinha = ws_Base.Range("B" & ws_Base.Rows.Count).End(xlUp).Row
                Set rngID = ws_Base.Range("B3:B" & uLinha)
                maiorID = WorksheetFunction.Max(rngID)
                
                ws_Base.Range("B" & linhaAlvo + 1) = maiorID + 1
                ws_Base.Cells(linhaAlvo + 1, Col_Status_Origem).Value = "Item_Secundario"
                ws_Base.Cells(linhaAlvo + 1, Col_Categoria).Value = "Individual"
                ws_Base.Cells(linhaAlvo + 1, Col_Alvo_Ref).Value = "Reposicao_Ativa"
                ws_Base.Cells(linhaAlvo + 1, 2).Interior.Color = RGB(200, 200, 200)
                
                ' Lógica de Cálculo de Valores
                Val_Ref_01 = ws_Base.Cells(linhaAlvo, Col_Val_01).Value
                Val_Ref_02 = ws_Base.Cells(linhaAlvo, Col_Val_Comp).Value
                Val_Ref_03 = ws_Base.Cells(linhaAlvo, Col_Fator_A).Value
                Val_Ref_04 = ws_Base.Cells(linhaAlvo, Col_Fator_B).Value
                Val_Ref_05 = ws_Base.Cells(linhaAlvo, Col_Val_02).Value
                
                If ws_Base.Cells(linhaAlvo, Col_Status_Origem).Value = "Novo_Registro" Then
                    custo_calc = 0
                Else
                    If Val_Ref_01 > 0 Then
                        custo_calc = Val_Ref_05 / Val_Ref_01
                    Else
                        custo_calc = 0
                    End If
                End If
                
                ' Redistribuição de Quantidades
                If custo_calc = 0 Then
                    ws_Base.Cells(linhaAlvo, Col_Val_01) = 0
                    ws_Base.Cells(linhaAlvo + 1, Col_Val_01) = 0
                ElseIf Round(Val_Ref_03 * Val_Ref_04, 0) < ws_Base.Cells(linhaAlvo, Col_Val_01) Then
                    ws_Base.Cells(linhaAlvo, Col_Val_01) = Round(Val_Ref_03 * Val_Ref_04, 0)
                    ws_Base.Cells(linhaAlvo + 1, Col_Val_01) = Val_Ref_01 - ws_Base.Cells(linhaAlvo, Col_Val_01)
                Else
                    ws_Base.Cells(linhaAlvo + 1, Col_Val_01) = 0
                End If
                
                ws_Base.Cells(linhaAlvo, Col_Val_Comp) = Round(Val_Ref_03 * Val_Ref_04, 0)
                ws_Base.Cells(linhaAlvo + 1, Col_Val_Comp) = Val_Ref_02 - ws_Base.Cells(linhaAlvo, Col_Val_Comp)
                
                ' Redistribuição Financeira
                If custo_calc = 0 Then
                    ws_Base.Cells(linhaAlvo, Col_Val_02) = 0
                    ws_Base.Cells(linhaAlvo + 1, Col_Val_02) = 0
                Else
                    ws_Base.Cells(linhaAlvo, Col_Val_02) = ws_Base.Cells(linhaAlvo, Col_Val_01) * custo_calc
                    ws_Base.Cells(linhaAlvo + 1, Col_Val_02) = ws_Base.Cells(linhaAlvo + 1, Col_Val_01) * custo_calc
                End If
                
            End If
        Next j
    
        ' Finalização do Sistema
        Call Rotina_Bloquear
        uLinhaLog = ws_Log.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With ws_Log
            .Range("A" & uLinhaLog).Value = "Ação Reposição"
            .Range("B" & uLinhaLog).Value = dataExec
            .Range("C" & uLinhaLog).Value = horaExec
            .Range("D" & uLinhaLog).Value = usuarioID
            .Range("E" & uLinhaLog).Value = "Finalizada"
        End With
        
        ' Relatório de Erros
        If listaErros.Count > 0 Then
            msgFinalErro = "Processamento concluído com exceções:" & vbCrLf & vbCrLf
            For Each registroErro In listaErros
                msgFinalErro = msgFinalErro & "- " & registroErro & vbCrLf
            Next registroErro
            MsgBox msgFinalErro, vbExclamation, "Log de Validação"
        End If

    MsgBox "Operação concluída com sucesso!", vbInformation
    Application.ScreenUpdating = True
    
End Sub


