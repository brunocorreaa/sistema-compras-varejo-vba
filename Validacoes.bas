Attribute VB_Name = "Validacoes"
Sub ValidarIntegridadeDados(ByVal caller As String)

    ' --- DECLARAÇÕES ---
    Dim db As Worksheet, ap As Worksheet, mn As Worksheet, Controle_Erro As Worksheet
    Dim last_row As Long, last_row_erro As Long
    Dim i As Long, carac_Tamanho_ID As Long
    Dim hasfilter As Boolean
    Dim usuario As String
    Dim colDict As Object, Celula As Range
    
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("BASE_PRINCIPAL")
    Set ap = ThisWorkbook.Sheets("Parametros")
    Set mn = ThisWorkbook.Sheets("Painel_Controle")
    Set Controle_Erro = ThisWorkbook.Sheets("Log_Erros")
    
    ' Desbloqueio de segurança
    db.Unprotect "senha_sistema"
    
    ' Auditoria
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    
    ' --- GESTÃO DE FILTROS ---
    hasfilter = False
    If db.AutoFilterMode Then
        If Not db.AutoFilter Is Nothing Then
            If db.AutoFilter.Range.Rows(1).Row = 2 Then hasfilter = True
        End If
    End If
    
    If hasfilter Then
        If db.FilterMode Then db.ShowAllData
    Else
        db.Rows("2:2").AutoFilter
    End If

    ' --- MAPEAMENTO DINÂMICO DE COLUNAS ---
    nomesColunas = Array("ID_Ref", "Status_Registro", "Subconjunto", "Cod_Subconjunto", "Contato_Principal", "Responsavel", "Entidade_Origem", "Entidade_Destino", _
                        "Ciclo_Anual", "Link_Menu", "Ponto_Equilibrio", "Codigo_SKU_Variantes", "Setor_Operacional", "Data_Limite", "Categoria_Tecnica", "Volume_Planejado", _
                        "Custo_Medio", "Valor_Total_Liquido", "Abrangencia_Geografica", "Capacidade_Lote", "Volume_Processado", "Fluxo_Logistico", "Meta_Target", _
                        "Classificacao_Projeto", "Modelo_Distribuicao", "Codigo_Rastreio", "Origem_Entrada", "Segmentacao", "Tipo_Entrada", "Agrupamento_A", "Agrupamento_B", _
                        "Matriz_Escalonamento", "Periodo_Referencia", "Sazonalidade", "Indice_Fiscal", "Dimensao", "Data_Entrada", "Mes_Competencia", "Preco_Unitario_Previsto")
    
    For Each nome In nomesColunas
        For Each Celula In db.Rows(2).Cells
            If Celula.Value = nome Then
                Select Case nome
                    Case "ID_Ref": Col_ID = Celula.Column
                    Case "Status_Registro": Col_Status = Celula.Column
                    Case "Volume_Planejado": Col_VolPlan = Celula.Column
                    Case "Custo_Medio": Col_CustoMed = Celula.Column
                    Case "Valor_Total_Liquido": Col_TotalLiq = Celula.Column
                    Case "Volume_Processado": Col_VolProc = Celula.Column
                    Case "Codigo_Rastreio": Col_Rastreio = Celula.Column
                    Case "Origem_Entrada": Col_Origem = Celula.Column
                    Case "Agrupamento_B": Col_AgrupB = Celula.Column
                    Case "Data_Limite": Col_DataLimite = Celula.Column
                    Case "Dimensao": Col_Dimensao = Celula.Column
                    Case "Matriz_Escalonamento": Col_Matriz = Celula.Column
                    Case "Ponto_Equilibrio": Col_PE = Celula.Column
                    Case "Codigo_SKU_Variantes": Col_SKU_Var = Celula.Column
                    Case "Indice_Fiscal": Col_Fiscal = Celula.Column
                End Select
                Exit For
            End If
        Next Celula
    Next nome
        
    last_row = db.Cells(Rows.Count, "B").End(xlUp).Row
    last_row_erro = Controle_Erro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    ' --- VALIDAÇÕES DE INTEGRIDADE ---

    ' V01: Check de Ativação Global
    If ap.Range("AO2").Value = 0 Then
        mn.Visible = xlSheetVeryHidden
        RegistrarErro Controle_Erro, "Sistema desativado por inconsistência crítica.", hoje, horaAtual, usuario
        InterromperProcesso "O sistema está bloqueado. Reinicie a aplicação."
    End If
    
    ' V02: Prevenção de Valores Negativos
    For i = 3 To last_row
        If db.Cells(i, Col_VolPlan).Value < 0 Or db.Cells(i, Col_CustoMed).Value < 0 Or _
           db.Cells(i, Col_TotalLiq).Value < 0 Or db.Cells(i, Col_VolProc).Value < 0 Then
           
            RegistrarErro Controle_Erro, "Valores negativos detectados em colunas quantitativas.", hoje, horaAtual, usuario
            InterromperProcesso "Erro na linha " & db.Cells(i, Col_ID).Value & ": Não são permitidos valores negativos."
        End If
    Next i

    ' V03: Verificação de Comprimento de Identificador (Código Rastreio)
    For i = 3 To last_row
        carac_Tamanho_ID = Len(db.Cells(i, Col_Rastreio).Text)
        ' Critérios de comprimento: 16, 10, 1 ou vazio
        If carac_Tamanho_ID <> 16 And carac_Tamanho_ID <> 10 And carac_Tamanho_ID <> 0 And carac_Tamanho_ID <> 1 And db.Cells(i, Col_Rastreio).Text <> "" Then
            RegistrarErro Controle_Erro, "Comprimento inválido de identificador de rastreio.", hoje, horaAtual, usuario
            InterromperProcesso "Erro na linha " & db.Cells(i, Col_ID).Value & ": Código de rastreio fora do padrão técnico."
        End If
    Next i

    ' V04: Consistência de Origem x Agrupamento
    For i = 3 To last_row
        If db.Cells(i, Col_Origem).Value = "Entrada_Manual" And db.Cells(i, Col_AgrupB).Value <> "" Then
            RegistrarErro Controle_Erro, "Conflito de origem: Registro manual não deve conter Agrupamento_B.", hoje, horaAtual, usuario
            InterromperProcesso "Erro na linha " & db.Cells(i, Col_ID).Value & ": Limpe o Agrupamento_B para entradas manuais."
        End If
    Next i

    ' V05: Validação de Cronograma (Datas)
    For i = 3 To last_row
        If db.Cells(i, Col_ID).Value <> "" Then
            If caller = "Operacao_Escrita" Or caller = "Modificacao" Then
                If Not IsDate(db.Cells(i, Col_DataLimite).Value) Then
                    RegistrarErro Controle_Erro, "Data inválida ou ausente.", hoje, horaAtual, usuario
                    InterromperProcesso "Verifique a data na linha " & db.Cells(i, Col_ID).Value
                End If
            End If
        End If
    Next i

    ' V07: Sincronismo de Delimitadores (Grades)
    ' Garante que o número de subitens (separados por ;) seja consistente entre colunas relacionadas
    For i = 3 To last_row
        qtd_dim = Len(db.Cells(i, Col_Dimensao)) - Len(Replace(db.Cells(i, Col_Dimensao).Value, ";", ""))
        qtd_mat = Len(db.Cells(i, Col_Matriz)) - Len(Replace(db.Cells(i, Col_Matriz).Value, ";", ""))
        
        If qtd_dim <> qtd_mat Then
            RegistrarErro Controle_Erro, "Divergência na contagem de subitens da Matriz/Dimensão.", hoje, horaAtual, usuario
            InterromperProcesso "Inconsistência de grade na linha " & db.Cells(i, Col_ID).Value
        End If
    Next i

    ' V12: Campos Fiscais/Obrigatórios
    For i = 3 To last_row
        If db.Cells(i, Col_ID).Value <> "" And db.Cells(i, Col_Fiscal).Value = "" Then
            RegistrarErro Controle_Erro, "Campo de índice fiscal obrigatório vazio.", hoje, horaAtual, usuario
            InterromperProcesso "O preenchimento do Índice Fiscal é obrigatório na linha " & db.Cells(i, Col_ID).Value
        End If
    Next i
    
    ' V13: Integridade de Lote: Volume vs. Unidade de Agrupamento
    For i = 3 To last_row
        If db.Cells(i, Col_ID).Value <> "" And db.Cells(i, Col_UnidadeLote).Value <> 0 And caller <> "RemoverRegistro" Then
            ' Verifica se o Volume Processado é múltiplo exato da Unidade do Lote (Pack)
            If (db.Cells(i, Col_VolProc).Value / db.Cells(i, Col_UnidadeLote).Value) <> Int(db.Cells(i, Col_VolProc).Value / db.Cells(i, Col_UnidadeLote).Value) Then

                RegistrarErro Controle_Erro, "Volume processado não é múltiplo da unidade do lote", hoje, horaAtual, usuario
                
                mn.Select
                Call bBloqueio
                MsgBox "Atenção! O volume processado deve ser múltiplo da Unidade de Lote na linha " & db.Cells(i, Col_ID).Value, vbExclamation, "Erro de Integridade"
                
                InterromperProcesso "Falha no cálculo de múltiplo de lote."
            End If
        End If
    Next i
    
    ' V14: Regras de Distribuição por Tipo de Fluxo (PE)
    For i = 3 To last_row

        ' Caso Tipo de Fluxo: DISTRIBUICAO_DIRETA - Requer neutralidade nos campos de escalonamento (PE)
        If db.Cells(i, Col_TipoFluxo).Value = "DISTRIBUICAO_DIRETA" And db.Cells(i, Col_ID).Value <> "" Then
            PE_valores = Split(db.Cells(i, Col_PE).Value, ";")
            
            For j = LBound(PE_valores) To UBound(PE_valores)
                If PE_valores(j) <> "0" And Len(Trim(PE_valores(j))) > 0 Then
                    RegistrarErro Controle_Erro, "Fluxo direto exige valores zerados ou vazios na distribuição", hoje, horaAtual, usuario
                    
                    mn.Select
                    Call bBloqueio
                    MsgBox "Atenção! Para Distribuição Direta, a grade aceita apenas zeros ou vazio na linha " & db.Cells(i, Col_ID).Value, vbExclamation, "Erro de Regra de Negócio"
                    InterromperProcesso "Inconsistência em Fluxo Direto."
                End If
            Next j
        End If
        
        ' Caso Tipo de Fluxo: LOTE_PADRAO - Requer valores inteiros e simétricos
        If db.Cells(i, Col_TipoFluxo).Value = "LOTE_PADRAO" And db.Cells(i, Col_ID).Value <> "" Then
            PE_valores = Split(db.Cells(i, Col_PE).Value, ";")
            primeiro_valor = PE_valores(0)
            
            For j = LBound(PE_valores) To UBound(PE_valores)
                If Not IsNumeric(PE_valores(j)) Or InStr(PE_valores(j), ".") > 0 Or PE_valores(j) <> primeiro_valor Then
                    RegistrarErro Controle_Erro, "Lote Padrão exige valores inteiros e simétricos", hoje, horaAtual, usuario
                    
                    mn.Select
                    Call bBloqueio
                    MsgBox "Atenção! Para Lote Padrão, a grade deve conter apenas números inteiros idênticos na linha " & db.Cells(i, Col_ID).Value, vbExclamation, "Erro de Simetria"
                    InterromperProcesso "Inconsistência em Lote Padrão."
                End If
            Next j
        End If
    Next i
    
    ' V15: Bloqueio de Registros com Status "Inativo/Cancelado"
    For i = 3 To last_row
        If Not (caller = "DuplicarRegistro" Or caller = "ExportarRelatorio") And _
            db.Cells(i, Col_Status).Value = "CANCELADO" And (Not IsEmpty(db.Cells(i, Col_ID).Value)) Then

            RegistrarErro Controle_Erro, "Tentativa de processar registro inativo", hoje, horaAtual, usuario
            
            mn.Select
            Call bBloqueio
            MsgBox "Atenção! O registro na linha " & db.Cells(i, Col_ID).Value & " está marcado como CANCELADO e não permite esta operação.", vbExclamation, "Status Inválido"
            InterromperProcesso "Operação não permitida para registros inativos."
        End If
    Next i
    
    ' V16 -- Domínio de Valores para Coluna de Origem (Input)
    If caller = "FinalizarFluxo" Or caller = "ProcessarRegistro" Or caller = "AtualizarDados" Then
        For i = 3 To last_row
            If db.Cells(i, Col_ID).Value <> "" And _
                Not (db.Cells(i, Col_Origem).Value = "RECORRENTE" Or db.Cells(i, Col_Origem).Value = "PROJETO") Then
                
                ' Verifica se o setor operacional está no escopo permitido
                If Not IsError(Application.Match(UCase(db.Cells(i, Col_SetorOperacional).Value), _
                    Array("OP_01", "OP_02", "OP_03", "OP_04", "OP_LOG", "OP_ESP"), 0)) Then

                    RegistrarErro Controle_Erro, "Origem de entrada inválida", hoje, horaAtual, usuario
                    
                    mn.Select
                    Call bBloqueio
                    MsgBox "Atenção! A origem na linha " & db.Cells(i, Col_ID).Value & " deve ser 'RECORRENTE' ou 'PROJETO'.", vbExclamation, "Erro de Domínio"
                    InterromperProcesso "Valor de entrada fora do padrão."
                End If
            End If
        Next i
    End If
    
    ' V17 -- Cronologia: Data de Execução vs. Data Atual
    If caller = "ProcessarRegistro" Then
        For i = 3 To last_row
            If db.Cells(i, Col_ID).Value <> "" And Date > db.Cells(i, Col_DataLimite).Value Then
                RegistrarErro Controle_Erro, "Data de execução retroativa detectada", hoje, horaAtual, usuario
                
                mn.Select
                Call bBloqueio
                MsgBox "Atenção! Operação bloqueada. A data de execução na linha " & db.Cells(i, Col_ID).Value & " é anterior à data atual.", vbExclamation, "Erro Cronológico"
                InterromperProcesso "Data retroativa não permitida."
            End If
        Next i
    End If
    
    ' V18 -- Controle Orçamentário (Saldo de Verba/Budget)
    If caller = "ProcessarRegistro" Then
        For i = 3 To last_row
            If db.Cells(i, Col_ID).Value <> "" Then
                setor = db.Cells(i, Col_SetorOperacional).Value
                MesVariavel = Format(db.Cells(i, Col_DataLimite).Value, "mmmm")
                AnoVariavel = Year(db.Cells(i, Col_DataLimite).Value)
                chaveOrcamento = AnoVariavel & MesVariavel & setor

                ' Busca saldo disponível na matriz de governança
                On Error Resume Next
                saldoDisponivel = WorksheetFunction.VLookup(chaveOrcamento, Sheets("Matriz_Orcamentaria").Range("A:N"), 14, 0)
                If Err.Number <> 0 Then saldoDisponivel = 0: Err.Clear
                On Error GoTo 0
                
                ValorProcessamento = 0
                MesNum = Month(db.Cells(i, Col_DataLimite).Value)
                
                ' Consolidação do volume financeiro do período para validação de teto
                For k = 3 To last_row
                    If IsDate(db.Cells(k, Col_DataLimite).Value) Then
                        If Month(db.Cells(k, Col_DataLimite).Value) = MesNum And db.Cells(k, Col_ID).Value <> "" And setor = db.Cells(k, Col_SetorOperacional).Value Then
                            ValorProcessamento = ValorProcessamento + CDbl(db.Cells(k, Col_TotalLiq).Value)
                        End If
                    End If
                Next k
                                     
                ' Validação de estouro de teto orçamentário
                If saldoDisponivel < 0 Or (saldoDisponivel - ValorProcessamento) < 0 Then
                    RegistrarErro Controle_Erro, "Excesso de limite orçamentário", hoje, horaAtual, usuario
                    
                    mn.Select
                    Call bBloqueio
                    MsgBox "Atenção! Limite orçamentário excedido para " & MesVariavel & "/" & AnoVariavel & vbCrLf & _
                           "Disponível: " & Format(saldoDisponivel, "Currency") & vbCrLf & _
                           "Solicitado: " & Format(ValorProcessamento, "Currency"), vbCritical, "Bloqueio de Governança"
                    InterromperProcesso "Budget insuficiente."
                End If
            End If
        Next i
    End If

    ' V19 -- Consistência de Status de Transmissão (Check Emitido)
    If caller = "FinalizarFluxo" Or caller = "AtualizarDados" Then
        For i = 3 To last_row
            If db.Cells(i, Col_StatusTransmissao).Value = "Não" And db.Cells(i, Col_ID).Value <> "" Then
                RegistrarErro Controle_Erro, "Tentativa de alteração em registro não transmitido", hoje, horaAtual, usuario
                mn.Select: Call bBloqueio
                MsgBox "Operação negada. O registro " & db.Cells(i, Col_ID).Value & " ainda não foi transmitido para o banco central.", vbExclamation
                InterromperProcesso "Status de transmissão incompatível."
            End If
        Next i
    ElseIf caller = "ProcessarRegistro" Or caller = "RemoverRegistro" Then
        For i = 3 To last_row
            If db.Cells(i, Col_StatusTransmissao).Value = "Sim" And db.Cells(i, Col_ID).Value <> "" Then
                RegistrarErro Controle_Erro, "Tentativa de reprocessar registro já transmitido", hoje, horaAtual, usuario
                mn.Select: Call bBloqueio
                MsgBox "Operação negada. O registro " & db.Cells(i, Col_ID).Value & " já consta como transmitido.", vbExclamation
                InterromperProcesso "Registro já consolidado."
            End If
        Next i
    End If
    
    ' V20/21/22 -- Integridade de Categorização e Planejamento
    For i = 3 To last_row
        ' V20: Campos Mandatórios de Hierarquia
        If db.Cells(i, Col_Subconjunto).Value = "" Or db.Cells(i, Col_CodSubconjunto).Value = "" Then
            RegistrarErro Controle_Erro, "Campos de subconjunto mandatórios vazios", hoje, horaAtual, usuario
            InterromperProcesso "Erro na linha " & db.Cells(i, Col_ID).Value & ": Classificação ausente."
        End If
        
        ' V21: Validação de Registros do tipo "PLANEJAMENTO"
        If db.Cells(i, Col_OrigemStatus).Value = "PLANEJAMENTO" Then
            If Trim(db.Cells(i, Col_RefCML).Value) = "" Or db.Cells(i, Col_RefCML).Value = 0 Then
                RegistrarErro Controle_Erro, "Registro de planejamento sem referência de controle", hoje, horaAtual, usuario
                InterromperProcesso "Linha " & db.Cells(i, Col_ID).Value & ": Referência CML obrigatória para planos."
            End If
        End If

        ' V22: Restrição de Dados em Registros "AVULSOS"
        If db.Cells(i, Col_OrigemStatus).Value = "AVULSO" Then
            ' Se for avulso, colunas de planejamento devem estar zeradas/vazias
            If db.Cells(i, Col_RefCML).Value <> 0 Or db.Cells(i, Col_VolPlan).Value <> 0 Then
                RegistrarErro Controle_Erro, "Dados de planejamento em registro avulso", hoje, horaAtual, usuario
                InterromperProcesso "Linha " & db.Cells(i, Col_ID).Value & ": Registros avulsos não devem conter dados de plano."
            End If
        End If
    Next i

End Sub

' --- SUBS DE APOIO ---
Private Sub RegistrarErro(ws As Worksheet, msg As String, dt As Variant, hr As Variant, user As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    ws.Range("A" & r).Value = msg
    ws.Range("B" & r).Value = dt
    ws.Range("C" & r).Value = hr
    ws.Range("D" & r).Value = user
End Sub

Private Sub InterromperProcesso(msg As String)
    MsgBox msg, vbCritical, "Erro de Validação de Dados"
    Call bBloqueio ' Rotina externa de travamento
    Err.Raise vbObjectError + 513, , "Processo interrompido por falha de integridade."
End Sub
