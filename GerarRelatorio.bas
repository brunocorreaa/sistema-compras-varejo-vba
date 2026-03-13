Attribute VB_Name = "GerarRelatorio"
Sub bGerarRelatorio()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' Declaracoes
    Dim db As Worksheet, mn As Worksheet, ap As Worksheet
    Dim Controle_Macro As Worksheet, Controle_Erro As Worksheet, rd As Worksheet
    Dim arquivo_Envio As Workbook
    Dim lastRow As Long, resposta As Integer
    Dim usuario As String, hoje As Date, horaAtual As String
    Dim caminho_pasta As String, Grupo_Selecionado As String
    Dim planilha As Workbook
    Dim wsCompilado As Worksheet
    Dim ult_linha_Compilado As Long, ult_linha_Plano As Long, ult_linha_Plano_filtro As Long
    Dim col As Variant
    Dim colunasParaConverter As Variant
    Dim cellWidth As Double
    Dim cellHeight As Double
    Dim imgAspectRatio As Double
    Dim cellAspectRatio As Double
    Dim wbNovo As Workbook
    Dim nomeArquivoNovo As String
    Dim caminhoDiretorio As String
    Dim cp As Worksheet
    Dim nome, nomeProcurado, nomesColunas() As Variant
    Dim imagemInserida As Boolean
    Dim Coluna As Long
    Dim diretorio As String
    Dim img As Picture
    Dim cell As Range
    Dim arrRangesL1 As Variant
    Dim arrRangesL2 As Variant
    Dim arrColors As Variant
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim CaminhoArquivo As String
    Dim nomeArquivo As String
    Dim abasParaCopiar As Variant
    Dim aba As Variant
    Dim ultLinhaOrigem As Long
    Dim rngDadosCopia As Range
    Dim rngDadosLimpeza As Range
    Dim ultLinhaDestino As Long
    Dim linhasParaExcluir As Long
    Dim abasParaConverter As Variant
    Dim wsFonte As Worksheet
    Dim ultLinha As Long, ultColuna As Long
    Dim dados As Variant, header As Variant
    Dim dictAgrupado As Object
    Dim i As Long, j As Long
    Dim chaveAgrup As Variant
    Dim dictLinha As Object
    Dim colAgrupamento As Variant, colSomar As Variant
    Dim idxCol As Object
    Dim linhaDestino As Long
    Dim codigo As Variant, lbValor As Double
    Dim codigoBusca As Variant
    Dim linhaCompilado As Long
    Dim valorReferencia As Variant
    Dim valorLB As Variant
    Dim mediaVendaDia As Double
    Dim Grupos_Selecionados_Array() As String, Classes_Selecionadas_Array() As String
    Dim G As Variant, c As Variant
    Dim colA_Fornecedor As Long, colA_Marca As Long, colA_Setor As Long, colA_Grupo As Long, colA_Classe As Long
    Dim colA_Subclasse As Long, colA_Codigo As Long, colA_ClusterCompra As Long, colA_EstacaoProduto As Long
    Dim colA_PrecoVendas As Long, colA_QuantidadeComprar As Long, colA_Observacao As Long
    Dim colB_Fornecedor As Long, colB_Marca As Long, colB_Setor As Long, colB_Grupo As Long, colB_Classe As Long
    Dim colB_Subclasse As Long, colB_Codigo As Long, colB_ClusterCompra As Long, colB_EstacaoProduto As Long
    Dim colB_PrecoVendas As Long, colB_QuantidadeComprar As Long, colB_Observacao As Long
    Dim cabecalhoA As Variant, cabecalhoB As Variant
    Dim lastRowB As Long, lastRowA As Long
    Dim repetRow As Long
    Dim anoAtual As Long
    Dim semestreAtual As Long
    Dim strGruposSelecionados As String
    Dim strClassesSelecionadas As String
    Dim strAnoSelecionado As String
    Dim strSemestreSelecionado As String
    Dim Coluna_Grupo As Long
    Dim Coluna_Classe As Long
    Dim Coluna_OrigemModelo As Long
    Dim cell_header As Range
    Dim Acoes_Selecionadas_Array() As String
    Dim Status_Selecionados_Array() As String
    Dim A As Variant, S As Variant

    ' Definir referências às planilhas (Nomes de abas mantidos conforme estrutura técnica)
    Set db = ThisWorkbook.Sheets("BASE_DADOS_PRINCIPAL")
    Set mn = ThisWorkbook.Sheets("Painel_Controle")
    Set ap = ThisWorkbook.Sheets("Config_Apoio")
    Set Controle_Macro = Sheets("Log_Execucao")
    Set Controle_Erro = ThisWorkbook.Sheets("Log_Erros")
    Set rd = ThisWorkbook.Sheets("Analise_Performance")

    ' Confirmação do usuário
    resposta = MsgBox("Você realmente quer executar o botão GERAR RELATÓRIO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Processamento")
    If resposta <> vbYes Then Exit Sub

    ' Dados do usuário e horário
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")

    ' Ultima linha do log
    last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1

    ' Registrar início da macro
    With Controle_Macro
        .Range("A" & last_row_macro).Value = "Gerar Relatorio"
        .Range("B" & last_row_macro).Value = hoje
        .Range("C" & last_row_macro).Value = horaAtual
        .Range("D" & last_row_macro).Value = usuario
        .Range("E" & last_row_macro).Value = "Iniciada"
    End With

    ' Desoculta colunas para processamento
    db.Columns.Hidden = False
    
    ' Executar validações de sistema
    Call Validacoes("GerarRelatorio")

    ' Chamar procedimentos de segurança
    Call bDesbloqueio
    
    ' Interface de seleção
    form_FiltroSelecaoRelatorio.Show
    
    ' Definição de caminhos locais
    caminho_pasta = ThisWorkbook.Path & "\"
    Set Compilado = Workbooks.Open(caminho_pasta & "Arquivo_Consolidado.xlsm")
    Set wsCompilado = Compilado.Sheets("Dados_Consolidados")
    Set wsVendas = Compilado.Sheets("DB_Historico_Vendas")
    Set wsRecebimentos = Compilado.Sheets("DB_Recebimentos")
    Set wsCarteira = Compilado.Sheets("DB_Carteira_Pedidos")
    Set wsAnalitico = Compilado.Sheets("DB_Analitico")
    Set wsVisaoFinal = Compilado.Sheets("Relatorio_Final")
    
    wsCompilado.Select
    
    ' Reset de dados temporários
    ult_linha_Compilado = wsCompilado.Cells(wsCompilado.Rows.Count, "B").End(xlUp).Row
    If ult_linha_Compilado >= 3 Then wsCompilado.Rows("3:" & ult_linha_Compilado).ClearContents
    
    ' Coleta de parâmetros definidos no Painel
    strGruposSelecionados = mn.Cells(3, 3).Value
    strClassesSelecionadas = mn.Cells(3, 4).Value
    strAcoesSelecionadas = mn.Cells(3, 5).Value
    strStatusSelecionados = mn.Cells(3, 6).Value
    strAnoSelecionado = mn.Cells(3, 7).Value
    strSemestreSelecionado = mn.Cells(3, 8).Value
    
    ' Validações de preenchimento
    If strGruposSelecionados = "" And strClassesSelecionadas = "" Then
        MsgBox "Nenhum parâmetro de grupo foi selecionado.", vbExclamation
        End ' Interrupção segura
    End If
    
    If strAnoSelecionado = "" Or strSemestreSelecionado = "" Then
        If Not Compilado Is Nothing Then
            Compilado.Close SaveChanges:=False
            Set Compilado = Nothing
        End If
        mn.Select
        Call bBloqueio
        MsgBox "Ano ou Semestre não selecionado.", vbExclamation
        End
    End If

    ' Tratamento de Arrays para Filtros
    If strGruposSelecionados <> "" Then
        Grupos_Selecionados_Array = Split(strGruposSelecionados, ",")
    Else
        ReDim Grupos_Selecionados_Array(0): Grupos_Selecionados_Array(0) = "*"
    End If

    If strClassesSelecionadas <> "" Then
        Classes_Selecionadas_Array = Split(strClassesSelecionadas, ",")
    Else
        ReDim Classes_Selecionadas_Array(0): Classes_Selecionadas_Array(0) = "*"
    End If
    
    ' Tratamento de Status e Ações (conversão de vazios)
    If strAcoesSelecionadas <> "" Then
        Acoes_Selecionadas_Array = Split(strAcoesSelecionadas, ",")
        For A = LBound(Acoes_Selecionadas_Array) To UBound(Acoes_Selecionadas_Array)
            If Acoes_Selecionadas_Array(A) = "" Then Acoes_Selecionadas_Array(A) = ""
        Next A
    Else
        ReDim Acoes_Selecionadas_Array(0): Acoes_Selecionadas_Array(0) = ""
    End If
    
    If strStatusSelecionados <> "" Then
        Status_Selecionados_Array = Split(strStatusSelecionados, ",")
        For S = LBound(Status_Selecionados_Array) To UBound(Status_Selecionados_Array)
            If Status_Selecionados_Array(S) = "" Then Status_Selecionados_Array(S) = ""
        Next S
    Else
        ReDim Status_Selecionados_Array(0): Status_Selecionados_Array(0) = ""
    End If

    ' Mapeamento dinâmico de cabeçalhos
    For Each cell_header In db.Rows(2).Cells
        Select Case cell_header.Value
            Case "Setor": Coluna_Setor = cell_header.Column
            Case "Grupo": Coluna_Grupo = cell_header.Column
            Case "Classe": Coluna_Classe = cell_header.Column
            Case "Ano": Coluna_Ano = cell_header.Column
            Case "Semestre": Coluna_Semestre = cell_header.Column
            Case "Status_Item": Coluna_OrigemModelo = cell_header.Column
        End Select
        If Coluna_Setor > 0 And Coluna_Grupo > 0 And Coluna_Classe > 0 And _
           Coluna_OrigemModelo > 0 And Coluna_Semestre > 0 And Coluna_Ano > 0 Then Exit For
    Next cell_header
    
    ' Verificação de integridade da estrutura da planilha
    If Coluna_Setor = 0 Or Coluna_Grupo = 0 Or Coluna_Classe = 0 Or _
       Coluna_OrigemModelo = 0 Or Coluna_Ano = 0 Or Coluna_Semestre = 0 Then
        If Not Compilado Is Nothing Then
            Compilado.Close SaveChanges:=False
            Set Compilado = Nothing
        End If
        mn.Select: Call bBloqueio
        MsgBox "Estrutura de colunas da base de dados não identificada.", vbCritical
        End
    End If
    
    ' Aplicação de Filtros Avançados na Base
    ult_linha_Plano = db.Cells(db.Rows.Count, "B").End(xlUp).Row
    
    With db
        If Not .AutoFilterMode Then .Rows(2).AutoFilter
        If .FilterMode Then .ShowAllData
        
        ' Filtros por parâmetros selecionados
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=Coluna_Grupo, Criteria1:=Grupos_Selecionados_Array, Operator:=xlFilterValues
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=Coluna_Classe, Criteria1:=Classes_Selecionadas_Array, Operator:=xlFilterValues
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=8, Criteria1:="<>" ' Filtro de Código Identificador
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=Coluna_Ano, Criteria1:=strAnoSelecionado
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=Coluna_Semestre, Criteria1:=strSemestreSelecionado
        .Range("A2:DW" & ult_linha_Plano).AutoFilter Field:=Coluna_OrigemModelo, Criteria1:="<>" & "Excluido"
    End With
    
    
' Copiar dados filtrados e colar no consolidado
    ult_linha_Plano_filtro = db.Cells(db.Rows.Count, "B").End(xlUp).Row
    
    ' Validacao para saber se o Grupo possui itens vinculados
    If ult_linha_Plano_filtro <= 2 Then
        
        ' Fechar o arquivo de Consolidado
        If Not Compilado Is Nothing Then
            Compilado.Close SaveChanges:=False
            Set Compilado = Nothing
        End If
        
        ' Resetar filtros na base
        db.Activate
        db.Range("A2").Select
        ActiveSheet.ShowAllData
        
        ' Registro de log de erro por falta de dados
        last_row = Controle_Erro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Erro
            .Range("A" & last_row).Value = "Erro: Grupos selecionados sem códigos vinculados"
            .Range("B" & last_row).Value = hoje
            .Range("C" & last_row).Value = horaAtual
            .Range("D" & last_row).Value = usuario
        End With
        
        ' Retorno ao menu e bloqueio de segurança
        mn.Select
        Call bBloqueio
        
        MsgBox "Atenção! Nenhum produto dentro dos Grupos selecionados possui identificador válido.", vbExclamation
        End ' Interrupção estruturada
        
    End If
    
    ' Transferir os dados para a planilha de consolidado
    db.Range("A3:DB" & ult_linha_Plano_filtro).Copy
    
    With wsCompilado
        .Cells(2, 1).PasteSpecial xlPasteValues
        .Range("DC2").Value = caminho_pasta
    End With
    
    ' Extensão de fórmulas de suporte
    ult_linha_Compilado = wsCompilado.Cells(Rows.Count, "B").End(xlUp).Row
    wsCompilado.Range("DC2:DW2").Select
    Selection.AutoFill Destination:=wsCompilado.Range("DC2:DW" & ult_linha_Compilado), Type:=xlFillDefault
    
    ' Normalização de colunas (Conversão de Texto para Número)
    colunasParaConverter = Array("G", "H", "AE", "AF")

    For Each col In colunasParaConverter
        wsCompilado.Columns(col).TextToColumns Destination:=wsCompilado.Range(col & "1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
            TrailingMinusNumbers:=True
    Next col
    
' < PARTE 2 ------------- Atualizar as bases de Datasets Externos

    ' Definição de diretórios de servidores e bases de dados
    CaminhoArquivo = "S:\PUBLICO\DADOS_PERFORMANCE\INDICADORES\"
    nomeArquivo = "DB_ANALITICO_MOVIMENTACOES_GERAL.xlsm"
    
    ' Identificadores das abas de origem
    abasParaCopiar = Array("Data_Analitico", "Data_Vendas", "Data_Recebimentos", "Data_Carteira", "Data_Vendas_Espec", "Data_Recebimentos_Espec")

    ' Acesso às bases de dados em modo leitura
    Set wbOrigem = Workbooks.Open(CaminhoArquivo & nomeArquivo, ReadOnly:=True)
    
    anoAtual = Year(Date)
    semestreAtual = IIf(Month(Date) <= 6, 1, 2)

    ' Processamento cíclico de atualização das bases
    For Each aba In abasParaCopiar
    
        Set wsOrigem = Nothing
        Set wsDestino = Nothing
        Set rngDadosCopia = Nothing
        Set rngDadosLimpeza = Nothing
        linhasParaExcluir = 0

        ' Mapeamento de abas de sistema para abas de relatório
        Set wsOrigem = wbOrigem.Sheets(aba)
        If aba = "Data_Vendas_Espec" Then
            Set wsDestino = Compilado.Sheets("DB_Historico_Vendas")
        ElseIf aba = "Data_Recebimentos_Espec" Then
            Set wsDestino = Compilado.Sheets("DB_Recebimentos")
        Else
            Set wsDestino = Compilado.Sheets(aba)
        End If
        
        ' Verificação de consistência de dados na origem
        If Application.WorksheetFunction.CountA(wsOrigem.Cells) > 0 Then
            ultLinhaOrigem = wsOrigem.Cells.SpecialCells(xlCellTypeLastCell).Row

            ' Definição de escopo de dados por tipo de dataset
            Select Case aba
                Case "Data_Recebimentos"
                    Set rngDadosCopia = wsOrigem.Range("A1", wsOrigem.Cells(ultLinhaOrigem, "D"))
                    Set rngDadosLimpeza = wsDestino.Range("A1", wsDestino.Cells(wsDestino.Rows.Count, "D"))
                    rngDadosLimpeza.ClearContents
                    linhasParaExcluir = 5

                Case "Data_Vendas"
                    Set rngDadosCopia = wsOrigem.Range("A1", wsOrigem.Cells(ultLinhaOrigem, "F"))
                    Set rngDadosLimpeza = wsDestino.Range("A1", wsDestino.Cells(wsDestino.Rows.Count, "F"))
                    rngDadosLimpeza.ClearContents
                    linhasParaExcluir = 5
                                        
                Case "Data_Recebimentos_Espec", "Data_Vendas_Espec"
                    ' Validação de período para bases específicas
                    If Int(strAnoSelecionado) = anoAtual And Int(strSemestreSelecionado) = semestreAtual Then
                        Dim colFim As String: colFim = IIf(InStr(aba, "Vendas"), "F", "D")
                        Set rngDadosCopia = wsOrigem.Range("A1", wsOrigem.Cells(ultLinhaOrigem, colFim))
                        Set rngDadosLimpeza = wsDestino.Range("A1", wsDestino.Cells(wsDestino.Rows.Count, colFim))
                        rngDadosLimpeza.ClearContents
                        linhasParaExcluir = 5
                    Else
                        GoTo ProximaAba
                    End If
                                        
                Case "Data_Carteira"
                    Set rngDadosCopia = wsOrigem.Range("A1", wsOrigem.Cells(ultLinhaOrigem, "AH"))
                    Set rngDadosLimpeza = wsDestino.Range("A1", wsDestino.Cells(wsDestino.Rows.Count, "AH"))
                    rngDadosLimpeza.ClearContents
                    linhasParaExcluir = 0

                Case "Data_Analitico"
                    Set rngDadosCopia = wsOrigem.Range("A1", wsOrigem.Cells(ultLinhaOrigem, "E"))
                    Set rngDadosLimpeza = wsDestino.Range("A1", wsDestino.Cells(wsDestino.Rows.Count, "E"))
                    rngDadosLimpeza.ClearContents
                    linhasParaExcluir = 5
            End Select
            
            ' Execução da transferência de valores
            rngDadosCopia.Copy
            wsDestino.Range("A1").PasteSpecial xlPasteValues
            Application.CutCopyMode = False

            ' Tratamento de cabeçalhos e formatação
            If linhasParaExcluir > 0 Then
                ultLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
                If ultLinhaDestino >= linhasParaExcluir Then
                    wsDestino.Rows("1:" & linhasParaExcluir).Delete Shift:=xlUp
                ElseIf ultLinhaDestino > 0 Then
                    wsDestino.Range("A1", wsDestino.Cells(ultLinhaDestino, 1)).EntireRow.Delete Shift:=xlUp
                End If
                wsDestino.Columns.AutoFit
            End If

            ' Injeção de cálculos dinâmicos pós-processamento
            Select Case aba
                Case "Data_Recebimentos", "Data_Recebimentos_Espec"
                    ultLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
                    wsDestino.Range("E1").Value = "Vol_Recebido_Calc"
                    If ultLinhaDestino >= 2 Then
                        wsDestino.Range("E2").Formula = "=D2*IFERROR(VLOOKUP(VALUE(A2),Dados_Consolidados!$J:$DO,110,FALSE),1)"
                        wsDestino.Range("E2").AutoFill Destination:=wsDestino.Range("E2:E" & ultLinhaDestino), Type:=xlFillDefault
                    End If
                    wsDestino.Columns("E").AutoFit

                Case "Data_Carteira"
                    ultLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
                    If ultLinhaDestino >= 2 Then
                        wsDestino.Range("AI1").Value = "Link_ID_Carteira"
                        wsDestino.Range("AI2").Formula = "=K2&A2"
                        wsDestino.Range("AI2").AutoFill Destination:=wsDestino.Range("AI2:AI" & ultLinhaDestino), Type:=xlFillDefault
                        
                        wsDestino.Range("AJ1").Value = "Fator_Unidade"
                        wsDestino.Range("AJ2").Formula = "=IFERROR(VLOOKUP(AI2,Dados_Consolidados!$DR:$DS,2,FALSE),0)"
                        wsDestino.Range("AJ2").AutoFill Destination:=wsDestino.Range("AJ2:AJ" & ultLinhaDestino), Type:=xlFillDefault

                        wsDestino.Range("AK1").Value = "Vol_Previsto_Final"
                        wsDestino.Range("AK2").Formula = "=IF(L2=""TIPO_ESPECIFICO_AGRUPADO"",M2*AJ2,M2)"
                        wsDestino.Range("AK2").AutoFill Destination:=wsDestino.Range("AK2:AK" & ultLinhaDestino), Type:=xlFillDefault
                    End If
                    wsDestino.Columns("AI:AK").AutoFit
            End Select
        Else
            MsgBox "Aba '" & aba & "' na base de origem não contém registros.", vbInformation
        End If
ProximaAba:
    Next aba

    ' Encerramento do acesso às bases externas
    If Not wbOrigem Is Nothing Then
        wbOrigem.Close SaveChanges:=False
        Set wbOrigem = Nothing
    End If
    
    ' Normalização final dos identificadores principais
    abasParaConverter = Array("DB_Historico_Vendas", "Data_Analitico", "DB_Recebimentos")

    For Each aba In abasParaConverter
        Set wsDestino = Compilado.Sheets(aba)
        wsDestino.Columns("A:A").TextToColumns Destination:=wsDestino.Range("A1"), DataType:=xlDelimited, _
            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Next aba

' < PARTE 3 ------------- Consolidação e Agrupamento para Visão Final

    Set wsCompilado = Compilado.Sheets("Dados_Consolidados")
    Set wsVisaoFinal = Compilado.Sheets("Relatorio_Final")

    ' Formatação de valores monetários para processamento
    wsCompilado.Columns("AE:AE").TextToColumns Destination:=wsCompilado.Range("AE1"), DataType:=xlDelimited, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
    ' Captura de matriz de dados em memória para performance
    ultLinha = wsCompilado.Cells(wsCompilado.Rows.Count, 2).End(xlUp).Row
    ultColuna = wsCompilado.Cells(1, wsCompilado.Columns.Count).End(xlToLeft).Column
    dados = wsCompilado.Range(wsCompilado.Cells(1, 1), wsCompilado.Cells(ultLinha, ultColuna)).Value
    header = wsCompilado.Range(wsCompilado.Cells(1, 1), wsCompilado.Cells(1, ultColuna)).Value

    ' Definição de Dimensões para Agrupamento (Group By)
    colAgrupamento = Array("Fornecedor", "Marca", "Setor", "Grupo", "Classe", _
                            "Subclasse", "Código", "Cluster", _
                            "Estação", "Preço_Unitario")
    
    ' Definição de Métricas para Agregação (Soma)
    colSomar = Array("Qtd Comprada", "Qtd de Venda", "Qtd Recebimento", "Qtd Carteira", _
                     "Qtd Estoque Lojas", "Qtd Estoque CD", _
                     "Compra Liquida", "Venda Valor", "Recebimento Valor", "Carteira Valor", _
                     "Estoque Valor Lojas", "Estoque Valor CD", "Venda Bruta")

    ' Mapeamento de índices de cabeçalho
    Set idxCol = CreateObject("Scripting.Dictionary")
    For j = 1 To ultColuna
        idxCol(header(1, j)) = j
    Next j

    ' Objeto de agregação principal
    Set dictAgrupado = CreateObject("Scripting.Dictionary")

    ' Lógica de Agrupamento via Dicionários aninhados
    For i = 2 To UBound(dados)
        chaveAgrup = ""
        For Each Chave In colAgrupamento
            chaveAgrup = chaveAgrup & "|" & dados(i, idxCol(Chave))
        Next

        If Not dictAgrupado.Exists(chaveAgrup) Then
            Set dictLinha = CreateObject("Scripting.Dictionary")
            For Each Chave In colAgrupamento
                dictLinha(Chave) = dados(i, idxCol(Chave))
            Next
            For Each Chave In colSomar
                dictLinha(Chave) = 0
            Next
            dictLinha("Qtd Estoque Total") = 0
            dictLinha("Valor Estoque Total") = 0
            dictLinha("LB%_soma") = 0
            dictLinha("LB%_cont") = 0
            
            ' Atributos qualitativos
            If idxCol.Exists("Dias de Venda") Then dictLinha("Dias de Venda") = dados(i, idxCol("Dias de Venda")) Else dictLinha("Dias de Venda") = ""
            If idxCol.Exists("Status_Giro") Then dictLinha("Status_Giro") = dados(i, idxCol("Status_Giro")) Else dictLinha("Status_Giro") = ""
            
            dictLinha("Qtd_A_Comprar") = 0
            dictLinha("Obs_Interna") = ""
            dictLinha("Acao_Sugerida") = ""
            dictLinha("Status_Operacional") = ""
            dictAgrupado.Add chaveAgrup, dictLinha
        End If

        Set dictLinha = dictAgrupado(chaveAgrup)
        For Each Chave In colSomar
            If idxCol.Exists(Chave) Then
                If IsNumeric(dados(i, idxCol(Chave))) Then
                    dictLinha(Chave) = dictLinha(Chave) + CDbl(dados(i, idxCol(Chave)))
                End If
            End If
        Next
        
        ' Consolidação de estoques
        dictLinha("Qtd Estoque Total") = dictLinha("Qtd Estoque Lojas") + dictLinha("Qtd Estoque CD")
        dictLinha("Valor Estoque Total") = dictLinha("Estoque Valor Lojas") + dictLinha("Estoque Valor CD")

        ' Acumuladores para cálculo de margem média (LB%)
        If idxCol.Exists("LB%") Then
            If IsNumeric(dados(i, idxCol("LB%"))) Then
                dictLinha("LB%_soma") = dictLinha("LB%_soma") + CDbl(dados(i, idxCol("LB%")))
                dictLinha("LB%_cont") = dictLinha("LB%_cont") + 1
            End If
        End If
    Next i
    
    ' Preparação da aba de saída de resultados
    wsVisaoFinal.Select
    wsVisaoFinal.Cells.Clear
    
    ' Remoção de elementos gráficos residuais
    wsVisaoFinal.Cells.Select
    On Error Resume Next
    ActiveSheet.DrawingObjects.Delete
    On Error GoTo 0
    
' Define a linha de início para os cabeçalhos do relatório final
    linhaDestino = 2
    j = 1

    ' Estruturação dos Cabeçalhos da Visão de Saída
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Midia": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Fornecedor": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Marca": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Setor": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Grupo": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Classe": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Subclasse": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Identificador": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Ref_Interna": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Cluster_Agrup": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Sazonalidade": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Preco_Unitario": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Qtd_Total_Comprada": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Qtd_Venda_Acum": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Qtd_Recebimento": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Qtd_Em_Transito": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Estoque_Total": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Estoque_Lojas": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Estoque_CD": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Compra_Liq": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Venda_Ref": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Receb_Ref": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Transito_Ref": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Estoque_Total": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Estoque_Lojas": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Valor_Estoque_CD": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "%_Giro": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Margem_Bruta": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Dias_Ativos": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Cobertura_Dias": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Range_Giro": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Venda_Financ": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Sugerido_Compra": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Notas": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Decisao": j = j + 1
    wsVisaoFinal.Cells(linhaDestino, j).Value = "Status_Fluxo"

    ' Preenchimento dos dados processados (Dicionário -> Planilha)
    linhaDestino = 3
    For Each chaveAgrup In dictAgrupado.Keys
        Set dictLinha = dictAgrupado(chaveAgrup)
        j = 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = "": j = j + 1 ' Espaço reservado para mídia/imagem
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Fornecedor"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Marca"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Setor"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Grupo"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Classe"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Subclasse"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Código"): j = j + 1
        
        ' Busca de Referência cruzada na base consolidada
        codigoBusca = dictLinha("Código")
        valorReferencia = ""
        With wsCompilado
            For linhaCompilado = 2 To .Cells(.Rows.Count, idxCol("Código")).End(xlUp).Row
                If .Cells(linhaCompilado, idxCol("Código")).Value = codigoBusca Then
                    If idxCol.Exists("Referencia") Then
                        valorReferencia = .Cells(linhaCompilado, idxCol("Referencia")).Value
                    End If
                    Exit For
                End If
            Next linhaCompilado
        End With
        
        wsVisaoFinal.Cells(linhaDestino, j).Value = valorReferencia: j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Cluster Compra"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Estação do Produto"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Preço de Vendas"): j = j + 1

        ' Inserção de métricas quantitativas
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Comprada"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd de Venda"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Recebimento"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Carteira"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Estoque Total"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Estoque Contábil - Lojas"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd Estoque Contábil - CD"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Compra Líquida Total"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Venda em CML"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Recebimento CML"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Carteira CML"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("CML Estoque Total"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Estoque Contábil em CML - Lojas"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Estoque Contábil em CML - CD"): j = j + 1

        ' Cálculo de Giro (% Sell-Through)
        If dictLinha("Qtd de Venda") + dictLinha("Qtd Estoque Total") <> 0 Then
            wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Qtd de Venda") / (dictLinha("Qtd de Venda") + dictLinha("Qtd Estoque Total"))
        Else
            wsVisaoFinal.Cells(linhaDestino, j).Value = 0
        End If
        j = j + 1

        ' Recuperação de Margem (LB%) via busca indexada
        codigoBusca = dictLinha("Código")
        valorLB = 0
        With wsCompilado
            For linhaCompilado = 2 To .Cells(.Rows.Count, idxCol("Código")).End(xlUp).Row
                If .Cells(linhaCompilado, idxCol("Código")).Value = codigoBusca Then
                    If idxCol.Exists("LB%") Then
                        If IsNumeric(.Cells(linhaCompilado, idxCol("LB%")).Value) Then
                            valorLB = CDbl(.Cells(linhaCompilado, idxCol("LB%")).Value)
                        End If
                    End If
                    Exit For
                End If
            Next linhaCompilado
        End With
        
        wsVisaoFinal.Cells(linhaDestino, j).Value = valorLB: j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Dias de Venda"): j = j + 1
        
        ' Cálculo de Cobertura de Estoque
        If dictLinha("Dias de Venda") <> 0 Then
            mediaVendaDia = IIf(dictLinha("Qtd de Venda") <> 0, dictLinha("Qtd de Venda") / dictLinha("Dias de Venda"), 0)
            wsVisaoFinal.Cells(linhaDestino, j).Value = IIf(mediaVendaDia <> 0, dictLinha("Qtd Estoque Total") / mediaVendaDia, 0)
        Else
            wsVisaoFinal.Cells(linhaDestino, j).Value = 0
        End If
        j = j + 1

        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Faixa de Dias de Venda"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Venda Financeira"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Quantidade a Comprar"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Observação"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Ação"): j = j + 1
        wsVisaoFinal.Cells(linhaDestino, j).Value = dictLinha("Status")

        linhaDestino = linhaDestino + 1
    Next
     
    ' --- SINCRONIZAÇÃO COM DADOS HISTÓRICOS DE DESEMPENHO ---
    
    ' Mapeamento dinâmico de colunas da base de Retorno (rd)
    cabecalhoB = rd.Rows(1).Value
    For i = LBound(cabecalhoB, 2) To UBound(cabecalhoB, 2)
        Select Case cabecalhoB(1, i)
            Case "Fornecedor": colB_Fornecedor = i
            Case "Marca": colB_Marca = i
            Case "Setor": colB_Setor = i
            Case "Grupo": colB_Grupo = i
            Case "Classe": colB_Classe = i
            Case "Subclasse": colB_Subclasse = i
            Case "Código": colB_Codigo = i
            Case "Cluster Compra": colB_ClusterCompra = i
            Case "Estação do Produto": colB_EstacaoProduto = i
            Case "Preço de Vendas": colB_PrecoVendas = i
            Case "Quantidade a Comprar": colB_QuantidadeComprar = i
            Case "Observação": colB_Observacao = i
            Case "Ação": colB_Acao = i
            Case "Status": colB_Status = i
        End Select
    Next i
    
    ' Mapeamento dinâmico de colunas da Visão Final (Destino)
    cabecalhoA = wsVisaoFinal.Rows(2).Value
    For i = LBound(cabecalhoA, 2) To UBound(cabecalhoA, 2)
        Select Case cabecalhoA(1, i)
            Case "Fornecedor": colA_Fornecedor = i
            Case "Marca": colA_Marca = i
            Case "Status": colA_Status = i
        End Select
    Next i
    
    ' Carregamento de chaves históricas em Dicionário para cruzamento veloz
    Set dictB = CreateObject("Scripting.Dictionary")
    lastRowB = rd.Cells(Rows.Count, "C").End(xlUp).Row
    
    For j = 2 To lastRowB
        keyB = rd.Cells(j, colB_Fornecedor).Value & "|" & rd.Cells(j, colB_Marca).Value & "|" & _
               rd.Cells(j, colB_Setor).Value & "|" & rd.Cells(j, colB_Grupo).Value & "|" & _
               rd.Cells(j, colB_Classe).Value & "|" & rd.Cells(j, colB_Subclasse).Value & "|" & _
               rd.Cells(j, colB_Codigo).Value & "|" & rd.Cells(j, colB_ClusterCompra).Value & "|" & _
               rd.Cells(j, colB_EstacaoProduto).Value & "|" & rd.Cells(j, colB_PrecoVendas).Value
        
        If Not dictB.Exists(keyB) Then dictB.Add keyB, j
    Next j
    
    ' Devolução de informações editadas pelo usuário (Qtd, Obs, Ação, Status)
    lastRowA = wsVisaoFinal.Cells(Rows.Count, "D").End(xlUp).Row
    For k = 3 To lastRowA
        keyA = wsVisaoFinal.Cells(k, colA_Fornecedor).Value & "|" & wsVisaoFinal.Cells(k, colA_Marca).Value
        
        If dictB.Exists(keyA) Then
            repetRow = dictB.Item(keyA)
            wsVisaoFinal.Cells(k, colA_QuantidadeComprar).Value = rd.Cells(repetRow, colB_QuantidadeComprar).Value
            wsVisaoFinal.Cells(k, colA_Observacao).Value = rd.Cells(repetRow, colB_Observacao).Value
            wsVisaoFinal.Cells(k, colA_Acao).Value = rd.Cells(repetRow, colB_Acao).Value
            wsVisaoFinal.Cells(k, colA_Status).Value = rd.Cells(repetRow, colB_Status).Value
        End If
    Next k
    
    ' --- APLICAÇÃO DE REGRAS DE FILTRO E LIMPEZA ---
    With wsVisaoFinal
        If .AutoFilterMode Then .ShowAllData
        .Range("A2").AutoFilter
        
        ' Aplica critérios de seleção de Ação e Status definidos no formulário
        .Range("A2:AJ" & lastRowA).AutoFilter Field:=colA_Acao, Criteria1:=Acoes_Selecionadas_Array, Operator:=xlFilterValues
        .Range("A2:AJ" & lastRowA).AutoFilter Field:=colA_Status, Criteria1:=Status_Selecionados_Array, Operator:=xlFilterValues
        
        Dim rngVisible As Range
        On Error Resume Next
        Set rngVisible = .Range("A3:AJ" & lastRowA).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not rngVisible Is Nothing Then
            ' Limpeza física de linhas ocultas para manter apenas os dados filtrados
            Dim rowstodelete As Range
            For r = lastRowA To 3 Step -1
                If .Rows(r).Hidden Then
                    If rowstodelete Is Nothing Then Set rowstodelete = .Rows(r) Else Set rowstodelete = Union(rowstodelete, .Rows(r))
                End If
            Next r
            
            If Not rowstodelete Is Nothing Then rowstodelete.Delete Shift:=xlUp
            .ShowAllData
            lastRowA = .Cells(Rows.Count, "D").End(xlUp).Row
        Else
            ' Tratamento de exceção: Filtro resultou em zero registros
            .ShowAllData
            If Not Compilado Is Nothing Then Compilado.Close SaveChanges:=False
            
            db.Activate
            ActiveSheet.ShowAllData
            
            ' Registro de Log de erro de saída vazia
            last_row = Controle_Erro.Cells(Rows.Count, "B").End(xlUp).Row + 1
            With Controle_Erro
                .Range("A" & last_row).Value = "Relatório sem dados após filtros de Ação/Status."
                .Range("B" & last_row).Value = hoje
                ' ...
            End With
            
            mn.Select
            Call bBloqueio
            MsgBox "Critérios de Ação/Status não retornaram dados.", vbInformation
            End
        End If
    End With

' --- CÁLCULOS DE RESUMO (SUBTOTAIS E PONDERADOS) ---
    With wsVisaoFinal
        ' Inserção da linha de Totalização no topo
        .Rows("2:2").Insert Shift:=xlDown
        .Range("A2").Value = "TOTAL GERAL"
        
        ' Fórmulas de Subtotal (Soma) para colunas de Volume e Valor (M a Z)
        .Range("M2").Formula = "=SUBTOTAL(9,M4:M" & lastRowA + 1 & ")"
        .Range("M2").AutoFill Destination:=.Range("M2:Z2"), Type:=xlFillDefault
        
        ' Cálculo de Giro (% Sell-Through) consolidado
        .Range("AA2").Formula = "=N2/(N2+Q2)"
        
        ' Cálculo de Margem Bruta (LB%) Ponderada (usa coluna auxiliar AK para ignorar linhas ocultas)
        .Range(.Cells(4, "AK"), .Cells(lastRowA + 1, "AK")).Formula = "=SUBTOTAL(103,D4)"
        .Cells(2, "AB").FormulaLocal = "=SOMARPRODUTO(AB4:AB" & lastRowA + 1 & ";AF4:AF" & lastRowA + 1 & ";AK4:AK" & lastRowA + 1 & ")/SOMARPRODUTO(AF4:AF" & lastRowA + 1 & ";AK4:AK" & lastRowA + 1 & ")"
        
        ' Outras métricas de resumo
        .Range("AC2") = "-" ' Dias de Venda (não somável)
        .Range("AD2").Formula = "=SUBTOTAL(1,AD4:AD" & lastRowA + 1 & ")" ' Média de Cobertura
        .Range("AE2") = "-" ' Faixa (texto)
        .Range("AF2").Formula = "=SUBTOTAL(9,AF4:AF" & lastRowA + 1 & ")" ' Venda Financeira Total
        .Range("AG2").Formula = "=SUBTOTAL(9,AG4:AG" & lastRowA + 1 & ")" ' Qtd Sugerida Total
    End With

    ' --- ESTILIZAÇÃO E IDENTIDADE VISUAL ---
    With wsVisaoFinal
        .Columns.AutoFit

        ' Definição dos Agrupadores Superiores (Categorias)
        .Range("B1").Value = "DIMENSÕES HIERÁRQUICAS"
        .Range("H1").Value = "DADOS DO ITEM"
        .Range("M1").Value = "VOLUMETRIA"
        .Range("T1").Value = "VALORES (CML)"
        .Range("AA1").Value = "KPIS DESEMPENHO"
        .Range("AG1").Value = "ESTRATÉGIA / PLANO"

        ' Configuração de Cores e Alinhamentos por Blocos
        arrRangesL1 = Array("A1", "B1:G1", "H1:L1", "M1:S1", "T1:Z1", "AA1:AF1", "AG1:AJ1")
        arrRangesL2_3 = Array("A2:A3", "B2:G3", "H2:L3", "M2:S3", "T2:Z3", "AA2:AF3", "AG2:AJ3")
        
        ' Paleta de Cores (Azul, Verde, Laranja)
        arrColorsL1 = Array(RGB(172, 185, 202), RGB(169, 208, 142), RGB(172, 185, 202), RGB(169, 208, 142), RGB(172, 185, 202), RGB(169, 208, 142), RGB(255, 217, 102))
        arrColorsL2_3 = Array(RGB(214, 220, 228), RGB(198, 224, 184), RGB(214, 220, 228), RGB(198, 224, 184), RGB(214, 220, 228), RGB(198, 224, 184), RGB(255, 230, 153))

        For colIdx = LBound(arrRangesL1) To UBound(arrRangesL1)
            With .Range(arrRangesL1(colIdx))
                .HorizontalAlignment = xlCenterAcrossSelection
                .Interior.Color = arrColorsL1(colIdx)
                .Font.Bold = True
            End With
            .Range(arrRangesL2_3(colIdx)).Interior.Color = arrColorsL2_3(colIdx)
        Next colIdx

        ' Ajustes de Altura e Tipografia
        .Rows(1).RowHeight = 25: .Rows(1).Font.Size = 14
        .Rows(3).RowHeight = 35: .Rows(3).WrapText = True
        
        ' Formatação de Números e Moeda
        .Range("M:AG").NumberFormat = "#,##0"
        .Range("AA:AB").NumberFormat = "0.00%"
        .Columns("L").NumberFormat = "0.00"
        
        ' Proteção e Congelamento
        .Columns("AK").EntireColumn.Hidden = True
        .Range("B4").Select: ActiveWindow.FreezePanes = True

        ' Listas de Validação (Drop-downs) para o Plano de Ação
        ultLinhaDestino = .Cells(.Rows.Count, "H").End(xlUp).Row
        With .Range("AI4:AI" & ultLinhaDestino).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Comprar,Cancelar,Prorrogar,Antecipar,Remarcar,Promover,Distribuir,Nenhuma"
        End With
    End With

    ' --- IMPORTAÇÃO DINÂMICA DE FOTOS ---
    diretorio = wsCompilado.Range("DC2") & "Fotos-Produtos\"
    last_row = wsVisaoFinal.Cells(Rows.Count, "H").End(xlUp).Row
    
    For i = 4 To last_row
        nomeFoto = wsVisaoFinal.Cells(i, Coluna_Referencia)
        Set cell = wsVisaoFinal.Range("A" & i)
        
        ' Verifica existência do arquivo (.jpg ou .png)
        arquivoImagem = ""
        If Dir(diretorio & nomeFoto & ".jpg") <> "" Then arquivoImagem = diretorio & nomeFoto & ".jpg"
        If arquivoImagem = "" And Dir(diretorio & nomeFoto & ".png") <> "" Then arquivoImagem = diretorio & nomeFoto & ".png"

        If arquivoImagem <> "" Then
            Set img = wsVisaoFinal.Pictures.Insert(arquivoImagem)
            With img
                ' Redimensionamento Proporcional à célula
                ratio = .Width / .Height
                If ratio > (cell.Width / cell.RowHeight) Then
                    .Width = cell.Width: .Height = cell.Width / ratio
                Else
                    .Height = cell.RowHeight: .Width = cell.RowHeight * ratio
                End If
                ' Centralização
                .Top = cell.Top + (cell.RowHeight - .Height) / 2
                .Left = cell.Left + (cell.Width - .Width) / 2
            End With
        Else
            cell.Value = "Sem Imagem"
        End If
    Next i

    ' --- EXPORTAÇÃO E FINALIZAÇÃO ---
    wsVisaoFinal.Copy ' Gera novo workbook
    Set wbNovo = ActiveWorkbook
    wbNovo.Sheets(1).Name = "Relatorio_Consolidado"
    
    ' Nomeação dinâmica baseada no setor
    setorNome = db.Cells(3, Coluna_Setor).Value
    nomeArquivo = "Relatorio_" & setorNome & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    wbNovo.SaveAs Filename:=ThisWorkbook.Path & "\" & nomeArquivo
    wbNovo.Close False
    
    ' Log de Auditoria e Feedback
    With Controle_Macro
        last_row_macro = .Cells(Rows.Count, "B").End(xlUp).Row + 1
        .Range("A" & last_row_macro).Resize(1, 5).Value = Array("Gerar Relatorio", Date, Time, usuario, "Finalizada")
    End With

    MsgBox "Relatório gerado com sucesso!", vbInformation
    ' Restaura configurações do Excel
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
    
End Sub



