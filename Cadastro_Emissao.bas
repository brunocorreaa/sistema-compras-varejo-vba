Attribute VB_Name = "Cadastro_Emissao"
Sub Emissao_Cadastro_Sist()
    
    Application.ScreenUpdating = False
        
    ' Declarações
    Dim dataExec As Date
    Dim horaExec As String, idUsuario As String
    Dim wsLogMacro As Worksheet
    Dim resposta As VbMsgBoxResult
    
    ' Atribuições
    Set wsLogMacro = ThisWorkbook.Sheets("LOG_SISTEMA")

    ' Confirmação de Segurança
    resposta = MsgBox("Deseja iniciar o processamento de CADASTRO DE REGISTROS?", _
               vbQuestion + vbYesNo + vbDefaultButton2, "Validação de Operação")
    
    If resposta <> vbYes Then Exit Sub
    
    ' Desbloqueio de estrutura
    Call bDesbloqueio

    ' Metadados da execução
    dataExec = Date
    horaExec = Format(Time, "hh:mm:ss")
    idUsuario = Environ("Username")

    ' Registro de Início
    Reg_Log wsLogMacro, "Processo_Cadastro", dataExec, horaExec, idUsuario, "Iniciada"
    
    ' Chamada do Processamento Principal
    Call Processar_Exportacao_Cadastro("ID_CHAMADA_CADASTRO")
    
    ' Registro de Finalização
    Reg_Log wsLogMacro, "Processo_Cadastro", dataExec, horaExec, idUsuario, "Finalizada"

    MsgBox "Processamento concluído com sucesso!"
    Application.ScreenUpdating = True
    
End Sub

' Sub auxiliar para registro de log (Melhor prática de código)
Sub Reg_Log(ws As Worksheet, acao As String, dt As Date, hr As String, usr As String, st As String)
    Dim uLinha As Long
    uLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    ws.Range("A" & uLinha).Value = acao
    ws.Range("B" & uLinha).Value = dt
    ws.Range("C" & uLinha).Value = hr
    ws.Range("D" & uLinha).Value = usr
    ws.Range("E" & uLinha).Value = st
End Sub

Sub Processar_Exportacao_Cadastro(ByVal id_origem As String)
        
    ' Declarações de Objetos
    Dim ws_Base As Worksheet, ws_Export As Worksheet, ws_Menu As Worksheet, ws_Auxiliar As Worksheet
    Dim wb_Destino As Workbook
    Dim dict_Unicos As Object
    Dim rng_Col As Range
    
    ' Declarações de Variáveis de Controle
    Dim uLinhaBase As Long, uLinhaAux As Long
    Dim i As Long, j As Long, p As Long, q As Long
    Dim nomeColBusca As String, pathPasta As String
    
    ' Atribuições de Objetos
    Set dict_Unicos = CreateObject("Scripting.Dictionary")
    Set ws_Base = ThisWorkbook.Sheets("DATA_MASTER")
    Set ws_Auxiliar = ThisWorkbook.Sheets("AUX_SISTEMA")
    Set ws_Menu = ThisWorkbook.Sheets("MENU_PRINCIPAL")
        
    ' Validações de Regras de Negócio (Camada de Segurança)
    Call Validar_Integridade_Dados(id_origem)
    Call Validacoes_Estruturais_Open
    Call bDesbloqueio
            
    ' --- MAPEAMENTO DINÂMICO DE COLUNAS (ORIGEM) ---
    Dim colMap As Variant
    colMap = Array("FORNECEDOR", "AGRUPAMENTO", "VALOR_LIQ_TOTAL", "ANO_REF", "CICLO", "MENU_REF", _
                   "DT_ENTRADA", "DT_SAIDA", "STATUS_VALIDACAO", "ID_REF", "DOC_COMPRA", "TAMANHO", _
                   "GRADE", "CLUSTER_ID", "MODALIDADE", "TEMA_REF", ws_Auxiliar.Range("AJ1").Value, _
                   "CATEGORIA", "QTD_VOL", "CONTATO_EMAIL", "REP_NOME", "LOJAS_QTD", "REF_PE", _
                   "TIPO_ORDEM", "DATA_ENTREGA", "PACK_SIZE", "DESC_ECOM", "VAR_ECOM", "EAN_VAR", _
                   "CATEG_EAN", "ORIGEM_MODELO", "SETOR_REF")

    ' Dicionário para armazenar índices de colunas (Substitui o Select Case exaustivo)
    Dim idxCols As Object: Set idxCols = CreateObject("Scripting.Dictionary")
    
    For Each Item In colMap
        For Each Celula In ws_Base.Rows(2).Cells
            If Celula.Value = Item Then
                idxCols(Item) = Celula.Column
                Exit For
            End If
        Next Celula
    Next Item
    
    ' --- TRATAMENTO DE DATAS (QUARTA-FEIRA MAIS PRÓXIMA) ---
    uLinhaBase = ws_Base.Cells(ws_Base.Rows.Count, "B").End(xlUp).Row
    Set rng_Col = ws_Base.Range(ws_Base.Cells(3, idxCols("DATA_ENTREGA")), ws_Base.Cells(uLinhaBase, idxCols("DATA_ENTREGA")))
    
    rng_Col.NumberFormat = "dd.mm.yyyy"
    For Each cel In rng_Col
        If IsDate(cel.Value) Then
            ' Lógica para ajustar para a próxima Quarta-Feira (Weekday 4)
            If Weekday(cel.Value) <> 4 Then
                cel.Value = CDate(cel.Value) + (4 - Weekday(CDate(cel.Value)) + 7) Mod 7
            End If
        End If
    Next cel
    
    ' --- GERAÇÃO DE LISTA DE ENVIO ÚNICA NA ABA AUXILIAR ---
    ws_Auxiliar.Range("AI2:AN" & ws_Auxiliar.Rows.Count).ClearContents
    
    With ws_Base
        .Range("$A$2:$DW$" & uLinhaBase).AutoFilter Field:=3, Criteria1:="<>"
        
        ' Transferência de blocos de dados para processamento de e-mail
        .Range(.Cells(2, idxCols("CONTATO_EMAIL")), .Cells(uLinhaBase, idxCols("CONTATO_EMAIL"))).SpecialCells(xlCellTypeVisible).Copy
        ws_Auxiliar.Range("AI1").PasteSpecial Paste:=xlPasteValues
        
        ' ... (Repete-se a lógica de cópia para as demais colunas AK, AL, AM, AN conforme original)
    End With
    
    ' --- TRATAMENTO DA ABA AUXILIAR PARA EXPORTAÇÃO ---
    ' Converte Temas em categorias operacionais (MTO/MTA)
    With ws_Auxiliar.Range("AL2:AL" & ws_Auxiliar.Cells(ws_Auxiliar.Rows.Count, "AL").End(xlUp).Row)
        .Replace "Classico", "TIPO_B", xlWhole
        .Replace "Fashion", "TIPO_B", xlWhole
        .Replace "Essencial", "TIPO_A", xlWhole
    End With
    
    ' Limpeza de duplicatas para gerar um arquivo por combinação única de critérios
    uLinhaAux = ws_Auxiliar.Cells(ws_Auxiliar.Rows.Count, "AI").End(xlUp).Row
    ws_Auxiliar.Range("AI1:AN" & uLinhaAux).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), header:=xlYes
    ws_Base.AutoFilterMode = False

    ' --- LOOP DE EXPORTAÇÃO PARA TEMPLATE EXTERNO ---
    pathPasta = ThisWorkbook.Path & "\"
    uLinhaAux = ws_Auxiliar.Cells(ws_Auxiliar.Rows.Count, "AI").End(xlUp).Row
    
    For q = 2 To uLinhaAux
        ' Abre template padrão (Anonimizado)
        Set wb_Destino = Workbooks.Open(pathPasta & "template_sistema_cadastro.xlsm")
        Set ws_Export = wb_Destino.Sheets(1)
        
        ' Limpa dados anteriores do template
        ws_Export.Rows("2:" & ws_Export.Rows.Count).ClearContents
        
        ' Mapeamento de colunas do template de destino
        Dim colDestino As Variant
        colDestino = Array("IMAGEM_REF", "ID_INDEX", "DT_ENTREGA_FINAL", "SISTEMA_ID", "OC_PACK", _
                           "OC_SIMPLES", "SKU_PACK", "SKU_ID", "BRAND", "NOME_ITEM", "REF_FABRICA")
        
    Next q
    
    Sub Processar_Itens_Cadastro(ByVal caller As String)

    last_row_db = db.Cells(db.Rows.Count, "B").End(xlUp).Row
    
    ' --- LOOP PRINCIPAL: Varredura da Base de Dados ---
    For j = 3 To last_row_db
        
        ' Identifica se o item possui múltiplos tamanhos (delimitados por ;)
        cont_carac = Len(db.Cells(j, Coluna_Tamanho)) - Len(Replace(db.Cells(j, Coluna_Tamanho), ";", "")) + 1
        
        ' Sub-loop para desmembrar a grade em linhas individuais (Filhos)
        For t = 1 To cont_carac
            
            ' Aplicação de Filtros e Regras de Validação (MTA vs Essencial)
            Filtro_Valido_MTA = (ap.Range("AL" & q).Value = "MTA")
            Filtro_Valido_Essencial = (db.Cells(j, Coluna_Tema).Value = "Essencial")
            
            ' Conversão de Tipo de Pedido para Código Numérico
            Tipo_Pedido_Filtro = IIf(db.Cells(j, Coluna_TipoPedido) = "Pack", 2, 1)
            
            ' --- VALIDAÇÃO DE PERTENCIMENTO AO LOTE ATUAL ---
            If db.Cells(j, 3).Value <> "" And _
               db.Cells(j, Coluna_Email) = ap.Range("AI" & q) And _
               Month(db.Cells(j, Coluna_DataDeEntrega)) = ap.Range("AK" & q) And _
               Filtro_Valido_MTA = Filtro_Valido_Essencial And _
               db.Cells(j, Coluna_Fabricante).Value = ap.Range("AN" & q) Then
                
                ' Tratamento para Cancelamento: Registra status e aplica tachado (Strikethrough)
                If caller = "CancelarLinha" Then
                    db.Cells(j, Coluna_OrigemDoModelo).Value = "Cancelado"
                    db.Rows(j).Font.Strikethrough = True
                End If
                
                ' Define próxima linha livre no template de destino
                k = cp.Sheets(1).Cells(Rows.Count, "B").End(xlUp).Row + 1
                
                ' --- TRANSFERÊNCIA DE DADOS COLUNA A COLUNA ---
                For i = 1 To contcolumns
                    valor_cabecalho = db.Cells(2, i).Value
                    
                    ' Busca correspondência de cabeçalho entre Origem e Destino
                    For Each cell In rng
                        If LCase(cell.Value) = LCase(valor_cabecalho) Then
                            f = cell.Column
                            
                            ' Lógica para desmembramento de Strings (Tamanho/Grade/EAN)
                            If (i = Coluna_Tamanho Or i = Coluna_Grade) And cont_carac > 1 Then
                                cp.Sheets(1).Cells(k, f).Value = Split(db.Cells(j, i).Value, ";")(t - 1)
                            ElseIf i = Coluna_DataDeEntrega Then
                                cp.Sheets(1).Cells(k, f).Value = Format(db.Cells(j, i).Value, "dd.mm.yyyy")
                            Else
                                cp.Sheets(1).Cells(k, f).Value = db.Cells(j, i).Value
                            End If
                        End If
                    Next cell
                Next i
                
                ' Regra de Negócio: Definição de Flag VEX (Setor_A/Setor_B)
                setor_atual = LCase(db.Cells(j, Coluna_Setor).Text)
                If (setor_atual = "Setor_A" And db.Cells(j, Coluna_Agregacao).Value > 0) Or setor_atual = "Setor_B" Then
                    cp.Sheets(1).Cells(k, Coluna_vex_cp) = "SIM"
                Else
                    cp.Sheets(1).Cells(k, Coluna_vex_cp) = "NAO"
                End If
            End If
        Next t
    Next j

    ' Formatação e Salvamento ---
    
    ' 1. Determinação do Grupo de Compradores (Baseado no Setor)
    Select Case LCase(db.Cells(3, Coluna_Setor).Text)
        Case "Setor_A", "Setor_B", "Setor_C": GrupoCompradores = "GRUPO_A"
        Case "Setor_D", "Setor_F", "Setor_G": GrupoCompradores = "GRUPO_B"
        Case Else: GrupoCompradores = "COORDENACAO_GERAL"
    End Select
    
    ' 2. Geração do Arquivo Final
    Set novoLivro = Workbooks.Add
    cp.Sheets(1).UsedRange.Copy novoLivro.Sheets(1).Range("A1")
    
    ' 3. Personalização Visual por Tipo de Ação (Cores de Cabeçalho)
    With novoLivro.Sheets(1).Rows("1:1").Interior
        If caller = "CancelarLinha" Then
            .Color = RGB(255, 199, 206) ' Vermelho Claro
            path_final = "S:\Caminho\Cancelados\"
        ElseIf caller = "EdicaoLinha" Then
            .Color = RGB(255, 235, 156) ' Amarelo
            path_final = "S:\Caminho\Editados\"
        Else
            .Color = RGB(198, 239, 206) ' Verde
            path_final = "S:\Caminho\Novos\"
        End If
    End With

    ' 4. Proteção de Estrutura (Trava colunas de sistema, libera colunas de edição E:H)
    With novoLivro.Sheets(1)
        .Protect Password:="PROTECAO_SISTEMA", UserInterfaceOnly:=True
        .Columns("D:H").Locked = False ' Permite edição apenas no intervalo necessário
    End With

    ' 5. Salvamento com Nomenclatura Padronizada
    nomeFinal = Format(Now, "yyyymmdd_hhmmss") & "_" & setor & "_" & nomeArquivo
    novoLivro.SaveAs path_final & nomeFinal & ".xlsx"
    novoLivro.Close SaveChanges:=True

End Sub

