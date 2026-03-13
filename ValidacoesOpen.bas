Attribute VB_Name = "ValidacoesOpen"
Sub ValidarProtocolosEntrada()

    ' --- DECLARAÇÕES ---
    Dim db As Worksheet, mn As Worksheet, ap As Worksheet, pd As Worksheet, Controle_Erro As Worksheet
    Dim last_row_erro As Long, ult_linha_protocolo As Long
    Dim i As Long
    Dim usuario As String, setor_operacional As String
    Dim colDict As Object, Celula As Range
    Dim nome As Variant, nomeProcurado As String
    Dim setoresPermitidos As Variant

    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_OPERACIONAIS")
    Set mn = ThisWorkbook.Sheets("Painel_Controle")
    Set ap = ThisWorkbook.Sheets("Parametros")
    Set pd = ThisWorkbook.Sheets("Log_Transacoes")
    Set Controle_Erro = ThisWorkbook.Sheets("Log_Erros")
    Set colDict = CreateObject("Scripting.Dictionary")

    ' Auditoria de Execução
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    
    ' --- MAPEAMENTO DE COLUNAS CRÍTICAS ---
    ' Definido com base no tipo de registro de entrada
    If db.Range("I3").Value = "TIPO_B" Then
        nomesColunas = Array("ID", "DATA_REFERENCIA", "ENTIDADE", "DESCRITIVO_ITEM", "COD_REFERENCIA", "ESCANEAMENTO", "SUBCONJUNTO", _
                             "COD_SUBCONJUNTO", "CATEGORIA_TECNICA", "AGRUPAMENTO", "COD_AGRUPAMENTO", "VOL_POR_LOTE", "CUSTO_UNITARIO", _
                             "VALOR_REPASSE", "MATRIZ_COR", "CICLO_VIDA", "METODO_PAGAMENTO", "COD_FISCAL", _
                             "FABRICANTE_ORIGEM", "CLASSIFICACAO_FISCAL", "GRUPO_GESTOR", "PARCEIRO_NEGOCIO", "NIVEL_LOGISTICO", "TIPO_FLUXO", "EMAIL_RESPONSAVEL", _
                             "AGENTE_EXTERNO", "ANO_CONTABIL", "PERIODO_REF", "CLASSE_GERAL", "MONTANTE_LIQUIDO", "SETOR_OPERACIONAL")
    Else
        nomesColunas = Array("ID", "DATA_REFERENCIA", "ENTIDADE", "DESCRITIVO_ITEM", "COD_REFERENCIA", "ESCANEAMENTO", "SUBCONJUNTO", _
                             "COD_SUBCONJUNTO", "CATEGORIA_TECNICA", "AGRUPAMENTO", "COD_AGRUPAMENTO", "VOL_POR_LOTE", "CUSTO_UNITARIO", _
                             "VALOR_REPASSE", "DIMENSAO_FISICA", "MATRIZ_COR", "CICLO_VIDA", "METODO_PAGAMENTO", "COD_FISCAL", _
                             "FABRICANTE_ORIGEM", "CLASSIFICACAO_FISCAL", "GRUPO_GESTOR", "PARCEIRO_NEGOCIO", "NIVEL_LOGISTICO", "TIPO_FLUXO", "EMAIL_RESPONSAVEL", _
                             "AGENTE_EXTERNO", "ANO_CONTABIL", "PERIODO_REF", "CLASSE_GERAL", "MONTANTE_LIQUIDO", "SETOR_OPERACIONAL")
    End If

    ' Indexação dinâmica das colunas na aba de Log
    For Each Celula In pd.Rows(1).Cells
        nomeProcurado = Celula.Value
        If Not IsError(Application.Match(nomeProcurado, nomesColunas, 0)) Then
            colDict(nomeProcurado) = Celula.Column
        End If
    Next Celula

    ' Definição do Escopo Operacional
    If db.Range("K3").Value = "OP_INF_01" Or db.Range("K3").Value = "OP_INF_02" Then
        setor_operacional = "AREA_INFANTIL"
    Else
        setor_operacional = db.Range("K3").Value
    End If

    ult_linha_protocolo = pd.Cells(Rows.Count, "A").End(xlUp).Row
    last_row_erro = Controle_Erro.Cells(Rows.Count, "B").End(xlUp).Row + 1

    ' V1 -- Sintaxe de Nomeclatura de Arquivos de Protocolo
    For i = 2 To ult_linha_protocolo
    
        PrefixoSistema = Mid(pd.Cells(i, "A").Value, 1, 3)    ' Padrão esperado "202"
        Delimitador_A = Mid(pd.Cells(i, "A").Value, 9, 1)    ' Esperado "_"
        Delimitador_B = Mid(pd.Cells(i, "A").Value, 16, 1)   ' Esperado "_"
        TagSetor = pd.Cells(i, "AV").Value                   ' Tag de validação de setor
    
        ' Verifica integridade da string do nome do arquivo
        If PrefixoSistema <> "202" Or Delimitador_A <> "_" Or Delimitador_B <> "_" Then

            RegistrarErro Controle_Erro, "Inconsistência de sintaxe no arquivo de protocolo", hoje, horaAtual, usuario

            mn.Select
            Call bBloqueio
            mn.Visible = xlSheetVeryHidden
            
            MsgBox "Falha de Sintaxe no arquivo: " & pd.Cells(i, "A").Value & vbCrLf & "Local: " & pd.Cells(i, "AW").Value, vbExclamation, "Erro de Protocolo"
            
            InterromperProcesso "Nomenclatura de arquivo fora do padrão de sistema."
        End If
    
        ' Domínios permitidos para Setores Operacionais
        setoresPermitidos = Array("OP_MAS", "OP_FEM", "OP_CAL", "AREA_INFANTIL", "OP_INF_01", "OP_INF_02", "OP_CASA", "OP_SUP", "OP_INT", "OP_TEC", "OP_ACC")
    
        If IsError(Application.Match(TagSetor, setoresPermitidos, 0)) Then
            
            RegistrarErro Controle_Erro, "Setor operacional não reconhecido no protocolo", hoje, horaAtual, usuario
                                                
            mn.Select
            Call bBloqueio
            mn.Visible = xlSheetVeryHidden
            
            MsgBox "O setor identificado no arquivo " & pd.Cells(i, "A").Value & " não consta na base de permissões.", vbExclamation, "Erro de Domínio"
            
            InterromperProcesso "Setor inválido ou sem acesso."
        End If
    Next i

    ' V2 -- Verificação de Campos Mandatórios (Células Vitais)
    For i = 2 To ult_linha_protocolo
        For Each nome In nomesColunas
            If colDict.Exists(nome) Then
                
                ' Verifica se campo está vazio para o setor ativo e data de corte superior a Abr/2025
                If Trim(pd.Cells(i, colDict(nome)).Value) = "" And setor_operacional = pd.Cells(i, "AV").Value And CDate(pd.Cells(i, "AX").Value) > "01/04/2025" Then
                    
                    RegistrarErro Controle_Erro, "Campo mandatório vazio em arquivo de log", hoje, horaAtual, usuario

                    mn.Select
                    Call bBloqueio
                    mn.Visible = xlSheetVeryHidden
                    
                    MsgBox "Processamento bloqueado: Arquivo com dados incompletos detectado." & vbCrLf & _
                           "Arquivo: " & pd.Cells(i, "A").Value & vbCrLf & _
                           "Campo Ausente: " & nome, vbExclamation, "Integridade de Dados"
                    
                    InterromperProcesso "Falha de preenchimento em colunas vitais."
                End If
            End If
        Next nome
    Next i
    
    ' Liberação do Sistema
    ap.Range("AO2").Value = 1
    mn.Visible = xlSheetVisible

End Sub

