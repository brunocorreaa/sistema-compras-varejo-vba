Attribute VB_Name = "BloqueioDesbloqueio"
' Declarações Globais Anonimizadas
Dim ws_Base As Worksheet, ws_Aux As Worksheet, ws_Reg As Worksheet, ws_Menu As Worksheet

Sub bDesbloqueio()
    
    ' Declarações locais
    Dim ws As Worksheet
    Dim abasParaProcessar As Variant
    Dim temFiltro As Boolean
    Dim nomeAba As Variant
    
    abasParaProcessar = Array("DATA_MASTER", "AUX_SISTEMA")
    
    ' Atribuições de Abas
    Set ws_Base = ThisWorkbook.Sheets("DATA_MASTER")
    Set ws_Aux = ThisWorkbook.Sheets("AUX_SISTEMA")
    Set ws_Reg = ThisWorkbook.Sheets("REGISTROS")
    Set ws_Menu = ThisWorkbook.Sheets("MENU_PRINCIPAL")
    
    ' 1. Desproteger aba principal (Senha Anonimizada)
    ws_Base.Unprotect "PROT_SISTEMA_2026"
    
    ' 2. Ocultar abas de sistema via rotina interna
    Call Rotina_Ocultar_Abas_Tecnicas
    
    ' 3. Mostrar todas as colunas nas abas operacionais
    ws_Base.Columns.Hidden = False
    ws_Aux.Columns.Hidden = False
    ws_Reg.Columns.Hidden = False
    
    ' 4. Gestão de Filtros Automáticos
    For Each nomeAba In abasParaProcessar
        Set ws = ThisWorkbook.Sheets(nomeAba)
        ws.Unprotect "PROT_SISTEMA_2026"
        
        temFiltro = False
        
        ' Verifica estrutura de filtro existente
        If ws.AutoFilterMode Then
            If Not ws.AutoFilter Is Nothing Then
                ' Ajusta linha de cabeçalho dependendo da aba
                If ws.AutoFilter.Range.Rows(1).Row = IIf(nomeAba = "DATA_MASTER", 2, 1) Then
                    temFiltro = True
                End If
            End If
        End If
        
        ' Aplica ou limpa o filtro conforme a necessidade
        With ws.Rows(IIf(nomeAba = "DATA_MASTER", 2, 1))
            If temFiltro Then
                If ws.AutoFilter.Filters.Count > 0 Then
                    On Error Resume Next
                    ws.ShowAllData
                    On Error GoTo 0
                End If
            Else
                ws.Range("A" & .Row, ws.Cells(.Row, ws.Columns.Count).End(xlToLeft)).AutoFilter
            End If
        End With
    Next nomeAba
    
End Sub

Sub bBloqueio()

    ' Declarações
    Dim ColIndex As Integer
    Dim ColTarget As Long
    Dim ListaBloqueio As Variant
    Dim i As Long, j As Long
    Dim itemBusca As String
    Dim uLinhaAux As Integer, uColBase As Long
    Dim filtroAtivo As Boolean
    Dim f As Filter

    ' Atribuições
    Set ws_Base = ThisWorkbook.Sheets("DATA_MASTER")
    Set ws_Aux = ThisWorkbook.Sheets("AUX_SISTEMA")
    Set ws_Menu = ThisWorkbook.Sheets("MENU_PRINCIPAL")
    Set ws_Reg = ThisWorkbook.Sheets("REGISTROS")
    
    ' Desproteger para configuração
    ws_Base.Unprotect "PROT_SISTEMA_2026"
    ws_Aux.Unprotect "PROT_SISTEMA_2026"
    
    ' Reset de visualização de colunas
    ws_Base.Select
    ws_Base.Columns("A:XFD").Hidden = False ' Abrange todas as colunas possíveis
    
    ' Ocultação dinâmica baseada em parâmetros da aba Auxiliar
    uLinhaAux = ws_Aux.Range("AB1").End(xlDown).Row

    For i = 2 To uLinhaAux
        uColBase = ws_Base.Range("B2").End(xlToRight).Column
        For j = 1 To uColBase
            ' Se o parâmetro na Auxiliar coincidir com o cabeçalho na Base, oculta
            If ws_Aux.Cells(i, 28).Value = ws_Base.Cells(2, j).Value Then
                ws_Base.Columns(j).Hidden = True
            End If
        Next j
    Next i

    ' Ocultar abas de sistema
    Call Rotina_Ocultar_Abas_Tecnicas
    
    ' Garantir integridade do filtro na linha de cabeçalho (Linha 2)
    temFiltro = False
    If ws_Base.AutoFilterMode Then
        If Not ws_Base.AutoFilter Is Nothing Then
            If ws_Base.AutoFilter.Range.Rows(1).Row = 2 Then temFiltro = True
        End If
    End If
    
    If temFiltro Then
        filtroAtivo = False
        For Each f In ws_Base.AutoFilter.Filters
            If f.On Then
                filtroAtivo = True
                Exit For
            End If
        Next f
        If filtroAtivo Then ws_Base.ShowAllData
    Else
        ws_Base.Range("A2", ws_Base.Cells(2, ws_Base.Columns.Count).End(xlToLeft)).AutoFilter
    End If

    ' Reset de bloqueio de células
    ws_Base.Cells.Locked = False
    ws_Base.Cells.FormulaHidden = False
    
    ' 5. Bloqueio de colunas específicas (Critérios Anonimizados)
    ListaBloqueio = Array("ID_REF", "STATUS_EMISSAO", "CATEGORIA", "AGRUPAMENTO", "NIVEL_01", "NIVEL_02", "COD_NIVEL", _
                          "REF_PRODUTO", "CLUSTER_ID", "QTD_TOTAL", "COD_CLUSTER", "ENTIDADE_FAB", _
                          "DESC_VAR", "VAL_UNIT_PLAN", "MARGEM_REF", "PLAN_VOL", "PLAN_CUSTO", "VAL_LIQ_TOTAL", _
                          "TEMA_REF", "MODALIDADE", "MES_REF", "ANO_REF", "CICLO_REF", "LOGISTICA", "TARGET_LOG", _
                          "PACK_SIZE", "COD_PROD", "DOC_COMPRA", "EAN_GTIN", "VARIANTE_TIPO", _
                          "CATEG_SISTEMA", "DESC_WEB", "LUCRO_EST", "CONTATO_ID", "REP_COMERCIAL", _
                          "ORIGEM_INFO", "COD_PRE_ID", "CHAVE_SISTEMA")
    
    For i = LBound(ListaBloqueio) To UBound(ListaBloqueio)
        itemBusca = ListaBloqueio(i)
        For Each Celula In ws_Base.Rows(2).Cells
            If Celula.Value = itemBusca Then
                ColTarget = Celula.Column
                ws_Base.Columns(ColTarget).Locked = True
                Exit For
            End If
        Next Celula
    Next i
            
    ' Bloqueio de proteção de cabeçalho superior
    ws_Base.Rows(1).Locked = True
            
    ' Atualização de Fórmulas Automáticas f(x)
    uColBase = ws_Base.Range("A1").End(xlToRight).Column
    uLinhaBase = ws_Base.Cells(Rows.Count, "B").End(xlUp).Row

    For i = 1 To uColBase
        ' Identifica colunas de cálculo automático via marcador f(x)
        If ws_Base.Cells(1, i).Value = "f(x)" And ws_Base.Cells(2, i).Value <> "ID_REF" Then
            ws_Base.Range(ws_Base.Cells(3, i), ws_Base.Cells(uLinhaBase, i)).Formula = ws_Base.Cells(3, i).Formula
        End If
    Next i
    
    ' Limpeza de resíduos abaixo da última linha válida
    If ws_Base.AutoFilterMode Then ws_Base.AutoFilter.ShowAllData
    uLinhaBase = ws_Base.Cells(ws_Base.Rows.Count, "B").End(xlUp).Row
    ws_Base.Rows(uLinhaBase + 1 & ":" & ws_Base.Rows.Count).Clear
    
    ' Posicionamento de tela
    ws_Base.Activate
    ws_Base.Cells(1, 1).Select
    ActiveWindow.ScrollColumn = 2

    ' Reset da aba Menu
    ws_Menu.Cells.Clear
    ws_Menu.Activate
    ws_Menu.Cells(1, 1).Select
        
    ' Proteção Final com permissões de usuário
    ws_Base.Protect Password:="PROT_SISTEMA_2026", DrawingObjects:=False, Contents:=True, Scenarios:=False, _
        AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True

End Sub

