Attribute VB_Name = "Klassmatt_Padrao"
Sub ProcessarDadosPadrao()

    Application.ScreenUpdating = False
    
    ' --- DECLARAÇÕES ---
    Dim db As Worksheet, mn As Worksheet, ap As Worksheet, Controle_Macro As Worksheet
    Dim i As Long, j As Long, linha_i As Long, colNumber As Long, w As Long
    Dim caminho_pasta As String, nomeProcurado As String, colLetter As String, resultado As String
    Dim nomeProcurado_Destino As String, nomeProcurado_Plano As String, prefixo As String, Tamanho As String
    Dim planilha As Workbook
    Dim Coluna, last_row, tam_cont_carac As Integer
    Dim Celula As Range, cell As Range, rng As Range
    Dim itens As Variant, valor_procv As Variant
    Dim Coluna_Tipo As Long, Coluna_TipoPedido As Long, Coluna_Filtro As Long
    Dim Coluna_Referencia As Long, Coluna_PrecoBruto As Long, Coluna_CodSubclasse As Long
    Dim Coluna_Setor As Long, Coluna_Agregacao As Long, Coluna_Tema As Long
    Dim Coluna_Subclasse As Long, Coluna_CompraLiquidaTotal As Long
    Dim Coluna_QtdComprada As Long, Coluna_Tamanho As Long, Coluna_CategoriaForn As Long
    Dim Coluna_Classe As Long, Coluna_NomeProduto As Long, Coluna_Marca As Long, Coluna_Cor As Long
    Dim Coluna_CodigoEANVariantes As Long
                
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_PRINCIPAIS")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set ap = ThisWorkbook.Sheets("Apoio")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
                        
    ' Confirmação de execução
    resposta = MsgBox("Você realmente deseja executar o processamento para o PADRÃO EXTERNO?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Sistema")
    
    If resposta <> vbYes Then Exit Sub

    ' Auditoria de execução
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Processamento Padrao"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Iniciada"
    End With

    ' Procedimentos auxiliares
    Call Validacoes("processamento")
    Call bDesbloqueio
            
    ' --- MAPEAMENTO DINÂMICO DE COLUNAS ---
    nomesColunas = Array("Tipo", "Tipo do Pedido", "Ir Menu", "Referencia", "Preço Bruto", "Peso Liquido", "Cód. Subclasse", _
                        "Setor", "Agregação", "Tema", "Subclasse", "Compra Líquida Total", "Qtd Comprada", "Tamanho", _
                        "Categoria EAN", "Classe", "Nome Produto", "Marca", "Cor", "Código EAN Variantes")
    
    For Each nome In nomesColunas
        For Each Celula In db.Rows(2).Cells
            If Celula.Value = nome Then
                Select Case nome
                    Case "Tipo": Coluna_Tipo = Celula.Column
                    Case "Tipo do Pedido": Coluna_TipoPedido = Celula.Column
                    Case "Ir Menu": Coluna_Filtro = Celula.Column
                    Case "Referencia": Coluna_Referencia = Celula.Column
                    Case "Preço Bruto": Coluna_PrecoBruto = Celula.Column
                    Case "Cód. Subclasse": Coluna_CodSubclasse = Celula.Column
                    Case "Setor": Coluna_Setor = Celula.Column
                    Case "Agregação": Coluna_Agregacao = Celula.Column
                    Case "Tema": Coluna_Tema = Celula.Column
                    Case "Subclasse": Coluna_Subclasse = Celula.Column
                    Case "Compra Líquida Total": Coluna_CompraLiquidaTotal = Celula.Column
                    Case "Qtd Comprada": Coluna_QtdComprada = Celula.Column
                    Case "Tamanho": Coluna_Tamanho = Celula.Column
                    Case "Categoria EAN": Coluna_CategoriaForn = Celula.Column
                    Case "Classe": Coluna_Classe = Celula.Column
                    Case "Nome Produto": Coluna_NomeProduto = Celula.Column
                    Case "Marca": Coluna_Marca = Celula.Column
                    Case "Cor": Coluna_Cor = Celula.Column
                    Case "Código EAN Variantes": Coluna_CodigoEANVariantes = Celula.Column
                End Select
                Exit For
            End If
        Next Celula
    Next nome
    
    last_row = db.Cells(Rows.Count, Coluna_Filtro).End(xlUp).Row
    contcolumns = Application.CountA(db.Range("A2:XFD2"))
    
    ' --- MANIPULAÇÃO DO ARQUIVO DE DESTINO ---
    caminho_pasta = ThisWorkbook.Path & "\"
    Set planilha = Workbooks.Open(caminho_pasta & "modelo_integracao.xlsx")
    
    ' Limpeza da área de dados do destino
    colNumber = planilha.Sheets(1).Range("B14").End(xlToRight).Column
    colLetter = GetColumnLetter(colNumber)
    Set rng = planilha.Sheets(1).Range("B14:" & colLetter & "14")
    rowLast = planilha.Sheets(1).Range("B14").End(xlDown).Row
    If rowLast < 1048576 Then planilha.Sheets(1).Range("C15:" & colLetter & rowLast).Clear
            
    ' --- LOOP DE TRANSFERÊNCIA DE DADOS ---
    For j = 3 To last_row
        If db.Cells(j, Coluna_Filtro).Value <> "" Then
            k = planilha.Sheets(1).Range("O" & planilha.Sheets(1).Rows.Count).End(xlUp).Row + 1
            
            For i = 1 To contcolumns
                valor_plano = db.Cells(2, i).Value
                
                For Each cell In rng
                    nomeProcurado_Destino = Replace(Trim(LCase(cell.Value)), "**", "")
                    nomeProcurado_Plano = LCase(valor_plano)
                    CodSubclasse = db.Cells(j, Coluna_CodSubclasse)
                    f = cell.Column
                    
                    ' Transferência direta baseada em nome de cabeçalho
                    If nomeProcurado_Destino = nomeProcurado_Plano And valor_plano <> "Nº" Then
                        planilha.Sheets(1).Cells(k, f).Value = db.Cells(j, i).Value
                        
                        ' Regra específica para produtos comerciais sem tamanho
                        If db.Cells(j, Coluna_Tipo).Value Like "*COMERCIALIZÁVEL*" And cell.Value = "TAMANHO" Then
                            planilha.Sheets(1).Cells(k, f).Value = ""
                        End If
                        
                        ' Regra para Packs e Variantes únicas
                        If db.Cells(j, Coluna_TipoPedido).Value = "Pack" And (cell.Value Like "*Variantes*") Then
                            planilha.Sheets(1).Cells(k, f).Value = ""
                        ElseIf LCase(db.Cells(j, Coluna_Tamanho).Text) = "u" And cell.Value = "Código EAN Variantes" Then
                            planilha.Sheets(1).Cells(k, f).Value = ""
                        End If
                    End If
                    
                    ' --- COLUNAS COM CÁLCULOS ESPECÍFICOS / PROCV ---
                    On Error Resume Next
                    Select Case nomeProcurado_Destino
                        Case "peso", "peso bruto"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:X"), 7, False)
                        Case "dimensoes"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:U"), 4, False)
                        Case "peso liquido", "peso produto"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:V"), 5, False)
                        Case "altura"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:AA"), 10, False)
                        Case "largura"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:Z"), 9, False)
                        Case "comprimento"
                            planilha.Sheets(1).Cells(k, f).Value = Application.VLookup(CodSubclasse, ap.Range("R:Y"), 8, False)
                        
                        Case "código ean / gtin"
                            If LCase(db.Cells(j, Coluna_Tamanho).Text) = "u" Then
                                planilha.Sheets(1).Cells(k, f).Value = db.Cells(j, Coluna_CodigoEANVariantes)
                            Else
                                Tamanho = Trim(db.Cells(j, Coluna_Tamanho).Value)
                                If Tamanho <> "" Then
                                    itens = Split(Tamanho, ";")
                                    resultado = "0"
                                    For w = 1 To UBound(itens): resultado = resultado & ";0": Next w
                                Else: resultado = "": End If
                                planilha.Sheets(1).Cells(k, f).Value = resultado
                            End If

                        Case "categoria ean variantes"
                            Tamanho = Trim(db.Cells(j, Coluna_Tamanho).Value)
                            If Tamanho <> "" And db.Cells(j, Coluna_CategoriaForn).Value <> "" Then
                                itens = Split(Tamanho, ";")
                                resultado = db.Cells(j, Coluna_CategoriaForn).Value
                                For w = 1 To UBound(itens): resultado = resultado & ";" & db.Cells(j, Coluna_CategoriaForn).Value: Next w
                            Else: resultado = "": End If
                            planilha.Sheets(1).Cells(k, f).Value = resultado

                        Case "categoria sistema"
                            Dim vClasse As String: vClasse = Trim(db.Cells(j, Coluna_Classe).Value)
                            valor_procv = Application.WorksheetFunction.VLookup(vClasse, ap.Range("M:N"), 2, False)
                            planilha.Sheets(1).Cells(k, f).Value = valor_procv

                        Case "texto e-commerce"
                            resultado = Trim(db.Cells(j, Coluna_NomeProduto).Value & " " & db.Cells(j, Coluna_Marca).Value & " " & db.Cells(j, Coluna_Referencia).Value & " " & db.Cells(j, Coluna_Cor).Value)
                            planilha.Sheets(1).Cells(k, f).Value = Application.WorksheetFunction.Proper(resultado)

                        Case "grupo de mercadorias"
                            planilha.Sheets(1).Cells(k, f).Value = db.Cells(j, Coluna_CodSubclasse) & ". " & db.Cells(j, Coluna_Subclasse)

                        Case "grupo de compradores"
                            Select Case LCase(db.Cells(j, Coluna_Setor))
                                Case "setor_a", "setor_b": GrupoCompradores = "COMPRADOR_01"
                                Case "setor_c", "setor_d": GrupoCompradores = "COMPRADOR_02"
                                Case Else: GrupoCompradores = "COMPRADOR_GERAL"
                            End Select
                            planilha.Sheets(1).Cells(k, f).Value = GrupoCompradores

                        Case "preço standard"
                            planilha.Sheets(1).Cells(k, f).Value = IIf(db.Cells(j, Coluna_QtdComprada).Value > 0, db.Cells(j, Coluna_CompraLiquidaTotal).Value / db.Cells(j, Coluna_QtdComprada).Value, 0)

                        ' --- VALORES CONSTANTES ---
                        Case "unidade  medida": planilha.Sheets(1).Cells(k, f) = "UN"
                        Case "tipo macro":      planilha.Sheets(1).Cells(k, f) = "PADRAO_SISTEMA"
                        Case "garantia":        planilha.Sheets(1).Cells(k, f) = "60"
                        Case "volume":          planilha.Sheets(1).Cells(k, f).Value = 6
                        Case "lead time":       planilha.Sheets(1).Cells(k, f).Value = 30
                    End Select
                    On Error GoTo 0
                Next cell
            Next i
        End If
    Next j
    
    ' --- FINALIZAÇÃO ---
    planilha.Save
    planilha.Close

    Call bBloqueio

    With Controle_Macro
        last_row_log = .Cells(Rows.Count, "B").End(xlUp).Row + 1
        .Range("A" & last_row_log).Value = "Processamento Padrao"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Finalizada"
    End With

    MsgBox "Processamento concluído com sucesso!", vbInformation
    Application.ScreenUpdating = True

End Sub
