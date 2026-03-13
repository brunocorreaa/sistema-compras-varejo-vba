VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_LinhaEdicao 
   Caption         =   "Formulário de Edição - Lojas Lebes"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865.001
   OleObjectBlob   =   "form_LinhaEdicao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_LinhaEdicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' FORMULÁRIO: form_Editor_Registros
' EVENTO: UserForm_Initialize
' DESCRIÇÃO: Carrega dados de produtos/pedidos para interface de edição,
'            utilizando mapeamento dinâmico de cabeçalhos e chaves compostas.
' =========================================================================

Private Sub UserForm_Initialize()

    ' Declarações de Escopo Local
    Dim wsDB As Worksheet, wsMenu As Worksheet, wsApoio As Worksheet, wsPedidos As Worksheet
    Dim last_row As Long, i As Integer
    Dim ChaveComposta As String
    Dim nomesColunas As Variant, nome As Variant
    Dim Celula As Range
    Dim resultado_vlookup As String
    Dim EmpresaReferencia As String
    Dim intervalo_pgto As Range
    
    ' Atribuições de Objetos (Anonimizados)
    Set wsDB = ThisWorkbook.Sheets("BASE_DADOS_PRODUTOS") ' Antigo LINHAS_COLECAO
    Set wsMenu = ThisWorkbook.Sheets("Painel_Controle")
    Set wsApoio = ThisWorkbook.Sheets("Parametros_Sistema")
    Set wsPedidos = ThisWorkbook.Sheets("BASE_PEDIDOS")

    ' Identificação da linha de referência baseada na seleção do usuário
    last_row = wsDB.Cells(Rows.Count, "B").End(xlUp).Row
    Valor_Linha = wsMenu.Cells(3, 2).Value
    Valor_Linha_ID = wsMenu.Cells(3, 3).Value

    ' -----------------------------------------------------------------
    ' MAPEAMENTO DINÂMICO DE COLUNAS (Cabeçalhos)
    ' -----------------------------------------------------------------
    nomesColunas = Array("Ciclo", "Ano_Ref", "ID_Subclasse", "Sazonalidade", "Nivel_Cluster", _
                         "Parceiro_Comercial", "Descricao_Item", "SKU_Referencia", "Brand", _
                         "Variante_Cor", "Preco_Final", "Preco_Custo", "Volume_Planejado", _
                         "Categoria_Estilo", "Data_Limite", "Modalidade_Pedido", "Grade_Tamanho", _
                         "Especificacao", "Flag_PE", "EAN_Variante", "ID_Faturamento")

    For Each nome In nomesColunas
        For Each Celula In wsDB.Rows(2).Cells
            If Celula.Value = nome Then
                Select Case nome
                    Case "Sazonalidade": colEstacao = Celula.Column
                    Case "Nivel_Cluster": colCluster = Celula.Column
                    Case "Parceiro_Comercial": colFornecedor = Celula.Column
                    Case "Descricao_Item": colNomeProd = Celula.Column
                    Case "SKU_Referencia": colRef = Celula.Column
                    Case "Brand": colMarca = Celula.Column
                    Case "Variante_Cor": colCor = Celula.Column
                    Case "Preco_Final": colVenda = Celula.Column
                    Case "Preco_Custo": colBruto = Celula.Column
                    Case "Volume_Planejado": colQtd = Celula.Column
                    Case "Categoria_Estilo": colTema = Celula.Column
                    Case "Data_Limite": colData = Celula.Column
                    Case "Modalidade_Pedido": colTipo = Celula.Column
                    Case "Grade_Tamanho": colTam = Celula.Column
                    Case "Especificacao": colGrade = Celula.Column
                    Case "Flag_PE": colPE = Celula.Column
                    Case "EAN_Variante": colEAN = Celula.Column
                    Case "ID_Subclasse": colSub = Celula.Column
                    Case "Ano_Ref": colAno = Celula.Column
                    Case "Ciclo": colSem = Celula.Column
                End Select
                Exit For
            End If
        Next Celula
    Next nome

    ' -----------------------------------------------------------------
    ' CARREGAMENTO DOS DADOS (MODO SOMENTE LEITURA - DISPLAY)
    ' -----------------------------------------------------------------
    
    ' Preenchimento de labels/comboboxes bloqueados para visualização
    Caixa_Sazonalidade = wsDB.Cells(Valor_Linha, colEstacao).Value
    Caixa_Sazonalidade.Locked = True

    Caixa_Cluster = wsDB.Cells(Valor_Linha, colCluster).Value
    Caixa_Cluster.Locked = True

    Caixa_Parceiro = wsDB.Cells(Valor_Linha, colFornecedor).Value
    Caixa_Parceiro.Locked = True

    ' Criação de Chave Única para busca cruzada em outras bases
    ' Chave: ID_Linha & Subclasse & Ano & Ciclo
    ChaveComposta = wsDB.Cells(Valor_Linha, 2) & "-" & _
                    wsDB.Cells(Valor_Linha, colSub) & "-" & _
                    wsDB.Cells(Valor_Linha, colAno) & "-" & _
                    wsDB.Cells(Valor_Linha, colSem)
    
    ' Exemplo de VLOOKUP e SUMIF via VBA
    ' Busca Preço de Custo na base de Pedidos
    Caixa_PrecoCusto = WorksheetFunction.VLookup(ChaveComposta, wsPedidos.Range("C:V"), 20, 0)
    Caixa_PrecoCusto.Locked = True

    ' Soma quantidade total baseada na chave composta
    Caixa_SomaQtd.Value = WorksheetFunction.SumIf(wsPedidos.Range("C:C"), ChaveComposta, wsPedidos.Range("U:U"))
    Caixa_SomaQtd.Locked = True

    ' -----------------------------------------------------------------
    ' LÓGICA DE CONDIÇÃO DE PAGAMENTO (REGRA DE NEGÓCIO)
    ' -----------------------------------------------------------------
    If UCase(wsApoio.Cells(1, "AJ").Value) = "BRAND" Then
        EmpresaReferencia = wsDB.Cells(Valor_Linha, colMarca)
    Else
        EmpresaReferencia = wsDB.Cells(Valor_Linha, colFornecedor)
    End If
    
    ' Busca código de condição de pagamento e formata com zeros à esquerda
    Set intervalo_pgto = wsApoio.Range("C:D")
    resultado_vlookup = Application.WorksheetFunction.VLookup(EmpresaReferencia, intervalo_pgto, 2, False)
    resultado_vlookup = String(4 - Len(resultado_vlookup), "0") & resultado_vlookup
    
    CaixaCondicaoPGTO = resultado_vlookup
    CaixaCondicaoPGTO.Locked = True

    ' -----------------------------------------------------------------
    ' CARREGAMENTO DOS CAMPOS PARA EDIÇÃO (INPUTS)
    ' -----------------------------------------------------------------
    
    With CaixaMotivo_Edit
        .List = Array("Ajuste Operacional", "Divergência Fornecedor", "Antecipação Fluxo", "Prorrogação Prazo")
    End With

    ' Carregamento de RowSource dinâmico para Combos de Edição
    With wsApoio
        Caixa_Sazonalidade_Edit.RowSource = .Name & "!K2:K" & .Range("K1").End(xlDown).Row
        Caixa_Cluster_Edit.RowSource = .Name & "!H2:H" & .Range("H1").End(xlDown).Row
    End With

    ' Replica valores atuais para os campos de edição
    Caixa_NomeProduto_Edit = wsDB.Cells(Valor_Linha, colNomeProd).Value
    Caixa_Referencia_Edit = wsDB.Cells(Valor_Linha, colRef).Value
    Caixa_PrecoVenda_Edit = wsDB.Cells(Valor_Linha, colVenda).Value
    
    With Caixa_Modalidade_Edit
        .List = Array("Modelo_A", "Modelo_B", "Modelo_C")
    End With
    Caixa_Modalidade_Edit = wsDB.Cells(Valor_Linha, colTipo).Value

    ' ID de Auditoria (Bloqueado)
    Caixa_ID_Auditoria = Valor_Linha_ID
    Caixa_ID_Auditoria.Locked = True
    
End Sub
' =========================================================================
' FORMULÁRIO: form_LinhaEdicao
' EVENTO: CommandButton1_Click (Salvar Alterações)
' DESCRIÇÃO: Processa a edição de registros, valida integridade de dados,
'            verifica saldo orçamentário (OTB) e gera logs de auditoria.
' =========================================================================

Private Sub btn_SalvarEdicao_Click()

    ' Declarações de Objetos e Variáveis
    Dim wsDB As Worksheet, wsMenu As Worksheet, wsErro As Worksheet, wsEdicao As Worksheet
    Dim colunasMapeadas As Variant, j As Integer
    Dim lastRowMenu As Long, lastRowErro As Long, lastRowAudit As Long
    Dim valorCLTNovo As Double, valorCLTAntigo As Double, saldoCML As Double
    Dim chaveOrcamento As String, usuario As String, dataExec As Date, horaExec As String

    ' Inicialização de Referências
    Set wsDB = ThisWorkbook.Sheets("LINHAS_COLECAO")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsErro = ThisWorkbook.Sheets("Controle-Erro")
    Set wsEdicao = ThisWorkbook.Sheets("Controle-Edicao")
    
    usuario = Environ("Username")
    dataExec = Date
    horaExec = Format(Time, "hh:mm:ss")
    
    ' Identificação da linha alvo (recuperada do contexto do formulário)
    idxLinhaDB = wsMenu.Cells(3, 2).Value
    idxLinhaPlano = wsMenu.Cells(3, 3).Value

    ' -----------------------------------------------------------------
    ' 1. VALIDAÇÕES DE INTEGRIDADE
    ' -----------------------------------------------------------------
    
    ' Validação: Preços e Quantidades devem ser numéricos
    If Not IsNumeric(txt_PrecoVenda_Edit) Or Not IsNumeric(txt_PrecoCusto_Edit) Then
        RegistrarErro "Erro de Tipo: Preço informado inválido", wsErro, dataExec, horaExec, usuario
        MsgBox "Verifique os valores de preço informados.", vbCritical
        Exit Sub
    End If

    ' Validação: Sincronismo de Grades (Tamanho vs EAN)
    ' Garante que a quantidade de tamanhos separados por ";" é igual à de EANs
    cntTam = Len(txt_Tamanho_Edit) - Len(Replace(txt_Tamanho_Edit, ";", ""))
    cntEAN = Len(txt_EAN_Edit) - Len(Replace(txt_EAN_Edit, ";", ""))
    
    If cntTam <> cntEAN Then
        RegistrarErro "Erro de Grade: Descompasso entre Tamanhos e EANs", wsErro, dataExec, horaExec, usuario
        MsgBox "A quantidade de tamanhos deve ser igual à quantidade de códigos EAN.", vbExclamation
        Exit Sub
    End If

    ' -----------------------------------------------------------------
    ' 2. REGRA DE NEGÓCIO: GESTÃO DE VERBA (OTB)
    ' -----------------------------------------------------------------
    
    ' Cálculo de impacto financeiro (Compra Líquida Total - CLT)
    ' Considera impostos dinâmicos recuperados da base (PIS/COFINS/ICMS)
    taxasFixas = 0.076 + 0.0165 ' PIS + COFINS
    taxaICMS = wsDB.Cells(idxLinhaDB, colICMS).Value
    
    valorCLTNovo = txt_Qtd_Edit * txt_Custo_Edit * (1 - (taxasFixas + taxaICMS))
    valorCLTAntigo = txt_Qtd_Original * txt_Custo_Original * (1 - (taxasFixas + taxaICMS))
    
    ' Se houver aumento de custo ou mudança de período, valida saldo orçamentário
    If valorCLTNovo > valorCLTAntigo Or mesOriginal <> mesEditado Then
        
        chaveOrcamento = Year(txt_DataEntrega_Edit) & Format(txt_DataEntrega_Edit, "mmmm") & setor
        saldoCML = Application.WorksheetFunction.VLookup(chaveOrcamento, Sheets("OTB").Range("A:N"), 14, 0)
        
        If (saldoCML - (valorCLTNovo - valorCLTAntigo)) < 0 Then
            RegistrarErro "Estouro de Verba: Tentativa de edição sem saldo OTB", wsErro, dataExec, horaExec, usuario
            MsgBox "Atenção! Verba insuficiente para este ajuste." & vbCrLf & "Saldo: " & FormatCurrency(saldoCML), vbCritical
            Exit Sub
        End If
    End If

    ' -----------------------------------------------------------------
    ' 3. AUDITORIA
    ' -----------------------------------------------------------------
    
    ' Validação: O usuário realmente alterou algo?
    ' Comparação binarizada entre campos originais e editados (Resumo no Menu)
    If VerificarMudancas() = False Then
        MsgBox "Nenhuma alteração detectada.", vbInformation
        Exit Sub
    End If

    ' Transfere dados validados do formulário para a Base Principal
    colunasAlvo = Array(colEstacao, colCluster, colFornecedor, colNome, colRef, colVenda, colCusto, colQtd, colData)
    
    For i = LBound(colunasAlvo) To UBound(colunasAlvo)
        ' Tratamento especial para datas para evitar inversão de formato (PT-BR vs US)
        If colunasAlvo(i) = colData Then
            wsDB.Cells(idxLinhaDB, colunasAlvo(i)) = CDate(txt_DataEntrega_Edit)
        Else
            wsDB.Cells(idxLinhaDB, colunasAlvo(i)) = ObterValorCampo(i)
        End If
    Next i

    ' Grava Log de Auditoria: Quem, Quando e Por que alterou?
    lastRowAudit = wsEdicao.Cells(Rows.Count, 1).End(xlUp).Row + 1
    With wsEdicao
        .Cells(lastRowAudit, 1) = dataExec
        .Cells(lastRowAudit, 2) = horaExec
        .Cells(lastRowAudit, 3) = usuario
        .Cells(lastRowAudit, 4) = combo_MotivoEdicao.Value
    End With

    MsgBox "Edição realizada com sucesso e log de auditoria gerado.", vbInformation
    Unload Me

End Sub
