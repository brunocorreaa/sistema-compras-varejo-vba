VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_LinhaClone 
   Caption         =   "Formulário de Clone - Lojas Lebes"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "form_LinhaClone.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_LinhaClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' FORMULÁRIO: form_LinhaClone
' DESCRIÇÃO: Interface para clonagem e edição de parâmetros de registros.
' =========================================================================

Private Sub CommandButton1_Click()

    ' Armazena as seleções na aba de controle para processamento posterior
    With Sheets("Painel_Principal")
        .Cells(2, 3) = CaixaCombinacao_Grupo.Value
        .Cells(2, 4) = CaixaCombinacao_Classe.Value
        .Cells(2, 5) = CaixaCombinacao_Subclasse.Value
        .Cells(2, 6) = CaixaCombinacao_TipoEnvio.Value
        .Cells(2, 7) = CaixaCombinacao_Objetivo.Value
        .Cells(2, 8) = CaixaCombinacao_MetodoDist.Value
        .Cells(2, 9) = CaixaCombinacao_Segmento.Value
        .Cells(2, 10) = CaixaCombinacao_Ano.Value
        .Cells(2, 11) = CaixaCombinacao_Periodo.Value
    End With
    
    Unload Me
        
End Sub

Private Sub CommandButton2_Click()

    ' Sinaliza cancelamento via variável global para interromper macros chamadoras
    CancelamentoSolicitado = True
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    ' Declarações de Escopo Local
    Dim wsDB As Worksheet, wsMenu As Worksheet, wsApoio As Worksheet
    Dim ultLinha As Long, i As Integer
    Dim colGrupo, colClasse, colSub, colEnvio, ColTarget, colDist, colTema, colAno, colSem As Integer
    Dim nomesColunas As Variant, nome As Variant
    Dim Celula As Range
    
    ' Atribuições de Objetos (Anonimizados)
    Set wsDB = ThisWorkbook.Sheets("BASE_REGISTROS") ' Antigo LINHAS_COLECAO
    Set wsMenu = ThisWorkbook.Sheets("Painel_Principal")
    Set wsApoio = ThisWorkbook.Sheets("Configuracoes")

    ' Recupera o ponteiro do registro ativo selecionado na interface
    Valor_Linha = wsMenu.Cells(2, 2).Value
    
    ' Busca ID técnico do registro na base de dados
    Valor_ID_Registro = wsDB.Cells(Valor_Linha, 2).Value
    
    ' Determina extensão da base para validações
    last_row = wsDB.Cells(Rows.Count, "B").End(xlUp).Row

    ' -----------------------------------------------------------------
    ' MAPEAMENTO DINÂMICO DE COLUNAS
    ' Objetivo: Localizar índices independente da posição da coluna
    ' -----------------------------------------------------------------
    nomesColunas = Array("Grupo", "Classe", "Subclasse", "Tipo_Operacao", "Alvo", "Logistica", "Categoria", "Ano", "Ciclo")

    For Each nome In nomesColunas
        ' Varredura no cabeçalho (Linha 2)
        For Each Celula In wsDB.Rows(2).Cells
            If Celula.Value <> "" Then
                ' Mapeia variáveis de coluna conforme cabeçalho encontrado
                Select Case Celula.Value
                    Case "Grupo": colGrupo = Celula.Column
                    Case "Classe": colClasse = Celula.Column
                    Case "Subclasse": colSub = Celula.Column
                    Case "Tipo_Operacao": colEnvio = Celula.Column
                    Case "Alvo": ColTarget = Celula.Column
                    Case "Logistica": colDist = Celula.Column
                    Case "Categoria": colTema = Celula.Column
                    Case "Ano": colAno = Celula.Column
                    Case "Ciclo": colSem = Celula.Column
                End Select
            End If
        Next Celula
    Next nome

    ' -----------------------------------------------------------------
    ' CARREGAMENTO DAS CAIXAS DE COMBINAÇÃO (Comboboxes)
    ' -----------------------------------------------------------------

    ' Fontes de Dados Dinâmicas da aba de Configurações
    With wsApoio
        CaixaCombinacao_Grupo.RowSource = .Name & "!O2:O" & .Range("O1").End(xlDown).Row
        CaixaCombinacao_Classe.RowSource = .Name & "!P2:P" & .Range("P1").End(xlDown).Row
        CaixaCombinacao_Subclasse.RowSource = .Name & "!Q2:Q" & .Range("Q1").End(xlDown).Row
    End With

    ' Itens Fixos: Tipo de Operação / Logística
    With CaixaCombinacao_TipoEnvio
        .List = Array("Operação_Principal", "Fase_01", "Fase_02", "Fase_03", "Fase_04", "Fase_05", "Emergencial")
    End With

    With CaixaCombinacao_Objetivo
        .List = Array("Apresentação", "Manutenção")
    End With

    With CaixaCombinacao_MetodoDist
        .List = Array("Modelo_A", "Modelo_B")
    End With

    With CaixaCombinacao_Segmento
        .List = Array("Perfil_Standard", "Perfil_Essencial", "Perfil_Premium")
    End With

    ' Datas Dinâmicas (Ano Atual e Subsequente)
    With CaixaCombinacao_Ano
        .Clear
        .AddItem Year(Date)
        .AddItem Year(Date) + 1
    End With

    With CaixaCombinacao_Periodo
        .List = Array("1º Semestre", "2º Semestre")
    End With

    ' -----------------------------------------------------------------
    ' PREENCHIMENTO DOS VALORES ATUAIS DO REGISTRO
    ' -----------------------------------------------------------------
    CaixaCombinacao_Grupo = wsDB.Cells(Valor_Linha, colGrupo).Value
    CaixaCombinacao_Classe = wsDB.Cells(Valor_Linha, colClasse).Value
    CaixaCombinacao_Subclasse = wsDB.Cells(Valor_Linha, colSub).Value
    CaixaCombinacao_TipoEnvio = wsDB.Cells(Valor_Linha, colEnvio).Value
    CaixaCombinacao_Objetivo = wsDB.Cells(Valor_Linha, ColTarget).Value
    CaixaCombinacao_MetodoDist = wsDB.Cells(Valor_Linha, colDist).Value
    CaixaCombinacao_Segmento = wsDB.Cells(Valor_Linha, colTema).Value
    CaixaCombinacao_Ano = wsDB.Cells(Valor_Linha, colAno).Value
    CaixaCombinacao_Periodo = wsDB.Cells(Valor_Linha, colSem).Value

    ' Bloqueio de ID para evitar alteração acidental de chave primária
    CaixaCombinacao_NumeroLinhaClone = Valor_ID_Registro
    CaixaCombinacao_NumeroLinhaClone.Locked = True

End Sub
