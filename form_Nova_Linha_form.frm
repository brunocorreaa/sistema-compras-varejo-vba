VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Nova_Linha_form 
   Caption         =   "Formulário de Nova Linha - Lojas Lebes"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "form_Nova_Linha_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Nova_Linha_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' FORMULÁRIO: form_Nova_Linha_form
' FUNÇÃO: Capturar metadados para criação de novos registros no plano
' =========================================================================

Private Sub CommandButton1_Click()
    ' Validação de Preenchimento: Garante a integridade dos dados antes da gravação
    Dim controles As Variant, c As Variant
    controles = Array(CaixaCombinacao_Grupo, CaixaCombinacao_Classe, CaixaCombinacao_Subclasse, _
                      CaixaCombinacao_EnvioRep, CaixaCombinacao_Target, CaixaCombinacao_Distribuicao, _
                      CaixaCombinacao_Tema, CaixaCombinacao_Ano, CaixaCombinacao_Semestre, CaixaCombinacao_Setor)

    For Each c In controles
        If c.Value = "" Then
            MsgBox "Atenção: Todos os campos são obrigatórios!", vbExclamation, "Validação"
            ' Reinicia o form para garantir que o Initialize rode novamente se necessário
            Unload Me
            form_Nova_Linha_form.Show
            Exit Sub
        End If
    Next c

    ' Persistência dos Dados na Aba Menu (Buffer de configuração)
    With Sheets("Menu")
        .Cells(2, 3) = CaixaCombinacao_Setor.Value
        .Cells(2, 4) = CaixaCombinacao_Grupo.Value
        .Cells(2, 5) = CaixaCombinacao_Classe.Value
        .Cells(2, 6) = CaixaCombinacao_Subclasse.Value
        .Cells(2, 7) = CaixaCombinacao_EnvioRep.Value
        .Cells(2, 8) = CaixaCombinacao_Target.Value
        .Cells(2, 9) = CaixaCombinacao_Distribuicao.Value
        .Cells(2, 10) = CaixaCombinacao_Tema.Value
        .Cells(2, 11) = CaixaCombinacao_Ano.Value
        .Cells(2, 12) = CaixaCombinacao_Semestre.Value
    End With
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim db As Worksheet: Set db = ThisWorkbook.Sheets("LINHAS_COLECAO")
    Dim ap As Worksheet: Set ap = ThisWorkbook.Sheets("Apoio")
    
    ' --- MAPEAMENTO DE COLUNAS ---
    ' Evita quebras de código caso a estrutura da planilha mude
    Dim nomesColunas As Variant: nomesColunas = Array("Setor", "Grupo", "Classe", "Subclasse", "Envio / Reposição", "Target", "Distribuição", "Tema", "Ano", "Semestre")
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    
    For Each nome In nomesColunas
        For Each Celula In db.Rows(2).Cells
            If Celula.Value = nome Then
                colMap.Add nome, Celula.Column
                Exit For
            End If
        Next Celula
    Next nome

    ' --- CARREGAMENTO DE OPÇÕES (COMBOBOXES) ---
    
    ' Setor: Regra específica para Mix Infantil
    Dim setorBase As String: setorBase = UCase(db.Cells(3, colMap("Setor")).Text)
    If setorBase = "INFANTIL MENINO" Or setorBase = "INFANTIL MENINA" Then
        With CaixaCombinacao_Setor
            .AddItem "INFANTIL MENINO"
            .AddItem "INFANTIL MENINA"
        End With
    Else
        CaixaCombinacao_Setor.Value = setorBase
    End If

    ' Fontes de Dados Dinâmicas (Aba Apoio)
    ' O uso de RowSource conecta o formulário diretamente às listas de cadastro
    On Error Resume Next ' Prevenção para listas vazias
    CaixaCombinacao_Grupo.RowSource = "Apoio!O2:O" & ap.Cells(Rows.Count, "O").End(xlUp).Row
    CaixaCombinacao_Classe.RowSource = "Apoio!P2:P" & ap.Cells(Rows.Count, "P").End(xlUp).Row
    CaixaCombinacao_Subclasse.RowSource = "Apoio!Q2:Q" & ap.Cells(Rows.Count, "Q").End(xlUp).Row
    On Error GoTo 0

    ' Datas Inteligentes: Sempre oferece o Ano Vigente e o Próximo
    With CaixaCombinacao_Ano
        .AddItem Year(Date)
        .AddItem Year(Date) + 1
    End With

End Sub

