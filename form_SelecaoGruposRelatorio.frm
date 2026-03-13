VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_SelecaoGruposRelatorio 
   Caption         =   "Formulário de Geração Relatório"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "form_SelecaoGruposRelatorio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_SelecaoGruposRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================
' FORMULÁRIO: form_Filtros
' FUNÇÃO: Coletar critérios de filtro múltiplo (Grupos, Classes, Status, Ano)
' =========================================================================

Private Sub CommandButton1_Click()
    Dim mn As Worksheet: Set mn = ThisWorkbook.Sheets("Menu")

    ' Coleta as seleções usando uma função auxiliar para evitar repetição de código
    Dim strGrupos As String: strGrupos = GetSelectedItems(Me.ListBox_Grupos)
    Dim strClasses As String: strClasses = GetSelectedItems(Me.ListBox_Classes)
    Dim strAcoes As String: strAcoes = GetSelectedItems(Me.ListBox_Acao)
    Dim strStatus As String: strStatus = GetSelectedItems(Me.ListBox_Status)

    ' Validação: Impede filtros vazios que poderiam ocultar todos os dados
    If strGrupos = "" And strClasses = "" Then
        MsgBox "Selecione ao menos um critério para filtrar.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Persistência dos Critérios na Aba Menu
    ' Essas células servirão de base para a Macro de Filtro Automático
    With mn
        .Cells(3, 3).Value = strGrupos
        .Cells(3, 4).Value = strClasses
        .Cells(3, 5).Value = strAcoes
        .Cells(3, 6).Value = strStatus
        .Cells(3, 7).Value = Me.ListBox_Ano.Value
        .Cells(3, 8).Value = Me.ListBox_Semestre.Value
    End With
    
    Unload Me
End Sub

' --- FUNÇÃO AUXILIAR: GetSelectedItems ---
' Transforma os itens selecionados de uma ListBox em uma string única
' -------------------------------------------------------------------------
Private Function GetSelectedItems(lst As MSForms.ListBox) As String
    Dim i As Long
    Dim res As String
    
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            Dim val As String: val = lst.List(i)
            ' Tratamento para critério de célula vazia
            If val = "Vazio" Then val = "="
            
            res = res & IIf(res = "", "", ",") & val
        End If
    Next i
    GetSelectedItems = res
End Function

' --- Selecionar Todos ---
Private Sub CheckBox_SelecionarTodosGrupos_Click()
    SetListItemsSelection Me.ListBox_Grupos, Me.CheckBox_SelecionarTodosGrupos.Value
End Sub

Private Sub SetListItemsSelection(lst As MSForms.ListBox, bState As Boolean)
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = bState
    Next i
End Sub
