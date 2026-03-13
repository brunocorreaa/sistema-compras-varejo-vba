Attribute VB_Name = "AtualizarLinha"
' Declarações de Variáveis Globais (Anonimizadas)
Dim ws_Base As Worksheet, ws_Suporte As Worksheet, ws_Registros As Worksheet
Dim ws_Processamento As Worksheet, ws_Financeiro As Worksheet

Sub AtualizarDadosSistema()
    
    ' Atribuições de Abas
    Set ws_Base = ThisWorkbook.Sheets("DADOS_MESTRE")
    Set ws_Suporte = ThisWorkbook.Sheets("AUXILIAR")
    Set ws_Registros = ThisWorkbook.Sheets("BASE_REGISTROS")
    Set ws_Processamento = ThisWorkbook.Sheets("PROCESSAMENTO_INTERNO")
    Set ws_Financeiro = ThisWorkbook.Sheets("CONTROLE_ORCAMENTO")
    
    ' Atualizar a planilha de Registros
    With ws_Registros
        
        ' Senha anonimizada (X)
        .Unprotect Password:="XXXXXXXX"
        
        ' Verificar se a tabela de conexão existe e atualizar
        Set tbl = .ListObjects("Tabela_Registros_Enviados")
        
        If Not tbl Is Nothing Then
            tbl.QueryTable.Refresh BackgroundQuery:=True
        Else
            MsgBox "Tabela de registros não encontrada.", vbExclamation
        End If
    
    End With
    
    ' Atualizar planilha de Processamento Interno
    With ws_Processamento
        
        ' Senha anonimizada (X)
        .Unprotect Password:="XXXXXXXX"
        
        ' Verificar se a tabela de conexão existe e atualizar
        Set tbl = .ListObjects("Tabela_Backup_Processamento")
        
        If Not tbl Is Nothing Then
            tbl.QueryTable.Refresh BackgroundQuery:=True
        Else
            MsgBox "Tabela de processamento não encontrada.", vbExclamation
        End If
        
    End With
    
    ' Atualizar Controle Financeiro/Orçamento
    With ws_Financeiro
        
        ' Verificar se a tabela de conexão existe e atualizar
        Set tbl = .ListObjects("Tabela_Orcamento_Anual")
        
        If Not tbl Is Nothing Then
            tbl.QueryTable.Refresh BackgroundQuery:=True
        Else
            MsgBox "Tabela de orçamento não encontrada.", vbExclamation
        End If
        
    End With
    
End Sub
