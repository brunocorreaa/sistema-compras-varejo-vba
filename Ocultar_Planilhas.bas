Attribute VB_Name = "Ocultar_Planilhas"
Sub OcultarPlanilhasSistema()

    ' Ocultar abas administrativas e de sistema (Nível VeryHidden)
    
    Sheets("Dados_Orcamento").Visible = xlSheetVeryHidden
    Sheets("Config-Abas").Visible = xlSheetVeryHidden
    Sheets("Config-Log_Macros").Visible = xlSheetVeryHidden
    Sheets("Config-Erros").Visible = xlSheetVeryHidden
    Sheets("Config-Edicao").Visible = xlSheetVeryHidden
    Sheets("Config-Arquivos").Visible = xlSheetVeryHidden
    Sheets("Resultados_KPI").Visible = xlSheetVeryHidden

End Sub

Sub MostrarPlanilhasSistema()

    ' Tornar visíveis abas de operação e configuração
    
    Sheets("Painel_Operacional").Visible = xlSheetVisible
    Sheets("Dados_Orcamento").Visible = xlSheetVisible
    Sheets("Config-Abas").Visible = xlSheetVisible
    Sheets("Config-Log_Macros").Visible = xlSheetVisible
    Sheets("Config-Erros").Visible = xlSheetVisible
    Sheets("Config-Edicao").Visible = xlSheetVisible
    Sheets("Config-Arquivos").Visible = xlSheetVisible
    Sheets("Menu").Visible = xlSheetVisible
    Sheets("Resultados_KPI").Visible = xlSheetVisible

End Sub
