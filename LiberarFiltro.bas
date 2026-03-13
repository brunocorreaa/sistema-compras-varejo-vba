Attribute VB_Name = "LiberarFiltro"
Sub AlternarFiltrosProtegidos()

    Application.ScreenUpdating = False
    
    ' --- DECLARAÇÕES E ATRIBUIÇÕES ---
    Dim ws As Worksheet
    Dim mn As Worksheet
    Dim Controle_Macro As Worksheet
    Dim sheetsArray As Variant
    Dim hasfilter As Boolean
    Dim sheetName As Variant
    
    ' Atribuições de abas principais
    Set mn = ThisWorkbook.Sheets("Menu")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
    
    ' Array de abas que sofrerão a ação (Nomes anonimizados)
    sheetsArray = Array("DADOS_PRINCIPAIS", "Apoio", "Registros")

    ' Auditoria de execução
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    ' Registro do início da macro
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Gestao de Filtros"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Iniciada"
    End With

    ' Procedimentos de acesso
    Call bDesbloqueio
    
    ' --- ITERAÇÃO SOBRE AS PLANILHAS ---
    For Each sheetName In sheetsArray
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Desproteção com senha anonimizada
        ws.Unprotect "SENHA_SISTEMA"
        
        hasfilter = False
        
        ' Verifica a existência de autofiltro ativo
        If ws.AutoFilterMode Then
            If Not ws.AutoFilter Is Nothing Then
                ' Define a linha de cabeçalho dependendo da aba
                If ws.AutoFilter.Range.Rows(1).Row = IIf(sheetName = "DADOS_PRINCIPAIS", 2, 1) Then
                    hasfilter = True
                End If
            End If
        End If
        
        ' Lógica: Se houver filtro, limpa os dados; se não houver, aplica o filtro
        With ws.Rows(IIf(sheetName = "DADOS_PRINCIPAIS", 2, 1))
            If hasfilter Then
                If ws.AutoFilter.Filters.Count > 0 Then
                    On Error Resume Next
                    ws.ShowAllData
                    On Error GoTo 0
                End If
            Else
                ' Aplica o filtro na região preenchida da linha de cabeçalho
                ws.Range("A" & .Row, ws.Cells(.Row, ws.Columns.Count).End(xlToLeft)).AutoFilter
            End If
        End With
    Next sheetName

    ' Reestabelece bloqueios de segurança
    Call bBloqueio
    
    ' Registro da finalização da macro
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Gestao de Filtros"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Finalizada"
    End With

    MsgBox "Filtros atualizados com sucesso!", vbInformation
    
    Application.ScreenUpdating = True

End Sub
