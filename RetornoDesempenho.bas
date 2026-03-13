Attribute VB_Name = "RetornoDesempenho"
Sub ProcessarRetornoPerformance()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' --- DECLARAÇÕES ---
    Dim wbSelecionado As Workbook
    Dim db As Worksheet, ap As Worksheet, mn As Worksheet, rd As Worksheet, Controle_Macro As Worksheet
    Dim arquivo_importacao As String
    Dim lastRowA As Long, lastRowB As Long, repetRow As Long
    Dim keyA As String, keyB As String
    Dim dictB As Object
    Dim cabecalhoA As Variant, cabecalhoB As Variant
    
    ' Colunas Mapeadas (A = Origem / B = Destino)
    Dim colA_Parceiro As Long, colA_ID As Long, colA_Setor As Long, colA_Grupo As Long
    Dim colA_Classe As Long, colA_Subclasse As Long, colA_Codigo As Long, colA_Categoria As Long
    Dim colA_Ciclo As Long, colA_ValorRef As Long, colA_QtdTotal As Long, colA_Obs As Long
    Dim colA_Operacao As Long, colA_Situacao As Long
    
    Dim colB_Parceiro As Long, colB_ID As Long, colB_Setor As Long, colB_Grupo As Long
    Dim colB_Classe As Long, colB_Subclasse As Long, colB_Codigo As Long, colB_Categoria As Long
    Dim colB_Ciclo As Long, colB_ValorRef As Long, colB_QtdTotal As Long, colB_Obs As Long
    Dim colB_Operacao As Long, colB_Situacao As Long, colB_Timestamp As Long
    
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_MESTRE")
    Set ap = ThisWorkbook.Sheets("Apoio")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set rd = ThisWorkbook.Sheets("Historico_Performance")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
    Set dictB = CreateObject("Scripting.Dictionary")
    
    ' Auditoria
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1

    With Controle_Macro
        .Range("A" & last_row_log).Value = "Importação Performance"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Iniciada"
    End With

    Call Validacoes("")

    ' --- SELEÇÃO DE ARQUIVO ---
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o arquivo de dados externos:"
        .Filters.Clear
        .Filters.Add "Arquivos de Dados:", "*.xlsx"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            arquivo_importacao = .SelectedItems(1)
        Else
            MsgBox "Operação cancelada.", vbExclamation
            Exit Sub
        End If
    End With
    
    Set wbSelecionado = Workbooks.Open(arquivo_importacao)
    Set ws = wbSelecionado.Sheets(1) ' Assume a primeira aba do arquivo externo
    
    ' --- MAPEAMENTO DINÂMICO DE CABEÇALHOS (ORIGEM) ---
    cabecalhoA = ws.Rows(3).Value
    For i = LBound(cabecalhoA, 2) To UBound(cabecalhoA, 2)
        Select Case cabecalhoA(1, i)
            Case "Parceiro": colA_Parceiro = i
            Case "ID_Ref": colA_ID = i
            Case "Setor": colA_Setor = i
            Case "Grupo": colA_Grupo = i
            Case "Classe": colA_Classe = i
            Case "Subclasse": colA_Subclasse = i
            Case "Código": colA_Codigo = i
            Case "Categoria": colA_Categoria = i
            Case "Ciclo": colA_Ciclo = i
            Case "Valor": colA_ValorRef = i
            Case "Qtd": colA_QtdTotal = i
            Case "Observação": colA_Obs = i
            Case "Operação": colA_Operacao = i
            Case "Status": colA_Situacao = i
        End Select
    Next i
    
    ' --- MAPEAMENTO DINÂMICO DE CABEÇALHOS (DESTINO) ---
    cabecalhoB = rd.Rows(1).Value
    For i = LBound(cabecalhoB, 2) To UBound(cabecalhoB, 2)
        Select Case cabecalhoB(1, i)
            Case "Parceiro": colB_Parceiro = i
            Case "ID_Ref": colB_ID = i
            Case "Setor": colB_Setor = i
            Case "Grupo": colB_Grupo = i
            Case "Classe": colB_Classe = i
            Case "Subclasse": colB_Subclasse = i
            Case "Código": colB_Codigo = i
            Case "Categoria": colB_Categoria = i
            Case "Ciclo": colB_Ciclo = i
            Case "Valor": colB_ValorRef = i
            Case "Qtd": colB_QtdTotal = i
            Case "Observação": colB_Obs = i
            Case "Operação": colB_Operacao = i
            Case "Status": colB_Situacao = i
            Case "Timestamp": colB_Timestamp = i
        End Select
    Next i
    
    ' --- CARGA DO DICIONÁRIO (VERIFICAR EXISTENTES) ---
    lastRowB = rd.Cells(Rows.Count, "D").End(xlUp).Row
    For j = 2 To lastRowB
        keyB = rd.Cells(j, colB_Parceiro).Value & "|" & rd.Cells(j, colB_ID).Value & "|" & _
               rd.Cells(j, colB_Setor).Value & "|" & rd.Cells(j, colB_Codigo).Value & "|" & _
               rd.Cells(j, colB_ValorRef).Value
               
        If Not dictB.Exists(keyB) Then dictB.Add keyB, j
    Next j
    
    ' --- PROCESSAMENTO E SINCRONIZAÇÃO ---
    lastRowA = ws.Cells(Rows.Count, "D").End(xlUp).Row
    For k = 4 To lastRowA
        keyA = ws.Cells(k, colA_Parceiro).Value & "|" & ws.Cells(k, colA_ID).Value & "|" & _
               ws.Cells(k, colA_Setor).Value & "|" & ws.Cells(k, colA_Codigo).Value & "|" & _
               ws.Cells(k, colA_ValorRef).Value
        
        If dictB.Exists(keyA) Then
            ' Atualiza registro existente
            repetRow = dictB.Item(keyA)
            rd.Cells(repetRow, colB_QtdTotal).Value = ws.Cells(k, colA_QtdTotal).Value
            rd.Cells(repetRow, colB_Obs).Value = ws.Cells(k, colA_Obs).Value
            rd.Cells(repetRow, colB_Operacao).Value = ws.Cells(k, colA_Operacao).Value
            rd.Cells(repetRow, colB_Situacao).Value = ws.Cells(k, colA_Situacao).Value
            rd.Cells(repetRow, colB_Timestamp).Value = Now
        Else
            ' Adiciona novo registro
            lastRowB = rd.Cells(Rows.Count, "D").End(xlUp).Row + 1
            rd.Cells(lastRowB, colB_Parceiro).Value = ws.Cells(k, colA_Parceiro).Value
            rd.Cells(lastRowB, colB_ID).Value = ws.Cells(k, colA_ID).Value
            rd.Cells(lastRowB, colB_Setor).Value = ws.Cells(k, colA_Setor).Value
            rd.Cells(lastRowB, colB_Grupo).Value = ws.Cells(k, colA_Grupo).Value
            rd.Cells(lastRowB, colB_Classe).Value = ws.Cells(k, colA_Classe).Value
            rd.Cells(lastRowB, colB_Subclasse).Value = ws.Cells(k, colA_Subclasse).Value
            rd.Cells(lastRowB, colB_Codigo).Value = ws.Cells(k, colA_Codigo).Value
            rd.Cells(lastRowB, colB_Categoria).Value = ws.Cells(k, colA_Categoria).Value
            rd.Cells(lastRowB, colB_Ciclo).Value = ws.Cells(k, colA_Ciclo).Value
            rd.Cells(lastRowB, colB_ValorRef).Value = ws.Cells(k, colA_ValorRef).Value
            rd.Cells(lastRowB, colB_QtdTotal).Value = ws.Cells(k, colA_QtdTotal).Value
            rd.Cells(lastRowB, colB_Obs).Value = ws.Cells(k, colA_Obs).Value
            rd.Cells(lastRowB, colB_Operacao).Value = ws.Cells(k, colA_Operacao).Value
            rd.Cells(lastRowB, colB_Situacao).Value = ws.Cells(k, colA_Situacao).Value
            rd.Cells(lastRowB, colB_Timestamp).Value = Now
            
            dictB.Add keyA, lastRowB
        End If
    Next k
    
    ' --- FINALIZAÇÃO ---
    wbSelecionado.Close SaveChanges:=False
    Call bBloqueio

    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Importação Performance"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Finalizada"
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    MsgBox "Processamento concluído!", vbInformation

End Sub
