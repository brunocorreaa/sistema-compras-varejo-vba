Attribute VB_Name = "Ocultar_Colunas"
Sub GerenciarVisibilidadeColunas()

    Application.ScreenUpdating = False
    
    ' --- DECLARAÇÕES ---
    Dim db As Worksheet
    Dim ap As Worksheet
    Dim mn As Worksheet
    Dim Controle_Macro As Worksheet
    Dim Controle_Aba As Worksheet
    Dim i As Long, j As Long, last_row_j As Long
            
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_PRINCIPAIS")
    Set ap = ThisWorkbook.Sheets("Apoio")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
    Set Controle_Aba = ThisWorkbook.Sheets("Config-Abas")
    
    ' Auditoria de execução
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")

    ' Registro do início da macro
    last_row = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row).Value = "Ocultar Colunas"
        .Range("B" & last_row).Value = hoje
        .Range("C" & last_row).Value = horaAtual
        .Range("D" & last_row).Value = usuario
        .Range("E" & last_row).Value = "Iniciada"
    End With

    ' Procedimentos auxiliares
    Call Validacoes("")
    Call bDesbloqueio
    
    ' --- RESET DE VISIBILIDADE ---
    ' Exibe todas as colunas antes de aplicar a nova regra
    db.Select
    db.Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = False
    
    ' --- LÓGICA DE OCULTAÇÃO DINÂMICA ---
    ' Define o limite baseado na lista de colunas a ocultar (Coluna AB na aba Apoio)
    last_row = ap.Range("AB1").End(xlDown).Row

    For i = 2 To last_row
        
        ' Identifica a última coluna com dados no cabeçalho (Linha 2)
        last_row_j = db.Range("B2").End(xlToRight).Column
        
        For j = 1 To last_row_j
            
            ' Se o nome na lista de apoio (coluna 28/AB) for igual ao cabeçalho, oculta a coluna
            If ap.Cells(i, 28).Value = db.Cells(2, j).Value Then
                
                db.Columns(j).Hidden = True
                
            End If
                    
        Next j
            
    Next i

    ' --- FINALIZAÇÃO E ATUALIZAÇÃO ---
    Call bBloqueio
    Call AtualizarRelatorios

    ' Registro da finalização no log
    last_row = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row).Value = "Ocultar Colunas"
        .Range("B" & last_row).Value = hoje
        .Range("C" & last_row).Value = horaAtual
        .Range("D" & last_row).Value = usuario
        .Range("E" & last_row).Value = "Finalizada"
    End With
        
    MsgBox "Configuração de colunas aplicada!", vbInformation

    Application.ScreenUpdating = True

End Sub
