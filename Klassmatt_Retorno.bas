Attribute VB_Name = "Klassmatt_Retorno"
Sub RetornoDadosSistema()

    Application.ScreenUpdating = False
    
    ' --- DECLARAÇÕES ---
    Dim db As Worksheet, mn As Worksheet, Controle_Macro As Worksheet
    Dim sh_externa As Worksheet
    Dim Coluna, last_row As Integer
    Dim i, j, colNumber As Long
    Dim caminho_pasta, colLetter As String
    Dim plan_externa As Workbook
    Dim Coluna_Filtro As Long, Coluna_Id_Destino As Long, Coluna_Id_Origem As Long
    Dim last_row_externa As Long
    
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_PRINCIPAIS")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")

    ' Confirmação de execução
    resposta = MsgBox("Você realmente deseja executar o RETORNO DE DADOS?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Sistema")
    
    If resposta <> vbYes Then Exit Sub

    ' Auditoria de execução
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Retorno Dados"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Iniciada"
    End With

    ' Procedimentos auxiliares
    Call Validacoes("RetornoDados")
    Call bDesbloqueio

    ' --- MANIPULAÇÃO DO ARQUIVO EXTERNO ---
    caminho_pasta = ThisWorkbook.Path & "\"
    Set plan_externa = Workbooks.Open(caminho_pasta & "modelo_integracao.xlsx")
    
    ' Atribuição da aba de retorno (originalmente 'CRIADAS')
    Set sh_externa = plan_externa.Sheets("RETORNO")

    ' --- MAPEAMENTO DE COLUNAS ---
    Coluna_Filtro = db.Rows(2).Find(What:="Ir Menu", LookIn:=xlValues, LookAt:=xlWhole).Column
    Coluna_Id_Destino = db.Rows(2).Find(What:="IdSistema", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    ' Buscar a coluna de ID na planilha externa (linha 14 conforme original)
    On Error Resume Next
    Coluna_Id_Origem = sh_externa.Rows(14).Find(What:="IdSistema", LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo 0
    
    If Coluna_Id_Origem = 0 Then
        MsgBox "Não foi possível encontrar a coluna de identificação ou a aba de retorno no arquivo externo.", vbCritical
        plan_externa.Close SaveChanges:=False
        Exit Sub
    End If

    ' --- LOOP DE RETORNO DE INFORMAÇÕES ---
    last_row = db.Cells(Rows.Count, Coluna_Filtro).End(xlUp).Row
    
    ' Indexador da planilha externa
    idx_externo = 3
    
    For j = 3 To last_row
        ' Verifica se a linha deve receber o retorno baseado no filtro
        If db.Cells(j, Coluna_Filtro) <> "" Then
            
            ' Transfere o valor da planilha externa para o banco de dados principal
            db.Cells(j, Coluna_Id_Destino).Value = sh_externa.Cells(idx_externo + 12, Coluna_Id_Origem).Value
            idx_externo = idx_externo + 1
            
        End If
    Next j
    
    plan_externa.Close SaveChanges:=False

    ' Finalização e Bloqueio
    Call bBloqueio

    With Controle_Macro
        last_row_log = .Cells(Rows.Count, "B").End(xlUp).Row + 1
        .Range("A" & last_row_log).Value = "Retorno Dados"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Finalizada"
    End With

    MsgBox "Processo de retorno concluído!", vbInformation

    Application.ScreenUpdating = True

End Sub
