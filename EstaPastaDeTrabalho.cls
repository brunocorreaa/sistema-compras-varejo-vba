VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Variável global para controle de ociosidade
Dim UltimaAcao As Date

' =========================================================================
' EVENTO: Workbook_BeforeClose
' DESCRIÇÃO: Registra o log de fechamento e o tempo total de uso do arquivo.
' =========================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Application.ScreenUpdating = False
    
        ' Declarações
        Dim wsLog, wsMenu As Worksheet
        Dim last_row As Integer
        Dim hoje As Date
        Dim horaAtual, usuario As String
        
        ' Atribuições (Nomes de abas anonimizados)
        Set wsLog = ThisWorkbook.Sheets("Controle_Log")
        Set wsMenu = ThisWorkbook.Sheets("Painel_Principal")
        
        ' Executa rotina de segurança/manutenção antes de fechar
        Call Rotina_Desbloqueio_Seguro
        
        ' Captura metadados da sessão
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")
        usuario = Environ("Username")
        
        ' Localiza última linha para registro
        last_row = wsLog.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        ' Verifica status do último log para calcular tempo de permanência
        If wsLog.Range("A" & last_row - 1) = "Status_Abertura" Then
            With wsLog
                .Range("A" & last_row).Value = "Status_Fechamento"
                .Range("B" & last_row).Value = hoje
                .Range("C" & last_row).Value = horaAtual
                .Range("D" & last_row).Value = usuario
                ' Cálculo de tempo de sessão em minutos
                .Range("E" & last_row).Value = Round(((hoje + TimeValue(horaAtual)) - (.Range("B" & last_row - 1).Value + .Range("C" & last_row - 1).Value)) * 24 * 60, 2)
            End With
        Else
            With wsLog
                .Range("A" & last_row).Value = "Status_Fechamento"
                .Range("B" & last_row).Value = hoje
                .Range("C" & last_row).Value = horaAtual
                .Range("D" & last_row).Value = usuario
                .Range("E" & last_row).Value = 0
            End With
        End If
                        
        ' Lógica de fechamento rápido: se salvo recentemente, fecha sem alertas
        If (Now - wsMenu.Cells(2, 2).Value) * 86400 < 10 Then
            Application.DisplayAlerts = False
            ThisWorkbook.Close SaveChanges:=False
        End If

    Application.ScreenUpdating = True

End Sub

' =========================================================================
' EVENTO: Workbook_BeforeSave
' DESCRIÇÃO: Registra o timestamp do último salvamento no Menu.
' =========================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    Sheets("Painel_Principal").Cells(2, 2).Value = Now
End Sub

' =========================================================================
' EVENTO: Workbook_Open
' DESCRIÇÃO: Rotinas de auditoria, lembretes de manutenção e verificação de integridade.
' =========================================================================
Private Sub Workbook_Open()
    
    Application.ScreenUpdating = False
    
        ' Declarações de objetos e variáveis de controle
        Dim wsLog_Arq As Worksheet, wsLog_Aba As Worksheet, wsApoio As Worksheet, wsMenu As Worksheet
        Dim TamanhoLimite_KB, TamanhoAtual_KB As Long
        Dim CaminhoArquivo, horaAtual, usuario As String
        Dim last_row As Integer
        Dim diaReferencia As Integer
        
        diaReferencia = Day(Date)
        
        ' Atribuições (Nomes genéricos para proteção de dados)
        Set wsLog_Arq = ThisWorkbook.Sheets("Controle_Log")
        Set wsLog_Aba = ThisWorkbook.Sheets("Log_Navegacao")
        Set wsApoio = ThisWorkbook.Sheets("Configuracoes")
        Set wsMenu = ThisWorkbook.Sheets("Painel_Principal")
        
        'Call Rotina_Desbloqueio_Seguro
        
        ' Desabilita triggers de execução interna até o fim da abertura
        With wsApoio.Range("AO2")
            .Value = 0
            .Font.ThemeColor = xlThemeColorDark1
        End With
        
        ' -----------------------------------------------------------------
        ' LEMBRETES DE MANUTENÇÃO PERIÓDICA (Anonimizado)
        ' -----------------------------------------------------------------
        ' Datas de lembrete ajustadas para lógica genérica
        If diaReferencia = 1 Or diaReferencia = 15 Then
            MsgBox "Aviso de Revisão Periódica:" & vbCrLf & vbCrLf & _
                   "Este é um lembrete automático para validar os parâmetros " & _
                   "de configuração na aba 'Configuracoes'." & vbCrLf & vbCrLf & _
                   "Certifique-se de que os clusters de dados estão atualizados.", _
                   vbInformation, "Governança de Dados"
        End If
          
        ' -----------------------------------------------------------------
        ' MONITORAMENTO DE PERFORMANCE (Tamanho do Arquivo)
        ' -----------------------------------------------------------------
        TamanhoLimite_KB = 40000 ' Limite definido pelo desenvolvedor (40MB)
        CaminhoArquivo = ActiveWorkbook.FullName
        TamanhoAtual_KB = FileLen(CaminhoArquivo) / 1024
        
        If TamanhoAtual_KB > TamanhoLimite_KB Then
            MsgBox "ALERTA DE PERFORMANCE:" & vbCrLf & _
                   "O arquivo excedeu o tamanho recomendado (" & TamanhoLimite_KB & " KB)." & vbCrLf & _
                   "Dicas para otimização:" & vbCrLf & _
                   "1. Limpar formatos em excesso." & vbCrLf & _
                   "2. Reduzir imagens ou objetos desnecessários.", vbExclamation
        End If
                                
        ' -----------------------------------------------------------------
        ' REGISTRO DE AUDITORIA DE ACESSO
        ' -----------------------------------------------------------------
        horaAtual = Format(Time, "hh:mm:ss")
        usuario = Environ("Username")
        last_row = wsLog_Arq.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        ' Verifica se houve fechamento irregular anterior
        If wsLog_Arq.Range("A" & last_row - 1).Value = "Status_Abertura" Then
            With wsLog_Arq
                .Range("A" & last_row).Value = "Status_Fechamento"
                .Range("B" & last_row).Value = Int(wsMenu.Cells(2, 2).Value)
                .Range("C" & last_row).Value = Format(wsMenu.Cells(2, 2).Value, "hh:mm:ss")
                .Range("D" & last_row).Value = .Range("D" & last_row - 1).Value
                .Range("E" & last_row).Value = Round(((.Range("B" & last_row).Value + .Range("C" & last_row).Value) - (.Range("B" & last_row - 1).Value + .Range("C" & last_row - 1).Value)) * 24 * 60, 2)
            End With
        End If
        
        ' Registro da nova abertura
        last_row = wsLog_Arq.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With wsLog_Arq
            .Range("A" & last_row).Value = "Status_Abertura"
            .Range("B" & last_row).Value = Date
            .Range("C" & last_row).Value = horaAtual
            .Range("D" & last_row).Value = usuario
            .Range("E" & last_row).Value = 0
        End With
        
        ' Registro inicial de navegação
        last_row = wsLog_Aba.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With wsLog_Aba
            .Range("A" & last_row).Value = wsMenu.Name
            .Range("B" & last_row).Value = Date
            .Range("C" & last_row).Value = horaAtual
            .Range("D" & last_row).Value = usuario
        End With
        
        wsMenu.Visible = xlSheetVisible
        
        ' Chamadas de subsistemas (Nomes anonimizados)
        Call Executar_Validacoes_Estruturais
        Call Rotina_Bloqueio_Seguro
        Call Sincronizar_Dados_Externos
                                
    Application.ScreenUpdating = True
    
End Sub

' =========================================================================
' MONITORES DE ATIVIDADE DO USUÁRIO
' DESCRIÇÃO: Captura mudanças de aba e cliques para log de utilização.
' =========================================================================
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Call RegistrarAcaoUsuario(Sh.Name)
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call RegistrarAcaoUsuario(Sh.Name)
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Call RegistrarAcaoUsuario(Sh.Name)
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Call RegistrarAcaoUsuario(Sh.Name)
End Sub

' =========================================================================
' SUBROTINA: RegistrarAcaoUsuario
' DESCRIÇÃO: Calcula tempo de inatividade e registra o fluxo do usuário.
' =========================================================================
Private Sub RegistrarAcaoUsuario(ByVal nomeAba As String)
    
    Dim wsLog_Aba As Worksheet
    Dim tempoInativo, totalSegundos As Double
    Dim last_row As Long
    Dim horaAtual, usuario As String
    Dim segAtual, segUltimo As Long
    
    Set wsLog_Aba = ThisWorkbook.Sheets("Log_Navegacao")
    
    horaAtual = Format(Time, "hh:mm:ss")
    usuario = Environ("Username")
    last_row = wsLog_Aba.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    If UltimaAcao <> 0 Then
        totalSegundos = (Now - UltimaAcao) * 86400
        segAtual = TimeValue(Format(Time, "hh:mm:ss")) * 86400
        
        ' Recupera timestamp do último registro para evitar duplicidade rápida
        If wsLog_Aba.Range("C" & last_row - 1).Value <> "" Then
            segUltimo = TimeValue(Format(wsLog_Aba.Range("C" & last_row - 1).Value, "hh:mm:ss")) * 86400
        Else
            segUltimo = 0
        End If
        
        ' Lógica: Registra se inativo por > 10 min ou se mudou de aba
        If Int(totalSegundos) > 600 Then
            With wsLog_Aba
                .Range("A" & last_row).Value = nomeAba
                .Range("B" & last_row).Value = Date
                .Range("C" & last_row).Value = horaAtual
                .Range("D" & last_row).Value = usuario
                .Range("E" & last_row).Value = Int(totalSegundos)
            End With
        ElseIf segAtual - segUltimo > 1 And wsLog_Aba.Range("A" & last_row - 1).Value <> nomeAba Then
            With wsLog_Aba
                .Range("A" & last_row).Value = nomeAba
                .Range("B" & last_row).Value = Date
                .Range("C" & last_row).Value = horaAtual
                .Range("D" & last_row).Value = usuario
                .Range("E" & last_row).Value = 0
            End With
        End If
    End If
    
    UltimaAcao = Now
    
End Sub

