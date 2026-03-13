Attribute VB_Name = "bPowerPoint"
Sub GerarApresentacaoDados()

    Application.ScreenUpdating = False

    ' --- DECLARAÇÕES ---
    Dim db As Worksheet, ap As Worksheet, mn As Worksheet
    Dim Controle_Macro As Worksheet, Controle_Aba As Worksheet
    Dim pptApp As Object, pptPresentation As Object, pptSlide As Object
    Dim pptTextbox As Object, pasta As Object, Arquivo As Object
    Dim i As Long, last_row As Long, TextboxCount As Integer, Coluna As Integer
    Dim SlideWidth As Single, SlideHeight As Single, TextboxWidth As Single, TextboxHeight As Single
    Dim LeftPos As Single, TopPos As Single
    Dim j As Long, StartRow As Long
    Dim nomeProcurado As String, imgPath_jpg As String, imgPath_png As String
    Dim Celula As Range, headersRow As Range, r As Range
    Dim Colunas(1 To 10) As Long, nomevar(1 To 10) As String
    Dim Coluna_1 As Long, Coluna_2 As Long, Coluna_3 As Long, Coluna_4 As Long, Coluna_5 As Long
    Dim Coluna_6 As Long, Coluna_7 As Long, Coluna_8 As Long, Coluna_9 As Long, Coluna_10 As Long
    
    ' --- ATRIBUIÇÕES ---
    Set db = ThisWorkbook.Sheets("DADOS_PRINCIPAIS")
    Set ap = ThisWorkbook.Sheets("Apoio")
    Set mn = ThisWorkbook.Sheets("Menu")
    Set Controle_Macro = ThisWorkbook.Sheets("Controle-Macro")
    Set Controle_Aba = ThisWorkbook.Sheets("Config-Abas")
    
    ' Confirmação de execução
    resposta = MsgBox("Deseja iniciar a geração do RELATÓRIO PPT?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de Sistema")
    
    If resposta <> vbYes Then Exit Sub

    ' Registro de Log
    usuario = Environ("Username")
    hoje = Date
    horaAtual = Format(Time, "hh:mm:ss")
    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Relatório PPT"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Iniciada"
    End With

    Call Validacoes("")
    Call bDesbloqueio
    
    ' --- MAPEAMENTO DE COLUNAS ---
    ' Nomes anonimizados para busca de cabeçalho
    nomesColunas = Array("Categoria", "Agrupamento", "Ir Menu", "Unidade", "Dimensao", "Padrao", "Classificacao", "Tipo Entrada", "CodigoRef", "Prazo", "Periodo")
    
    For Each nome In nomesColunas
        nomeProcurado = nome
        For Each Celula In db.Rows(2).Cells
            If Celula.Value = nomeProcurado Then
                Select Case nomeProcurado
                    Case "Categoria": Coluna_Classe = Celula.Column
                    Case "Agrupamento": Coluna_Grupo = Celula.Column
                    Case "Ir Menu": Coluna_IrMenu = Celula.Column
                    Case "Unidade": Coluna_Setor = Celula.Column
                    Case "Dimensao": Coluna_Tamanho = Celula.Column
                    Case "Padrao": Coluna_Grade = Celula.Column
                    Case "Classificacao": Coluna_Cluster = Celula.Column
                    Case "Tipo Entrada": Coluna_TipoPedido = Celula.Column
                    Case "CodigoRef": Coluna_Referencia = Celula.Column
                    Case "Prazo": Coluna_DataEntrega = Celula.Column
                    Case "Periodo": Coluna_MesRecebimento = Celula.Column
                End Select
                Exit For
            End If
        Next Celula
    Next nome
    
    Set headersRow = db.Rows(2)
    
    ' Loop para variáveis dinâmicas (AD2:AD11 na aba Apoio)
    For i = 1 To 10
        nomeProcurado = ap.Range("AD" & i + 1).Value
        Set r = headersRow.Find(What:=nomeProcurado, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then Colunas(i) = r.Column
        
        ' Remove prefixo para campos de hierarquia principal
        If nomeProcurado = "Unidade" Or nomeProcurado = "Agrupamento" Or nomeProcurado = "Categoria" Then
            nomevar(i) = ""
        Else
            nomevar(i) = nomeProcurado & ": "
        End If
    Next i
    
    ' Atribuição das colunas mapeadas
    Coluna_1 = Colunas(1): Coluna_2 = Colunas(2): Coluna_3 = Colunas(3): Coluna_4 = Colunas(4): Coluna_5 = Colunas(5)
    Coluna_6 = Colunas(6): Coluna_7 = Colunas(7): Coluna_8 = Colunas(8): Coluna_9 = Colunas(9): Coluna_10 = Colunas(10)

    ' --- INTEGRAÇÃO POWERPOINT ---
    caminho_pasta = ThisWorkbook.Path & "\"
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations.Open(caminho_pasta & "modelo_apresentacao.pptx")
    
    SlideWidth = pptPresentation.PageSetup.SlideWidth
    SlideHeight = pptPresentation.PageSetup.SlideHeight
    TextboxWidth = 165
    TextboxHeight = 200
    
    ' Configuração do Slide de Capa
    Set pptSlideOne = pptPresentation.Slides(1)
    Set pptTextbox = pptSlideOne.Shapes.AddTextbox(1, 220, 250, 700, 700)
    pptTextbox.TextFrame.TextRange.Text = UCase("RELATÓRIO DE DADOS - " & db.Cells(3, Coluna_Setor).Value)
    pptTextbox.TextFrame.TextRange.Font.Size = 32
    pptTextbox.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    
    Set pptTextbox = pptSlideOne.Shapes.AddTextbox(1, 30, 500, 600, 50)
    pptTextbox.TextFrame.TextRange.Text = UCase(usuario & " - Unidade: " & db.Cells(3, Coluna_Setor).Value)
    pptTextbox.TextFrame.TextRange.Font.Size = 12
    pptTextbox.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    
    lastSlideIndex = 2
    Set pptSlide = pptPresentation.Slides.Add(lastSlideIndex, 12) ' 12 = ppLayoutBlank
    
    ' --- LOOP DE DADOS PARA SLIDES ---
    For A = 2 To ap.Cells(Rows.Count, "AE").End(xlUp).Row
        mes_rec = ap.Cells(A, "AE").Value
        
        Set pptTextbox = pptSlide.Shapes.AddTextbox(1, 380, 250, 800, 20)
        pptTextbox.TextFrame.TextRange.Text = UCase(mes_rec)
        pptTextbox.TextFrame.TextRange.Font.Size = 36
        pptTextbox.TextFrame.TextRange.Font.Bold = True
        
        lastSlideIndex = lastSlideIndex + 1
        Set pptSlide = pptPresentation.Slides.Add(lastSlideIndex, 12)

        For b = 2 To ap.Cells(Rows.Count, "H").End(xlUp).Row
            cluster = ap.Cells(b, "H").Value
            For c = 2 To ap.Cells(Rows.Count, "O").End(xlUp).Row
                Agrup = ap.Cells(c, "O").Value
                For d = 2 To ap.Cells(Rows.Count, "P").End(xlUp).Row
                    Categ = ap.Cells(d, "P").Value

                    For i = 3 To db.Cells(Rows.Count, "B").End(xlUp).Row
                        ' Lógica de Data/Mês
                        If db.Cells(i, Coluna_DataEntrega).Value <> "" Then
                            mes_num = Month(db.Cells(i, Coluna_DataEntrega).Value)
                        Else
                            On Error Resume Next
                            mes_num = WorksheetFunction.Match(db.Cells(i, Coluna_DataEntrega).Value, Array("Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"), 0)
                            On Error GoTo 0
                        End If
                        
                        nome_mes = StrConv(Format(DateSerial(Year(Now), mes_num, 1), "mmmm"), vbProperCase)

                        ' Filtros de validação para inclusão no slide
                        If UCase(db.Cells(i, Coluna_IrMenu).Value) <> "" And _
                           UCase(nome_mes) = UCase(mes_rec) And _
                           UCase(db.Cells(i, Coluna_Cluster).Value) = UCase(cluster) And _
                           UCase(db.Cells(i, Coluna_Grupo).Value) = UCase(Agrup) And _
                           UCase(db.Cells(i, Coluna_Classe).Value) = UCase(Categ) Then

                            If TextboxCount = 0 Or TextboxCount = 1 Then
                                Set pptTextbox = pptSlide.Shapes.AddTextbox(1, 80, 10, 800, 15)
                                pptTextbox.TextFrame.TextRange.Text = UCase(Agrup & " - " & nome_mes & " - Ref: " & cluster)
                                pptTextbox.TextFrame.TextRange.Font.Size = 26
                                pptTextbox.TextFrame.TextRange.Font.Bold = True
                            End If

                            If LeftPos + TextboxWidth <= SlideWidth And TextboxCount < 6 Then
                                ' Inserção de texto
                                Set pptTextbox = pptSlide.Shapes.AddTextbox(1, LeftPos, 60, TextboxWidth, TextboxHeight)
                                pptTextbox.TextFrame.TextRange.Text = _
                                    nomevar(1) & db.Cells(i, Coluna_1).Value & vbCrLf & _
                                    nomevar(2) & db.Cells(i, Coluna_2).Value & vbCrLf & _
                                    nomevar(3) & db.Cells(i, Coluna_3).Value & vbCrLf & _
                                    nomevar(4) & db.Cells(i, Coluna_4).Value & vbCrLf & _
                                    nomevar(5) & db.Cells(i, Coluna_5).Value & vbCrLf & _
                                    nomevar(6) & db.Cells(i, Coluna_6).Value & vbCrLf & _
                                    nomevar(7) & db.Cells(i, Coluna_7).Value & vbCrLf & _
                                    nomevar(8) & db.Cells(i, Coluna_8).Value & vbCrLf & _
                                    nomevar(9) & db.Cells(i, Coluna_9).Value & vbCrLf & _
                                    nomevar(10) & db.Cells(i, Coluna_10).Value
                                
                                ' Lógica de Imagem
                                imgPath_jpg = caminho_pasta & "Arquivos_Midia\" & db.Cells(i, Coluna_Referencia).Value & ".jpg"
                                imgPath_png = caminho_pasta & "Arquivos_Midia\" & db.Cells(i, Coluna_Referencia).Value & ".png"

                                If Dir(imgPath_jpg) <> "" Then
                                    pptSlide.Shapes.AddPicture imgPath_jpg, 0, 1, LeftPos, 250, 165, 200
                                ElseIf Dir(imgPath_png) <> "" Then
                                    pptSlide.Shapes.AddPicture imgPath_png, 0, 1, LeftPos, 250, 165, 200
                                Else
                                    Set pptNoImg = pptSlide.Shapes.AddTextbox(1, LeftPos, 300, 165, 200)
                                    pptNoImg.TextFrame.TextRange.Text = "Mídia Indisponível"
                                End If

                                TextboxCount = TextboxCount + 1
                                LeftPos = LeftPos + TextboxWidth + 20
                            Else
                                ' Novo slide por falta de espaço
                                lastSlideIndex = lastSlideIndex + 1
                                Set pptSlide = pptPresentation.Slides.Add(lastSlideIndex, 12)
                                LeftPos = 30
                                TextboxCount = 1

                            End If
                        End If
                    Next i
                Next d
                
                ' Limpeza de slides vazios ao final dos loops internos
                If pptPresentation.Slides(lastSlideIndex).Shapes.Count = 0 Then
                    LeftPos = 30
                    TextboxCount = 1
                Else
                    lastSlideIndex = lastSlideIndex + 1
                    Set pptSlide = pptPresentation.Slides.Add(lastSlideIndex, 12)
                    LeftPos = 30
                    TextboxCount = 1
                End If
            Next c
        Next b
    Next A

    ' --- FINALIZAÇÃO ---
    Set pptSlide = Nothing: Set pptPresentation = Nothing: Set pptApp = Nothing
    Call bBloqueio

    last_row_log = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
    With Controle_Macro
        .Range("A" & last_row_log).Value = "Relatório PPT"
        .Range("B" & last_row_log).Value = hoje
        .Range("C" & last_row_log).Value = horaAtual
        .Range("D" & last_row_log).Value = usuario
        .Range("E" & last_row_log).Value = "Finalizada"
    End With

    MsgBox "Apresentação gerada com sucesso!", vbInformation
    Application.ScreenUpdating = True

End Sub

