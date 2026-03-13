Attribute VB_Name = "Fornecedor_Email"
Sub EnviarEmail()

    Application.ScreenUpdating = False
        
        ' Declarações
        Dim db As Worksheet, ap As Worksheet, mn As Worksheet
        Dim Controle_Macro As Worksheet, Controle_Erro As Worksheet
        Dim Coluna, horario, last_row, tam_cont_carac, grade_cont_carac, soma_grade As Integer
        Dim i, j, coluna_cadastro As Long
        Dim rng As Range, rng_y As Range
        Dim caminho_pasta, nomeProcurado As String
        Dim dtToday As Date
        Dim pasta As Object
        Dim Celula As Range, opcoes As Range
        
        ' Atribuições (Nomes de abas anonimizados)
        Set db = ThisWorkbook.Sheets("BASE_DADOS")
        Set ap = ThisWorkbook.Sheets("Apoio_Tecnico")
        Set mn = ThisWorkbook.Sheets("Menu_Principal")
        Set Controle_Macro = ThisWorkbook.Sheets("LOG_SISTEMA")
        Set Controle_Erro = ThisWorkbook.Sheets("LOG_ERRO")
        
        ' Validação de execução
        resposta = MsgBox("Deseja executar o envio de E-MAIL PARA PARCEIROS?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação")
        
        If resposta <> vbYes Then
            Exit Sub
        End If

        ' Auditoria
        usuario = Environ("Username")
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")
        
        ' Registro de Log Inicial
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Envio de Email Informativo"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Iniciada"
        End With

        Call Validacoes("")
        Call bDesbloqueio

        caminho_pasta = ThisWorkbook.Path & "\"
        Set pasta = CreateObject("Scripting.filesystemobject").GetFolder(caminho_pasta)
        
        '<<< MAPEAMENTO DE COLUNAS >>>
        ' Cabeçalhos generalizados
        nomesColunas = Array("Tamanho", "Grade", "Agrupamento", ap.Range("AJ1").Value, "Quantidade", "Email", "Ref_Interna", "Tipo_Entrada", "Pack_Size", "Descritivo_A", "Descritivo_B", "Categoria_ID")
        
        For Each nome In nomesColunas
            For Each Celula In db.Rows(2).Cells
                If Celula.Value = nome Then
                    Select Case nome
                        Case "Tamanho": Coluna_Tamanho = Celula.Column
                        Case "Grade": Coluna_Grade = Celula.Column
                        Case "Agrupamento": Coluna_Cluster = Celula.Column
                        Case ap.Range("AJ1").Value: Coluna_Envio = Celula.Column
                        Case "Quantidade": Coluna_Qnt_Comprada = Celula.Column
                        Case "Email": Coluna_Email = Celula.Column
                        Case "Ref_Interna": Coluna_PE = Celula.Column
                        Case "Tipo_Entrada": Coluna_TipoPedido = Celula.Column
                        Case "Pack_Size": Coluna_Tamanho_do_Pack = Celula.Column
                        Case "Descritivo_A": Coluna_TextoE_commerce = Celula.Column
                        Case "Descritivo_B": Coluna_Texto_Ecommerce_Variantes = Celula.Column
                        Case "Categoria_ID": Coluna_Categoria_EAN = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next nome
        
        ' Processamento de filtragem de destinatários
        last_row = db.Range("B1").End(xlDown).Row
        ap.Range("AI2:AM" & ap.Cells(ap.Rows.Count, 1).End(xlUp).Row).Clear
        
        With db
            .Range("$A$2:$DW$" & last_row).AutoFilter Field:=3, Criteria1:="<>"
            .Range(.Cells(2, Coluna_Email), .Cells(last_row, Coluna_Email)).SpecialCells(xlCellTypeVisible).Copy
            ap.Range("AI1").PasteSpecial Paste:=xlPasteValues
            .Range(.Cells(2, Coluna_Envio), .Cells(last_row, Coluna_Envio)).SpecialCells(xlCellTypeVisible).Copy
            ap.Range("AJ1").PasteSpecial Paste:=xlPasteValues
        End With
            
        last_row_ap = ap.Cells(ap.Rows.Count, "AI").End(xlUp).Row
        ap.Range("AI1:AJ" & last_row_ap).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
        db.AutoFilterMode = False

        '<<< LOOP DE ENVIO >>>
        last_envio = ap.Cells(ap.Rows.Count, "AI").End(xlUp).Row
        
        For q = 2 To last_envio
            FornecedorMarca = ap.Range("AJ" & q).Value
            contcolumns = Application.CountA(db.Range("A2:XFD2"))
            
            ' Abertura do template (Nome do arquivo mantido como padrao)
            Set planilha = Workbooks.Open(caminho_pasta & "template_externo.xlsx")
            Set rng = planilha.Sheets(1).Range("A2:W2")
            
            planilha.Sheets(1).Rows("3:1000").Clear ' Limpeza preventiva
            
            For j = 3 To db.Cells(Rows.Count, "B").End(xlUp).Row
                ' Lógica de contagem de variantes por delimitador ";"
                If db.Cells(j, Coluna_TipoPedido).Value = "Pack" Then
                    cont_carac = 1
                Else
                    cont_carac = Len(db.Cells(j, Coluna_Tamanho)) - Len(Replace(db.Cells(j, Coluna_Tamanho), ";", "")) + 1
                End If
                
                For t = 1 To cont_carac
                    If db.Cells(j, 3) <> "" And db.Cells(j, Coluna_Envio) = FornecedorMarca Then
                        k = planilha.Sheets(1).Range("A1048576").End(xlUp).Row + 1
                        
                        For i = 1 To contcolumns
                            valor_cabecalho = db.Cells(2, i).Value
                            For Each cell In rng
                                If LCase(cell.Value) = LCase(valor_cabecalho) Then
                                    f = cell.Column
                                    If (i = Coluna_Tamanho Or i = Coluna_Grade Or i = Coluna_PE) And cont_carac > 1 Then
                                        On Error Resume Next
                                        planilha.Sheets(1).Cells(k, f).Value = Split(db.Cells(j, i).Value, ";")(t - 1)
                                        On Error GoTo 0
                                    Else
                                        planilha.Sheets(1).Cells(k, f).Value = db.Cells(j, i).Value
                                    End If
                                    
                                    ' Cálculo de distribuição de quantidades
                                    qnt_comprada = db.Cells(j, Coluna_Qnt_Comprada).Value
                                    If cont_carac = 1 Then
                                        planilha.Sheets(1).Cells(k, 12) = qnt_comprada
                                    Else
                                        tam_total = db.Cells(j, Coluna_Tamanho_do_Pack).Value
                                        planilha.Sheets(1).Cells(k, 12) = (qnt_comprada / tam_total) * planilha.Sheets(1).Cells(k, 13)
                                    End If
                                End If
                            Next cell
                        Next i
                    End If
                Next t
            Next j
            
            ' Configuração de validação de dados no arquivo gerado
            Set opcoes = planilha.Sheets(2).Range("A2:A10")
            ultima_linha_fornecedor = planilha.Sheets(1).Cells(planilha.Sheets(1).Rows.Count, 1).End(xlUp).Row
            
            With planilha.Sheets(1).Columns("W").Validation ' Exemplo de coluna fixada para origem
                .Delete
                .Add Type:=xlValidateList, Formula1:="=Planilha2!" & opcoes.Address
            End With
            
            ' Formatação Visual (Campos obrigatórios em amarelo)
            Set rng_y = planilha.Sheets(1).Range("A3:W" & ultima_linha_fornecedor)
            On Error Resume Next
            rng_y.SpecialCells(xlCellTypeBlanks).Interior.Color = 65535
            On Error GoTo 0
            
            planilha.Save
            planilha.Close
            
            '<<< INTERFACE COM OUTLOOK >>>
            enviarpara = ap.Range("AI" & q)
            If enviarpara <> "" Then
                horario = Hour(Now)
                tratamento = IIf(horario > 11, "Boa tarde,", "Bom dia,")
                
                Mensagem = "Prezado(a) Parceiro(a),<br/><br/>" _
                         & tratamento & "<br/><br/>" _
                         & "Encaminhamos em anexo o formulário para preenchimento dos dados técnicos.<br/>" _
                         & "<b>Solicitamos o preenchimento integral dos campos destacados em amarelo.</b><br/><br/>" _
                         & "Atenciosamente,<br/>Equipe de Gestão"

                Set OutApp = CreateObject("Outlook.Application")
                Set OutMail = OutApp.CreateItem(0)
                
                With OutMail
                    .To = enviarpara
                    .Subject = "Solicitação de Dados Técnicos - " & FornecedorMarca & " - " & Date
                    .HTMLBody = Mensagem
                    .Attachments.Add caminho_pasta & "template_externo.xlsx"
                    .Display ' Alterar para .Send para envio automático
                End With
            End If
        Next q
    
        Call bBloqueio

        ' Registro de Log Final
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Envio de Email Informativo"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Finalizada"
        End With

    MsgBox "Processo de comunicação concluído!"
    Application.ScreenUpdating = True

End Sub

