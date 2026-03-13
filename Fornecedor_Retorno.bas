Attribute VB_Name = "Fornecedor_Retorno"
Sub RetornoFornecedor()

    Application.ScreenUpdating = False

        ' Declarações
        Dim planilha As Workbook
        Dim db As Worksheet, mn As Worksheet, Controle_Macro As Worksheet, Controle_Erro As Worksheet, db_f As Worksheet
        Dim i As Long, j As Long, p As Long, Z As Long
        Dim forn_last_row As Long, last_row As Long, last_row_macro As Long
        Dim cont_carac As Integer, Coluna As Integer, f As Integer
        Dim caminho_pasta As String, nomeProcurado As String, usuario As String, ID As String
        Dim dict As Object
        Dim Celula As Range, cell As Range
        
        ' Atribuições (Nomes de abas anonimizados)
        Set db = ThisWorkbook.Sheets("BASE_DADOS")
        Set mn = ThisWorkbook.Sheets("Menu_Principal")
        Set Controle_Macro = ThisWorkbook.Sheets("LOG_SISTEMA")
        Set Controle_Erro = ThisWorkbook.Sheets("LOG_ERRO")
 
        ' Validação de execução
        If MsgBox("Deseja importar os dados de RETORNO DO PARCEIRO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") <> vbYes Then
            Exit Sub
        End If
 
        ' Auditoria e Log Inicial
        usuario = Environ("Username")
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Processamento Retorno Parceiro"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Iniciada"
        End With

        Call Validacoes("")
        Call bDesbloqueio
                
        '<<< MAPEAMENTO DE COLUNAS >>>
        nomesColunas = Array("ID_Ref", "Flag_Acao", "Tamanho", "Grade", "Agrupamento", "Tipo_Entrada", "EAN_Variante", "Categoria_ID")
        
        For Each nome In nomesColunas
            For Each Celula In db.Rows(2).Cells
                If Celula.Value = nome Then
                    Select Case nome
                        Case "ID_Ref": Coluna_N = Celula.Column
                        Case "Flag_Acao": Coluna_IrMenu = Celula.Column
                        Case "Tamanho": Coluna_Tamanho = Celula.Column
                        Case "Grade": Coluna_Grade = Celula.Column
                        Case "Agrupamento": Coluna_Cluster = Celula.Column
                        Case "Tipo_Entrada": Coluna_TipoPedido = Celula.Column
                        Case "EAN_Variante": Coluna_Cod_EAN_Var = Celula.Column
                        Case "Categoria_ID": Coluna_Categoria_EAN = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next nome
        
        '<<< PROCESSAMENTO DOS ARQUIVOS >>>
        caminho_pasta = ThisWorkbook.Path & "\"
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' Abertura do arquivo de retorno (Nome genérico)
        On Error Resume Next
        Set planilha = Workbooks.Open(caminho_pasta & "template_externo.xlsx")
        If Err.Number <> 0 Then
            MsgBox "Arquivo 'template_externo.xlsx' não encontrado!", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        
        Set db_f = planilha.Sheets(1)
        last_row = db.Cells(Rows.Count, "B").End(xlUp).Row
        forn_last_row = db_f.Cells(Rows.Count, "A").End(xlUp).Row
        contcolumns = Application.CountA(db.Range("A2:XFD2"))
        
        For j = 3 To last_row
            ' Calcula variantes baseadas no delimitador ";"
            cont_carac = Len(db.Cells(j, Coluna_Tamanho)) - Len(Replace(db.Cells(j, Coluna_Tamanho), ";", "")) + 1
            irMenu = db.Cells(j, Coluna_IrMenu).Value
                
            If irMenu <> "" Then
                For i = 1 To contcolumns
                    valor_cabecalho = db.Cells(2, i).Value
                    
                    For Each cell In db_f.Range("A2:EQ2")
                        If cell.Value = valor_cabecalho Then
                            For Z = 3 To forn_last_row
                                If db.Cells(j, 2).Value = db_f.Cells(Z, 1).Value Then
                                    ' Ignora colunas estruturais que não devem ser sobrescritas
                                    If cell.Value <> "Tamanho" And cell.Value <> "Grade" And cell.Value <> "Ref_Interna" Then
                                        f = cell.Column
                                        
                                        ' Lógica específica para concatenação de EANs
                                        If cell.Value = "EAN_Variante" And cont_carac > 1 Then
                                            dict.RemoveAll
                                            For p = 3 To forn_last_row
                                                ID = db_f.Cells(p, 1).Value
                                                codigoEAN = db_f.Cells(p, 19).Value ' Coluna fixa do EAN no retorno
                                                
                                                If db.Cells(j, 2).Value = ID Then
                                                    If Not dict.Exists(ID) Then
                                                        dict(ID) = codigoEAN
                                                    Else
                                                        dict(ID) = dict(ID) & ";" & codigoEAN
                                                    End If
                                                End If
                                            Next p
                                            db.Cells(j, i).Value = dict(db.Cells(j, 2).Value)
                                        Else
                                            db.Cells(j, i).Value = db_f.Cells(Z, f).Value
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next Z
                        End If
                    Next cell
                Next i
            End If
        Next j
        
        planilha.Close SaveChanges:=False
    
        '<<< TRATAMENTO PARA ITENS TIPO "PACK" >>>
        For i = 3 To last_row
            If db.Cells(i, Coluna_TipoPedido).Value = "Pack" And db.Cells(i, Coluna_IrMenu).Value <> "" Then
                texto_grade = db.Cells(i, Coluna_Grade).Value
                qtdeVariantes = Len(texto_grade) - Len(Replace(texto_grade, ";", "")) + 1
                
                ' Replica o EAN principal para todas as variações da grade do Pack
                ean_base = db.Cells(i, Coluna_Cod_EAN_Var).Value
                strRept = Left(WorksheetFunction.Rept(ean_base & ";", qtdeVariantes), _
                          Len(WorksheetFunction.Rept(ean_base & ";", qtdeVariantes)) - 1)
                
                db.Cells(i, Coluna_Cod_EAN_Var) = strRept
            End If
        Next i
        
        '<<< VALIDAÇÃO DE INTEGRIDADE (Tamanho vs EAN) >>>
        For i = 3 To last_row
            If db.Cells(i, Coluna_IrMenu).Value <> "" Then
                tam_cont = Len(db.Cells(i, Coluna_Tamanho)) - Len(Replace(db.Cells(i, Coluna_Tamanho), ";", "")) + 1
                ean_cont = Len(db.Cells(i, Coluna_Cod_EAN_Var)) - Len(Replace(db.Cells(i, Coluna_Cod_EAN_Var), ";", "")) + 1
                
                If tam_cont <> ean_cont Then
                    ' Registro de Erro
                    err_row = Controle_Erro.Cells(Rows.Count, "B").End(xlUp).Row + 1
                    With Controle_Erro
                        .Range("A" & err_row).Value = "Divergência: Tamanho (" & tam_cont & ") vs EAN (" & ean_cont & ")"
                        .Range("B" & err_row).Value = Date
                        .Range("C" & err_row).Value = Format(Now, "hh:mm:ss")
                        .Range("D" & err_row).Value = usuario
                    End With
                    
                    Call bBloqueio
                    MsgBox "Erro de integridade na linha " & i & ". A quantidade de Tamanhos não condiz com a de EANs.", vbCritical
                    Exit Sub
                End If
            End If
        Next i

        Call bBloqueio

        ' Registro de Log Final
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Processamento Retorno Parceiro"
            .Range("B" & last_row_macro).Value = Date
            .Range("C" & last_row_macro).Value = Format(Now, "hh:mm:ss")
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Finalizada"
        End With

    MsgBox "Importação concluída com sucesso!"
    Application.ScreenUpdating = True

End Sub
