Attribute VB_Name = "CompraMTA"
Sub bProcessarDados()

    Application.ScreenUpdating = False

        ' Declaração de variáveis para as planilhas e objetos
        Dim db As Worksheet, mn As Worksheet, Controle_Macro As Worksheet, wkbImport As Workbook, pasta As Object
        Dim skuDict As Object, lastRowA As Long, lastRowH As Long, i As Long, j As Long, n As Long
        Dim currentSKU As String, currentSKU2 As String, zerosComPontoEVirgula As String
        Dim maiorValor As Double, ult_Linha_Proc As Long, contador As Long, last_row As Long
        Dim last_col As Long, Valor_Linha As Long, Coluna_EnvioRep As Long, Coluna_Target As Long
        Dim Coluna_Distribuicao As Long, Coluna_Tema As Long, Coluna_OrigemDoModelo As Long
        Dim Coluna_PlanoQtd As Long, Coluna_QtdComprada As Long, Coluna_Tamanho_do_Pack As Long
        Dim Coluna_Qtd_Lojas As Long, Coluna_Plano_Apoio As Long, Coluna_Semestre As Long
        Dim Coluna_Tamanho As Long, Coluna_Grade As Long, Coluna_Ano As Long
        Dim nomesColunas As Variant, nome As Variant, Celula As Range
        Dim caminho_pasta As String
        
        ' Atribuindo os objetos às planilhas (Nomes anonimizados)
        Set db = ThisWorkbook.Sheets("BASE_DADOS")
        Set mn = ThisWorkbook.Sheets("Menu")
        Set Controle_Macro = ThisWorkbook.Sheets("LOG_SISTEMA")
        
        ' Confirmar execução
        resposta = MsgBox("Você realmente quer executar o processo de IMPORTAÇÃO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação de uso")
        
        If resposta <> vbYes Then
            Exit Sub
        End If

        ' Informações de auditoria
        usuario = Environ("Username")
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")

        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Processo Importação"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Iniciada"
        End With
        
        ' Chamadas de sub-rotinas
        Call Rotina_Validacoes("")
        Call bDesbloqueio

        last_row = db.Cells(Rows.Count, "B").End(xlUp).Row
        contcolumns = Application.CountA(db.Range("A2:XFD2"))
                
        ' Cabeçalhos de coluna (Mantidos por serem referências lógicas, mas podem ser ajustados se forem proprietários)
        nomesColunas = Array("Ano", "ID_REF", "Tamanho", "Grade", "Status_Entrega", "Target", "Fluxo", "Categoria", "Origem_Entrada", "Plano_Qtd", "Qtd_Efetiva", "Tam_Pack", "Qtd_Locais", "Tipo_Pedido", "Plano_Apoio", "Semestre")
        
        For Each nome In nomesColunas
            For Each Celula In db.Rows(2).Cells
                If Celula.Value = nome Then
                    Select Case nome
                        Case "Ano": Coluna_Ano = Celula.Column
                        Case "Status_Entrega": Coluna_EnvioRep = Celula.Column
                        Case "Target": Coluna_Target = Celula.Column
                        Case "Fluxo": Coluna_Distribuicao = Celula.Column
                        Case "Categoria": Coluna_Tema = Celula.Column
                        Case "Origem_Entrada": Coluna_OrigemDoModelo = Celula.Column
                        Case "Plano_Qtd": Coluna_PlanoQtd = Celula.Column
                        Case "Qtd_Efetiva": Coluna_QtdComprada = Celula.Column
                        Case "Tam_Pack": Coluna_Tamanho_do_Pack = Celula.Column
                        Case "Qtd_Locais": Coluna_Qtd_Lojas = Celula.Column
                        Case "Tipo_Pedido": Coluna_TipoPedido = Celula.Column
                        Case "Plano_Apoio": Coluna_Plano_CML = Celula.Column
                        Case "Semestre": Coluna_Semestre = Celula.Column
                        Case "Tamanho": Coluna_Tamanho = Celula.Column
                        Case "Grade": Coluna_Grade = Celula.Column
                        Case "ID_REF": Coluna_PE = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next nome
        
        ' Abrir o arquivo de suporte (Nome anonimizado)
        caminho_pasta = ThisWorkbook.Path & "\"
        Set pasta = CreateObject("Scripting.filesystemobject").GetFolder(caminho_pasta)
        Set wkbImport = Workbooks.Open(caminho_pasta & "Template_Dados.xlsx")

        ' Validação dos dados importados
        lastRowD = wkbImport.Sheets(1).Cells(wkbImport.Sheets(1).Rows.Count, "D").End(xlUp).Row
        
        For i = 2 To lastRowD
            With wkbImport.Sheets(1)
                Select Case .Cells(i, 5).Value
                    Case Year(Date), Year(Date) - 1, Year(Date) + 1, ""
                    Case Else
                        MsgBox "Erro na linha " & i & ": Ano inválido."
                        Exit Sub
                End Select
                
                Select Case .Cells(i, 6).Value
                    Case 1, 2, ""
                    Case Else
                        MsgBox "Erro na linha " & i & ": Semestre inválido."
                        Exit Sub
                End Select
            End With
        Next i
        
        lastRowA = wkbImport.Sheets(1).Cells(wkbImport.Sheets(1).Rows.Count, "D").End(xlUp).Row

        ' Processamento de duplicatas e estruturação de dados
        wkbImport.Sheets(1).Columns("D:D").Copy
        wkbImport.Sheets(1).Range("M1").PasteSpecial Paste:=xlPasteValues
        wkbImport.Sheets(1).Columns("M:M").RemoveDuplicates Columns:=1, header:=xlYes

        lastRowH = wkbImport.Sheets(1).Cells(wkbImport.Sheets(1).Rows.Count, "M").End(xlUp).Row
        
        ' Lógica de Fórmulas e Dicionários
        wkbImport.Sheets(1).Range("N1").Value = "Contador"
        wkbImport.Sheets(1).Range("N2").FormulaR1C1 = "=COUNTIF(C[-10],RC[-1])"
        If lastRowH > 2 Then wkbImport.Sheets(1).Range("N2").AutoFill Destination:=wkbImport.Sheets(1).Range("N2:N" & lastRowH)
        
        wkbImport.Sheets(1).Range("O1").Value = "PosicaoRef"
        wkbImport.Sheets(1).Range("O2").FormulaR1C1 = "=MATCH(RC[-2],C[-11],0)"
        If lastRowH > 2 Then wkbImport.Sheets(1).Range("O2").AutoFill Destination:=wkbImport.Sheets(1).Range("O2:O" & lastRowH)
        
        For i = 2 To lastRowA
            UltimaPosicao = InStrRev(wkbImport.Sheets(1).Range("B" & i).Value, ", ")
            wkbImport.Sheets(1).Range("J" & i).Value = UltimaPosicao
        Next i
        
        Set skuDict = CreateObject("Scripting.Dictionary")
        For i = 2 To lastRowA
            currentSKU = wkbImport.Sheets(1).Cells(i, "D").Value
            If Not skuDict.Exists(currentSKU) And wkbImport.Sheets(1).Cells(i, "C").Value <> 0 Then
                skuDict.Add currentSKU, wkbImport.Sheets(1).Cells(i, "J").Value
            ElseIf wkbImport.Sheets(1).Cells(i, "J").Value <> "" And wkbImport.Sheets(1).Cells(i, "C").Value <> 0 Then
                skuDict(currentSKU) = skuDict(currentSKU) & ";" & wkbImport.Sheets(1).Cells(i, "J").Value
            End If
        Next i
        
        ' Transferência para a aba Menu
        wkbImport.Sheets(1).Columns("M:M").Copy
        mn.Visible = True
        mn.Activate
        mn.Range("A1").PasteSpecial Paste:=xlPasteValues
        
        ult_linha = mn.Range("A" & Rows.Count).End(xlUp).Row
        
        ' Loop de inserção na Base de Dados principal
        For n = 2 To ult_linha
            Valor_Linha = mn.Cells(n, 1).Value
            last_row = db.Cells(Rows.Count, "B").End(xlUp).Row
            
            For i = 3 To last_row
                If db.Range("B" & i).Value = Valor_Linha Then
                    Valor_Linha = db.Range("B" & i).Row
                    Exit For
                End If
            Next i
                
            db.Rows(Valor_Linha + 1).Insert Shift:=xlDown
            db.Rows(Valor_Linha).Copy Destination:=db.Rows(Valor_Linha + 1)
            
            maiorValor = WorksheetFunction.Max(db.Range("B3:B" & last_row + 1))
            db.Range("B" & Valor_Linha + 1).Value = maiorValor + 1
            db.Cells(Valor_Linha + 1, 2).Interior.Color = RGB(200, 200, 200) ' Cor neutra
            
            ' Limpeza de colunas de planejamento (nomes alterados para genéricos)
            For i = 1 To db.Range("A1").End(xlToRight).Column
                ColunaNome = db.Cells(2, i).Value
                If ColunaNome = "Campo_A" Or ColunaNome = "Campo_B" Or ColunaNome = "Plano_Qtd" Then
                    db.Cells(Valor_Linha + 1, i).ClearContents
                End If
            Next i

            ' Inserção dos dados processados
            db.Cells(Valor_Linha + 1, Coluna_OrigemDoModelo).Value = "Inserida"
        Next n
        
        ' Finalização e Backup
        wkbImport.SaveAs caminho_pasta & "Backup_Log\" & Format(Now, "yyyymmdd_hhmmss") & "_Backup_Dados.xlsx"
        wkbImport.Close SaveChanges:=False
        
        Call bBloqueio
        Call AtualizarPedidos
        
        ' Log de Finalização
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Processo Importação"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Finalizada"
        End With
    
    MsgBox "Processamento concluído com sucesso!"
    Application.ScreenUpdating = True

End Sub
