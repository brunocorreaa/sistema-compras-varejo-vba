Attribute VB_Name = "ExcluirLinha"
Sub bExcluirLinha()

    Application.ScreenUpdating = False

        ' Declarações
        Dim db As Worksheet
        Dim mn As Worksheet
        Dim Controle_Macro As Worksheet
        Dim last_row As Long
        Dim rng As Range
        Dim maiorValor As Double
        
        ' Atribuições (Nomes de abas anonimizados)
        Set db = ThisWorkbook.Sheets("BASE_DADOS")
        Set mn = ThisWorkbook.Sheets("Menu")
        Set Controle_Macro = ThisWorkbook.Sheets("LOG_SISTEMA")
    
        ' Validação de execução
        resposta = MsgBox("Você realmente quer executar a operação: EXCLUIR REGISTRO?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação")
        
        If resposta <> vbYes Then
            Exit Sub
        End If

        ' Auditoria
        usuario = Environ("Username")
        hoje = Date
        horaAtual = Format(Time, "hh:mm:ss")

        ' Registro de Log
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Exclusão Registro"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Iniciada"
        End With

        ' Chamada de validações internas
        Call Validacoes("ExcluirRegistro")

        ' Desbloqueio de segurança
        Call bDesbloqueio

        '<<< MAPEAMENTO DE COLUNAS >>>'
    
        ' Nomes das colunas anonimizados conforme regras de negócio genéricas
        nomesColunas = Array("Origem_Entrada", "Status_Processamento")
        
        For Each nome In nomesColunas
            Coluna = 0
            nomeProcurado = nome
            For Each Celula In db.Rows(2).Cells
                If Celula.Value = nomeProcurado Then
                    Select Case nomeProcurado
                        Case "Origem_Entrada": Coluna_OrigemDoModelo = Celula.Column
                        Case "Status_Processamento": Coluna_PedidoEmitido = Celula.Column
                    End Select
                    Exit For
                End If
            Next Celula
        Next nome
        
    
        ' Determinação da última linha preenchida na Base
        last_row = db.Cells(Rows.Count, "B").End(xlUp).Row

        ' Coleta de IDs únicos via Dictionary para processamento em lote
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        For j = 3 To last_row
            ' Verifica flag de seleção na Coluna C
            If db.Cells(j, "C").Value <> "" Then
                If Not dict.Exists(db.Cells(j, "B").Value) Then
                    dict.Add db.Cells(j, "B").Value, True
                End If
            End If
        Next j
        
        ' Conversão das chaves em Array
        Valor_Linha = Join(dict.Keys, ";")
        valores = Split(Valor_Linha, ";")
        Valor_Linha_Count = UBound(valores) + 1
        
        ' Loop de exclusão baseado nos IDs coletados
        For k = 0 To Valor_Linha_Count - 1
            
            Valor_Linha = valores(k)
            
            ' Localização da linha física no Excel
            db.Select
            For m = 3 To last_row
                If db.Range("B" & m).Text = Valor_Linha Then
                    Valor_Linha_Plano = db.Range("B" & m).Row
                    Exit For
                End If
            Next m
            
            ' Regra de Negócio: Só exclui se a origem for "Inserida" e o status for "Não" (Não Processado)
            If db.Cells(Valor_Linha_Plano, Coluna_OrigemDoModelo).Value = "Inserida" And db.Cells(Valor_Linha_Plano, Coluna_PedidoEmitido).Value = "Não" Then
                db.Rows(Valor_Linha_Plano & ":" & Valor_Linha_Plano).Delete
            Else
                MsgBox "O registro " & Valor_Linha & " não pode ser excluído: origem protegida ou status já processado."
            End If
        
        Next k
        
        ' Reativação da segurança
        Call bBloqueio
        
        ' Log de Finalização
        last_row_macro = Controle_Macro.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With Controle_Macro
            .Range("A" & last_row_macro).Value = "Exclusão Registro"
            .Range("B" & last_row_macro).Value = hoje
            .Range("C" & last_row_macro).Value = horaAtual
            .Range("D" & last_row_macro).Value = usuario
            .Range("E" & last_row_macro).Value = "Finalizada"
        End With
    
    MsgBox "Operação concluída!"
    
    Application.ScreenUpdating = True
    
End Sub
