Attribute VB_Name = "bFunction"
' Converte o índice numérico de uma coluna em sua letra correspondente (ex: 1 -> A)
Function ObterLetraColuna(ByVal numeroColuna As Long) As String
    On Error Resume Next
    ObterLetraColuna = Split(Cells(1, numeroColuna).Address, "$")(1)
    On Error GoTo 0
End Function

' Realiza busca binária simulada via Array em duas colunas com tratamento de strings
Function BuscarPorDuasCondicoes(CriterioA As String, CriterioB As String) As Variant
    
    Dim wsApoio As Worksheet
    Dim uLinha As Long
    Dim i As Long
    Dim stringTratadaB As String
    Dim matrizBusca() As Variant
    Dim matrizResultado() As Variant
    
    ' Define a planilha de referência
    Set wsApoio = ThisWorkbook.Sheets("Apoio")
    
    ' Limpeza de caracteres especiais do Critério B (Removendo ".", "/", "-")
    stringTratadaB = Replace(Replace(Replace(CriterioB, ".", ""), "/", ""), "-", "")
    
    ' Identifica a última célula preenchida
    uLinha = wsApoio.Cells(wsApoio.Rows.Count, "R").End(xlUp).Row
    
    ' Se a planilha estiver vazia, encerra
    If uLinha < 1 Then
        BuscarPorDuasCondicoes = "Não Encontrado"
        Exit Function
    End If
    
    ' Carrega os dados em memória (Arrays) para performance otimizada
    ' Matriz R (Busca 1), P (Busca 2) e Q (Resultado)
    matrizBusca = wsApoio.Range("P1:R" & uLinha).Value
    
    ' Percorre a matriz em memória
    For i = 1 To UBound(matrizBusca, 1)
        ' Verifica Coluna R (Índice 3 da matriz) e Coluna P tratada (Índice 1 da matriz)
        If matrizBusca(i, 3) = CriterioA Then
            If Replace(Replace(Replace(matrizBusca(i, 1), ".", ""), "/", ""), "-", "") = stringTratadaB Then
                ' Retorna o valor da Coluna Q (Índice 2 da matriz)
                BuscarPorDuasCondicoes = matrizBusca(i, 2)
                Exit Function
            End If
        End If
    Next i
    
    ' Retorno padrão caso não encontre correspondência
    BuscarPorDuasCondicoes = "Não Encontrado"

End Function
