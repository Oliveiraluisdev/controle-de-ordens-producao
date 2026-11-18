Attribute VB_Name = "mdlOP"
'==============================================================================
' mdlOP - Regras das OPs
' Controle de Ordens por Maquina
'
' Cadastro, edicao, exclusao, arquivamento, restauracao, validacao,
' duplicidade e checklist. Uma unica logica serve para todas as maquinas.
'==============================================================================
Option Explicit

'==============================================================================
' LOCALIZACAO
'==============================================================================

'------------------------------------------------------------------------------
' Procura a OP em todas as maquinas e no ARQUIVO.
' Devolve True se encontrar, com a aba e o indice da linha na tabela.
'   ignorarAba / ignorarIdx: pula uma linha especifica (usado na edicao).
'------------------------------------------------------------------------------
Public Function LocalizarOP(ByVal numeroOP As String, _
                            ByRef wsAchada As Worksheet, _
                            ByRef tblAchada As ListObject, _
                            ByRef idxAchado As Long, _
                            Optional ByVal ignorarTabela As String = "", _
                            Optional ByVal ignorarIdx As Long = 0) As Boolean

    Dim chave As String
    Dim nome As Variant
    Dim ws As Worksheet

    chave = ChaveOP(numeroOP)
    If Len(chave) = 0 Then Exit Function

    For Each nome In ListarMaquinas
        Set ws = AbaMaquina(CStr(nome))
        If Not ws Is Nothing Then
            If ProcurarNaAba(ws, chave, tblAchada, idxAchado, ignorarTabela, ignorarIdx) Then
                Set wsAchada = ws
                LocalizarOP = True
                Exit Function
            End If
        End If
    Next nome

    If PlanilhaExiste(ABA_ARQUIVO) Then
        Set ws = ThisWorkbook.Worksheets(ABA_ARQUIVO)
        If ProcurarNaAba(ws, chave, tblAchada, idxAchado, ignorarTabela, ignorarIdx) Then
            Set wsAchada = ws
            LocalizarOP = True
        End If
    End If

End Function

'------------------------------------------------------------------------------
' Procura a OP em todos os blocos de uma aba.
'------------------------------------------------------------------------------
Private Function ProcurarNaAba(ByVal ws As Worksheet, ByVal chave As String, _
                               ByRef tblAchada As ListObject, _
                               ByRef idxAchado As Long, _
                               ByVal ignorarTabela As String, _
                               ByVal ignorarIdx As Long) As Boolean

    Dim tbl As ListObject
    Dim i As Long

    For Each tbl In TabelasDaAba(ws)
        If Not tbl.DataBodyRange Is Nothing Then
            For i = 1 To tbl.ListRows.Count
                If Not (StrComp(tbl.Name, ignorarTabela, vbTextCompare) = 0 And _
                        i = ignorarIdx) Then
                    If ChaveOP(CStr(tbl.DataBodyRange.Cells(i, COL_OP).Value)) = chave Then
                        Set tblAchada = tbl
                        idxAchado = i
                        ProcurarNaAba = True
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next tbl

End Function

'------------------------------------------------------------------------------
' Linha selecionada pelo usuario na aba ativa.
' Devolve True se a celula ativa estiver numa linha de OP valida.
'------------------------------------------------------------------------------
Public Function LinhaSelecionada(ByRef tbl As ListObject, ByRef idx As Long) As Boolean

    Dim ws As Worksheet

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    Set ws = ActiveSheet

    Set tbl = TabelaDaCelula(ws, ActiveCell)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Then Exit Function

    idx = ActiveCell.Row - tbl.DataBodyRange.Row + 1
    If Len(Trim$(CStr(tbl.DataBodyRange.Cells(idx, COL_OP).Value))) = 0 Then Exit Function

    LinhaSelecionada = True

End Function

'==============================================================================
' VALIDACAO
'==============================================================================
Public Function ValidarCampos(ByVal maquina As String, ByVal numeroOP As String, _
                              ByVal item As String, ByVal qtdTexto As String, _
                              ByVal cliente As String, ByVal mes As String, _
                              ByRef msg As String) As Boolean

    If AbaMaquina(maquina) Is Nothing Then
        msg = "Selecione a maquina."
        Exit Function
    End If

    If Len(Trim$(numeroOP)) = 0 Then
        msg = "Informe o numero da OP."
        Exit Function
    End If

    If Not TextoNumericoValido(numeroOP) Then
        msg = "Informe um numero de OP valido (somente numeros)."
        Exit Function
    End If

    If Len(Trim$(item)) = 0 Then
        msg = "Informe o item."
        Exit Function
    End If

    If Not TextoNumericoValido(qtdTexto) Then
        msg = "Informe uma quantidade valida."
        Exit Function
    End If

    If CDbl(SomenteDigitos(qtdTexto)) <= 0 Then
        msg = "Informe uma quantidade valida."
        Exit Function
    End If

    If Len(Trim$(cliente)) = 0 Then
        msg = "Informe o cliente."
        Exit Function
    End If

    If Len(Trim$(mes)) = 0 Then
        msg = "Informe o mes."
        Exit Function
    End If

    If Len(NormalizarMes(mes)) = 0 Then
        msg = "Informe um mes valido. Exemplos: OUTUBRO/2026, out/2026, 10/2026."
        Exit Function
    End If

    ValidarCampos = True

End Function

Private Function MensagemDuplicada(ByVal wsAchada As Worksheet) As String
    If wsAchada.Name = ABA_ARQUIVO Then
        MensagemDuplicada = "Esta OP ja esta cadastrada (esta no ARQUIVO)." & vbNewLine & _
                            "Verifique os dados antes de continuar."
    Else
        MensagemDuplicada = "Esta OP ja esta cadastrada na " & wsAchada.Name & "." & _
                            vbNewLine & "Verifique os dados antes de continuar."
    End If
End Function

'==============================================================================
' LEITURA E GRAVACAO DE LINHAS
' Uma linha e transportada em dois vetores:
'   valores(1 To 10) - conteudo das colunas (a 10 so existe no ARQUIVO)
'   notas(1 To 4)    - comentarios com a data de cada marcacao
'==============================================================================
Private Sub LerLinha(ByVal tbl As ListObject, ByVal idx As Long, _
                     ByRef valores() As Variant, ByRef notas() As String)

    Dim c As Long
    Dim cel As Range

    ReDim valores(1 To COL_MES)
    ReDim notas(1 To QTD_TAREFAS)

    For c = 1 To tbl.ListColumns.Count
        If c <= COL_MES Then valores(c) = tbl.DataBodyRange.Cells(idx, c).Value
    Next c

    For c = 1 To QTD_TAREFAS
        Set cel = tbl.DataBodyRange.Cells(idx, COL_TAREFA1 + c - 1)
        If Not cel.Comment Is Nothing Then notas(c) = cel.Comment.Text
    Next c

End Sub

Private Sub EscreverLinha(ByVal tbl As ListObject, ByVal idx As Long, _
                          ByRef valores() As Variant, ByRef notas() As String)

    Dim c As Long
    Dim cel As Range
    Dim linha As Range

    Set linha = tbl.ListRows(idx).Range

    ' OP sempre como texto, para preservar "62.620"
    linha.Cells(1, COL_OP).NumberFormat = "@"
    linha.Cells(1, COL_OP).Value = CStr(valores(COL_OP))

    linha.Cells(1, COL_ITEM).NumberFormat = "@"
    linha.Cells(1, COL_ITEM).Value = CStr(valores(COL_ITEM))

    linha.Cells(1, COL_QTD).NumberFormat = "#,##0"
    linha.Cells(1, COL_QTD).Value = valores(COL_QTD)

    linha.Cells(1, COL_CLIENTE).NumberFormat = "@"
    linha.Cells(1, COL_CLIENTE).Value = CStr(valores(COL_CLIENTE))

    For c = 1 To QTD_TAREFAS
        Set cel = linha.Cells(1, COL_TAREFA1 + c - 1)
        If CStr(valores(COL_TAREFA1 + c - 1)) = CaixaMarcada Then
            cel.Value = CaixaMarcada
        Else
            cel.Value = CaixaVazia
        End If
        cel.ClearComments
        If Len(notas(c)) > 0 Then
            cel.AddComment notas(c)
            cel.Comment.Shape.TextFrame.AutoSize = True
        End If
    Next c

    linha.Cells(1, COL_CADASTRO).NumberFormat = "dd/mm/yyyy"
    linha.Cells(1, COL_CADASTRO).Value = valores(COL_CADASTRO)

    ' Colunas que so existem no ARQUIVO
    If tbl.ListColumns.Count >= COL_MAQUINA Then
        linha.Cells(1, COL_MAQUINA).NumberFormat = "@"
        linha.Cells(1, COL_MAQUINA).Value = CStr(valores(COL_MAQUINA))
    End If

    If tbl.ListColumns.Count >= COL_MES Then
        linha.Cells(1, COL_MES).NumberFormat = "@"
        linha.Cells(1, COL_MES).Value = CStr(valores(COL_MES))
    End If

    FormatarLinha tbl, idx

End Sub

'------------------------------------------------------------------------------
' Aparencia padrao de uma linha + destaque de checklist completo.
'------------------------------------------------------------------------------
Private Sub FormatarLinha(ByVal tbl As ListObject, ByVal idx As Long)

    Dim linha As Range
    Dim tarefas As Range
    Dim completa As Boolean
    Dim c As Long

    Set linha = tbl.ListRows(idx).Range
    Set tarefas = linha.Cells(1, COL_TAREFA1).Resize(1, QTD_TAREFAS)

    With tarefas
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 13
    End With

    linha.Cells(1, COL_QTD).HorizontalAlignment = xlRight
    linha.Cells(1, COL_CADASTRO).HorizontalAlignment = xlCenter

    completa = True
    For c = 1 To QTD_TAREFAS
        If CStr(tarefas.Cells(1, c).Value) <> CaixaMarcada Then
            completa = False
            Exit For
        End If
    Next c

    If completa Then
        linha.Interior.Color = COR_CONCLUIDA
        linha.Font.Color = COR_TEXTO_CONCLUIDA
    Else
        linha.Interior.Pattern = xlNone
        linha.Font.ColorIndex = xlAutomatic
        linha.Cells(1, COL_CADASTRO).Font.Color = COR_TEXTO_CINZA
    End If

    linha.Cells(1, COL_OP).Font.Bold = True

End Sub

'------------------------------------------------------------------------------
' Remove a linha-modelo vazia que a tabela traz quando ainda nao tem OPs.
'------------------------------------------------------------------------------
Private Sub RemoverLinhasVazias(ByVal tbl As ListObject)

    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(tbl.DataBodyRange.Cells(i, COL_OP).Value))) = 0 And _
           Len(Trim$(CStr(tbl.DataBodyRange.Cells(i, COL_ITEM).Value))) = 0 Then
            tbl.ListRows(i).Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' Posicao que a OP deve ocupar para manter a ordem numerica.
'------------------------------------------------------------------------------
Private Function PosicaoOrdenada(ByVal tbl As ListObject, ByVal numeroOP As String, _
                                 Optional ByVal ignorarIdx As Long = 0) As Long

    Dim i As Long
    Dim alvo As Double

    alvo = ValorOP(numeroOP)

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            If i <> ignorarIdx Then
                If ValorOP(CStr(tbl.DataBodyRange.Cells(i, COL_OP).Value)) > alvo Then
                    PosicaoOrdenada = i
                    Exit Function
                End If
            End If
        Next i
    End If

    PosicaoOrdenada = tbl.ListRows.Count + 1

End Function

'------------------------------------------------------------------------------
' Insere uma linha vazia na posicao pedida e devolve o indice dela.
'------------------------------------------------------------------------------
Private Function InserirLinha(ByVal tbl As ListObject, ByVal posicao As Long) As Long

    If tbl.ListRows.Count = 0 Then
        tbl.ListRows.Add
        InserirLinha = 1
    ElseIf posicao > tbl.ListRows.Count Then
        tbl.ListRows.Add AlwaysInsert:=True
        InserirLinha = tbl.ListRows.Count
    Else
        tbl.ListRows.Add Position:=posicao, AlwaysInsert:=True
        InserirLinha = posicao
    End If

End Function

'------------------------------------------------------------------------------
' Move uma linha (com checklist e comentarios) para a tabela de destino,
' na ordem numerica correta, e apaga a original.
' Funciona tanto dentro da mesma tabela quanto entre tabelas.
' Devolve o indice final da linha no destino.
'------------------------------------------------------------------------------
Private Function MoverLinha(ByVal tblOrigem As ListObject, ByVal idxOrigem As Long, _
                            ByVal tblDestino As ListObject, _
                            ByRef valores() As Variant, ByRef notas() As String) As Long

    Dim mesmaTabela As Boolean
    Dim posicao As Long
    Dim idxNovo As Long

    mesmaTabela = (tblOrigem.Name = tblDestino.Name)

    If mesmaTabela Then
        posicao = PosicaoOrdenada(tblDestino, CStr(valores(COL_OP)), idxOrigem)
    Else
        RemoverLinhasVazias tblDestino
        posicao = PosicaoOrdenada(tblDestino, CStr(valores(COL_OP)))
    End If

    ' Mesma posicao na mesma tabela: basta regravar
    If mesmaTabela And (posicao = idxOrigem Or posicao = idxOrigem + 1) Then
        EscreverLinha tblDestino, idxOrigem, valores, notas
        MoverLinha = idxOrigem
        Exit Function
    End If

    idxNovo = InserirLinha(tblDestino, posicao)
    EscreverLinha tblDestino, idxNovo, valores, notas

    If mesmaTabela Then
        If idxNovo <= idxOrigem Then idxOrigem = idxOrigem + 1
        tblOrigem.ListRows(idxOrigem).Delete
        If idxOrigem < idxNovo Then idxNovo = idxNovo - 1
    Else
        tblOrigem.ListRows(idxOrigem).Delete
    End If

    MoverLinha = idxNovo

End Function

'------------------------------------------------------------------------------
' Seleciona a celula da OP de uma linha e deixa ela visivel.
'------------------------------------------------------------------------------
Private Sub MostrarLinha(ByVal tbl As ListObject, ByVal idx As Long)

    On Error Resume Next
    tbl.Parent.Activate
    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
    tbl.DataBodyRange.Cells(idx, COL_OP).Select
    On Error GoTo 0

End Sub

'==============================================================================
' CADASTRO
'==============================================================================
Public Function CadastrarOP(ByVal maquina As String, ByVal numeroOP As String, _
                            ByVal item As String, ByVal qtdTexto As String, _
                            ByVal cliente As String, ByVal mes As String, _
                            ByRef msg As String) As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsAchada As Worksheet
    Dim tblAchada As ListObject
    Dim idxAchado As Long
    Dim valores() As Variant
    Dim notas() As String
    Dim idxNovo As Long
    Dim c As Long

    On Error GoTo TrataErro

    If Not ValidarCampos(maquina, numeroOP, item, qtdTexto, cliente, mes, msg) Then
        Exit Function
    End If

    If LocalizarOP(numeroOP, wsAchada, tblAchada, idxAchado) Then
        msg = MensagemDuplicada(wsAchada)
        Exit Function
    End If

    mes = NormalizarMes(mes)
    Set ws = AbaMaquina(maquina)

    OtimizarInicio
    Set tbl = TabelaDoMes(ws, mes, True)      ' cria o bloco do mes se faltar
    OtimizarFim

    If tbl Is Nothing Then
        msg = "Nao consegui criar o bloco do mes " & mes & " na " & ws.Name & "."
        Exit Function
    End If

    ReDim valores(1 To COL_MES)
    ReDim notas(1 To QTD_TAREFAS)

    valores(COL_OP) = FormatarMilhar(numeroOP)
    valores(COL_ITEM) = UCase$(Trim$(item))
    valores(COL_QTD) = CDbl(SomenteDigitos(qtdTexto))
    valores(COL_CLIENTE) = UCase$(Trim$(cliente))
    For c = 1 To QTD_TAREFAS
        valores(COL_TAREFA1 + c - 1) = CaixaVazia     ' sempre desmarcado
    Next c
    valores(COL_CADASTRO) = Date

    OtimizarInicio

    RemoverLinhasVazias tbl
    idxNovo = InserirLinha(tbl, PosicaoOrdenada(tbl, CStr(valores(COL_OP))))
    EscreverLinha tbl, idxNovo, valores, notas

    AdicionarNaLista CFG_ITENS, valores(COL_ITEM)
    AdicionarNaLista CFG_CLIENTES, valores(COL_CLIENTE)
    AdicionarNaLista CFG_QUANTIDADES, valores(COL_QTD)

    OtimizarFim
    MostrarLinha tbl, idxNovo

    msg = "OP " & valores(COL_OP) & " cadastrada na " & ws.Name & " - " & mes & "."
    CadastrarOP = True
    Exit Function

TrataErro:
    OtimizarFim
    msg = "Erro ao cadastrar: " & Err.Description

End Function

'==============================================================================
' EDICAO
'==============================================================================
Public Function EditarOP(ByVal abaOrigem As String, ByVal opOriginal As String, _
                         ByVal maquina As String, ByVal numeroOP As String, _
                         ByVal item As String, ByVal qtdTexto As String, _
                         ByVal cliente As String, ByVal mes As String, _
                         ByRef msg As String) As Boolean

    Dim wsOrig As Worksheet
    Dim tblOrig As ListObject
    Dim tblDest As ListObject
    Dim idxOrig As Long
    Dim wsAchada As Worksheet
    Dim tblAchada As ListObject
    Dim idxAchado As Long
    Dim valores() As Variant
    Dim notas() As String
    Dim idxFinal As Long

    On Error GoTo TrataErro

    If Not ValidarCampos(maquina, numeroOP, item, qtdTexto, cliente, mes, msg) Then
        Exit Function
    End If

    mes = NormalizarMes(mes)

    ' Relocaliza a OP original (a planilha pode ter mudado desde a abertura)
    Set wsOrig = ThisWorkbook.Worksheets(abaOrigem)
    If Not ProcurarNaAba(wsOrig, ChaveOP(opOriginal), tblOrig, idxOrig, "", 0) Then
        msg = "Nao encontrei mais a OP " & opOriginal & " na " & abaOrigem & "."
        Exit Function
    End If

    ' Duplicidade, ignorando a propria OP
    If LocalizarOP(numeroOP, wsAchada, tblAchada, idxAchado, tblOrig.Name, idxOrig) Then
        msg = MensagemDuplicada(wsAchada)
        Exit Function
    End If

    OtimizarInicio
    Set tblDest = TabelaDoMes(AbaMaquina(maquina), mes, True)
    OtimizarFim

    If tblDest Is Nothing Then
        msg = "Nao consegui criar o bloco do mes " & mes & "."
        Exit Function
    End If

    ' Le tudo (checklist, notas, data de cadastro) e troca so os dados
    LerLinha tblOrig, idxOrig, valores, notas
    valores(COL_OP) = FormatarMilhar(numeroOP)
    valores(COL_ITEM) = UCase$(Trim$(item))
    valores(COL_QTD) = CDbl(SomenteDigitos(qtdTexto))
    valores(COL_CLIENTE) = UCase$(Trim$(cliente))

    OtimizarInicio

    idxFinal = MoverLinha(tblOrig, idxOrig, tblDest, valores, notas)

    AdicionarNaLista CFG_ITENS, valores(COL_ITEM)
    AdicionarNaLista CFG_CLIENTES, valores(COL_CLIENTE)
    AdicionarNaLista CFG_QUANTIDADES, valores(COL_QTD)

    OtimizarFim
    MostrarLinha tblDest, idxFinal

    msg = "OP " & valores(COL_OP) & " atualizada (" & mes & ")."
    If tblDest.Parent.Name <> wsOrig.Name Then
        msg = msg & vbNewLine & "Movida da " & wsOrig.Name & " para a " & tblDest.Parent.Name & "."
    End If

    EditarOP = True
    Exit Function

TrataErro:
    OtimizarFim
    msg = "Erro ao editar: " & Err.Description

End Function

'==============================================================================
' EXCLUSAO
'==============================================================================
Public Sub ExcluirOPSelecionada()

    Dim tbl As ListObject
    Dim idx As Long
    Dim numeroOP As String

    If Not EhAbaMaquina(ActiveSheet) Then
        Alertar "Use este botao dentro da aba de uma maquina."
        Exit Sub
    End If

    If Not LinhaSelecionada(tbl, idx) Then
        Alertar "Clique em uma linha da tabela para selecionar a OP."
        Exit Sub
    End If

    numeroOP = CStr(tbl.DataBodyRange.Cells(idx, COL_OP).Value)

    If Not Confirmar("Excluir a OP " & numeroOP & "?" & vbNewLine & vbNewLine & _
                     "Esta acao nao pode ser desfeita.") Then Exit Sub

    OtimizarInicio
    tbl.ListRows(idx).Delete
    OtimizarFim

End Sub

'==============================================================================
' ARQUIVAMENTO E RESTAURACAO
'==============================================================================
Public Sub ArquivarOPSelecionada()

    Dim tbl As ListObject
    Dim tblArq As ListObject
    Dim idx As Long
    Dim valores() As Variant
    Dim notas() As String
    Dim maquina As String

    On Error GoTo TrataErro

    If Not EhAbaMaquina(ActiveSheet) Then
        Alertar "Use este botao dentro da aba de uma maquina."
        Exit Sub
    End If

    If Not LinhaSelecionada(tbl, idx) Then
        Alertar "Clique em uma linha da tabela para selecionar a OP."
        Exit Sub
    End If

    maquina = ActiveSheet.Name
    LerLinha tbl, idx, valores, notas
    valores(COL_MES) = MesDaTabela(tbl)

    If Not Confirmar("Arquivar a OP " & valores(COL_OP) & "?" & vbNewLine & vbNewLine & _
                     "Ela sai da " & maquina & " e vai para o ARQUIVO." & vbNewLine & _
                     "Voce pode restaura-la depois.") Then Exit Sub

    valores(COL_MAQUINA) = maquina
    Set tblArq = TabelaDaAba(ThisWorkbook.Worksheets(ABA_ARQUIVO))

    OtimizarInicio
    MoverLinha tbl, idx, tblArq, valores, notas
    OtimizarFim

    Avisar "OP " & valores(COL_OP) & " arquivada."
    Exit Sub

TrataErro:
    OtimizarFim
    Alertar "Erro ao arquivar: " & Err.Description

End Sub

Public Sub RestaurarOPSelecionada()

    Dim tbl As ListObject
    Dim tblDest As ListObject
    Dim idx As Long
    Dim idxFinal As Long
    Dim valores() As Variant
    Dim notas() As String
    Dim wsDest As Worksheet
    Dim mes As String

    On Error GoTo TrataErro

    If ActiveSheet.Name <> ABA_ARQUIVO Then
        Alertar "Use este botao dentro da aba ARQUIVO."
        Exit Sub
    End If

    If Not LinhaSelecionada(tbl, idx) Then
        Alertar "Clique em uma linha da tabela para selecionar a OP."
        Exit Sub
    End If

    LerLinha tbl, idx, valores, notas
    Set wsDest = AbaMaquina(CStr(valores(COL_MAQUINA)))

    If wsDest Is Nothing Then
        Alertar "A maquina de origem (" & valores(COL_MAQUINA) & ") nao existe mais." & _
                vbNewLine & "Crie a maquina antes de restaurar esta OP."
        Exit Sub
    End If

    If Not Confirmar("Restaurar a OP " & valores(COL_OP) & " para a " & _
                     wsDest.Name & "?") Then Exit Sub

    mes = NormalizarMes(CStr(valores(COL_MES)))
    If Len(mes) = 0 Then mes = MesAtual

    valores(COL_MAQUINA) = vbNullString
    valores(COL_MES) = vbNullString

    OtimizarInicio
    Set tblDest = TabelaDoMes(wsDest, mes, True)
    If tblDest Is Nothing Then
        OtimizarFim
        Alertar "Nao consegui criar o bloco do mes " & mes & " na " & wsDest.Name & "."
        Exit Sub
    End If
    idxFinal = MoverLinha(tbl, idx, tblDest, valores, notas)
    OtimizarFim

    MostrarLinha tblDest, idxFinal
    Avisar "OP " & valores(COL_OP) & " restaurada na " & wsDest.Name & " - " & mes & "."
    Exit Sub

TrataErro:
    OtimizarFim
    Alertar "Erro ao restaurar: " & Err.Description

End Sub

'==============================================================================
' CHECKLIST - chamado pelo duplo clique (EstaPastaDeTrabalho)
'==============================================================================
Public Sub TratarDuploClique(ByVal ws As Object, ByVal alvo As Range, ByRef cancelar As Boolean)

    Dim tbl As ListObject
    Dim coluna As Long
    Dim idx As Long

    If Not EhAbaMaquina(ws) Then Exit Sub
    If alvo.CountLarge <> 1 Then Exit Sub

    Set tbl = TabelaDaCelula(ws, alvo)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(alvo, tbl.DataBodyRange) Is Nothing Then Exit Sub

    coluna = alvo.Column - tbl.Range.Column + 1
    If coluna < COL_TAREFA1 Or coluna > COL_TAREFA1 + QTD_TAREFAS - 1 Then Exit Sub

    idx = alvo.Row - tbl.DataBodyRange.Row + 1
    If Len(Trim$(CStr(tbl.DataBodyRange.Cells(idx, COL_OP).Value))) = 0 Then Exit Sub

    cancelar = True                   ' nao entra no modo de edicao da celula

    Application.EnableEvents = False
    On Error GoTo Fim

    If CStr(alvo.Value) = CaixaMarcada Then
        alvo.Value = CaixaVazia
        alvo.ClearComments            ' desmarcou: a data some junto
    Else
        alvo.Value = CaixaMarcada
        alvo.ClearComments
        alvo.AddComment "Marcado em " & Format$(Now, "dd/mm/yyyy hh:nn")
        alvo.Comment.Shape.TextFrame.AutoSize = True
    End If

    FormatarLinha tbl, idx

Fim:
    Application.EnableEvents = True

End Sub

'==============================================================================
' ABERTURA DO FORMULARIO
'==============================================================================
Public Sub AbrirFormNovo(ByVal maquina As String)

    Dim f As frmOP

    Set f = New frmOP
    f.PrepararNovo maquina
    f.Show

End Sub

Public Sub AbrirFormEdicao()

    Dim tbl As ListObject
    Dim idx As Long
    Dim f As frmOP
    Dim linha As Range

    If Not EhAbaMaquina(ActiveSheet) Then
        Alertar "Use este botao dentro da aba de uma maquina."
        Exit Sub
    End If

    If Not LinhaSelecionada(tbl, idx) Then
        Alertar "Clique em uma linha da tabela para selecionar a OP."
        Exit Sub
    End If

    Set linha = tbl.ListRows(idx).Range

    Set f = New frmOP
    f.PrepararEdicao ActiveSheet.Name, _
                     CStr(linha.Cells(1, COL_OP).Value), _
                     CStr(linha.Cells(1, COL_ITEM).Value), _
                     CStr(linha.Cells(1, COL_QTD).Value), _
                     CStr(linha.Cells(1, COL_CLIENTE).Value), _
                     MesDaTabela(tbl)
    f.Show

End Sub

