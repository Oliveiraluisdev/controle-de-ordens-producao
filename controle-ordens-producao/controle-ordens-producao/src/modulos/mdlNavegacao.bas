Attribute VB_Name = "mdlNavegacao"
'==============================================================================
' mdlNavegacao - MENU, botoes, busca rapida, setores, nova e exclusao de maquina
' Controle de Ordens por Maquina
'
' Todos os botoes sao criados por codigo. Rode InstalarSistema uma vez depois
' de colar o codigo, e sempre que quiser redesenhar os botoes.
'==============================================================================
Option Explicit

Private Const PREFIXO_BOTAO As String = "btnSis_"      ' identifica botoes do sistema
Private Const COR_BOTAO As Long = 2039583              ' RGB(31, 31, 31)
Private Const COR_BOTAO_CLARO As Long = 15921906       ' RGB(242, 242, 242)

'==============================================================================
' INSTALACAO
'==============================================================================
Public Sub InstalarSistema()

    Dim nome As Variant
    Dim ws As Worksheet

    On Error GoTo TrataErro
    OtimizarInicio

    AtribuirSetoresPadrao
    MigrarBlocos
    MontarMenu

    For Each nome In ListarMaquinas
        Set ws = AbaMaquina(CStr(nome))
        If Not ws Is Nothing Then MontarBotoesMaquina ws
    Next nome

    MontarBotoesArquivo ThisWorkbook.Worksheets(ABA_ARQUIVO)

    OtimizarFim
    IrParaMenu

    Avisar "Sistema instalado." & vbNewLine & vbNewLine & _
           "Os botoes foram criados no MENU, nas maquinas e no ARQUIVO."
    Exit Sub

TrataErro:
    OtimizarFim
    Alertar "Erro na instalacao: " & Err.Description

End Sub

'==============================================================================
' BOTOES (acoes chamadas pelos botoes das abas)
'==============================================================================
Public Sub BotaoNovaOP()
    If EhAbaMaquina(ActiveSheet) Then
        AbrirFormNovo ActiveSheet.Name
    Else
        AbrirFormNovo vbNullString
    End If
End Sub

Public Sub BotaoEditarOP()
    AbrirFormEdicao
End Sub

Public Sub BotaoExcluirOP()
    ExcluirOPSelecionada
End Sub

Public Sub BotaoArquivarOP()
    ArquivarOPSelecionada
End Sub

Public Sub BotaoRestaurarOP()
    RestaurarOPSelecionada
End Sub

Public Sub IrParaMenu()
    ThisWorkbook.Worksheets(ABA_MENU).Activate
    ThisWorkbook.Worksheets(ABA_MENU).Range(CELULA_BUSCA).Select
End Sub

Public Sub IrParaArquivo()
    ThisWorkbook.Worksheets(ABA_ARQUIVO).Activate
End Sub

'------------------------------------------------------------------------------
' Botao de maquina no MENU. O nome da maquina fica no texto alternativo
' do proprio botao.
'------------------------------------------------------------------------------
Public Sub IrParaMaquina()

    Dim nomeBotao As String
    Dim ws As Worksheet

    On Error Resume Next
    nomeBotao = CStr(Application.Caller)
    Set ws = AbaMaquina(ThisWorkbook.Worksheets(ABA_MENU).Shapes(nomeBotao).AlternativeText)
    On Error GoTo 0

    If ws Is Nothing Then
        Alertar "Nao encontrei a aba desta maquina."
        Exit Sub
    End If

    ws.Activate
    ws.Range("A5").Select

End Sub

'==============================================================================
' BUSCA RAPIDA
'==============================================================================
Public Sub BuscarOP()

    Dim texto As String
    Dim wsAchada As Worksheet
    Dim tbl As ListObject
    Dim idx As Long

    texto = Trim$(CStr(ThisWorkbook.Worksheets(ABA_MENU).Range(CELULA_BUSCA).Value))

    If Len(ChaveOP(texto)) = 0 Then
        Alertar "Digite o numero da OP no campo de busca."
        Exit Sub
    End If

    If Not LocalizarOP(texto, wsAchada, tbl, idx) Then
        Alertar "A OP " & FormatarMilhar(texto) & " nao foi encontrada."
        Exit Sub
    End If

    Application.EnableEvents = False
    On Error Resume Next
    wsAchada.Activate
    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
    tbl.ListRows(idx).Range.Select
    ActiveWindow.ScrollRow = IIf(tbl.ListRows(idx).Range.Row > 8, _
                                 tbl.ListRows(idx).Range.Row - 3, 1)
    On Error GoTo 0
    Application.EnableEvents = True

End Sub

'==============================================================================
' NOVA MAQUINA
'==============================================================================
Public Sub NovaMaquina()

    Dim nome As String
    Dim msg As String
    Dim modelo As Worksheet
    Dim ultima As Worksheet
    Dim nova As Worksheet
    Dim tbl As ListObject
    Dim blocos As Collection
    Dim i As Long
    Dim setor As String

    nome = UCase$(Trim$(InputBox("Nome da nova maquina:" & vbNewLine & vbNewLine & _
                                 "Exemplo: MAQUINA1 19", "Nova maquina")))
    If Len(nome) = 0 Then Exit Sub

    If Not NomeMaquinaValido(nome, msg) Then
        Alertar msg
        Exit Sub
    End If

    setor = EscolherDaLista("Em qual setor fica a " & nome & "?", ListarSetores, True, _
                            "Setor da maquina")
    If Len(setor) = 0 Then Exit Sub

    ' Modelo: a primeira maquina existente. Posicao: depois da ultima.
    Set modelo = PrimeiraMaquina()
    Set ultima = UltimaMaquina()
    If modelo Is Nothing Then
        Alertar "Nao ha nenhuma aba de maquina para servir de modelo."
        Exit Sub
    End If

    On Error GoTo TrataErro
    OtimizarInicio

    modelo.Copy After:=ultima
    Set nova = ThisWorkbook.Worksheets(ultima.Index + 1)
    nova.Name = nome
    nova.Range("A1").Value = nome

    ' A copia vem com os blocos da maquina modelo: mantem so um, vazio
    Set blocos = TabelasDaAba(nova)
    For i = blocos.Count To 2 Step -1
        Set tbl = blocos(i)
        nova.Rows(LinhaDoTitulo(tbl) & ":" & _
                  (tbl.Range.Row + tbl.Range.Rows.Count)).Delete Shift:=xlUp
    Next i

    Set tbl = TabelaDaAba(nova)
    If Not tbl.DataBodyRange Is Nothing Then
        For i = tbl.ListRows.Count To 1 Step -1
            tbl.ListRows(i).Delete
        Next i
    End If
    tbl.Name = NomeTabelaLivre(nome)
    EscreverTitulo tbl, MesAtual

    AdicionarNaLista CFG_MAQUINAS, nome
    DefinirSetor nome, setor
    MontarMenu

    OtimizarFim

    nova.Activate
    nova.Range("A5").Select
    Avisar "Maquina " & nome & " criada no setor " & setor & "."
    Exit Sub

TrataErro:
    OtimizarFim
    Alertar "Erro ao criar a maquina: " & Err.Description

End Sub

'==============================================================================
' EXCLUIR MAQUINA
' So exclui maquina SEM OPs, para nunca apagar dados por acidente.
'==============================================================================
Public Sub ExcluirMaquina()

    Dim nome As String
    Dim ws As Worksheet
    Dim qtdOPs As Long
    Dim qtdArquivadas As Long
    Dim aviso As String

    nome = EscolherDaLista("Qual maquina deseja excluir?", ListarMaquinas, False, _
                           "Excluir maquina")
    If Len(nome) = 0 Then Exit Sub

    Set ws = AbaMaquina(nome)

    ' 1. Bloqueia se ainda houver OPs na maquina
    If Not ws Is Nothing Then
        qtdOPs = ContarOPsDaAba(ws, "")
        If qtdOPs > 0 Then
            Alertar "A " & nome & " ainda tem " & qtdOPs & _
                    IIf(qtdOPs = 1, " OP.", " OPs.") & vbNewLine & vbNewLine & _
                    "Arquive as OPs ou mude-as de maquina (EDITAR OP) " & _
                    "antes de excluir."
            Exit Sub
        End If
    End If

    ' 2. Avisa sobre OPs arquivadas que perderiam a origem
    qtdArquivadas = ContarOPsDaAba(ThisWorkbook.Worksheets(ABA_ARQUIVO), nome)

    aviso = "Excluir a maquina " & nome & "?" & vbNewLine & vbNewLine
    If qtdArquivadas > 0 Then
        aviso = aviso & "Atencao: " & qtdArquivadas & _
                IIf(qtdArquivadas = 1, " OP desta maquina esta", _
                                       " OPs desta maquina estao") & _
                " no ARQUIVO e nao poderao mais ser restauradas." & vbNewLine & vbNewLine
    End If
    aviso = aviso & "Esta acao nao pode ser desfeita."

    If Not Confirmar(aviso) Then Exit Sub

    On Error GoTo TrataErro
    OtimizarInicio

    ThisWorkbook.Worksheets(ABA_MENU).Activate
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    RemoverMaquinaConfig nome
    MontarMenu

    OtimizarFim
    Avisar "Maquina " & nome & " excluida."
    Exit Sub

TrataErro:
    Application.DisplayAlerts = True
    OtimizarFim
    Alertar "Erro ao excluir a maquina: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' Conta as OPs de uma aba, somando todos os blocos de mes. Com filtroMaquina,
' conta so as daquela maquina (usado no ARQUIVO, que tem a coluna MAQUINA).
'------------------------------------------------------------------------------
Private Function ContarOPsDaAba(ByVal ws As Worksheet, ByVal filtroMaquina As String) As Long

    Dim tbl As ListObject
    Dim total As Long

    For Each tbl In TabelasDaAba(ws)
        total = total + ContarOPs(tbl, filtroMaquina)
    Next tbl

    ContarOPsDaAba = total

End Function

Private Function ContarOPs(ByVal tbl As ListObject, ByVal filtroMaquina As String) As Long

    Dim i As Long
    Dim total As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Len(Trim$(CStr(tbl.DataBodyRange.Cells(i, COL_OP).Value))) > 0 Then
            If Len(filtroMaquina) = 0 Then
                total = total + 1
            ElseIf StrComp(CStr(tbl.DataBodyRange.Cells(i, COL_MAQUINA).Value), _
                           filtroMaquina, vbTextCompare) = 0 Then
                total = total + 1
            End If
        End If
    Next i

    ContarOPs = total

End Function

'------------------------------------------------------------------------------
' Mostra uma lista numerada e devolve o item escolhido (por numero ou nome).
' Com permitirNovo, aceita tambem um valor que nao esta na lista.
' Devolve "" se o usuario cancelar.
'------------------------------------------------------------------------------
Private Function EscolherDaLista(ByVal pergunta As String, ByVal lista As Collection, _
                                 ByVal permitirNovo As Boolean, _
                                 ByVal titulo As String) As String

    Dim texto As String
    Dim resposta As String
    Dim i As Long
    Dim n As Long

    texto = pergunta & vbNewLine & vbNewLine
    For i = 1 To lista.Count
        texto = texto & "   " & i & " - " & lista(i) & vbNewLine
    Next i

    If permitirNovo Then
        texto = texto & vbNewLine & "Digite o numero, ou o nome de um setor novo."
    Else
        texto = texto & vbNewLine & "Digite o numero ou o nome."
    End If

    resposta = UCase$(Trim$(InputBox(texto, titulo)))
    If Len(resposta) = 0 Then Exit Function

    ' Escolha pelo numero
    If resposta Like "#" Or resposta Like "##" Then
        n = CLng(resposta)
        If n >= 1 And n <= lista.Count Then
            EscolherDaLista = CStr(lista(n))
            Exit Function
        End If
    End If

    ' Escolha pelo nome
    For i = 1 To lista.Count
        If StrComp(CStr(lista(i)), resposta, vbTextCompare) = 0 Then
            EscolherDaLista = CStr(lista(i))
            Exit Function
        End If
    Next i

    If permitirNovo Then
        EscolherDaLista = resposta
    Else
        Alertar "Opcao invalida: " & resposta
    End If

End Function

'------------------------------------------------------------------------------
' Preenche o setor das maquinas que ainda nao tem setor, pelo inicio do nome.
' So age em celulas vazias: nunca sobrescreve um setor ja definido.
'------------------------------------------------------------------------------
Private Sub AtribuirSetoresPadrao()

    Dim cfg As Worksheet
    Dim nome As Variant
    Dim n As String

    Set cfg = ThisWorkbook.Worksheets(ABA_CONFIG)

    If Len(CStr(cfg.Cells(1, CFG_SETOR_MAQUINA).Value)) = 0 Then
        cfg.Cells(1, CFG_SETOR_MAQUINA).Value = "SETOR"
        cfg.Cells(1, CFG_SETOR_MAQUINA).Font.Bold = True
    End If
    If Len(CStr(cfg.Cells(1, CFG_SETORES).Value)) = 0 Then
        cfg.Cells(1, CFG_SETORES).Value = "SETORES (ordem do MENU)"
        cfg.Cells(1, CFG_SETORES).Font.Bold = True
    End If

    If ListarSetores.Count = 0 Then
        AdicionarNaLista CFG_SETORES, "AREA 1"
        AdicionarNaLista CFG_SETORES, "AREA 2"
        AdicionarNaLista CFG_SETORES, "AREA 3"
    End If

    For Each nome In ListarMaquinas
        If Len(SetorDaMaquina(CStr(nome))) = 0 Then
            n = UCase$(CStr(nome))
            If n Like "MAQUINA1*" Then
                DefinirSetor CStr(nome), "AREA 1"
            ElseIf n Like "MAQUINA2*" Or n Like "MAQUINA4*" Then
                DefinirSetor CStr(nome), "AREA 2"
            ElseIf InStr(n, "MAQUINA6") > 0 Then
                DefinirSetor CStr(nome), "AREA 3"
            End If
        End If
    Next nome

End Sub

Private Function NomeMaquinaValido(ByVal nome As String, ByRef msg As String) As Boolean

    Dim proibidos As Variant
    Dim p As Variant

    If Len(nome) > 31 Then
        msg = "O nome pode ter no maximo 31 caracteres."
        Exit Function
    End If

    proibidos = Array("\", "/", "?", "*", "[", "]", ":")
    For Each p In proibidos
        If InStr(nome, p) > 0 Then
            msg = "O nome nao pode conter os caracteres  \ / ? * [ ] :"
            Exit Function
        End If
    Next p

    If nome = ABA_MENU Or nome = ABA_ARQUIVO Or nome = ABA_CONFIG Then
        msg = "Este nome e reservado pelo sistema."
        Exit Function
    End If

    If PlanilhaExiste(nome) Then
        msg = "Ja existe uma aba chamada " & nome & "."
        Exit Function
    End If

    NomeMaquinaValido = True

End Function

Private Function PrimeiraMaquina() As Worksheet

    Dim nome As Variant

    For Each nome In ListarMaquinas
        Set PrimeiraMaquina = AbaMaquina(CStr(nome))
        If Not PrimeiraMaquina Is Nothing Then Exit Function
    Next nome

End Function

Private Function UltimaMaquina() As Worksheet

    Dim nome As Variant
    Dim ws As Worksheet

    For Each nome In ListarMaquinas
        Set ws = AbaMaquina(CStr(nome))
        If Not ws Is Nothing Then
            If UltimaMaquina Is Nothing Then
                Set UltimaMaquina = ws
            ElseIf ws.Index > UltimaMaquina.Index Then
                Set UltimaMaquina = ws
            End If
        End If
    Next nome

End Function

'==============================================================================
' DESENHO DOS BOTOES
'==============================================================================
Public Sub MontarMenu()

    Dim ws As Worksheet
    Dim setores As Collection
    Dim setor As Variant
    Dim maquinas As Collection
    Dim i As Long
    Dim cont As Long
    Dim x0 As Single, y As Single
    Dim shp As Shape

    Set ws = ThisWorkbook.Worksheets(ABA_MENU)
    ApagarBotoes ws

    ' A area das maquinas e desenhada pelo sistema: limpa textos antigos
    ws.Range("A11:Z500").ClearContents

    ' Linha de acoes
    x0 = ws.Range("B5").Left
    y = ws.Range("B5").Top
    CriarBotao ws, "novaop", "+ NOVA OP", x0, y, 130, 30, "BotaoNovaOP", True
    CriarBotao ws, "novamaq", "NOVA MAQUINA", x0 + 140, y, 130, 30, "NovaMaquina", False
    CriarBotao ws, "excmaq", "EXCLUIR MAQUINA", x0 + 280, y, 130, 30, "ExcluirMaquina", False
    CriarBotao ws, "arquivo", "ARQUIVO", x0 + 420, y, 130, 30, "IrParaArquivo", False
    CriarBotao ws, "importar", "IMPORTAR DA PROGRAMACAO", x0 + 560, y, 190, 30, _
               "ImportarDaProgramacao", True

    ' Busca
    CriarBotao ws, "buscar", "BUSCAR", ws.Range("D9").Left + 4, ws.Range("B9").Top, _
               80, ws.Range("B9").Height, "BuscarOP", True

    ' Setores da lista, na ordem; no fim, as maquinas sem setor
    Set setores = New Collection
    For Each setor In ListarSetores
        setores.Add CStr(setor)
    Next setor
    setores.Add SETOR_SEM

    y = ws.Range("B12").Top

    For Each setor In setores

        Set maquinas = MaquinasDoSetor(CStr(setor))

        If maquinas.Count > 0 Then

            CriarTitulo ws, "sec_" & cont, CStr(setor), x0, y
            y = y + 26

            ' Botoes em grade de 4 colunas
            For i = 0 To maquinas.Count - 1
                Set shp = CriarBotao(ws, "maq_" & cont & "_" & i, CStr(maquinas(i + 1)), _
                                     x0 + (i Mod 4) * 140, y + (i \ 4) * 40, _
                                     130, 32, "IrParaMaquina", False)
                shp.AlternativeText = CStr(maquinas(i + 1))
            Next i

            y = y + ((maquinas.Count - 1) \ 4 + 1) * 40 + 22

        End If

        cont = cont + 1

    Next setor

End Sub

'------------------------------------------------------------------------------
' Maquinas de um setor, na ordem da _CONFIG.
' SETOR_SEM reune as maquinas sem setor ou com setor fora da lista.
'------------------------------------------------------------------------------
Private Function MaquinasDoSetor(ByVal setor As String) As Collection

    Dim res As Collection
    Dim nome As Variant
    Dim s As String

    Set res = New Collection

    For Each nome In ListarMaquinas
        s = SetorDaMaquina(CStr(nome))
        If setor = SETOR_SEM Then
            If Len(s) = 0 Or Not SetorNaLista(s) Then res.Add CStr(nome)
        ElseIf StrComp(s, setor, vbTextCompare) = 0 Then
            res.Add CStr(nome)
        End If
    Next nome

    Set MaquinasDoSetor = res

End Function

Private Function SetorNaLista(ByVal setor As String) As Boolean

    Dim v As Variant

    For Each v In ListarSetores
        If StrComp(CStr(v), setor, vbTextCompare) = 0 Then
            SetorNaLista = True
            Exit Function
        End If
    Next v

End Function

'------------------------------------------------------------------------------
' Titulo de setor no MENU (caixa de texto sem borda nem fundo).
'------------------------------------------------------------------------------
Private Sub CriarTitulo(ByVal ws As Worksheet, ByVal id As String, _
                        ByVal texto As String, ByVal x As Single, ByVal y As Single)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, 400, 22)

    With shp
        .Name = PREFIXO_BOTAO & id
        .Placement = xlFreeFloating
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 2: .MarginTop = 0: .MarginBottom = 0
            With .TextRange
                .Text = texto
                .Font.Name = "Segoe UI"
                .Font.Size = 13
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = COR_BOTAO
            End With
        End With
    End With

End Sub

Public Sub MontarBotoesMaquina(ByVal ws As Worksheet)

    Dim x As Single, y As Single

    ApagarBotoes ws

    x = ws.Range("A2").Left + 2
    y = ws.Range("A2").Top + 2

    CriarBotao ws, "menu", "< MENU", x, y, 70, 26, "IrParaMenu", False
    CriarBotao ws, "nova", "+ NOVA OP", x + 76, y, 90, 26, "BotaoNovaOP", True
    CriarBotao ws, "editar", "EDITAR OP", x + 172, y, 90, 26, "BotaoEditarOP", False
    CriarBotao ws, "excluir", "EXCLUIR OP", x + 268, y, 90, 26, "BotaoExcluirOP", False
    CriarBotao ws, "arquivar", "ARQUIVAR OP", x + 364, y, 96, 26, "BotaoArquivarOP", False

End Sub

Public Sub MontarBotoesArquivo(ByVal ws As Worksheet)

    Dim x As Single, y As Single

    ApagarBotoes ws

    x = ws.Range("A2").Left + 2
    y = ws.Range("A2").Top + 2

    CriarBotao ws, "menu", "< MENU", x, y, 70, 26, "IrParaMenu", False
    CriarBotao ws, "restaurar", "RESTAURAR OP", x + 76, y, 110, 26, "BotaoRestaurarOP", True

End Sub

Private Sub ApagarBotoes(ByVal ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PREFIXO_BOTAO)) = PREFIXO_BOTAO Then
            ws.Shapes(i).Delete
        End If
    Next i

End Sub

Private Function CriarBotao(ByVal ws As Worksheet, ByVal id As String, _
                            ByVal texto As String, _
                            ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single, _
                            ByVal macro As String, ByVal principal As Boolean) As Shape

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)

    With shp
        .Name = PREFIXO_BOTAO & id
        .OnAction = macro
        .Placement = xlFreeFloating          ' nao se move com filtros
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = COR_BOTAO
        .Line.Weight = 0.75
        .Adjustments.item(1) = 0.12

        If principal Then
            .Fill.ForeColor.RGB = COR_BOTAO
        Else
            .Fill.ForeColor.RGB = COR_BOTAO_CLARO
        End If

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2
            With .TextRange
                .Text = texto
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                .Font.Bold = msoTrue
                If principal Then
                    .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    .Font.Fill.ForeColor.RGB = COR_BOTAO
                End If
            End With
        End With
    End With

    Set CriarBotao = shp

End Function

