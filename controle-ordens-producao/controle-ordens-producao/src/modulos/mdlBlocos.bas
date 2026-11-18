Attribute VB_Name = "mdlBlocos"
'==============================================================================
' mdlBlocos - Blocos de mes dentro da aba de cada maquina
' Controle de Ordens por Maquina
'
' Cada aba de maquina tem um bloco por mes:
'
'     MES: OUTUBRO/2026          <- titulo (uma celula, na coluna A)
'     OP | ITEM | QTD | ...      <- cabecalho da tabela daquele mes
'     62.920 ...
'
'     MES: NOVEMBRO/2026
'     OP | ITEM | QTD | ...
'
' Os blocos ficam em ordem cronologica: o mes mais antigo em cima. Dentro de
' cada bloco, as OPs continuam em ordem numerica.
'==============================================================================
Option Explicit

Private Const LINHAS_NOVO_BLOCO As Long = 4    ' titulo + cabecalho + 1 linha + branco

'------------------------------------------------------------------------------
' Prefixo do titulo. Montado com ChrW para o acento nunca se perder.
'------------------------------------------------------------------------------
Public Function PrefixoMes() As String
    PrefixoMes = "M" & ChrW(202) & "S: "       ' "MES: " com E maiusculo acentuado
End Function

'==============================================================================
' LEITURA
'==============================================================================

'------------------------------------------------------------------------------
' Todas as tabelas (blocos) de uma aba, de cima para baixo.
'------------------------------------------------------------------------------
Public Function TabelasDaAba(ByVal ws As Worksheet) As Collection

    Dim res As Collection
    Dim lo As ListObject
    Dim i As Long
    Dim inserido As Boolean

    Set res = New Collection

    For Each lo In ws.ListObjects
        inserido = False
        For i = 1 To res.Count
            If lo.Range.Row < res(i).Range.Row Then
                res.Add lo, , i
                inserido = True
                Exit For
            End If
        Next i
        If Not inserido Then res.Add lo
    Next lo

    Set TabelasDaAba = res

End Function

'------------------------------------------------------------------------------
' Tabela que contem a celula informada (Nothing se a celula estiver fora).
'------------------------------------------------------------------------------
Public Function TabelaDaCelula(ByVal ws As Worksheet, ByVal cel As Range) As ListObject

    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If Not Intersect(cel, lo.Range) Is Nothing Then
            Set TabelaDaCelula = lo
            Exit Function
        End If
    Next lo

End Function

'------------------------------------------------------------------------------
' Linha do titulo de um bloco (fica sempre logo acima do cabecalho).
'------------------------------------------------------------------------------
Public Function LinhaDoTitulo(ByVal tbl As ListObject) As Long
    LinhaDoTitulo = tbl.Range.Row - 1
End Function

'------------------------------------------------------------------------------
' Mes de um bloco, lido do titulo. "" se o bloco nao tiver titulo.
'------------------------------------------------------------------------------
Public Function MesDaTabela(ByVal tbl As ListObject) As String

    Dim texto As String

    If tbl Is Nothing Then Exit Function
    If tbl.Range.Row < 2 Then Exit Function

    texto = Trim$(CStr(tbl.Parent.Cells(LinhaDoTitulo(tbl), tbl.Range.Column).Value))
    If InStr(1, texto, PrefixoMes, vbTextCompare) = 1 Then
        MesDaTabela = NormalizarMes(Mid$(texto, Len(PrefixoMes) + 1))
    End If

End Function

'------------------------------------------------------------------------------
' Escreve o titulo do bloco.
'------------------------------------------------------------------------------
Public Sub EscreverTitulo(ByVal tbl As ListObject, ByVal mes As String)

    Dim cel As Range

    Set cel = tbl.Parent.Cells(LinhaDoTitulo(tbl), tbl.Range.Column)

    With cel
        .NumberFormat = "@"
        .Value = PrefixoMes & mes
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(31, 31, 31)
    End With

    tbl.Parent.Rows(LinhaDoTitulo(tbl)).RowHeight = 26

End Sub

'==============================================================================
' LOCALIZACAO E CRIACAO DE BLOCOS
'==============================================================================

'------------------------------------------------------------------------------
' Bloco de um mes dentro da aba. Com criarSeFaltar, cria o bloco quando aquele
' mes ainda nao existe na maquina.
'------------------------------------------------------------------------------
Public Function TabelaDoMes(ByVal ws As Worksheet, ByVal mes As String, _
                            ByVal criarSeFaltar As Boolean) As ListObject

    Dim tbl As ListObject
    Dim alvo As String

    alvo = NormalizarMes(mes)
    If Len(alvo) = 0 Then Exit Function

    For Each tbl In TabelasDaAba(ws)
        If MesDaTabela(tbl) = alvo Then
            Set TabelaDoMes = tbl
            Exit Function
        End If
    Next tbl

    If criarSeFaltar Then Set TabelaDoMes = CriarBloco(ws, alvo)

End Function

'------------------------------------------------------------------------------
' Cria um bloco novo, na posicao cronologica correta.
'------------------------------------------------------------------------------
Private Function CriarBloco(ByVal ws As Worksheet, ByVal mes As String) As ListObject

    Dim modelo As ListObject
    Dim tbl As ListObject
    Dim novo As ListObject
    Dim linhaTitulo As Long
    Dim colunas As Long
    Dim col As Long
    Dim chave As Double

    Set modelo = TabelaDaAba(ws)
    If modelo Is Nothing Then Exit Function

    colunas = modelo.ListColumns.Count
    col = modelo.Range.Column
    chave = ChaveMes(mes)

    ' Onde o bloco entra: antes do primeiro bloco de mes posterior
    For Each tbl In TabelasDaAba(ws)
        If ChaveMes(MesDaTabela(tbl)) > chave Then
            linhaTitulo = LinhaDoTitulo(tbl)
            Exit For
        End If
    Next tbl

    If linhaTitulo > 0 Then
        ws.Rows(linhaTitulo & ":" & (linhaTitulo + LINHAS_NOVO_BLOCO - 1)).Insert Shift:=xlDown
    Else
        linhaTitulo = UltimaLinhaUsada(ws) + 2
    End If

    ' Cabecalho copiado de um bloco existente (traz toda a formatacao)
    modelo.HeaderRowRange.Copy ws.Cells(linhaTitulo + 1, col)
    Application.CutCopyMode = False
    ws.Range(ws.Cells(linhaTitulo + 2, col), _
             ws.Cells(linhaTitulo + 2, col + colunas - 1)).ClearContents

    Set novo = ws.ListObjects.Add(xlSrcRange, _
               ws.Range(ws.Cells(linhaTitulo + 1, col), _
                        ws.Cells(linhaTitulo + 2, col + colunas - 1)), , xlYes)
    novo.Name = NomeTabelaLivre(ws.Name & "_" & mes)
    novo.TableStyle = modelo.TableStyle

    EscreverTitulo novo, mes

    Set CriarBloco = novo

End Function

'------------------------------------------------------------------------------
' Ultima linha com qualquer conteudo na aba.
'------------------------------------------------------------------------------
Public Function UltimaLinhaUsada(ByVal ws As Worksheet) As Long

    Dim ultima As Range

    On Error Resume Next
    Set ultima = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                               SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If ultima Is Nothing Then
        UltimaLinhaUsada = 4
    Else
        UltimaLinhaUsada = ultima.Row
    End If

End Function

'==============================================================================
' MIGRACAO
' Prepara as abas que foram criadas antes dos blocos de mes:
'   - poe um titulo em cada tabela que ainda nao tem
'   - acrescenta a coluna MES no ARQUIVO
'   - solta o congelamento de paineis (com varios blocos ele atrapalha)
'==============================================================================
Public Sub MigrarBlocos()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblArq As ListObject
    Dim cab As Range

    For Each ws In ThisWorkbook.Worksheets
        If EhAbaMaquina(ws) Then

            For Each tbl In TabelasDaAba(ws)
                If Len(MesDaTabela(tbl)) = 0 Then
                    ws.Rows(tbl.Range.Row).Insert Shift:=xlDown
                    EscreverTitulo tbl, MesAtual
                End If
            Next tbl

            On Error Resume Next
            ws.Activate
            ActiveWindow.FreezePanes = False
            On Error GoTo 0

        End If
    Next ws

    ' A coluna DATA foi abandonada: se existir de uma versao anterior, sai
    For Each ws In ThisWorkbook.Worksheets
        If EhAbaMaquina(ws) Or ws.Name = ABA_ARQUIVO Then
            For Each tbl In TabelasDaAba(ws)
                RemoverColuna tbl, "DATA"
            Next tbl
        End If
    Next ws

    ' Coluna MES no ARQUIVO
    Set tblArq = TabelaDaAba(ThisWorkbook.Worksheets(ABA_ARQUIVO))
    GarantirColuna tblArq, COL_MES, "M" & ChrW(202) & "S", 16

End Sub

'------------------------------------------------------------------------------
' Remove uma coluna pelo titulo, se ela existir.
'------------------------------------------------------------------------------
Private Sub RemoverColuna(ByVal tbl As ListObject, ByVal titulo As String)

    Dim coluna As Long

    If tbl Is Nothing Then Exit Sub

    coluna = ColunaPorTitulo(tbl, titulo)
    If coluna > 0 Then tbl.ListColumns(coluna).Delete

End Sub

'------------------------------------------------------------------------------
' Garante que a tabela tenha uma coluna com aquele titulo, na posicao pedida.
' Se ja existir (em qualquer posicao), nao faz nada.
'------------------------------------------------------------------------------
Private Sub GarantirColuna(ByVal tbl As ListObject, ByVal posicao As Long, _
                           ByVal titulo As String, ByVal largura As Double)

    Dim cab As Range

    If tbl Is Nothing Then Exit Sub
    If ColunaPorTitulo(tbl, titulo) > 0 Then Exit Sub
    If tbl.ListColumns.Count < posicao - 1 Then Exit Sub

    tbl.ListColumns.Add Position:=posicao
    Set cab = tbl.HeaderRowRange.Cells(1, posicao)

    ' Formatacao copiada do cabecalho vizinho
    tbl.HeaderRowRange.Cells(1, posicao - 1).Copy
    cab.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    cab.Value = titulo
    tbl.Parent.Columns(tbl.Range.Column + posicao - 1).ColumnWidth = largura

End Sub
