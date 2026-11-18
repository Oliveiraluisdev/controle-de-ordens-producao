Attribute VB_Name = "mdlConfig"
'==============================================================================
' mdlConfig - Constantes do sistema e leitura da aba _CONFIG
' Controle de Ordens por Maquina
'
' A lista de maquinas e as sugestoes do formulario (itens, clientes e
' quantidades) ficam na aba oculta _CONFIG. Nada disso esta fixo no codigo.
'==============================================================================
Option Explicit

'--- Abas especiais -----------------------------------------------------------
Public Const ABA_MENU As String = "MENU"
Public Const ABA_ARQUIVO As String = "ARQUIVO"
Public Const ABA_CONFIG As String = "_CONFIG"

'--- Celula de busca rapida no MENU (mesclada B9:C9) --------------------------
Public Const CELULA_BUSCA As String = "B9"

'--- Colunas da tabela de OPs -------------------------------------------------
Public Const COL_OP As Long = 1
Public Const COL_ITEM As Long = 2
Public Const COL_QTD As Long = 3
Public Const COL_CLIENTE As Long = 4
Public Const COL_TAREFA1 As Long = 5        ' PW
Public Const QTD_TAREFAS As Long = 4        ' PW, IMPRESSAO, RELATORIO, SISTEMA
Public Const COL_CADASTRO As Long = 9
Public Const COL_MAQUINA As Long = 10       ' existe somente no ARQUIVO
Public Const COL_MES As Long = 11           ' existe somente no ARQUIVO

'--- Colunas da aba _CONFIG ---------------------------------------------------
Public Const CFG_MAQUINAS As Long = 1        ' coluna A - nome da maquina
Public Const CFG_SETOR_MAQUINA As Long = 2   ' coluna B - setor de cada maquina
Public Const CFG_SETORES As Long = 3         ' coluna C - setores, na ordem do MENU
Public Const CFG_ITENS As Long = 5           ' coluna E
Public Const CFG_CLIENTES As Long = 7        ' coluna G
Public Const CFG_QUANTIDADES As Long = 9     ' coluna I

'--- Setor usado para maquinas sem setor definido -----------------------------
Public Const SETOR_SEM As String = "OUTRAS"

'--- Aparencia ----------------------------------------------------------------
Public Const COR_CONCLUIDA As Long = 14540253    ' RGB(221, 221, 221) - linha completa
Public Const COR_TEXTO_CONCLUIDA As Long = 5921370  ' RGB(90, 90, 90)
Public Const COR_TEXTO_CINZA As Long = 8421504      ' RGB(128, 128, 128)

'------------------------------------------------------------------------------
' Caracteres do checklist.
'------------------------------------------------------------------------------
Public Function CaixaVazia() As String
    CaixaVazia = ChrW(9744)      ' caixa vazia
End Function

Public Function CaixaMarcada() As String
    CaixaMarcada = ChrW(9745)    ' caixa marcada
End Function

'------------------------------------------------------------------------------
' Le uma coluna da _CONFIG (a partir da linha 2) e devolve os valores.
'------------------------------------------------------------------------------
Public Function ListaConfig(ByVal coluna As Long) As Collection

    Dim ws As Worksheet
    Dim res As Collection
    Dim ultima As Long
    Dim i As Long
    Dim v As Variant

    Set res = New Collection
    Set ws = ThisWorkbook.Worksheets(ABA_CONFIG)

    ultima = ws.Cells(ws.Rows.Count, coluna).End(xlUp).Row

    For i = 2 To ultima
        v = ws.Cells(i, coluna).Value
        If Len(Trim$(CStr(v))) > 0 Then res.Add v
    Next i

    Set ListaConfig = res

End Function

'------------------------------------------------------------------------------
' Acrescenta um valor a uma lista da _CONFIG, se ainda nao existir.
'------------------------------------------------------------------------------
Public Sub AdicionarNaLista(ByVal coluna As Long, ByVal valor As Variant)

    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long
    Dim alvo As String

    alvo = Trim$(CStr(valor))
    If Len(alvo) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(ABA_CONFIG)
    ultima = ws.Cells(ws.Rows.Count, coluna).End(xlUp).Row

    For i = 2 To ultima
        If StrComp(Trim$(CStr(ws.Cells(i, coluna).Value)), alvo, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next i

    If ultima < 1 Then ultima = 1

    If IsNumeric(valor) And coluna = CFG_QUANTIDADES Then
        ws.Cells(ultima + 1, coluna).Value = CDbl(valor)
    Else
        ws.Cells(ultima + 1, coluna).NumberFormat = "@"
        ws.Cells(ultima + 1, coluna).Value = alvo
    End If

End Sub

'------------------------------------------------------------------------------
' Lista de maquinas, na ordem da _CONFIG.
'------------------------------------------------------------------------------
Public Function ListarMaquinas() As Collection
    Set ListarMaquinas = ListaConfig(CFG_MAQUINAS)
End Function

'------------------------------------------------------------------------------
' Setores, na ordem em que aparecem no MENU.
'------------------------------------------------------------------------------
Public Function ListarSetores() As Collection
    Set ListarSetores = ListaConfig(CFG_SETORES)
End Function

'------------------------------------------------------------------------------
' Linha da maquina na _CONFIG (0 se nao estiver cadastrada).
'------------------------------------------------------------------------------
Public Function LinhaMaquinaConfig(ByVal nomeMaquina As String) As Long

    Dim ws As Worksheet
    Dim ultima As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(ABA_CONFIG)
    ultima = ws.Cells(ws.Rows.Count, CFG_MAQUINAS).End(xlUp).Row

    For i = 2 To ultima
        If StrComp(Trim$(CStr(ws.Cells(i, CFG_MAQUINAS).Value)), _
                   Trim$(nomeMaquina), vbTextCompare) = 0 Then
            LinhaMaquinaConfig = i
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' Setor de uma maquina ("" se nao tiver).
'------------------------------------------------------------------------------
Public Function SetorDaMaquina(ByVal nomeMaquina As String) As String

    Dim lin As Long

    lin = LinhaMaquinaConfig(nomeMaquina)
    If lin > 0 Then
        SetorDaMaquina = UCase$(Trim$(CStr( _
            ThisWorkbook.Worksheets(ABA_CONFIG).Cells(lin, CFG_SETOR_MAQUINA).Value)))
    End If

End Function

'------------------------------------------------------------------------------
' Grava o setor de uma maquina e inclui o setor na lista, se for novo.
'------------------------------------------------------------------------------
Public Sub DefinirSetor(ByVal nomeMaquina As String, ByVal setor As String)

    Dim lin As Long

    setor = UCase$(Trim$(setor))
    lin = LinhaMaquinaConfig(nomeMaquina)
    If lin = 0 Or Len(setor) = 0 Then Exit Sub

    With ThisWorkbook.Worksheets(ABA_CONFIG).Cells(lin, CFG_SETOR_MAQUINA)
        .NumberFormat = "@"
        .Value = setor
    End With

    AdicionarNaLista CFG_SETORES, setor

End Sub

'------------------------------------------------------------------------------
' Remove a maquina (nome e setor) da _CONFIG.
'------------------------------------------------------------------------------
Public Sub RemoverMaquinaConfig(ByVal nomeMaquina As String)

    Dim ws As Worksheet
    Dim lin As Long

    lin = LinhaMaquinaConfig(nomeMaquina)
    If lin = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(ABA_CONFIG)
    ws.Range(ws.Cells(lin, CFG_MAQUINAS), ws.Cells(lin, CFG_SETOR_MAQUINA)).Delete Shift:=xlUp

End Sub

'------------------------------------------------------------------------------
' A aba informada e de uma maquina?
'------------------------------------------------------------------------------
Public Function EhAbaMaquina(ByVal ws As Object) As Boolean

    Dim nome As Variant

    If TypeName(ws) <> "Worksheet" Then Exit Function

    For Each nome In ListarMaquinas
        If StrComp(ws.Name, CStr(nome), vbTextCompare) = 0 Then
            EhAbaMaquina = True
            Exit Function
        End If
    Next nome

End Function

'------------------------------------------------------------------------------
' Aba da maquina pelo nome, ou Nothing se nao existir.
'------------------------------------------------------------------------------
Public Function AbaMaquina(ByVal nomeMaquina As String) As Worksheet

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, Trim$(nomeMaquina), vbTextCompare) = 0 Then
            If EhAbaMaquina(ws) Then Set AbaMaquina = ws
            Exit Function
        End If
    Next ws

End Function

'------------------------------------------------------------------------------
' Primeira tabela da aba. No ARQUIVO existe so uma; nas maquinas, use as
' funcoes de mdlBlocos, porque cada mes tem a sua.
'------------------------------------------------------------------------------
Public Function TabelaDaAba(ByVal ws As Worksheet) As ListObject
    If ws.ListObjects.Count > 0 Then Set TabelaDaAba = ws.ListObjects(1)
End Function

'------------------------------------------------------------------------------
' Nome de tabela ainda nao usado na pasta de trabalho.
'------------------------------------------------------------------------------
Public Function NomeTabelaLivre(ByVal base As String) As String

    Dim limpo As String
    Dim candidato As String
    Dim ch As String
    Dim i As Long
    Dim n As Long

    limpo = "tbl"
    For i = 1 To Len(base)
        ch = Mid$(base, i, 1)
        If ch Like "[A-Za-z0-9]" Then limpo = limpo & ch Else limpo = limpo & "_"
    Next i

    candidato = limpo
    Do While TabelaExiste(candidato)
        n = n + 1
        candidato = limpo & "_" & n
    Loop

    NomeTabelaLivre = candidato

End Function

Public Function ColunaPorTitulo(ByVal tbl As ListObject, ByVal titulo As String) As Long

    Dim i As Long

    If tbl Is Nothing Then Exit Function

    For i = 1 To tbl.ListColumns.Count
        If StrComp(Trim$(CStr(tbl.HeaderRowRange.Cells(1, i).Value)), _
                   Trim$(titulo), vbTextCompare) = 0 Then
            ColunaPorTitulo = i
            Exit Function
        End If
    Next i

End Function

Public Function TabelaExiste(ByVal nome As String) As Boolean

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nome, vbTextCompare) = 0 Then
                TabelaExiste = True
                Exit Function
            End If
        Next lo
    Next ws

End Function

