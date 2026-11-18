Attribute VB_Name = "mdlUtil"
'==============================================================================
' mdlUtil - Funcoes auxiliares reutilizaveis
' Controle de Ordens por Maquina
'==============================================================================
Option Explicit

Private Const NOMES_MESES As String = "JANEIRO|FEVEREIRO|MARCO|ABRIL|MAIO|JUNHO|" & _
                                      "JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO"


'------------------------------------------------------------------------------
' Desliga atualizacao de tela e eventos durante operacoes em lote.
'------------------------------------------------------------------------------
Public Sub OtimizarInicio()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Public Sub OtimizarFim()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
End Sub

'------------------------------------------------------------------------------
' Mantem apenas os digitos de um texto. "62.620" -> "62620"
'------------------------------------------------------------------------------
Public Function SomenteDigitos(ByVal texto As String) As String

    Dim i As Long
    Dim ch As String
    Dim buf As String

    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next i

    SomenteDigitos = buf

End Function

'------------------------------------------------------------------------------
' Verdadeiro se o texto contem apenas digitos, pontos e espacos
' (e ao menos um digito). Usado para validar OP e quantidade.
'------------------------------------------------------------------------------
Public Function TextoNumericoValido(ByVal texto As String) As Boolean

    Dim i As Long
    Dim ch As String

    texto = Trim$(texto)
    If Len(SomenteDigitos(texto)) = 0 Then Exit Function

    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or ch = "." Or ch = " ") Then Exit Function
    Next i

    TextoNumericoValido = True

End Function

'------------------------------------------------------------------------------
' Chave de comparacao da OP: so digitos, sem zeros a esquerda.
' "62.620", "62620" e " 062620 " geram a mesma chave.
'------------------------------------------------------------------------------
Public Function ChaveOP(ByVal texto As String) As String

    Dim d As String

    d = SomenteDigitos(texto)
    Do While Len(d) > 1 And Left$(d, 1) = "0"
        d = Mid$(d, 2)
    Loop

    ChaveOP = d

End Function

'------------------------------------------------------------------------------
' Formata digitos com ponto de milhar, sem converter em numero.
' "62620" -> "62.620"
'------------------------------------------------------------------------------
Public Function FormatarMilhar(ByVal texto As String) As String

    Dim d As String
    Dim buf As String
    Dim i As Long
    Dim cont As Long

    d = ChaveOP(texto)

    For i = Len(d) To 1 Step -1
        buf = Mid$(d, i, 1) & buf
        cont = cont + 1
        If cont Mod 3 = 0 And i > 1 Then buf = "." & buf
    Next i

    FormatarMilhar = buf

End Function

'------------------------------------------------------------------------------
' Valor numerico da OP, usado apenas para ordenar.
'------------------------------------------------------------------------------
Public Function ValorOP(ByVal texto As String) As Double

    Dim d As String

    d = ChaveOP(texto)
    If Len(d) > 0 Then ValorOP = CDbl(d)

End Function

'------------------------------------------------------------------------------
' Mensagens padronizadas.
'------------------------------------------------------------------------------
Public Sub Avisar(ByVal mensagem As String)
    MsgBox mensagem, vbInformation, "Controle de Ordens"
End Sub

Public Sub Alertar(ByVal mensagem As String)
    MsgBox mensagem, vbExclamation, "Controle de Ordens"
End Sub

Public Function Confirmar(ByVal mensagem As String) As Boolean
    Confirmar = (MsgBox(mensagem, vbQuestion + vbYesNo + vbDefaultButton2, _
                        "Controle de Ordens") = vbYes)
End Function

'------------------------------------------------------------------------------
' Verifica se existe uma aba com esse nome.
'------------------------------------------------------------------------------
Public Function PlanilhaExiste(ByVal nome As String) As Boolean

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nome, vbTextCompare) = 0 Then
            PlanilhaExiste = True
            Exit Function
        End If
    Next ws

End Function

'------------------------------------------------------------------------------
' Ordena uma colecao (numeros ou textos). Volume sempre pequeno.
'------------------------------------------------------------------------------
Public Function OrdenarColecao(ByVal origem As Collection) As Collection

    Dim res As Collection
    Dim i As Long, j As Long
    Dim inserido As Boolean

    Set res = New Collection

    For i = 1 To origem.Count
        inserido = False
        For j = 1 To res.Count
            If MenorQue(origem(i), res(j)) Then
                res.Add origem(i), , j
                inserido = True
                Exit For
            End If
        Next j
        If Not inserido Then res.Add origem(i)
    Next i

    Set OrdenarColecao = res

End Function

Private Function MenorQue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNumeric(a) And IsNumeric(b) Then
        MenorQue = (CDbl(a) < CDbl(b))
    Else
        MenorQue = (StrComp(CStr(a), CStr(b), vbTextCompare) < 0)
    End If
End Function

'==============================================================================
' MESES
' O mes e sempre guardado no formato "OUTUBRO/2026".
'==============================================================================
'------------------------------------------------------------------------------
' Nome do mes pelo numero (1 a 12).
'------------------------------------------------------------------------------
Public Function NomeDoMes(ByVal numero As Long) As String

    Dim nomes() As String

    If numero < 1 Or numero > 12 Then Exit Function
    nomes = Split(NOMES_MESES, "|")
    NomeDoMes = nomes(numero - 1)

End Function

'------------------------------------------------------------------------------
' Numero do mes pelo nome ou abreviacao ("OUT", "OUTUBRO"). 0 se nao reconhecer.
'------------------------------------------------------------------------------
Public Function NumeroDoMes(ByVal nome As String) As Long

    Dim nomes() As String
    Dim i As Long
    Dim alvo As String

    alvo = UCase$(Trim$(nome))
    alvo = Replace(alvo, ChrW(199), "C")      ' cedilha de MARCO
    If Len(alvo) < 3 Then Exit Function

    nomes = Split(NOMES_MESES, "|")
    For i = 0 To UBound(nomes)
        If Left$(nomes(i), 3) = Left$(alvo, 3) Then
            NumeroDoMes = i + 1
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' Mes de uma data, no formato do sistema.
'------------------------------------------------------------------------------
Public Function MesDaData(ByVal quando As Date) As String
    MesDaData = NomeDoMes(Month(quando)) & "/" & Year(quando)
End Function

Public Function MesAtual() As String
    MesAtual = MesDaData(DiaDeTurno(Now))
End Function

'------------------------------------------------------------------------------
' Dia de producao da fabrica: o dia comeca as 6h e vai ate as 5h59 do dia
' seguinte. 01/10 as 03:00 ainda e dia 30/09.
'------------------------------------------------------------------------------
Public Function DiaDeTurno(ByVal quando As Date) As Date
    DiaDeTurno = Int(CDbl(quando) - CDbl(TimeSerial(6, 0, 0)))
End Function

'------------------------------------------------------------------------------
' Padroniza o que o usuario digitou. Aceita "10/2026", "out/26", "OUTUBRO 2026"
' e ate so "outubro" (assume o ano atual). Devolve "" se nao entender.
'------------------------------------------------------------------------------
Public Function NormalizarMes(ByVal texto As String) As String

    Dim t As String
    Dim partes() As String
    Dim numero As Long
    Dim ano As Long

    t = UCase$(Trim$(texto))
    If Len(t) = 0 Then Exit Function

    t = Replace(t, "-", "/")
    t = Replace(t, ".", "/")
    t = Replace(t, " DE ", "/")
    t = Replace(t, " ", "/")
    Do While InStr(t, "//") > 0
        t = Replace(t, "//", "/")
    Loop

    partes = Split(t, "/")

    ' Mes
    If IsNumeric(partes(0)) Then
        numero = CLng(partes(0))
    Else
        numero = NumeroDoMes(partes(0))
    End If
    If numero < 1 Or numero > 12 Then Exit Function

    ' Ano
    If UBound(partes) >= 1 Then
        If Not IsNumeric(partes(1)) Then Exit Function
        ano = CLng(partes(1))
        If ano < 100 Then ano = 2000 + ano
        If ano < 2000 Or ano > 2199 Then Exit Function
    Else
        ano = Year(Date)
    End If

    NormalizarMes = NomeDoMes(numero) & "/" & ano

End Function

'------------------------------------------------------------------------------
' Chave de ordenacao do mes (ano * 12 + numero). 0 se o mes for invalido.
'------------------------------------------------------------------------------
Public Function ChaveMes(ByVal mes As String) As Double

    Dim t As String
    Dim partes() As String

    t = NormalizarMes(mes)
    If Len(t) = 0 Then Exit Function

    partes = Split(t, "/")
    ChaveMes = CDbl(partes(1)) * 12 + NumeroDoMes(partes(0))

End Function

'------------------------------------------------------------------------------
' Meses sugeridos no formulario: alguns antes e alguns depois do mes atual.
'------------------------------------------------------------------------------
Public Function ListaMeses(ByVal antes As Long, ByVal depois As Long) As Collection

    Dim res As Collection
    Dim i As Long

    Set res = New Collection
    For i = -antes To depois
        res.Add MesDaData(DateAdd("m", i, Date))
    Next i

    Set ListaMeses = res

End Function

