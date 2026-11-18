Attribute VB_Name = "mdlImportar"
'==============================================================================
' mdlImportar - Importacao das OPs a partir da Programacao Mensal
' Controle de Ordens por Maquina
'
' COMO FUNCIONA, PASSO A PASSO
'   1. Voce escolhe o arquivo da programacao (pode estar no OneDrive).
'   2. O sistema faz uma COPIA dele na pasta temporaria do Windows.
'   3. Abre a copia, escondida e somente leitura, sem atualizar vinculos.
'   4. Le as OPs da aba (mes) escolhida.
'   5. Fecha a copia e APAGA o arquivo temporario - inclusive se der erro.
'
' O arquivo original nunca e aberto nem alterado, e nenhum vinculo fica
' guardado na planilha do Controle.
'
' LAYOUT ESPERADO DA PROGRAMACAO
'   Uma aba por mes. Dentro da aba, blocos:
'     linha com o nome da maquina sozinho na coluna A  (ex.: MAQUINA1 1)
'     linha de cabecalho: CLIENTE | CAV. | PRODUTO | OP | LOTE | QUANT...
'     uma linha por OP, ate a primeira linha em branco
'==============================================================================
Option Explicit

'--- Titulos procurados no cabecalho de cada bloco ----------------------------
' As programacoes tem colunas em ordens diferentes (Embalagem, RPC, TPA...),
' por isso as colunas sao descobertas pelo cabecalho, e nao fixas.
Private Const TIT_CLIENTE As String = "CLIENTE"
Private Const TIT_PRODUTO As String = "PRODUTO"
Private Const TIT_OP As String = "OP"
Private Const TIT_QTD As String = "QUANT"
Private Const TIT_DATA As String = "DATA"
Private Const TIT_INICIAL As String = "INICIAL"

Private Const ULTIMA_COLUNA As Long = 20

' Linhas de texto solto que nao sao nome de maquina
Private Const LIXO As String = "PROGRAMA|DATA DA ATUALIZA|ESTQ|TOTAL|OBS|LEGENDA"

'--- Estado da importacao -----------------------------------------------------
Private mWbProg As Workbook
Private mCopia As String
Private mCalcAnterior As XlCalculation
Private mSegurancaAnterior As MsoAutomationSecurity

'==============================================================================
' PONTO DE ENTRADA (botao IMPORTAR no MENU)
'==============================================================================
Public Sub ImportarDaProgramacao()

    Dim f As frmImportar

    Set f = New frmImportar
    f.Show

End Sub

'==============================================================================
' ABERTURA E FECHAMENTO DA COPIA
'==============================================================================

'------------------------------------------------------------------------------
' Pede o arquivo, copia para a pasta temporaria e abre a copia escondida.
' Devolve "" se deu tudo certo, ou a mensagem de erro.
'------------------------------------------------------------------------------
Public Function AbrirProgramacao(ByRef nomeArquivo As String) As String

    Dim caminho As Variant

    FecharProgramacao                     ' garante que nada ficou aberto

    caminho = Application.GetOpenFilename( _
              "Planilhas do Excel (*.xls*), *.xls*", , _
              "Escolha o arquivo da Programacao")

    If VarType(caminho) = vbBoolean Then Exit Function   ' cancelou

    On Error GoTo TrataErro

    Application.StatusBar = "Copiando a programacao... aguarde."

    mCopia = Environ$("TEMP") & "\ctrlops_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    FileCopy CStr(caminho), mCopia

    Application.StatusBar = "Abrindo a programacao... pode demorar, o arquivo e grande."
    Application.ScreenUpdating = False

    ' Abre no modo mais leve possivel: sem recalcular, sem macros, sem avisos
    mCalcAnterior = Application.Calculation
    mSegurancaAnterior = Application.AutomationSecurity
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.DisplayAlerts = False

    Set mWbProg = Workbooks.Open(Filename:=mCopia, ReadOnly:=True, UpdateLinks:=0, _
                                 AddToMru:=False, Notify:=False)

    Application.DisplayAlerts = True

    On Error Resume Next
    mWbProg.Windows(1).Visible = False
    On Error GoTo TrataErro

    nomeArquivo = Mid$(CStr(caminho), InStrRev(CStr(caminho), "\") + 1)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Function

TrataErro:
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    AbrirProgramacao = "Nao consegui abrir a programacao: " & Err.Description
    FecharProgramacao

End Function

'------------------------------------------------------------------------------
' Fecha a copia e apaga o arquivo temporario.
'------------------------------------------------------------------------------
Public Sub FecharProgramacao()

    On Error Resume Next

    If Not mWbProg Is Nothing Then
        mWbProg.Close SaveChanges:=False
        Set mWbProg = Nothing
    End If

    If Len(mCopia) > 0 Then
        If Len(Dir$(mCopia)) > 0 Then Kill mCopia
        mCopia = vbNullString
    End If

    ' Devolve o Excel ao estado normal
    If mCalcAnterior <> 0 Then
        Application.Calculation = mCalcAnterior
        mCalcAnterior = 0
    End If
    If mSegurancaAnterior <> 0 Then
        Application.AutomationSecurity = mSegurancaAnterior
        mSegurancaAnterior = 0
    End If

    Application.StatusBar = False
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' Abas da programacao aberta.
'------------------------------------------------------------------------------
Public Function AbasDaProgramacao() As Collection

    Dim res As Collection
    Dim ws As Worksheet

    Set res = New Collection
    If mWbProg Is Nothing Then
        Set AbasDaProgramacao = res
        Exit Function
    End If

    For Each ws In mWbProg.Worksheets
        If ws.Visible = xlSheetVisible Then res.Add ws.Name
    Next ws

    Set AbasDaProgramacao = res

End Function

'==============================================================================
' LEITURA DAS OPs
' Cada OP lida vira uma linha de texto com os campos separados por "|":
'   OP | ITEM | QTD | CLIENTE | DATA | MAQUINA | MES | SITUACAO
'==============================================================================
'------------------------------------------------------------------------------
' Converte qualquer celula em texto, inclusive celulas com erro de formula
' (#N/D, #VALOR). Sem isso, uma unica celula com erro derruba a leitura toda.
'------------------------------------------------------------------------------
Private Function TextoDaCelula(ByVal valor As Variant) As String

    On Error Resume Next

    If IsError(valor) Then Exit Function
    If IsNull(valor) Then Exit Function
    If IsEmpty(valor) Then Exit Function

    TextoDaCelula = Trim$(CStr(valor))

End Function

Public Function LerOPs(ByVal nomeAba As String) As Collection

    Dim ws As Worksheet
    Dim res As Collection
    Dim dados As Variant
    Dim ultima As Long
    Dim lin As Long
    Dim maquinaProg As String
    Dim maquina As String
    Dim numeroOP As String
    Dim texto As String
    Dim dia As Date
    Dim temData As Boolean

    Dim cCliente As Long, cItem As Long, cOP As Long, cQtd As Long, cData As Long

    Set res = New Collection
    Set LerOPs = res

    If mWbProg Is Nothing Then Exit Function

    On Error GoTo TrataErro
    Set ws = mWbProg.Worksheets(nomeAba)

    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ultima < 2 Then Exit Function
    If ultima > 5000 Then ultima = 5000

    Application.StatusBar = "Lendo as OPs de " & nomeAba & "..."
    dados = ws.Range(ws.Cells(1, 1), ws.Cells(ultima, ULTIMA_COLUNA)).Value

    On Error Resume Next            ' uma linha problematica nao para as demais

    For lin = 1 To ultima

        If EhCabecalho(dados, lin, cCliente, cItem, cOP, cQtd, cData) Then

            ' Achou o cabecalho do bloco: as colunas valem daqui para baixo

        ElseIf EhInicioDeBloco(dados, lin, cItem) Then

            ' Titulo com o nome da maquina. Testado ANTES da linha de OP porque
            ' as vezes o titulo tem sujeira nas outras colunas.
            maquinaProg = TextoDaCelula(dados(lin, 1))
            maquina = MaquinaDoSistema(maquinaProg)

        ElseIf TemOP(dados, lin, cOP) Then

            ' Linha de OP
            numeroOP = FormatarMilhar(SomenteDigitos(TextoDaCelula(dados(lin, cOP))))

            If Len(maquina) > 0 Or Len(maquinaProg) > 0 Then

                temData = False
                If cData > 0 Then
                    temData = IsDate(dados(lin, cData))
                    If temData Then dia = DiaDeTurno(CDate(dados(lin, cData)))
                End If

                res.Add numeroOP & "|" & _
                        UCase$(TextoDaCelula(dados(lin, cItem))) & "|" & _
                        SomenteDigitos(TextoDaCelula(dados(lin, cQtd))) & "|" & _
                        UCase$(TextoDaCelula(dados(lin, cCliente))) & "|" & _
                        IIf(temData, Format$(dia, "dd/mm/yyyy"), "") & "|" & _
                        IIf(Len(maquina) > 0, maquina, maquinaProg) & "|" & _
                        IIf(temData, MesDaData(dia), MesAtual) & "|" & _
                        Situacao(numeroOP, maquina, maquinaProg)

            End If

        End If

    Next lin

    On Error GoTo 0
    Application.StatusBar = False
    Exit Function

TrataErro:
    Application.StatusBar = False
    Alertar "Nao consegui ler a aba " & nomeAba & ": " & Err.Description

End Function

'------------------------------------------------------------------------------
' A linha e um cabecalho de bloco? Se for, descobre as colunas.
'------------------------------------------------------------------------------
Private Function EhCabecalho(ByRef dados As Variant, ByVal lin As Long, _
                             ByRef cCliente As Long, ByRef cItem As Long, _
                             ByRef cOP As Long, ByRef cQtd As Long, _
                             ByRef cData As Long) As Boolean

    Dim col As Long
    Dim titulo As String
    Dim achouCliente As Long, achouItem As Long, achouOP As Long
    Dim achouQtd As Long, achouData As Long

    For col = 1 To ULTIMA_COLUNA

        titulo = UCase$(Trim$(Replace(Replace(TextoDaCelula(dados(lin, col)), _
                 ChrW(10), " "), ChrW(13), " ")))
        Do While InStr(titulo, "  ") > 0
            titulo = Replace(titulo, "  ", " ")
        Loop

        If titulo = TIT_CLIENTE Then
            If achouCliente = 0 Then achouCliente = col
        ElseIf titulo = TIT_PRODUTO Then
            If achouItem = 0 Then achouItem = col
        ElseIf titulo = TIT_OP Then
            If achouOP = 0 Then achouOP = col
        ElseIf InStr(1, titulo, TIT_QTD, vbTextCompare) = 1 Then
            If achouQtd = 0 Then achouQtd = col
        ElseIf InStr(1, titulo, TIT_DATA, vbTextCompare) = 1 And _
               InStr(1, titulo, TIT_INICIAL, vbTextCompare) > 0 Then
            If achouData = 0 Then achouData = col
        End If

    Next col

    If achouCliente > 0 And achouOP > 0 And achouItem > 0 Then
        cCliente = achouCliente
        cItem = achouItem
        cOP = achouOP
        cQtd = achouQtd
        cData = achouData
        EhCabecalho = True
    End If

End Function

'------------------------------------------------------------------------------
' A linha e o titulo de um bloco de maquina?
' Basta o texto da coluna A parecer nome de maquina e a coluna do produto
' estar vazia. As outras colunas podem ter qualquer coisa (horas, datas soltas).
'------------------------------------------------------------------------------
Private Function EhInicioDeBloco(ByRef dados As Variant, ByVal lin As Long, _
                                 ByVal cItem As Long) As Boolean

    Dim texto As String

    texto = TextoDaCelula(dados(lin, 1))
    If Len(texto) = 0 Then Exit Function
    If EhLixo(texto) Then Exit Function
    If Not EhTituloDeMaquina(texto) Then Exit Function

    EhInicioDeBloco = VazioNaColuna(dados, lin, cItem)

End Function

'------------------------------------------------------------------------------
' A linha tem numero de OP? Testado em passos separados porque o VBA avalia
' os dois lados de um And, e dados(lin, 0) estouraria antes do cabecalho.
'------------------------------------------------------------------------------
Private Function TemOP(ByRef dados As Variant, ByVal lin As Long, _
                       ByVal cOP As Long) As Boolean

    If cOP <= 0 Then Exit Function
    TemOP = (Len(SomenteDigitos(TextoDaCelula(dados(lin, cOP)))) > 0)

End Function

Private Function VazioNaColuna(ByRef dados As Variant, ByVal lin As Long, _
                               ByVal col As Long) As Boolean

    If col <= 0 Then
        VazioNaColuna = True
    Else
        VazioNaColuna = (Len(TextoDaCelula(dados(lin, col))) = 0)
    End If

End Function

'------------------------------------------------------------------------------
' O texto parece um titulo de maquina?
' Vale quando a primeira palavra e a mesma de alguma maquina cadastrada
' (MAQUINA1, MAQUINA2, MAQUINA3, MAQUINA4, MAQUINA5, MAQUINA6...). Assim linhas
' como "OP LOGISTICA" ou "LIMPEZA DA ROSCA" nao sao confundidas com maquina.
'------------------------------------------------------------------------------
Private Function EhTituloDeMaquina(ByVal texto As String) As Boolean

    Dim nome As Variant
    Dim primeira As String

    primeira = PrimeiraPalavra(texto)
    If Len(primeira) = 0 Then Exit Function

    For Each nome In ListarMaquinas
        If StrComp(PrimeiraPalavra(CStr(nome)), primeira, vbTextCompare) = 0 Then
            EhTituloDeMaquina = True
            Exit Function
        End If
    Next nome

End Function

Private Function PrimeiraPalavra(ByVal texto As String) As String

    Dim t As String

    t = UCase$(Trim$(Replace(texto, ChrW(10), " ")))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    If Len(t) = 0 Then Exit Function
    PrimeiraPalavra = Split(t, " ")(0)

End Function

'------------------------------------------------------------------------------
' Textos que aparecem soltos na coluna A mas nao sao maquina.
'------------------------------------------------------------------------------
Private Function EhLixo(ByVal texto As String) As Boolean

    Dim partes() As String
    Dim i As Long

    If Len(texto) > 40 Then
        EhLixo = True
        Exit Function
    End If

    partes = Split(LIXO, "|")
    For i = LBound(partes) To UBound(partes)
        If InStr(1, texto, partes(i), vbTextCompare) > 0 Then
            EhLixo = True
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' Situacao da OP: vazio = pode importar.
'------------------------------------------------------------------------------
Private Function Situacao(ByVal numeroOP As String, ByVal maquina As String, _
                          ByVal maquinaProg As String) As String

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim idx As Long

    If Len(maquina) = 0 Then
        Situacao = "MAQUINA NAO CADASTRADA (" & maquinaProg & ")"
        Exit Function
    End If

    If LocalizarOP(numeroOP, ws, tbl, idx) Then
        If ws.Name = ABA_ARQUIVO Then
            Situacao = "JA CADASTRADA (ARQUIVO)"
        Else
            Situacao = "JA CADASTRADA (" & ws.Name & ")"
        End If
    End If

End Function

'------------------------------------------------------------------------------
' Acha a maquina do sistema que corresponde ao nome usado na programacao.
' Tenta, nesta ordem:
'   1. o nome como esta, so padronizando o numero (MAQUINA1 1 -> MAQUINA1 01)
'   2. a primeira palavra + o numero do fim
'      (MAQUINA4 AUXILIAR 1 -> MAQUINA4 01)
' Devolve "" se nao existir aba para ela.
'------------------------------------------------------------------------------
Public Function MaquinaDoSistema(ByVal nomeProg As String) As String

    Dim nome As String
    Dim numero As String
    Dim candidato As String

    nome = NomeMaquinaPadrao(nomeProg)

    If Not AbaMaquina(nome) Is Nothing Then
        MaquinaDoSistema = nome
        Exit Function
    End If

    numero = NumeroDoFim(nome)
    If Len(numero) > 0 Then
        candidato = Split(nome, " ")(0) & " " & numero
        If Not AbaMaquina(candidato) Is Nothing Then
            MaquinaDoSistema = candidato
            Exit Function
        End If
    End If

End Function

'------------------------------------------------------------------------------
' Numero do fim do nome, com dois digitos. "" se nao terminar em numero.
'------------------------------------------------------------------------------
Private Function NumeroDoFim(ByVal texto As String) As String

    Dim i As Long
    Dim ch As String
    Dim numero As String

    For i = Len(texto) To 1 Step -1
        ch = Mid$(texto, i, 1)
        If ch >= "0" And ch <= "9" Then
            numero = ch & numero
        Else
            Exit For
        End If
    Next i

    If Len(numero) = 1 Then numero = "0" & numero
    NumeroDoFim = numero

End Function

'------------------------------------------------------------------------------
' "MAQUINA1 1" -> "MAQUINA1 01". Deixa o resto do nome como esta.
'------------------------------------------------------------------------------
Public Function NomeMaquinaPadrao(ByVal nome As String) As String

    Dim texto As String
    Dim numero As String
    Dim i As Long
    Dim ch As String

    texto = UCase$(Trim$(nome))
    texto = Replace(texto, ChrW(10), " ")
    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    ' Numero no fim do nome
    For i = Len(texto) To 1 Step -1
        ch = Mid$(texto, i, 1)
        If ch >= "0" And ch <= "9" Then
            numero = ch & numero
        ElseIf ch = " " And Len(numero) > 0 Then
            Exit For
        ElseIf Len(numero) > 0 Then
            Exit For
        End If
    Next i

    If Len(numero) = 1 Then
        texto = Left$(texto, Len(texto) - 1) & "0" & numero
    End If

    NomeMaquinaPadrao = texto

End Function

'==============================================================================
' IMPORTACAO
' Recebe as linhas escolhidas (mesmo formato de LerOPs) e cadastra cada uma.
'==============================================================================
Public Function ImportarLinhas(ByVal linhas As Collection, ByRef relatorio As String) As Long

    Dim linha As Variant
    Dim campos() As String
    Dim msg As String
    Dim importadas As Long
    Dim puladas As Long
    Dim erros As String

    For Each linha In linhas

        campos = Split(CStr(linha), "|")

        If Len(campos(7)) > 0 Then
            puladas = puladas + 1
        Else
            If CadastrarOP(campos(5), campos(0), campos(1), campos(2), campos(3), _
                           campos(6), msg) Then
                importadas = importadas + 1
            Else
                erros = erros & vbNewLine & "OP " & campos(0) & ": " & msg
            End If
        End If

    Next linha

    relatorio = importadas & IIf(importadas = 1, " OP importada.", " OPs importadas.")
    If puladas > 0 Then
        relatorio = relatorio & vbNewLine & puladas & _
                    IIf(puladas = 1, " ja estava cadastrada e foi pulada.", _
                                     " ja estavam cadastradas e foram puladas.")
    End If
    If Len(erros) > 0 Then relatorio = relatorio & vbNewLine & vbNewLine & "Nao entraram:" & erros

    ImportarLinhas = importadas

End Function

