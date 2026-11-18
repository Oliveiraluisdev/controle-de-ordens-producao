VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOP 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmOP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmOP - Formulario de cadastro e edicao de OP
' Controle de Ordens por Maquina
'
' Os campos sao criados por codigo quando o formulario abre. Na instalacao,
' basta inserir um UserForm VAZIO, renomear para frmOP e colar este codigo.
'
' No cadastro, o formulario NAO fecha depois de salvar: limpa so o numero da
' OP e mantem maquina, item, quantidade e cliente, para cadastrar varias em
' sequencia. Na edicao, fecha normalmente.
'==============================================================================
Option Explicit

'--- Controles ----------------------------------------------------------------
Private lblTitulo As MSForms.Label
Private cboMaquina As MSForms.ComboBox
Private txtOP As MSForms.TextBox
Private cboItem As MSForms.ComboBox
Private cboQtd As MSForms.ComboBox
Private cboCliente As MSForms.ComboBox
Private cboMes As MSForms.ComboBox
Private WithEvents btnSalvar As MSForms.CommandButton
Attribute btnSalvar.VB_VarHelpID = -1
Private WithEvents btnCancelar As MSForms.CommandButton
Attribute btnCancelar.VB_VarHelpID = -1
Private lblStatus As MSForms.Label

'--- Estado -------------------------------------------------------------------
Private mEdicao As Boolean
Private mAbaOrigem As String
Private mOPOriginal As String
Private mContador As Long          ' OPs cadastradas desde que o form abriu

'--- Layout -------------------------------------------------------------------
Private Const MARGEM As Single = 18
Private Const LARG_ROTULO As Single = 78
Private Const LARG_CAMPO As Single = 190
Private Const ALT_CAMPO As Single = 22
Private Const PASSO As Single = 34

'==============================================================================
' MONTAGEM
'==============================================================================
Private Sub UserForm_Initialize()

    Dim y As Single

    Me.Caption = "Controle de Ordens"
    Me.BackColor = RGB(255, 255, 255)
    Me.Width = MARGEM * 2 + LARG_ROTULO + LARG_CAMPO + 12
    Me.Height = 330

    '--- Faixa de titulo ---
    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Left = 0: .Top = 0
        .Width = Me.InsideWidth: .Height = 36
        .BackColor = RGB(31, 31, 31)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 12: .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .Caption = vbNewLine & "CADASTRAR NOVA OP"
    End With

    y = 54

    Set cboMaquina = NovoCombo("cboMaquina", "Maquina", y, True):   y = y + PASSO
    Set txtOP = NovoTexto("txtOP", "OP", y):                          y = y + PASSO
    Set cboItem = NovoCombo("cboItem", "Item", y, False):             y = y + PASSO
    Set cboQtd = NovoCombo("cboQtd", "Quantidade", y, False):         y = y + PASSO
    Set cboCliente = NovoCombo("cboCliente", "Cliente", y, False):    y = y + PASSO
    Set cboMes = NovoCombo("cboMes", "Mes", y, False):                y = y + PASSO + 6

    '--- Botoes ---
    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Left = MARGEM + LARG_ROTULO: .Top = y
        .Width = 110: .Height = 30
        .Caption = "CADASTRAR"
        .BackColor = RGB(31, 31, 31)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .Default = True                      ' Enter salva
    End With

    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar", True)
    With btnCancelar
        .Left = MARGEM + LARG_ROTULO + 118: .Top = y
        .Width = 72: .Height = 30
        .Caption = "Fechar"
        .Font.Name = "Segoe UI": .Font.Size = 9
        .Cancel = True                       ' Esc fecha
    End With

    y = y + 38

    '--- Linha de status (confirma cada cadastro sem abrir mensagem) ---
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = MARGEM: .Top = y
        .Width = LARG_ROTULO + LARG_CAMPO: .Height = 30
        .Font.Name = "Segoe UI": .Font.Size = 9
        .ForeColor = RGB(0, 110, 50)
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Caption = vbNullString
    End With

    Me.Height = y + 30 + MARGEM + (Me.Height - Me.InsideHeight)

    CarregarListas

End Sub

Private Function NovoRotulo(ByVal texto As String, ByVal y As Single) As MSForms.Label

    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add("Forms.Label.1", , True)
    With lbl
        .Left = MARGEM: .Top = y + 4
        .Width = LARG_ROTULO: .Height = 16
        .Caption = texto
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    Set NovoRotulo = lbl

End Function

Private Function NovoTexto(ByVal nome As String, ByVal rotulo As String, _
                           ByVal y As Single) As MSForms.TextBox

    Dim txt As MSForms.TextBox

    NovoRotulo rotulo, y
    Set txt = Me.Controls.Add("Forms.TextBox.1", nome, True)
    With txt
        .Left = MARGEM + LARG_ROTULO: .Top = y
        .Width = LARG_CAMPO: .Height = ALT_CAMPO
        .Font.Name = "Segoe UI": .Font.Size = 11
    End With

    Set NovoTexto = txt

End Function

Private Function NovoCombo(ByVal nome As String, ByVal rotulo As String, _
                           ByVal y As Single, ByVal somenteLista As Boolean) As MSForms.ComboBox

    Dim cbo As MSForms.ComboBox

    NovoRotulo rotulo, y
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    With cbo
        .Left = MARGEM + LARG_ROTULO: .Top = y
        .Width = LARG_CAMPO: .Height = ALT_CAMPO
        .Font.Name = "Segoe UI": .Font.Size = 11
        .MatchEntry = fmMatchEntryComplete     ' completa enquanto digita
        If somenteLista Then
            .Style = fmStyleDropDownList       ' maquina: so da lista
        Else
            .Style = fmStyleDropDownCombo      ' demais: livre + sugestoes
        End If
    End With

    Set NovoCombo = cbo

End Function

'------------------------------------------------------------------------------
' Preenche as listas a partir da _CONFIG.
'------------------------------------------------------------------------------
Private Sub CarregarListas()

    Dim v As Variant

    For Each v In ListarMaquinas
        cboMaquina.AddItem CStr(v)
    Next v

    For Each v In OrdenarColecao(ListaConfig(CFG_ITENS))
        cboItem.AddItem CStr(v)
    Next v

    For Each v In OrdenarColecao(ListaConfig(CFG_QUANTIDADES))
        cboQtd.AddItem FormatarMilhar(CStr(v))
    Next v

    For Each v In OrdenarColecao(ListaConfig(CFG_CLIENTES))
        cboCliente.AddItem CStr(v)
    Next v

    ' Meses sugeridos: dois para tras e nove para a frente
    For Each v In ListaMeses(2, 9)
        cboMes.AddItem CStr(v)
    Next v

End Sub

'==============================================================================
' MODOS (chamados por mdlOP antes do Show)
'==============================================================================
Public Sub PrepararNovo(ByVal maquina As String)

    mEdicao = False
    lblTitulo.Caption = vbNewLine & "CADASTRAR NOVA OP"
    btnSalvar.Caption = "CADASTRAR"

    cboMes.Text = MesAtual

    If Len(maquina) > 0 Then SelecionarMaquina maquina

End Sub

Public Sub PrepararEdicao(ByVal maquina As String, ByVal numeroOP As String, _
                          ByVal item As String, ByVal qtdTexto As String, _
                          ByVal cliente As String, ByVal mes As String)

    mEdicao = True
    mAbaOrigem = maquina
    mOPOriginal = numeroOP

    lblTitulo.Caption = vbNewLine & "EDITAR OP " & numeroOP
    btnSalvar.Caption = "SALVAR"
    btnCancelar.Caption = "Cancelar"

    SelecionarMaquina maquina
    txtOP.Text = numeroOP
    cboItem.Text = item
    cboQtd.Text = FormatarMilhar(qtdTexto)
    cboCliente.Text = cliente
    cboMes.Text = mes

End Sub

Private Sub SelecionarMaquina(ByVal maquina As String)

    Dim i As Long

    For i = 0 To cboMaquina.ListCount - 1
        If StrComp(cboMaquina.List(i), maquina, vbTextCompare) = 0 Then
            cboMaquina.ListIndex = i
            Exit Sub
        End If
    Next i

End Sub

' Obs.: os campos sao lidos com .Text (e nao .Value), que nunca devolve Null.
'==============================================================================
' EVENTOS
'==============================================================================
Private Sub UserForm_Activate()
    On Error Resume Next
    If Len(cboMaquina.Text) = 0 Then
        cboMaquina.SetFocus
    Else
        txtOP.SetFocus
    End If
End Sub

Private Sub btnSalvar_Click()

    Dim msg As String
    Dim ok As Boolean

    If mEdicao Then
        ok = EditarOP(mAbaOrigem, mOPOriginal, cboMaquina.Text, txtOP.Text, _
                      cboItem.Text, cboQtd.Text, cboCliente.Text, cboMes.Text, msg)
    Else
        ok = CadastrarOP(cboMaquina.Text, txtOP.Text, _
                         cboItem.Text, cboQtd.Text, cboCliente.Text, cboMes.Text, msg)
    End If

    If Not ok Then
        Alertar msg
        Exit Sub
    End If

    If mEdicao Then
        ' Edicao: fecha e confirma
        Unload Me
        Avisar msg
    Else
        ' Cadastro: continua aberto para a proxima OP
        mContador = mContador + 1
        lblStatus.Caption = msg & vbNewLine & mContador & _
                            IIf(mContador = 1, " OP cadastrada", " OPs cadastradas") & _
                            " nesta sequencia."
        txtOP.Text = vbNullString
        txtOP.SetFocus
    End If

End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

