VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportar 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmImportar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmImportar - Escolha das OPs vindas da Programacao
' Controle de Ordens por Maquina
'
' Na instalacao: inserir um UserForm VAZIO, renomear para frmImportar e colar
' este codigo. Os controles sao criados sozinhos quando o formulario abre.
'==============================================================================
Option Explicit

Private lblTitulo As MSForms.Label
Private lblArquivo As MSForms.Label
Private cboAba As MSForms.ComboBox
Private lstOPs As MSForms.ListBox
Private lblStatus As MSForms.Label
Private WithEvents btnArquivo As MSForms.CommandButton
Attribute btnArquivo.VB_VarHelpID = -1
Private WithEvents btnLer As MSForms.CommandButton
Attribute btnLer.VB_VarHelpID = -1
Private WithEvents btnTodas As MSForms.CommandButton
Attribute btnTodas.VB_VarHelpID = -1
Private WithEvents btnNenhuma As MSForms.CommandButton
Attribute btnNenhuma.VB_VarHelpID = -1
Private WithEvents btnImportar As MSForms.CommandButton
Attribute btnImportar.VB_VarHelpID = -1
Private WithEvents btnFechar As MSForms.CommandButton
Attribute btnFechar.VB_VarHelpID = -1

Private mLinhas As Collection          ' OPs lidas, no formato de mdlImportar

Private Const MARGEM As Single = 14

'==============================================================================
' MONTAGEM
'==============================================================================
Private Sub UserForm_Initialize()

    Me.Caption = "Importar da Programacao"
    Me.BackColor = RGB(255, 255, 255)
    Me.Width = 720
    Me.Height = 520

    Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    With lblTitulo
        .Left = 0: .Top = 0: .Width = Me.InsideWidth: .Height = 36
        .BackColor = RGB(31, 31, 31): .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 12: .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .Caption = vbNewLine & "IMPORTAR OPs DA PROGRAMACAO"
    End With

    Set btnArquivo = Me.Controls.Add("Forms.CommandButton.1", "btnArquivo", True)
    With btnArquivo
        .Left = MARGEM: .Top = 48: .Width = 130: .Height = 26
        .Caption = "ESCOLHER ARQUIVO"
        .BackColor = RGB(31, 31, 31): .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
    End With

    Set lblArquivo = Me.Controls.Add("Forms.Label.1", "lblArquivo", True)
    With lblArquivo
        .Left = MARGEM + 140: .Top = 54: .Width = 540: .Height = 16
        .Font.Name = "Segoe UI": .Font.Size = 9
        .ForeColor = RGB(120, 120, 120)
        .BackStyle = fmBackStyleTransparent
        .Caption = "Nenhum arquivo aberto."
    End With

    Set cboAba = Me.Controls.Add("Forms.ComboBox.1", "cboAba", True)
    With cboAba
        .Left = MARGEM: .Top = 84: .Width = 160: .Height = 22
        .Font.Name = "Segoe UI": .Font.Size = 10
        .Style = fmStyleDropDownList
        .Enabled = False
    End With

    Set btnLer = Me.Controls.Add("Forms.CommandButton.1", "btnLer", True)
    With btnLer
        .Left = MARGEM + 170: .Top = 84: .Width = 110: .Height = 22
        .Caption = "LER OPs"
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .Enabled = False
    End With

    Set btnTodas = Me.Controls.Add("Forms.CommandButton.1", "btnTodas", True)
    With btnTodas
        .Left = MARGEM + 400: .Top = 84: .Width = 130: .Height = 22
        .Caption = "Marcar disponiveis"
        .Font.Name = "Segoe UI": .Font.Size = 9
    End With

    Set btnNenhuma = Me.Controls.Add("Forms.CommandButton.1", "btnNenhuma", True)
    With btnNenhuma
        .Left = MARGEM + 538: .Top = 84: .Width = 130: .Height = 22
        .Caption = "Desmarcar todas"
        .Font.Name = "Segoe UI": .Font.Size = 9
    End With

    Set lstOPs = Me.Controls.Add("Forms.ListBox.1", "lstOPs", True)
    With lstOPs
        .Left = MARGEM: .Top = 114: .Width = 682: .Height = 320
        .Font.Name = "Segoe UI": .Font.Size = 9
        .ColumnCount = 7
        .ColumnWidths = "60 pt;70 pt;70 pt;90 pt;65 pt;110 pt;190 pt"
        .ColumnHeads = False
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption        ' mostra as caixinhas
    End With

    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = MARGEM: .Top = 444: .Width = 430: .Height = 30
        .Font.Name = "Segoe UI": .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Caption = "Escolha o arquivo da programacao para comecar."
    End With

    Set btnImportar = Me.Controls.Add("Forms.CommandButton.1", "btnImportar", True)
    With btnImportar
        .Left = 460: .Top = 444: .Width = 150: .Height = 32
        .Caption = "IMPORTAR MARCADAS"
        .BackColor = RGB(31, 31, 31): .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
    End With

    Set btnFechar = Me.Controls.Add("Forms.CommandButton.1", "btnFechar", True)
    With btnFechar
        .Left = 618: .Top = 444: .Width = 78: .Height = 32
        .Caption = "Fechar"
        .Font.Name = "Segoe UI": .Font.Size = 9
        .Cancel = True
    End With

    Set mLinhas = New Collection

End Sub

'==============================================================================
' EVENTOS
'==============================================================================
Private Sub btnArquivo_Click()

    Dim erroMsg As String
    Dim nomeArquivo As String
    Dim aba As Variant

    erroMsg = AbrirProgramacao(nomeArquivo)

    If Len(erroMsg) > 0 Then
        Alertar erroMsg
        Exit Sub
    End If

    If Len(nomeArquivo) = 0 Then Exit Sub       ' cancelou

    lblArquivo.Caption = nomeArquivo & "   (copia temporaria, somente leitura)"

    cboAba.Clear
    For Each aba In AbasDaProgramacao
        cboAba.AddItem CStr(aba)
    Next aba

    SelecionarAbaDoMes

    cboAba.Enabled = True
    btnLer.Enabled = True
    lblStatus.Caption = "Escolha a aba do mes e clique em LER OPs."

End Sub

'--- Deixa marcada a aba com o nome do mes atual, se existir -------------------
Private Sub SelecionarAbaDoMes()

    Dim mesAgora As String
    Dim i As Long

    mesAgora = Split(MesAtual, "/")(0)

    ' O nome da aba pode ser "SETEMBRO" ou "TPA SEMI SETEMBRO"
    For i = 0 To cboAba.ListCount - 1
        If InStr(1, cboAba.List(i), mesAgora, vbTextCompare) > 0 Then
            cboAba.ListIndex = i
            Exit Sub
        End If
    Next i

End Sub

Private Sub btnLer_Click()

    Dim linha As Variant
    Dim campos() As String
    Dim i As Long
    Dim disponiveis As Long

    If Len(cboAba.Text) = 0 Then
        lblStatus.Caption = "Escolha a aba do mes."
        Exit Sub
    End If

    Me.MousePointer = fmMousePointerHourGlass
    lstOPs.Clear

    Set mLinhas = LerOPs(cboAba.Text)

    For Each linha In mLinhas
        campos = Split(CStr(linha), "|")
        lstOPs.AddItem campos(0)                    ' OP
        i = lstOPs.ListCount - 1
        lstOPs.List(i, 1) = campos(1)               ' ITEM
        lstOPs.List(i, 2) = FormatarMilhar(campos(2))   ' QTD
        lstOPs.List(i, 3) = campos(3)               ' CLIENTE
        lstOPs.List(i, 4) = campos(4)               ' DATA
        lstOPs.List(i, 5) = campos(5)               ' MAQUINA
        lstOPs.List(i, 6) = campos(7)               ' SITUACAO
        If Len(campos(7)) = 0 Then disponiveis = disponiveis + 1
    Next linha

    Me.MousePointer = fmMousePointerDefault

    lblStatus.Caption = mLinhas.Count & " OPs na aba " & cboAba.Text & ".   " & _
                        disponiveis & " podem ser importadas." & vbNewLine & _
                        "Marque as que voce quer trazer."

End Sub

Private Sub btnTodas_Click()

    Dim i As Long

    For i = 0 To lstOPs.ListCount - 1
        lstOPs.Selected(i) = (Len(Trim$(CStr(lstOPs.List(i, 6)))) = 0)
    Next i

End Sub

Private Sub btnNenhuma_Click()

    Dim i As Long

    For i = 0 To lstOPs.ListCount - 1
        lstOPs.Selected(i) = False
    Next i

End Sub

Private Sub btnImportar_Click()

    Dim escolhidas As Collection
    Dim i As Long
    Dim relatorio As String

    Set escolhidas = New Collection

    For i = 0 To lstOPs.ListCount - 1
        If lstOPs.Selected(i) Then escolhidas.Add mLinhas(i + 1)
    Next i

    If escolhidas.Count = 0 Then
        lblStatus.Caption = "Marque ao menos uma OP."
        Exit Sub
    End If

    If Not Confirmar("Importar " & escolhidas.Count & _
                     IIf(escolhidas.Count = 1, " OP?", " OPs?")) Then Exit Sub

    Me.MousePointer = fmMousePointerHourGlass
    ImportarLinhas escolhidas, relatorio
    Me.MousePointer = fmMousePointerDefault

    Avisar relatorio

    ' Relê para atualizar a situacao das que acabaram de entrar
    btnLer_Click

End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    FecharProgramacao          ' fecha a copia e apaga o arquivo temporario
End Sub

