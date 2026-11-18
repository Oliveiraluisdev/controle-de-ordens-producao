VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' EstaPastaDeTrabalho - Eventos da pasta de trabalho
' Controle de Ordens por Maquina
'
' Este codigo vai DENTRO do EstaPastaDeTrabalho (em "Microsoft Excel Objetos"),
' e nao em um modulo comum.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Ao abrir o arquivo, vai para o MENU.
'------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Worksheets(ABA_MENU).Activate
End Sub

'------------------------------------------------------------------------------
' Duplo clique em qualquer aba de maquina: marca ou desmarca o checklist.
'------------------------------------------------------------------------------
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, _
                                            ByVal Target As Range, _
                                            Cancel As Boolean)
    TratarDuploClique Sh, Target, Cancel
End Sub

'------------------------------------------------------------------------------
' Busca rapida: digitar a OP no campo do MENU e apertar Enter.
'------------------------------------------------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If Sh.Name <> ABA_MENU Then Exit Sub
    If Intersect(Target, Sh.Range(CELULA_BUSCA)) Is Nothing Then Exit Sub
    If Len(Trim$(CStr(Sh.Range(CELULA_BUSCA).Value))) = 0 Then Exit Sub

    BuscarOP

End Sub

