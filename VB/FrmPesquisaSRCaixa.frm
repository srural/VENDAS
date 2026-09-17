VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPesquisaSRCaixa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "FrmPesquisaSRCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6765
   Begin VB.CheckBox chDenominacao 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4470
      TabIndex        =   7
      Top             =   270
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar Progresso 
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   4230
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chCond 
      Caption         =   "Qualquer parte do campo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4470
      TabIndex        =   4
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton BtSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   285
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   585
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "OK"
      Height          =   285
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   585
   End
   Begin VB.TextBox txtPesquisa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3075
   End
   Begin MSComctlLib.ListView lvPesquisa 
      Height          =   3795
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   6694
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "CLICodCliente"
         Text            =   "Código"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "CLINome"
         Text            =   "Nome"
         Object.Width           =   7762
      EndProperty
   End
   Begin MSComctlLib.ListView LvPesquisa2 
      Height          =   3645
      Left            =   90
      TabIndex        =   6
      Top             =   540
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   6429
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "CLICodCliente"
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "CLINome"
         Text            =   "Nome"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Endereco"
         Object.Width           =   4764
      EndProperty
   End
End
Attribute VB_Name = "FrmPesquisaSRCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Pesquisa As ADODB.Recordset 'ado de pesquisa
Public SqlPesquisa As String 'sql q monta tela de pesquisa
Public Campos1 As String  'campos da pesquisa
Public Campos2 As String  'campos da pesquisa
Public Campos3 As String  'campos da pesquisa
Public Condicao1 As String 'campo condicao do rel
Public Tabela As String
Public Ordenado As String 'ordenado
Public CampoExtra1 As String 'campo para inserir no sql sem aparecer na lista
Public TipoResultado As Currency
Public QtdCampos As Currency
Public ResultadoPesquisa As String 'grava o resultado para os formularios o pegarem

Private Sub BtOk_Click()
Dim CondCampo As String
On Error GoTo Linha
DoEvents

    Set Pesquisa = New ADODB.Recordset
    
    If QtdCampos = 1 Then
        SqlPesquisa = "Select " & Campos1 & ", " & Campos2 & IIf(CampoExtra1 <> "", "," & CampoExtra1, "") & " From " & Tabela
    ElseIf QtdCampos = 2 Then
        SqlPesquisa = "Select " & Campos1 & ", " & Campos2 & ", " & Campos3 & IIf(CampoExtra1 <> "", "," & CampoExtra1, "") & " From " & Tabela
    End If
    
    Pesquisa.CacheSize = 30
    Pesquisa.CursorType = adOpenForwardOnly
    Pesquisa.CursorLocation = adUseClient
    
    If QtdCampos = 2 And chDenominacao.Value = 1 Then
        CondCampo = Campos3
    Else
        CondCampo = Campos2
    End If

    If IsNumeric(txtPesquisa) Then
        SqlPesquisa = SqlPesquisa & " Where " & Campos1 & " = " & txtPesquisa.Text & IIf(Condicao1 <> "", " and " & Condicao1, "")
    ElseIf Not IsNumeric(txtPesquisa) And txtPesquisa <> "" And chCond.Value = 1 Then
        SqlPesquisa = SqlPesquisa & " Where " & CondCampo & " Like '%" & txtPesquisa.Text & "%'" & IIf(Condicao1 <> "", " and " & Condicao1, "")
    ElseIf Not IsNumeric(txtPesquisa) And txtPesquisa <> "" And chCond.Value = 0 Then
        SqlPesquisa = SqlPesquisa & " Where " & CondCampo & " Like '" & txtPesquisa.Text & "%'" & IIf(Condicao1 <> "", " and " & Condicao1, "")
    Else
        SqlPesquisa = SqlPesquisa & IIf(Condicao1 <> "", " Where " & Condicao1, "")
    End If

    If Condicao1 <> "" Then
        If Ordenado = "" Then
            SqlPesquisa = SqlPesquisa & " Order by " & Campos2
        Else
            SqlPesquisa = SqlPesquisa & " Order by " & Ordenado
        End If
    ElseIf Ordenado <> "" Then
        SqlPesquisa = SqlPesquisa & " Order by " & Ordenado
    End If
    
    Pesquisa.Open SqlPesquisa, BancoSRCaixa, adOpenStatic, adLockOptimistic, adCmdText

    If Pesquisa.Eof = False And IsNumeric(txtPesquisa) = True Then
'        ResultadoPesquisa = txtPesquisa
'        Unload Me
'        Exit Sub
        ProcuraForm txtPesquisa
    End If


    PreencheLvPesq
    'SelecionaCampo

Linha:
    LogErros
End Sub

Sub PreencheLvPesq()
On Error GoTo Linha

Dim lstPesq As ListItem
Dim X, y As Currency

lvPesquisa.ListItems.Clear
LvPesquisa2.ListItems.Clear

If Pesquisa.RecordCount = 0 Then Exit Sub

If Pesquisa.RecordCount >= 100 Then
    Progresso.Max = Pesquisa.RecordCount / 100
Else
    Progresso.Max = Pesquisa.RecordCount
End If

X = 0
y = 0

Do While Not Pesquisa.Eof
    Dim varTemp(0 To 1) As String
    varTemp(0) = IIf(IsNull(Pesquisa(0)), 0, Pesquisa(0))
    varTemp(1) = LimpaFormata(IIf(IsNull(Pesquisa(1)), 0, Pesquisa(1)))
If QtdCampos = 1 Then
    Set lstPesq = lvPesquisa.ListItems.Add(, , varTemp(0))
    lstPesq.SubItems(1) = varTemp(1)
ElseIf QtdCampos = 2 Then
    Set lstPesq = LvPesquisa2.ListItems.Add(, , varTemp(0))
    lstPesq.SubItems(1) = varTemp(1)
    lstPesq.SubItems(2) = Pesquisa(2)
End If

    Pesquisa.MoveNext
    X = X + 1
    If Pesquisa.RecordCount >= 100 Then
        If X = 100 Then
            y = y + 1
            Progresso.Value = y
            X = 0
        End If
    Else
        Progresso.Value = X
    End If
Loop
    
    If Pesquisa.Eof = False Then Pesquisa.MoveFirst
    Progresso.Value = 0
Linha:
    
End Sub

Private Sub BtSair_Click()
    ResultadoPesquisa = 0
    Unload Me
End Sub

Private Sub Form_Activate()
   
    'If CfgPesquisa = "1" Then
       BtOk_Click
    'Else
    '   lvPesquisa.ListItems.Clear
    '   LvPesquisa2.ListItems.Clear
    'End If
    
End Sub

Private Sub Form_Load()
'On Error Resume Next

    Me.Left = (MDIPrincipal.Width - Me.Width) - 200
    Me.Top = (MDIPrincipal.Height - Me.Height) - 500
    
    If QtdCampos = 1 Then
        lvPesquisa.Visible = True
        LvPesquisa2.Visible = False
        chDenominacao.Visible = False
    ElseIf QtdCampos = 2 Then
        lvPesquisa.Visible = False
        LvPesquisa2.Visible = True
        chDenominacao.Visible = True
        chDenominacao.Caption = "Pesquisar por " & Mid(Campos3, 4, Len(Campos3))
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
SqlPesquisa = ""
Campos1 = ""
Campos2 = ""
Campos3 = ""
Condicao1 = ""
CampoExtra1 = ""
TipoResultado = 0
QtdCampos = 0
Ordenado = ""
End Sub

Private Sub lvPesquisa_DblClick()
'On Error Resume Next
    lvPesquisa_KeyPress 13
End Sub

Private Sub lvPesquisa_GotFocus()
On Error Resume Next

    txtPesquisa.Text = lvPesquisa.SelectedItem.Text

End Sub

Private Sub lvPesquisa_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error Resume Next
    lvPesquisa_GotFocus
End Sub

Private Sub lvPesquisa_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        
        If TipoResultado <= 2 Then
            ProcuraForm lvPesquisa.SelectedItem.Text
        Else
            ProcuraForm lvPesquisa.ListItems.Item(lvPesquisa.SelectedItem.Index).ListSubItems(1).Text
        End If
        
        Unload Me
    End If
End Sub

Private Sub lvPesquisa2_DblClick()
'On Error Resume Next
    lvPesquisa2_KeyPress 13
End Sub

Private Sub lvPesquisa2_GotFocus()
'On Error Resume Next
    txtPesquisa.Text = LvPesquisa2.SelectedItem.Text
End Sub

Private Sub lvPesquisa2_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error Resume Next
    lvPesquisa2_GotFocus
End Sub

Private Sub lvPesquisa2_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    If KeyAscii = 13 Then
        'If TipoResultado <= 2 Then
            ProcuraForm LvPesquisa2.SelectedItem.Text
        'Else
        '    ProcuraForm LvPesquisa2.ListItems(1).Text
        'End If
        
        Unload Me
        
    End If
End Sub

Private Sub ProcuraForm(Resultado As String)
On Error GoTo Linha

    If TipoResultado = 1 Then
        
        Formulário.Registros.Requery
        Formulário.Registros.Find CampoIndice.DataField & " = " & Resultado, , , 0
            
        If Formulário.Registros.Eof = False Then
            Status 2
            MostraDados
        Else
            MsgBox "Ocorrência não encontrada!", vbInformation
            txtPesquisa.SetFocus
            SelecionaCampo
            Exit Sub
        End If
        
    ElseIf TipoResultado = 2 Then
        ResultadoPesquisa = Resultado
        SendKeys "{enter}"
    ElseIf TipoResultado = 3 Then
        ResultadoPesquisa = Resultado
    End If
        Unload Me
        
Linha:
    LogErros
End Sub

Private Sub TxtPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
    If KeyCode = 38 Then
        If QtdCampos = 1 Then lvPesquisa.SetFocus
        If QtdCampos = 2 Then LvPesquisa2.SetFocus
    ElseIf KeyCode = 40 Then
        If QtdCampos = 1 Then lvPesquisa.SetFocus
        If QtdCampos = 2 Then LvPesquisa2.SetFocus
    ElseIf KeyCode = 13 Then
        BtOk_Click
    End If
End Sub

Public Function LimpaFormata(Valor As String) As String
'On Error Resume Next
'Funcao que traz o codigo do plano de contas formatado
    LimpaFormata = Replace(Valor, "§", " - ")
End Function
