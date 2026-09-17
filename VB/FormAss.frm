VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormAss 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12780
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -225
      ScaleHeight     =   345
      ScaleWidth      =   12990
      TabIndex        =   23
      Top             =   7920
      Width           =   13020
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12795
      TabIndex        =   21
      Top             =   0
      Width           =   12795
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Assistências"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   135
         TabIndex        =   22
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.CommandButton BtNovo 
      Caption         =   "&Novo"
      Height          =   330
      Left            =   11325
      TabIndex        =   20
      Top             =   5715
      Width           =   1230
   End
   Begin VB.CommandButton BtSalvar 
      Caption         =   "&Gravar"
      Height          =   330
      Left            =   11325
      TabIndex        =   8
      Top             =   6210
      Width           =   1230
   End
   Begin VB.CommandButton BtExluir 
      Caption         =   "&Excluir"
      Height          =   330
      Left            =   11325
      TabIndex        =   19
      Top             =   6720
      Width           =   1230
   End
   Begin VB.CommandButton BtLimpa 
      Caption         =   "&Limpa Tela"
      Height          =   330
      Left            =   10305
      TabIndex        =   18
      Top             =   900
      Width           =   1500
   End
   Begin VB.Frame Frame 
      Enabled         =   0   'False
      Height          =   3750
      Left            =   135
      TabIndex        =   11
      Top             =   1395
      Width           =   12435
      Begin VB.TextBox TxtContato 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3225
         MaxLength       =   30
         TabIndex        =   4
         Top             =   405
         Width           =   4260
      End
      Begin VB.TextBox TxtProblema 
         Appearance      =   0  'Flat
         Height          =   1320
         Left            =   135
         MaxLength       =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1035
         Width           =   12075
      End
      Begin VB.TextBox TxtTecnico 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7650
         MaxLength       =   30
         TabIndex        =   6
         Top             =   405
         Width           =   4560
      End
      Begin VB.TextBox TxtSolucao 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   135
         MaxLength       =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2610
         Width           =   12120
      End
      Begin VB.TextBox TxtProtocolo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   3
         Top             =   405
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contato"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Protocolo"
         Height          =   195
         Left            =   1575
         TabIndex        =   15
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Técnico"
         Height          =   195
         Left            =   7650
         TabIndex        =   14
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P&roblema"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   810
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Solução"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   2385
         Width           =   585
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Ass"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4230
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT  *  FROM Ass"
      Top             =   5985
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Frame FrmEnt 
      Caption         =   "Entidade"
      Height          =   690
      Left            =   135
      TabIndex        =   10
      Top             =   675
      Width           =   9915
      Begin VB.TextBox TxtCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8370
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin VB.Data DatDados 
         Caption         =   "Ent"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   1665
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From ENT WHERE CodEntidade = 1 or CodEntidade >= 5 Order by Nome"
         Top             =   225
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormAss.frx":0000
         DataField       =   "Nome"
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   225
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSDBGrid.DBGrid dbgrid1 
      Bindings        =   "FormAss.frx":001A
      Height          =   2445
      Left            =   180
      OleObjectBlob   =   "FormAss.frx":0034
      TabIndex        =   9
      ToolTipText     =   "Duplo clique para editar ou excluir"
      Top             =   5265
      Width           =   10950
   End
End
Attribute VB_Name = "FormAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String
Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
Criterio = Criterio & Codigo
DatDados(1).Recordset.FindFirst Criterio
If DatDados(1).Recordset.NoMatch Then
    MsgBox "Entidade não Localizada", vbExclamation, App.Title
    DbcPesquisa.Text = ""
    DbcPesquisa.SetFocus
Else
    DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
    DatDados(0).RecordSource = "SELECT * FROM Ass WHERE Entidade = " & Val(Codigo) & " ORDER BY CodAss"
    DatDados(0).Refresh
    
    txtCod.Text = DbcPesquisa.BoundText
    If DatDados(0).Recordset.RecordCount <> 0 Then
         DatDados(0).Recordset.MoveLast
    End If
End If

End Sub


Private Sub BtExluir_Click()
If MsgBox("Deseja excluir o registro", vbYesNo, App.Title) = vbYes Then
    DatDados(0).Recordset.Delete
    If DatDados(0).Recordset.RecordCount > 0 Then
        DatDados(0).Recordset.MoveLast
    End If
    BtLimpa_Click
End If
End Sub

Private Sub BtLimpa_Click()
'DbcPesquisa = ""
MebData = "__/__/____"
TxtProtocolo = ""
TxtContato = ""
TxtProblema = ""
TxtTecnico = ""
TxtSolucao = ""
End Sub

Private Sub BtNovo_Click()
BtLimpa_Click
If IsNumeric(DbcPesquisa.BoundText) Then
'    Frame2.Enabled = True
    BtSalvar.Enabled = True
    BtNovo.Enabled = False
    BtExluir.Enabled = False
    DatDados(0).Recordset.AddNew
    DatDados(0).Recordset("Entidade") = txtCod
    DatDados(0).Recordset.Update
    DatDados(0).Recordset.Bookmark = DatDados(0).Recordset.LastModified
    
'    DatDados(0).Recordset.Edit
    MebData = Date
    TxtProtocolo = "ASS" & Right(Year(Date), 2) & psCompString(1, DatDados(0).Recordset("CodAss"), 6)
    TxtContato.SetFocus
End If
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtSalvar_Click()

With DatDados(0).Recordset
    If .RecordCount > 0 Then
        .Edit
        .Fields("Data") = MebData
        .Fields("Protocolo") = TxtProtocolo
        .Fields("Contato") = TxtContato
        .Fields("Problema") = TxtProblema
        .Fields("Tecnico") = TxtTecnico
        .Fields("Solucao") = TxtSolucao
        .Update
        BtNovo.Enabled = True
        BtSalvar.Enabled = False
        Frame.Enabled = False
    End If
End With

End Sub

Private Sub DbcPesquisa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DbcPesquisa.Text <> "" Then
        If Not IsNumeric(DbcPesquisa.Text) Then
            If IsNumeric(DbcPesquisa.BoundText) Then
                Pesquisa DbcPesquisa.BoundText
            End If
        Else
            txtCod.Text = DbcPesquisa.BoundText
            Pesquisa DbcPesquisa.Text
        End If
    End If
End If
End Sub

Private Sub dbgrid1_Click()
BtLimpa_Click
End Sub

Private Sub DBGrid1_DBLClick()
    BtExluir.Enabled = True
    If IsDate(DBGrid1.Columns(2)) Then
        MebData = DBGrid1.Columns(2)
    End If
    TxtProtocolo = DBGrid1.Columns(3)
    TxtContato = DBGrid1.Columns(4)
    TxtProblema = DBGrid1.Columns(5)
    TxtTecnico = DBGrid1.Columns(6)
    TxtSolucao = DBGrid1.Columns(7)
    Frame.Enabled = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Dim Componente As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = 100  '(MDIPrincipal.Width - Me.Width) / 2
Me.Top = 100 '(MDIPrincipal.Height - Me.Height) / 6

DatDados(0).RecordSource = "SELECT * FROM Ass  WHERE Entidade = " & Val(DbcPesquisa.BoundText) & "  ORDER BY CodAss"
DatDados(0).Refresh

DatDados(1).RecordSource = "SELECT * FROM Ent  WHERE Tipo = 1 or Tipo >= 5 Order By Nome"
DatDados(1).Refresh

Frame.Enabled = False

BtLimpa_Click
End Sub


