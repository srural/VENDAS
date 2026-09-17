VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelConv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6120
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5355
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -495
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   16
      Top             =   5805
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   14
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Convênios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   135
         TabIndex        =   15
         Top             =   45
         Width           =   4245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Desconto"
      Height          =   1095
      Left            =   405
      TabIndex        =   7
      Top             =   1485
      Width           =   3435
      Begin MSMask.MaskEdBox Meb1 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   495
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Meb2 
         Height          =   285
         Left            =   1755
         TabIndex        =   12
         Top             =   495
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   2970
         TabIndex        =   13
         Top             =   540
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convênio"
         Height          =   195
         Left            =   1755
         TabIndex        =   11
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   1395
         TabIndex        =   10
         Top             =   540
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   915
      TabIndex        =   6
      Top             =   5130
      Width           =   1590
   End
   Begin VB.Data DatVend 
      Caption         =   "Convenios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodConv, Nome_Conv From Cov Order by Nome_Conv"
      Top             =   3825
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir"
      Top             =   5130
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Filtro"
      Height          =   600
      Left            =   405
      TabIndex        =   0
      Top             =   810
      Width           =   3435
      Begin VB.OptionButton Opt 
         Caption         =   "Cadastro"
         Height          =   375
         Index           =   3
         Left            =   1890
         TabIndex        =   3
         Top             =   180
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Movimento"
         Height          =   375
         Index           =   0
         Left            =   450
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin MSDBCtls.DBCombo DbcVend 
      Bindings        =   "FormRelConv.frx":0000
      Height          =   1740
      Left            =   405
      TabIndex        =   4
      Top             =   2970
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3043
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   "Nome_Conv"
      BoundColumn     =   "CodConv"
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Convenio"
      Height          =   195
      Left            =   405
      TabIndex        =   5
      Top             =   2700
      Width           =   675
   End
End
Attribute VB_Name = "FormRelConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
'Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
'Dim Data1 As String
'Dim Data2 As String
Dim Periodo As String
Dim SQL As String
Dim CamposSql As String
Dim CondicaoWhere As String
Dim Ordenado As String

'MDIPrincipal.ComRel.ShowPrinter


'MDIPrincipal.RptRel.ReportFileName = Caminho & "\RelVend.RPT"

If DbcVend.Text = "" Then
    MsgBox "Selecione pelo menos um convênio !", vbInformation, App.Title
    Exit Sub
End If

If Opt(3).value = True Then
    CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, ENT.Fone, ENT.Convenio FROM ENT "
    Ordenado = " Order by Ent.Nome"
    CondicaoWhere = " WHERE Ent.Convenio = " & DbcVend.BoundText
    SQL = CamposSql & CondicaoWhere & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelConv.Relatorio.Recordset = varPesquisa(1)
    FrmRelConv.ReportField3.Campo = "" ' = False
    FrmRelConv.ReportField11.Caption = ""  ' = False
    FrmRelConv.ReportField5.Caption = ""  '= False
    FrmRelConv.ReportField7.Campo = "" ' = False
Else
    CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, ENT.Convenio, Sum(DUP.Valor) AS SomaDeValor, DUP.Cond FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Convenio, DUP.Cond "
    Ordenado = " Order by Ent.Nome"
    CondicaoWhere = " HAVING DUP.Cond =False  AND  ENT.Convenio = " & DbcVend.BoundText
    SQL = CamposSql & CondicaoWhere & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelConv.Relatorio.Recordset = varPesquisa(1)
    FrmRelConv.VarPerc1 = Meb1
    FrmRelConv.VarPerc2 = Meb2
    FrmRelConv.ReportField12.Caption = Meb1
    FrmRelConv.ReportField13.Caption = Meb2
'    FrmRelConv.ReportField5.Visible = True
'    FrmRelConv.ReportField7.Visible = True

End If

FrmRelConv.rpCabTitulo.Caption = "Relatório de Convênio"
FrmRelConv.rpCabSubTitulo.Caption = DbcVend.Text
FrmRelConv.Config

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DatVend.DatabaseName = Caminho & "\Dados.Mdb"

'DbcVend.Enabled = False

Meb1 = 0
Meb2 = 0

End Sub

Private Sub Opt_Click(Index As Integer)

'If Index = 3 Then
'    If Opt(3).Value = True Then
'        DbcVend.Enabled = True
'    End If

'End If

End Sub
