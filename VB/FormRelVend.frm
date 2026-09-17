VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelVend 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8685
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   11
      Top             =   3555
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
      TabIndex        =   9
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Clientes por Vendedor"
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
         Left            =   90
         TabIndex        =   10
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   3930
      TabIndex        =   8
      Top             =   3015
      Width           =   1590
   End
   Begin VB.Data DatVend 
      Caption         =   "Vendedores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3870
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7"
      Top             =   1620
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   3015
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Filtro"
      Height          =   1860
      Left            =   810
      TabIndex        =   0
      Top             =   945
      Width           =   1590
      Begin VB.OptionButton Opt 
         Caption         =   "Especifico"
         Height          =   375
         Index           =   3
         Left            =   270
         TabIndex        =   5
         Top             =   1350
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Inativos"
         Height          =   375
         Index           =   2
         Left            =   270
         TabIndex        =   3
         Top             =   990
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ativos"
         Height          =   375
         Index           =   1
         Left            =   270
         TabIndex        =   2
         Top             =   630
         Width           =   825
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin MSDBCtls.DBCombo DbcVend 
      Bindings        =   "FormRelVend.frx":0000
      Height          =   1740
      Left            =   2565
      TabIndex        =   6
      Top             =   1080
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   3043
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      Style           =   1
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor"
      Height          =   195
      Left            =   2565
      TabIndex        =   7
      Top             =   855
      Width           =   690
   End
End
Attribute VB_Name = "FormRelVend"
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

CamposSql = " SELECT ENT.Classificacao, Ent.Cpf, Ent.Cgc, Ent.Rg, Ent.InscrEst, ENT.NomeVendedor, ENT.CodEntidade, ENT.Ativo, ENT.Tipo, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, EntVend.Nome, EntVend.CodEntidade FROM ENT AS EntVend INNER JOIN ENT ON EntVend.CodEntidade = ENT.Vendedor "
Ordenado = " Order by EntVend.Nome, Ent.Cidade, Ent.Nome"

'MDIPrincipal.RptRel.ReportFileName = Caminho & "\RelVend.RPT"

CondicaoWhere = " Where (Ent.Tipo = 1 OR Ent.Tipo >= 5) "

If OPt(3).value = True Then
    CondicaoWhere = CondicaoWhere & " and EntVend.CodEntidade = " & DbcVend.BoundText + " and Ent.Ativo = true "
Else
    If OPt(1).value = True Then
            CondicaoWhere = CondicaoWhere + " and Ent.Ativo = true "
        ElseIf OPt(2).value = True Then
            CondicaoWhere = CondicaoWhere + " and Ent.Ativo = false "
    End If
End If

    SQL = CamposSql & CondicaoWhere & Ordenado

    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelVend.Relatorio.Recordset = varPesquisa(1)
    FrmRelVend.rpCabTitulo.Caption = "Relatório de Vendedores"
    FrmRelVend.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelVend.Config

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DatVend.DatabaseName = Caminho & "\Dados.Mdb"

DbcVend.Enabled = False
End Sub

Private Sub Opt_Click(Index As Integer)

If Index = 3 Then
    If OPt(3).value = True Then
        DbcVend.Enabled = True
    End If

End If

End Sub
