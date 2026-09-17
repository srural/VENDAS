VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormComposicao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7320
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11985
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   6
      Top             =   6975
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11985
      TabIndex        =   4
      Top             =   0
      Width           =   11985
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta de produto pela Composição"
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
         TabIndex        =   5
         Top             =   45
         Width           =   6990
      End
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6525
      TabIndex        =   3
      Top             =   900
      Width           =   1275
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   900
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Data DatDados 
      Caption         =   "Produto"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   3960
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3555
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormComposicao.frx":0000
      Height          =   5340
      Left            =   225
      OleObjectBlob   =   "FormComposicao.frx":001A
      TabIndex        =   1
      Top             =   1350
      Width           =   11580
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormComposicao.frx":0D62
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   900
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "descrição_Produto"
      BoundColumn     =   "CodPrd"
      Text            =   ""
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   675
      Width           =   645
   End
End
Attribute VB_Name = "FormComposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Dim Componentes As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select CodPrd, Descrição_Produto From PRD Order By Descrição_Produto"
DatDados(0).Refresh

Me.Left = 200 '(MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 '(MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub DbcPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If DbcPesquisa <> "" And Not IsNumeric(DbcPesquisa.Text) Then
    Criterio = "Prd.descrição_produto Like " & """" & DbcPesquisa & "*"""
    DatDados(0).Recordset.FindFirst Criterio
    If DatDados(0).Recordset.NoMatch Then
        Beep
        DbcPesquisa.SetFocus
    Else
        txtCod = DatDados(0).Recordset.Fields("CodPrd")
        DatDados(1).RecordSource = "SELECT Comp.Qtd, COMP.PrdComp, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.CodPrd FROM PRD INNER JOIN COMP ON PRD.CodPrd = COMP.Produto WHERE COMP.PrdComp = " & txtCod & " ORDER BY PRD.Descrição_Produto"

        DatDados(1).Refresh
    End If
End If

End Sub
