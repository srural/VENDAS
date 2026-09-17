VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmLanca 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8565
   ClientLeft      =   4095
   ClientTop       =   3075
   ClientWidth     =   17130
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   17130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton fechaparcial 
      BackColor       =   &H8000000A&
      Caption         =   "&Fecha Parcial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7515
      Width           =   1620
   End
   Begin VB.Frame FrameEnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cadastro"
      Height          =   2625
      Left            =   3915
      TabIndex        =   28
      Top             =   1935
      Visible         =   0   'False
      Width           =   7665
      Begin VB.CommandButton CmdSalvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salvar"
         Height          =   375
         Left            =   6075
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2025
         Width           =   1320
      End
      Begin VB.TextBox Email 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Email"
         Height          =   315
         Left            =   945
         MaxLength       =   60
         TabIndex        =   33
         Tag             =   "S"
         Top             =   1440
         Width           =   6435
      End
      Begin VB.TextBox Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Nome"
         Height          =   315
         Left            =   945
         MaxLength       =   60
         TabIndex        =   29
         Tag             =   "S"
         Top             =   540
         Width           =   6435
      End
      Begin MSMask.MaskEdBox Fone 
         DataField       =   "Celular"
         Height          =   315
         Left            =   5625
         TabIndex        =   32
         Top             =   990
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         Mask            =   "(##)#.####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cpf 
         DataField       =   "CPF"
         DataSource      =   "DatDados(1)"
         Height          =   315
         Left            =   945
         TabIndex        =   30
         Tag             =   "S"
         Top             =   990
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Rg 
         DataField       =   "RG"
         DataSource      =   "DatDados(1)"
         Height          =   315
         Left            =   3330
         TabIndex        =   31
         Tag             =   "S"
         Top             =   990
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.###.###-C"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG:"
         Height          =   195
         Index           =   29
         Left            =   2880
         TabIndex        =   39
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Index           =   55
         Left            =   315
         TabIndex        =   38
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Index           =   37
         Left            =   435
         TabIndex        =   37
         Top             =   990
         Width           =   345
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         Height          =   195
         Index           =   8
         Left            =   5175
         TabIndex        =   36
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   35
         Top             =   540
         Width           =   465
      End
   End
   Begin VB.CommandButton BtPesquisa 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmLanca.frx":0000
      Height          =   435
      Left            =   7740
      Picture         =   "FrmLanca.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Pesquisar"
      Top             =   1080
      Width           =   555
   End
   Begin VB.Data DatPesquisaG 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   630
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox txtcomplemento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8415
      TabIndex        =   24
      Top             =   1080
      Width           =   2505
   End
   Begin VB.CommandButton CmbTotaliza 
      BackColor       =   &H8000000B&
      Caption         =   "&Confirma  [F9]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3330
      Width           =   2295
   End
   Begin VB.CommandButton CmbImprimi 
      BackColor       =   &H8000000A&
      Caption         =   "&Ver Conta  [F8]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3945
      Width           =   2295
   End
   Begin VB.CommandButton CmbLista 
      BackColor       =   &H8000000A&
      Caption         =   "&Lista Conta  [F7]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4605
      Width           =   2295
   End
   Begin VB.CommandButton CmbFecha 
      BackColor       =   &H8000000A&
      Caption         =   "&Fecha Conta  [F6]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton CmdLimpa 
      BackColor       =   &H8000000A&
      Caption         =   "&Limpa  [F10]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5940
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -315
      ScaleHeight     =   330
      ScaleWidth      =   17445
      TabIndex        =   16
      Top             =   8235
      Width           =   17475
   End
   Begin VB.TextBox TxtNroMesa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10710
      TabIndex        =   14
      Top             =   7515
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11025
      TabIndex        =   15
      Top             =   7605
      Width           =   1590
   End
   Begin VB.CommandButton Transfere 
      Caption         =   "&Transfere Mesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8460
      Picture         =   "FrmLanca.frx":0A64
      TabIndex        =   13
      Top             =   7515
      Width           =   2070
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7695
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSDBCtls.DBCombo DbcProd 
      Bindings        =   "FrmLanca.frx":0D6E
      DataSource      =   "DatDados(0)"
      Height          =   5925
      Left            =   2835
      TabIndex        =   12
      Top             =   1890
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   10451
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   "Descrição_Produto"
      BoundColumn     =   "CodPrd"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   17145
      TabIndex        =   10
      Top             =   0
      Width           =   17145
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "LANÇA MESAS"
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
         TabIndex        =   11
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.TextBox TxtCartao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   180
      TabIndex        =   0
      Top             =   1035
      Width           =   1200
   End
   Begin VB.TextBox TxtDadosLanca 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   2655
      Width           =   2310
   End
   Begin VB.TextBox TxtDadosLanca 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   1890
      Width           =   2310
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\AATESTE\AATesteBica\Posto\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7110
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSComctlLib.ListView ListVda 
      Height          =   5370
      Left            =   8415
      TabIndex        =   3
      Top             =   1890
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   9472
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CodPrd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Produto"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qtd"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Unitário"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Complemento"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSDBCtls.DBCombo TxtPessoa 
      Bindings        =   "FrmLanca.frx":0D88
      Height          =   405
      Left            =   1620
      TabIndex        =   27
      Top             =   1080
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8415
      TabIndex        =   25
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label MebTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   14325
      TabIndex        =   18
      Top             =   7650
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome (opcional)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1620
      TabIndex        =   17
      Top             =   810
      Width           =   1650
   End
   Begin VB.Label LblCard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1500
      Left            =   225
      TabIndex        =   9
      Top             =   6615
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa/Cartão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   795
      Width           =   1185
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   16200
      TabIndex        =   5
      Top             =   7335
      Width           =   570
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade / Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   1620
      Width           =   1965
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto / Cód. Barra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   2430
      Width           =   1950
   End
   Begin VB.Label Mens 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   8175
      Width           =   270
   End
End
Attribute VB_Name = "FrmLanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Banco As Database
Public Formulario As Form
Public Registros As DAO.Recordset
Public Clientes As DAO.Recordset
Public Vendas As DAO.Recordset
Public Bicos As DAO.Recordset
Public Bico As DAO.Recordset
Public Encerrante As DAO.Recordset
'Public Ofertas As Dao.RecordSet
Public TotTemp As DAO.Recordset
Public Linha As Integer
Public ContPro As Integer
Public Soma As Double
'Public Caminho As String
Public Flag2 As Boolean
Public Pessoa As String
Public NroCupom As String
Public Itens As Integer
Public Ponto As Integer
Public Config As DAO.Recordset
Public PistaAnt As String
Public PosicaoAba As Long
Public LstPesq As ListItem
Public PrdAnt As Long
Public Varvda As Recordset
Public VarProdutos As Recordset
Public VarFlagVer As Boolean
Public VarX As Integer

Public CodCard As Double

'Public Automa As Dao.RecordSet

Private Sub CmbCalc_Click()
    Shell ("calc")
End Sub

Private Sub CmbCancela_Click()

Dim sql As String

If MsgBox("Confirma cancelamento da venda ?", vbYesNo) = vbYes Then
    ListVda.ListItems.Clear
        sql = "UPDATE Conv SET Cancelado = -1, Totaliza = -1   WHERE Cliente = " & CodCard & " AND Flag = 0 " ' AND Maquina = " & """" & LimpaCampo(Left(Maquina, 15)) & """"
        FrmFrente.Banco.Execute sql
'        DatDados(1).Refresh
        LimpaTela
Else
    txtdadoslanca_KeyPress 2, 13
End If
End Sub



Private Sub BtPesquisa_Click()
FrameEnt.Visible = True
Nome.SetFocus
End Sub

Private Sub CmbFecha_Click()
If MsgBox("Confirma fechamento da mesa " & LblCard, vbYesNo, App.Path) = vbYes Then

    Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
    
    VarConv.Requery
    If VarConv.RecordCount > 0 Then
        If VarConv!Cupom <> 0 Then
            FrmFrente.CliCartao = VarConv!Cupom
        End If
    End If

    FrmFrente.TxtTmp2 = Trim(LblCard.Caption)
'    FrmFrente.TxtSenha = Trim(LblCard.Caption)
    FrmFrente.TxtDados(0) = "/" & Trim(LblCard.Caption)
    SendKeys ("{ENTER}")
    Unload Me

'    txtdadoslanca(0).SetFocus

End If
End Sub

Private Sub CmbImprimi_Click()
Dim VarConv As Recordset
'Dim soma As Currency
Soma = 0
VarFlagVer = True
VarX = 0
If TxtCartao <> "" And MebTotal = "" Then
    If IsNumeric(TxtCartao) Then
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Sum(CONV.Valor) AS SomaDeValor, CONV.Cliente, CONV.Flag, CONV.Totaliza From CONV GROUP BY CONV.Cliente, CONV.Flag, CONV.Totaliza HAVING cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 ")
        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
        sql = "SELECT Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA "
        VarConv.Requery
        If VarConv.RecordCount > 0 Then
'                MebTotal = VarConv!somadevalor
            VarConv.MoveFirst
'            txtdadoslanca(0).Enabled = False
'            txtdadoslanca(2).Enabled = False
            ListVda.ListItems.Clear
                        
            While Not VarConv.EOF
                
                If VarConv.Fields("Valor") <> 0 Then
                    Set LstPesq = ListVda.ListItems.Add(, , VarConv!CodVda)

                
                    LstPesq.SubItems(1) = VarConv.Fields("Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(2) = VarConv.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(3) = Numero(Format(VarConv.Fields("Qtd"), " ##,##0.00"), 2)  'SeCampoNulo(Consulta, "Conta")
                    LstPesq.SubItems(4) = Numero(Format(VarConv.Fields("Valor") / VarConv.Fields("qtd"), "##,##0.00"), 2)  'SeCampoNulo(Consulta, "DataOper")
                    LstPesq.SubItems(5) = Format(VarConv.Fields("Valor"), "##,##0.00")  'SeCampoNulo(Consulta, "BomPara")
'                    LstPesq.SubItems(6) = VarConv.Fields("Cupom")   'SeCampoNulo(Consulta, "Agencia")
                    Soma = Soma + VarConv!Valor
                    MebTotal.Caption = Format(Soma, "     ##,##0.00")
                    VarX = VarX + 1
                End If
                VarConv.MoveNext
            Wend
            VarX = VarX + 1
        End If
    End If
End If
'Text1.SetFocus
End Sub

Private Sub CmbLista_Click()
'CmbImprimi_Click
Dim Comissao As Currency
Dim VarConv As Recordset
'Dim soma As Currency
Soma = 0

If TxtCartao <> "" And MebTotal = "" Then
    If IsNumeric(TxtCartao) Then
    
        sql = "UPDATE Conv SET funcionario = 99999   WHERE Cliente = " & TxtCartao
        FrmFrente.Banco.Execute sql

'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Sum(CONV.Valor) AS SomaDeValor, CONV.Cliente, CONV.Flag, CONV.Totaliza From CONV GROUP BY CONV.Cliente, CONV.Flag, CONV.Totaliza HAVING cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 ")
        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT CONV.Funcionario, CONV.Produto, First(CONV.CodVda) AS PrimeiroDeCodVda, Sum(CONV.Valor) AS SomaDeValor, Sum(CONV.Qtd) AS SomaDeQtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd GROUP BY Conv.Funcionario, CONV.Produto, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto, CONV.Cancelado HAVING (((CONV.Cliente)= " & TxtCartao & ") AND ((CONV.Flag)=-1) AND ((CONV.Totaliza)=0) AND ((CONV.Cancelado)=0)) ORDER BY First(CONV.CodVda)")
        
        VarConv.Requery
        If VarConv.RecordCount > 0 Then
                                  
'                MebTotal = VarConv!somadevalor
            VarConv.MoveFirst
            TxtDadosLanca(0).Enabled = False
            TxtDadosLanca(2).Enabled = False
            ListVda.ListItems.Clear
                        
            While Not VarConv.EOF
'                VarConv.Edit
'                VarConv!Funcionario = 1
'                VarConv.Update
                If VarConv.Fields("SomaDeValor") <> 0 Then
                    Set LstPesq = ListVda.ListItems.Add(, , VarConv!PrimeiroDeCodVda)
                    
                    LstPesq.SubItems(1) = VarConv.Fields("Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(2) = VarConv.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(3) = Numero(Format(VarConv.Fields("SomaDeQtd"), " ##,##0.00"), 2)  'SeCampoNulo(Consulta, "Conta")
                    LstPesq.SubItems(4) = Numero(Format(VarConv.Fields("SomaDeValor") / VarConv.Fields("SomaDeqtd"), "##,##0.00"), 2)  'SeCampoNulo(Consulta, "DataOper")
                    LstPesq.SubItems(5) = Format(VarConv.Fields("SomaDeValor"), "##,##0.00")  'SeCampoNulo(Consulta, "BomPara")
                    Soma = Soma + VarConv!SomaDeValor
                    MebTotal.Caption = Format(Soma, "     ##,##0.00")

                End If
                VarConv.MoveNext
            Wend
            ListVda.Enabled = False
        End If
    End If
End If

Iniciar (False)
    
    Print #1, "========================================"
    If VarImpressoraNf <> "TERMICA" Then
        Print #1, Gde & "MESA ====> " & LblCard & Nor
    Else
        Print #1, "MESA ====> " & LblCard
    End If
    Print #1, "----------------------------------------"
    Print #1, ""
    For x = 1 To ListVda.ListItems.Count
'            Print #1, Lista.List(x)
'            Print #1, Lista.List(x)
'             Print #1, ListVda.ListItems.Item(X).Text; Tab(1);
'             Print #1, Lista.ListItems.Item(x).SubItems(1); Tab(51);
             Print #1, Alinhar(ListVda.ListItems.Item(x).SubItems(3), 5, 1); Tab(7);
             Print #1, Left(ListVda.ListItems.Item(x).SubItems(2), 25); Tab(34);
             Print #1, Alinhar(ListVda.ListItems.Item(x).SubItems(5), 7, 2)
            
'            Linha = Linha + 1
    Next
    
    If VarComissao = 0 Then
        Print #1, Tab(29); "------------"
        Print #1, Tab(18); "  TOTAL =>>>>>>  "; Format(MebTotal, " ##0.00")
    Else
        Comissao = (CCur(MebTotal) * (VarComissao / 100))
        Print #1, Tab(29); "------------"
        Print #1, Tab(18); "SUB-TOTAL =>>"; Alinhar((MebTotal), 10, 2)
        Print #1, Tab(18); "    10 %  =>>"; Alinhar(CStr(Comissao), 10, 2)
        Print #1, Tab(18); "    TOTAL =>>"; Alinhar(CStr(CCur(MebTotal) + Comissao), 10, 2)
     End If
    
    Print #1, ""
    Print #1, "========================================"
    Print #1, "       Agradecemos a Preferencia      "
    For x = 1 To NroLinha
        Print #1, ""
    Next
    Print #1, Chr(27) & "i"

Fim (False)
'Text1.SetFocus
End Sub

Private Sub CmbTotaliza_Click()
Dim sql As String
Dim Flag As Boolean
Dim VarCont As Integer
Flag = False
If VarFlagVer = False Then
    VarCont = 1
Else
    VarCont = VarX
End If


If Soma > 0 Then
'    If MsgBox("Confirma ?", vbYesNo) = vbYes Then
        sql = "UPDATE Conv SET Flag = -1   WHERE Cliente = " & CodCard & " AND Cancelado = 0  " 'AND Maquina = " & """" & LimpaCampo(Left(Maquina, 15)) & """"
        FrmFrente.Banco.Execute sql
    '        DatDados(1).Refresh
        If VarImprCozinha <> "" Then
            
          For x = VarCont To ListVda.ListItems.Count
             Criterio = "CodPrd = " & ListVda.ListItems.Item(x).SubItems(1)
             VarProdutos.FindFirst Criterio
             If Not VarProdutos.NoMatch Then
                 If VarProdutos.Fields("Local") = "1" Then
                    Flag = True
                 End If
              End If
          Next
          If Flag = True Then
              Open VarImprCozinha For Output As #1
    
                
                Print #1, "========================================"
 '               If VarImpressoraNf <> "TERMICA" Then
 '                   Print #1, Gde & "MESA ====> " & LblCard & "  - " & TxtPessoa & Nor
 '               Else
                    Print #1, "MESA ====> " & LblCard & "  - " & Left(TxtPessoa, 13); Tab(31); "- " & Time
 '               End If
                Print #1, "----------------------------------------"
                Print #1, ""
                For x = VarCont To ListVda.ListItems.Count
                
                    Criterio = "CodPrd = " & ListVda.ListItems.Item(x).SubItems(1)
                    VarProdutos.FindFirst Criterio
                    If Not VarProdutos.NoMatch Then
                       If VarProdutos.Fields("Local") = "1" Then
            '              Print #1, Lista.List(x)
            '              Print #1, Lista.List(x)
            '              Print #1, ListVda.ListItems.Item(X).Text; Tab(1);
            '              Print #1, Lista.ListItems.Item(x).SubItems(1); Tab(51);
                           Print #1, ListVda.ListItems.Item(x).SubItems(3); Tab(7);
                           Print #1, Left(ListVda.ListItems.Item(x).SubItems(2), 28)
                           Print #1, Left(ListVda.ListItems.Item(x).SubItems(6), 28)
                
            '              Print #1, ListVda.ListItems.Item(x).SubItems(4); Tab(34);
    '                      Print #1, Alinhar(ListVda.ListItems.Item(x).SubItems(5) / 100, 7, 2)
            '              Print #1, Lista.ListItems.Item(x).SubItems(4)
                        
            '              Linha = Linha + 1
                       End If
                    End If
                Next
                Print #1, ""
                Print #1, "========================================"
                For x = 1 To NroLinha
                    Print #1, ""
                Next
                Print #1, Chr(27) & "i"  '& Chr(27) & "v50"
                            
                Close #1
                    
            End If
        End If
        Flag = False
        If VarImprBar <> "" Then
            
          For x = VarCont To ListVda.ListItems.Count
             Criterio = "CodPrd = " & ListVda.ListItems.Item(x).SubItems(1)
             VarProdutos.FindFirst Criterio
             If Not VarProdutos.NoMatch Then
                 If VarProdutos.Fields("Local") = "2" Then
                    Flag = True
                 End If
              End If
          Next
          If Flag = True Then
              Open VarImprBar For Output As #1
    
                
                Print #1, "========================================"
'                If VarImpressoraNf <> "TERMICA" Then
'                    Print #1, Gde & "MESA ====> " & LblCard & "  - " & TxtPessoa & Nor
'                Else
                    Print #1, "MESA ====> " & LblCard & "  - " & Left(TxtPessoa, 13); Tab(31); "- " & Time
'                End If
                Print #1, "----------------------------------------"
                Print #1, ""
                For x = VarCont To ListVda.ListItems.Count
                
                    Criterio = "CodPrd = " & ListVda.ListItems.Item(x).SubItems(1)
                    VarProdutos.FindFirst Criterio
                    If Not VarProdutos.NoMatch Then
                       If VarProdutos.Fields("Local") = "2" Then
            '              Print #1, Lista.List(x)
            '              Print #1, Lista.List(x)
            '              Print #1, ListVda.ListItems.Item(X).Text; Tab(1);
            '              Print #1, Lista.ListItems.Item(x).SubItems(1); Tab(51);
                           Print #1, ListVda.ListItems.Item(x).SubItems(3); Tab(7);
                           Print #1, Left(ListVda.ListItems.Item(x).SubItems(2), 28)
                           Print #1, Left(ListVda.ListItems.Item(x).SubItems(6), 28)
            '              Print #1, ListVda.ListItems.Item(x).SubItems(4); Tab(34);
    '                      Print #1, Alinhar(ListVda.ListItems.Item(x).SubItems(5) / 100, 7, 2)
            '              Print #1, Lista.ListItems.Item(x).SubItems(4)
                        
            '              Linha = Linha + 1
                       End If
                    End If
                Next
                Print #1, ""
                Print #1, "========================================"
                For x = 1 To NroLinha
                    Print #1, ""
                Next
                Print #1, Chr(27) & "i"  '& Chr(27) & "v50"
                            
                Close #1
                    
            End If
        End If
        
        If Soma > 0 Then
            If MsgBox("Confirma Impressão ?", vbYesNo + vbDefaultButton2) = vbYes Then
                ImprimiCta VarCont
            Else
                ListVda.ListItems.Clear
                LimpaTela
                Unload Me
            End If
        Else
            ListVda.ListItems.Clear
            LimpaTela
            Unload Me
        End If

'    Else
'        txtdadoslanca_KeyPress 2, 13
'    End If
    CmbImprimi.Enabled = True
    CmbLista.Enabled = True
    CmbFecha.Enabled = True

Else
    sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom, Qtd, Produto, Flag,  Valor ) Values ("
    sql = sql & CodCard & " ,"
    sql = sql & """" & LimpaCampo(Left(TxtPessoa, 15)) & """" & " , #"
    sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
    sql = sql & """" & CStr(Time) & """" & ", "
    sql = sql & IIf(IsNumeric(txtcomplemento.Text), txtcomplemento, 0) & " , "
    sql = sql & 0 & " , "
    sql = sql & 0 & " , "
    sql = sql & -1 & " , "
    sql = sql & Numero(0, 2) & ")"
    
'    FrmFrente.Banco.Execute sql
    
    Varvda.Requery
'     VarVda.Bookmark = VarVda.LastModified
    
    Varvda.MoveLast
    Unload Me
End If

End Sub


Private Sub ImprimiCta(Contador As Integer)
Dim x As Integer

Iniciar (False)
    
    Print #1, "========================================"
    If VarImpressoraNf <> "TERMICA" Then
    
        Print #1, "MESA ====> " & LblCard & "  - " & Left(TxtPessoa, 12); Tab(30); "- " & Time
    Else
'        Print #1, "MESA ====> " & LblCard & "  - " & TxtPessoa
        Print #1, "MESA ====> " & LblCard & "  - " & Left(TxtPessoa, 12); Tab(30); "- " & Time
    End If
    Print #1, "----------------------------------------"
    Print #1, ""
    For x = Contador To ListVda.ListItems.Count
'            Print #1, Lista.List(x)
'            Print #1, Lista.List(x)
'             Print #1, ListVda.ListItems.Item(X).Text; Tab(1);
'             Print #1, Lista.ListItems.Item(x).SubItems(1); Tab(51);
             Print #1, ListVda.ListItems.Item(x).SubItems(3); Tab(7);
             Print #1, Left(ListVda.ListItems.Item(x).SubItems(2), 28); Tab(34);
'             Print #1, ListVda.ListItems.Item(x).SubItems(4); Tab(34);
             Print #1, Alinhar(ListVda.ListItems.Item(x).SubItems(5) / 100, 7, 2)
'             Print #1, Lista.ListItems.Item(x).SubItems(4)
            
'            Linha = Linha + 1
    Next
    Print #1, ""
    Print #1, "========================================"
    For x = 1 To NroLinha
        Print #1, ""
    Next
    Print #1, Chr(27) & "i" '& Chr(27) & "v50"
    

Fim (False)
ListVda.ListItems.Clear
LimpaTela
Unload Me
FrmFrente.AtuMesa

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdLimpa_Click()
'LimpaTela
End Sub

Private Sub CmdSalvar_Click()
If Nome <> "" Then
    DatPesquisaG.Recordset.AddNew
    DatPesquisaG.Recordset.Fields("Nome") = Nome
    DatPesquisaG.Recordset.Fields("Cpf") = Cpf
    DatPesquisaG.Recordset.Fields("Rg") = Rg
    DatPesquisaG.Recordset.Fields("Fone") = Fone
    DatPesquisaG.Recordset.Fields("email") = Email
    DatPesquisaG.Recordset.Fields("Tipo") = 1
    DatPesquisaG.Recordset.Fields("Classificacao") = 1
    DatPesquisaG.Recordset.Fields("Ativo") = 1
    DatPesquisaG.Recordset.Update
    MsgBox "Cliente Gravado com Sucesso !", vbInformation
    TxtPessoa.Text = Nome
    FrameEnt.Visible = False
    FrmFrente.Clientes.Requery
    FrmFrente.Registros.Requery
    FrmFrente.DatCli.Recordset.Requery
    
    TxtPessoa.SetFocus
Else
    FrameEnt.Visible = False
    TxtPessoa.SetFocus
End If

End Sub

Private Sub DbcProd_GotFocus()
    DbcProd.BackColor = &HC0FFFF
    DatDados(0).Recordset.MoveFirst
End Sub

Private Sub DbcProd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If IsNumeric(DbcProd.BoundText) Then
        If TxtDadosLanca(0).Text = "" Then
            TxtDadosLanca(0).Text = 1
        End If
        
        If TxtDadosLanca(0).Text = 0 Then
            TxtDadosLanca(0).Text = 1
        End If
        
        EfetuaVenda (DbcProd.BoundText)
'        SendKeys ("{ENTER}")
        TxtDadosLanca(0).Text = ""
        TxtDadosLanca(2).Text = ""
        TxtDadosLanca(0).SetFocus
    End If
End If
End Sub

Private Sub DbcProd_LostFocus()
DbcProd.BackColor = &H80000005
End Sub

Private Sub DblPesquisaG_DblClick()
Dim sql As String
'FramePesquisaG.Visible = False
sql = "Nome = " & """" & DblPesquisaG.BoundText & """"
DatPesquisaG.Recordset.FindFirst sql
TxtPessoa = DatPesquisaG.Recordset.Fields("Nome")
txtcomplemento = DatPesquisaG.Recordset.Fields("CodEntidade")
'MebDados(1) = DatPesquisaG.Recordset.Fields("CEP")
'MebDados(16) = DatPesquisaG.Recordset.Fields("CodCidade")
TxtDadosLanca(0).SetFocus
End Sub

Private Sub fechaparcial_Click()
Dim sql As String
sql = "DELETE CONV.*, CONV.Cliente FROM CONV WHERE CONV.Cliente= 204"

If MsgBox("Confirma fechamento parcial da mesa  " & LblCard, vbYesNo, App.Path) = vbYes Then
    FrmFrente.Banco.Execute sql
    For x = 1 To ListVda.ListItems.Count
        
        If ListVda.ListItems(x).Checked Then
                        
            sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom,Flag, Qtd, Produto,  Valor ) Values ("
            sql = sql & 204 & " ,"
            sql = sql & """" & LimpaCampo(Left(Pessoa, 15)) & """" & " , #"
            sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
            sql = sql & """" & CStr(Time) & """" & ", "
            sql = sql & 0 & " , "
            sql = sql & -1 & ", "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(3) / 100, " ##,##0.00"), 2) & " , "
            sql = sql & ListVda.ListItems.Item(x).SubItems(1) & " , "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(5), "##,##0.00"), 2) & ")"
            FrmFrente.Banco.Execute sql
            
            Criterio = "CodVda = " & ListVda.ListItems.Item(x)
            Varvda.FindFirst Criterio
            If Not Varvda.NoMatch Then
                Varvda.Edit
                Varvda!cancelado = -1
                Varvda.Update
            End If
            
        End If
    Next
    ListVda.ListItems.Clear
    LimpaTela
    
    Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & 204 & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
    
    VarConv.Requery
    If VarConv.RecordCount > 0 Then
        If VarConv!Cupom <> 0 Then
            FrmFrente.CliCartao = VarConv!Cupom
        End If
    End If

    FrmFrente.TxtTmp2 = 204
'    FrmFrente.TxtSenha = Trim(LblCard.Caption)
    FrmFrente.TxtDados(0) = "/" & 204
    Unload Me
'    txtdadoslanca(0).SetFocus
    SendKeys ("{ENTER}")
    
End If

End Sub

Private Sub Form_Activate()
If TxtCartao = "" Then
    TxtCartao.SetFocus
Else
    TxtDadosLanca(0).SetFocus
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmFrente.Timer2.Enabled = True
End Sub

Private Sub ListVda_DblClick()
Dim Criterio As String
Dim VarItem As String
If ListVda.ListItems.Count > 0 Then
    If MsgBox("Confirma exlusão deste item ? " & ListVda.SelectedItem.Text, vbYesNo, App.Title) = vbYes Then
        VarItem = ListVda.SelectedItem.Index
        Criterio = "CodVda = " & ListVda.SelectedItem.Text
        Varvda.FindFirst Criterio
        If Not Varvda.NoMatch Then
            Soma = MebTotal
            Soma = Soma - Varvda("Valor")
            MebTotal = Soma
            ListVda.ListItems.Remove CDec(VarItem)
            
            Varvda.Edit
            Varvda!cancelado = -1
            VarX = VarX - 1
'            VarVda.Delete
            Varvda.Update
        End If
    End If
End If
If TxtDadosLanca(0).Enabled = True Then
    TxtDadosLanca(0).SetFocus
End If
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim PesquisaSql As String
Dim Retorno As Integer
Dim Status As String
Dim x As Integer
Dim P As Long
Dim Pista As String
Dim Ponto1 As String
Dim Ponto2 As String
Dim Ponto3 As String
Dim Ponto4 As String
Dim Comando As String
Dim VarCodigo As Double
PistaAnt = ""
Caminho = App.Path
VarFlagVer = False

'If App.PrevInstance Then
'    MsgBox "Existe uma aplicação do " & App.ProductName & " aberta nesta máquina, um ou mais programas abertos pode causar lentidão no sistema!", 16
'    End
'End If

'FrmFrente.Timer2.Enabled = False
'If Dir(App.Path & "\Config.ini") <> "" Then
    
'    Caminho = LerINI("Config", "LocalBco", App.Path & "\config.ini")
'    Impressora = LerINI("Config", "Ecf", App.Path & "\config.ini")
'    PortaF = LerINI("Config", "Porta", App.Path & "\config.ini")
'    PortaNF = LerINI("Config", "Impr", App.Path & "\config.ini")
'    Leitor = LerINI("Config", "Leitor", App.Path & "\config.ini")
'    VarTef = LerINI("Config", "Tef", App.Path & "\config.ini")
'    VarTipoVda = LerINI("Config", "TipoVda", App.Path & "\config.ini")
'    VarTermica = LerINI("Config", "Termica", App.Path & "\config.ini")
'    VarComprovante = LerINI("Config", "Comprovante", App.Path & "\config.ini")
'    VarPrecoLiberado = LerINI("Config", "PrecoLiberado", App.Path & "\config.ini")
'    VarTipoPreco = LerINI("Config", "TipoPreco", App.Path & "\config.ini")
'
'End If

If dir(App.Path & "Cond.Txt") <> "" Then
    Nor = Chr(18)
    Cond = Chr(15)
Else
    Nor = ""
    Cond = ""
End If
Gde = Chr(14)
Negi = Chr(27) + Chr(69)
NegF = Chr(27) + Chr(70)
SaltoPag = Chr(12)

'FrmFrente.Caminho = App.Path
Soma = 0
For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

'P = NameOfPC(Maquina)

'Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Dados.mdb")
'Set Config = Banco.OpenRecordset("Select Cfg.* FROM Cfg")

'Config.MoveFirst
'Porta = Config.Fields("Porta")

DatPesquisaG.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"


'If FramePesquisaG.Visible = False Then
'    If TxtPessoa.Text <> "" Then
'       PesquisaGeral FrmFrente.Caminho & "\dados.mdb", "SELECT CodEntidade, Nome From Ent  where Nome like " & """" & "*" & TxtPessoa.Text & "*" & """" & "  ORDER BY Nome", "Nome"
'    Else
       PesquisaGeral FrmFrente.Caminho & "\dados.mdb", "SELECT CodEntidade, Nome, Cpf, Rg, Ativo, Tipo, Classificacao, Fone , Email From Ent  ORDER BY Nome", "Nome"
'    End If
'Else
'    FramePesquisaG.Visible = False
'End If


Set Registros = FrmFrente.Banco.OpenRecordset("SELECT PRD.* , BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto  WHERE Ativo = true ORDER BY PRD.CodPrd ")

Set Vendas = FrmFrente.Banco.OpenRecordset("SELECT * FROM CONV")

Set Varvda = FrmFrente.Banco.OpenRecordset("Select * From Conv Order By CodVda")

Set VarProdutos = FrmFrente.Banco.OpenRecordset("Select CodPrd, Concorrencia, Local From Prd Order By CodPrd")

'DatDados(1).RecordSource = "SELECT vda.pedido, vda.bico, Vda.Unitario, Prd.Estoque, Vda.BaixarEst, vda.cliente, VDA.CodAba, VDA.CodVda, Prd.CodPrd, PRD.Descrição_Produto, VDA.Qtd, PRD.Preco, VDA.Valor, Vda.Maquina FROM PRD INNER JOIN VDA ON PRD.CodPrd = VDA.Produto WHERE VDA.Flag = False and Vda.Maquina = '" & LimpaCampo(Maquina) & "' ORDER BY VDA.CodVda"
'DatDados(1).Refresh

LblHora = Time
LblData = Date

DatDados(0).RecordSource = "SELECT CodPrd, Descrição_Produto FROM Prd  WHERE Ativo = true ORDER BY Descrição_produto"
DatDados(0).Refresh

'sql = "UPDATE Conv SET Conv.Cancelado = -1   WHERE Flag = 0  " 'AND Maquina = " & """" & Trim(LimpaCampo(Left(FrmFrente.Maquina, 15))) & """"
'FrmFrente.Banco.Execute sql

'If TxtCartao <> "" Then
'    FrmFrente.Banco.Execute "UPDATE CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd SET CONV.Valor = [prd].[venda] WHERE CONV.Totaliza)=False AND CONV.Cliente = " & TxtCartao
'End If


End Sub

Function LimpaTela()
TxtDadosLanca(0) = ""

txtcomplemento = ""
MebTotal = ""
TxtCartao = ""
Soma = 0
ContPro = 0
Valor = 0
txtpgto = ""
LblCard = ""
CmbImprimi.Enabled = True
CmbLista.Enabled = True
CmbFecha.Enabled = True


'txtdadoslanca(2).SetFocus
'txtdadoslanca(0).SetFocus

TxtDadosLanca(0).Enabled = True
TxtDadosLanca(0).SetFocus
TxtDadosLanca(2).Enabled = True
TxtDadosLanca(2).SetFocus
TxtCartao.Enabled = True
TxtCartao.SetFocus

Linha = 0
Itens = 0
ListVda.ListItems.Clear

End Function


Private Sub Timer1_Timer()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then
    LimpaTela
End If

If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Transfere_Click()
If TxtNroMesa = "" Then
    MsgBox "Indentifique o numero da mesa destino ", vbInformation
    TxtNroMesa.SetFocus
Else
    For x = 1 To ListVda.ListItems.Count
        If ListVda.ListItems(x).Checked Then
                        
            sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom,Flag, Qtd, Produto,  Valor ) Values ("
            sql = sql & TxtNroMesa & " ,"
            sql = sql & """" & LimpaCampo(Left(Pessoa, 15)) & """" & " , #"
            sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
            sql = sql & """" & CStr(Time) & """" & ", "
            sql = sql & 0 & " , "
            sql = sql & -1 & ", "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(3) / 100, " ##,##0.00"), 2) & " , "
            sql = sql & ListVda.ListItems.Item(x).SubItems(1) & " , "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(5), "##,##0.00"), 2) & ")"
            FrmFrente.Banco.Execute sql
            
            Criterio = "CodVda = " & ListVda.ListItems.Item(x)
            Varvda.FindFirst Criterio
            If Not Varvda.NoMatch Then
                Varvda.Edit
                Varvda!cancelado = -1
                Varvda.Update
            End If
            
        End If
    Next
    ListVda.ListItems.Clear
    LimpaTela
    Unload Me
    
  
End If

End Sub

Private Sub TxtCartao_GotFocus()
TxtCartao.BackColor = &HC0FFFF
End Sub

Private Sub TxtCartao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub TxtCartao_KeyPress(KeyAscii As Integer)
'Dim VarConv As Recordset
'Dim Soma As Currency
'Soma = 0
If KeyAscii = 13 And TxtCartao <> "" Then
    If IsNumeric(TxtCartao) Then
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Sum(CONV.Valor) AS SomaDeValor, CONV.Cliente, CONV.Flag, CONV.Totaliza From CONV GROUP BY CONV.Cliente, CONV.Flag, CONV.Totaliza HAVING cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 ")
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ")
        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT CONV.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, Conv.Maquina FROM CONV WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
        
        VarConv.Requery
        If VarConv.RecordCount > 0 Then
            If IsNull(VarConv!Maquina) Then
                TxtPessoa = ""
            Else
                TxtPessoa = VarConv!Maquina
            End If
            If VarConv!Cupom <> 0 Then
                txtcomplemento = VarConv!Cupom
            End If
            
        End If
'                MebTotal = VarConv!somadevalor
'            VarConv.MoveFirst
'            While Not VarConv.EOF
            
'                Set LstPesq = ListVda.ListItems.Add(, , VarConv!CodVda)
                
'                LstPesq.SubItems(1) = VarConv.Fields("Produto")   'SeCampoNulo(Consulta, "Agencia")
'                LstPesq.SubItems(2) = VarConv.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
'                LstPesq.SubItems(3) = Numero(Format(VarConv.Fields("Qtd"), " ##,##0.000"), 3)  'SeCampoNulo(Consulta, "Conta")
'                LstPesq.SubItems(4) = Numero(Format(VarConv.Fields("Valor") / VarConv.Fields("qtd"), "##,##0.000"), 3)  'SeCampoNulo(Consulta, "DataOper")
'                LstPesq.SubItems(5) = Format(VarConv.Fields("Valor"), "##,##0.00")  'SeCampoNulo(Consulta, "BomPara")
'                Soma = Soma + VarConv!Valor
'                MebTotal.Text = Format(Soma, "     ##,##0.00")'''

'                VarConv.MoveNext
'            Wend
'        End If
        LblCard = TxtCartao
        VarItem = 0
        CodCard = TxtCartao

        TxtCartao.Enabled = False
'        If txtdadoslanca(0). = True Then
            TxtDadosLanca(0).SetFocus
'        End If
    Else
        MsgBox "Mesa não indentificada !", vbInformation
        TxtCartao = ""
        TxtCartao.SetFocus
    End If
End If
End Sub

Private Sub TxtCartao_LostFocus()

If TxtCartao = "" Then
        MsgBox "Mesa não indentificada !", vbInformation
        TxtCartao = ""
        LblCard = ""
        TxtCartao.Enabled = True
        TxtCartao.SetFocus
Else
    If TxtCartao <= 0 Or TxtCartao > 204 Then
            MsgBox "Mesa Inválida !", vbInformation
            TxtCartao = ""
            LblCard = ""
            TxtCartao.Enabled = True
            TxtCartao.SetFocus
    End If

End If

If IsNumeric(TxtCartao) Then TxtCartao_KeyPress 13
TxtCartao.BackColor = &H80000005
End Sub

Private Sub txtcomplemento_GotFocus()
txtcomplemento.BackColor = &HC0FFFF
End Sub

Private Sub txtcomplemento_LostFocus()
txtcomplemento.BackColor = &H80000005
End Sub

Private Sub txtdadoslanca_Change(Index As Integer)
Dim Criterio As String
If Index = 2 Then
    If TxtDadosLanca(2).Text <> "" Then
        If IsNumeric(TxtDadosLanca(2).Text) Then
             Criterio = "CodPrd = " & TxtDadosLanca(2).Text
             DatDados(0).Recordset.FindFirst Criterio
             If Not DatDados(0).Recordset.NoMatch Then
                   DbcProd.Text = DatDados(0).Recordset.Fields("Descrição_Produto")
             End If
        End If
    End If
End If
End Sub

Private Sub txtdadoslanca_GotFocus(Index As Integer)
TxtDadosLanca(Index).BackColor = &HC0FFFF
End Sub

Private Sub txtdadoslanca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Status As String
Dim sql As String
'If Index = 2 Then
'    If KeyCode = 113 Then
'        FrmPesquisa.Show vbModal
'        txtdadoslanca(2).SetFocus
'        txtdadoslanca_KeyPress 2, 13
'    End If
'End If

If KeyCode = 119 And CmbImprimi.Enabled = True Then
    CmbImprimi_Click
End If

If KeyCode = 118 And CmbLista.Enabled = True Then
     CmbLista_Click
End If

If KeyCode = 117 And CmbFecha.Enabled = True Then
     CmbFecha_Click
End If


If KeyCode = 120 Then
    CmbTotaliza_Click
End If

'If KeyCode = 121 Then
'    LimpaTela
'End If

If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub txtdadoslanca_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Criterio As String
Dim Status As String
Dim buf_ret As String * 100
Select Case Index
    
    Case 0
'        If CodCli <> 0 Then
            If KeyAscii = 13 And TxtDadosLanca(0) <> "" Then
                If Len(TxtDadosLanca(0)) > 8 Then
                    If IsNumeric(TxtDadosLanca(0)) Then
                        TxtTmp = "1"
                        EfetuaVenda (TxtDadosLanca(0))
                    Else
                        TxtDadosLanca(0) = ""
                        TxtDadosLanca(0).SetFocus
                    End If
                Else
                    If IsNumeric(TxtDadosLanca(0)) Then
                        If Len(TxtDadosLanca(0)) < 8 Then
                            TxtDadosLanca(2).SetFocus
                        Else
                            TxtDadosLanca(0) = ""
                            TxtDadosLanca(0).SetFocus
                        End If
                    Else
                        TxtDadosLanca(0) = ""
                        TxtDadosLanca(0).SetFocus
                    End If
                End If
            End If
    Case 2
        If KeyAscii = 27 Then
            TxtDadosLanca(2) = ""
            TxtDadosLanca(0) = ""
            TxtDadosLanca(0).SetFocus
        End If
        
        If KeyAscii = 13 And TxtDadosLanca(2) <> "" Then
            If IsNumeric(TxtDadosLanca(2)) And TxtDadosLanca(0) <> "" Then
                EfetuaVenda (TxtDadosLanca(2))
            Else
                TxtTmp = ""
                TxtDadosLanca(2) = ""
                TxtDadosLanca(2).SetFocus
            End If
        End If
        If KeyAscii = 13 And TxtDadosLanca(2) = "" And TxtDadosLanca(0) <> "" Then
                DbcProd.SetFocus
        End If
End Select

End Sub

Private Sub IncluirLinha(ValorVenda As String)
Dim quantidade As Currency
Dim Valor As Currency
Dim ConvertVr As String
Dim ConvertQtd As String
Dim Retorno As Integer
Dim TESTE As String
Dim TESTE2 As String
Dim Produto As String
Dim CodigoBarra As String
Dim Status As String
Dim sql As String
Dim Abasteci As Long
Dim Bico As Byte
Dim Posicao As Byte
Dim VarCodigo As Double

On Error GoTo erro
'Bico = 0


If Len(TxtDadosLanca(0)) > 8 Then
    quantidade = 1
Else
    quantidade = CDec(TxtDadosLanca(0))
End If

If Len(Trim(ValorVenda)) > 8 Then
    If Left(ValorVenda, 1) = 2 Then
            quantidade = CStr((CDec(Mid(ValorVenda, 8, 5)) / 100) / Registros.Fields("Venda"))
            MebUnit.Text = Registros.Fields("Venda") * quantidade
    End If
End If
'If TxtTmp.Text = "" Then
    Valor = IIf(IsNull(Format(Registros.Fields("Venda"), "##,##0.000")), 0, Format(Registros.Fields("Venda"), "##,##0.00"))
'Else
'    If ChkPzo Then
'        Valor = Format(MebUnit, "##,##0.000")
'    Else
'        Valor = Format(TxtTmp.Text, "##,##0.000")
'    End If
'    TxtTmp.Text = ""
'End If

'If TxtPessoa <> "" Then
    Pessoa = TxtPessoa
'End If

If Valor = 0 Then
    FrmValor.Show vbModal
    Valor = ValorUnit
End If
If TxtDadosLanca(0).Text <> "" Then
    If IsNull(Registros.Fields("BarBarra")) Then
        CodigoBarra = "0"
    Else
        CodigoBarra = Registros.Fields("BarBarra")
    End If
    
    Produto = LimpaPrd()
    ConvertQtd = Format(quantidade, "##,##0.00")
    ConvertVr = Format(Valor, "###,##0.000")
    If MebTotal <> "" Then
        Soma = MebTotal  ' Format(Soma, "     ##,##0.00") + (Format(Valor, "     ##,##0.000") * Format(quantidade, "     ##,##0.00"))
    Else
        Soma = 0
    End If
    Soma = Soma + (Valor * quantidade)
    MebTotal.Caption = Format(Soma, "     ##,##0.00")
    
    
    
    sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom, Qtd, Produto,  Valor ) Values ("
    sql = sql & CodCard & " ,"
    sql = sql & """" & LimpaCampo(Left(Pessoa, 15)) & """" & " , #"
    sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
    sql = sql & """" & CStr(Time) & """" & ", "
    sql = sql & IIf(IsNumeric(txtcomplemento.Text), txtcomplemento, 0) & " , "
    sql = sql & Numero(Format(quantidade, " ##,##0.00"), 2) & " , "
    sql = sql & Registros.Fields("CodPrd") & " , "
    sql = sql & Numero(Format(Valor, "##,##0.00") * Format(quantidade, "##,##0.00"), 2) & ")"
    FrmFrente.Banco.Execute sql
    
    Varvda.Requery
'     VarVda.Bookmark = VarVda.LastModified
    
    Varvda.MoveLast
    
    VarCodigo = Varvda("CodVda")
    
    Set LstPesq = ListVda.ListItems.Add(, , VarCodigo)
    
    LstPesq.SubItems(1) = Registros.Fields("CodPrd")   'SeCampoNulo(Consulta, "Agencia")
    LstPesq.SubItems(2) = Registros.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
    LstPesq.SubItems(3) = Numero(Format(quantidade, " ##,##0.00"), 2)  'SeCampoNulo(Consulta, "Conta")
    LstPesq.SubItems(4) = Numero(Format(Valor, "##,##0.00"), 2)  'SeCampoNulo(Consulta, "DataOper")
    LstPesq.SubItems(5) = Numero(Format(Valor, "##,##0.000") * Format(quantidade, "#0.000"), 2) 'SeCampoNulo(Consulta, "BomPara")
    If txtcomplemento <> "" Then
            LstPesq.SubItems(6) = txtcomplemento
    End If
    CmbImprimi.Enabled = False
    CmbLista.Enabled = False
    CmbFecha.Enabled = False

            
End If

erro:
If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If

End Sub

Private Sub EfetuaVenda(Valor As String)
Dim Terminação As String
Terminação = """"
If Len(Trim(Valor)) > 5 Then
    If Left(Valor, 1) = 2 Then
        Criterio = Registros.Fields("BarBarra").Name & " = "
        Criterio = Criterio & Terminação & Mid(Valor, 2, 6) & Terminação
    Else
        Criterio = Registros.Fields("BarBarra").Name & " = "
        Criterio = Criterio & Terminação & Valor & Terminação
    End If
Else
    Criterio = Registros.Fields(1).Name & " = "
    Criterio = Criterio & Valor
End If
Registros.FindFirst Criterio
If Registros.NoMatch Then
    MsgBox "Informação não Localizada", vbExclamation, App.Title
Else
    If Len(Trim(Valor)) > 5 Then
        TxtTmp = Registros!Venda
    End If
'    txtdadoslanca(1).Text = Registros.Fields("Descrição_Produto").Value
'    MebUnit.Text = Registros.Fields("venda").Value
    IncluirLinha (Valor)
    TxtDadosLanca(0) = ""
    TxtDadosLanca(2) = ""
    TxtDadosLanca(0).SetFocus
End If
End Sub

Function LimpaPrd() As String
Dim Produto As String
Dim PrdSujo As String
Dim x As Integer
Produto = ""
PrdSujo = Left(Registros.Fields("Descrição_Produto"), 38)
For x = 1 To Len(PrdSujo)
    If Mid$(PrdSujo, x, 1) >= Chr(32) And Mid$(PrdSujo, x, 1) <= Chr(122) Then
            Produto = Produto + Mid$(PrdSujo, x, 1)
    End If
Next
LimpaPrd = psCompString(3, Produto, 38)
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"

If KeyCode = 121 Then
    LimpaTela
End If

If KeyCode = 27 Then
    Unload Me
End If

End Sub


Private Sub ListaBicos_LostFocus()
ListaBicos.BackColor = &HFFFFFF ' Branco
MebValor.BackColor = &H80FFFF ' Amarelo
MebValor.SetFocus
End Sub

Private Sub ListaPreco_GotFocus()
ListaPreco.BackColor = &H80FFFF ' Amarelo
End Sub

Private Sub ListaPreco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If ListaPreco.ListIndex = 2 Then
    If VarPrecoLiberado Then
        MebUnitario.Enabled = True
        MebUnitario = Format(Registros.Fields("Venda"), "##,#0.000")
        MebUnitario.SetFocus
    Else
        MsgBox "Operação não aceita", vbInformation, App.Title
    End If
Else
    MebUnitario.Enabled = False
    MebUnitario = Format(ListaPreco.Text, "##,#0.000")
    MebValor = TxtDadosLanca(0) * ListaPreco.Text
    If Bico.RecordCount > 0 Then
        ListaBicos.SetFocus
    Else
        MebValor.SetFocus
    End If
End If
    
End If
End Sub

Private Sub ListaPreco_LostFocus()
ListaPreco.BackColor = &HFFFFFF ' Branco
End Sub

Private Sub MebUnitario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MebUnitario <> 0 Then
        MebValor = TxtDadosLanca(0) * MebUnitario
    End If
    If Bico.RecordCount > 0 Then
        ListaBicos.SetFocus
    Else
        MebValor.SetFocus
    End If
End If
End Sub

Private Sub MebValor_GotFocus()
MebValor.SelStart = 0
MebValor.SelLength = 12
End Sub

Private Sub MebUnitario_GotFocus()
MebValor.SelStart = 0
MebValor.SelLength = 12
End Sub

Private Sub MebValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MebValor <> "" Then
        If IsNumeric(MebValor) Then
            If MebValor <> 0 Then
                If MebUnitario <> 0 Then
                    TxtTmp.Text = MebUnitario
                    If Bico.RecordCount > 0 Then
                        LblPos.Caption = CInt(Mid$(ListaBicos.Text, 5, 2))
                        LblBico.Caption = CInt(Left(ListaBicos.Text, 2))
                    End If
                    TxtDadosLanca(0) = Round(Format(MebValor / MebUnitario, "##,##0.000"), 3)
                    MebUnit.Text = TxtTmp
                    Criterio = Registros.Fields(0).Name & " = "
                    Criterio = Criterio & PrdAnt
                    Registros.FindFirst Criterio
                    IncluirLinha (Valor)
                    TxtDadosLanca(0) = ""
                    TxtDadosLanca(2) = ""
                    TxtDadosLanca(0).SetFocus
                Else
                    MsgBox "Preço unitário não pode ser 0 (zero)", vbInformation, App.Title
                    MebUnitario.Enabled = True
                    MebUnitario.SetFocus
                End If
            Else
                MebValor.SetFocus
            End If
        Else
            MsgBox "Valor inválido", vbInformation, App.Title
            MebValor.SetFocus
        End If
    Else
        MsgBox "Valor inválido", vbInformation, App.Title
        MebValor.SetFocus
    End If
End If

End Sub

Private Sub MebValor_LostFocus()
MebValor.BackColor = &HFFFFFF ' Branco
FrmPreco.Visible = False
End Sub


Private Sub txtdadoslanca_LostFocus(Index As Integer)
TxtDadosLanca(Index).BackColor = &H80000005
End Sub

Private Sub TxtPessoa_GotFocus()
TxtPessoa.BackColor = &HC0FFFF
End Sub

Private Sub TxtPessoa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtDadosLanca(0).SetFocus
End If
End Sub

Public Sub PesquisaGeral(Tabela As String, sql As String, Campo As String)
'    FramePesquisaG.Visible = True

    DatPesquisaG.DatabaseName = Tabela
    DatPesquisaG.RecordSource = sql
'    Formulário.DatPesquisaG.DataField = "Cidade"
'    Formulário.DatPesquisaG.ListField = Campo
    DatPesquisaG.Refresh
'    DblPesquisaG.Refresh
'    Formulário.DblPesquisaG.SetFocus

End Sub

Private Sub TxtPessoa_LostFocus()
TxtPessoa.BackColor = &H80000005

If IsNumeric(TxtPessoa.BoundText) Then
    txtcomplemento = TxtPessoa.BoundText
End If
End Sub
