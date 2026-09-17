VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FrmPedidos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14790
   Icon            =   "FrmPedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   14790
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   14790
      TabIndex        =   130
      Top             =   0
      Width           =   14820
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   360
         TabIndex        =   138
         Top             =   45
         Width           =   2085
      End
      Begin VB.Label LblAtraso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12285
         TabIndex        =   136
         Top             =   45
         Width           =   1950
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8430
      Left            =   0
      TabIndex        =   26
      Top             =   630
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14870
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmPedidos.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdDescAtacado"
      Tab(0).Control(1)=   "CmdPedMeiaFolha"
      Tab(0).Control(2)=   "BtPedPrazo"
      Tab(0).Control(3)=   "DatDados(10)"
      Tab(0).Control(4)=   "BtXml"
      Tab(0).Control(5)=   "DatDados(5)"
      Tab(0).Control(6)=   "BtAtualiza"
      Tab(0).Control(7)=   "BtContEntrega"
      Tab(0).Control(8)=   "ImprNota"
      Tab(0).Control(9)=   "BtSelo"
      Tab(0).Control(10)=   "ImprimeDupi"
      Tab(0).Control(11)=   "BtConsolida"
      Tab(0).Control(12)=   "BtDesconto"
      Tab(0).Control(13)=   "BtCarne"
      Tab(0).Control(14)=   "BtCopiaPed"
      Tab(0).Control(15)=   "BtCancelaNt"
      Tab(0).Control(16)=   "ChkErro"
      Tab(0).Control(17)=   "TxtPesqNota"
      Tab(0).Control(18)=   "TxtCodPed"
      Tab(0).Control(19)=   "DatDados(4)"
      Tab(0).Control(20)=   "Frame2"
      Tab(0).Control(21)=   "DbcDados(3)"
      Tab(0).Control(22)=   "GridPed"
      Tab(0).Control(23)=   "DtData"
      Tab(0).Control(24)=   "DbcDados(4)"
      Tab(0).Control(25)=   "LblNomePedido"
      Tab(0).Control(26)=   "Label13"
      Tab(0).Control(27)=   "Label21"
      Tab(0).Control(28)=   "Label12"
      Tab(0).Control(29)=   "Label20"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Pedido"
      TabPicture(1)   =   "FrmPedidos.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "SSTabItens"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cMBoRDEM"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Tabela de Preços"
      TabPicture(2)   =   "FrmPedidos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid3"
      Tab(2).Control(1)=   "DBGTabela"
      Tab(2).Control(2)=   "DatDados(9)"
      Tab(2).Control(3)=   "DatDados(8)"
      Tab(2).Control(4)=   "Frame8"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton CmdDescAtacado 
         Caption         =   "Desc.Atacado"
         Height          =   375
         Left            =   -64875
         TabIndex        =   144
         Top             =   7425
         Width           =   1365
      End
      Begin VB.CommandButton CmdPedMeiaFolha 
         Caption         =   "Pedido Meia Folha"
         Height          =   375
         Left            =   -63435
         TabIndex        =   143
         Top             =   7425
         Width           =   1635
      End
      Begin VB.CommandButton BtPedPrazo 
         Caption         =   "Inconsistência"
         Height          =   375
         Left            =   -61725
         TabIndex        =   142
         Top             =   7425
         Width           =   1230
      End
      Begin VB.CommandButton cMBoRDEM 
         Caption         =   "..."
         Height          =   330
         Left            =   8010
         TabIndex        =   141
         Top             =   7110
         Width           =   465
      End
      Begin VB.Data DatDados 
         Caption         =   "convenio"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   10
         Left            =   -67935
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   855
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Frame Frame8 
         Caption         =   "Pesquise"
         Height          =   960
         Left            =   -74820
         TabIndex        =   132
         Top             =   450
         Width           =   6720
         Begin VB.TextBox TxtPesq 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   180
            TabIndex        =   133
            Top             =   360
            Width           =   6090
         End
      End
      Begin VB.CommandButton BtXml 
         Caption         =   "Importa &XML"
         Height          =   330
         Left            =   -71880
         TabIndex        =   131
         Top             =   7920
         Width           =   1515
      End
      Begin VB.Data DatDados 
         Caption         =   "Produtos"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   8
         Left            =   -70545
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3150
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Data DatDados 
         Caption         =   "Loja"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   9
         Left            =   -67080
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Data DatDados 
         Caption         =   "Pedido"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   5
         Left            =   -67890
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4095
         Visible         =   0   'False
         Width           =   2085
      End
      Begin MSDBGrid.DBGrid DBGTabela 
         Bindings        =   "FrmPedidos.frx":0060
         Height          =   6435
         Left            =   -74820
         OleObjectBlob   =   "FrmPedidos.frx":007A
         TabIndex        =   32
         Top             =   1575
         Width           =   12075
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "FrmPedidos.frx":110B
         Height          =   3690
         Left            =   -62625
         OleObjectBlob   =   "FrmPedidos.frx":1125
         TabIndex        =   33
         Top             =   1575
         Width           =   2205
      End
      Begin VB.CommandButton BtAtualiza 
         Caption         =   "&Atualiza"
         Height          =   315
         Left            =   -61185
         TabIndex        =   125
         Top             =   450
         Width           =   810
      End
      Begin VB.CommandButton BtContEntrega 
         Caption         =   "&Controle de Entrega"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -67500
         TabIndex        =   124
         Top             =   7920
         Width           =   1695
      End
      Begin VB.CommandButton ImprNota 
         Caption         =   "&Nota Fiscal"
         Height          =   330
         Left            =   -70230
         TabIndex        =   123
         Top             =   7920
         Width           =   1140
      End
      Begin VB.CommandButton BtSelo 
         Caption         =   "&Ident. de Volume"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -69030
         TabIndex        =   122
         Top             =   7920
         Width           =   1425
      End
      Begin VB.CommandButton ImprimeDupi 
         Caption         =   "&Promissória"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74880
         TabIndex        =   121
         Top             =   7920
         Width           =   1200
      End
      Begin VB.CommandButton BtConsolida 
         Caption         =   "&Consolida Pedido"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73605
         TabIndex        =   120
         Top             =   7920
         Width           =   1515
      End
      Begin VB.CommandButton BtDesconto 
         Caption         =   "&Desconto"
         Height          =   330
         Left            =   -65685
         TabIndex        =   119
         Top             =   7920
         Width           =   960
      End
      Begin VB.CommandButton BtCarne 
         Caption         =   "&Carnê"
         Height          =   330
         Left            =   -64620
         TabIndex        =   118
         Top             =   7920
         Width           =   1110
      End
      Begin VB.CommandButton BtCopiaPed 
         Caption         =   "Co&pia Pedido"
         Height          =   330
         Left            =   -63360
         TabIndex        =   117
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton BtCancelaNt 
         Caption         =   "Cance&la Nro Nota"
         Height          =   330
         Left            =   -61905
         TabIndex        =   116
         Top             =   7920
         Width           =   1470
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   135
         TabIndex        =   104
         Top             =   7065
         Width           =   4170
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   75
            TabIndex        =   109
            Top             =   975
            Width           =   3750
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   108
            Top             =   135
            Width           =   3750
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   107
            Top             =   345
            Width           =   3750
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   106
            Top             =   555
            Width           =   3750
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   105
            Top             =   765
            Width           =   3750
         End
      End
      Begin VB.CheckBox ChkErro 
         Caption         =   "Erro"
         Height          =   330
         Left            =   -62535
         TabIndex        =   94
         Top             =   855
         Width           =   600
      End
      Begin VB.Frame Frame7 
         Enabled         =   0   'False
         Height          =   675
         Left            =   7410
         TabIndex        =   79
         Top             =   360
         Width           =   5325
         Begin MSMask.MaskEdBox MebHora 
            Height          =   270
            Left            =   1530
            TabIndex        =   80
            Top             =   330
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebData 
            Height          =   270
            Left            =   150
            TabIndex        =   81
            Top             =   330
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebFat 
            Height          =   270
            Left            =   2730
            TabIndex        =   83
            Top             =   330
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebNota 
            Height          =   270
            Left            =   4260
            TabIndex        =   84
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Fatura"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2700
            TabIndex        =   87
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   90
            TabIndex        =   86
            Top             =   120
            Width           =   570
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   2640
            X2              =   2640
            Y1              =   120
            Y2              =   630
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   1440
            X2              =   1440
            Y1              =   90
            Y2              =   630
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   4110
            X2              =   4110
            Y1              =   90
            Y2              =   660
         End
      End
      Begin TabDlg.SSTab SSTabItens 
         Height          =   5970
         Left            =   90
         TabIndex        =   52
         Top             =   1080
         Width           =   14550
         _ExtentX        =   25665
         _ExtentY        =   10530
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Cadastro"
         TabPicture(0)   =   "FrmPedidos.frx":1B10
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ObsEnt"
         Tab(0).Control(1)=   "DatDados(6)"
         Tab(0).Control(2)=   "Frame6"
         Tab(0).Control(3)=   "TxtObs"
         Tab(0).Control(4)=   "Frame(0)"
         Tab(0).Control(5)=   "GridParc"
         Tab(0).Control(6)=   "Label28"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Produtos"
         TabPicture(1)   =   "FrmPedidos.frx":1B2C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Grid"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "DatDados(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "DatDados(7)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.TextBox ObsEnt 
            Appearance      =   0  'Flat
            Height          =   1305
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   129
            Top             =   2760
            Width           =   8415
         End
         Begin VB.Data DatDados 
            Caption         =   "Parcelas"
            Connect         =   "Access"
            DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   360
            Index           =   6
            Left            =   -65055
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2475
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame Frame6 
            Caption         =   "Custo Pedido"
            Enabled         =   0   'False
            Height          =   600
            Left            =   -66315
            TabIndex        =   111
            Top             =   5040
            Width           =   5730
            Begin MSMask.MaskEdBox MebMargemPed 
               Height          =   270
               Left            =   2655
               TabIndex        =   112
               Top             =   270
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   476
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   -2147483633
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebCustoPed 
               Height          =   270
               Left            =   450
               TabIndex        =   113
               Top             =   270
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   476
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   -2147483633
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   1935
               TabIndex        =   128
               Top             =   270
               Width           =   495
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "R$"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   105
               TabIndex        =   115
               Top             =   270
               Width           =   180
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   195
               Left            =   3420
               TabIndex        =   114
               Top             =   270
               Width           =   120
            End
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            Height          =   1350
            Left            =   -74865
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Top             =   4410
            Width           =   8415
         End
         Begin VB.Frame Frame 
            Height          =   2325
            Index           =   0
            Left            =   -74910
            TabIndex        =   67
            Top             =   315
            Width           =   8430
            Begin VB.CommandButton BtPesquisa 
               Appearance      =   0  'Flat
               DisabledPicture =   "FrmPedidos.frx":1B48
               Height          =   345
               Left            =   5535
               Picture         =   "FrmPedidos.frx":207A
               Style           =   1  'Graphical
               TabIndex        =   140
               ToolTipText     =   "Pesquisar"
               Top             =   1080
               Width           =   465
            End
            Begin VB.TextBox TxtCfop 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5865
               TabIndex        =   5
               Top             =   390
               Width           =   1230
            End
            Begin VB.TextBox txtPedidoRef 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4485
               TabIndex        =   4
               Top             =   390
               Width           =   1290
            End
            Begin VB.TextBox TxtDesc 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2760
               TabIndex        =   3
               Top             =   390
               Width           =   1635
            End
            Begin VB.Data DatDados 
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Index           =   3
               Left            =   2070
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "Select CodEntidade,Nome, Tipo, Comissao From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
               Top             =   1710
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.Data DatDados 
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Index           =   0
               Left            =   1935
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   ""
               Top             =   1080
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox ChkFunRural 
               Caption         =   "FunRural"
               Height          =   240
               Left            =   7380
               TabIndex        =   7
               Top             =   1125
               Width           =   960
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   3
               Left            =   6165
               TabIndex        =   9
               Top             =   1710
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSDBCtls.DBCombo DbcDados 
               Bindings        =   "FrmPedidos.frx":25AC
               Height          =   315
               Index           =   0
               Left            =   135
               TabIndex        =   6
               Top             =   1080
               Width           =   5280
               _ExtentX        =   9313
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   1
               Left            =   150
               TabIndex        =   1
               Top             =   390
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   0
               Left            =   1395
               TabIndex        =   2
               Top             =   390
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSDBCtls.DBCombo DbcDados 
               Bindings        =   "FrmPedidos.frx":25C6
               DataField       =   "Nome"
               Height          =   315
               Index           =   2
               Left            =   135
               TabIndex        =   8
               Top             =   1710
               Width           =   5280
               _ExtentX        =   9313
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ListField       =   "Nome"
               BoundColumn     =   "CodEntidade"
               Text            =   ""
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "&Entidade"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   78
               Top             =   855
               Width           =   630
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C.F.Op"
               Height          =   195
               Left            =   5910
               TabIndex        =   76
               Top             =   180
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Emissão"
               Height          =   195
               Left            =   150
               TabIndex        =   75
               Top             =   195
               Width           =   585
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ped.Ref."
               Height          =   195
               Left            =   4485
               TabIndex        =   74
               Top             =   180
               Width           =   585
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entrega"
               Height          =   195
               Left            =   1395
               TabIndex        =   73
               Top             =   180
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto (10+5+3) %"
               Height          =   195
               Left            =   2760
               TabIndex        =   72
               Top             =   180
               Width           =   1530
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Vendedor/Comprador"
               Height          =   195
               Left            =   135
               TabIndex        =   71
               Top             =   1470
               Width           =   1530
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Comissão (%)"
               Height          =   195
               Left            =   6165
               TabIndex        =   70
               Top             =   1485
               Width           =   930
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   195
               Index           =   3
               Left            =   6165
               TabIndex        =   69
               Top             =   855
               Width           =   540
            End
            Begin VB.Label LblCodigo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   6165
               TabIndex        =   68
               Top             =   1080
               Width           =   1080
            End
         End
         Begin VB.Data DatDados 
            Caption         =   "Loja"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   7
            Left            =   4500
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3240
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Data DatDados 
            Caption         =   "Itens"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   2
            Left            =   2340
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Frame Frame 
            Height          =   1350
            Index           =   1
            Left            =   60
            TabIndex        =   56
            Top             =   330
            Width           =   14310
            Begin VB.TextBox TxtRefItens 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6795
               MaxLength       =   40
               TabIndex        =   12
               Top             =   375
               Width           =   3480
            End
            Begin VB.Data DatDados 
               Caption         =   "Produtos"
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   360
               Index           =   1
               Left            =   2565
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "Select CodPrd, Descrição_Produto, CodBar, Complemento, Estoque, Ativo  From PRD Where Ativo = True  Order By Descrição_Produto"
               Top             =   360
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.ComboBox DbcLoja 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "System"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   9585
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   930
               Width           =   2595
            End
            Begin VB.ComboBox CmbAcab 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   10485
               TabIndex        =   13
               Top             =   375
               Width           =   3240
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   5
               Left            =   6450
               TabIndex        =   18
               Top             =   930
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.0000;(#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   2
               Left            =   4980
               TabIndex        =   17
               Top             =   930
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   4
               Left            =   90
               TabIndex        =   10
               Top             =   375
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   6
               Left            =   12225
               TabIndex        =   21
               Top             =   930
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.000;(#,##0.000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebAcrescimo 
               Height          =   315
               Left            =   3285
               TabIndex        =   16
               Top             =   930
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   8
               Left            =   7995
               TabIndex        =   19
               Top             =   930
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "#,##0.0000;(#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebVlr 
               Height          =   315
               Left            =   90
               TabIndex        =   14
               Top             =   930
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.000;(#,##0.000)"
               PromptChar      =   "_"
            End
            Begin MSDBCtls.DBCombo DbcDados 
               Bindings        =   "FrmPedidos.frx":25E0
               Height          =   315
               Index           =   1
               Left            =   1305
               TabIndex        =   11
               Top             =   360
               Width           =   5235
               _ExtentX        =   9234
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ListField       =   "Descrição_Produto"
               BoundColumn     =   "CodPrd"
               Text            =   ""
            End
            Begin MSMask.MaskEdBox MebComissao 
               Height          =   315
               Left            =   1665
               TabIndex        =   15
               Top             =   930
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comissão"
               Height          =   195
               Left            =   1650
               TabIndex        =   139
               Top             =   720
               Width           =   675
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Referência"
               Height          =   195
               Left            =   6750
               TabIndex        =   110
               Top             =   180
               Width           =   780
            End
            Begin VB.Label QtdCorpo 
               Height          =   285
               Left            =   13860
               TabIndex        =   85
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Produto"
               Height          =   195
               Left            =   1335
               TabIndex        =   66
               Top             =   165
               Width           =   555
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   6450
               TabIndex        =   65
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc(%)"
               Height          =   195
               Left            =   4995
               TabIndex        =   64
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Financeiro"
               Height          =   195
               Left            =   12210
               TabIndex        =   63
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Loja/Depósito/Grade"
               Height          =   195
               Left            =   9585
               TabIndex        =   62
               Top             =   750
               Width           =   1515
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Complemento"
               Height          =   195
               Left            =   10485
               TabIndex        =   61
               Top             =   180
               Width           =   960
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acréscimo (10+5+3)%"
               Height          =   195
               Left            =   3315
               TabIndex        =   60
               Top             =   720
               Width           =   1530
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   7950
               TabIndex        =   59
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Produto"
               Height          =   195
               Left            =   90
               TabIndex        =   58
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Quantidade"
               Height          =   195
               Left            =   90
               TabIndex        =   57
               Top             =   150
               Width           =   825
            End
         End
         Begin MSDBGrid.DBGrid Grid 
            Bindings        =   "FrmPedidos.frx":25FA
            Height          =   4035
            Left            =   45
            OleObjectBlob   =   "FrmPedidos.frx":2614
            TabIndex        =   55
            Top             =   1800
            Width           =   14355
         End
         Begin MSDBGrid.DBGrid GridParc 
            Bindings        =   "FrmPedidos.frx":40B8
            Height          =   4560
            Left            =   -66315
            OleObjectBlob   =   "FrmPedidos.frx":40D2
            TabIndex        =   127
            Top             =   405
            Width           =   5700
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observação"
            Height          =   195
            Left            =   -74865
            TabIndex        =   82
            Top             =   4170
            Width           =   870
         End
      End
      Begin VB.TextBox TxtPesqNota 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65325
         TabIndex        =   50
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox TxtCodPed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -64020
         TabIndex        =   48
         Top             =   855
         Width           =   1185
      End
      Begin VB.Frame Frame5 
         Height          =   1320
         Left            =   4455
         TabIndex        =   37
         Top             =   6990
         Width           =   3465
         Begin VB.CommandButton ImprNota1 
            Height          =   555
            Left            =   2760
            Picture         =   "FrmPedidos.frx":516E
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Imprimi Nota"
            Top             =   720
            Width           =   600
         End
         Begin VB.CommandButton BtOk 
            Enabled         =   0   'False
            Height          =   555
            Left            =   720
            Picture         =   "FrmPedidos.frx":5B18
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Grava Pedido"
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton BtCancela 
            Enabled         =   0   'False
            Height          =   555
            Left            =   2091
            Picture         =   "FrmPedidos.frx":6182
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Apaga Pedido Atual"
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton BtImprimir 
            Enabled         =   0   'False
            Height          =   555
            Left            =   757
            Picture         =   "FrmPedidos.frx":65C4
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Imprimi Pedido Matricial"
            Top             =   720
            Width           =   600
         End
         Begin VB.CommandButton BtAlterar 
            Enabled         =   0   'False
            Height          =   555
            Left            =   1440
            Picture         =   "FrmPedidos.frx":6C2E
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Altera Pedido"
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton BtParcelas 
            Enabled         =   0   'False
            Height          =   555
            Left            =   2760
            Picture         =   "FrmPedidos.frx":7070
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Calcula Parcelas"
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton BtTransportadora 
            Enabled         =   0   'False
            Height          =   555
            Left            =   1424
            Picture         =   "FrmPedidos.frx":74B2
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Pedido de Entrega"
            Top             =   720
            Width           =   600
         End
         Begin VB.CommandButton BtPedProm 
            Enabled         =   0   'False
            Height          =   555
            Left            =   2070
            Picture         =   "FrmPedidos.frx":77BC
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Imprimi Pedido Laser/Jato "
            Top             =   720
            Width           =   600
         End
         Begin VB.CommandButton BtNovo 
            Height          =   555
            Left            =   90
            Picture         =   "FrmPedidos.frx":7AC6
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Pedido Novo"
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton BtCupom 
            Enabled         =   0   'False
            Height          =   555
            Left            =   90
            Picture         =   "FrmPedidos.frx":8130
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Imprimi Pedido Tipo Cupom"
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1320
         Left            =   8550
         TabIndex        =   34
         Top             =   6990
         Width           =   6105
         Begin VB.TextBox TxtDados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   6
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   101
            Tag             =   "N"
            Top             =   945
            Width           =   1335
         End
         Begin VB.TextBox TxtDados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   5
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   99
            Tag             =   "N"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtDados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   4
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   98
            Tag             =   "N"
            Top             =   945
            Width           =   1200
         End
         Begin VB.TextBox TxtDados 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   3
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   97
            Tag             =   "N"
            Top             =   360
            Width           =   1200
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   405
            Left            =   4560
            ScaleHeight     =   405
            ScaleWidth      =   1455
            TabIndex        =   88
            Top             =   240
            Width           =   1455
            Begin VB.TextBox TxtDados 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   1
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   89
               Tag             =   "N"
               Top             =   30
               Width           =   1335
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   405
            Left            =   4560
            ScaleHeight     =   405
            ScaleWidth      =   1455
            TabIndex        =   53
            Top             =   810
            Width           =   1455
            Begin VB.TextBox TxtDados 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   2
               Left            =   75
               Locked          =   -1  'True
               TabIndex        =   54
               Tag             =   "N"
               Top             =   30
               Width           =   1335
            End
         End
         Begin MSMask.MaskEdBox MebTotDesc 
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   780
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDesc 
            Height          =   315
            Left            =   45
            TabIndex        =   46
            Top             =   390
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   945
            TabIndex        =   103
            Top             =   450
            Width           =   120
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Substituição"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1260
            TabIndex        =   102
            Top             =   765
            Width           =   885
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Base C/ Sub."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1260
            TabIndex        =   100
            Top             =   180
            Width           =   945
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00C0C0C0&
            X1              =   1170
            X2              =   1170
            Y1              =   135
            Y2              =   1275
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Ipi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2745
            TabIndex        =   96
            Top             =   735
            Width           =   150
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2745
            TabIndex        =   95
            Top             =   180
            Width           =   375
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00C0C0C0&
            X1              =   4050
            X2              =   4050
            Y1              =   135
            Y2              =   1275
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   60
            TabIndex        =   92
            Top             =   150
            Width           =   690
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   91
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Sub"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   90
            Top             =   150
            Width           =   285
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   36
            Top             =   750
            Width           =   345
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C0C0C0&
            X1              =   1200
            X2              =   6075
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C0C0C0&
            X1              =   2640
            X2              =   2640
            Y1              =   150
            Y2              =   1290
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Entidade"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   4
         Left            =   -73065
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   855
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Frame Frame2 
         Height          =   750
         Left            =   -74865
         TabIndex        =   30
         Top             =   450
         Width           =   1170
         Begin VB.OptionButton Op1 
            Caption         =   "Entrada"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   24
            Top             =   450
            Width           =   900
         End
         Begin VB.OptionButton Op1 
            Caption         =   "Saída"
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   23
            Top             =   180
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pedido"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   13185
         TabIndex        =   28
         Top             =   360
         Width           =   1440
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   29
            Tag             =   "N"
            Top             =   210
            Width           =   1245
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Natureza da Operação"
         Height          =   630
         Index           =   3
         Left            =   90
         TabIndex        =   27
         Top             =   360
         Width           =   6015
         Begin VB.ComboBox CmbDados 
            Appearance      =   0  'Flat
            DataField       =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            ItemData        =   "FrmPedidos.frx":843A
            Left            =   105
            List            =   "FrmPedidos.frx":8471
            TabIndex        =   0
            Top             =   240
            Width           =   5820
         End
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FrmPedidos.frx":85EE
         Height          =   315
         Index           =   3
         Left            =   -73650
         TabIndex        =   31
         Top             =   855
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin MSDBGrid.DBGrid GridPed 
         Bindings        =   "FrmPedidos.frx":8608
         Height          =   5955
         Left            =   -74925
         OleObjectBlob   =   "FrmPedidos.frx":8622
         TabIndex        =   25
         Top             =   1350
         Width           =   14505
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   315
         Left            =   -61770
         TabIndex        =   126
         Top             =   855
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   43778049
         CurrentDate     =   37501
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FrmPedidos.frx":A27D
         Height          =   315
         Index           =   4
         Left            =   -68430
         TabIndex        =   134
         Top             =   855
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome_Conv"
         BoundColumn     =   "CodConv"
         Text            =   ""
      End
      Begin VB.Label LblNomePedido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74865
         TabIndex        =   137
         Top             =   7470
         Width           =   9735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "&Convênio"
         Height          =   195
         Left            =   -68430
         TabIndex        =   135
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nota"
         Height          =   195
         Left            =   -65325
         TabIndex        =   51
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pedido"
         Height          =   195
         Left            =   -64020
         TabIndex        =   49
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "&Entidade"
         Height          =   195
         Left            =   -73650
         TabIndex        =   35
         Top             =   630
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 0 - Entrada de Produtos(E)
' 1 - Venda Vista(S)
' 2 - Venda Parcelada(S)
' 3 - Ordem de Serviço(S)
' 4 - Orçamento (S)
' 5 - Consignação (S)
' 6 - Doação/Bonificação(S)
' 7 - Consórcio (S)
' 8 - Transferência (Remessa)(S)
' 9 - Transferência (Retorno)(E)
'10 - Devolução de Venda(E)
'11 - Devolução de Compra(S)
'12 - Pedido Compra(E)
'13 - Pedido Entrega (S)



Public Tabela As Recordset
Dim Itens As Recordset
Dim Parcelas As Recordset
Public pedidos As Recordset
Dim soma As Recordset
Public Entidade As Recordset
Dim Produtos As Recordset
Public VarCusto As Currency
Dim Acabamento As Recordset
Public Senha As Boolean
Public VarLoja As String
Public PcoPzo As Byte
Public Indices As Recordset
Public ValIndices As Currency
Public VarBarra As Recordset
Public Multa As Boolean
Public VarLista1 As Recordset
Public VarLibreaPreco As Boolean

Dim Desc(1 To 10) As Double

Public VarObs As String
Public VarOperacao As String
Public ContComplemento As Boolean
Public VarTotalPrd As Currency
Public VarContPed As Double
Public VarGrade As Boolean
Public Dolar As Currency
Public RetornoPesq As String

Private Sub BtAgrupa_Click()
If DatDados(5).Recordset.RecordCount > 0 Then
    If TxtCodPed = "" Then
        If MsgBox("Deseja agrupar itens da nota ?", vbYesNo, App.Title) = vbYes Then
            
        End If
'        FormEntregaFutura.TxtPedido = GridPed.Columns(2)
'        FormEntregaFutura.Show
'        Exit Sub
    End If
End If

End Sub

Private Sub BtAlterar_Click()
Dim Baixa As Boolean
Dim VarCont As Boolean
Dim SQL As String

VarCont = True
Senha = False
If Tabela!NroNt <> 0 Then
     MsgBox "Atenção este pedido tem nota fiscal, a alteração pode afetar a contabilidade", vbExclamation, App.Title
End If

If CmbDados(3).ListIndex = 1 Then
    SQL = " DELETE REC.*, REC.CodDupli FROM REC WHERE REC.CodDupli =  " & TxtDados(0)
    Banco.Execute SQL
End If

If CmbDados(3).ListIndex <> 5 Then
    FormSenha2.Show vbModal
Else
    Senha = True
End If
VarOperacao = Right(CmbDados(3).Text, 1)
If Senha = True Then


    For Each Componentes In Frame
        Componentes.Enabled = True
        Frame3.Enabled = True
    Next
'   CmbDados(3).Enabled = False
    SSTab1.TabEnabled(0) = False
    BtNovo.Enabled = False
    BtAlterar.Enabled = False
    BtCancela.Enabled = True
    BtOK.Enabled = True
    BtCupom.Enabled = False
    BtImprimir.Enabled = False
    ImprNota1.Enabled = False
    BtPedProm.Enabled = False
    BtTransportadora.Enabled = False
    BtParcelas.Enabled = True
'    BtObs.Enabled = True
    MebDados(4).SetFocus
    Open Caminho & "\Altera.txt" For Append As #1
    Print #1, "Pedido; " & VarCodPed & "; " & DbcDados(0).Text & "; " & Date & "; " & Time & "; " & Maquina & "; " & CmbDados(3).Text & "; " & Alinhar(TxtDados(2), 10, 2)
    Close #1
    Select Case CmbDados(3).ListIndex
       Case 0
            Baixa = False
        Case 1
            Baixa = True
        Case 2
            Baixa = True
        Case 3
            Baixa = True
        Case 6
            Baixa = True
        Case 8
            Baixa = True
        Case 9
            Baixa = False
        Case 10
            Baixa = False
        Case 11
            Baixa = True
        Case 13
            Baixa = True
        Case 14
            Baixa = True
        Case 15
            Baixa = False
        Case 16
            Baixa = False
    End Select
    
    Select Case CmbDados(3).ListIndex
            Case 3
                VarCont = False
            Case 4
                VarCont = False
            Case 5
                VarCont = False
            Case 7
                VarCont = False
            Case 12
                VarCont = False
            Case 13
                VarCont = False
            Case 14
'                VarCont = False
                
    End Select
    
    If VarCont = True Then
        If DatDados(2).Recordset.RecordCount > 0 Then
            If DatDados(2).Recordset.Eof = True Then
                DatDados(2).Recordset.MovePrevious
            End If
        End If
        
        If DatDados(2).Recordset.RecordCount > 0 Then
'            If DatDados(2).Recordset("Ped.Entregue") = 0 Then
'                DatDados(2).Recordset.Edit
'                DatDados(2).Recordset("Ped.Entregue") = 1
'                DatDados(2).Recordset.Update
            If Dir(App.Path & "\" & Trim(TxtDados(0)) & ".txt") <> "" Then
                Kill App.Path & "\" & Trim(TxtDados(0)) & ".txt"
            End If

            DatDados(2).Recordset.MoveFirst
            While Not DatDados(2).Recordset.Eof
                If DatDados(2).Recordset("Itp.Entregue") = True Then
                    If Not IsNull(Trim(DatDados(2).Recordset.Fields("Loja"))) Then
                        Open App.Path & "\" & TxtDados(0) & ".txt" For Append As #1
                        Print #1, psCompString(2, Numero2(DatDados(2).Recordset.Fields("Qtd")), 10) & psCompString(1, DatDados(2).Recordset.Fields("Produto"), 8) & psCompString(3, Trim(DatDados(2).Recordset.Fields("Loja")), 20) & IIf(Baixa, 0, 1)
                        Close #1
                    End If
                End If
                DatDados(2).Recordset.MoveNext
            Wend
'                BaixaEstoque TxtDados(0), Not Baixa, CmbDados(3).ListIndex
'             End If
         End If
'        DatDados(5).Recordset.Edit
'        DatDados(5).Recordset("Entregue") = 0
'        DatDados(5).Recordset.Update
    End If
If Baixa = True Then
    If Admin = False Then
        Txtcfop.Enabled = False
    Else
        Txtcfop.Enabled = True
    End If
Else
    Txtcfop.Enabled = True
End If
    
End If
End Sub

Private Sub BtAtualiza_Click()
If TxtPesqNota.Text = "" Then
    If Op1(0).value = True Then
        DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.DataEmiss = #" & Format(DtData, "mm/dd/yyyy") & "# and Ped.Operacao >= 1 ORDER BY PED.CodPed"
    Else
        DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.DataEmiss = #" & Format(DtData, "mm/dd/yyyy") & "# and Ped.Operacao = 0  ORDER BY PED.CodPed"
    End If
Else
    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat,  Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.NroNt  = " & TxtPesqNota & " ORDER BY PED.CodPed"
    TxtPesqNota = ""
End If
    DatDados(5).Refresh
    If DatDados(5).Recordset.RecordCount <> 0 Then
'       DatDados(5).Recordset.MoveFirst
       DatDados(5).Recordset.MoveLast
       VarContPed = DatDados(5).Recordset.RecordCount
'       GridPed.Refresh
'       DatDados(5).Refresh
    End If

End Sub

Private Sub BtCancela_Click()
Dim VARPASS As Recordset
    Set VARPASS = Banco.OpenRecordset("SELECT * FROM PWD WHERE NOME = 'EXCPEDIDO'")
    VARPASS.Requery
    If VARPASS.RecordCount = 1 Then
        Senha = False
        FormSenha2.Tela = "a"
        FormSenha2.Show vbModal
        If Senha = True Then
            ApagaPedido
            VarContPed = VarContPed - 1
            LIMPACampo
        End If
    Else
            ApagaPedido
            VarContPed = VarContPed - 1
            LIMPACampo
    End If
End Sub

Private Sub BtCancelaNt_Click()
If Admin Then
    If MsgBox("Confirma o cancelamento da nota => " & GridPed.Columns(7) & " ?", vbYesNo, App.Title) = vbYes Then
        DatDados(5).Recordset.Edit
        DatDados(5).Recordset("NroNt") = 0
        DatDados(5).Recordset("Cfo") = ""
        'DatDados(5).Recordset("DtFat") = ""
        DatDados(5).Recordset.Update
    End If
End If
End Sub

Private Sub BtCarne_Click()
If Op1(0).value = True Then

    If GridPed.Columns(3) = "Venda Prazo" Then 'Or GridPed.Columns(3) = 3 Or GridPed.Columns(3) = 7 Then
        FrmParcelas.CodDoc = GridPed.Columns(2)
        FrmParcelas.CodEnt = GridPed.Columns(0)
        FrmParcelas.MebVlrPed = Format(GridPed.Columns(8), "#,##0.00")
        FrmParcelas.TxtNome = psCompString(1, GridPed.Columns(0), 6) & "-" & DatDados(4).Recordset("Nome")
        FrmParcelas.OK.Enabled = False
        FrmParcelas.GridParc.Enabled = False
        FrmParcelas.Show vbModal
    Else
        MsgBox "Natureza da Operação Inválida", vbInformation, App.Title
    End If

End If

End Sub

Private Sub BtConsolida_Click()

FormConsolida.VarOper = DatDados(5).Recordset("Operacao")
FormConsolida.Show vbModal

End Sub

Private Sub BtContEntrega_Click()
If DatDados(5).Recordset.RecordCount > 0 Then
    If TxtCodPed = "" Then
       FormEntregaFutura.TxtPedido = GridPed.Columns(2)
       FormEntregaFutura.Show
       Exit Sub
    End If
End If
End Sub

Private Sub BtCopiaPed_Click()
Dim SQL As String
Dim VarCodPed As Double

If GridPed.Row = -1 Then
    MsgBox "Você deve selecionar algum pedido !!!", 16
Else
    If MsgBox("Deseja copiar o pedido numero " & GridPed.Columns(2), vbYesNo, App.Title) = vbYes Then
        SQL = " INSERT INTO PED  (Operacao, Entidade, DataEmiss, Total, Desconto, Vendedor, Entregue, Transportadora, Comissao, IcmSubst, Ipi, Icms, ValorFrete, NroPedido, SubTotal, TotalPrd, FormaPgto ) SELECT 4, PED.Entidade, PED.DataEmiss, PED.Total, PED.Desconto, PED.Vendedor, False AS Expr1, PED.Transportadora, PED.Comissao, PED.IcmSubst, PED.Ipi, PED.Icms, PED.ValorFrete, PED.NroPedido, PED.SubTotal, PED.TotalPrd, PED.FormaPgto FROM PED WHERE PED.CodPed = " & GridPed.Columns(2)
        Banco.Execute SQL
        
        Set pedidos = Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
        pedidos.Requery
        If pedidos.RecordCount > 0 Then
            pedidos.MoveLast
            VarCodPed = pedidos!CodPed
        End If
        
        Set pedidos = Banco.OpenRecordset("Select * From Itp Where Pedido = " & GridPed.Columns(2))
        pedidos.Requery
        
        
        SQL = " INSERT INTO ITP (Pedido, Produto, Complemento, Qtd, ValorUnit, Valor, Desconto, CustoDia, Loja, Entregue, AliqIcms, AliqIPI, ValorIPI, CST, Comissao , VlrVdaDia, VendaDia ) SELECT  " & VarCodPed & ", ITP.Produto, ITP.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.CustoDia, ITP.Loja, False AS Expr1, iif(isnull(ITP.AliqIcms),0,ITP.AliqIcms), iif(isnull(ITP.AliqIPI),0,ITP.AliqIPI), IIF(ISNULL(ITP.ValorIPI),0,ITP.ValorIPI), iif(ISNULL(Itp.Cst),0,Itp.Cst), IIF(ISNULL(ITP.Comissao),0,ITP.Comissao), Itp.VlrVdaDia, Itp.VendaDia FROM ITP WHERE Itp.Pedido = " & GridPed.Columns(2)
        Banco.Execute SQL
        
        '    Sql = "INSERT INTO PAG (Fornecedor, Data, DtVenc, Valor, VlrPago, Numero , NroNota ,Obs, Conta, CodCta, Tipo, DtPgto, Pedido, Cond ) VALUES ( "
        '    Sql = Sql & Datdados(1).Recordset.Fields("CodEntidade") & " , #"
        '    Sql = Sql & Format(DtData, "mm/dd/yyyy") & "# , #"
        '    Sql = Sql & Format(DtVenc, "mm/dd/yyyy") & "# , "
        '    Sql = Sql & Numero2(MebValor.Text) & " ,"
        '    Sql = Sql & Numero2(MebValor.Text) & " ,"
        '    Sql = Sql & """" & (TxtDupli) & """" & " ,"
        '    Sql = Sql & """" & (TxtNota) & """" & " ,"
        '    Sql = Sql & """" & (TxtObs) & """" & " ,"
        '    Sql = Sql & """" & (DbcConta) & """" & " ,"
        '    Sql = Sql & """" & (MebCta) & """" & " ,"
        '    Sql = Sql & """" & CmbTipo.ListIndex & """" & " ,#"
        '    Sql = Sql & Format(Dtpgto, "mm/dd/yyyy") & "# , "
        '    Sql = Sql & TxtPedido & " ,"
        '    Sql = Sql & ChkPago & ")"
        '    Banco.Execute Sql
    
        
        
        MsgBox "Pedido copiado com sucesso !"
'        FrmNota.CodPed = GridPed.Columns(2)
    End If
End If

End Sub

Private Sub BtCupom_Click()
Dim varSubTotal As Currency
Dim VarTotal As Currency
Dim Linhas As Currency

'GoTo Laser

Linhas = 0

'On Error GoTo erro
If MsgBox("Confirma Impressão em " & Porta, vbYesNo, App.Title) = vbYes Then
    
    Iniciar (False)
    
    Criterio = "CodEntidade =  " & LblCodigo
    
    
    If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
        VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Uf", App.Path & "\nfe.ini") <> "" Then
        VarUfEmite = LerINI("EMITENTE", "Uf", App.Path & "\nfe.ini")
    End If
    
    
    If LerINI("EMITENTE", "RazaoSocial", App.Path & "\nfe.ini") <> "" Then
        VarRazaoEmite = LerINI("EMITENTE", "RazaoSocial", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Fantasia", App.Path & "\nfe.ini") <> "" Then
        VarFantasiaEmite = LerINI("EMITENTE", "Fantasia", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Logradouro", App.Path & "\nfe.ini") <> "" Then
        VarEndEmite = LerINI("EMITENTE", "Logradouro", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Nro", App.Path & "\nfe.ini") <> "" Then
        VarNroEmite = LerINI("EMITENTE", "Nro", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Bairro", App.Path & "\nfe.ini") <> "" Then
        VarBairroEmite = LerINI("EMITENTE", "Bairro", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Cidade", App.Path & "\nfe.ini") <> "" Then
        VarCidadeEmite = LerINI("EMITENTE", "Cidade", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "InscEst", App.Path & "\nfe.ini") <> "" Then
        VarInscEmite = LerINI("EMITENTE", "InscEst", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Cep", App.Path & "\nfe.ini") <> "" Then
        VarCepEmite = LerINI("EMITENTE", "Cep", App.Path & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
        VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
    End If
    
              
    Set Empresa = Banco.OpenRecordset("CFG")
    DatDados(0).Recordset.FindFirst Criterio
        
    If DatDados(2).Recordset.RecordCount = 0 Then Exit Sub
        
        DatDados(2).Recordset.MoveFirst
            
        Print #1, ""
        Print #1, ""
        If VarRazaoEmite = "" Then
            Print #1, Centralizar(Left(Empresa("Cabecalho1"), 20), 40)
            Print #1, Centralizar(Left(Empresa("Cabecalho2"), 40), 40)
        Else
            Print #1, VarRazaoEmite
            Print #1, VarEndEmite & "," & VarNroEmite
            Print #1, VarCidadeEmite & "-" & VarUfEmite & "  " & VarBairroEmite
            Print #1, Format(VarCnpjEmite, "@@.@@@.@@@/@@@@-@@") & "    " & Format(VarInscEmite, "@@@.@@@.@@@.@@@")
        End If
        Print #1, "========================================"
'        Print #1, ""
        Print #1, DatDados(2).Recordset("DataEmiss") & "  " & DatDados(2).Recordset("Hora") & "     Ped.:" & psCompString(1, VarCodPed, 6)
    '    Print #1, ""
        Print #1, DatDados(0).Recordset("CodEntidade") & "-" & Left(DatDados(0).Recordset("Nome"), 30)
        Print #1, DatDados(0).Recordset("Endereco") & " , " & DatDados(0).Recordset("Nro")
        Print #1, Left(DatDados(0).Recordset("Bairro"), 15) & " - " & Left(DatDados(0).Recordset("Cidade"), 20)
        Print #1, Format(DatDados(0).Recordset("Fone"), "(###)-###-####")
        Print #1, "----------------------------------------"
        Print #1, Tab(1); "CODIGO"; Tab(8); "PRODUTO"
        Print #1, Tab(1); "QTD"; Tab(15); "UNIT."; Tab(30); "TOTAL"
        Print #1, "----------------------------------------"
       
        Do While Not DatDados(2).Recordset.Eof
            Print #1, Tab(1); psCompString(1, Trim(DatDados(2).Recordset("Produto")), 6) & "-"; Tab(8); Left(DatDados(2).Recordset("Descrição_Produto"), 35)
            Print #1, Alinhar(DatDados(2).Recordset("Qtd"), 8, 2); Tab(10); " X "; Tab(15); Alinhar(DatDados(2).Recordset("VlrVdaDia") / DatDados(2).Recordset("Qtd"), 8, 2); Tab(32); Alinhar(DatDados(2).Recordset("VlrVdaDia"), 9, 2)
            
            Linhas = Linhas + 1
            varSubTotal = varSubTotal + DatDados(2).Recordset("VlrVdaDia")
            DatDados(2).Recordset.MoveNext
        Loop
    
    DatDados(2).Refresh
    
    VarTotal = varSubTotal
    
    Print #1, ""
    Print #1, Tab(15); "SUB-TOTAL:"; Tab(29); Alinhar(CStr(varSubTotal), 12, 2)
    If MebTotDesc <> "" Then
        Print #1, Tab(15); "DESCONTO.:"; Tab(29); Alinhar(CStr(MebTotDesc), 12, 2)
        If IsNumeric(MebTotDesc) Then
            VarTotal = VarTotal - MebTotDesc
        End If
    End If
    
    If DatDados(2).Recordset.Fields("Operacao") = 2 Then
        Print #1, Tab(15); "TOTAL....:"; Tab(29); Alinhar(CStr(VarTotal), 12, 2)
    Else
        If DatDados(0).Recordset("Classificacao") = 3 Then
            varSubTotal = VarTotal * (2.3 / 100)
            VarTotal = VarTotal - varSubTotal
            Print #1, Tab(15); "FUN-RURAL:"; Tab(29); Alinhar((CStr(varSubTotal)), 12, 2)
            Print #1, Tab(15); "TOTAL....:"; Tab(29); Alinhar(CStr(VarTotal), 12, 2)
        Else
            Print #1, Tab(15); "TOTAL....:"; Tab(29); Alinhar(CStr(VarTotal), 12, 2)
        End If
    End If
    
    Print #1, ""
    If TxtObs <> "" Then
        Print #1, TxtObs
        Print #1, ""
    End If
    
    Set Parcelas = Banco.OpenRecordset("SELECT DUP.* FROM DUP Where Dup.Pedido = " & VarCodPed & " Order by Dup.NroPcl")
    
    If Parcelas.RecordCount > 0 Then
    
        Print #1, Centralizar("CONDICOES DE PAGAMENTO", 40)
        Print #1, ""
        
        Do While Not Parcelas.Eof
            Print #1, Tab(1); Parcelas("NroPcl"); Tab(7); Parcelas("DtVenc"); Tab(20); Format(Parcelas("Valor"), "#,##0.00;-#,##0.00"); Tab(28); Parcelas("DtVEnc") - Parcelas("DtEmiss"); Tab(34); "DIAS"
            Parcelas.MoveNext
        Loop
    End If
    
'    For x = 1 To 30 - Linhas
'        Print #1, ""
'    Next
    
'    Print #1, ""
    Print #1, ""
    Print #1, "Ass.Cli_________________________________"
    Print #1, ""
    Print #1, Left(CmbDados(3).Text, 15) & " - " & Left(DbcDados(2).Text, 20)
    Print #1, "Comprovante sem efeito fiscal"
    Print #1, ""
    Print #1, Centralizar(Empresa("MensagemRodape"), 40)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Fim (False)
End If
Erro:
    
    If Err.Number = 76 Then
        MsgBox "A porta " & Porta & " não foi localizada !!!", 16
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, 16
    End If
Close #1

End Sub

Private Sub BtFabrica_Click()
Dim Nome As String
Dim Apaga As String
Dim SQL As Recordset
Dim Etq As Recordset
Nome = DatDados(4).Recordset.Fields("Nome")
Apaga = "Delete * From Etq"
Banco.Execute Apaga

Set Etq = Banco.OpenRecordset("Select * From Etq")

Set SQL = Banco.OpenRecordset("SELECT ITP.Produto, PED.QtdVolume, PED.NroNt, ITP.Qtd, PED.CodPed FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE  PED.CodPed = " & GridPed.Columns(2) & "  ORDER BY PED.NroNt")
SQL.Requery
If SQL.RecordCount > 0 Then
    SQL.MoveFirst
    
    While Not SQL.Eof
        For X = 1 To SQL!Qtd
            Etq.AddNew
            Etq!Produto = SQL!Produto
            Etq!Qtd = SQL!QtdVolume
            Etq!Preco = SQL!NroNt
            Etq.Update
        Next
        SQL.MoveNext
    Wend
    'MDIPrincipal.ComRel.ShowPrinter
    MDIPrincipal.RptRel.ReportFileName = Caminho & "\IDVOLUME.RPT"
    
    MDIPrincipal.RptRel.Formulas(0) = "Nome = '" & Nome & "' "
    MDIPrincipal.RptRel.SelectionFormula = "" '{Ped.CodPed} = " & varCodPed
    
    Imprimir
    
    Banco.Execute Apaga

End If

End Sub

Private Sub BtDesconto_Click()
Dim varDesconto As Currency
Dim SQL As String
If MsgBox("Confirma desconto para o pedido " & GridPed.Columns(2), vbYesNo, App.Title) = vbYes Then
    varDesconto = InputBox("Percentual de desconto (%)")
    If IsNumeric(varDesconto) Then
        If varDesconto <> 0 Then
            SQL = "UPDATE ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed SET PED.Total = [ped].[total]*((100-" & varDesconto & ")/100), ITP.ValorUnit = [itp].[valorunit]*((100-" & varDesconto & ")/100), ITP.Valor = [itp].[valor]*((100-" & varDesconto & ")/100), ITP.VlrVdaDia = [itp].[vlrvdadia]*((100-" & varDesconto & ")/100), ITP.CustoDia = [itp].[custodia]*((100-" & varDesconto & ")/100), ITP.VendaDia = [itp].[vendadia]*((100-" & varDesconto & ")/100) WHERE (((PED.CodPed)=" & GridPed.Columns(2) & "))"
            Banco.Execute SQL
        Else
            MsgBox "Valor inválido", vbInformation, App.Title
        End If
    Else
        MsgBox "Valor inválido", vbInformation, App.Title
    End If
End If

End Sub

Private Sub BtImprimir_Click()
Dim varSubTotal As Currency
Dim VarTotal As Currency
Dim Linhas As Currency
Dim Icms As Currency
Dim TabIcms As Recordset
Dim TotIcms As Currency
Dim Base As Currency
Dim Transport As Recordset
Dim VarQtd As Currency
Dim VarNroNota As String
Dim Doc1 As String
Dim Doc2 As String
Dim VarSub As Currency

'GoTo Laser

Linhas = 0
TotIcms = 0
Base = 0
VarQtd = 0

'On Error GoTo erro
If MsgBox("Confirma Impressão em " & Porta, vbYesNo, App.Title) = vbYes Then
    If DatDados(0).Recordset("Uf") <> "SP" Then
        Set TabIcms = Banco.OpenRecordset("Select * From Icm Where Uf = " & """" & DatDados(0).Recordset("Uf") & """")
        If TabIcms.RecordCount = 0 Then
            MsgBox "Entidade sem ICMS Cadastrado !", vbExclamation, App.Title
            Exit Sub
        End If
    End If
    Iniciar (False)
    If DatDados(2).Recordset.RecordCount > 0 Then
        DatDados(2).Recordset.MoveFirst
        Set Transport = Banco.OpenRecordset("Select * From Ent Where CodEntidade = " & DatDados(2).Recordset.Fields("Transportadora"))
    End If
    
    Criterio = "CodEntidade =  " & DbcDados(0).BoundText
              
    Set Empresa = Banco.OpenRecordset("CFG")
    
    DatDados(0).Recordset.FindFirst Criterio
        
    If DatDados(2).Recordset.RecordCount = 0 Then Exit Sub
        DatDados(2).Recordset.MoveFirst
        If Not IsNull(DatDados(2).Recordset("PedSub")) Then
            VarSub = DatDados(2).Recordset("PedSub")
        End If
        
        Print #1, "==============================================================================="
        Print #1, Gde & Centralizar(Left(Empresa("Cabecalho1"), 40), 40) & Nor
        Print #1, Centralizar(Left(Empresa("Cabecalho2"), 80), 80)
        Print #1, "==============================================================================="
        Print #1, "Data:"; DatDados(2).Recordset("DataEmiss"); Tab(17); " - "; DatDados(2).Recordset("Hora"); Tab(29); "Saida: "; IIf(IsNull(DatDados(2).Recordset("DtSaida")), "", DatDados(2).Recordset("DtSaida")); Tab(49); "Ped.Ref: "; psCompString(1, txtPedidoRef, 6); Tab(66); "Pedido: "; psCompString(1, VarCodPed, 6)
        Print #1, "-------------------------------------------------------------------------------"
        Print #1, psCompString(1, DatDados(0).Recordset("CodEntidade"), 6) & "-" & Left(DatDados(0).Recordset("Nome"), 40); Tab(60); "Prev.Ent:"; Tab(70); Left(DatDados(2).Recordset("PrevEntrega"), 10)
        Print #1, DatDados(0).Recordset("Endereco") & " " & DatDados(0).Recordset("Nro"); Tab(60); "CondPgt.:"; Tab(70); Left(DatDados(2).Recordset("CondPgto"), 10)
        Print #1, DatDados(0).Recordset("Bairro") & "    " & DatDados(0).Recordset("Cidade") & "  " & DatDados(0).Recordset("Uf"); Tab(40); Format(DatDados(0).Recordset("Cep"), "@@.@@@-@@"); Tab(60); "Desconto:"; Tab(70); Left(TxtDesc, 10)
        Print #1, "Fone:"; Tab(7); Format(DatDados(0).Recordset("Fone"), "(@@)@@@@-@@@@");
        If DatDados(0).Recordset("Classificacao") = 2 Then
                Doc1 = Format(DatDados(0).Recordset("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                Doc2 = DatDados(0).Recordset("InscrEst")
                Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
            Else
                Doc1 = Format(DatDados(0).Recordset("Rg"), "@@.@@@.@@@-@")
                Doc2 = Format(DatDados(0).Recordset("Cpf"), "@@@.@@@.@@@.@@")
                Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                Print #1, Tab(52); "CPF:"; Tab(62); Doc2
        End If
        Print #1, "-------------------------------------------------------------------------------"
        Print #1, "CODIGO"; Tab(8); "PRODUTO"; Tab(57); "QTD"; Tab(66); "UNIT."; Tab(75); "TOTAL"
        Print #1, "-------------------------------------------------------------------------------"
    
        VarNroNota = DatDados(2).Recordset("NroNt")
    
        Do While Not DatDados(2).Recordset.Eof
            Print #1, Tab(1); psCompString(1, Trim(DatDados(2).Recordset("Produto")), 6); Tab(8); Left(DatDados(2).Recordset("Descrição_Produto"), 35); Tab(44); Left(DatDados(2).Recordset("Referencia"), 8); Tab(53); Alinhar(DatDados(2).Recordset("Qtd"), 8, 2); Tab(62); Alinhar(DatDados(2).Recordset("VlrVdaDia") / DatDados(2).Recordset("Qtd"), 9, 2); Tab(71); Alinhar(DatDados(2).Recordset("VlrVdaDia"), 9, 2);
'            If DatDados(0).Recordset("Uf") = "SP" Then
'                Icms = DatDados(2).Recordset("VlrVdaDia") * (DatDados(2).Recordset("Icm") / 100)
'                Print #1, Tab(75); Alinhar(DatDados(2).Recordset("Icm"), 5, 2)
'            Else
'                Icms = DatDados(2).Recordset("VlrVdaDia") * (TabIcms!Icms / 100)
'                Print #1, Tab(75); Alinhar(TabIcms!Icms, 5, 2)
'            End If
'            If DatDados(2).Recordset("Tributo") = 1 Then
'                TotIcms = TotIcms + Icms
'                Base = Base + DatDados(2).Recordset("VlrVdaDia")
'            End If
            
            Linhas = Linhas + 1
            varSubTotal = varSubTotal + DatDados(2).Recordset("VlrVdaDia")
            VarQtd = VarQtd + DatDados(2).Recordset("Qtd")
            DatDados(2).Recordset.MoveNext
        Loop
    
    DatDados(2).Refresh
    
    VarTotal = varSubTotal
    If Not IsNull(DatDados(2).Recordset("Obs")) And DatDados(2).Recordset("Obs") <> "" Then
        Print #1, ""
        Print #1, Tab(10); Left(DatDados(2).Recordset("Obs"), 50)
        Print #1, Tab(10); Mid$(DatDados(2).Recordset("Obs"), 51, 50)
        Print #1, Tab(10); Mid$(DatDados(2).Recordset("Obs"), 102, 50)
    End If
    
    Print #1, ""
    Print #1, "Total de Itens: " & Alinhar(CStr(VarQtd), 9, 3)
    
    Print #1, ""
'    Print #1, Tab(1); "Base Calculo:"; Tab(15); Alinhar(CStr(Base), 12, 2); Tab(29); "Icms:"; Tab(35); Alinhar(CStr(TotIcms), 9, 2);
    Print #1, Tab(56); "SUB-TOTAL:"; Tab(68); Alinhar(CStr(varSubTotal), 12, 2)
    If MebTotDesc <> "" Then
        Print #1, Tab(56); "DESCONTO.:"; Tab(68); Alinhar(CStr(MebTotDesc), 12, 2)
        If IsNumeric(MebTotDesc) Then
            VarTotal = VarTotal - MebTotDesc
        End If
    End If
    
    If DatDados(2).Recordset.Fields("Operacao") <> 3 Then
        If DatDados(2).Recordset("ValorFrete") <> 0 Then
            Print #1, Tab(56); "FRETE....:"; Tab(68); Alinhar(CStr(DatDados(2).Recordset("ValorFrete")), 12, 2)
            VarTotal = VarTotal + CCur(DatDados(2).Recordset("ValorFrete"))
        End If
        If DatDados(2).Recordset("PedSub") <> 0 Then
            Print #1, Tab(56); "Subst....:"; Tab(68); Alinhar(CStr(DatDados(2).Recordset("PedSub")), 12, 2)
            VarTotal = VarTotal + CCur(DatDados(2).Recordset("PedSub"))
        End If

        Print #1, Tab(56); "TOTAL....:"; Tab(68); Alinhar(CStr(VarTotal), 12, 2)
    Else
        If DatDados(0).Recordset("Classificacao") = 3 Then
            varSubTotal = VarTotal * (2.3 / 100)
            VarTotal = VarTotal - varSubTotal
            Print #1, Tab(56); "FUN-RURAL:"; Tab(68); Alinhar((CStr(varSubTotal)), 12, 2)
            Print #1, Tab(56); "TOTAL....:"; Tab(68); Alinhar(CStr(VarTotal), 12, 2)
        Else
            If DatDados(2).Recordset("ValorFrete") <> 0 Then
                Print #1, Tab(56); "FRETE....:"; Tab(68); Alinhar(CStr(DatDados(2).Recordset("ValorFrete")), 12, 2)
                VarTotal = VarTotal + CCur(DatDados(2).Recordset("ValorFrete"))
            End If
            If DatDados(2).Recordset("PedSub") <> 0 Then
                Print #1, Tab(56); "Subst....:"; Tab(68); Alinhar(CStr(DatDados(2).Recordset("PedSub")), 12, 2)
                VarTotal = VarTotal + CCur(DatDados(2).Recordset("PedSub"))
            End If
            
            
            Print #1, Tab(56); "TOTAL....:"; Tab(68); Alinhar(CStr(VarTotal), 12, 2)
        End If
    End If
    Print #1, ""
    Print #1, Cond & "(" & UCase(PassaExtenso(CStr(Round(VarTotal, 2)))) & "********************************************************************************)" & Nor
    Print #1, ""
    If CmbDados(3).ListIndex = 0 And Not IsNull(MebDados(0)) Then
        Print #1, "Previsão de entrega/pagamento: " & Format(MebDados(0), "dd/mm/yyyy")
    End If
    
    
    Set Parcelas = Banco.OpenRecordset("SELECT DUP.* FROM DUP Where Dup.Pedido = " & VarCodPed & " Order by Dup.NroPcl")
    
    If Parcelas.RecordCount > 0 Then
    
        Print #1, Centralizar("CONDICOES DE PAGAMENTO", 80)
        Print #1, ""
        
        Do While Not Parcelas.Eof
            Print #1, Tab(10); Parcelas("NroPcl"); Tab(17); Parcelas("DtVenc"); Tab(30); Format(Parcelas("Valor"), "#,##0.00;-#,##0.00"); Tab(40); Parcelas("DtVEnc") - Parcelas("DtEmiss"); Tab(50); "DIAS"
            Parcelas.MoveNext
        Loop
    End If
    Print #1, ""
    Print #1, Left(CmbDados(3).Text, 15) & " - " & Left(DbcDados(2).Text, 30)
    If Transport.RecordCount > 0 Then
        Print #1, "Transportadora: " & Transport!Nome
    Else
        Print #1, ""
    End If
    Print #1, ""
    Print #1, " Recebi os produtos acima - Ass.Cli ----------------------------------"
    
    Print #1, ""
    Print #1, Centralizar(Empresa("MensagemRodape"), 80)
    
    If Parcelas.RecordCount > 0 Then
    
        If MsgBox("Imprimir duplicata mercantil ?", vbYesNo, App.Title) = vbYes Then
              Parcelas.MoveFirst
              If Parcelas.RecordCount > 0 Then
                  VarDtVenc = Parcelas!DtVenc
              Else
                  VarDtVenc = ""
              End If
        
              Print #1, "==============================================================================="
              Print #1, Gde & NomeInscr & Nor
              Print #1, VarEndEmpresa; Tab(40); CidadeEmpresa
              Print #1, "Insc.Est.: " & VarInscr; Tab(30); "CNPJ: " & Cnpj
              Print #1, "==============================================================================="
              Print #1, "         FATURA           |       DUPLICATA         |         DATA             "
              Print #1, "   Valor     |   Numero   |   Valor     |  Numero   | Emissão     | Vencimento "
              Print #1, "-------------------------------------------------------------------------------"
    '                                            Print #1, "             |            |"; Tab(30); Alinhar(CStr(MebTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, LblTot, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
              If VarNroNota = 0 Then
                  Print #1, "             |            |"; Tab(30); Alinhar(CStr(VarTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, VarCodPed, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
              Else
                  Print #1, Tab(2); Alinhar(CStr(VarTotal), 10, 2); Tab(14); "|"; Tab(17); psCompString(1, VarCodPed, 6); Tab(27); "|"; Tab(30); Alinhar(CStr(VarTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, VarNroNota, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
              End If
              Print #1, "-------------------------------------------------------------------------------"
              Print #1, " Nome do Sacado.: "; DatDados(0).Recordset("Nome")
              Print #1, " Endereco.......: "; Trim(DatDados(0).Recordset("Endereco")) & ", " & DatDados(0).Recordset("Nro")
              Print #1, " Municipio......: "; Left(DatDados(0).Recordset("Cidade"), 20) & " - " & Left(DatDados(0).Recordset("Uf"), 2) & "  " & Format(DatDados(0).Recordset("Cep"), "@@.@@@-@@")
              Print #1, " CNPJ/CPF.......: "; Doc1; Tab(45); "Insc.Est./RG.: "; Doc2
              Print #1, "-------------------------------------------------------------------------------"
              Print #1, " Valor por Extenso: "; Cond & UCase(PassaExtenso(CStr(VarTotal))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *" & Nor
              Print #1, "-------------------------------------------------------------------------------"
              Print #1, " Reconheco/emos a exatidao desta DUPLICATA de VENDA MERCANTIL, na importancia  "
              Print #1, " acima que pagarei/emos a " & Negi & NomeInscr & NegF
              Print #1, " ou a sua ordem na praca e vencimento acima indicados."
              Print #1, " "
              Print #1, "           ____/____/_______            _______________________________"
              Print #1, "            Data do aceite                  Assinatura do Sacado       "
        End If
        Print #1, SaltoPag
    Else
        Print #1, SaltoPag
    End If
    Fim (False)
End If
Erro:
    
    If Err.Number = 76 Then
        MsgBox "A porta " & Porta & " não foi localizada !!!", 16
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, 16
    End If
Close #1
    
End Sub

Private Sub BtNovo_Click()
LIMPACampo
MebDados(1).Text = Format(Date, "dd/mm/yyyy")

SSTab1.TabEnabled(0) = False
SSTabItens.Tab = 0
CmbDados(3).SetFocus
End Sub

Private Sub BtOK_Click()
Dim Criterio As String
Dim Resp As String
Dim VarCont As Boolean
Dim VarDesc As Currency
Dim SQL As String
VarCont = True
On Error GoTo Erro


Select Case CmbDados(3).ListIndex
        Case 3
            VarCont = False
        Case 4
            VarCont = False
        Case 5
            VarCont = False
        Case 7
            VarCont = False
        Case 12
            VarCont = False
        Case 13
'            VarCont = False
        Case 14
'            VarCont = False
End Select

Set pedidos = Banco.OpenRecordset("Select * From Itp Where Pedido = " & VarCodPed)
pedidos.Requery
If pedidos.RecordCount = 0 Then
    Set pedidos = Banco.OpenRecordset("Select * From Ped Where CodPed = " & VarCodPed)
    pedidos.Requery
    
    Open Caminho & "\Delete.txt" For Append As #1
    Print #1, "Pedido; " & VarCodPed & "; " & DbcDados(0).Text & "; " & Date & "; " & Time & "; " & Maquina & "; " & CmbDados(3).Text & "; " & Alinhar(TxtDados(2), 10, 2)
    Close #1
    
    If pedidos.RecordCount > 0 Then
        pedidos.Delete
    End If
    
    BtAlterar.Enabled = False
    BtCancela.Enabled = False
    BtOK.Enabled = False
    BtCupom.Enabled = False
    BtImprimir.Enabled = False
    BtPedProm.Enabled = False
    BtTransportadora.Enabled = False
    BtParcelas.Enabled = False
    BtNovo.Enabled = True
    ImprNota1.Enabled = False
    
    MsgBox "O Pedido atual foi excluido !!!", 16
    LIMPACampo
    Exit Sub
End If

Resp = MsgBox("Confirma Pedido ?", vbYesNoCancel)

If Resp = vbCancel Then
    MebDados(4).SetFocus
    Exit Sub
ElseIf Resp = vbNo Then
            ApagaPedido
'            Frame(1).Enabled = False
'            LIMPACampo
     ElseIf Resp = vbYes Then
         If Tabela.RecordCount <> 0 Then
            If IsNumeric(TxtDados(0)) Then
                If MebDados(3) = 0 Then
                    SQL = "UPDATE PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido SET ITP.Comissao = [itp].[vlrvdadia]*([prd].[comissao]/100) Where Pedido = " & TxtDados(0)
                    Banco.Execute SQL
                Else
                    SQL = "UPDATE PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido SET ITP.Comissao = [itp].[vlrvdadia]* " & Numero4(MebDados(3) / 100) & " Where Pedido = " & TxtDados(0)
                    Banco.Execute SQL
                End If
            End If
            
            Criterio = "CodPed = " & TxtDados(0)
            Tabela.FindFirst Criterio
            Tabela.Edit
            Tabela.Fields("Vendedor") = DbcDados(2).BoundText
            Tabela.Fields("DataEmiss") = MebDados(1).Text
            Tabela.Fields("Hora") = Time
            If MebTotDesc <> "" Then
                Tabela.Fields("Desconto") = MebTotDesc
            Else
                Tabela.Fields("Desconto") = 0
            End If
            If MebDados(0).Text <> "__/__/____" Then
                Tabela.Fields("DtSaida") = MebDados(0).Text
            Else
                Tabela.Fields("DtSaida") = Null '"__/__/____"
            End If
            
            Tabela.Fields("Operacao") = CmbDados(3).ListIndex
            If IsNumeric(DbcDados(0).BoundText) Then
                Tabela.Fields("Entidade") = DbcDados(0).BoundText
            End If
            If IsNumeric(TxtDados(2).Text) Then
                Tabela("Total") = CCur(TxtDados(2).Text)
            End If
            
            If IsNumeric(TxtDados(1).Text) Then
                Tabela("SubTotal") = CCur(TxtDados(1).Text)
            End If
            
            If IsNumeric(VarTotalPrd) Then
                Tabela("TotalPrd") = CCur(VarTotalPrd)
            End If
            
            If MebTotDesc <> "" Then
                Tabela.Fields("Desconto") = MebTotDesc
            Else
                Tabela.Fields("Desconto") = 0
            End If
            Tabela("NroPedido") = txtPedidoRef.Text
            Tabela("Cfo") = Txtcfop.Text
                                    
            Tabela!Obs = TxtObs
            Tabela!FunRural = ChkFunRural.value
            Tabela.Update
            CmbDados(3).Enabled = True
            Frame(0).Enabled = False
            Frame(1).Enabled = False
            Frame(3).Enabled = False
            Frame3.Enabled = True
            BtAlterar.Enabled = True
            BtNovo.Enabled = True
            BtCancela.Enabled = False
            BtOK.Enabled = False
            BtImprimir.Enabled = True
            BtCupom.Enabled = True
            BtPedProm.Enabled = True
            BtTransportadora.Enabled = True
            BtParcelas.Enabled = True
            ImprNota1.Enabled = True
                        
            CmbDados(3).Enabled = True
            
            If VarCont Then
                FormEntrega.Show vbModal
            End If
                        
            
            Totaliza
            Tabela.Edit
            
            If MebTotDesc <> "" Then
                If MebTotDesc > 0 Then
                    If MsgBox("Voce deseja distribuir o desconto entre os produtos ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                        VarDesc = (MebTotDesc / TxtDados(1))
                        SQL = " UPDATE ITP SET ITP.ValorUnit = [itp].[valorunit]-([itp].[valorunit]* " & Numero4(Format(VarDesc, "###,##0.0000")) & "), ITP.Valor = [itp].[valor]-([itp].[valor]* " & Numero4(Format(VarDesc, "###,##0.0000")) & "), ITP.VlrVdaDia = ITP.VlrVdaDia - ([itp].[vlrvdadia]* " & Numero4(Format(VarDesc, "###,##0.0000")) & "), ITP.Comissao = [itp].[Comissao]-([itp].[Comissao]* " & Numero4(Format(VarDesc, "###,##0.0000")) & ") WHERE (((ITP.Pedido)= " & TxtDados(0) & "))"
                        Banco.Execute SQL
                        Tabela.Fields("Desconto") = 0
                        MebDesc = 0
                        MebTotDesc = 0
                        TxtDados(1) = TxtDados(2)
                        
                    Else
                        Tabela.Fields("Desconto") = MebTotDesc
                    End If
                 Else
                        Tabela.Fields("Desconto") = 0
                 End If
            Else
                Tabela.Fields("Desconto") = 0
            End If
            
            If IsNumeric(TxtDados(2).Text) Then
                Tabela("Total") = CCur(TxtDados(2).Text)
            End If
            
            If IsNumeric(TxtDados(1).Text) Then
                Tabela("SubTotal") = CCur(TxtDados(1).Text)
            End If
            
            If IsNumeric(VarTotalPrd) Then
                Tabela("TotalPrd") = CCur(VarTotalPrd)
            End If
            
                        
            Tabela.Update
                                    
            If CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 7 Then
                FrmParcelas.CodDoc = TxtDados(0)
                FrmParcelas.CodEnt = LblCodigo
                FrmParcelas.MebVlrPed = Format(TxtDados(2), "#,##0.00")
                FrmParcelas.TxtNome = Label(3) ' psCompString(1, LblCodigo, 6) & "-" &  ' DbcDados(0).Text
                FrmParcelas.Show vbModal
            Else
                If VarOperacao <> "" Then
                    SQL = " DELETE DUP.*, DUP.Pedido FROM DUP WHERE DUP.Pedido = " & TxtDados(0)
                    Banco.Execute SQL
                    VarOperacao = ""
                End If
            End If
            
            If CmbDados(3).ListIndex = 10 Then
                If MsgBox("Deseja relacionar essa devolução a algum pedido ?", vbYesNo, App.Title) = vbYes Then
                    FormDevolucao.Show
                    
                End If
            End If
         End If
End If

SSTab1.TabEnabled(0) = True
Exit Sub
    
Erro:
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End Sub

Private Sub BtParcelas_Click()
Dim Parc As Currency
Dim SQL As String
If CmbDados(3).ListIndex = 0 Then
    FormLancaPag.DbcPesquisa = LblCodigo
    FormLancaPag.MebValor = TxtDados(2)
'    FormLancaPag.TxtDupli = TxtDados(0) & "/1"
    FormLancaPag.TxtNota = txtPedidoRef
    FormLancaPag.TxtPedido = TxtDados(0)
    FormLancaPag.Show vbModal
Else
    If CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 3 Or CmbDados(3).ListIndex = 7 Or CmbDados(3).ListIndex = 13 Or CmbDados(3).ListIndex = 14 Or CmbDados(3).ListIndex = 4 Then
        FrmParcelas.CodDoc = TxtDados(0)
        FrmParcelas.CodEnt = DbcDados(0).BoundText
        FrmParcelas.MebVlrPed = Format(TxtDados(2), "#,##0.00")
        FrmParcelas.TxtNome = Label(3)  ' DbcDados(0).Text
        FrmParcelas.Show vbModal
    Else
        MsgBox "Natureza da Operação Inválida", vbInformation, App.Title
    End If
End If
End Sub

Private Sub BtPedPrazo_Click()
If TxtPesqNota.Text = "" Then
'    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.DataEmiss = #" & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed"
    
    DatDados(5).RecordSource = "SELECT DISTINCTROW PED.CodPed, PED.Operacao, PED.Entidade, PED.DataEmiss, PED.Total, PED.Desconto, PED.Vendedor, Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.NroPedido FROM PED LEFT JOIN DUP ON PED.CodPed = DUP.Pedido WHERE PED.Operacao = 2 AND PED.DataEmiss = #" & Format(DtData, "mm/dd/yyyy") & "# AND DUP.Pedido Is Null "
End If
DatDados(5).Refresh
If DatDados(5).Recordset.RecordCount <> 0 Then
    DatDados(5).Recordset.MoveLast
    VarContPed = DatDados(5).Recordset.RecordCount
End If

End Sub

Private Sub BtPedProm_Click()
    funcaoImpressao 0
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub


Private Sub BtPesquisa_Click()
    FrmPesquisaGeral.Campos1 = "CodEntidade"
    FrmPesquisaGeral.Campos2 = "Nome"
    FrmPesquisaGeral.Tabela = "ENT"
    FrmPesquisaGeral.TipoResultado = 1
    FrmPesquisaGeral.QtdCampos = 1
    
    FrmPesquisaGeral.Caption = "Pesquisa de Entidades"
    FrmPesquisaGeral.Show vbModal
    DbcDados(0).Text = RetornoPesq
    DbcDados(0).SetFocus

End Sub

Private Sub BtSelo_Click()
Dim Nome As String
Dim Apaga As String
Dim SQL As Recordset
Dim Etq As Recordset
Dim VarSql As String
Nome = DatDados(4).Recordset.Fields("Nome")
Apaga = "Delete * From Etq"
Banco.Execute Apaga

Set Etq = Banco.OpenRecordset("SELECT ETQ.*, PRD.Descrição_Produto, ETQ.Codigo FROM ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd ORDER BY ETQ.Codigo")

Set SQL = Banco.OpenRecordset("SELECT ITP.Produto, ITP.Complemento, PED.QtdVolume, PED.NroNt, ITP.Qtd, PED.CodPed, PRD.Descrição_Produto FROM (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.CodPed = " & GridPed.Columns(2) & "  ORDER BY PED.NroNt")
' VarSql = " SELECT ITP.Produto, ITP.Complemento, PED.QtdVolume, PED.NroNt, ITP.Qtd, PED.CodPed FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE  PED.CodPed = " & GridPed.Columns(2) & "  ORDER BY PED.NroNt "

SQL.Requery
If SQL.RecordCount > 0 Then
    SQL.MoveFirst
    
    While Not SQL.Eof
        For X = 1 To SQL!Qtd
            Etq.AddNew
            Etq!Produto = SQL!Produto
            Etq!Qtd = SQL!QtdVolume
            Etq!Preco = SQL!NroNt
            Etq!Obs = Left(Trim(SQL!Descrição_Produto) & "  -  " & Trim(SQL!Complemento), 80)
            Etq.Update
        Next
        SQL.MoveNext
    Wend
    Etq.Requery
    FormIdentificador.TxtNome = Nome
    FormIdentificador.TxtCod = GridPed.Columns(0)

    FormIdentificador.Show 'vbModal
    

End If

End Sub

Private Sub BtTransportadora_Click()
    funcaoImpressao 1
End Sub

Private Sub BtXml_Click()
FormXml.Show
End Sub

Private Sub ChkErro_Click()

If ChkErro.value = 1 Then
    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.Entregue = True  ORDER BY PED.CodPed"
    DatDados(5).Refresh
    If DatDados(5).Recordset.RecordCount <> 0 Then
       DatDados(5).Recordset.MoveLast
       ImprimeDupi.Enabled = True
       BtConsolida.Enabled = True
       BtSelo.Enabled = True
       BtContEntrega.Enabled = True
       ImprNota.Enabled = True
    Else
       ImprimeDupi.Enabled = False
       BtConsolida.Enabled = False
       BtSelo.Enabled = False
       BtContEntrega.Enabled = False
       ImprNota.Enabled = False
    End If
    GridPed.Refresh
End If

End Sub

Private Sub CmbAcab_Change()
QtdCorpo = Len(Trim(CmbAcab))
If Len(Trim(CmbAcab)) >= Itens("Complemento").Size Then
    MsgBox "Texto muito grande para esse campo", vbInformation
    CmbAcab = Left((Trim(CmbAcab)), Itens("Complemento").Size)
    CmbAcab.SetFocus
End If
'If QtdCorpo >= 20 Then
'    ContComplemento = True
'Else
'    ContComplemento = False
'End If

End Sub

Private Sub CmbAcab_LostFocus()
If CmbAcab.Text <> "" Then
    MebAcrescimo.Enabled = True
    MebVlr.SetFocus
    MebAcrescimo.Text = 0
Else
    MebAcrescimo.Enabled = True
    MebVlr.SetFocus
    MebAcrescimo.Text = 0
End If
End Sub



Private Sub CmbDados_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then
    If CmbDados(3).ListIndex < 0 Then
        MsgBox "Natureza inválida !", vbInformation, App.Title
        CmbDados(3).ListIndex = 1
        CmbDados(3).SetFocus
    End If
End If
End Sub

Private Sub CmbDados_LostFocus(Index As Integer)
Dim VarTemp As String
Select Case Index
    Case 3
        If VarOperacao <> "" Then
            If VarOperacao <> "=" Then
                VarTemp = CmbDados(3).Text
                If Right(CmbDados(3).Text, 1) <> VarOperacao Then
                    MsgBox "Atenção ! Essa operação pode acarretar divergência de estoque", vbInformation, App.Title
'                    CmbDados(3).Text = VarTemp
'                    CmbDados(3).SetFocus
'                    Exit Sub
                End If
            End If
        End If
        '    If DbcDados(0).Text = "" Then
                If CmbDados(3).ListIndex = 0 Or CmbDados(3).ListIndex = 11 Or CmbDados(3).ListIndex = 12 Then
        '            Frame4.Caption = "Identificação do Fornecedor"
                    DatDados(0).RecordSource = "Select * From Ent Where (Tipo = 2 or Tipo = 5) and ativo = true   Order by Nome"
'                    TxtCfop.Enabled = True
                ElseIf CmbDados(3).ListIndex = 8 Or CmbDados(3).ListIndex = 9 Then
        '                Frame4.Caption = "Identificação do Cliente/Fornecedor"
                        DatDados(0).RecordSource = "Select * From ENT Order by Nome"
                    Else
        '                Frame4.Caption = "Identificação do Cliente"
                        DatDados(0).RecordSource = "Select * From ENT where (Tipo = 1 or tipo >=5) and ativo = true  Order by Nome"
'                        TxtCfop.Enabled = False
                End If
'                If DbcDados(0).Text = "" Then
                    DatDados(0).Refresh
'                End If
        '    End If
            If CmbDados(3).Text <> "" Then
                Frame(0).Enabled = True
                Frame3.Enabled = True
        '       MebDados(1).SetFocus
            End If
         If DbcDados(0).Text <> "" And CmbDados(3).ListIndex = 2 Then

         End If
         
End Select
End Sub


Function PegaTipo(Operacao As Integer) As String
Dim Retorno As String


Select Case Operacao
    Case 0
        Retorno = "Entrada de Produtos"
    Case 1
        Retorno = "Venda Vista"
    Case 2
        Retorno = "Venda Parcelada"
    Case 3
        Retorno = "Ordem de Serviço"
    Case 4
        Retorno = "Orçamento"
    Case 5
        Retorno = "Consignação"
    Case 6
        Retorno = "Retirada/Doação"
    Case 7
        Retorno = "Consórcio"
    Case 8
        Retorno = "Transferência (Remessa)"
    Case 9
        Retorno = "Transferência (Retorno)"
    Case 10
        Retorno = "Devolução de Venda"
    Case 11
        Retorno = "Devolução de Compra"
End Select

PegaTipo = Retorno

End Function



Private Sub cMBoRDEM_Click()
If TxtDados(0) <> "" Then
    DatDados(2).RecordSource = "SELECT Itp.Entregue, Tributo,Icm, Descrição_Produto,CondPgto, Itp.Referencia, Itp.complemento,Ped.Entregue,DataEmiss,Ped.Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & TxtDados(0) & " Order by Descrição_Produto "
    DatDados(2).Refresh
End If
End Sub

Private Sub CmdDescAtacado_Click()
Dim SQL As String
Dim VarPerc As String
VarPerc = ""
VarPerc = InputBox("Percentual de sobre o custo (%)")
If VarPerc <> "" Then
    SQL = "UPDATE ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd SET ITP.Valor = (([prd].[custo]+([prd].[custo]*(" & VarPerc & "/100)))*[itp].[qtd]), ITP.VlrVdaDia = (([prd].[custo]+([prd].[custo]*(" & VarPerc & "/100)))*[itp].[qtd]) ,  ITP.ValorUnit = (([prd].[custo]+([prd].[custo]*(" & VarPerc & "/100))))  WHERE (((ITP.Pedido)= " & GridPed.Columns(2) & " ))"
    Banco.Execute SQL
End If
End Sub

Private Sub CmdPedMeiaFolha_Click()
    funcaoImpressao 4
End Sub

Private Sub Command1_Click()
    funcaoImpressao 5
End Sub

Private Sub dbcdados_Change(Index As Integer)
Dim Criterio As String
On Error Resume Next
Select Case Index
Case 0
    If DbcDados(0).BoundText <> "" Then
        Criterio = "CodEntidade = " & DbcDados(0).BoundText
        Entidade.FindFirst Criterio
        Label(0).Caption = Trim(Entidade("Endereco")) & " ," & Entidade("Nro")
        Label(1).Caption = Trim(Entidade("Cidade")) & "-" & Entidade("Uf")
        Label(3).Caption = psCompString(1, Entidade("CodEntidade"), 6) & " - " & Entidade("Nome")
        If Entidade!Classificacao = 2 Then
            If Entidade!CGC <> "" Then
                Label(2).Caption = Format(Entidade!CGC, "@@.@@@.@@@/@@@@-@@") & "   |   " & Entidade!InscrEst
            Else
                Label(2).Caption = Format(Entidade!Cpf, "@@@.@@@.@@@-@@") & "   |   " & Entidade!InscrEst
            End If
        ElseIf Entidade!Classificacao = 1 Then
            Label(2).Caption = Format(Entidade!Cpf, "@@@.@@@.@@@-@@") & "   |   " & Entidade!InscrEst
        End If
            
        If Left(Entidade("InscrEst"), 1) = "P" And CmbDados(3).ListIndex = 0 Then
              ChkFunRural.value = 1
        Else
              ChkFunRural.value = 0
        End If
        If Not IsNull(Entidade("Fone")) Then
            Label(4).Caption = Format(Entidade("Fone"), "(##) ####-####")
            If Not IsNull(Entidade("Fax")) Then
                Label(4).Caption = Format(Entidade("Fone"), "(##) ####-####") & "   --   " & Format(Entidade("Fax"), "(##) ####-####")
            End If
        End If
                
    '        Label(2).Caption = Entidade("Nome")
        LblCodigo.Caption = Entidade("CodEntidade")
        Criterio = "CodEntidade = " & Entidade("CodEntidade")
        DatDados(0).Recordset.FindFirst Criterio
    End If
Case 1
    If Trim(DbcDados(1).Text) <> "" And IsNumeric(DbcDados(1).BoundText) Then
        If DbcDados(1).Text <> "" And Not IsNumeric(DbcDados(1).Text) Then
            Criterio = "Prd.CodPrd = " & DbcDados(1).BoundText 'Datdados(1).Recordset.Fields("CodPrd")
            Produtos.FindFirst Criterio
        End If
        If Not Produtos.NoMatch And DbcDados(1).BoundText <> "" Then
            If CmbDados(3).ListIndex = 0 Or CmbDados(3).ListIndex = 11 Or CmbDados(3).ListIndex = 12 Or CmbDados(3).ListIndex = 15 Then
                If Not IsNull(Produtos("Preco")) Then
                    If Produtos("Preco") <> 0 Then
                        If ContCompra = True Then
                            MebDados(5).Text = Round(Produtos("Preco"), 4)
                        Else
                            MebDados(5).Text = Round(Produtos("Custo"), 4)
                        End If
                    Else
                        MebDados(5).Text = Round(Produtos("Custo"), 4)
                    End If
                Else
                    MebDados(5).Text = Round(Produtos("Custo"), 4)
                End If
                MebVlr = MebDados(5).Text
                MebDados(2).Text = 0
            Else
                If CmbDados(3).ListIndex = 2 Then
                    If VarPzoProduto = True Then
                        If VarPcoPrazo = True Then
                            MebDados(5).Text = Round(Produtos("PcoPrazo"), 3)
                        Else
                            MebDados(5).Text = Round(Produtos("Venda"), 3) ' + (Produtos("Venda") * (ValIndices / 100)), 3)
                        End If
                    Else
                        If VarPcoPrazo = True Then
                            MebDados(5).Text = Round(Produtos("Venda") + (Produtos("Venda") * (ValIndices / 100)), 3)
                        Else
                            MebDados(5).Text = Round(Produtos("Venda"), 3)
                        End If
                    
                    End If
                    
                Else
                    If Produtos("Promocao") = True Then
                        MebDados(5).Text = Round(Produtos("Venda_Promocao"), 3)
                    Else
                        MebDados(5).Text = Round(Produtos("Venda"), 3)
                    End If
                End If
                If Dolar <> 0 Then
                    If Produtos("Est_ideal") <> 0 Then
                         MebDados(5).Text = Produtos("est_ideal") * Round(Dolar, 3)
                    Else
                        MebDados(5).Text = Round(Produtos("Venda"), 3)
                    End If
                End If
                
                If VarLista1.RecordCount > 0 Then
                    VarLista1.FindFirst "CodPrd = " & DbcDados(1).BoundText
                    If Not VarLista1.NoMatch Then
                        MebDados(5).Text = Round(VarLista1!Valor, 3)
                    End If
                End If
                                
                MebVlr = MebDados(5).Text
'                MebDados(2).Text = Produtos("Comissao")
                MebComissao.Text = Produtos("Comissao")
            End If
            VarCusto = Produtos("Custo")
        Else
            MebDados(5).Text = ""
            MebDados(2).Text = ""
        End If
    End If
Case 3
    If DbcDados(3).BoundText <> "" Then
        Criterio = "CodEntidade = " & DbcDados(3).BoundText
        DatDados(4).Recordset.FindFirst Criterio
    End If
End Select
Exit Sub

erroped:

    MsgBox Err.Description & Err.Number

End Sub


Private Sub DbcDados_GotFocus(Index As Integer)
If Index <> 2 Then
    If Index = 1 And MebVlr = "" Then
        SendKeys "{DEL}"
    End If
End If

If Index = 1 Then
'        If DbcDados(1).Text = "" Then
'            FrmPesquisa.Show vbModal
'        End If
'    End If
    If MebDados(4).Text = "" Then
'        If VarGrade = False Then
            DbcDados(1).Text = ""
'        End If
        MebDados(4).SetFocus
    End If
End If


End Sub

'Private Sub DbcDados_GotFocus(Index As Integer)
'DbcDados(1).Text = ""
'End Sub

Private Sub DbcDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Me.KeyPreview = False
Dim Falta As Boolean
Falta = False
'On Error Resume Next

If KeyCode = 9 Or KeyCode = 13 Then
    Select Case Index
        Case 0
            If DbcDados(0).Text = "" Or DbcDados(0).Text = "(Nenhum)" Then
                DbcDados(0).SetFocus
                Exit Sub
            End If
            Dim Criterio As String
            If IsNumeric(DbcDados(0).Text) Then
               Criterio = "CodEntidade = " & DbcDados(0)
               DatDados(0).Recordset.FindFirst Criterio
               If DatDados(0).Recordset.NoMatch Then
                   MsgBox "Entidade não encontrada !!!", 16
                   DbcDados(0).Text = ""
                   DbcDados(0).SetFocus
               End If
            Else
                If IsNumeric(DbcDados(0).BoundText) Then
                    If DbcDados(0).BoundText = 0 Then
                        If DbcDados(0).Text <> "" Then
                            MsgBox "Entidade não encontrada !!!", 16
                            DbcDados(0).SetFocus
                            Exit Sub
                        Else
                            SSTab1.Tab = 0
                            GridCliFor.SetFocus
                        End If
                    End If
                End If
            End If
'            DbcDados(0).Text = DatDados(0).Recordset.Fields("Nome")
            If IsNumeric(DbcDados(0).BoundText) Then
                If VarLimiteNaParcela = True And CmbDados(3).ListIndex = 2 Then
                    With DatDados(0).Recordset
                        If .Fields("Endereco") = "" Then Falta = True
                        If .Fields("Nro") = "" Then Falta = True
                        If .Fields("Bairro") = "" Then Falta = True
                        If .Fields("Cidade") = "" Then Falta = True
                        If .Fields("Uf") = "" Then Falta = True
                        If .Fields("Cep") = "" Then Falta = True
'                        If .Fields("Cpf") = "" Then Falta = True
'                        If .Fields("Rg") = "" Then Falta = True
'                        If .Fields("Cpf") = "" Then Falta = True
                        If .Fields("Profissao") = "" Then Falta = True
                        If .Fields("Trabalho") = "" Then Falta = True
                        If .Fields("Fone_Trabalho") = "" Then Falta = True
                        If .Fields("Salario") = "" Then Falta = True
                        If .Fields("Refer1") = "" Then Falta = True
                        If .Fields("Refer2") = "" Then Falta = True
                        If .Fields("Refer3") = "" Then Falta = True
                    End With
                    If Falta = True Then
                        MsgBox "Cadastro Incompleto !!", vbInformation, App.Title
                        DbcDados(0).Text = ""
                        DbcDados(0).SetFocus
                        
                        Exit Sub
                    End If
                End If
                If DatDados(0).Recordset.Fields("Status") = 4 And (DbcDados(0).BoundText <> 0 And CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 5) Then
                      MsgBox "Cliente Bloqueado ! ", vbExclamation, App.Title
                      DbcDados(0).Text = ""
                      DbcDados(0).SetFocus
                      Exit Sub
                 End If
            
                If DbcDados(0).BoundText <> 0 And (CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 5) Then
'                       If VarLimiteNaParcela = True And CmbDados(3).ListIndex = 2 Then
                            Set Parcelas = Banco.OpenRecordset("SELECT Entidade,  Sum(DUP.Valor) AS SomaDeValor FROM Dup WHERE  DtVenc < #" & Format(Date, "mm/dd/yyyy") & "# GROUP BY Entidade, Cond HAVING Entidade = " & DbcDados(0).BoundText & " AND Cond = FALSE")
'                       Else
'                            Set Parcelas = Banco.OpenRecordset("SELECT Entidade,  Sum(DUP.Valor) AS SomaDeValor FROM Dup WHERE  DtVenc < #" & Format(Date, "mm/dd/yyyy") & "# GROUP BY Entidade, Cond HAVING Entidade = " & DbcDados(0).BoundText & " AND Cond = FALSE")
'                       End If
                       Parcelas.Requery
                       If Parcelas.RecordCount <> 0 Then
                           LblAtraso = Format(Parcelas!SomaDeValor, " ##,##0.00")
                           
                            MsgBox "Cliente tem parcelas atrasadas, favor consultar o setor de cobrança !", vbExclamation, App.Title
                            If DatDados(0).Recordset.Fields("Status") = 5 Then
                                 MsgBox "Cliente Bloqueado ! ", vbExclamation, App.Title
                                 DbcDados(0).Text = ""
                                 DbcDados(0).SetFocus
                                 Exit Sub
                            End If
                       End If
                End If
                If Entidade!Desconto <> "" Then
                    TxtDesc.ToolTipText = Entidade!Desconto
                End If
                If VarAssistencia Then
                    SQL = "SELECT PED.Entidade, PED.DataEmiss, PED.DtSaida, PED.Operacao FROM PED WHERE PED.Entidade = " & DbcDados(0).BoundText & " AND PED.DataEmiss<>[ped].[dtsaida] AND PED.DtSaida <= #" & Format(Date, "mm/dd/yyyy") & "# AND PED.Operacao =14"
                    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, PED.Entidade, PED.DataEmiss, PED.DtSaida, PED.Operacao FROM PED WHERE PED.Entidade = " & DbcDados(0).BoundText & " AND PED.DataEmiss<>[ped].[dtsaida] AND PED.DtSaida <= #" & Format(Date, "mm/dd/yyyy") & "# AND PED.Operacao =14"
                    DatDados(5).Refresh
                    If DatDados(5).Recordset.RecordCount > 0 Then
                            MsgBox "Cliente tem assistência pendente, favor verificar o setor responsavel ! ", vbExclamation, App.Title
                    End If
                    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, PED.CodPed, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=3,'Serviço',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.DataEmiss, PED.DtFat, Ped.DtSaida, PED.Total, PED.Desconto FROM PED WHERE (((PED.Entidade)=0)) ORDER BY PED.CodPed"
                    
                End If
                
                If Entidade!Vendedor <> 0 Then
                    Criterio = "CodEntidade = " & Entidade!Vendedor
                    DatDados(3).Recordset.FindFirst Criterio
                    DbcDados(2).Text = DatDados(3).Recordset("Nome")
                    If DatDados(3).Recordset("Comissao") <> "" Then
                        MebDados(3).Text = DatDados(3).Recordset("Comissao")
                    End If
                End If
            End If
            ObsEnt.Text = ""
            If Not IsNull(Entidade!Obs) Then
                ObsEnt.Text = Entidade!Obs
            End If
            
            If Not IsNull(Entidade!PesAut) Then
                ObsEnt.Text = ObsEnt.Text & " " & Entidade!PesAut
            End If
                        
         Case 1
'            If varqtdfixo = 1 And CmbDados(3).ListIndex = 5 Then
                
'            End If
            If DbcDados(1).Text = "" Then
                 FrmPesquisa.Show vbModal
            End If
                           
            If IsNumeric(DbcDados(1).Text) Then
                If Len(CDec(DbcDados(1))) >= 8 Then
                    Set VarBarra = Banco.OpenRecordset("select * From Bar Where BarBarra = " & """" & DbcDados(1) & """")
                    If VarBarra.RecordCount > 0 Then
                        Criterio = "CodPrd = " & VarBarra!Produto
                    Else
                        Criterio = "CodPrd = " & CDbl(DbcDados(1))
                    End If
                Else
                    Criterio = "CodPrd = " & CDbl(DbcDados(1))
                End If
                DatDados(1).Recordset.FindFirst Criterio
                If DatDados(1).Recordset.NoMatch Then
                     GoTo Prd_Nao_Encontrado
                Else
'                    DbcDados(1).Text = DatDados(1).Recordset.Fields("Descrição_Produto")
                    DbcDados(1).BoundText = DatDados(1).Recordset.Fields("CodPrd")
                End If
            Else
               If Trim(DbcDados(1).Text) <> "" And IsNumeric(DbcDados(1).BoundText) Then
                    If DbcDados(1).BoundText = 0 Then
                         GoTo Prd_Nao_Encontrado
                    Else
                        Criterio = "CodPrd = " & DbcDados(1).BoundText
                        DatDados(1).Recordset.FindFirst Criterio
                    End If
                Else
                         GoTo Prd_Nao_Encontrado
                End If
            End If
            
            CmbAcab.Text = ""
            If Acabamento.RecordCount <> 0 Then
                CmbAcab.Enabled = True
                Acabamento.MoveFirst
                CmbAcab.Clear
                While Not Acabamento.Eof
                    CmbAcab.AddItem Acabamento("Descricao")
                    Acabamento.MoveNext
                Wend
            End If
            If CmbDados(3).ListIndex = 12 Then
                If DatDados(1).Recordset("CodBar") <> "" Then
                    CmbAcab.Enabled = True
                    CmbAcab.Text = DatDados(1).Recordset("CodBar")
                End If
            Else
                If DatDados(1).Recordset("complemento") <> "" Then
                    CmbAcab.Enabled = True
                    CmbAcab.Text = DatDados(1).Recordset("complemento")
                End If
            End If
            If CmbAcab.Text = "" Then
                If DatDados(1).Recordset("Marca") <> "" Then
                    CmbAcab.Enabled = True
                    CmbAcab.Text = DatDados(1).Recordset("Marca")
                End If
            End If
'            DbcLoja.Text = ""
            DbcLoja.Enabled = True
            If CmbDados(3).ListIndex = 0 Then
                DatDados(7).RecordSource = "SELECT CLOJ.* FROM CLOJ ORDER BY CLOJ.Descricao"
            Else
                DatDados(7).RecordSource = "SELECT LOJ.*, CLOJ.* FROM CLOJ INNER JOIN LOJ ON CLOJ.Loja = LOJ.DescricaoLoja WHERE Produto = " & DatDados(1).Recordset.Fields("CodPrd") & " ORDER BY Descricaoloja"
            End If
            DatDados(7).Refresh
'            If VarLojaUnica = False Then
            
                If DatDados(7).Recordset.RecordCount <> 0 Then
                    DatDados(7).Recordset.MoveFirst
                    DbcLoja.Clear
                    While Not DatDados(7).Recordset.Eof
                        If CmbDados(3).ListIndex = 0 Then
                            
                            DbcLoja.AddItem DatDados(7).Recordset.Fields("Descricao")
                        Else
                            DbcLoja.AddItem DatDados(7).Recordset.Fields("DescricaoLoja") & "  -  " & DatDados(7).Recordset("Estoque")
                        End If
                        DatDados(7).Recordset.MoveNext
                    Wend
                    DbcLoja.ListIndex = 0
                End If
'            Else
'               If DatDados(7).Recordset.RecordCount <> 0 Then
'                    DatDados(7).Recordset.MoveFirst
'                    DbcLoja.Clear
'                    If CmbDados(3).ListIndex = 0 Then
'                        DbcLoja.AddItem DatDados(7).Recordset.Fields("Descricao")
'                    Else
'                        DbcLoja.AddItem DatDados(7).Recordset.Fields("DescricaoLoja") & "  -  " & DatDados(7).Recordset("Estoque")
'                    End If
'                    DbcLoja.ListIndex = 0
'               End If
'            End If
            If DatDados(1).Recordset("Estoque") <= 0 Then
                DbcLoja.BackColor = &HC0C0FF
                If VarBloquearEst = True Then
                    If CmbDados(3).ListIndex = 1 Or CmbDados(3).ListIndex = 2 Then
                        MsgBox "Produto sem estoque !", vbInformation
'                        DbcDados(0).Text = ""
                        DbcDados(1).SetFocus
                    End If
                End If
            Else
                DbcLoja.BackColor = &H80000005
            End If
            If (CmbDados(3).ListIndex = 5 Or CmbDados(3).ListIndex = 1 Or CmbDados(3).ListIndex = 2) And VarGrade = False And VarPrecoPedido = False Then
                MebDados(6).Text = MebVlr
                MebDados_LostFocus 6
'                DbcDados(1).Text = ""
                
                If VarQtdFixa <> 1 Then
'                    MebDados(4).Text = VarQtdFixa

'                    SendKeys "{tab}"
'                    DbcDados(1).Text = ""
'                    DbcDados(1).SetFocus
'                    DbcDados(1).SetFocus
                End If
                MebDados(4).SetFocus
                
            End If
            
            If VarComposicao = True Then
                Dim VarDbComp As Recordset
                Set VarDbComp = Banco.OpenRecordset("Select * from Comp Where Produto =  " & DatDados(1).Recordset.Fields("CodPrd"))
                If VarDbComp.RecordCount > 0 Then
                    
                End If
            End If
            
        Case 2
                If DbcDados(2).Text = "" Then
                    DbcDados(2).SetFocus
                    Exit Sub
                End If
                If IsNumeric(DbcDados(2).Text) Then
                   Criterio = "CodEntidade = " & DbcDados(2)
                   DatDados(3).Recordset.FindFirst Criterio
                   If DatDados(3).Recordset.NoMatch Then
                       MsgBox "Vendedor não encontrado !!!", 16
                       DbcDados(2).Text = ""
                       DbcDados(2).SetFocus
                       Exit Sub
                   Else
                       DbcDados(2).Text = DatDados(3).Recordset.Fields("Nome")
                   End If
                Else
                    If DbcDados(2).BoundText = "" Then
                        If DbcDados(2).Text <> "" Then
                            MsgBox "Vendedor não encontrado !!!", 16
                            DbcDados(2).SetFocus
                            Exit Sub
                        End If
                    End If
                End If
        
        End Select
    End If
    
Me.KeyPreview = True

Exit Sub

Prd_Nao_Encontrado:
    MsgBox "Produto não encontrado !!!", 16
'    DbcDados(1).Text = ""
    DbcDados(1).SetFocus
End Sub


Private Sub DbcDados_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
'         On Error Resume Next
    If Index = 3 Then
        If IsNumeric(DbcDados(3).BoundText) Then
            DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE (((PED.Entidade)=" & DbcDados(3).BoundText & ")) ORDER BY PED.CodPed"
            DatDados(5).Refresh
            If DatDados(5).Recordset.RecordCount <> 0 Then
               VarContPed = DatDados(5).Recordset.RecordCount
               DatDados(5).Recordset.MoveLast
               ImprimeDupi.Enabled = True
               BtConsolida.Enabled = True
               BtSelo.Enabled = True
               BtContEntrega.Enabled = True
               ImprNota.Enabled = True
            Else
               ImprimeDupi.Enabled = False
               BtConsolida.Enabled = False
               BtSelo.Enabled = False
               BtContEntrega.Enabled = False
               ImprNota.Enabled = False
            End If
            GridPed.Refresh
        End If

        If Not IsNumeric(DbcDados(0).BoundText) Then
            DbcDados(3).Text = DatDados(4).Recordset("Nome")
        End If
    End If
    
    If Index = 2 Then
        
        Criterio = "Nome = " & """" & DbcDados(2).Text & """"
        DatDados(3).Recordset.FindFirst Criterio
    
        If DatDados(3).Recordset("Comissao") <> "" Then
            MebDados(3).Text = DatDados(3).Recordset("Comissao")
        End If
    End If
            
    If Index = 4 Then
    
        If IsNumeric(DbcDados(4).BoundText) Then
            
            DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ent.Convenio, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE (((Ent.Convenio)=" & DbcDados(4).BoundText & ")) ORDER BY PED.CodPed"
            'Sql = "SELECT Ped.Convenio, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE (((PED.Convenio)=" & DbcDados(4).BoundText & ")) ORDER BY PED.CodPed"
            DatDados(5).Refresh
            If DatDados(5).Recordset.RecordCount <> 0 Then
               VarContPed = DatDados(5).Recordset.RecordCount
               DatDados(5).Recordset.MoveLast
               ImprimeDupi.Enabled = True
               BtConsolida.Enabled = True
               BtSelo.Enabled = True
               BtContEntrega.Enabled = True
               ImprNota.Enabled = True
            Else
               ImprimeDupi.Enabled = False
               BtConsolida.Enabled = False
               BtSelo.Enabled = False
               BtContEntrega.Enabled = False
               ImprNota.Enabled = False
            End If
            GridPed.Refresh
        End If
            
    End If
    
End If
End Sub

Private Sub DbcDados_LostFocus(Index As Integer)
Dim Criterio As String
Dim NomeObs As String
If Index = 0 Then
    If Not IsNumeric(DbcDados(0).BoundText) Then
        MsgBox "Entidade não encontrada", vbExclamation, App.Title
        DbcDados(0).Text = ""
        DbcDados(0).SetFocus
    Else
    
            If IsNumeric(DbcDados(0).BoundText) Then
                If VarLimiteNaParcela = True And CmbDados(3).ListIndex = 2 Then
                    With DatDados(0).Recordset
                        If .Fields("Endereco") = "" Then Falta = True
                        If .Fields("Nro") = "" Then Falta = True
                        If .Fields("Bairro") = "" Then Falta = True
                        If .Fields("Cidade") = "" Then Falta = True
                        If .Fields("Uf") = "" Then Falta = True
                        If .Fields("Cep") = "" Then Falta = True
'                        If .Fields("Cpf") = "" Then Falta = True
'                        If .Fields("Rg") = "" Then Falta = True
'                        If .Fields("Cpf") = "" Then Falta = True
                        If .Fields("Profissao") = "" Then Falta = True
                        If .Fields("Trabalho") = "" Then Falta = True
                        If .Fields("Fone_Trabalho") = "" Then Falta = True
                        If .Fields("Salario") = "" Then Falta = True
                        If .Fields("Refer1") = "" Then Falta = True
                        If .Fields("Refer2") = "" Then Falta = True
                        If .Fields("Refer3") = "" Then Falta = True
                    End With
                    If Falta = True Then
                        MsgBox "Cadastro Incompleto !!", vbInformation, App.Title
                        DbcDados(0).Text = ""
                        DbcDados(0).SetFocus
                        
                        Exit Sub
                    End If
                End If
                If DatDados(0).Recordset.Fields("Status") = 4 And (DbcDados(0).BoundText <> 0 And CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 5) Then
                      MsgBox "Cliente Bloqueado ! ", vbExclamation, App.Title
                      DbcDados(0).Text = ""
                      DbcDados(0).SetFocus
                      Exit Sub
                 End If
            
                If DbcDados(0).BoundText <> 0 And (CmbDados(3).ListIndex = 2 Or CmbDados(3).ListIndex = 5) Then
'                       If VarLimiteNaParcela = True And CmbDados(3).ListIndex = 2 Then
                            Set Parcelas = Banco.OpenRecordset("SELECT Entidade,  Sum(DUP.Valor) AS SomaDeValor FROM Dup WHERE  DtVenc < #" & Format(Date, "mm/dd/yyyy") & "# GROUP BY Entidade, Cond HAVING Entidade = " & DbcDados(0).BoundText & " AND Cond = FALSE")
'                       Else
'                            Set Parcelas = Banco.OpenRecordset("SELECT Entidade,  Sum(DUP.Valor) AS SomaDeValor FROM Dup WHERE  DtVenc < #" & Format(Date, "mm/dd/yyyy") & "# GROUP BY Entidade, Cond HAVING Entidade = " & DbcDados(0).BoundText & " AND Cond = FALSE")
'                       End If
                       Parcelas.Requery
                       If Parcelas.RecordCount <> 0 Then
                           LblAtraso = Format(Parcelas!SomaDeValor, " ##,##0.00")
                           
                            MsgBox "Cliente tem parcelas atrasadas, favor consultar o setor de cobrança !", vbExclamation, App.Title
                            If DatDados(0).Recordset.Fields("Status") = 5 Then
                                 MsgBox "Cliente Bloqueado ! ", vbExclamation, App.Title
                                 DbcDados(0).Text = ""
                                 DbcDados(0).SetFocus
                                 Exit Sub
                            End If
                       End If
                End If
            End If
        Set VarLista1 = Banco.OpenRecordset("SELECT HISTEST.Nome, HISTEST.CodPrd, HISTEST.Valor, HISTEST.Qtd, HISTEST.Ordem FROM HISTEST WHERE HISTEST.Nome = " & """" & DbcDados(0).BoundText & """" & " ORDER BY HISTEST.CodPrd")
        VarLista1.Requery
        
        CmbDados(3).Enabled = False
        
        If Dir(App.Path & "\ObsFixa.Txt") <> "" Then
            
            Open App.Path & "\ObsFixa.Txt" For Input As #1
            Do
                Input #1, NomeObs
                TxtObs = TxtObs + NomeObs + " - "
            Loop Until NomeObs = ""
            Close #1
ErroObs:
            If Err.Number = 62 Then
                Close #1
            End If
            
            
        End If
    End If
End If

End Sub


Private Sub DbcLoja_LostFocus()

If CmbDados(3).ListIndex = 0 Then
    
    If LerINI("Pedido", "Grade", App.Path & "\config.ini") <> "" Then
        If LerINI("Pedido", "Grade", App.Path & "\config.ini") = 1 Then
'            FormListaGrade.Show vbModal

'               FormGrade.Show vbModal
               
 '               MebDados(4).Text = ""
 '               MebDados(5).Text = ""
 '               MebDados(6).Text = ""
 '               MebVlr = ""
 '               MebDados(8).Text = ""
 '               TxtRefItens = ""
 '               CmbAcab = ""
                
                
 '               Totaliza
 '               TxtDados(2).Text = Format(CCur(TxtDados(1)) + CCur(TxtDados(3)) + CCur(TxtDados(4)), "#,##0.00")
        
 '               If VarQtdFixa <> 0 Then
 '                   MebDados(4).Text = VarQtdFixa
  '                  DbcDados(1).SetFocus
  '              Else
  '                  MebDados(4).SetFocus
  '              End If
               
               
               
        End If
    End If
        
End If



End Sub

Private Sub DBGTabela_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DBGTabela.Columns(0) = "" Then
    DatDados(9).RecordSource = "SELECT * FROM Loj WHERE Produto = 0 ORDER BY DescricaoLoja"
Else
    DatDados(9).RecordSource = "SELECT * FROM Loj WHERE Produto = " & DBGTabela.Columns(0) & " ORDER BY DescricaoLoja"
End If
DatDados(9).Refresh
End Sub

Private Sub Form_Activate()
Set Formulário = Me
    Me.Left = 50
    Me.Top = 30

If Admin = False Then
    MebDados(1) = Date
    MebDados(0) = Date
'    MebDados(1).Enabled = False
End If


'BtNovo.Enabled = False
If SSTab1.Tab = 0 Then
    DbcDados(3).SetFocus
End If

If TxtCodPed <> "" Then
    TxtCodPed.SetFocus
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyPress = 0
    End If
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

Private Sub Form_Load()
Dim Componentes As Variant

DtData = Date
ContComplemento = False

ValIndices = 0
VarContPed = 0
Dolar = 0

VarGrade = False

If VarLojaUnica = True Then
'    DBGrid3.Visible = False
End If

If LerINI("Pedido", "Grade", App.Path & "\config.ini") <> "" Then
    VarGrade = LerINI("Pedido", "Grade", App.Path & "\config.ini")
End If

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

'If Admin = False Then
'    MebDados(1) = Date
'    MebDados(1).Enabled = False
'End If

If VarCustoPed = True Then
    If Admin Then
        Frame6.Visible = True
    Else
        Frame6.Visible = False
    End If
Else
    Frame6.Visible = False
End If

If LerINI("Financeiro", "CobrarMulta", App.Path & "\config.ini") <> "" Then
    Multa = LerINI("Financeiro", "CobrarMulta", App.Path & "\config.ini")
Else
    Multa = False
End If

If LerINI("Financeiro", "LiberaPreco", App.Path & "\config.ini") <> "" Then
    VarLibreaPreco = LerINI("Financeiro", "LiberaPreco", App.Path & "\config.ini")
Else
    VarLibreaPreco = False
End If

    
'    Me.Left = (MDIPrincipal.Width - Me.Width) / 2
'    Me.Top = (MDIPrincipal.Height - Me.Height) / 6
    
'    Me.Left = 100
'    Me.Top = 50
    
    Set Tabela = Banco.OpenRecordset("Select * From PED  Order by CodPed Desc", dbOpenDynaset)
    Set Itens = Banco.OpenRecordset("Select * From Itp  Order By Pedido")
    
'    DatDados(2).RecordSource = "SELECT PED.*, ITP.*, PRD.* ,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed = 0"

'    DatDados(2).Refresh
    
    Set Indices = Banco.OpenRecordset("SELECT * FROM Ind WHERE NroIndice = 99")
    Indices.Requery
    If Indices.RecordCount > 0 Then
        Dolar = Indices!Indice
    End If
    
    PcoPzo = LerINI("Pedido", "PcoPrazo", App.Path & "\config.ini")
    Set Indices = Banco.OpenRecordset("SELECT * FROM Ind WHERE NroIndice = 3")
    If PcoPzo = 1 Then
        Indices.Requery
        If Indices.RecordCount > 0 Then
            ValIndices = Indices!Perc
        Else
            ValIndices = 0
        End If
    Else
        ValIndices = 0
    End If
    

    
    Set Entidade = Banco.OpenRecordset("Select * From Ent ")
    Set Produtos = Banco.OpenRecordset("SELECT PRD.*, GRU.*, PRD.CodPrd FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY PRD.CodPrd")
        
    DatDados(1).RecordSource = "Select CodPrd, Descrição_Produto, CodBar, Complemento, Estoque, Ativo, Marca  From PRD Where Ativo = true Order By Descrição_Produto"
                
    DatDados(5).RecordSource = "SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, PED.CodPed, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=3,'Serviço',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.DataEmiss, PED.DtFat, Ped.DtSaida, PED.Total, PED.Desconto FROM PED WHERE (((PED.Entidade)=0)) ORDER BY PED.CodPed"
    
    DatDados(10).RecordSource = "Select * From Cov Order By Nome_Conv"
    
'    DatDados(9).RecordSource = "SELECT * FROM Loj ORDER BY DescricaoLoja"
    
'    DatDados(5).Refresh
    Frame(0).Enabled = False
    Frame(1).Enabled = False
    Frame3.Enabled = True
    Op1_Click 0
    
    SSTabItens.Tab = 0
    
   ' SSTabItens.TabEnabled(1) = False
    
    VarCodPed = 0
    
    Set Acabamento = Banco.OpenRecordset("Select * From Acab Order by Descricao")
    Acabamento.Requery
    
    LIMPACampo
    MebDados(3).Text = 0
    TxtCodPed = ""
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

If BtOK.Enabled = True Then
    Cancel = True
    BtOK_Click
End If

End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Grid.Columns(6) = (Grid.Columns(6)) - CCur(Grid.Columns(6)) * (Grid.Columns(8)) / 100
    Grid.Columns(7) = CCur(Grid.Columns(6)) * CCur(Grid.Columns(5))
End Sub

Private Sub Grid_DblClick()
Dim Loja As Recordset
Dim Baixa As Boolean
Dim Prod As Recordset
Dim Resp As String
Dim SQL As String
Dim VarCont As Boolean
VarCont = True
If BtOK.Enabled = True Then
        
    If DatDados(2).Recordset.RecordCount = 0 Then Exit Sub
    Resp = MsgBox("Sim para Alterar, Não para Excluir", vbYesNoCancel + vbDefaultButton3, App.Title)
    
    If Resp = vbNo Then
        If DatDados(2).Recordset.RecordCount > 1 Then
            DatDados(2).Recordset.Delete
            If DatDados(2).Recordset.RecordCount > 1 Then DatDados(2).Recordset.MoveNext
        Else
            MsgBox "Exclui o pedido inteiro", vbInformation, App.Title
        End If
    ElseIf Resp = vbYes Then
        FormEditar.LblProduto.Caption = Grid.Columns(1)
        FormEditar.MebQtd = Grid.Columns(5)
        FormEditar.MebDesc = Grid.Columns(8)
        FormEditar.MebUnitario = Grid.Columns(6)
        FormEditar.MebFina = Grid.Columns(9)
        FormEditar.MebTot = Grid.Columns(7)
        FormEditar.txtLblItens = DatDados(2).Recordset("CodItp")
        If Not IsNull(DatDados(2).Recordset("Complemento")) Then
            FormEditar.txtComplemento = DatDados(2).Recordset("Complemento")
        End If
        If Not IsNull(DatDados(2).Recordset("Referencia")) Then
            FormEditar.TxtRef = DatDados(2).Recordset("Referencia")
        End If
        
        FormEditar.Show vbModal
        Grid.Refresh
    End If
    Totaliza
    
    MebDados(4).SetFocus

End If
End Sub


Private Sub GridPed_DblClick()
Dim Criterio As String

LIMPACampo
For Each Componentes In Frame
    Componentes.Enabled = False
    Frame3.Enabled = True
Next

BtAlterar.Enabled = True
BtCancela.Enabled = False
BtOK.Enabled = False
BtImprimir.Enabled = True
BtCupom.Enabled = True
BtNovo.Enabled = True
ImprNota1.Enabled = True

BtPedProm.Enabled = True
BtTransportadora.Enabled = True
        
If TxtCodPed = "" Then
   TxtCodPed = GridPed.Columns(2)
End If

If TxtPesqNota <> "" Then
'    If ChkNfe = 1 Then
'        Set Tabela = Banco.OpenRecordset("SELECT PED.*, NOTA.*, PED.NroNt, NOTA.SerieNota FROM PED INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed Where NroNt = " & TxtPesqNota & " And serienota = '55' ")
'    Else
        Set Tabela = Banco.OpenRecordset("Select ped.* From Ped Where NroNt = " & TxtPesqNota)
'    End If
    If Tabela.RecordCount > 0 Then
        TxtCodPed = Tabela!CodPed
    Else
        MsgBox "Nota não existente !", vbInformation, App.Title
    End If

Else
    Set Tabela = Banco.OpenRecordset("Select ped.* From Ped Where CodPed = " & TxtCodPed)
End If
Tabela.Requery
If Tabela.RecordCount = 0 Then Exit Sub

    CmbDados(3).ListIndex = Tabela("Operacao")
    If CmbDados(3).ListIndex = 0 Then
            DatDados(0).RecordSource = "Select * From Ent Where (Tipo = 2 or Tipo = 5) and ativo = true  Order by Nome"
            DatDados(0).Refresh
        Else
            DatDados(0).RecordSource = "Select * From Ent Where (Tipo = 1 or Tipo = 5 Or Tipo = 6 Or Tipo = 7) and ativo = true Order by Nome"
            DatDados(0).Refresh
'            ElseIf CmbDados(3).Text = "Transferência" Then
'                    DatDados(0).RecordSource = "Select * From ENT Order by Nome"
'                    DatDados(0).Refresh
    End If
    DbcDados(0).BoundText = Tabela("Entidade")
    Criterio = "CodEntidade = " & Tabela("Entidade")
    DatDados(0).Recordset.FindFirst Criterio
    If Not DatDados(0).Recordset.NoMatch Then
        DbcDados(0).Text = DatDados(0).Recordset("Nome")
    End If
    
    DbcDados(2).BoundText = Tabela("Vendedor")
    If Not IsNull(Tabela!Hora) Then
        MebHora.Text = Tabela!Hora
    End If
    If Not IsNull(Tabela("DataEmiss")) Then
        MebDados(1).Text = Tabela("DataEmiss")
    End If
    If Not IsNull(Tabela("DataEmiss")) Then
        MebData = Tabela("DataEmiss")
    End If
    
    If Not IsNull(Tabela("DtSaida")) Then
        MebDados(0).Text = Tabela("DtSaida")
    End If
    
    If Not IsNull(Tabela("DtFat")) Then
        MebFat = Tabela("Dtfat")
    Else
        MebFat = ""
    End If
    
    If Not IsNull(Tabela("NroNt")) Then
        If Tabela("NroNt") <> 0 Then
            MebNota = Tabela("NroNt")
        Else
            MebNota = ""
        End If
    End If
        
    If Not IsNull(Tabela("NroPedido")) Then
        txtPedidoRef.Text = Tabela("NroPedido")
    End If
    If Not IsNull(Tabela("cfo")) Then
         Txtcfop.Text = Tabela("Cfo")
    End If
    If Tabela!FunRural = True Then
        ChkFunRural.value = 1
    Else
        ChkFunRural.value = 0
    End If
    
    If Not IsNull(Tabela("Desconto")) Then
        MebTotDesc = Tabela("desconto")
    End If
    
    If Not IsNull(Tabela("Comissao")) Then
        MebDados(3) = Tabela("Comissao")
    End If
    
    If Not IsNull(Tabela("Obs")) Then
        TxtObs = Tabela("Obs")
    End If
    
    
    VarCodPed = TxtCodPed
    
    SSTab1.Tab = 1
    SSTabItens.Tab = 1
                    
        
    DatDados(2).RecordSource = "SELECT Itp.Entregue, Tributo,Icm, Descrição_Produto,CondPgto, Itp.Referencia, Itp.complemento,Ped.Entregue,DataEmiss,Ped.Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
    
'    DatDados(2).RecordSource = "SELECT PED.*, ITP.*, PRD.*,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
    
    DatDados(2).Refresh
    If DatDados(2).Recordset.RecordCount > 0 Then
        DatDados(2).Recordset.MoveFirst
        While Not DatDados(2).Recordset.Eof
            If Len(Trim(DatDados(2).Recordset("Complemento"))) > 20 Then
               ContComplemento = True
            End If
            DatDados(2).Recordset.MoveNext
        Wend
    End If
    MebDados(4).Text = ""
    MebDados(5).Text = ""
'    DbcDados(1).Text = ""
    
    Totaliza
    TxtDados(0).Text = VarCodPed
    
    DatDados(6).RecordSource = "SELECT DUP.* FROM DUP WHERE DUP.Pedido = " & TxtDados(0) & " ORDER BY DUP.NroPcl"
    DatDados(6).Refresh
    
    
'    BtTransportadora.Enabled = False
     BtParcelas.Enabled = False
    
'    CmbDados(3).Enabled = False

End Sub
Private Sub Totaliza()
    If VarCodPed <> 0 Then
        Set soma = Banco.OpenRecordset("SELECT Itp.Pedido, Sum(Itp.Vlrvdadia) AS SomaDeValor, Sum(Itp.CustoDia * Itp.Qtd) AS SomaDeCusto, Sum(Itp.Qtd) AS SomaDeQtd, Sum(Valor) AS SomaDeValorPrd FROM PED INNER JOIN Itp ON Ped.CodPed = Itp.Pedido Where Pedido = " & VarCodPed & " GROUP BY Itp.Pedido")
        
        If soma.RecordCount > 0 And Not Tabela.Eof Then
                                   
            TxtDados(1).Text = Format(soma("SomaDeValor"), "#,##0.00")
            VarTotalPrd = Format(soma("SomaDeValorPrd"), "#,##0.00")
            
            If IsNull(Tabela("ValorFrete")) Then
                TxtDados(3).Text = 0
            Else
                TxtDados(3).Text = Format(Tabela("ValorFrete"), "#,##0.00")
            End If
            TxtDados(4).Text = Format(Tabela("Ipi"), "#,##0.00")
            
            If IsNull(Tabela("PedSub")) Then
                TxtDados(6).Text = 0
            Else
                TxtDados(6).Text = Format(Tabela("PedSub"), "#,##0.00")
            End If
            
            If IsNull(Tabela("PedBaseSub")) Then
                TxtDados(5).Text = 0
            Else
                TxtDados(5).Text = Format(Tabela("PedbaseSub"), "#,##0.00")
            End If
'            If Tabela("TipoFrete") = "FOB" Then
                TxtDados(2).Text = Format(soma("SomaDeValor") + CCur(TxtDados(3)) + CCur(TxtDados(4)) + CCur(TxtDados(6)), "#,##0.00")
'            Else
'                TxtDados(2).Text = Format(soma("SomaDeValor") + CCur(TxtDados(4)) + CCur(TxtDados(6)), "#,##0.00")
'            End If
            If MebTotDesc <> "" And MebTotDesc <> "0" Then
                If TxtDados(2) <> "" Then
                    TxtDados(2) = Format(TxtDados(1) - MebTotDesc, "#,##0.00")
'                    TxtDados(1) = Format(TxtDados(2), "#,##0.00")
                End If
            End If
            
            If VarCustoPed = True Then
                MebCustoPed = Format(soma("SomaDeCusto"), "#,##0.00")
                If MebCustoPed <> 0 Then
                    MebMargemPed = ((TxtDados(2) / (soma("SomaDeCusto"))) - 1) * 100
                Else
                    MebMargemPed = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub GridPed_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DatDados(5).Recordset.RecordCount > 0 Then
        TxtDados(0).Text = DatDados(5).Recordset("CodPed")
        DatDados(4).Recordset.FindFirst "CodEntidade = " & DatDados(5).Recordset("Entidade")
        LblNomePedido = DatDados(4).Recordset("Nome")
    End If
    
End Sub

Private Sub ImprimeDupi_Click()
    funcaoImpressao 3
End Sub

Private Sub ImprNota_Click()
If GridPed.Row = -1 Then
    MsgBox "Você deve selecionar algum pedido !!!", 16
Else
    FrmNota.Show
    FrmNota.CodPed = GridPed.Columns(2)
    FrmNota.SSTabNota.Tab = 1
'    FrmNota.CodPed_KeyPress 13
End If
End Sub

Private Sub ImprNota1_Click()
    FrmNota.Show
    FrmNota.CodPed = TxtDados(0)
    FrmNota.SSTabNota.Tab = 1

End Sub

Private Sub Label3_Click()
MebDados(1).Enabled = True
End Sub

Private Sub MebAcrescimo_KeyPress(KeyAscii As Integer)

If KeyAscii = 46 Then
    SendKeys "{Backspace}"
    SendKeys ","
    Exit Sub
End If

If KeyAscii <> 43 Then
    If KeyAscii <> 44 Then
        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                If (KeyAscii < 48 Or KeyAscii > 57) Then
                    MsgBox "Caractere inválido p/ este campo", vbInformation, App.Title
                    MebAcrescimo.SetFocus
                End If
            End If
        End If
    End If
End If

'If Not IsNumeric(MebAcrescimo) Then
'    MebAcrescimo = 0
'    MebAcrescimo.SetFocus
'End If


End Sub

Private Sub MebAcrescimo_LostFocus()
Dim X As Integer
Dim y As Integer
Dim VarDesc As String
Dim TotAcresc As Currency
TotAcresc = 0
VarDesc = ""
y = 1
For X = 1 To Len(MebAcrescimo)
    If Mid$(MebAcrescimo, X, 1) = "+" Then
        TotAcresc = TotAcresc + CCur(VarDesc)
        VarDesc = ""
        y = y + 1
    Else
        VarDesc = VarDesc + Mid$(MebAcrescimo, X, 1)
    End If
Next
If VarDesc <> "" Then
     TotAcresc = TotAcresc + CCur(VarDesc)
End If

'If MebAcrescimo <> "" Then
    If MebDados(5) <> "" Then
        MebDados(5) = MebVlr + (MebVlr * (TotAcresc / 100))
    End If
'End If
End Sub

Private Sub MebComissao_LostFocus()
   If VarGrade = True Then
           If MebComissao = "" Then
                FormListaGrade.VarComiss = 0
           Else
                FormListaGrade.VarComiss = MebComissao
           End If
           FormListaGrade.Show vbModal

        Totaliza
    End If
End Sub

Private Sub MebDados_GotFocus(Index As Integer)
Dim X As Integer
'On Error Resume Next

Select Case Index
    Case 1
        MebDados(1).Text = Format(Date, "dd/mm/yyyy")
    Case 0
        MebDados(0).Text = Format(Date, "dd/mm/yyyy")
    Case 5
        For X = 1 To 10
            MebDados(5) = MebDados(5) - (MebDados(5) * (CDec(Desc(X)) / 100))
        Next
                   
        MebDados(6).Text = MebDados(5).Text
        
        Frame(0).Enabled = False
        Frame3.Enabled = True
            
        If DatDados(2).Recordset.RecordCount > 0 Then
            BtAlterar.Enabled = False
            BtCancela.Enabled = True
            BtOK.Enabled = True
            BtImprimir.Enabled = True
            BtCupom.Enabled = True
            BtPedProm.Enabled = True
            BtTransportadora.Enabled = True
            ImprNota1.Enabled = True
        End If
    
End Select
End Sub

Private Sub MebDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Select Case Index

    Case 4
    
    If KeyCode = 13 Or KeyCode = 9 Then
        If MebDados(4).Text = "" Then
            If BtOK.Enabled = False Then
                MebDados(4).Text = ""
                MebDados(4).SetFocus
            Else
                If CmbDados(3).ListIndex <> 0 Then
                    BtOK.SetFocus
                Else
                    MebDados(4).Text = ""
                    MebDados(4).SetFocus
                End If
            End If
            Exit Sub
        End If
        
        If MebDados(4).Text = "9999" Then
            MebDados(4).Text = ""
            FormListaVda.Show vbModal
            Totaliza
        End If
        
    End If
    
End Select
End Sub

Private Sub MebDados_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
    SendKeys "{Backspace}"
    SendKeys ","
    Exit Sub
End If

If KeyAscii <> 43 Then
    If KeyAscii <> 44 Then
        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                If (KeyAscii < 48 Or KeyAscii > 57) Then
                    MsgBox "Caractere inválido p/ este campo", vbInformation, App.Title
                    MebDados(Index).SetFocus
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub MebDados_LostFocus(Index As Integer)

Dim varValUnit As Currency
Dim varValTot As Currency
Dim F As Integer

Select Case Index
    Case 1
        MebDados(0).Text = MebDados(1).Text
        MebDados(3).Text = 0
    Case 2
        If MebDados(5) <> "" Then
            If MebDados(2) <> "" Then
                If CmbDados(3).ListIndex = 0 Then
                    MebDados(5) = Round(MebDados(5) - (MebDados(5) * (MebDados(2) / 100)), 4)
                Else
                    MebDados(5) = Round(MebDados(5) - (MebDados(5) * (MebDados(2) / 100)), 3)
                End If
            End If
        End If
        
        If Index = 2 Or Index = 5 Or Index = 6 Or Index = 8 Then
        
            If Not IsNumeric(MebDados(Index)) Then
                MebDados(Index) = 0
                MebDados(Index).SetFocus
            End If
        
        End If
        
    Case 3
        If DbcDados(0).BoundText = "" Then
            MsgBox "É Necessário ter selecionado algum fornecedor !!!", 16
            DbcDados(0).SetFocus
            Exit Sub
        End If
        If Not IsDate(MebDados(1)) Then
             MsgBox "Data Inválida !!!", 16
            MebDados(1).SetFocus
            Exit Sub
        End If
'        If Not IsDate(MebDados(0)) Then
'            MsgBox "Data Inválida !!!", 16
'            MebDados(0).SetFocus
'            Exit Sub
'        End If
        If TxtDados(2) = "" Then
            Tabela.AddNew
        Else
            Criterio = "CodPed = " & TxtDados(0)
            Tabela.FindFirst Criterio
            Tabela.Edit
        End If
        Tabela("Entidade") = DbcDados(0).BoundText
        
        If DbcDados(2).BoundText <> "" Then
            If IsNumeric(DbcDados(2).BoundText) Then
                Tabela("Vendedor") = DbcDados(2).BoundText
                Tabela("DataEmiss") = MebDados(1).Text
                If MebDados(0).Text <> "__/__/____" Then
                    Tabela("DtSaida") = MebDados(0).Text
                Else
                    Tabela.Fields("DtSaida") = Null  ' "__/__/____"
                End If
                Tabela("Operacao") = CmbDados(3).ListIndex
    '            tabela("Entregue") = ChkDados(0).Value
                Tabela("NroPedido") = txtPedidoRef.Text
                Tabela("Cfo") = Txtcfop.Text
'                If MebDados(3).Text <> 0 Then
                    Tabela("Comissao") = MebDados(3).Text
'                Else
'                    Tabela("Comissao") = MebComissao.Text
'                End If
                Tabela("Contato") = Left(Trim(Maquina), 10)

                Tabela.Update
                Tabela.Bookmark = Tabela.LastModified
            Else
                MsgBox "Vendedor inválido !", vbInformation, App.Title
                DbcDados(2).SetFocus
                Exit Sub
            End If
        Else
            Tabela.Update
            Exit Sub
        End If
        
'        Tabela.Requery
        Frame(1).Enabled = True
'        BtObs.Enabled = True
        If TxtDados(2) = "" Then
            VarCodPed = Tabela("CodPed")
        End If
        TxtDados(0).Text = VarCodPed
        MebDados(4).SetFocus
        SSTab1.TabEnabled(0) = False
        MebHora = Time
'        If CmbDados(3).ListIndex = 3 Then
        
'            BtCancela.Enabled = True
'            BtOk.Enabled = True
        
'        End If
        MebData = MebDados(1)
        SSTabItens.Tab = 1
        If VarQtdFixa <> 0 Then
            MebDados(4).Text = VarQtdFixa
            DbcDados(1).SetFocus
        End If
        
        
    Case 4
        If Len(MebDados(4).Text) >= 8 Then
            DbcDados(1).Text = MebDados(4).Text
            SendKeys "{TAB}"
            MebDados(4).Text = 1
        End If
    Case 5
        If MebDados(4).Text = "" Then
            MsgBox "É necessário informar a quantidade !!!", 16
            MebDados(4).SetFocus
            Exit Sub
        End If
    
        If MebDados(5).Text <> "" Then
            If CmbDados(3).ListIndex = 0 Then
                MebDados(8).Text = MebDados(4) * Round(MebDados(5), 4)
            Else
                MebDados(8).Text = MebDados(4) * Round(MebDados(5), CasasDec)
            End If
        End If
    
        If MebDados(5).Text <> "" Then
            MebDados(6).Text = MebDados(5)
        End If
        
'        If CmbDados(3).ListIndex = 0 Then
'            If mebcrescimo <> 0 Or MebAcrescimo <> "" Then
'                MebDados(5).Text = MebDados(5) + (MebDados(5) * (MebAcrescimo / 100))
'            End If
'        End If
    
        If MebDados(5).Text <> "" Then
            MebDados(6).Text = MebDados(5)
        End If
    
    Case 6
        If MebDados(6).Text < CCur(MebDados(5)) Then
            MsgBox "Financeiro não pode ser menor que o valor do produto", vbInformation, App.Title
            MebDados(6).SetFocus
            Exit Sub
        End If
        If Not IsNumeric(DbcDados(1).BoundText) Then
            MsgBox "É necessário estar selecionado algum produto !!!", 16
            DbcDados(1).SetFocus
            Exit Sub
        End If
        If Trim(MebDados(4).Text) = "" Then
            MsgBox "É necessário informar a quantidade !!!", 16
            MebDados(4).SetFocus
            Exit Sub
        End If
        If Trim(MebDados(5)) = "" Then
            MsgBox "É necessário informar o valor !!!", 16
            MebDados(5).SetFocus
            Exit Sub
        End If
               
        If Trim(MebDados(6)) = "" Then
            MsgBox "É necessário informar o valor !!!", 16
            MebDados(6).SetFocus
            Exit Sub
        End If
        
        If Trim(DbcLoja.Text) = "" Then
            MsgBox "É necessário informar uma loja valor !!!", 16
            MebDados(4).SetFocus
            Exit Sub
        End If
            If CmbDados(3).ListIndex > 0 Then
                    VarLoja = ""
                    For F = 1 To Len(DbcLoja.Text)
                        If Mid$(DbcLoja, F, 1) <> Chr(32) Then
                            VarLoja = VarLoja + Mid$(DbcLoja, F, 1)
                        Else
                            F = Len(DbcLoja.Text)
                        End If
                    Next
                    Criterio = "DescricaoLoja = " & """" & VarLoja & """" & " and Produto = " & DatDados(1).Recordset.Fields("CodPrd")
                    DatDados(7).Recordset.FindFirst Criterio
                    If DatDados(7).Recordset.NoMatch Then
                        MsgBox "Local de estoque não existe", vbInformation, App.Title
                        MebDados(4).SetFocus
                        Exit Sub
                    End If
            Else
'                If VarLojaUnica = True Then
'                    VarLoja = "LOJA"
'                Else
                    VarLoja = DbcLoja
'                End If
            End If
'        Else
'            DbcLoja.Enabled = False
'            VarLoja = "LOJA"
'        End If
        If VarGrade Then
            CmbAcab = CmbAcab & VarLoja
        End If

        If Trim(DbcDados(1).BoundText) = "" Or Not IsNumeric(DbcDados(1).BoundText) Then Exit Sub
        If DbcDados(0).BoundText = "" Or Not IsNumeric(DbcDados(0).BoundText) Then Exit Sub
        
'        Set Extra = Banco.OpenRecordset("Select * From PRD Where CodPrd = " & DbcDados(1).BoundText)
        
'        If CmbDados(3).ListIndex = 0 Then
'            Extra.Edit
'            Extra!Preco = MebVlr
'            Extra.Update
'        End If


        If Multa = True And (CmbDados(3).ListIndex = 1 Or CmbDados(3).ListIndex = 2) And VarLibreaPreco = False Then
            If CCur(MebDados(6)) < MebVlr Then
                MsgBox "Valor não pode ser menor que a tabela", vbInformation
                MebVlr.SetFocus
                Exit Sub
            End If
               
        End If
        
        Set Itens = Banco.OpenRecordset("ITP")
        
        Itens.AddNew
        Itens("Pedido") = VarCodPed
        Itens("Produto") = DbcDados(1).BoundText
        If IsNumeric(MebDados(2).Text) Then Itens("Desconto") = Round(MebDados(2).Text, CasasDec)
        If IsNumeric(MebDados(4).Text) Then Itens("Qtd") = MebDados(4).Text
        If CmbDados(3).ListIndex = 0 Or CmbDados(3).ListIndex = 11 Then
            If ContCompra = False Then
                If IsNumeric(MebDados(5).Text) Then varValUnit = Round((CCur(MebDados(5).Text)), 4) ' - CCur(MebDados(5).Text) * Round((Extra("Desconto") / 100), 3)), 4)
            Else
                If IsNumeric(MebDados(5).Text) Then varValUnit = Round((CCur(MebVlr)), 4) ' - CCur(MebVlr) * Round((Extra("Desconto") / 100), 3)), 4)
            End If
            Itens("ValorUnit") = Round(varValUnit, 4)
        Else
            If IsNumeric(MebDados(5).Text) Then varValUnit = Round((CCur(MebDados(5).Text)), CasasDec) ' - CCur(MebDados(5).Text) * Round((Extra("Desconto") / 100), 3)), CasasDec)
            Itens("ValorUnit") = Round(varValUnit, CasasDec)
        End If
        
        If IsNumeric(MebDados(4).Text) Then varValTot = CCur(varValUnit) * CCur(MebDados(4).Text)
        
        If CmbDados(3).ListIndex = 0 Or CmbDados(3).ListIndex = 11 Then
            Itens("Valor") = Round(varValTot, 4)
            If IsNumeric(MebDados(5).Text) Then Itens("CustoDia") = Round((CCur(MebDados(5).Text)), 4) ' - CCur(MebDados(5).Text) * Round((Extra("Desconto") / 100), 2)), 4)
            
'            Itens("CustoDia") = Round(VarCusto, 4)
                    
            Itens("VendaDia") = Round(Produtos("Venda"), 4)
            If ContCompra = False Then
                Itens("VlrVdaDia") = Round(CCur(MebDados(6).Text), 4) * CCur(MebDados(4).Text)
            Else
                Itens("VlrVdaDia") = Round(CCur(MebVlr), 4) * CCur(MebDados(4).Text)
            End If
        Else
            Itens("Valor") = Round(varValTot, CasasDec)
            Itens("CustoDia") = Round(VarCusto, CasasDec)
                    
            Itens("VendaDia") = Round(Produtos("Venda"), CasasDec)
            Itens("VlrVdaDia") = Round(CCur(MebDados(6).Text), CasasDec) * CCur(MebDados(4).Text)
            
            If Not IsNull(MebComissao) Then
                If IsNumeric(MebComissao) Then
                    Itens("Comissao") = Format((Round(CCur(MebDados(6).Text), CasasDec) * CCur(MebDados(4).Text)) * (MebComissao / 100), "##,##0.00")
                End If
            End If
                    
        End If
        If DbcLoja.Text <> "" Then
            Itens("Loja") = VarLoja
        End If
        Itens("Complemento") = CmbAcab
        Itens("Referencia") = TxtRefItens
'        Itens("Entregue") = -1
        Itens.Update
                      
        DatDados(2).RecordSource = "SELECT Itp.Entregue, Tributo,Icm,Descrição_Produto,CondPgto, Itp.Referencia,Itp.complemento, Ped.Entregue,DataEmiss,Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
'        Sql = "SELECT Tributo,Icm,Descrição_Produto,CondPgto, Itp.Referencia,Itp.complemento, Ped.Entregue,DataEmiss,Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
        DatDados(2).Refresh
                
        
        If DatDados(2).Recordset.RecordCount > 0 Then
            DatDados(2).Recordset.MoveLast

'           BtTransportadora.Enabled = True
'            BtParcelas.Enabled = True
            
'            BtAlterar.Enabled = False
            BtCancela.Enabled = True
            BtOK.Enabled = True
'            BtImprimir.Enabled = True
'            BtPedProm.Enabled = True
            
        End If
        MebDados(4).Text = ""
        MebDados(5).Text = ""
        MebDados(6).Text = ""
'        DbcDados(1).Text = ""
        MebVlr = ""
        MebComissao = 0
        MebDados(8).Text = ""
        TxtRefItens = ""
        CmbAcab = ""
        
        
'        Set Extra = Banco.OpenRecordset("SELECT Itp.Pedido, Sum(Itp.Vlrvdadia) AS SomaDeValor, Sum(Itp.Valor) AS SomaDeValorPrd  FROM PED INNER JOIN Itp ON Ped.CodPed = Itp.Pedido Where Pedido = " & VarCodPed & " GROUP BY Itp.Pedido")
'        TxtDados(2).Text = Format(Extra("SomaDeValor"), "#,##0.00")
'        TxtDados(1).Text = Format(Extra("SomaDeValor"), "#,##0.00")
        Totaliza
'        VarTotalPrd = Extra("SomaDeValorPrd")
        If ChkFunRural.value = 1 Then
            MebDesc = 2.3
            MebTotDesc = CCur((TxtDados(1).Text) * 2.3) / 100
            TxtDados(2) = Format(CCur(TxtDados(1)) - CCur(MebTotDesc), "#,##0.00")
        End If
        TxtDados(2).Text = Format(CCur(TxtDados(1)) + CCur(TxtDados(3)) + CCur(TxtDados(4)), "#,##0.00")

        If VarQtdFixa <> 0 Then
            MebDados(4).Text = VarQtdFixa
            DbcDados(1).SetFocus
        Else
            MebDados(4).SetFocus
        End If
        
End Select
End Sub

Private Sub MebDesc_LostFocus()
If MebDesc <> "" Then
    MebTotDesc = TxtDados(2) * (MebDesc / 100)
'    MebTotDesc.SetFocus
End If
End Sub

Private Sub MebTotDesc_LostFocus()

Totaliza

End Sub

Private Sub MebVlr_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    SendKeys "{Backspace}"
    SendKeys ","
    Exit Sub
End If

If KeyAscii <> 43 Then
    If KeyAscii <> 44 Then
        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                If (KeyAscii < 48 Or KeyAscii > 57) Then
                    MsgBox "Caractere inválido p/ este campo", vbInformation, App.Title
                    MebVlr.SetFocus
                End If
            End If
        End If
    End If
End If

If Not IsNumeric(MebVlr) Then
    MebVlr = 0
    MebVlr.SetFocus
End If


End Sub

Private Sub MebVlr_LostFocus()
    If Multa = True And (CmbDados(3).ListIndex = 1 Or CmbDados(3).ListIndex = 2) And VarLibreaPreco = False Then
        If CCur(MebVlr) <> MebDados(5) Then
            MsgBox "Valor não pode ser alterado", vbInformation
            MebVlr.SetFocus
        End If
    End If
End Sub

Private Sub Op1_Click(Index As Integer)
If Op1(0).value = True Then
    DatDados(4).RecordSource = "Select CodEntidade, Nome From Ent Where Tipo = 1  Or Tipo =  5 Or Tipo =  6 Or Tipo =  7  Order by Nome"
    DatDados(4).Refresh
Else
    DatDados(4).RecordSource = "Select CodEntidade, Nome From Ent Where Tipo = 2 Or Tipo =  5 Order by Nome"
    DatDados(4).Refresh
End If
End Sub

Private Sub ApagaPedido()
Dim Loja As Recordset
Dim Baixa As Boolean
Dim Prod As Recordset
Dim VarCont As Boolean
VarCont = True
Dim SQL As String
If MsgBox("Confirma exclusão do Pedido " & VarCodPed & " ? ", vbYesNo, App.Title) = vbYes Then
               
    If Dir(App.Path & "\" & Trim(TxtDados(0)) & ".txt") <> "" Then
        Kill App.Path & "\" & Trim(TxtDados(0)) & ".txt"
    End If
               
        
    Select Case CmbDados(3).ListIndex
       Case 0
            Baixa = False
        Case 1
            Baixa = True
        Case 2
            Baixa = True
        Case 3
            Baixa = True
        Case 6
            Baixa = True
        Case 8
            Baixa = True
        Case 9
            Baixa = False
        Case 10
            Baixa = False
        Case 11
            Baixa = True
        Case 14
            Baixa = True
        Case 15
            Baixa = False
        Case 16
            Baixa = False
    End Select
    
    Select Case CmbDados(3).ListIndex
            Case 3
                VarCont = False
            Case 4
                VarCont = False
            Case 5
                VarCont = False
            Case 7
                VarCont = False
            Case 12
                VarCont = False
            Case 13
                VarCont = False
            Case 14
'                VarCont = False
    End Select
    If VarCont = True Then
        BaixaEstoque TxtDados(0), Not Baixa, CmbDados(3).ListIndex
    End If
        
    
    SQL = "DELETE * from dup WHERE Pedido = " & VarCodPed
    Banco.Execute SQL
    
    If LerINI("PEDIDO", "Exclui", App.Path & "\Config.ini") <> "" Then
        If LerINI("PEDIDO", "Exclui", App.Path & "\Config.ini") = True Then
    
            SQL = "DELETE * from ITP WHERE ITP.Pedido = " & CCur(VarCodPed)
            Banco.Execute SQL
            
            SQL = "DELETE * from Ped WHERE CodPed = " & VarCodPed
            Banco.Execute SQL
            
        Else
                     
            SQL = "UPDATE PED SET PED.Operacao = 4, PED.DtSaida = #" & Format(Date, "mm/dd/yyyy") & "#, PED.NroPedido = '999999', PED.PrevEntrega = " & """" & CmbDados(3).Text & "CANCELADO" & """" & " WHERE PED.CodPed = " & VarCodPed
            Banco.Execute SQL
        End If
    Else
            SQL = "UPDATE PED SET PED.Operacao = 4, PED.DtSaida = #" & Format(Date, "mm/dd/yyyy") & "#, PED.NroPedido = '999999', PED.PrevEntrega = " & """" & CmbDados(3).Text & "CANCELADO" & """" & " WHERE PED.CodPed = " & VarCodPed
            Banco.Execute SQL
    
    End If
    SQL = "DELETE * from Nota WHERE CodPed = " & VarCodPed
    Banco.Execute SQL
    
    SQL = "Delete * From Rec Where CodDupli = " & VarCodPed
    Banco.Execute SQL

    Open Caminho & "\Delete.txt" For Append As #1
    Print #1, "VENDAS; Pedido; " & VarCodPed & "; " & Date & "; " & Time & "; " & Maquina & "; R$ " & Format(TxtDados(2), "#,##0.00;- #,##0.00")
    Close #1
             
    BtAlterar.Enabled = False
    BtCancela.Enabled = False
    BtOK.Enabled = False
    BtImprimir.Enabled = False
    BtCupom.Enabled = False
    ImprNota1.Enabled = False
    
    BtPedProm.Enabled = False
    SSTab1.TabEnabled(0) = True
    Frame(1).Enabled = False
    LIMPACampo
    
End If
End Sub

Public Sub LIMPACampo()
Dim X As Integer
        For Each Componente In TxtDados
            Componente.Text = ""
        Next
        For Each Componente In MebDados
            If IsDate(Componente.Text) Or Componente.Text = "__/__/____" Then
                Componente.Text = "__/__/____"
            Else
                Componente.Text = ""
            End If
        Next
        For Each Componente In DbcDados
            Componente.Text = ""
        Next
        For Each Componente In CmbDados
            Componente.Text = ""
        Next
        For Each Componente In Label
            Componente.Caption = ""
        Next
'        DbcLoja.Text = ""
        TxtDesc = ""
        For X = 1 To 10
            Desc(X) = 0
        Next
        
        DatDados(2).RecordSource = "SELECT itp.entregue, Tributo,icm,Descrição_Produto,CondPgto, Itp.Referencia, Itp.complemento, Ped.Entregue,DataEmiss,Ped.Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed = 9.9"
        DatDados(2).Refresh
                
        Frame(0).Enabled = True
        Frame3.Enabled = True
        CmbDados(3).Enabled = True

        BtAlterar.Enabled = False
        BtCancela.Enabled = False
        BtOK.Enabled = False
        BtImprimir.Enabled = False
        BtCupom.Enabled = False
        ImprNota1.Enabled = False
        
        BtPedProm.Enabled = False
        BtTransportadora.Enabled = False
        BtNovo.Enabled = False
 '       BtTransportadora.Enabled = False
        BtParcelas.Enabled = False
        TxtDados(2).Text = ""
        Frame(3).Enabled = True
'        CmbDados(3).SetFocus
        MebDesc.Text = ""
        MebTotDesc = 0
        CmbDados(3).ListIndex = 1
        TotalPed = 0
        VarOperacao = ""
        CmbAcab = ""
        TxtObs = ""
        VarTotalPrd = 0
        Txtcfop = ""
        TxtRefItens = ""
        MebComissao = ""
End Sub

Private Sub SSTab1_Click(previoustab As Integer)
If SSTab1.Tab = 2 Then
    If DBGTabela.Columns(0) = "" Then
        DatDados(8).RecordSource = "Select CodPrd, Descrição_Produto, Embalagem, Estoque, Venda, Ativo  From PRD   Where Ativo = True Order By Descrição_Produto"
        DatDados(8).Refresh
    End If
    TxtPesq.SetFocus
End If

If SSTab1.Tab = 0 Then
    If DatDados(5).Recordset.RecordCount > 0 Then
       If VarContPed <> DatDados(5).Recordset.RecordCount Then
            DatDados(5).Refresh
            DatDados(5).Recordset.MoveLast
       End If
    End If
    TxtCodPed = ""
    DbcDados(3).SetFocus
End If

End Sub

Private Sub SSTabItens_Click(previoustab As Integer)
If IsNumeric(TxtDados(0)) Then
    DatDados(6).RecordSource = "SELECT DUP.* FROM DUP WHERE DUP.Pedido = " & TxtDados(0) & " ORDER BY DUP.NroPcl"
    DatDados(6).Refresh
End If
End Sub

Private Sub TxtCodPed_GotFocus()
TxtPesqNota = ""
End Sub

Private Sub TxtCodPed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtCodPed <> "" Then
         If IsNumeric(TxtCodPed) Then
            GridPed_DblClick
         Else
            TxtCodPed = ""
            TxtCodPed.SetFocus
         End If
    End If
End If
End Sub



Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 43 Then
    If KeyAscii <> 44 Then
        If KeyAscii <> 8 Then
            If KeyAscii <> 13 Then
                If (KeyAscii < 48 Or KeyAscii > 57) Then
                    MsgBox "Caractere inválido p/ este campo", vbInformation, App.Title
                    TxtDesc.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub TxtDesc_LostFocus()
Dim X As Integer
Dim y As Integer
Dim VarDesc As String
VarDesc = ""
y = 1
For X = 1 To Len(TxtDesc)
    If Mid$(TxtDesc, X, 1) = "+" Then
        Desc(y) = VarDesc
        VarDesc = ""
        y = y + 1
    Else
        VarDesc = VarDesc + Mid$(TxtDesc, X, 1)
    End If
Next
If VarDesc <> "" Then
        Desc(y) = VarDesc
End If

End Sub

Private Sub TxtObs_GotFocus()
FrmPedidos.KeyPreview = False
End Sub

Private Sub TxtObs_LostFocus()

FrmPedidos.KeyPreview = True
SSTabItens.Tab = 0
End Sub

Private Sub TxtPesq_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If IsNumeric(TxtPesq.Text) Then
    Criterio = "CodPrd = " & TxtPesq
    DatDados(8).Recordset.FindFirst Criterio
    If DatDados(8).Recordset.NoMatch Then
       Beep
       TxtPesq.SetFocus
    Else
        DBGTabela.Refresh
    End If
Else
    If TxtPesq <> "" Then
        Criterio = "Prd.descrição_produto  Like " & """" & TxtPesq & "*"""
        DatDados(8).Recordset.FindFirst Criterio
        If DatDados(8).Recordset.NoMatch Then
           Beep
           TxtPesq.SetFocus
        Else
            DBGTabela.Refresh
        End If
    End If
End If
End Sub

Private Sub TxtPesq_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    SendKeys "{Tab}"
End If
End Sub


Private Sub TxtPesqNota_GotFocus()
TxtCodPed = ""
End Sub

Private Sub TxtPesqNota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtPesqNota <> "" Then
         If IsNumeric(TxtPesqNota) Then
            Set Tabela = Banco.OpenRecordset("Select ped.* From Ped Where NroNt = " & TxtPesqNota)
            Tabela.Requery
            If Tabela.RecordCount > 1 Then
                GridPed_DblClick
            End If
         Else
            TxtPesqNota = ""
            TxtPesqNota.SetFocus
         End If
    End If
End If

End Sub

Function funcaoImpressao(Index As Integer)

Dim X As Integer
Dim Criterio As String
Dim RegTransp As Recordset
Dim RegVend As Recordset
Dim SQL As String
Dim FormPed As Form
Dim VarSomaPcl As Currency
'On Error GoTo Erroimpr

If Index = 3 Then
    Set FormPed = FrmRelPedPromi
Else
    Set FormPed = FrmRelPed
End If

Select Case Index
    Case 0, 1, 3
        If TxtDados(0).Text <> "" Then
            Set pedidos = Banco.OpenRecordset("SELECT * From Ped Where CodPed = " & Val(TxtDados(0)))
            Set Parcelas = Banco.OpenRecordset("SELECT * From Dup Where Pedido = " & TxtDados(0))
            SQL = "SELECT Ent.cgc, Ent.InscrEst, Ent.Rg, ENT.Classificacao, Ped.Ipi, Ped.PedSub,Ent.CodEntidade,Ent.Fantasia, ENT.Nome, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor, Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento,Itp.Referencia, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, Itp.Valor, ITP.Loja, ENT.Endereço_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob , PED.ValorFrete, Ped.Desconto, ITP.CodItp, Ped.NroPedido, Prd.complemento FROM PRD INNER JOIN (ENT INNER JOIN (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON ENT.CodEntidade = PED.Entidade) ON PRD.CodPrd = ITP.Produto "
            SQL = SQL & " GROUP BY Ent.cgc, Ent.InscrEst, Ent.Rg, Ent.Classificacao, Ped.ipi, Ped.PedSub, Ent.CodEntidade,Ent.Fantasia, ENT.Nome, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor , Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Itp.Referencia, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, Itp.Valor, ITP.Loja, ENT.Endereço_Entrega, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob ,PED.ValorFrete,Ped.Desconto,PED.NroPedido,ITP.CodItp, Prd.complemento "
            SQL = SQL & " HAVING (((PED.CodPed)= " & TxtDados(0) & ")) "
            If VarOrdemPed Then
                SQL = SQL & " ORDER BY ITP.Produto"
            Else
                SQL = SQL & " ORDER BY ITP.CodItp"
            End If
            Set varPesquisa(1) = Banco.OpenRecordset(SQL)
            
            If varPesquisa(1).RecordCount <= 0 Then
                MsgBox "Pedido sem itens pra imprimir", vbInformation
                Exit Function
            End If
            If ContComplemento = False Then
                varPesquisa(1).MoveFirst
                
                While Not varPesquisa(1).Eof
                    If varPesquisa(1).Fields("Referencia") <> "" Then
                        ContComplemento = True
                    End If
                    varPesquisa(1).MoveNext
                Wend
            End If
            varPesquisa(1).MoveFirst
            Set RegTransp = Banco.OpenRecordset("Select Nome, Endereco, Nro From Ent Where CodEntidade = " & pedidos("Transportadora"))
            Set RegVend = Banco.OpenRecordset("Select Nome From Ent Where CodEntidade = " & pedidos("Vendedor"))
            
            Criterio = "CodPed = " & TxtDados(0).Text
            pedidos.FindFirst Criterio
            
            Set FormPed.Relatorio.Recordset = varPesquisa(1)
            
            FormPed.rpCabTitulo.Caption = NomeEmpresa & " - " & Cabecalho1
            FormPed.rpCabInformacao.Caption = Cabecalho2
            
                   
            If varPesquisa(1).Fields("Classificacao") = 2 Then
                If Not IsNull(varPesquisa(1).Fields("Cgc")) Then
                    FormPed.RpCgc.Caption = Format(varPesquisa(1).Fields("Cgc"), "@@.@@@.@@@/@@@@-@@")
                End If
                If Not IsNull(varPesquisa(1).Fields("InscrEst")) Then
                    FormPed.RpInscr.Caption = varPesquisa(1).Fields("InscrEst")
                End If
            Else
                If Not IsNull(varPesquisa(1).Fields("Cpf")) Then
                    FormPed.RpCgc.Caption = varPesquisa(1).Fields("Cpf")  '  DatDados(0).Recordset.Fields("Cpf")
                End If
                If Not IsNull(varPesquisa(1).Fields("Rg")) Then
                    FormPed.RpInscr.Caption = varPesquisa(1).Fields("Rg")
                End If
            End If
            If Not IsNull(pedidos("CondPgto")) Then
                FormPed.RpCondpgto.Caption = pedidos("CondPgto")
            End If
            If Not IsNull(pedidos("PrevEntrega")) Then
                FormPed.RpPrevEntrega.Caption = pedidos("PrevEntrega")
            End If
            If FormPed.Name = "FrmRelPed" Then
                If Not IsNull(pedidos("FormaPgto")) Then
                    FormPed.RpFormaPgto.Caption = pedidos("FormaPgto")
                End If
                
                If Not IsNull(pedidos("Contato")) Then
                    FormPed.RpContato.Caption = pedidos("Contato")
                End If
            End If
            
            FormPed.RpOperacao.Caption = Format(CmbDados(3).Text, ">")
            
            If Not RegTransp.Eof Then
                FormPed.Transport.Caption = RegTransp("Nome")
            End If
            
            If Not RegVend.Eof Then
                FormPed.Vendedor.Caption = RegVend("Nome")
            End If
            
            If VarPedCompleto = True Then
                FormPed.CabGrupo(1).Mostrar = True
            Else
                FormPed.CabGrupo(1).Mostrar = False
            End If
            
                        
            If CmbDados(3).ListIndex = 4 Then
                FormPed.Label21.Visible = False
                FormPed.Label32.Visible = False
                FormPed.Rod.Mostrar = False
                FormPed.RodGrupo(2).Mostrar = False
            End If
            If Not IsNull(pedidos!Obs) Then
                FormPed.Obs1.Caption = Left(pedidos("Obs"), 60)
                FormPed.Obs2.Caption = Mid(pedidos("Obs"), 61, 60)
                FormPed.Obs3.Caption = Mid(pedidos("Obs"), 122, 60)
                FormPed.Obs4.Caption = Mid(pedidos("Obs"), 182, 60)
            End If
            
            
            Set Parcelas = Banco.OpenRecordset("SELECT * From Dup Where Pedido = " & TxtDados(0) & " Order By CodDup")
            X = 0
            VarSomaPcl = 0
            Parcelas.Requery
            If Parcelas.RecordCount > 0 Then
                If Parcelas.RecordCount <= 24 Then
                    Parcelas.MoveFirst
                    While Not Parcelas.Eof
                          FormPed.Pcl(X).Caption = TxtDados(0).Text & "/" & Format(Parcelas.Fields("NroPcl"), "00")
                          FormPed.Venc(X).Caption = Parcelas.Fields("DtVenc")
                          FormPed.Valor(X).Caption = Format(Parcelas.Fields("Valor"), "#,##0.00;(#,##0.00)")
                          VarSomaPcl = VarSomaPcl + Format(Parcelas.Fields("Valor"), "#,##0.00;(#,##0.00)")
                          X = X + 1
                          Parcelas.MoveNext
                    Wend
                    Parcelas.MoveFirst
                Else
                    MsgBox "Numero de parcelas maior que o suportado pelo pedido", vbInformation, App.Title
                End If
            Else
                MsgBox "Não foi gerado financeiro para este pedido", vbInformation, App.Title
            End If
            
            If Index <> 3 Then
                If VarFinanceiro = True Then
                    If Parcelas.RecordCount > 0 Then
                        FormPed.RodGrupo(2).Mostrar = True
                    Else
                        FormPed.RodGrupo(2).Mostrar = False
                    End If
                Else
                    FormPed.RodGrupo(2).Mostrar = False
                End If
            End If
        
            If Index = 1 Or Index = 3 Then
                If Index = 1 Then
                    For Each Componentes In FormPed.Shape
                        Componentes.Visible = False
                    Next
                
                    For Each Componentes In FormPed.Pcl
                        Componentes.Mostrar = False
                    Next
                
                    For Each Componentes In FormPed.Valor
                        Componentes.Mostrar = False
                    Next
                    
                    For Each Componentes In FormPed.Venc
                        Componentes.Mostrar = False
                    Next
                    
                    For Each Componentes In FormPed.lblInfGrupo
                        Componentes.Caption = ""
                    Next
                
                    For Each Componentes In FormPed.rpGrupo
                        Componentes.Mostrar = False
                    Next
                    
                    FormPed.LabelCondicao.Visible = False
                
                    FormPed.rpDetalhes(0).Campo = ""
                    FormPed.rpDetalhes(0).Caption = ""
                    FormPed.rpDetalhes(0).Formula = False
                    FormPed.lblGrupo(0).Caption = ""
                    
                    FormPed.rpDetalhes(1).Campo = "Loja"
                    FormPed.rpDetalhes(1).Caption = "Loja"
                    FormPed.rpDetalhes(1).Formula = False
                    FormPed.lblGrupo(1).Caption = "Local"
                    
                    FormPed.rpDetalhes(2).Campo = "forEntregue"
                    FormPed.rpDetalhes(2).Formula = True
                    FormPed.lblGrupo(2).Caption = "Entregue"
                
                    FormPed.rpDetalhes(3).Campo = "Qtd"
                    FormPed.rpDetalhes(3).Formula = False
                    FormPed.lblGrupo(3).Caption = "Qtd."
                End If
            End If
            
                        
            If Index = 3 Then
                If Parcelas.RecordCount > 0 Then
                    FormPed.SecPromissoria.Mostrar = True
                    DatDados(5).Recordset.MoveLast
                    FormPed.NroPedido1.Caption = Format(DatDados(5).Recordset("CodPed"), "000000")

                    FormPed.QuantiaExtenso.Caption = UCase(PassaExtenso(CStr(VarSomaPcl))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                    FormPed.ValorPromissoria.Caption = Format(VarSomaPcl, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
                    If VarPromissoriaBca = False Then
                        FormPed.ExtensoDia.Caption = valorextenso(Day(Parcelas("DtVenc"))) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(Parcelas("DtVenc")))) & " DO ANO DE " & valorextenso(Year(Parcelas("DtVenc")))
                        FormPed.DiaVenc.Caption = Day(Parcelas("DtVenc"))
                        FormPed.MesVenc.Caption = UCase(RetornaMes(Month(Parcelas("DtVenc"))))
                        FormPed.AnoVenc.Caption = Year(Parcelas("DtVenc"))
                        FormPed.Pagavel.Caption = CidadeEmpresa
                        FormPed.NomeEmpresa.Caption = NomeInscr
                        FormPed.EmpresaCGCCPF.Caption = Cnpj
                        FormPed.RpDia.Caption = Day(Date)
                        FormPed.RpMes.Caption = UCase(RetornaMes(Month(Date)))
                        FormPed.RpAno.Caption = Year(Date)
                    
                    End If
                    FormPed.Emitente.Caption = varPesquisa(1).Fields("Nome")
                    FormPed.Emitente2.Caption = varPesquisa(1).Fields("Nome")
                    If Not IsNull(varPesquisa(1).Fields("Cpf")) Then
                        FormPed.EmitenteCpf.Caption = varPesquisa(1).Fields("CPF")
                    End If

                    FormPed.EmitenteEndereco.Caption = varPesquisa(1).Fields("Endereco") & " - " & varPesquisa(1).Fields("Nro")
                    FormPed.EmitenteCidade.Caption = varPesquisa(1).Fields("Cidade") & "-" & varPesquisa(1).Fields("UF")
                Else
                    MsgBox "Não foi gerado financeiro para este pedido", vbInformation, App.Title
                End If
            End If
'            If NomeEmpresa = "SERRARIA SÃO JOSÉ" Then
'                FormPed.Label46.Visible = True
'                FormPed.Label48.Visible = True
'                FormPed.Shape4.Visible = True
'            Else
'                FormPed.Label46.Visible = False
'                FormPed.Label48.Visible = False
'                FormPed.Shape4.Visible = False
'            End If

'            If Len(Trim(varPesquisa(1).Fields("Itp.Complemento"))) <= 15 Then
            If ContComplemento = False Then
                FormPed.Det.Height = 4
                FormPed.Line2.y1 = 190
                FormPed.Line2.y2 = 190
                FormPed.ReportField29.Mostrar = False
            Else
                FormPed.Det.Height = 7
                FormPed.Line2.y1 = 390
                FormPed.Line2.y2 = 390
'                FormPed.ReportField15.Mostrar = False
                FormPed.ReportField29.Mostrar = True
            End If

            If Index = 0 Or Index = 1 Then
                If LerINI("PEDIDO", "Chopp", App.Path & "\Config.ini") <> "" Then
                     If LerINI("PEDIDO", "Chopp", App.Path & "\Config.ini") = 1 Then
                         FormPed.RodGrupo(2).Mostrar = True
                         FormPed.Image1.Visible = True
                         FormPed.Shape(0).Visible = False
                         FormPed.Shape(2).Visible = False
                         FormPed.Shape(3).Visible = False
                         FormPed.Shape(4).Visible = False
                     Else
                         FormPed.Image1.Visible = False
                     End If
                Else
                     FormPed.Image1.Visible = False
                End If
            End If
                    
            FormPed.Config

        Else
            MsgBox "Nenhum Pedido Selecionado !", vbExclamation, App.Title
        End If
    Case 4
    
        If TxtDados(0).Text <> "" Then
                  
         Dim RegPed As Recordset
         Dim sql2 As String
         Dim SQL3 As String
         If DatDados(5).Recordset("Entregue") = 0 And DatDados(5).Recordset("CondPgto") = "APP" Then
         
             SQL3 = "UPDATE ITP SET ITP.Entregue = True WHERE ITP.Pedido = " & DatDados(5).Recordset.Fields("CodPed")
             Banco.Execute SQL3
             
             DatDados(5).Recordset.Edit
             DatDados(5).Recordset("Entregue") = 1
             DatDados(5).Recordset.Update
         
                      
             BaixaEstoque DatDados(5).Recordset.Fields("CodPed"), True, DatDados(5).Recordset.Fields("Operacao")
             
'             If Dir("C:\Program Files\PlugBot\8088\csv-data\plugpedidos\produto.csv") <> "" Then
'                Kill "C:\Program Files\PlugBot\8088\csv-data\plugpedidos\produto.csv"
'             End If

            
         End If
         
         sql2 = "SELECT Ped.Ipi, Ped.PedSub,Ent.CodEntidade, ENT.Nome, Ent.Fantasia, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor, Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob , PED.ValorFrete, Ped.Desconto, ITP.CodItp, Ped.NroPedido, Prd.complemento FROM PRD INNER JOIN (ENT INNER JOIN (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON ENT.CodEntidade = PED.Entidade) ON PRD.CodPrd = ITP.Produto "
         sql2 = sql2 & " GROUP BY Ped.ipi, Ped.PedSub, Ent.CodEntidade, ENT.Nome, Ent.Fantasia, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor , Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob ,PED.ValorFrete,Ped.Desconto,PED.NroPedido,ITP.CodItp, Prd.complemento "
         sql2 = sql2 & " HAVING (((PED.CodPed)= " & TxtDados(0) & ")) ORDER BY ITP.CodItp"
         Set varPesquisa(1) = Banco.OpenRecordset(sql2)
         
                         
         
         Set FrmRelPed2.Relatorio.Recordset = varPesquisa(1)
         
         
         FrmRelPed2.rpCabTitulo.Caption = Cabecalho1
         FrmRelPed2.rpCabInformacao.Caption = Cabecalho2
        
'         If DbGrid.Columns(5) = "7" Then
'            FrmRelPed.ReportField10.Caption = "ORÇAMENTO"
'         End If
'         If DbGrid.Columns(5) = "9" Then
'            FrmRelPed.ReportField10.Caption = "DEVOLUÇÃO"
'         End If

         

             FrmRelPed2.CabGrupo(1).Mostrar = True

            Me.Hide
            FrmRelPed2.Relatorio.Copias = 1
    
            FrmRelPed2.Config
        Else
            MsgBox "Nenhum Pedido Selecionado !", vbExclamation, App.Title
        End If
        
End Select
           

Exit Function
Erroimpr:
     
     MsgBox Err.Description & Err.Number, vbCritical, Caminho


End Function

Private Sub TxtRefItens_LostFocus()


If TxtRefItens <> "" Then
    ContComplemento = True
End If
End Sub
