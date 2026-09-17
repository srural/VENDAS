VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormProdutos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8460
   ClientLeft      =   2100
   ClientTop       =   3150
   ClientWidth     =   14130
   Icon            =   "FormProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14130
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   14565
      TabIndex        =   205
      Top             =   8100
      Width           =   14595
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   14160
      TabIndex        =   192
      Top             =   0
      Width           =   14190
      Begin VB.Label LblTitulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   7245
         TabIndex        =   194
         Top             =   90
         Width           =   6585
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   193
         Top             =   45
         Width           =   5820
      End
   End
   Begin TabDlg.SSTab Sstab1 
      Height          =   7530
      Left            =   0
      TabIndex        =   49
      Top             =   585
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   13282
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   9
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "FormProdutos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PctFundo(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PctBarra"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdDuplica"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Fornecedores"
      TabPicture(1)   =   "FormProdutos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDados(10)"
      Tab(1).Control(1)=   "LblDados(16)"
      Tab(1).Control(2)=   "LblDados(17)"
      Tab(1).Control(3)=   "LblDados(18)"
      Tab(1).Control(4)=   "DbgDados(0)"
      Tab(1).Control(5)=   "MebData"
      Tab(1).Control(6)=   "MebCusto"
      Tab(1).Control(7)=   "DbcFornec"
      Tab(1).Control(8)=   "DatDados(3)"
      Tab(1).Control(9)=   "TxtCond"
      Tab(1).Control(10)=   "CmdDados(0)"
      Tab(1).Control(11)=   "CmdDados(1)"
      Tab(1).Control(12)=   "CmdDados(2)"
      Tab(1).Control(13)=   "CmdAtual"
      Tab(1).Control(14)=   "DatDados(2)"
      Tab(1).Control(15)=   "PctFundo(1)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Con&sulta"
      TabPicture(2)   =   "FormProdutos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "DBGrid2"
      Tab(2).Control(2)=   "TxtPesquisa"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(4)=   "Frame7"
      Tab(2).Control(5)=   "ChkFaltas"
      Tab(2).Control(6)=   "Frame16"
      Tab(2).Control(7)=   "Frame17"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "&Composição"
      TabPicture(3)   =   "FormProdutos.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ChkDados(1)"
      Tab(3).Control(1)=   "DatDados(6)"
      Tab(3).Control(2)=   "ChkAtuVda"
      Tab(3).Control(3)=   "CmdComp"
      Tab(3).Control(4)=   "MebDados(18)"
      Tab(3).Control(5)=   "MebTotComp"
      Tab(3).Control(6)=   "CmdAtuComp"
      Tab(3).Control(7)=   "CmbLanca"
      Tab(3).Control(8)=   "DatDados(7)"
      Tab(3).Control(9)=   "DbgDados(1)"
      Tab(3).Control(10)=   "DbcProdComp"
      Tab(3).Control(11)=   "MebQtd"
      Tab(3).Control(12)=   "MebDados(19)"
      Tab(3).Control(13)=   "MebDados(24)"
      Tab(3).Control(14)=   "Label12"
      Tab(3).Control(15)=   "Label11"
      Tab(3).Control(16)=   "Label10"
      Tab(3).Control(17)=   "LblDados(30)"
      Tab(3).Control(18)=   "LblDados(29)"
      Tab(3).ControlCount=   19
      TabCaption(4)   =   "&Movimento"
      TabPicture(4)   =   "FormProdutos.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DatDados(16)"
      Tab(4).Control(1)=   "DatDados(15)"
      Tab(4).Control(2)=   "DatDados(14)"
      Tab(4).Control(3)=   "DatDados(13)"
      Tab(4).Control(4)=   "DatDados(12)"
      Tab(4).Control(5)=   "GridEst2"
      Tab(4).Control(6)=   "CmdAtualiza"
      Tab(4).Control(7)=   "List2"
      Tab(4).Control(8)=   "List1"
      Tab(4).Control(9)=   "DatDados(9)"
      Tab(4).Control(10)=   "DatDados(8)"
      Tab(4).Control(11)=   "DBGridManual"
      Tab(4).Control(12)=   "DtDataMov"
      Tab(4).Control(13)=   "Lbl2"
      Tab(4).Control(14)=   "lbl1"
      Tab(4).Control(15)=   "LblSaida"
      Tab(4).Control(16)=   "LblEntrada"
      Tab(4).Control(17)=   "LblDados(35)"
      Tab(4).Control(18)=   "Label17"
      Tab(4).ControlCount=   19
      TabCaption(5)   =   "&Grade"
      TabPicture(5)   =   "FormProdutos.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DatDados(11)"
      Tab(5).Control(1)=   "DatDados(10)"
      Tab(5).Control(2)=   "DbgGrade"
      Tab(5).Control(3)=   "DbgIndices"
      Tab(5).Control(4)=   "Web"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "&Foto"
      TabPicture(6)   =   "FormProdutos.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "LblFoto"
      Tab(6).Control(1)=   "ImgFoto"
      Tab(6).Control(2)=   "CmdFicha"
      Tab(6).Control(3)=   "DatDadosCest"
      Tab(6).ControlCount=   4
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   16
         Left            =   -67035
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6255
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Data DatDadosCest 
         Caption         =   "Pesquisa"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -62460
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT CEST.Descricao, CEST.Segmento, CEST.CEST, CEST.[NCM/SH] FROM CEST ORDER BY CEST.Descricao"
         Top             =   4635
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   15
         Left            =   -66855
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5715
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   14
         Left            =   -66945
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4455
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   13
         Left            =   -66945
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4815
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   12
         Left            =   -66900
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5085
         Visible         =   0   'False
         Width           =   1950
      End
      Begin TabDlg.SSTab GridEst2 
         Height          =   3885
         Left            =   -74820
         TabIndex        =   206
         Top             =   450
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   6853
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   6
         TabsPerRow      =   8
         TabHeight       =   520
         TabCaption(0)   =   "Loja 1"
         TabPicture(0)   =   "FormProdutos.frx":00D0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "GridEstat"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Loja 2"
         TabPicture(1)   =   "FormProdutos.frx":00EC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "GridEstat2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Loja 3"
         TabPicture(2)   =   "FormProdutos.frx":0108
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "DBGrid4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Loja 4"
         TabPicture(3)   =   "FormProdutos.frx":0124
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "DBGrid5(0)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Loja 5"
         TabPicture(4)   =   "FormProdutos.frx":0140
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "DBGrid5(1)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Loja 6"
         TabPicture(5)   =   "FormProdutos.frx":015C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "DBGrid5(2)"
         Tab(5).ControlCount=   1
         Begin MSDBGrid.DBGrid GridEstat 
            Bindings        =   "FormProdutos.frx":0178
            Height          =   3435
            Left            =   135
            OleObjectBlob   =   "FormProdutos.frx":0192
            TabIndex        =   207
            Top             =   90
            Width           =   13530
         End
         Begin MSDBGrid.DBGrid GridEstat2 
            Bindings        =   "FormProdutos.frx":1A77
            Height          =   3435
            Left            =   -74865
            OleObjectBlob   =   "FormProdutos.frx":1A92
            TabIndex        =   208
            Top             =   90
            Width           =   13530
         End
         Begin MSDBGrid.DBGrid DBGrid4 
            Bindings        =   "FormProdutos.frx":3378
            Height          =   3435
            Left            =   -74865
            OleObjectBlob   =   "FormProdutos.frx":3393
            TabIndex        =   210
            Top             =   90
            Width           =   13530
         End
         Begin MSDBGrid.DBGrid DBGrid5 
            Bindings        =   "FormProdutos.frx":4C76
            Height          =   3435
            Index           =   0
            Left            =   -74865
            OleObjectBlob   =   "FormProdutos.frx":4C91
            TabIndex        =   214
            Top             =   90
            Width           =   13530
         End
         Begin MSDBGrid.DBGrid DBGrid5 
            Bindings        =   "FormProdutos.frx":6577
            Height          =   3435
            Index           =   1
            Left            =   -74865
            OleObjectBlob   =   "FormProdutos.frx":6592
            TabIndex        =   215
            Top             =   90
            Width           =   13530
         End
         Begin MSDBGrid.DBGrid DBGrid5 
            Bindings        =   "FormProdutos.frx":7E78
            Height          =   3435
            Index           =   2
            Left            =   -74865
            OleObjectBlob   =   "FormProdutos.frx":7E93
            TabIndex        =   222
            Top             =   90
            Width           =   13530
         End
      End
      Begin VB.Frame Frame19 
         Height          =   555
         Left            =   8505
         TabIndex        =   201
         Top             =   5715
         Width           =   2445
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Est_Ideal"
            Height          =   315
            Index           =   3
            Left            =   885
            TabIndex        =   202
            Tag             =   "S"
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indexador"
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   203
            Top             =   225
            Width           =   705
         End
      End
      Begin VB.CommandButton CmdFicha 
         Caption         =   "Ficha Técnica"
         Height          =   735
         Left            =   -62265
         TabIndex        =   191
         Top             =   585
         Width           =   1095
      End
      Begin VB.Frame Frame17 
         Caption         =   "Instruções Situação Tributaria para Simples Nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -68655
         TabIndex        =   189
         Top             =   4995
         Width           =   7575
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1995
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   190
            Text            =   "FormProdutos.frx":9779
            Top             =   225
            Width           =   7215
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Instruções Situação Tributaria para Regime Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -74775
         TabIndex        =   187
         Top             =   4995
         Width           =   5955
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1860
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   188
            Text            =   "FormProdutos.frx":9A34
            Top             =   225
            Width           =   5775
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Loja"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   11
         Left            =   -65730
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * FROM CLoj  ORDER BY Loja"
         Top             =   3105
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Data DatDados 
         Caption         =   "GRADE"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas_2_0\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   10
         Left            =   -72840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LOJ"
         Top             =   3195
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.CheckBox ChkDados 
         Caption         =   "Baixar &Estoque da Composição"
         DataField       =   "BaixarComp"
         Height          =   285
         Index           =   1
         Left            =   -67440
         TabIndex        =   181
         Tag             =   "S"
         Top             =   6795
         Width           =   2535
      End
      Begin VB.CommandButton CmdAtualiza 
         Caption         =   "Atualiza"
         Height          =   330
         Left            =   -72885
         TabIndex        =   119
         Top             =   4545
         Width           =   825
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3150
         ItemData        =   "FormProdutos.frx":9BEF
         Left            =   -62265
         List            =   "FormProdutos.frx":9BF1
         TabIndex        =   116
         Top             =   4335
         Width           =   960
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3150
         ItemData        =   "FormProdutos.frx":9BF3
         Left            =   -64830
         List            =   "FormProdutos.frx":9C27
         TabIndex        =   115
         Top             =   4335
         Width           =   2490
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   9
         Left            =   -72210
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6255
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\DADOS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   8
         Left            =   -66810
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data DatDados 
         Caption         =   "Composicao"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistemas em Desenvolvimento\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   6
         Left            =   -67665
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3060
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.CheckBox ChkAtuVda 
         Caption         =   "Atualiza Preço Venda"
         Height          =   345
         Left            =   -69555
         TabIndex        =   112
         Top             =   6750
         Width           =   1935
      End
      Begin VB.CommandButton CmdComp 
         Caption         =   "&Imprimi Composição"
         Height          =   330
         Left            =   -63330
         TabIndex        =   111
         Top             =   675
         Width           =   1875
      End
      Begin MSMask.MaskEdBox MebDados 
         DataField       =   "MaoObra"
         Height          =   330
         Index           =   18
         Left            =   -74685
         TabIndex        =   105
         Top             =   6795
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdDuplica 
         Caption         =   "Duplicar Registro"
         Height          =   360
         Left            =   9360
         TabIndex        =   104
         Top             =   585
         Width           =   1500
      End
      Begin VB.PictureBox PctBarra 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   10665
         ScaleHeight     =   525
         ScaleWidth      =   3135
         TabIndex        =   100
         Top             =   1590
         Width           =   3165
      End
      Begin VB.Frame Frame5 
         Caption         =   "Preço"
         Enabled         =   0   'False
         Height          =   690
         Left            =   11025
         TabIndex        =   98
         Top             =   405
         Width           =   2730
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Venda"
            Height          =   420
            Index           =   2
            Left            =   90
            TabIndex        =   99
            Top             =   180
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   741
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
      End
      Begin MSMask.MaskEdBox MebTotComp 
         Height          =   330
         Left            =   -62895
         TabIndex        =   97
         Top             =   6480
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.000;(#,##0.000)"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdAtuComp 
         Caption         =   "&Atualiza Custo"
         Height          =   330
         Left            =   -62895
         TabIndex        =   96
         Top             =   6885
         Width           =   1725
      End
      Begin VB.CommandButton CmbLanca 
         Caption         =   "&Lançar"
         Height          =   330
         Left            =   -66765
         TabIndex        =   92
         Top             =   675
         Width           =   1005
      End
      Begin VB.Data DatDados 
         Caption         =   "Produtos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   7
         Left            =   -73515
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT CodPrd, Descrição_Produto  FROM Prd Order By Descrição_Produto"
         Top             =   675
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   6360
         Index           =   0
         Left            =   135
         ScaleHeight     =   6360
         ScaleWidth      =   13875
         TabIndex        =   50
         Top             =   1125
         Width           =   13875
         Begin VB.Frame Frame20 
            Caption         =   "Preço a Atacado"
            ForeColor       =   &H8000000D&
            Height          =   1155
            Left            =   6075
            TabIndex        =   211
            Top             =   4545
            Width           =   2205
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Margem_Concorrencia"
               Height          =   315
               Index           =   27
               Left            =   135
               TabIndex        =   212
               Tag             =   "S"
               Top             =   540
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   -2147483641
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Atacado"
               Height          =   195
               Index           =   56
               Left            =   135
               TabIndex        =   213
               Top             =   315
               Width           =   1005
            End
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "Chefe_Grupo"
            Height          =   315
            Index           =   14
            Left            =   9090
            MaxLength       =   15
            TabIndex        =   11
            Tag             =   "S"
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton BtBarra 
            Height          =   330
            Left            =   12600
            Picture         =   "FormProdutos.frx":9DCE
            Style           =   1  'Graphical
            TabIndex        =   204
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Frame Frame18 
            Height          =   555
            Left            =   8370
            TabIndex        =   198
            Top             =   5175
            Width           =   2445
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Venda_Concorrencia"
               Height          =   315
               Index           =   25
               Left            =   855
               TabIndex        =   199
               Tag             =   "S"
               Top             =   180
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   -2147483641
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gelado"
               Height          =   195
               Left            =   180
               TabIndex        =   200
               Top             =   225
               Width           =   510
            End
         End
         Begin VB.CheckBox ChkDados 
            Caption         =   "&Site"
            DataField       =   "Concorrencia"
            Height          =   195
            Index           =   6
            Left            =   11310
            TabIndex        =   195
            Tag             =   "S"
            Top             =   105
            Width           =   1410
         End
         Begin VB.Frame Frame9 
            Caption         =   "Observação"
            Height          =   1725
            Left            =   8370
            TabIndex        =   182
            Top             =   2835
            Width           =   2430
            Begin VB.TextBox txtDados 
               BorderStyle     =   0  'None
               DataField       =   "Obs"
               Height          =   1455
               Index           =   4
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   183
               Tag             =   "S"
               Top             =   180
               Width           =   2265
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Dados de Tabela Vinculada"
            ForeColor       =   &H8000000D&
            Height          =   600
            Left            =   90
            TabIndex        =   177
            Top             =   5715
            Width           =   8190
            Begin VB.Label LblLab 
               Appearance      =   0  'Flat
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   5265
               TabIndex        =   180
               Top             =   225
               Width           =   2280
            End
            Begin VB.Label LblApres 
               Appearance      =   0  'Flat
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   180
               TabIndex        =   179
               Top             =   360
               Width           =   7380
            End
            Begin VB.Label LblDes 
               Appearance      =   0  'Flat
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   180
               TabIndex        =   178
               Top             =   180
               Width           =   4980
            End
         End
         Begin VB.Frame Frame14 
            ForeColor       =   &H8000000D&
            Height          =   600
            Left            =   8370
            TabIndex        =   176
            Top             =   5715
            Width           =   2475
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Balança"
               DataField       =   "Chefe"
               Height          =   285
               Index           =   4
               Left            =   180
               TabIndex        =   45
               Tag             =   "S"
               Top             =   225
               Width           =   960
            End
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Pesado"
               DataField       =   "Troca"
               Height          =   285
               Index           =   5
               Left            =   1350
               TabIndex        =   46
               Tag             =   "S"
               Top             =   225
               Width           =   1050
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Parametros Fiscais"
            ForeColor       =   &H8000000D&
            Height          =   1005
            Left            =   5985
            TabIndex        =   173
            Top             =   3465
            Width           =   2295
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               DataField       =   "Fator"
               Height          =   285
               Index           =   11
               Left            =   840
               MaxLength       =   10
               TabIndex        =   44
               Tag             =   "S"
               Top             =   600
               Width           =   1275
            End
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               DataField       =   "CodAbc"
               Height          =   315
               Index           =   10
               Left            =   840
               MaxLength       =   10
               TabIndex        =   43
               Tag             =   "S"
               Top             =   270
               Width           =   1275
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "cBenef"
               Height          =   195
               Index           =   43
               Left            =   270
               TabIndex        =   175
               Top             =   270
               Width           =   510
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "cClassTrib"
               Height          =   195
               Index           =   42
               Left            =   90
               TabIndex        =   174
               Top             =   630
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Estoque"
            ForeColor       =   &H8000000D&
            Height          =   1020
            Left            =   5985
            TabIndex        =   170
            Top             =   1440
            Width           =   2295
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Mínimo"
               Height          =   315
               Index           =   1
               Left            =   1155
               TabIndex        =   39
               Tag             =   "S"
               ToolTipText     =   "Se o estoque mínimo estiver zero, ele não será controlado"
               Top             =   270
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Validade"
               Height          =   315
               Index           =   20
               Left            =   1155
               TabIndex        =   40
               Tag             =   "S"
               Top             =   615
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Validade"
               Height          =   195
               Index           =   9
               Left            =   435
               TabIndex        =   172
               Top             =   615
               Width           =   615
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estoque &Mín."
               Height          =   195
               Index           =   6
               Left            =   90
               TabIndex        =   171
               Top             =   270
               Width           =   960
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Peso"
            ForeColor       =   &H8000000D&
            Height          =   945
            Left            =   5985
            TabIndex        =   167
            Top             =   2475
            Width           =   2295
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "PesoBruto"
               Height          =   315
               Index           =   28
               Left            =   810
               TabIndex        =   41
               Tag             =   "S"
               ToolTipText     =   "Se o estoque mínimo estiver zero, ele não será controlado"
               Top             =   180
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "PesoLiquido"
               Height          =   315
               Index           =   29
               Left            =   795
               TabIndex        =   42
               Tag             =   "S"
               Top             =   540
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Bruto"
               Height          =   195
               Index           =   37
               Left            =   300
               TabIndex        =   169
               Top             =   225
               Width           =   375
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Liquido"
               Height          =   195
               Index           =   36
               Left            =   150
               TabIndex        =   168
               Top             =   585
               Width           =   510
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Estatisticas"
            ForeColor       =   &H8000000D&
            Height          =   1380
            Left            =   8370
            TabIndex        =   158
            Top             =   1440
            Width           =   2445
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "DtAlt"
               Height          =   315
               Index           =   10
               Left            =   105
               TabIndex        =   159
               Tag             =   "S"
               Top             =   420
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Preco"
               Height          =   315
               Index           =   5
               Left            =   1230
               TabIndex        =   160
               Tag             =   "S"
               Top             =   420
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "DtVenda"
               Height          =   315
               Index           =   8
               Left            =   1230
               TabIndex        =   161
               Tag             =   "S"
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "DtCompra"
               Height          =   315
               Index           =   6
               Left            =   105
               TabIndex        =   162
               Tag             =   "S"
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Ultima Venda"
               Height          =   195
               Index           =   15
               Left            =   1185
               TabIndex        =   166
               Top             =   750
               Width           =   1035
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Ultima Compra"
               Height          =   195
               Index           =   22
               Left            =   15
               TabIndex        =   165
               Top             =   750
               Width           =   1110
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data Preço"
               Height          =   195
               Left            =   105
               TabIndex        =   164
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Compra"
               Height          =   195
               Left            =   1230
               TabIndex        =   163
               Top             =   210
               Width           =   630
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Preço Desconto"
            ForeColor       =   &H8000000D&
            Height          =   1605
            Left            =   4455
            TabIndex        =   155
            Top             =   4095
            Width           =   1485
            Begin VB.CheckBox ChkDados 
               Caption         =   "c/ &Desconto"
               DataField       =   "Promocao"
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   36
               Tag             =   "S"
               Top             =   240
               Width           =   1200
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Margem_Promocao"
               Height          =   315
               Index           =   16
               Left            =   90
               TabIndex        =   37
               Tag             =   "S"
               Top             =   450
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.0;- #,##0.0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Venda_Promocao"
               Height          =   315
               Index           =   17
               Left            =   105
               TabIndex        =   38
               Tag             =   "S"
               Top             =   990
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   -2147483641
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Desc."
               Height          =   195
               Index           =   23
               Left            =   105
               TabIndex        =   157
               Top             =   810
               Width           =   825
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   34
               Left            =   1170
               TabIndex        =   156
               Top             =   495
               Width           =   225
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Preço a Prazo"
            ForeColor       =   &H8000000D&
            Height          =   1605
            Left            =   2880
            TabIndex        =   152
            Top             =   4095
            Width           =   1485
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "MargemPrazo"
               Height          =   315
               Index           =   14
               Left            =   90
               TabIndex        =   34
               Tag             =   "S"
               Top             =   285
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.0;- #,##0.0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "PcoPrazo"
               Height          =   315
               Index           =   15
               Left            =   90
               TabIndex        =   35
               Tag             =   "S"
               Top             =   990
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   -2147483641
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   4
               Left            =   1155
               TabIndex        =   154
               Top             =   300
               Width           =   180
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Prazo"
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   153
               Top             =   765
               Width           =   810
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Preços"
            ForeColor       =   &H8000000D&
            Height          =   2505
            Left            =   2880
            TabIndex        =   144
            Top             =   1440
            Width           =   3045
            Begin VB.CommandButton Btcusto 
               Caption         =   ">"
               Height          =   285
               Left            =   2610
               TabIndex        =   145
               ToolTipText     =   "Atualiza para custo real"
               Top             =   225
               Width           =   285
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Margem"
               Height          =   315
               Index           =   22
               Left            =   1200
               TabIndex        =   28
               Tag             =   "S"
               Top             =   900
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.0;- #,##0.0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Venda"
               Height          =   345
               Index           =   23
               Left            =   1200
               TabIndex        =   30
               Tag             =   "S"
               Top             =   1290
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   -2147483641
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Custo"
               Height          =   315
               Index           =   21
               Left            =   1200
               TabIndex        =   27
               Tag             =   "S"
               Top             =   555
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebPerc 
               DataField       =   "Margem"
               Height          =   315
               Left            =   2175
               TabIndex        =   29
               Top             =   900
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.0;- #,##0.0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "CustoMedio"
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   33
               Tag             =   "S"
               Top             =   2070
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox CustoReal 
               Height          =   315
               Left            =   1200
               TabIndex        =   31
               Top             =   1665
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "CustoPrd"
               Height          =   315
               Index           =   34
               Left            =   1200
               TabIndex        =   26
               Tag             =   "S"
               Top             =   210
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.000;- #,##0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebcustoIva 
               DataField       =   "Margem"
               Height          =   315
               Left            =   2295
               TabIndex        =   32
               Top             =   1665
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.0;- #,##0.0"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Margem"
               Height          =   195
               Left            =   570
               TabIndex        =   151
               Top             =   930
               Width           =   570
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Venda"
               Height          =   195
               Index           =   14
               Left            =   675
               TabIndex        =   150
               Top             =   1320
               Width           =   465
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C&usto"
               Height          =   195
               Index           =   5
               Left            =   720
               TabIndex        =   149
               Top             =   585
               Width           =   405
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Custo Médio"
               Height          =   195
               Index           =   0
               Left            =   285
               TabIndex        =   148
               Top             =   2115
               Width           =   885
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Custo Real"
               Height          =   195
               Index           =   1
               Left            =   345
               TabIndex        =   147
               Top             =   1725
               Width           =   780
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Custo &Compra"
               Height          =   195
               Index           =   45
               Left            =   135
               TabIndex        =   146
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Dados Fiscais"
            ForeColor       =   &H8000000D&
            Height          =   4215
            Left            =   90
            TabIndex        =   127
            Top             =   1440
            Width           =   2715
            Begin VB.CommandButton BtCest 
               Height          =   315
               Left            =   2340
               Picture         =   "FormProdutos.frx":A210
               Style           =   1  'Graphical
               TabIndex        =   219
               ToolTipText     =   "Vincula com tabela de preços"
               Top             =   2745
               Width           =   330
            End
            Begin VB.CommandButton BtNcm 
               Height          =   315
               Left            =   2340
               Picture         =   "FormProdutos.frx":A742
               Style           =   1  'Graphical
               TabIndex        =   218
               ToolTipText     =   "Vincula com tabela de preços"
               Top             =   2385
               Width           =   330
            End
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               DataField       =   "ClasseFiscal"
               Height          =   315
               Index           =   12
               Left            =   1035
               MaxLength       =   15
               TabIndex        =   22
               Tag             =   "S"
               Top             =   2745
               Width           =   1290
            End
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               DataField       =   "Classe"
               Height          =   315
               Index           =   7
               Left            =   1035
               MaxLength       =   15
               TabIndex        =   21
               Tag             =   "S"
               ToolTipText     =   "Se nao souber o numero interio coloca só os 2 primeiros digitos"
               Top             =   2385
               Width           =   1290
            End
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               DataField       =   "SitTrib"
               Height          =   315
               Index           =   8
               Left            =   1035
               MaxLength       =   10
               TabIndex        =   23
               Tag             =   "S"
               ToolTipText     =   "Ver Instruções de preenchimento na guia de consulta"
               Top             =   3105
               Width           =   1290
            End
            Begin VB.ComboBox CmbDados 
               Appearance      =   0  'Flat
               DataField       =   "Aliquota"
               Height          =   315
               Index           =   0
               ItemData        =   "FormProdutos.frx":AC74
               Left            =   1035
               List            =   "FormProdutos.frx":AC8A
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Tag             =   "SI"
               Top             =   585
               Width           =   1560
            End
            Begin VB.ComboBox CmbDados 
               Appearance      =   0  'Flat
               DataField       =   "Tributo"
               Height          =   315
               Index           =   1
               ItemData        =   "FormProdutos.frx":ACBF
               Left            =   1035
               List            =   "FormProdutos.frx":ACCF
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Tag             =   "SI"
               Top             =   225
               Width           =   1560
            End
            Begin VB.TextBox txtDados 
               Appearance      =   0  'Flat
               DataField       =   "Iva"
               Height          =   315
               Index           =   9
               Left            =   1035
               MaxLength       =   10
               TabIndex        =   25
               Tag             =   "S"
               Top             =   3825
               Width           =   1290
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Icm"
               Height          =   315
               Index           =   7
               Left            =   1035
               TabIndex        =   17
               Tag             =   "S"
               Top             =   945
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Ipi"
               Height          =   315
               Index           =   9
               Left            =   1035
               TabIndex        =   18
               Tag             =   "S"
               Top             =   1305
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "CfopPrd"
               Height          =   315
               Index           =   33
               Left            =   1035
               TabIndex        =   24
               Tag             =   "S"
               Top             =   3465
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               PromptInclude   =   0   'False
               Format          =   "@.@@@"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "FretePrd"
               Height          =   315
               Index           =   32
               Left            =   1035
               TabIndex        =   19
               Tag             =   "S"
               Top             =   1665
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Redutor"
               Height          =   315
               Index           =   31
               Left            =   1035
               TabIndex        =   20
               Tag             =   "S"
               Top             =   2025
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;- #,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   53
               Left            =   2400
               TabIndex        =   143
               Top             =   3870
               Width           =   180
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   52
               Left            =   2400
               TabIndex        =   142
               Top             =   1710
               Width           =   180
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   51
               Left            =   2400
               TabIndex        =   141
               Top             =   2025
               Width           =   180
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   46
               Left            =   2400
               TabIndex        =   140
               Top             =   990
               Width           =   180
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   47
               Left            =   2400
               TabIndex        =   139
               Top             =   1350
               Width           =   180
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frete"
               Height          =   195
               Index           =   33
               Left            =   585
               TabIndex        =   138
               Top             =   1681
               Width           =   360
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ICM/Entrada"
               Height          =   195
               Index           =   48
               Left            =   30
               TabIndex        =   137
               Top             =   2045
               Width           =   915
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CEST"
               Height          =   195
               Index           =   50
               Left            =   480
               TabIndex        =   136
               Top             =   2775
               Width           =   420
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "T&ipo Tributo"
               Height          =   195
               Index           =   8
               Left            =   90
               TabIndex        =   135
               Top             =   225
               Width           =   855
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Aliquota"
               Height          =   195
               Index           =   12
               Left            =   375
               TabIndex        =   134
               Top             =   589
               Width           =   570
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ICM/Saída"
               Height          =   195
               Index           =   24
               Left            =   150
               TabIndex        =   133
               Top             =   953
               Width           =   795
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "I.P.I."
               Height          =   195
               Index           =   25
               Left            =   615
               TabIndex        =   132
               Top             =   1317
               Width           =   330
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N.C.M."
               Height          =   195
               Index           =   2
               Left            =   450
               TabIndex        =   131
               Top             =   2409
               Width           =   495
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Sit.Tributária"
               Height          =   195
               Index           =   32
               Left            =   60
               TabIndex        =   130
               Top             =   3137
               Width           =   885
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C.F.O.P."
               Height          =   195
               Index           =   44
               Left            =   345
               TabIndex        =   129
               Top             =   3501
               Width           =   600
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "I.V.A."
               Height          =   195
               Index           =   41
               Left            =   525
               TabIndex        =   128
               Top             =   3870
               Width           =   390
            End
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "Marca"
            Height          =   315
            Index           =   13
            Left            =   4290
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "S"
            Top             =   750
            Width           =   1305
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "Complemento"
            Height          =   315
            Index           =   3
            Left            =   915
            MaxLength       =   100
            TabIndex        =   5
            Tag             =   "S"
            Top             =   390
            Width           =   6345
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   10935
            ScaleHeight     =   465
            ScaleWidth      =   2730
            TabIndex        =   102
            Top             =   5805
            Width           =   2760
            Begin VB.Label LblEstoque 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1260
               TabIndex        =   216
               Top             =   90
               Width           =   1185
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estoque Atual"
               Height          =   195
               Index           =   7
               Left            =   30
               TabIndex        =   103
               Top             =   150
               Width           =   990
            End
         End
         Begin VB.CommandButton CmdBarra 
            Caption         =   "&Exclui Barra"
            Height          =   330
            Left            =   11205
            TabIndex        =   101
            Top             =   1080
            Width           =   1110
         End
         Begin VB.Data DatDados 
            Caption         =   "Loja"
            Connect         =   "Access"
            DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   5
            Left            =   11430
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   4815
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.CheckBox ChkDados 
            Caption         =   "&Ativo"
            DataField       =   "Ativo"
            Height          =   195
            Index           =   2
            Left            =   10350
            TabIndex        =   4
            Tag             =   "S"
            Top             =   105
            Width           =   735
         End
         Begin VB.TextBox TxtCodGru 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   4590
            TabIndex        =   13
            Top             =   1095
            Width           =   990
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "CodBar"
            Height          =   315
            Index           =   2
            Left            =   7380
            MaxLength       =   15
            TabIndex        =   10
            Tag             =   "S"
            Top             =   720
            Width           =   1620
         End
         Begin VB.Data DatDados 
            Caption         =   "CodBar"
            Connect         =   "Access"
            DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   4
            Left            =   11250
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2340
            Visible         =   0   'False
            Width           =   1800
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "FormProdutos.frx":AD03
            Height          =   1845
            Left            =   10935
            OleObjectBlob   =   "FormProdutos.frx":AD1D
            TabIndex        =   47
            Top             =   1530
            Width           =   2760
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "Local"
            Height          =   315
            Index           =   6
            Left            =   2220
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "S"
            Top             =   750
            Width           =   1440
         End
         Begin VB.CheckBox ChkDados 
            Caption         =   "&Foto"
            DataField       =   "Foto"
            Height          =   195
            Index           =   0
            Left            =   9405
            TabIndex        =   3
            Tag             =   "S"
            Top             =   105
            Width           =   690
         End
         Begin VB.Data DatDados 
            Caption         =   "Grupos"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   0
            Left            =   1530
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT GRU.Descrição_Grupo, GRU.CodGru  FROM GRU ORDER BY GRU.Descrição_Grupo"
            Top             =   1095
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Embalagem"
            Height          =   315
            Index           =   5
            Left            =   915
            MaxLength       =   5
            TabIndex        =   6
            Tag             =   "S"
            Top             =   750
            Width           =   780
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Descrição_Produto"
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
            Index           =   1
            Left            =   3090
            MaxLength       =   40
            TabIndex        =   2
            Tag             =   "S"
            Top             =   45
            Width           =   5715
         End
         Begin VB.TextBox txtDados 
            Appearance      =   0  'Flat
            DataField       =   "CodPrd"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   1
            Tag             =   "S"
            Top             =   45
            Width           =   1290
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FormProdutos.frx":B714
            DataField       =   "Grupo"
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   12
            Tag             =   "SI"
            Top             =   1095
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12648447
            ListField       =   "Descrição_Grupo"
            BoundColumn     =   "CodGru"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FormProdutos.frx":B72E
            DataField       =   "Fabricante"
            Height          =   315
            Index           =   1
            Left            =   6765
            TabIndex        =   14
            Tag             =   "SI"
            Top             =   1095
            Width           =   4350
            _ExtentX        =   7673
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "FormProdutos.frx":B748
            Height          =   2235
            Left            =   10935
            OleObjectBlob   =   "FormProdutos.frx":B762
            TabIndex        =   48
            Top             =   3450
            Width           =   2760
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Comissao"
            Height          =   315
            Index           =   30
            Left            =   6405
            TabIndex        =   9
            Tag             =   "S"
            Top             =   765
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   -2147483641
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0;- #,##0.0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Estoque"
            Height          =   315
            Index           =   4
            Left            =   12420
            TabIndex        =   217
            Tag             =   "S"
            Top             =   0
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0;- #,##0.0"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Chefe Grupo"
            Height          =   195
            Index           =   54
            Left            =   9090
            TabIndex        =   209
            Top             =   495
            Width           =   900
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
            Height          =   195
            Index           =   49
            Left            =   3705
            TabIndex        =   126
            Top             =   795
            Width           =   495
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   40
            Left            =   7065
            TabIndex        =   125
            Top             =   765
            Width           =   180
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comi&ssão"
            Height          =   195
            Index           =   39
            Left            =   5655
            TabIndex        =   124
            Top             =   795
            Width           =   675
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Compl."
            Height          =   195
            Index           =   38
            Left            =   390
            TabIndex        =   123
            Top             =   465
            Width           =   480
         End
         Begin VB.Label QtdCorpo 
            Height          =   285
            Left            =   8850
            TabIndex        =   122
            Top             =   60
            Width           =   300
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "&Local"
            Height          =   195
            Index           =   31
            Left            =   1740
            TabIndex        =   95
            Top             =   795
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fab./Fornec."
            Height          =   195
            Left            =   5745
            TabIndex        =   88
            Top             =   1170
            Width           =   930
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "CodFab/CodBarra"
            Height          =   195
            Index           =   26
            Left            =   7380
            TabIndex        =   67
            Top             =   495
            Width           =   1290
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Em&balagem"
            Height          =   195
            Index           =   11
            Left            =   -75
            TabIndex        =   54
            Top             =   795
            Width           =   945
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "D&escrição "
            Height          =   195
            Index           =   1
            Left            =   2295
            TabIndex        =   53
            Top             =   75
            Width           =   765
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Có&digo"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   52
            Top             =   45
            Width           =   495
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "&Grupo"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   51
            Top             =   1125
            Width           =   435
         End
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   -74820
         ScaleHeight     =   735
         ScaleWidth      =   5160
         TabIndex        =   80
         Top             =   540
         Width           =   5160
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Custo"
            Height          =   360
            Index           =   11
            Left            =   45
            TabIndex        =   85
            Tag             =   "N"
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483644
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Margem"
            Height          =   360
            Index           =   12
            Left            =   2025
            TabIndex        =   86
            Tag             =   "N"
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483644
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Venda"
            Height          =   360
            Index           =   13
            Left            =   3150
            TabIndex        =   87
            Tag             =   "N"
            Top             =   270
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483644
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Venda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   3150
            TabIndex        =   84
            Top             =   45
            Width           =   555
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Margem (%)"
            Height          =   195
            Index           =   27
            Left            =   2025
            TabIndex        =   83
            Top             =   45
            Width           =   825
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Custo"
            Height          =   195
            Index           =   19
            Left            =   45
            TabIndex        =   82
            Top             =   45
            Width           =   405
         End
         Begin VB.Label LblDados 
            DataField       =   "Nome"
            Height          =   375
            Index           =   20
            Left            =   240
            TabIndex        =   81
            Top             =   2640
            Width           =   3135
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Cotação"
         Connect         =   "Access"
         DatabaseName    =   "caminho & ""\Produtos.mdb"""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   2
         Left            =   -69600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CmdAtual 
         Caption         =   "&Atualiza Preços"
         Height          =   375
         Left            =   -67035
         TabIndex        =   79
         Top             =   810
         Width           =   1620
      End
      Begin VB.CheckBox ChkFaltas 
         Caption         =   "Só as Faltas"
         Height          =   285
         Left            =   -63615
         TabIndex        =   78
         Top             =   720
         Width           =   1275
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ordem"
         Height          =   645
         Left            =   -66090
         TabIndex        =   75
         Top             =   450
         Width           =   2310
         Begin VB.OptionButton Opt 
            Caption         =   "&Grupos"
            Height          =   285
            Index           =   3
            Left            =   1260
            TabIndex        =   77
            Top             =   225
            Width           =   915
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Produtos"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   76
            Top             =   225
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Calcular"
         Height          =   645
         Left            =   -68340
         TabIndex        =   72
         Top             =   450
         Width           =   2085
         Begin VB.OptionButton Opt 
            Caption         =   "&Venda"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   74
            Top             =   225
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Margem"
            Height          =   285
            Index           =   0
            Left            =   1035
            TabIndex        =   73
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.TextBox TxtPesquisa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74820
         TabIndex        =   69
         Top             =   675
         Width           =   6195
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pesquisas"
         Height          =   645
         Left            =   210
         TabIndex        =   68
         Top             =   390
         Width           =   9030
         Begin VB.Data DatImporta 
            Caption         =   "Importa"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   7065
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   135
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox TxtBusca 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7110
            TabIndex        =   196
            Text            =   "Importar Produto"
            Top             =   225
            Width           =   1815
         End
         Begin VB.Data DatDados 
            Caption         =   "Produtos"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   1
            Left            =   1170
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   225
            Visible         =   0   'False
            Width           =   2025
         End
         Begin MSDBCtls.DBCombo DbcPesquisa 
            Bindings        =   "FormProdutos.frx":C151
            Height          =   360
            Left            =   105
            TabIndex        =   0
            Top             =   225
            Width           =   6840
            _ExtentX        =   12065
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ListField       =   "descrição_Produto"
            BoundColumn     =   "CodPrd"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton CmdDados 
         Cancel          =   -1  'True
         Caption         =   "&Limpar"
         Height          =   375
         Index           =   2
         Left            =   -63075
         TabIndex        =   66
         Top             =   810
         Width           =   1620
      End
      Begin VB.CommandButton CmdDados 
         Caption         =   "&Excluir Fornecedor"
         Height          =   375
         Index           =   1
         Left            =   -65055
         TabIndex        =   64
         Top             =   810
         Width           =   1620
      End
      Begin VB.CommandButton CmdDados 
         Caption         =   "&Lançar"
         Height          =   375
         Index           =   0
         Left            =   -63480
         TabIndex        =   59
         Top             =   6930
         Width           =   1080
      End
      Begin VB.TextBox TxtCond 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -67080
         MaxLength       =   40
         TabIndex        =   58
         Tag             =   "N"
         Top             =   6975
         Width           =   2925
      End
      Begin VB.Data DatDados 
         Caption         =   "Fornecedor"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Index           =   3
         Left            =   -73695
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
         Top             =   6930
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSDBCtls.DBCombo DbcFornec 
         Bindings        =   "FormProdutos.frx":C16B
         DataField       =   "Nome"
         Height          =   315
         Left            =   -74730
         TabIndex        =   55
         Tag             =   "N"
         Top             =   6975
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MebCusto 
         Height          =   315
         Left            =   -70320
         TabIndex        =   56
         Tag             =   "N"
         Top             =   6975
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   315
         Left            =   -68610
         TabIndex        =   57
         Tag             =   "N"
         Top             =   6975
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBGrid.DBGrid DbgDados 
         Bindings        =   "FormProdutos.frx":C185
         Height          =   5340
         Index           =   0
         Left            =   -74775
         OleObjectBlob   =   "FormProdutos.frx":C19F
         TabIndex        =   65
         Tag             =   "N"
         Top             =   1305
         Width           =   13650
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "FormProdutos.frx":D0A2
         Height          =   3495
         Left            =   -74820
         OleObjectBlob   =   "FormProdutos.frx":D0BC
         TabIndex        =   71
         Top             =   1305
         Width           =   13785
      End
      Begin MSDBGrid.DBGrid DbgDados 
         Bindings        =   "FormProdutos.frx":E4AF
         Height          =   5160
         Index           =   1
         Left            =   -74730
         OleObjectBlob   =   "FormProdutos.frx":E4C9
         TabIndex        =   89
         Tag             =   "N"
         Top             =   1125
         Width           =   13605
      End
      Begin MSDBCtls.DBCombo DbcProdComp 
         Bindings        =   "FormProdutos.frx":F3C8
         DataField       =   "Nome"
         Height          =   315
         Left            =   -74730
         TabIndex        =   90
         Tag             =   "N"
         Top             =   675
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Descrição_Produto"
         BoundColumn     =   "CodPrd"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MebQtd 
         Height          =   315
         Left            =   -68520
         TabIndex        =   91
         Tag             =   "N"
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Format          =   "#,##0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDados 
         DataField       =   "Perdas"
         Height          =   330
         Index           =   19
         Left            =   -72975
         TabIndex        =   107
         Top             =   6795
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDados 
         DataField       =   "OutrosCusto"
         Height          =   330
         Index           =   24
         Left            =   -71265
         TabIndex        =   109
         Top             =   6795
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSDBGrid.DBGrid DBGridManual 
         Bindings        =   "FormProdutos.frx":F3E2
         Height          =   2130
         Left            =   -74820
         OleObjectBlob   =   "FormProdutos.frx":F3FC
         TabIndex        =   113
         Top             =   5220
         Width           =   7365
      End
      Begin MSComCtl2.DTPicker DtDataMov 
         Height          =   315
         Left            =   -74370
         TabIndex        =   118
         Top             =   4545
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   41222145
         CurrentDate     =   36759
      End
      Begin MSDBGrid.DBGrid DbgGrade 
         Bindings        =   "FormProdutos.frx":10474
         Height          =   6240
         Left            =   -74775
         OleObjectBlob   =   "FormProdutos.frx":1048E
         TabIndex        =   185
         Tag             =   "N"
         Top             =   630
         Width           =   6675
      End
      Begin MSDBGrid.DBGrid DbgIndices 
         Bindings        =   "FormProdutos.frx":11376
         Height          =   6270
         Left            =   -67665
         OleObjectBlob   =   "FormProdutos.frx":11391
         TabIndex        =   186
         Top             =   630
         Width           =   6135
      End
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   5730
         Left            =   -67665
         TabIndex        =   197
         Top             =   1080
         Width           =   5775
         ExtentX         =   10186
         ExtentY         =   10107
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Label Lbl2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   -66765
         TabIndex        =   221
         Top             =   6840
         Width           =   1500
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   -66765
         TabIndex        =   220
         Top             =   6480
         Width           =   1500
      End
      Begin VB.Image ImgFoto 
         Height          =   6585
         Left            =   -74505
         ToolTipText     =   "Coloque dentro da pasta fotos um imagem.jpg (500x350) com o nome igual ao codigo do produto"
         Top             =   540
         Width           =   12030
      End
      Begin VB.Label LblFoto 
         AutoSize        =   -1  'True
         Caption         =   "NÃO TEM FOTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1350
         Left            =   -71655
         TabIndex        =   184
         Top             =   3180
         Width           =   8175
      End
      Begin VB.Label LblSaida 
         BackStyle       =   0  'Transparent
         Caption         =   "Total das Saidas:"
         Height          =   195
         Left            =   -69510
         TabIndex        =   121
         Top             =   4770
         Width           =   2040
      End
      Begin VB.Label LblEntrada 
         BackStyle       =   0  'Transparent
         Caption         =   "Total das Entradas:"
         Height          =   285
         Left            =   -69510
         TabIndex        =   120
         Top             =   4455
         Width           =   2040
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         Height          =   195
         Index           =   35
         Left            =   -74820
         TabIndex        =   117
         Top             =   4590
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimentação feita por entrada e saída manual"
         Height          =   195
         Left            =   -74820
         TabIndex        =   114
         Top             =   4950
         Width           =   3390
      End
      Begin VB.Label Label12 
         Caption         =   "&Outros Custos (%)"
         Height          =   240
         Left            =   -71265
         TabIndex        =   110
         Top             =   6570
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "&Perdas (%)"
         Height          =   240
         Left            =   -72975
         TabIndex        =   108
         Top             =   6570
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "&Mão de Obra (%)"
         Height          =   240
         Left            =   -74685
         TabIndex        =   106
         Top             =   6570
         Width           =   1320
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Quantidade"
         Height          =   195
         Index           =   30
         Left            =   -68520
         TabIndex        =   94
         Top             =   450
         Width           =   825
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Produtos"
         Height          =   195
         Index           =   29
         Left            =   -74730
         TabIndex        =   93
         Top             =   450
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "&Pesquisa"
         Height          =   240
         Left            =   -74820
         TabIndex        =   70
         Top             =   450
         Width           =   870
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Con&dição Pagamento"
         Height          =   195
         Index           =   18
         Left            =   -67080
         TabIndex        =   63
         Top             =   6750
         Width           =   1530
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         Height          =   195
         Index           =   17
         Left            =   -68610
         TabIndex        =   62
         Top             =   6750
         Width           =   345
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Fornecedor"
         Height          =   195
         Index           =   16
         Left            =   -74730
         TabIndex        =   61
         Top             =   6750
         Width           =   810
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "C&usto"
         Height          =   195
         Index           =   10
         Left            =   -70320
         TabIndex        =   60
         Top             =   6750
         Width           =   405
      End
   End
End
Attribute VB_Name = "FormProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Dim FlagAbertura As Boolean
Dim Produto As Double
Public ValorAnt As Boolean
Public FlagAltera As Boolean

Public CTextBox As Boolean      'Caixa de Texto
Public CMaskEditBox As Boolean  'Mascara
Public CComboBox As Boolean     'Combo
Public CListBox As Boolean      'Lista
Public CdbCombo As Boolean      'DBCombo
Public CDbList As Boolean       'DBLista
Public CCheckBox As Boolean     'Caixa de Checagem
Public CDbgrid As Boolean       'Grade
Public ComandBotom As Boolean   'Botão
Public PesquisaDb As Boolean   'Pesquisa

Public SelectFromSQL  As String  ' Representa o Conteudo das Cláusulas Select e From do Sql Utilizado p/ acessar os registros deste formulário
Public WhereSQL As String        ' Representa o Conteudo da Cláusula Where do Sql Utilizado p/ acessar os registros deste formulário
Public FiltroSQL As String       ' Critério imposto pelo usuário
Public OrderBySQL As String      ' Representa a cláusula Order By do Sql do formulário Ativo
Public PesquisaSql As String     ' Representa a pesquisa F2
Public SelectFromFiltro As String  ' Representa a relação de campos que sera exibida na visualização no formulário de filtragem

Public Tabela As Recordset
Dim REG As Recordset
Public VarCodigo As Double
Public VarRegime  As Byte
Public VarPreco As Currency
Public VarEstoque As String

Public LocalizacaoBasica As String ' Armazena o nome do campo que sera utilizado na ferramenta de localizaco basica

Private Sub DefineObjetos()
'Define quais objeto foram utilizados neste formulário

CTextBox = True
CMaskEditBox = True
CComboBox = True
CListBox = False
CdbCombo = True
CDbList = False
CCheckBox = True
CDbgrid = False
PesquisaDb = True
End Sub


Private Sub BtBarra_Click()

Dim Prod As Recordset

    Open VarImprBarra For Output As #1
    If IsNumeric(TxtDados(0)) Then
        Print #1, Chr(2) + "O0000"
        Print #1, Chr(2) + "M0300"
        Print #1, Chr(2) + "c0000"
        Print #1, Chr(2) + "f000"
        Print #1, Chr(2) + "e"
        Print #1, Chr(2) + "LC0000"
        Print #1, "H09"
        Print #1, "D11"
        Print #1, "SC"
        Print #1, "PC"
        Print #1, "R0000"
        Print #1, "z"
        Print #1, "W"
        Print #1, Chr(94) + "01"
        Print #1, "191100500800020" & psCompString(3, TxtDados(1), 40)
        Print #1, "1F1202000200020" & psCompString(1, DBGrid1.Columns(1), 13)  ' Codigo Barras
'       Print #1, "192300500000125" & "R$"
        Print #1, "192300400100160" & "R$"
        If ChkDados(3).value = True Then
            Print #1, "142300000000210" & Alinhar(MebDados(16), 6, 2) ' Valor do Produto
        Else
            Print #1, "142300000000210" & Alinhar(MebDados(23), 6, 2) ' Valor do Produto
        End If
        Print #1, "191100100500020" & psCompString(3, NomeEmpresa, 20)
        Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
        Print #1, "E"
        
    End If
    Close #1
End Sub

Private Sub BtCest_Click()
FrmPesquisaCest.Show vbModal
End Sub

Private Sub Btcusto_Click()
MebDados(21) = CustoReal
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtNcm_Click()
FrmPesquisaNcm.Show vbModal
End Sub

Private Sub BtTab_Click()
FrmPesquisaTab.Show vbModal
End Sub

Private Sub ChkFaltas_Click()
If OPt(2).value = True Then
    Opt_Click 2
ElseIf OPt(3).value = True Then
    Opt_Click 3
End If
End Sub

Private Sub CmbDados_LostFocus(Index As Integer)

If Index = 1 Then
    If CmbDados(1).ListIndex = 0 Then
        If VarRegime = 1 Then
            TxtDados(8).Text = "102"
        Else
            TxtDados(8).Text = "000"
        End If
        MebDados(33).Text = "5102"
    End If

    If CmbDados(1).ListIndex = 1 Then
        If VarRegime = 1 Then
            TxtDados(8).Text = "500"
        Else
            TxtDados(8).Text = "060"
        End If
        MebDados(33).Text = "5405"
    End If
    
    If CmbDados(1).ListIndex = 2 Then
        If VarRegime = 1 Then
            TxtDados(8).Text = "400"
        Else
            TxtDados(8).Text = "040"
        End If
        MebDados(33).Text = "5102"
    End If
    

End If

If Index = 0 Then
    If CmbDados(0).Text = "7 %" Then
        MebDados(7) = 7
    ElseIf CmbDados(0).Text = "12 %" Then
            MebDados(7) = 12
        ElseIf CmbDados(0).Text = "18 %" Then
                MebDados(7) = 18
            ElseIf CmbDados(0).Text = "25 %" Then
                    MebDados(7) = 25
                ElseIf CmbDados(0).Text = "2,16 %" Then
                    MebDados(7).Text = 2.16
    End If
End If
End Sub

Private Sub CmbLanca_Click()
Dim Criterio As String
Dim SQL As String
If IsNumeric(DbcProdComp.BoundText) = False Then
    MsgBox "Selecione um produto", vbCritical, App.Title
    DbcProdComp.SetFocus
    Exit Sub
End If
If MebQtd = "" Or Not IsNumeric(MebQtd) Then
    MsgBox "Digite uma quantidade válida", vbCritical, App.Title
    MebQtd.SetFocus
    Exit Sub
End If
Criterio = "Produto = " & Registros.Fields("CodPrd") & " and PrdComp = " & DbcProdComp.BoundText
DatDados(6).Recordset.FindFirst Criterio
If DatDados(6).Recordset.NoMatch Then
    SQL = "INSERT INTO Comp ( Produto, PrdComp, Qtd ) Values ("
    SQL = SQL & Formulário.Registros.Fields("CodPrd") & ","
    SQL = SQL & DbcProdComp.BoundText & " , "
    SQL = SQL & Numero4(MebQtd) & ")"
'    SQL = SQL & """" & TxtDados(1) & """" & ")"
    Banco.Execute SQL
    DatDados(6).Refresh
    
    AtualizaComp
Else
    MsgBox "Produto já faz parte desta composição", vbInformation, App.Title
End If
DbcProdComp.BoundText = ""
MebQtd = ""
DbcProdComp.SetFocus

End Sub

Private Sub CmdAtual_Click()
Registros.Edit

Registros.Fields("Custo") = DbgDados(0).Columns(3)
Registros.Fields("Venda") = Registros.Fields("custo") + (Registros.Fields("custo") * (Registros.Fields("Margem") / 100))
Registros.Fields("DtAlt") = Date
Registros.Update
MostraDados
PosMovimentacao

End Sub

Private Sub CmdAtualiza_Click()
SSTab1_Click 4
End Sub

Private Sub CmdBarra_Click()
'On Error Resume Next
DatDados(4).Recordset.Delete

End Sub

Private Sub CmdAtuComp_Click()
Dim VarOutrosCustos As Currency
Dim VarMaoObra As Currency
Dim VarPerdas As Currency
VarOutrosCustos = 0
VarMaoObra = 0
VarPerdas = 0
If MebTotComp <> "" Then
    If MsgBox("Confirma atualização de custos ?", vbYesNo, App.Title) = vbYes Then
        VarOutrosCustos = MebTotComp * (MebDados(24) / 100)
        VarMaoObra = MebTotComp * (MebDados(18) / 100)
        VarPerdas = MebTotComp * (MebDados(19) / 100)
        Registros.Edit
        Registros.Fields("Custo") = CCur(MebTotComp) + VarMaoObra + VarPerdas + VarOutrosCustos
        Registros.Update
        If ChkAtuVda.value = True Then
            Registros.Edit
            Registros.Fields("Venda") = Registros.Fields("Custo") + (Registros.Fields("Custo") * (Registros.Fields("Margem") / 100))
            Registros.Update
        Else
            Registros.Edit
            Registros.Fields("Margem") = ((Registros.Fields("Venda") / Registros.Fields("Custo") - 1) * 100)
            Registros.Update
        End If
        MostraDados
        MsgBox "Custo atualizado !", vbInformation, App.Title
    End If
End If
End Sub

Private Sub CmdComp_Click()
Dim SQL As String
If DatDados(6).Recordset.RecordCount > 0 Then

        SQL = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Custo, PRD.Venda, PRD.Estoque, COMP.Qtd, COMP.PrdComp, ProdutoComp.Descrição_Produto, ProdutoComp.Estoque FROM PRD AS ProdutoComp INNER JOIN (PRD INNER JOIN COMP ON PRD.CodPrd = COMP.Produto) ON ProdutoComp.CodPrd = COMP.PrdComp Where Prd.Ativo = True and Prd.CodPrd = " & TxtDados(0).Text & " ORDER BY PRD.Descrição_Produto"

        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelComposicao.Relatorio.Recordset = varPesquisa(1)
        FrmRelComposicao.rpCabTitulo.Caption = "Relatório de Composição de Produtos"
        FrmRelComposicao.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelComposicao.Config

End If
End Sub

Private Sub CmdDados_Click(Index As Integer)

On Error GoTo Erro

Dim SQL As String
Dim Total As Recordset
Dim Valor As Currency
Dim Mascara As String
Dim Criterio As String
Select Case Index
    Case 0
        If IsNumeric(DbcFornec.BoundText) = False Then
            MsgBox "Selecione o Fornecedor", vbCritical, App.Title
            DbcFornec.SetFocus
            Exit Sub
        End If
        If MebCusto = "" Then
            MebCusto.Text = Registros.Fields("Custo")
        End If
        If MebData = "__/__/____" Then
            MebData = Date
        End If
        Criterio = "CodPrd = " & Registros.Fields("CodPrd") & " and CodFor =" & DbcFornec.BoundText
        DatDados(2).Recordset.FindFirst Criterio
        If DatDados(2).Recordset.NoMatch Then
            SQL = "INSERT INTO Cot ( CodPrd, CodFor, Data, Preço, Condições ) Values ("
            SQL = SQL & Formulário.Registros.Fields("CodPrd") & ","
            SQL = SQL & DbcFornec.BoundText & " , #"
            SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , "
            SQL = SQL & Numero2(MebCusto) & " , "
            SQL = SQL & """" & TxtCond & """" & ")"
        Else
            'UPDATE COT SET COT.Condições = "kjkjk", COT.Preço = 50 WHERE (((COT.CodPrd)=25) AND ((COT.CodFor)=2));
            SQL = "UPDATE Cot SET cot.Preço = " & Numero2(MebCusto) & ","
            SQL = SQL & "cot.Data = " & "#" & Format(MebData, "mm/dd/yyyy") & "#,"
            SQL = SQL & "cot.condições = " & """" & TxtCond & """" & ","
            SQL = SQL & "cot.codfor = " & DbcFornec.BoundText
            SQL = SQL & " WHERE (((COT.CODPRD)="
            SQL = SQL & Registros.Fields("CodPrd") & ") AND ((COT.CODFOR)="
            SQL = SQL & DatDados(2).Recordset.Fields("CodFor") & "))"
        End If
        Banco.Execute SQL
        DbcFornec.BoundText = ""
        Mascara = MebCusto.Mask
        MebCusto.Mask = ""
        MebCusto.Text = ""
        MebCusto.Mask = Mascara
        Mascara = MebData.Mask
        MebData.Mask = ""
        MebData.Text = ""
        MebData.Mask = Mascara
        TxtCond = ""
        DatDados(2).Refresh
    Case 1
        If DatDados(2).Recordset.Eof = False Then
            If MsgBox(" Confirma Exclusão do Item ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                SQL = "DELETE FROM cot WHERE CodPrd = " & Registros.Fields("CodPrd") & " AND CodFor =" & DatDados(2).Recordset.Fields("CodFor")
                Banco.Execute SQL
                DatDados(2).Refresh
            End If
        End If
    Case 2
        DbcFornec.BoundText = ""
        Mascara = MebCusto.Mask
        MebCusto.Mask = ""
        MebCusto.Text = ""
        MebCusto.Mask = Mascara
        Mascara = MebData.Mask
        MebData.Mask = ""
        MebData.Text = ""
        MebData.Mask = Mascara
        TxtCond = ""
        DatDados(2).Refresh
End Select

Erro:
If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If
End Sub


Private Sub CmdFechar_Click(Index As Integer)
Unload Me
End Sub

Private Sub AtualizaComp()
MebTotComp = 0
If DatDados(6).Recordset.RecordCount <> 0 Then
    DatDados(6).Recordset.MoveFirst
    While Not DatDados(6).Recordset.Eof
        If DbgDados(1).Columns(4) <> "" Then
            MebTotComp = CCur(MebTotComp) + CCur(DbgDados(1).Columns(4))
        End If
        DatDados(6).Recordset.MoveNext
    Wend
End If

End Sub


Private Sub CmdDuplica_Click()
Dim Codigo As String
Dim Produto As String
Dim Grupo As String
Dim Fabricante As String
Dim Embalagem As String
Dim LocalEst As String
Dim CodFab As String
Dim Custo As String
Dim Margem As String
Dim Venda As String
Dim EstMin As String
Dim Tributo As String
Dim Aliquota As String
Dim Icm As String
Dim Ipi As String
Dim VarComplemento As String
Dim VarClasse As String
Dim VarSitTrib As String
Dim VarCfop  As String
Dim VarMarca As String

If MsgBox("Tem certeza que deseja duplicar esse registro", vbYesNo, App.Title) = vbYes Then
   Codigo = TxtDados(0)
   Produto = TxtDados(1)
   Grupo = DbcDados(0).Text
   Fabricante = DbcDados(1).Text
   Embalagem = TxtDados(5)
   LocalEst = TxtDados(6)
   CodFab = TxtDados(2)
   Custo = MebDados(21)
   Margem = MebDados(22)
   Venda = MebDados(23)
   EstMin = MebDados(1)
   Tributo = CmbDados(1)
   Aliquota = CmbDados(0)
   Icm = MebDados(7)
   Ipi = MebDados(9)
   VarComplemento = TxtDados(3)
   VarClasse = TxtDados(7)
   VarSitTrib = TxtDados(8)
   VarCfop = MebDados(33)
   VarMarca = TxtDados(13)
   
   Manutencoes 1
   
   TxtDados(0) = Codigo
   TxtDados(1) = Produto
   DbcDados(0).Text = Grupo
   DbcDados(1).Text = Fabricante
   TxtDados(5) = Embalagem
   TxtDados(6) = LocalEst
   TxtDados(2) = CodFab
   MebDados(21) = Custo
   MebDados(22) = Margem
   MebDados(23) = Venda
   MebDados(1) = EstMin
   If Tributo <> "" Then
        CmbDados(1).Text = Tributo
   End If
   If Aliquota <> "" Then
        CmbDados(0) = Aliquota
   End If
   MebDados(7) = Icm
   MebDados(9) = Ipi
   ChkDados(2).value = 1
   TxtDados(3) = VarComplemento
   TxtDados(7) = VarClasse
   TxtDados(8) = VarSitTrib
   MebDados(33) = VarCfop
   TxtDados(13) = VarMarca
   
   
   Set REG = Banco.OpenRecordset("Select CodPrd,CodBar From Prd Order By CodPrd")
   If REG.RecordCount > 0 Then
        REG.MoveLast
'        If Len(REG!CodPrd) < 7 Then
            TxtDados(0).Text = CDec(REG!CodPrd) + 1
'        End If
   Else
        TxtDados(0).Text = 1
   End If
        
End If
End Sub



Private Sub CmdFicha_Click()
Dim word As New word.Application

If Dir(App.Path & "\ficha\" & Trim(TxtDados(0)) & ".doc") <> "" Then
    word.Documents.Open App.Path & "\ficha\" & Trim(TxtDados(0)) & ".doc"
    word.Visible = True
    word.WindowState = wdWindowStateMaximize

End If

End Sub

Private Sub dbcdados_Change(Index As Integer)
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"

If Index = 1 Then
    If DbcDados(1).Text <> "" And Not IsNumeric(DbcDados(1).Text) Then
        Criterio = "Nome Like " & """" & DbcDados(1) & "*"""
        DatDados(3).Recordset.FindFirst Criterio
        If DatDados(3).Recordset.NoMatch Then
            Beep
            DbcDados(1).SetFocus
        End If
    End If
End If

If Index = 0 Then
    TxtCodGru = DbcDados(0).BoundText
End If
End Sub

Private Sub DbcDados_Click(Index As Integer, Area As Integer)
If Index = 0 Then
'      If DbcDados(0).BoundColumn <> "" Then
'        DatDados(0).Recordset.Bookmark = DbcDados(0).SelectedItem
'        CmbDados(0).Text = DatDados(0).Recordset.Fields("Uf")
'    End If
End If
End Sub

Private Sub DbcDados_LostFocus(Index As Integer)
Dim Criterio As String
Dim Terminacao As String
Dim Temp As String
Terminacao = """"

If Index = 1 Then
    If DbcDados(1).Text <> "" And Not IsNumeric(DbcDados(1).Text) Then
        Criterio = "Nome = " & """" & DbcDados(1) & """"
        DatDados(3).Recordset.FindFirst Criterio
        If DatDados(3).Recordset.NoMatch Then
            Beep
            DbcDados(1).SetFocus
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub DbcFornec_LostFocus()
If DbcFornec.Text <> "" Then
    MebCusto.SetFocus
End If
End Sub


Private Sub DbcPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And DbcPesquisa = "" Then
        FrmPesquisaPrd.Show vbModal
'        txtDados(2).SetFocus
'        TxtDados_KeyPress 2, 13
    End If
End Sub

Private Sub DbcPesquisa_LostFocus()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"

If IsNumeric(DbcPesquisa.Text) Then
    If Len(DbcPesquisa) > VarTamBarra Then
        DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto ORDER BY BAR.BarBarra"
        DatDados(4).Refresh
        Criterio = "BarBarra = "
        Criterio = Criterio & Terminacao & DbcPesquisa.Text & Terminacao
        DatDados(4).Recordset.FindFirst Criterio
        If DatDados(4).Recordset.NoMatch Then
            DbcPesquisa.Text = ""
            DbcPesquisa.SetFocus
        Else
            Criterio = "CodPrd = "
            Criterio = Criterio & DatDados(4).Recordset.Fields("Produto")
            Registros.FindFirst Criterio
            DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((BAR.Produto)=" & TxtDados(0) & ")) ORDER BY BAR.BarBarra"
            DatDados(4).Refresh
            MostraDados
            PosMovimentacao
        End If
    Else
        Criterio = "Prd.CodPrd = " & DbcPesquisa.BoundText
        DatDados(1).Recordset.FindFirst Criterio
        If DatDados(1).Recordset.NoMatch Then
            Beep
            DbcPesquisa.Enabled = True
            DbcPesquisa.Text = ""
            DbcPesquisa.SetFocus
        Else
            If Len(DbcPesquisa) >= 1 Then
                Registros.FindFirst Criterio
                MostraDados
                PosMovimentacao
            End If
        End If
    End If
End If
End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo

'Produto = Registros.Fields("CodPrd")
End Sub

Private Sub DBGDados_BeforeUpdate(Index As Integer, Cancel As Integer)
Select Case Index
    Case 1
'        DatDados(1).Recordset.Fields("CodPrd") = Produto
    
    Case 2
        DatDados(2).Recordset.Fields("CodPrd") = Produto
    Case 3
'        DatDados(3).Recordset.Fields("CodPrd") = Cliente

End Select
End Sub

Private Sub DbcPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If DbcPesquisa <> "" And Not IsNumeric(DbcPesquisa.Text) Then
    Criterio = "Prd.descrição_produto Like " & """" & DbcPesquisa & "*"""
    DatDados(1).Recordset.FindFirst Criterio
    If DatDados(1).Recordset.NoMatch Then
        Beep
'        DbcPesquisa.SetFocus
    Else
        If Len(DbcPesquisa) >= 1 Then
            Registros.FindFirst Criterio
            MostraDados
            PosMovimentacao
        End If
    End If
End If

End Sub

Private Sub DbcPesquisa_Click(Area As Integer)
'DbcPesquisa_Change
End Sub

Private Sub DbgDados_Change(Index As Integer)
If Index = 0 Then
    DbgDados(0).Columns(2) = Date
End If
End Sub

Private Sub DbgDados_Click(Index As Integer)
'    If DatDados(2).Recordset.EOF = False Then
'        DbcFornec.BoundText = DatDados(2).Recordset.Fields(3)
'        MebCusto = DatDados(2).Recordset.Fields(0)
'        MebData = DatDados(2).Recordset.Fields(1)
'        TxtCond = DatDados(2).Recordset.Fields(2)
'    End If
End Sub


Private Sub DbgDados_DblClick(Index As Integer)
If Index = 1 Then
    If DatDados(6).Recordset.RecordCount <> 0 Then
        If MsgBox("Deseja excluir produto da composição", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            DatDados(6).Recordset.Delete
            DatDados(6).Refresh
            AtualizaComp
        End If
    End If
End If
End Sub

Private Sub DbgGrade_dblClick()
If MsgBox("Deseja excluir ?", vbYesNo, App.Title) = vbYes Then
    DatDados(5).Recordset.Delete
End If

End Sub

Private Sub DbgIndices_dblClick()
DatDados(5).Recordset.AddNew
DatDados(5).Recordset.Fields("Produto") = TxtDados(0)
DatDados(5).Recordset.Fields("DescricaoLoja") = DbgIndices.Columns(2)
DatDados(5).Recordset.Update
End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
    DatDados(4).Recordset("Produto") = CCur(TxtDados(0).Text)
End Sub


Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  
'  DatDados(4).Recordset.MoveNext
End If
End Sub

Private Sub DBGrid2_Click()
Dim Criterio As String

Criterio = "Prd.CodPrd  =" & DBGrid2.Columns(0)
Registros.FindFirst Criterio
MostraDados
PosMovimentacao
ValorAnt = DBGrid2.Columns(5)

End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
DBGrid2_Click
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
DBGrid2_Click
End Sub

Private Sub DBGrid3_AfterColUpdate(ByVal ColIndex As Integer)
    If DatDados(5).Recordset.RecordCount <> 0 Then
        DatDados(5).Recordset.Edit
    Else
        DatDados(5).Recordset.AddNew
    End If
    DatDados(5).Recordset("Produto") = CCur(TxtDados(0).Text)
    DatDados(5).Recordset.Update
End Sub

Private Sub LblDados_DblClick(Index As Integer)
TxtDados(0).Enabled = True
End Sub

Private Sub MebDados_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 17 Then
    If KeyAscii = 13 Then
    '    SSTabPrd.Tab = 1
    End If
End If

If KeyAscii = 46 Then
    SendKeys "{Backspace}"
    SendKeys ","
End If

End Sub

Private Sub Opt_Click(Index As Integer)

If OPt(2).value = True Then
    If ChkFaltas.value = 0 Then
        DatDados(1).RecordSource = "SELECT Prd.Estoque, Prd.Essenciais, PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, Prd.MargemPrazo, Prd.PcoPrazo FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY PRD.Descrição_Produto"
    Else
        DatDados(1).RecordSource = "SELECT Prd.Estoque, Prd.Essenciais, PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda,  Prd.MargemPrazo, prd.falta, Prd.PcoPrazo FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo  where prd.falta = -1 ORDER BY PRD.Descrição_Produto"
    End If
ElseIf OPt(3).value = True Then
    If ChkFaltas.value = 0 Then
        DatDados(1).RecordSource = "SELECT Prd.Estoque, Prd.Essenciais, PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, Prd.MargemPrazo, Prd.PcoPrazo FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY  gru.descrição_grupo,PRD.Descrição_Produto"
    Else
        DatDados(1).RecordSource = "SELECT Prd.Estoque, Prd.Essenciais, PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, prd.falta , Prd.MargemPrazo, Prd.PcoPrazo FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo  where prd.falta = -1 ORDER BY gru.descrição_grupo,PRD.Descrição_Produto"
    End If
End If
DatDados(1).Refresh

End Sub

Private Sub SSTab1_Click(previoustab As Integer)
Dim VarSomaVda As Currency
Dim VarTiraVda As Currency
Dim VarCont(1 To 16) As Currency
Dim X As Byte
VarSomaVda = 0
VarTiraVda = 0
For X = 1 To 16
    VarCont(X) = 0
Next
Dim VarLbl1 As Currency
Dim VarLbl2 As Currency

Dim SQL As String
VarLbl1 = 0
VarLbl2 = 0
Select Case SSTab1.Tab

    Case 0
        SSTab1.Tab = 0
        If PctFundo(0).Enabled = False Then
            DbcPesquisa.SetFocus
        End If
         AbilitaBotoes (True)
    Case 1
        SSTab1.Tab = 1
        AbilitaBotoes (False)
        DbcFornec.SetFocus
    Case 2
        SSTab1.Tab = 2
        AbilitaBotoes (False)
        TxtPesquisa.SetFocus
    Case 4
        If IsNumeric(TxtDados(0)) Then
            If Dir("c:\sidcomp\Urupes\dados.mdb") = "" Then

                DatDados(8).RecordSource = "SELECT Ped.NroPedido, Ped.NroNt, Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                SQL = "SELECT Ped.NroNt, Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                DatDados(8).Refresh
                If DatDados(8).Recordset.RecordCount > 0 Then
                    DatDados(8).Recordset.MoveFirst
                    While Not DatDados(8).Recordset.Eof
                        Select Case DatDados(8).Recordset("Operacao")
                               Case 0
                                     VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                     If DatDados(8).Recordset("NroPedido") <> "" And DatDados(8).Recordset("NroPedido") <> "*" Then
                                        VarLbl1 = VarLbl1 + DatDados(8).Recordset("Qtd")
                                     End If
                               Case 1
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                         If DatDados(8).Recordset("NroNt") <> 0 Then
                                            VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                         End If
                                     
                               Case 2
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                         If DatDados(8).Recordset("NroNt") <> 0 Then
                                            VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                         End If
                               Case 3
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                        If DatDados(8).Recordset("NroNt") <> 0 Then
                                           VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                        End If
                               
                               Case 4
                               Case 5
                               Case 6
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                         If DatDados(8).Recordset("NroNt") <> 0 Then
                                            VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                         End If
                               Case 7
                               Case 8
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                        If DatDados(8).Recordset("NroNt") <> 0 Then
                                           VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                        End If
                               Case 9
                                     VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                         If DatDados(8).Recordset("NroPedido") <> "" And DatDados(8).Recordset("NroPedido") <> "*" Then
                                            VarLbl1 = VarLbl1 + DatDados(8).Recordset("Qtd")
                                         End If
                               Case 10
                                     VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                        If DatDados(8).Recordset("NroPedido") <> "" And DatDados(8).Recordset("NroPedido") <> "*" Then
                                           VarLbl1 = VarLbl1 + DatDados(8).Recordset("Qtd")
                                        End If
                              Case 11
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                         If DatDados(8).Recordset("NroNt") <> 0 Then
                                            VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                         End If
                               Case 12
                               Case 13
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                        If DatDados(8).Recordset("NroNt") <> 0 Then
                                           VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                        End If
                               Case 14
                                     VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                        If DatDados(8).Recordset("NroNt") <> 0 Then
                                           VarLbl2 = VarLbl2 + DatDados(8).Recordset("Qtd")
                                        End If
                               
                        End Select
                        VarCont(DatDados(8).Recordset("Operacao") + 1) = VarCont(DatDados(8).Recordset("Operacao") + 1) + DatDados(8).Recordset("Qtd")
                        DatDados(8).Recordset.MoveNext
                    Wend
                    List2.Clear
                    For X = 1 To 16
                        List2.AddItem VarCont(X)
                    Next
                    LblEntrada = "Total das Entradas:   " & VarSomaVda
                    LblSaida = "Total das Saidas:   " & VarTiraVda
                    lbl1 = Format(VarLbl1, "###,##0.00")
                    Lbl2 = Format(VarLbl2, "###,##0.00")
                
                
                End If
    
          Else
                If Dir("c:\sidcomp\Vendas\dados.mdb") <> "" Then
                    DatDados(8).DatabaseName = "c:\sidcomp\Vendas\Dados.mdb"
                    DatDados(8).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(8).Refresh
                    If DatDados(8).Recordset.RecordCount > 0 Then
                        DatDados(8).Recordset.MoveFirst
                        While Not DatDados(8).Recordset.Eof
                            Select Case DatDados(8).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(8).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(8).Recordset("Qtd")
                                                                        
                            End Select
                            VarCont(DatDados(8).Recordset("Operacao") + 1) = VarCont(DatDados(8).Recordset("Operacao") + 1) + DatDados(8).Recordset("Qtd")
                            DatDados(8).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                        
                    End If
                
                End If
                
                If Dir("c:\sidcomp\Urupes\dados.mdb") <> "" Then
                    DatDados(12).DatabaseName = "c:\sidcomp\Urupes\Dados.mdb"
                    
                    DatDados(12).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(12).Refresh
                    If DatDados(12).Recordset.RecordCount > 0 Then
                        DatDados(12).Recordset.MoveFirst
                        While Not DatDados(12).Recordset.Eof
                            Select Case DatDados(12).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(12).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(12).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(12).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(12).Recordset("Qtd")
                                   
                            End Select
                            VarCont(DatDados(12).Recordset("Operacao") + 1) = VarCont(DatDados(12).Recordset("Operacao") + 1) + DatDados(12).Recordset("Qtd")
                            DatDados(12).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                    End If
                End If
                
                
                
                If Dir("c:\sidcomp\novoh\dados.mdb") <> "" Then
                    DatDados(13).DatabaseName = "c:\sidcomp\novoh\Dados.mdb"
                    
                    DatDados(13).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(13).Refresh
                    If DatDados(13).Recordset.RecordCount > 0 Then
                        DatDados(13).Recordset.MoveFirst
                        While Not DatDados(13).Recordset.Eof
                            Select Case DatDados(13).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(13).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(13).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(13).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(13).Recordset("Qtd")
                                   
                            End Select
                            VarCont(DatDados(13).Recordset("Operacao") + 1) = VarCont(DatDados(13).Recordset("Operacao") + 1) + DatDados(13).Recordset("Qtd")
                            DatDados(13).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                    End If
                End If
            
                If Dir("c:\sidcomp\Itajobi\dados.mdb") <> "" Then
                    DatDados(14).DatabaseName = "c:\sidcomp\Itajobi\Dados.mdb"
                    
                    DatDados(14).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(14).Refresh
                    If DatDados(14).Recordset.RecordCount > 0 Then
                        DatDados(14).Recordset.MoveFirst
                        While Not DatDados(14).Recordset.Eof
                            Select Case DatDados(14).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(14).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(14).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(14).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(14).Recordset("Qtd")
                                   
                            End Select
                            VarCont(DatDados(14).Recordset("Operacao") + 1) = VarCont(DatDados(14).Recordset("Operacao") + 1) + DatDados(14).Recordset("Qtd")
                            DatDados(14).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                    End If
                End If
                
                If Dir("c:\sidcomp\Marapoama\dados.mdb") <> "" Then
                    DatDados(15).DatabaseName = "c:\sidcomp\Marapoama\Dados.mdb"
                    
                    DatDados(15).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(15).Refresh
                    If DatDados(15).Recordset.RecordCount > 0 Then
                        DatDados(15).Recordset.MoveFirst
                        While Not DatDados(15).Recordset.Eof
                            Select Case DatDados(15).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(15).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(15).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(15).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(15).Recordset("Qtd")
                                   
                            End Select
                            VarCont(DatDados(15).Recordset("Operacao") + 1) = VarCont(DatDados(15).Recordset("Operacao") + 1) + DatDados(15).Recordset("Qtd")
                            DatDados(15).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                    End If
                End If
                
                If Dir("c:\sidcomp\Irapua\dados.mdb") <> "" Then
                    DatDados(16).DatabaseName = "c:\sidcomp\Irapua\Dados.mdb"
                    
                    DatDados(16).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, ITP.ValorUnit, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " AND PED.DataEmiss >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY PED.DataEmiss"
                    DatDados(16).Refresh
                    If DatDados(16).Recordset.RecordCount > 0 Then
                        DatDados(16).Recordset.MoveFirst
                        While Not DatDados(16).Recordset.Eof
                            Select Case DatDados(16).Recordset("Operacao")
                                   Case 0
                                         VarSomaVda = VarSomaVda + DatDados(16).Recordset("Qtd")
                                   Case 1
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 2
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 3
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 4
                                   Case 5
                                   Case 6
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 7
                                   Case 8
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 9
                                         VarSomaVda = VarSomaVda + DatDados(16).Recordset("Qtd")
                                   Case 10
                                         VarSomaVda = VarSomaVda + DatDados(16).Recordset("Qtd")
                                  Case 11
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 12
                                   Case 13
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   Case 14
                                         VarTiraVda = VarTiraVda + DatDados(16).Recordset("Qtd")
                                   
                            End Select
                            VarCont(DatDados(16).Recordset("Operacao") + 1) = VarCont(DatDados(16).Recordset("Operacao") + 1) + DatDados(16).Recordset("Qtd")
                            DatDados(16).Recordset.MoveNext
                        Wend
                        List2.Clear
                        For X = 1 To 16
                            List2.AddItem VarCont(X)
                        Next
                        LblEntrada = "Total das Entradas:   " & VarSomaVda
                        LblSaida = "Total das Saidas:   " & VarTiraVda
                    End If
                End If
                
            End If
        End If
    
        If IsNumeric(TxtDados(0)) Then
            DatDados(9).RecordSource = "SELECT EST.Tipo, EST.Produto, EST.Data, EST.Qtd, EST.Loja, EST.LojaDestino FROM EST WHERE EST.Produto = " & TxtDados(0) & " AND EST.DATA >= #" & Format(DtDataMov, "mm/dd/yyyy") & "# ORDER BY EST.Data "
            DatDados(9).Refresh
            If DatDados(9).Recordset.RecordCount > 0 Then
                DatDados(9).Recordset.MoveLast
            End If
        End If
    
    
End Select
End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)
Formulário.Left = 200
Formulário.Top = 100

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
PosMovimentacao
DbcPesquisa.SetFocus
SSTab1.Tab = 0

'SSTabPrd.Tab = 0
'SSTabPrd.TabEnabled(0) = True
'SSTabPrd.TabEnabled(1) = True
'SSTabPrd.TabEnabled(2) = True
'SSTabPrd.TabEnabled(3) = True
FlagAltera = False

VarCodigo = 0
AtualizaComp
DtDataMov = Date - 180

VarEstoque = 0

VarEstoque = LblEstoque


End Sub

Public Sub PreConfirmacao() 'Será executada antes das rotinas de salvar ou cancelar

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Teste As String
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
Set Registros = Banco.OpenRecordset("SELECT PRD.*, GRU.* FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY Prd.Descrição_Produto")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
Opt_Click 2

LocalizacaoBasica = "Prd.Descrição_Produto"
DefineObjetos
SelectFromSQL = "Select * From Prd"
SelectFromFiltro = "SELECT Prd.Descrição_Produto, Prd.Custo, Prd.Venda  FROM Prd"
PesquisaSql = "SELECT Prd.CodPrd, Prd.Descrição_Produto FROM Prd ORDER BY Prd.Descrição_Produto"
FlagAbertura = False
OrderBySQL = "[Prd].[Descrição_Produto]"

DatDados(2).RecordSource = "SELECT  COT.Preço, COT.Data, COT.Condições, cot.codfor, cot.codprd, ent.nome,ent.fone, ent.Vendedor FROM PRD INNER JOIN (ENT INNER JOIN COT ON Ent.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((Cot.CodPrd)=0)) ORDER BY Cot.Preço"
DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((BAR.Produto)=0)) ORDER BY BAR.BarBarra"
DatDados(5).RecordSource = "SELECT * From Loj Where Produto = 0"

DatDados(6).RecordSource = "SELECT COMP.PrdComp,Comp.Produto, PRD.Descrição_Produto,  PRD.Local, COMP.Qtd, [comp].[qtd]*[prd].[Venda] AS Expr1 FROM COMP INNER JOIN PRD ON COMP.PrdComp = PRD.CodPrd WHERE (((COMP.Produto)= 0)) ORDER BY  PRD.Descrição_Produto"

TxtDados(1).MaxLength = Registros.Fields("Descrição_produto").Size

DatDadosCest.DatabaseName = Caminho & "\Minutos.mdb"

' Set Tabela = Banco.OpenRecordset("SELECT TABELA.MED_ABC, TABELA.MED_DES, TABELA.MED_APR, TABELA.MED_BARRA, TABELA.LAB_NOM FROM TABELA ORDER BY TABELA.MED_ABC")
' If Tabela.RecordCount > 0 Then
'     BtTab.Enabled = True
' End If

If VarPzoProduto = True Then
    Frame11.Enabled = True
'    Frame12.Enabled = True
Else
    Frame11.Enabled = False
'    Frame12.Enabled = False
End If

If Dir(App.Path & "\Importa.mdb") <> "" Then
    DatImporta.DatabaseName = App.Path & "\importa.mdb"
    DatImporta.RecordSource = "SELECT BAR.BarBarra, Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Tributo, PRD.Aliquota, PRD.Classe, PRD.SitTrib, PRD.CfOpPrd FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto"

End If
    



If LerINI("Pedido", "Grade", App.Path & "\config.ini") <> "" Then
    If LerINI("Pedido", "Grade", App.Path & "\config.ini") = 1 Then
        SSTab1.TabEnabled(5) = True
    Else
    SSTab1.TabEnabled(5) = False
    End If
Else
    SSTab1.TabEnabled(5) = False
End If

If LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini") <> "" Then
    VarRegime = LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini")
'        VarRegime = 3
'    End If
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
    Confirmacoes (15)
End If

AtivarMenu (False)

End Sub

Public Sub PosInclusao()
Dim Componentes As Variant

DbcPesquisa = ""

' sera executada logo depois que o usuario iniciar o processo de inclusao addnew

'DbcPesquisa = ""
For Each Componentes In MebDados
    Componentes.Text = 0
Next

Set REG = Banco.OpenRecordset("Select CodPrd,CodBar From Prd Order By CodPrd")
If REG.RecordCount > 0 Then
    REG.MoveLast
    If Len(REG!CodPrd) < VarTamCod Then
        TxtDados(0).Text = CDec(REG!CodPrd) + 1
    End If
Else
    TxtDados(0).Text = 1
End If
TxtDados(8).Text = "000"
MebDados(8).Text = ""
MebDados(6).Text = ""
DbcPesquisa.Enabled = False
ChkDados(2).value = 1
TxtDados(9).Text = 0
TxtDados(11).Text = 0
'TxtDados(10).Text = 0
MebDados(32).Text = 0
MebDados(34).Text = 0
MebDados(31).Text = 0
'DbcDados(0).Text = "DIVERSOS"
LblEstoque = 0
VarEstoque = 0
'DatDados(0).Recordset.MoveFirst

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
CmdDuplica.Enabled = False
'Salvar
'Registros.Edit
'TxtDados(1).SetFocus
End Sub

Public Function PermiteInclusao() As Boolean

' Será executado antes do processo de inclusao (addnew)
' Aqui deve ser feitos os processo p/ que o usuário pode ou não incluir um novo registro
' MsgBox "Neste formulário não pode ser feitas Inclusão de novos registros", vbInformation, App.Title

PermiteInclusao = True
'PosMovimentacao

End Function

Public Sub PosAlteracao()
Dim Componentes As Variant
' sera executada logo depois que o usuario iniciar o processo de alteracao
'Txtdados(0).BackColor = QBColor(8) 'Muda a Cor de fundo do código

For Each Componentes In MebDados
    If Componentes.Text = "" And Registros.Fields(Componentes.DataField).Type <> dbDate Then
        Componentes.Text = 0
    End If
Next
If DbcDados(0).Text <> "" Then
    FlagAltera = True
End If

Set REG = Banco.OpenRecordset("Select CodPrd,CodBar From Prd Order By CodPrd")
VarCodigo = TxtDados(0)
TxtDados(1).SetFocus
DbcPesquisa.Enabled = False

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
CmdDuplica.Enabled = False
'txtDados(0).Enabled = False
If TxtDados(9).Text = "" Then
    TxtDados(9).Text = 0
End If

If TxtDados(11).Text = "" Then
    TxtDados(11).Text = 0
End If

If TxtDados(10).Text = "" Then
'    TxtDados(10).Text = 0
End If

If MebDados(32).Text = "" Then
    MebDados(32).Text = 0
End If

If MebDados(34).Text = "" Then
    MebDados(34).Text = 0
End If

If MebDados(31).Text = "" Then
    MebDados(31).Text = 0
End If


End Sub

Public Function PermiteAlteracao() As Boolean
' Será executado antes do processo de Alteracao
If VarPermitePrd = False Then
    PermiteAlteracao = True
Else
    PermiteAlteracao = False
End If

If ChkDados(3).value = 1 Then
    VarPreco = MebDados(17).Text
Else
    VarPreco = MebDados(23).Text
End If

Open Caminho & "\Altera.txt" For Append As #1
Print #1, "Produto; " & Registros.Fields(1) & "; " & Registros.Fields(2) & "; " & Date & "; " & Time & "; " & Maquina
Close #1

End Function

Public Sub PosExclusao()
' sera executada logo depois que o usuario iniciar o processo de exclusao
Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
' Será executado antes do processo de exclusao
Dim Mensagem As String
Dim Codigo As Double
Dim SQL As String

'If VarPermitePrd = False Then
'    PermiteExclusao = True
'Else
'    PermiteExclusao = False
'End If

'    PermiteExclusao = True
DatDados(8).RecordSource = "SELECT Ped.Contato, PED.Operacao, PED.CodPed, PED.DataEmiss, ENT.Nome, ITP.Qtd, ITP.VlrVdaDia, ITP.Produto, Itp.CustoDia, Itp.VendaDia FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE ITP.Produto = " & TxtDados(0) & " ORDER BY PED.DataEmiss"
DatDados(8).Refresh

If DatDados(8).Recordset.RecordCount > 0 Then
    MsgBox "Produto já tem movimento ! A Exclusão acarreta perca de informação.", vbInformation, App.Title
'    PermiteExclusao = False
'    Exit Function
End If


Mensagem = "Confirma Exclusão do Produto " & vbCr & Registros.Fields(2)
Codigo = Registros.Fields(1)
If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes And VarPermitePrd = False Then
    ' nesta posicao do codigo podemos incluir os procedimentos que deverao
    ' apagar todos os registro que dependem do registro que está sendo excluido, no caso de relacionamento entre tabelas
    ' ex : poderiamos excluir os itens de um pedido antes de excluir o cabeçalho

    
    SQL = "DELETE BAR.*, BAR.Produto FROM BAR WHERE BAR.Produto = " & TxtDados(0)
    Banco.Execute SQL
    
    SQL = "DELETE loj.*, LOJ.Produto FROM LOJ WHERE LOJ.Produto = " & TxtDados(0)
    Banco.Execute SQL
    
    SQL = "DELETE Comp.*, COMP.Produto FROM COMP WHERE COMP.Produto = " & TxtDados(0)
    Banco.Execute SQL
    
    SQL = "DELETE Cot.*, CodPrd FROM COT WHERE CodPrd = " & TxtDados(0)
    Banco.Execute SQL
        
    Open Caminho & "\Delete.txt" For Append As #1
    Print #1, "Produto; " & Registros.Fields(1) & "; " & Registros.Fields(2) & "; " & Date & "; " & Time & "; " & Maquina
    Close #1
    
    PermiteExclusao = True
    
Else
    PermiteExclusao = False
End If
End Function

Public Sub PosConfirmacao() 'Será executada deposi dos processos de salvar ou cancelar
Dim Criterio As String


DbcPesquisa.Enabled = True
DbcPesquisa.SetFocus
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = True
CmdDuplica.Enabled = True
TxtDados(0).Enabled = False

End Sub

Public Function PermiteCancelamento() As Boolean
' Será executada antes do processo de cancelamento ser executado (cancelupdate)
' podemos aqui impedir o usuario de cancelar uma inclusao ou uma alteracao

PermiteCancelamento = True
'TxtBarra = ""
End Function

Public Function PermiteSalvar() As Boolean
' Será axecutada antes de atribuir os campos do formulário no bco de dados
' Podemos fazer qualquer tipo de consistencia que não possa ser feita pelo bco de  dados
' Ex: O estoque máximo sempre deverá ser maior que o estoque minimo
If FlagAltera = False Then
    Dim RegPrd As Recordset
        Set RegPrd = Banco.OpenRecordset("Select Descrição_Produto From Prd Where Descrição_Produto = '" & TxtDados(1).Text & "'")
        If RegPrd.RecordCount > 0 Then
            If MsgBox("Produto duplicado, Deseja continuar ?", vbYesNo, App.Title) = vbYes Then
                PermiteSalvar = True
            Else
                PermiteSalvar = False
            End If
        Else
                PermiteSalvar = True
        End If
Else
    PermiteSalvar = True
End If

End Function

Public Function PermiteContinuarSalvar() As Boolean
' Será executada logo depois de atribuir os campos ao bco de dados
' Possui as mesma função do permitir salvar, Inclusive podendo atribuir valores a campos do bco de dados que não aparecem na tela
' Ex: Aqui vc pode atribuir um valor ao bco que não está no formulario
'If VarCodigo <> 0 Then
'    If VarCodigo <> TxtDados(0) Then
'        MsgBox "Você está tentando mudar o código do produto, isto podera alterar as vendas", vbInformation, App.Title
'    End If
'End If
If FlagAltera = False And TxtDados(1).Text <> "" Then
    If LerINI("Pedido", "Grade", App.Path & "\config.ini") <> "" Then
        If LerINI("Pedido", "Grade", App.Path & "\config.ini") = 1 Then
 '           FormGradePrd.Show
        Else
            VarEstoque = InputBox("Digite o valor do estoque inicial !")
            If VarEstoque = "" Then VarEstoque = 0
        End If
     End If
End If
FlagAltera = False
PermiteContinuarSalvar = True
LblEstoque = 0
''Datdados(1).Refresh

End Function

Public Function PosSalvamento() As Boolean
Dim Criterio As String
'Será executado assim que todos os dados forem gravados no bco

'If DatDados(5).Recordset.RecordCount = 1 Then
If LocalEstoque <> "" Then
    Criterio = "Produto = " & TxtDados(0).Text & " AND DescricaoLoja = " & """" & LocalEstoque & """"
    DatDados(5).Recordset.FindFirst Criterio
    If DatDados(5).Recordset.NoMatch Then
       DatDados(5).Recordset.AddNew
       DatDados(5).Recordset.Fields("Produto") = TxtDados(0).Text
       DatDados(5).Recordset.Fields("DescricaoLoja") = LocalEstoque
       DatDados(5).Recordset.Fields("Estoque") = VarEstoque
       DatDados(5).Recordset.Update
    End If
End If

Criterio = "CODPRD = " & TxtDados(0).Text
Registros.FindFirst Criterio
If Not Registros.NoMatch Then
    Registros.Edit
    Registros!estoque = VarEstoque
    Registros.Update
End If


Criterio = "Produto = " & TxtDados(0).Text
DatDados(4).Recordset.FindFirst Criterio
If DatDados(4).Recordset.NoMatch Then
   DatDados(4).Recordset.AddNew
   DatDados(4).Recordset.Fields("Produto") = TxtDados(0).Text
   DatDados(4).Recordset.Fields("BarBarra") = TxtDados(0).Text
   DatDados(4).Recordset.Update
   If Len(TxtDados(2)) >= 7 Then
        DatDados(4).Recordset.AddNew
        DatDados(4).Recordset.Fields("Produto") = TxtDados(0).Text
        DatDados(4).Recordset.Fields("BarBarra") = TxtDados(2).Text
        DatDados(4).Recordset.Update
   Else
        DatDados(4).Recordset.AddNew
        DatDados(4).Recordset.Fields("Produto") = TxtDados(0).Text
        DatDados(4).Recordset.Fields("BarBarra") = "789" & psCompString(1, TxtDados(0).Text, 7)
        DatDados(4).Recordset.Update
   
   
   End If
   
End If

If IsNumeric(DbcDados(1).BoundText) Then
    If DbcDados(1).BoundText > 0 Then
        Criterio = "CodPrd = " & TxtDados(0).Text & " AND CodFor = " & DbcDados(1).BoundText
        DatDados(2).Recordset.FindFirst Criterio
        If DatDados(2).Recordset.NoMatch Then
           DatDados(2).Recordset.AddNew
           DatDados(2).Recordset.Fields("CodPrd") = TxtDados(0).Text
           DatDados(2).Recordset.Fields("CodFor") = DbcDados(1).BoundText
           DatDados(2).Recordset.Fields("Data") = Date
           DatDados(2).Recordset.Fields("Preço") = 0
           DatDados(2).Recordset.Fields("Condições") = " "
           DatDados(2).Recordset.Update
        End If
    End If
End If

DatDados(0).Refresh 'Atualiza a lista
'DatDados(1).Refresh
DatDados(4).Refresh
'DatDados(1).Refresh

Criterio = "CodPrd = " & TxtDados(0)
DatDados(1).Recordset.FindFirst Criterio
If Not DatDados(1).Recordset.NoMatch Then
    
    
        If ChkDados(3).value = -1 Then
            If VarPreco <> MebDados(17).Text Then
                DatDados(1).Recordset.Edit
                DatDados(1).Recordset("Essenciais") = -1
                DatDados(1).Recordset.Update
            End If
        Else
            If VarPreco <> MebDados(23).Text Then
                DatDados(1).Recordset.Edit
                DatDados(1).Recordset("Essenciais") = -1
                DatDados(1).Recordset.Update
            End If
        End If
        
        
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset("Estoque") = VarEstoque
        DatDados(1).Recordset.Update

End If


'AtualizaChefe
'MebDados_lostfocus 0

PosSalvamento = True
PosMovimentacao

End Function

Public Sub PosCancelamento()
PosMovimentacao
'sera executado logo após do processo de cancelamento da aplicaçaõ

End Sub
Public Sub PosMovimentacao()
Dim Criterio As String
If Registros.RecordCount > 0 Then
    If Registros.Fields("Foto") Then
        If Dir(App.Path & "\fotos\" & Trim(TxtDados(0)) & ".jpg") <> "" Then
            ImgFoto.Visible = True
            LblFoto.Visible = False
    '        ImgFoto.Picture = LoadPicture(App.Path & "\fotos\" & Trim(CStr(Registros.Fields("CodPrd"))) & ".jpg")
            ImgFoto.Picture = LoadPicture(App.Path & "\fotos\" & Trim(TxtDados(0)) & ".jpg")
        Else
            ImgFoto.Visible = False
            LblFoto.Visible = True
        End If
    Else
        ImgFoto.Visible = False
        LblFoto.Visible = True
    End If
'    FormProdutos.Caption = "===>    " & txtDados(1).Text
    LblTitulo.Caption = TxtDados(1).Text
    If IsNumeric(DbcDados(0).BoundText) Then
        TxtCodGru = DbcDados(0).BoundText
    End If
'    If TxtDados(10).Text = "" Then
'        LblDes = ""
'        LblApres = ""
'        LblLab = ""
'    Else
'        If Tabela.RecordCount > 0 Then
'            Criterio = "Med_Barra = " & """" & TxtDados(2) & """"
'            Tabela.FindFirst Criterio
'            If Not Tabela.NoMatch Then
'                LblDes = Tabela!MED_DES
'                LblApres = Tabela!MED_APR
'                LblLab = Tabela!lab_Nom
'            Else
 '               LblDes = ""
'                LblApres = ""
'                LblLab = ""
'            End If
'        End If
'    End If
    If Len(TxtDados(2)) >= 8 And IsNumeric(TxtDados(2)) Then
        PctBarra.Visible = True
        PctBarra.ScaleMode = 3
        PctBarra.Height = PctBarra.Height * (1.2 * 30 / PctBarra.ScaleHeight)
        PctBarra.FontSize = 6
        Call DrawBarcode(TxtDados(2), PctBarra)
    Else
        PctBarra.Visible = False
    End If
    
    If MebDados(21) <> "0" And MebDados(21) <> "" Then
       If MebDados(23) <> "0" And MebDados(23) <> "" Then
            MebPerc.Text = (1 - (CCur(MebDados(21)) / CCur(MebDados(23)))) * 100
       End If
    Else
        MebPerc = 0
    End If
    
    LblEstoque = 0
    
    If DatDados(5).Recordset.RecordCount > 0 Then
        DatDados(5).Recordset.MoveFirst
        While Not DatDados(5).Recordset.Eof
            LblEstoque = LblEstoque + DatDados(5).Recordset("Estoque")
            DatDados(5).Recordset.MoveNext
        Wend
    End If
    
    VarEstoque = LblEstoque
    
    AtualizaComp
End If
End Sub


Public Sub Atualizar()
' Atualiza a variavel registro incluindo os novos registros cadastrados e excluindo os que foram deletados
Registros.Requery
DatDados(0).Refresh
DatDados(1).Refresh
DatDados(4).Refresh
End Sub

Private Sub MebCusto_GotFocus()
MebCusto.SelStart = 0
MebCusto.SelLength = 12

End Sub

Private Sub MebDados_GotFocus(Index As Integer)
'SendKeys "+{end}"
MebDados(Index).SelStart = 0
MebDados(Index).SelLength = 12
End Sub

Private Sub MebDados_LostFocus(Index As Integer)
Dim VarMargem As Currency
Dim Temp As Double
Dim Criterio As String
Dim VarCustoPrd As Currency
Dim VarCustoPrd2 As Currency
Dim VarImposto As Currency
Dim VarRedutor As Currency
Dim VarFrete As Currency
Dim VarIpi As Currency

VarCustoPrd = 0
VarMargem = 0
VarImposto = 0
VarCustoPrd2 = 0
VarRedutor = 0
VarFrete = 0
VarIpi = 0

If IsNumeric(MebDados(Index)) Then
        
        If Index = 34 Then

            VarFrete = CCur(MebDados(34)) * (CCur(MebDados(32))) / 100
            VarIpi = (CCur(MebDados(34)) + VarFrete) * (CCur(MebDados(9))) / 100
            VarRedutor = CCur(MebDados(34)) * (CCur(MebDados(31))) / 100
            
            VarCustoPrd = CCur(MebDados(34)) + VarIpi + VarFrete
            If TxtDados(9) <> 0 Then
                VarCustoPrd2 = VarCustoPrd + (VarCustoPrd * CCur(TxtDados(9)) / 100)
                VarImposto = VarCustoPrd2 * (CCur(MebDados(7)) / 100)
            End If
            CustoReal = VarCustoPrd + (VarImposto - VarRedutor)
            If MebDados(34) <> 0 Then
                MebDados(21) = MebDados(34)
            End If
            If MebDados(34) <> 0 Then
                MebcustoIva = ((CCur(CustoReal) / CCur(MebDados(34))) - 1) * 100
            End If
        End If

        If Index = 21 Then
                MebDados(23).Text = CCur(MebDados(21)) + (CCur(MebDados(21)) * (CCur(MebDados(22)) / 100))
                MebDados(10) = Date
    
'                temp = TxtDados(0).Text
'                Salvar
'                Criterio = "CodPrd = " & temp
'                Registros.FindFirst Criterio
       '        Registros.Bookmark = Registros.LastModified
       '        Registros.Edit
                
        End If
        
        If Index = 22 Then
            If MebDados(21) <> "0" And MebDados(21) <> "" Then
                If VarPrecoCima = 1 Then
                    VarMargem = (100 / (1 - (MebDados(22) / 100))) - 100
                    MebDados(23).Text = CCur(MebDados(21)) + (CCur(MebDados(21)) * (CCur(VarMargem) / 100))
                Else
                    MebDados(23).Text = CCur(MebDados(21)) + (CCur(MebDados(21)) * (CCur(MebDados(22)) / 100))
                End If
                MebDados(10) = Date
            End If
        End If
        
        If Index = 23 Then
            If MebDados(21) <> "0" And MebDados(21) <> "" Then
                MebDados(22).Text = ((CCur(MebDados(23)) / CCur(MebDados(21))) - 1) * 100
                If MebDados(23) <> "0" And MebDados(23) <> "" Then
                    MebPerc.Text = (1 - (CCur(MebDados(21)) / CCur(MebDados(23)))) * 100
                End If
                MebDados(10) = Date
            End If
        End If
        
        If Index = 21 Or Index = 22 Or Index = 23 Then
            If MebDados(14) <> "0" And MebDados(14) <> "" Then
                MebDados(15).Text = CCur(MebDados(23)) + (CCur(MebDados(23)) * (CCur(MebDados(14)) / 100))
            Else
                MebDados(15).Text = CCur(MebDados(23))
            End If
            If MebDados(16) <> "0" And MebDados(14) <> "" Then
                MebDados(17).Text = CCur(MebDados(23)) - (CCur(MebDados(23)) * (CCur(MebDados(16)) / 100))
            Else
                MebDados(17).Text = CCur(MebDados(23))
            End If
            
        End If
                        
        If Index = 14 Then
            If MebDados(23) <> "0" And MebDados(23) <> "" Then
                MebDados(15).Text = CCur(MebDados(23)) + (CCur(MebDados(23)) * (CCur(MebDados(14)) / 100))
            End If
        End If
                        
        If Index = 16 Then
            If MebDados(23) <> "0" And MebDados(23) <> "" Then
                MebDados(17).Text = CCur(MebDados(23)) - (CCur(MebDados(23)) * (CCur(MebDados(16)) / 100))
            End If
        End If
        
        If Index = 15 Then
            If MebDados(23) <> "0" And MebDados(23) <> "" Then
                MebDados(14).Text = ((CCur(MebDados(15)) / CCur(MebDados(23))) - 1) * 100
            End If
        End If

        If Index = 17 Then
            If MebDados(23) <> "0" And MebDados(23) <> "" Then
                MebDados(16).Text = (1 - (CCur(MebDados(17)) / CCur(MebDados(23)))) * 100
            End If
        End If

Else
    If Index <> 6 Then
        MebDados(Index) = ""
        MebDados(Index).SetFocus
    End If
End If
End Sub

Private Sub MebData_GotFocus()
MebData.SelStart = 0
MebData.SelLength = 12
End Sub

Private Sub SSTab2_DblClick()

End Sub

Private Sub TxtBusca_GotFocus()
TxtBusca = ""
End Sub

Private Sub TxtBusca_LostFocus()
Dim Teste As String
Dim Codigo As String
Dim Produto As String
Dim Grupo As String
Dim Fabricante As String
Dim Embalagem As String
Dim LocalEst As String
Dim CodFab As String
Dim Custo As String
Dim Margem As String
Dim Venda As String
Dim EstMin As String
Dim Tributo As String
Dim Aliquota As String
Dim Icm As String
Dim Ipi As String
Dim VarComplemento As String
Dim VarClasse As String
Dim VarSitTrib As String
Dim VarCfop  As String
Dim VarMarca As String
Dim VarImporta As Recordset
Dim varPesquisa As String
Teste = ""
If Dir(App.Path & "\Importa.mdb") <> "" Then
 '   Datimport.RecordSource = "SELECT BAR.BarBarra, Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Tributo, PRD.Aliquota, PRD.Classe, PRD.SitTrib, PRD.CfOpPrd FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto"
 '   DatImporta.RecordSource = "SELECT Prd.CodBar, Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Tributo, PRD.Aliquota, PRD.Classe, PRD.SitTrib, PRD.CfOpPrd FROM PRD ORDER BY CodPrd"
    
    DatImporta.Refresh
    If DatImporta.Recordset.RecordCount > 0 Then
        DatImporta.Recordset.MoveLast
        DatImporta.Refresh
        DatImporta.Recordset.MoveFirst
        DatImporta.Refresh
        varPesquisa = "BarBarra = " & """" & TxtBusca & """"
        DatImporta.Recordset.FindFirst varPesquisa
        If Not DatImporta.Recordset.NoMatch Then
        
            If MsgBox("Tem certeza que deseja importar esse produto", vbYesNo, App.Title) = vbYes Then
               Codigo = DatImporta.Recordset("CodPrd")
               Produto = DatImporta.Recordset("Descrição_Produto")
               Grupo = 1
               Embalagem = DatImporta.Recordset("Embalagem")

               Tributo = DatImporta.Recordset("Tributo")
               Aliquota = DatImporta.Recordset("Aliquota")

               VarClasse = DatImporta.Recordset("Classe")
               VarSitTrib = DatImporta.Recordset("SitTrib")
               VarCfop = DatImporta.Recordset("CfopPrd")
               VarMarca = DatImporta.Recordset("Barbarra")
               
               Manutencoes 1
               
               TxtDados(0) = Codigo
               TxtDados(1) = Produto
               DbcDados(0).Text = "DIVERSOS"

               TxtDados(5) = Embalagem

               If Tributo <> "" Then
                    CmbDados(1).ListIndex = Tributo - 1
               End If
               If Aliquota <> "" Then
                    CmbDados(0).ListIndex = Aliquota - 1
               End If
               ChkDados(2).value = 1
               TxtDados(7) = VarClasse
               TxtDados(8) = VarSitTrib
               MebDados(33) = VarCfop
               TxtDados(2) = VarMarca
               
               
               Set REG = Banco.OpenRecordset("Select CodPrd,CodBar From Prd Order By CodPrd")
               If REG.RecordCount > 0 Then
                    REG.MoveLast
            '        If Len(REG!CodPrd) < 7 Then
                        TxtDados(0).Text = CDec(REG!CodPrd) + 1
            '        End If
               Else
                    TxtDados(0).Text = 1
               End If
                    
            End If
        
        End If
    End If
Else
    Web.Navigate "http://sandbox.buscape.com/service/findOfferList/564771466d477a4458664d3d/?barcode=" & Trim(TxtBusca)
    
        MsgBox "Usando buscape para pesquisar produto !", vbInformation, App.Title
    'End If
    Teste = Web.Document.Body.InnerText
    If Left(Teste, 36) <> " The website cannot display the page" Then
        RecuperarListagemDeContas
    Else
        MsgBox "Endereço de pesquisa não encontrado ", vbInformation, App.Title
        TxtBusca = "Pesquisa Buscapé"
    
    End If

End If

End Sub

Private Sub RecuperarListagemDeContas()

    Dim sTempFile As String
    Dim sTempLineInput As String
    Dim sTempItemIndex As String
    Dim iItemIndex As Integer
    Dim VCodigo As Double
    
    iItemIndex = 0

    sTempFile = App.Path & "\TempSourceCode.txt"

    Open sTempFile For Output As #1
    Print #1, Web.Document.Body.InnerText
    Close #1

    Open sTempFile For Input As #1
    While Not Eof(1)
        Line Input #1, sTempLineInput

        If Left(sTempLineInput, 13) = "  <offerName>" And iItemIndex = 0 Then
            sTempLineInput = Mid$(sTempLineInput, 14, 200)
            For iItemIndex = 1 To Len(sTempLineInput)
                If Mid(sTempLineInput, iItemIndex, 1) <> "<" Then
                    sTempItemIndex = sTempItemIndex & Mid(sTempLineInput, iItemIndex, 1)
                Else
                    iItemIndex = Len(sTempLineInput)
'                    sTempLineInput = sTempItemIndex
                End If
            Next

        End If
    Wend
    Close #1
    If sTempItemIndex = "" Then
        MsgBox "Produto não cadastrada no Buscapé ", vbInformation, App.Title
    Else
       If MsgBox(sTempItemIndex, vbYesNo, App.Title) = vbYes Then
            VCodigo = TxtDados(0)
            
            Manutencoes 1
            TxtDados(0) = VCodigo + 1
            TxtDados(2) = TxtBusca
            TxtDados(5) = "UN"
            TxtDados(1) = Format(sTempItemIndex, ">")
            Set REG = Banco.OpenRecordset("Select CodPrd,CodBar From Prd Order By CodPrd")
            If REG.RecordCount > 0 Then
                 REG.MoveLast
                 If Len(REG!CodPrd) < 7 Then
                     TxtDados(0).Text = CDec(REG!CodPrd) + 1
                 End If
            Else
                 TxtDados(0).Text = 1
            End If
            
        End If
    End If
    TxtBusca = "Pesquisa Buscapé"

End Sub

Private Sub TxtCond_gotfocus()
TxtCond.SelStart = 0
TxtCond.SelLength = 12

End Sub

Private Sub Txtdados_Change(Index As Integer)
Dim Criterio As String
If Index = 1 Then
    QtdCorpo = Len(TxtDados(1))
End If

If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
    If IsNumeric(TxtDados(0)) = False Then
        DatDados(2).RecordSource = "SELECT  COT.Preço, COT.Data, COT.Condições, cot.codfor, cot.codprd, ent.nome,ent.fone, ent.Vendedor FROM PRD INNER JOIN (ENT INNER JOIN COT ON Ent.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((Cot.CodPrd)=0)) ORDER BY Cot.Preço"
        DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((BAR.Produto)=0)) ORDER BY BAR.BarBarra"
        DatDados(5).RecordSource = "SELECT * From Loj Where Produto = 0 ORDER BY DescricaoLoja"
        DatDados(6).RecordSource = "SELECT COMP.PrdComp,Comp.Produto, PRD.Descrição_Produto,  PRD.Local, COMP.Qtd, [comp].[qtd]*[prd].[Venda] AS Expr1 FROM COMP INNER JOIN PRD ON COMP.PrdComp = PRD.CodPrd WHERE (((COMP.Produto)= 0)) ORDER BY  PRD.Descrição_Produto"

    Else
        DatDados(2).RecordSource = "SELECT  COT.Preço, COT.Data, COT.Condições, cot.codfor, cot.codprd, ent.nome ,ent.fone, ent.Vendedor FROM PRD INNER JOIN (ENT INNER JOIN COT ON Ent.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((Cot.CodPrd)= " & TxtDados(0) & " )) ORDER BY Cot.Preço"
        DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((BAR.Produto)=" & TxtDados(0) & ")) ORDER BY BAR.BarBarra"
        DatDados(5).RecordSource = "SELECT * From Loj Where Produto = " & CCur(TxtDados(0).Text) & " ORDER BY DescricaoLoja"
        DatDados(6).RecordSource = "SELECT COMP.PrdComp, Comp.Produto, PRD.Descrição_Produto, PRD.Local, COMP.Qtd, [comp].[qtd]*[prd].[Venda] AS Expr1 FROM COMP INNER JOIN PRD ON COMP.PrdComp = PRD.CodPrd WHERE (((COMP.Produto)= " & CCur(TxtDados(0).Text) & ")) ORDER BY  PRD.Descrição_Produto"
    End If
    
    DatDados(2).Refresh
    DatDados(4).Refresh
    DatDados(5).Refresh
    DatDados(6).Refresh

'    If DatDados(5).Recordset.RecordCount = 1 Then
'        DBGrid1.Height = 4150
'    End If


'    PosMovimentacao
End If

End Sub
Private Sub AtualizaChefe()
Dim SQL As String
If Registros.Fields("Chefe_grupo") <> "" Then
    SQL = "UPDATE PRD SET PRD.Margem = " & Numero2(MebDados(22)) & ","
    SQL = SQL + " Prd.Custo = " & Numero2(MebDados(21)) & ","
    SQL = SQL + " Prd.Venda = " & Numero2(MebDados(23)) & ","
    SQL = SQL + " Prd.PcoPrazo = " & Numero2(MebDados(15)) & ","
    SQL = SQL + " Prd.Venda_Concorrencia = " & Numero2(MebDados(25))
'    SQL = SQL + " Prd.concorrencia = " & ChkDados(1).value & ","
'    SQL = SQL + " Prd.promocao = " & ChkDados(0).value
    SQL = SQL + " WHERE (((Prd.Chefe_Grupo)= " & """" & Registros.Fields("CodPrd") & """" & "))"
    Banco.Execute SQL
End If
End Sub

Private Sub TxtDados_GotFocus(Index As Integer)
If Index = 4 Then
    FormProdutos.KeyPreview = False
Else
    FormProdutos.KeyPreview = True
End If

If Index = 12 Then
    If TxtDados(7) <> "" Then
        DatDadosCest.RecordSource = "SELECT CEST.Descricao, CEST.Segmento, CEST.CEST, CEST.[NCM/SH] FROM CEST  WHERE  CEST.[NCM/SH] = " & """" & Format(LIMPACampo2(TxtDados(7)), "@@@@.@@.@@") & """" & "ORDER BY CEST.Descricao"
        DatDadosCest.Refresh
        If DatDadosCest.Recordset.RecordCount > 0 Then
            TxtDados(12) = DatDadosCest.Recordset.Fields("CEST")
        End If
    End If
End If


End Sub

Private Sub TxtDados_LostFocus(Index As Integer)
Dim Criterio As String
Dim Temp As Long
If Index = 0 And TxtDados(1).Text = "" Then
    If TxtDados(0) <> "" Then
        Criterio = "CodPrd = " & TxtDados(0).Text
        If TxtDados(1).Text = "" Then
            DatDados(1).Recordset.FindFirst Criterio
            If Not DatDados(1).Recordset.NoMatch Then
                MsgBox "Produto já cadastrado", vbExclamation, App.Path
                TxtDados(0).Text = ""
                TxtDados(0).SetFocus
            End If
        End If
    Else
        MsgBox "O campo codigo tem que ser preenchido", vbInformation, App.Title
        TxtDados(0).SetFocus
        Exit Sub
    End If
End If

If Index = 4 Then
    FormProdutos.KeyPreview = True
End If


End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" And Not IsNumeric(TxtPesquisa.Text) Then
    Criterio = "Prd.descrição_produto  Like " & """" & TxtPesquisa & "*"""
    DatDados(1).Recordset.FindFirst Criterio
    If DatDados(1).Recordset.NoMatch Then
        Beep
        TxtPesquisa.SetFocus
        SendKeys "{BACKSPACE}"
    Else
        If Len(TxtPesquisa) >= 1 Then
            Registros.FindFirst Criterio
            DBGrid2.Refresh
            MostraDados
            PosMovimentacao
        End If
    End If
End If

End Sub


Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim Criterio As String

If DBGrid2.Columns(3) <> "" And DBGrid2.Columns(4) <> "" And DBGrid2.Columns(5) <> "" Then
    If OPt(1).value = True Then
         DBGrid2.Columns(5) = DBGrid2.Columns(3) + (DBGrid2.Columns(3) * (DBGrid2.Columns(4) / 100))
    ElseIf OPt(0).value = True Then
         DBGrid2.Columns(4) = (((DBGrid2.Columns(5) / DBGrid2.Columns(3)) - 1) * 100)
    End If
    
    If VarPzoProduto = True Then
          DBGrid2.Columns(7) = DBGrid2.Columns(5) + (DBGrid2.Columns(5) * (DBGrid2.Columns(6) / 100))
    End If
    
    
    If Registros.Fields("Venda") <> ValorAnt Then
        Registros.Edit
        Registros.Fields("DtAlt") = Date
        Registros.Update
    End If
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

Private Sub TxtPesquisa_LostFocus()
Dim Criterio As String
If IsNumeric(TxtPesquisa) Then
    Criterio = "Prd.CodPrd = " & TxtPesquisa
    DatDados(1).Recordset.FindFirst Criterio
    If DatDados(1).Recordset.NoMatch Then
        Beep
        TxtPesquisa.Text = ""
        TxtPesquisa.SetFocus
    Else
        If Len(TxtPesquisa) >= 1 Then
            Registros.FindFirst Criterio
            MostraDados
            PosMovimentacao
        End If
    End If
End If
End Sub

