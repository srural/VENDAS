VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmConfiguracao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12165
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   12360
      TabIndex        =   74
      Top             =   7020
      Width           =   12390
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14655
      TabIndex        =   72
      Top             =   0
      Width           =   14655
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Configurações do Sistema"
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
         TabIndex        =   73
         Top             =   45
         Width           =   7800
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "CFG"
      Connect         =   "Access"
      DatabaseName    =   "C:\Rede\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   8505
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4815
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.CommandButton BtOk 
      Height          =   420
      Left            =   11340
      Picture         =   "FrmConfiguracao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Gravar"
      Top             =   1170
      Width           =   435
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5940
      Left            =   360
      TabIndex        =   7
      Top             =   810
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   10478
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "FrmConfiguracao.frx":0532
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Nota"
      TabPicture(1)   =   "FrmConfiguracao.frx":054E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "DatDados(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Diversos"
      TabPicture(2)   =   "FrmConfiguracao.frx":056A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label18"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "ChkEstoque"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "TxtTamBarra"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TxtTamCod"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "ChkTranspNt"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "ChkPrecoCima"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "ChkParcela"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "ChkTipoPgto"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "TxtBcoFin"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "ChkFinanc"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "ChkFunRural"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "ChkPrazo"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "ChkEditar"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "ChkReceber"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "ChkCompra"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "ChkPzoProduto"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "TxtEvento"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "TxtTaxa"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      Begin VB.TextBox TxtTaxa 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5265
         MaxLength       =   255
         TabIndex        =   77
         Top             =   2430
         Width           =   1830
      End
      Begin VB.TextBox TxtEvento 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         MaxLength       =   30
         TabIndex        =   75
         Top             =   2520
         Width           =   4125
      End
      Begin VB.CheckBox ChkPzoProduto 
         Caption         =   "&Preço a Prazo por Produto"
         Height          =   315
         Left            =   150
         TabIndex        =   67
         Top             =   900
         Width           =   2355
      End
      Begin VB.CheckBox ChkCompra 
         Caption         =   "&Calcular Preço de Compra"
         Height          =   375
         Left            =   2595
         TabIndex        =   66
         Top             =   870
         Width           =   2190
      End
      Begin VB.CheckBox ChkReceber 
         Caption         =   "Permitir &Receber Conta no Frente"
         Height          =   315
         Left            =   5280
         TabIndex        =   65
         Top             =   855
         Width           =   2745
      End
      Begin VB.CheckBox ChkEditar 
         Caption         =   "Permitir &Alterar Venda no Frente"
         Height          =   315
         Left            =   5280
         TabIndex        =   64
         Top             =   570
         Width           =   2745
      End
      Begin VB.CheckBox ChkPrazo 
         Caption         =   "Preço a Prazo &no Pedido"
         Height          =   315
         Left            =   2595
         TabIndex        =   63
         Top             =   570
         Width           =   2145
      End
      Begin VB.CheckBox ChkFunRural 
         Caption         =   "&FunRural"
         Height          =   315
         Left            =   150
         TabIndex        =   62
         Top             =   570
         Width           =   1125
      End
      Begin VB.CheckBox ChkFinanc 
         Caption         =   "&Interface com Financeiro"
         Height          =   315
         Left            =   150
         TabIndex        =   61
         Top             =   1230
         Width           =   2265
      End
      Begin VB.TextBox TxtBcoFin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   135
         MaxLength       =   255
         TabIndex        =   60
         Top             =   1770
         Width           =   4125
      End
      Begin VB.CheckBox ChkTipoPgto 
         Caption         =   "&Mostrar Tipo Recebimento Caixa"
         Height          =   315
         Left            =   2595
         TabIndex        =   59
         Top             =   1215
         Width           =   2625
      End
      Begin VB.CheckBox ChkParcela 
         Caption         =   "&Obs Automática"
         Height          =   315
         Left            =   5280
         TabIndex        =   58
         Top             =   1140
         Width           =   1950
      End
      Begin VB.CheckBox ChkPrecoCima 
         Caption         =   "&Calcular Preço de Cima p/ baixo"
         Height          =   375
         Left            =   5280
         TabIndex        =   57
         Top             =   1410
         Width           =   2790
      End
      Begin VB.CheckBox ChkTranspNt 
         Caption         =   "&Preenche Dados Nota"
         Height          =   375
         Left            =   8310
         TabIndex        =   56
         Top             =   540
         Width           =   2115
      End
      Begin VB.TextBox TxtTamCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   8310
         MaxLength       =   255
         TabIndex        =   55
         Top             =   1755
         Width           =   1650
      End
      Begin VB.TextBox TxtTamBarra 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   8310
         MaxLength       =   255
         TabIndex        =   54
         Top             =   1170
         Width           =   1650
      End
      Begin VB.CheckBox ChkEstoque 
         Caption         =   "&Manter Histórico Estoque"
         Height          =   375
         Left            =   5280
         TabIndex        =   53
         Top             =   1710
         Width           =   2790
      End
      Begin VB.Data DatDados 
         Caption         =   "CNT"
         Connect         =   "Access"
         DatabaseName    =   "C:\Rede\Sistemas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -70005
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3960
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Frame Frame2 
         Height          =   1470
         Left            =   -74865
         TabIndex        =   38
         Top             =   495
         Width           =   10500
         Begin VB.TextBox TxtEspacamento 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Espacamento"
            DataSource      =   "DatDados(1)"
            Height          =   285
            Left            =   165
            MaxLength       =   3
            TabIndex        =   47
            Top             =   945
            Width           =   1245
         End
         Begin VB.TextBox TxtNota 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "NroNota"
            DataSource      =   "DatDados(1)"
            Height          =   285
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   46
            Top             =   945
            Width           =   1050
         End
         Begin VB.TextBox TxtDetalhe 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "TotalLinhaDetalhe"
            DataSource      =   "DatDados(1)"
            Height          =   285
            Left            =   2895
            MaxLength       =   50
            TabIndex        =   45
            Top             =   945
            Width           =   1335
         End
         Begin VB.CheckBox ChkSimples 
            Caption         =   "&Super Simples"
            Height          =   285
            Left            =   150
            TabIndex        =   44
            Top             =   360
            Width           =   1395
         End
         Begin VB.CheckBox ChkDeducao 
            Caption         =   "&Dedução Icms"
            Height          =   285
            Left            =   1650
            TabIndex        =   43
            Top             =   360
            Width           =   1365
         End
         Begin VB.CheckBox ChkCalcular 
            Caption         =   "&Calcular Icms"
            Height          =   285
            Left            =   4710
            TabIndex        =   42
            Top             =   360
            Width           =   1365
         End
         Begin VB.CheckBox ChkDesconto 
            Caption         =   "&Impr.Desconto"
            Height          =   285
            Left            =   3180
            TabIndex        =   41
            Top             =   360
            Width           =   1335
         End
         Begin VB.ListBox Lista 
            Appearance      =   0  'Flat
            Height          =   1005
            ItemData        =   "FrmConfiguracao.frx":0586
            Left            =   7320
            List            =   "FrmConfiguracao.frx":05DB
            TabIndex        =   40
            Top             =   345
            Width           =   3045
         End
         Begin VB.TextBox TxtLinhaFatura 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "TotalLinhaDetalhe"
            DataSource      =   "DatDados(1)"
            Height          =   285
            Left            =   4440
            MaxLength       =   50
            TabIndex        =   39
            Top             =   945
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&Espaçamento"
            Height          =   195
            Left            =   165
            TabIndex        =   52
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Numero No&ta"
            Height          =   195
            Left            =   1605
            TabIndex        =   51
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Qtd Linha Detalhe"
            Height          =   195
            Left            =   2895
            TabIndex        =   50
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Variaveis da Nota"
            Height          =   195
            Left            =   7320
            TabIndex        =   49
            Top             =   135
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Qtd Linha Fatura"
            Height          =   195
            Left            =   4440
            TabIndex        =   48
            Top             =   720
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3825
         Left            =   -74865
         TabIndex        =   27
         Top             =   1980
         Width           =   10500
         Begin VB.OptionButton OptNota 
            Caption         =   "Promissória"
            Height          =   315
            Index           =   9
            Left            =   7290
            TabIndex        =   71
            Top             =   210
            Width           =   1125
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Diversos"
            Height          =   315
            Index           =   8
            Left            =   9345
            TabIndex        =   36
            Top             =   210
            Width           =   1035
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Salto"
            Height          =   315
            Index           =   7
            Left            =   8520
            TabIndex        =   35
            Top             =   210
            Width           =   720
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Duplicata"
            Height          =   315
            Index           =   6
            Left            =   6165
            TabIndex        =   34
            Top             =   210
            Width           =   1125
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Boleto"
            Height          =   315
            Index           =   5
            Left            =   5265
            TabIndex        =   33
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "FatCompra"
            Height          =   315
            Index           =   4
            Left            =   4065
            TabIndex        =   32
            Top             =   210
            Width           =   1095
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Fat.Venda"
            Height          =   315
            Index           =   3
            Left            =   2940
            TabIndex        =   31
            Top             =   210
            Width           =   1155
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Totais"
            Height          =   315
            Index           =   2
            Left            =   2115
            TabIndex        =   30
            Top             =   210
            Width           =   810
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Detalhe"
            Height          =   315
            Index           =   1
            Left            =   1185
            TabIndex        =   29
            Top             =   210
            Width           =   885
         End
         Begin VB.OptionButton OptNota 
            Caption         =   "Cabeçalho"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   28
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
         Begin MSDBGrid.DBGrid GridNt 
            Bindings        =   "FrmConfiguracao.frx":08A7
            Height          =   3165
            Left            =   90
            OleObjectBlob   =   "FrmConfiguracao.frx":08C1
            TabIndex        =   37
            Top             =   540
            Width           =   10245
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1470
         Left            =   -74910
         TabIndex        =   14
         Top             =   1800
         Width           =   9765
         Begin VB.TextBox TxtCabec 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   180
            MaxLength       =   40
            TabIndex        =   21
            Top             =   405
            Width           =   4005
         End
         Begin VB.TextBox TxtCabec2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   4410
            MaxLength       =   50
            TabIndex        =   20
            Top             =   405
            Width           =   4935
         End
         Begin VB.TextBox TxtImpressora 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   180
            MaxLength       =   50
            TabIndex        =   19
            Top             =   990
            Width           =   2115
         End
         Begin VB.TextBox TxtCasasDec 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   18
            Top             =   990
            Width           =   1485
         End
         Begin VB.CheckBox ChkCustoPed 
            Caption         =   "&Mostrar Custo Pedido"
            Height          =   315
            Left            =   4065
            TabIndex        =   17
            Top             =   975
            Width           =   2025
         End
         Begin VB.TextBox TxtLoja 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   6165
            MaxLength       =   50
            TabIndex        =   16
            Top             =   990
            Width           =   1875
         End
         Begin VB.CheckBox ChkReajusta 
            Caption         =   "&Reajusta Conta"
            Height          =   315
            Left            =   8235
            TabIndex        =   15
            Top             =   990
            Width           =   1470
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Cabeçalho &Principal"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   180
            Width           =   1410
         End
         Begin VB.Label Label25 
            Caption         =   "Cabeçalho &Secundário"
            Height          =   240
            Left            =   4410
            TabIndex        =   25
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "&Porta Impressora"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   765
            Width           =   1185
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&Nro Casas Decimais"
            Height          =   195
            Left            =   2460
            TabIndex        =   23
            Top             =   765
            Width           =   1425
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "&Loja Estoque Automática"
            Height          =   195
            Left            =   6165
            TabIndex        =   22
            Top             =   765
            Width           =   1770
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1470
         Left            =   -74910
         TabIndex        =   8
         Top             =   330
         Width           =   9735
         Begin VB.TextBox TxtMens 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3870
            MaxLength       =   40
            TabIndex        =   4
            ToolTipText     =   "Grava"
            Top             =   975
            Width           =   3555
         End
         Begin VB.CheckBox ChkBoleto 
            Caption         =   "Boleto em Formulário"
            Height          =   285
            Left            =   7830
            TabIndex        =   2
            Top             =   390
            Width           =   1845
         End
         Begin MSMask.MaskEdBox MebCnpj 
            Height          =   285
            Left            =   4410
            TabIndex        =   1
            Top             =   405
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtEmpresa 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   150
            MaxLength       =   40
            TabIndex        =   0
            Top             =   405
            Width           =   4125
         End
         Begin VB.TextBox TxtObsCarne 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   150
            MaxLength       =   40
            TabIndex        =   3
            ToolTipText     =   "Grava"
            Top             =   975
            Width           =   3555
         End
         Begin VB.TextBox TxtDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   7590
            MaxLength       =   50
            TabIndex        =   5
            Top             =   975
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Mensagem Rodapé"
            Height          =   195
            Left            =   3870
            TabIndex        =   13
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label Label13 
            Caption         =   "&Nome da Empresa"
            Height          =   240
            Left            =   150
            TabIndex        =   12
            Top             =   180
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "&CNPJ Empresa"
            Height          =   240
            Left            =   4410
            TabIndex        =   11
            Top             =   180
            Width           =   1590
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "&Observação Carnê"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Desconto &Frente"
            Height          =   195
            Left            =   7590
            TabIndex        =   9
            Top             =   750
            Width           =   1185
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Cartão da Comissão"
         Height          =   195
         Left            =   5265
         TabIndex        =   78
         Top             =   2205
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   " Nome do Evento"
         Height          =   240
         Left            =   90
         TabIndex        =   76
         Top             =   2295
         Width           =   2730
      End
      Begin VB.Label Label6 
         Caption         =   "&Caminho Bco Dados Financeiro"
         Height          =   240
         Left            =   135
         TabIndex        =   70
         Top             =   1545
         Width           =   2730
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Tamanho Cod.Produto"
         Height          =   195
         Left            =   8310
         TabIndex        =   69
         Top             =   1530
         Width           =   1605
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Tamanho Min.Cod.Barra"
         Height          =   195
         Left            =   8310
         TabIndex        =   68
         Top             =   945
         Width           =   1725
      End
   End
End
Attribute VB_Name = "FrmConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecCfg As Recordset

Private Sub BtOK_Click()

RecCfg.Edit
RecCfg!NomeEmpresa = TxtEmpresa
RecCfg!obsparcelas = TxtObsCarne
RecCfg!MensagemRodape = TxtMens
RecCfg!Cabecalho1 = TxtCabec
RecCfg!Cabecalho2 = TxtCabec2
RecCfg!Porta = TxtImpressora
RecCfg!Espacamento = TxtEspacamento
RecCfg!NroNota = TxtNota
RecCfg!TotalLinhaDetalhe = TxtDetalhe
RecCfg.Update

If ChkBoleto.value = 1 Then
    EscreveINI "BOLETO", "Tipo", "FORMULARIO", App.Path & "\config.ini"
Else
    EscreveINI "BOLETO", "Tipo", "A4", App.Path & "\config.ini"
End If
EscreveINI "Config", "Cnpj", MebCnpj, App.Path & "\config.ini"
EscreveINI "Config", "Empresa", TxtEmpresa, App.Path & "\config.ini"
EscreveINI "DescontoFrente", "Valor", TxtDesc, App.Path & "\config.ini"
EscreveINI "Pedido", "FunRural", ChkFunRural.value, App.Path & "\config.ini"
EscreveINI "Pedido", "PcoPrazo", chkprazo.value, App.Path & "\config.ini"
EscreveINI "Pedido", "PcoCompra", ChkCompra.value, App.Path & "\config.ini"
EscreveINI "Pedido", "PzoProduto", ChkPzoProduto.value, App.Path & "\config.ini"
EscreveINI "Nota", "Simples", ChkSimples.value, App.Path & "\config.ini"
EscreveINI "Nota", "Deducao", ChkDeducao.value, App.Path & "\config.ini"
EscreveINI "Nota", "Desconto", ChkDesconto.value, App.Path & "\config.ini"
EscreveINI "Nota", "CalculaIcms", ChkCalcular.value, App.Path & "\config.ini"
EscreveINI "Nota", "QtdLinhaFatura", TxtLinhaFatura.Text, App.Path & "\config.ini"
EscreveINI "DescontoFrente", "Editar", ChkEditar.value, App.Path & "\config.ini"
EscreveINI "DescontoFrente", "Receber", ChkReceber.value, App.Path & "\config.ini"
EscreveINI "Financeiro", "Comunica", ChkFinanc.value, App.Path & "\config.ini"
EscreveINI "Financeiro", "LocalBcoFin", TxtBcoFin.Text, App.Path & "\config.ini"

EscreveINI "Pedido", "CasasDec", TxtCasasDec.Text, App.Path & "\config.ini"
EscreveINI "Pedido", "MostrarCusto", ChkCustoPed.value, App.Path & "\config.ini"
EscreveINI "Produto", "LocalEstoque", TxtLoja.Text, App.Path & "\config.ini"
EscreveINI "Pedido", "MostraTipoPgtoCaixa", ChkTipoPgto.value, App.Path & "\config.ini"
EscreveINI "Pedido", "ObsAutomatica", ChkParcela.value, App.Path & "\config.ini"
EscreveINI "Pedido", "PrecoCima", ChkPrecoCima.value, App.Path & "\config.ini"
EscreveINI "Nota", "TranspNota", ChkTranspNt.value, App.Path & "\config.ini"
EscreveINI "Produto", "NroCod", TxtTamCod, App.Path & "\config.ini"
EscreveINI "Produto", "TamBarra", TxtTamBarra, App.Path & "\config.ini"

EscreveINI "Config", "Evento", TxtEvento, App.Path & "\config.ini"


EscreveINI "Pedido", "ReajustaConta", ChkReajusta, App.Path & "\config.ini"
EscreveINI "Produto", "HistEstoque", ChkEstoque, App.Path & "\config.ini"

EscreveINI "Parceiro", "TaxaCartao", TxtTaxa, App.Path & "\config.ini"


SSTab1.Tab = 0
End Sub

Private Sub ChkPrazo_LostFocus()
SSTab1.Tab = 2
End Sub



Private Sub Form_Activate()
Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyPress = 0
    End If
End Sub


Private Sub Form_Load()
Dim Boleto As String
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next
OptNota_Click 0

Set RecCfg = Banco.OpenRecordset("Select * From Cfg")

TxtEmpresa = RecCfg!NomeEmpresa
MebCnpj = LerINI("Config", "Cnpj", App.Path & "\config.ini")
Boleto = LerINI("BOLETO", "Tipo", App.Path & "\config.ini")
If Boleto = "FORMULARIO" Then
    ChkBoleto.value = 1
Else
    ChkBoleto.value = 0
End If
If Not IsNull(RecCfg!obsparcelas) Then
    TxtObsCarne = RecCfg!obsparcelas
End If
TxtMens = RecCfg!MensagemRodape
TxtDesc = LerINI("DescontoFrente", "Valor", App.Path & "\config.ini")
TxtCabec = RecCfg!Cabecalho1
TxtCabec2 = RecCfg!Cabecalho2
TxtImpressora = RecCfg!Porta
ChkFunRural.value = LerINI("Pedido", "FunRural", App.Path & "\config.ini")
chkprazo.value = LerINI("Pedido", "PcoPrazo", App.Path & "\config.ini")
If LerINI("Pedido", "PcoCompra", App.Path & "\config.ini") <> "" Then
    ChkCompra.value = LerINI("Pedido", "PcoCompra", App.Path & "\config.ini")
Else
    ChkCompra.value = 0
End If

If LerINI("Pedido", "PzoProduto", App.Path & "\config.ini") <> "" Then
    ChkPzoProduto.value = LerINI("Pedido", "PzoProduto", App.Path & "\config.ini")
Else
    ChkPzoProduto.value = 0
End If

ChkSimples.value = LerINI("Nota", "Simples", App.Path & "\config.ini")
ChkDeducao.value = LerINI("Nota", "Deducao", App.Path & "\config.ini")
ChkDesconto.value = LerINI("Nota", "Desconto", App.Path & "\config.ini")
ChkCalcular.value = LerINI("Nota", "CalculaIcms", App.Path & "\config.ini")
TxtLinhaFatura.Text = LerINI("Nota", "QtdLinhaFatura", App.Path & "\config.ini")

TxtEspacamento = RecCfg!Espacamento
If Not IsNull(RecCfg!NroNota) Then
    TxtNota = RecCfg!NroNota
    TxtDetalhe = RecCfg!TotalLinhaDetalhe
End If

If LerINI("DescontoFrente", "Editar", App.Path & "\config.ini") <> "" Then
    ChkEditar.value = LerINI("DescontoFrente", "Editar", App.Path & "\config.ini")
Else
    ChkEditar.value = 0
End If

If LerINI("DescontoFrente", "Receber", App.Path & "\config.ini") <> "" Then
    ChkReceber.value = LerINI("DescontoFrente", "Receber", App.Path & "\config.ini")
Else
    ChkReceber.value = 0
End If

If LerINI("Financeiro", "Comunica", App.Path & "\config.ini") <> "" Then
    ChkFinanc.value = LerINI("Financeiro", "Comunica", App.Path & "\config.ini")
Else
    ChkFinanc.value = 0
End If
TxtBcoFin.Text = LerINI("Financeiro", "LocalBcoFin", App.Path & "\config.ini")

TxtCasasDec.Text = LerINI("Pedido", "CasasDec", App.Path & "\config.ini")
If LerINI("Pedido", "MostrarCusto", App.Path & "\config.ini") <> "" Then
    ChkCustoPed.value = LerINI("Pedido", "MostrarCusto", App.Path & "\config.ini")
Else
    ChkCustoPed.value = 0
End If

TxtLoja.Text = LerINI("Produto", "LocalEstoque", App.Path & "\config.ini")

If LerINI("Pedido", "MostraTipoPgtoCaixa", App.Path & "\config.ini") <> "" Then
    ChkTipoPgto.value = LerINI("Pedido", "MostraTipoPgtoCaixa", App.Path & "\config.ini")
Else
    ChkTipoPgto.value = 0
End If

If LerINI("Pedido", "ObsAutomatica", App.Path & "\config.ini") <> "" Then
    ChkParcela.value = LerINI("Pedido", "ObsAutomatica", App.Path & "\config.ini")
Else
    ChkParcela.value = 0
End If

If LerINI("Pedido", "PrecoCima", App.Path & "\config.ini") <> "" Then
    ChkPrecoCima.value = LerINI("Pedido", "PrecoCima", App.Path & "\config.ini")
Else
    ChkPrecoCima.value = 0
End If

If LerINI("Nota", "TranspNota", App.Path & "\config.ini") <> "" Then
    ChkTranspNt.value = LerINI("Nota", "TranspNota", App.Path & "\config.ini")
Else
    ChkTranspNt.value = 0
End If

If LerINI("Produto", "NroCodigo", App.Path & "\config.ini") <> "" Then
    TxtTamCod = LerINI("Produto", "NroCodigo", App.Path & "\config.ini")
Else
    TxtTamCod = 8
End If

If LerINI("Produto", "TamBarra", App.Path & "\config.ini") <> "" Then
    TxtTamBarra = LerINI("Produto", "TamBarra", App.Path & "\config.ini")
Else
    TxtTamBarra = 8
End If

If LerINI("Pedido", "ReajustaConta", App.Path & "\config.ini") <> "" Then
    ChkReajusta = LerINI("Pedido", "ReajustaConta", App.Path & "\config.ini")
Else
    ChkReajusta = 0
End If

If LerINI("Produto", "HistEstoque", App.Path & "\config.ini") <> "" Then
    ChkEstoque = LerINI("Produto", "HistEstoque", App.Path & "\config.ini")
Else
    ChkEstoque = 0
End If

If LerINI("Config", "Evento", App.Path & "\config.ini") <> "" Then
    TxtEvento = LerINI("Config", "Evento", App.Path & "\config.ini")
Else
    TxtEvento = ""
End If

If LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini") <> "" Then
    TxtTaxa = LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini")
Else
    TxtEvento = ""
End If




SSTab1.Tab = 0

End Sub

Private Sub OptNota_Click(Index As Integer)
Dim Valor As Byte
If Index <= 2 Then
    Valor = Index + 1
End If

If Index > 2 And Index <= 4 Then
    Valor = Index + 2
End If

If Index = 5 Then
    Valor = 10
End If

If Index = 6 Then
    Valor = 15
End If

If Index = 7 Then
    Valor = 4
End If

If Index = 8 Then
    Valor = 99
End If

If Index = 9 Then
    Valor = 11
End If

DatDados(0).RecordSource = "Select * From CNT Where TipoDados =  " & Valor & " Order by Codigo"
DatDados(0).Refresh

End Sub


Private Sub TxtDesc_LostFocus()
SSTab1.Tab = 1
End Sub
