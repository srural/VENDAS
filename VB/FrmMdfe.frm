VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmMdfe 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14085
   Icon            =   "FrmMdfe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14085
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   14205
      TabIndex        =   41
      Top             =   8145
      Width           =   14235
      Begin VB.Label LblMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   42
         Top             =   45
         Width           =   8460
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14100
      TabIndex        =   37
      Top             =   0
      Width           =   14100
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
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   7065
         TabIndex        =   39
         Top             =   90
         Width           =   6270
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Mdf-e"
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
         TabIndex        =   38
         Top             =   45
         Width           =   5190
      End
   End
   Begin TabDlg.SSTab tabCad 
      Height          =   7545
      Left            =   0
      TabIndex        =   14
      Top             =   570
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   13309
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmMdfe.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdEncerrar"
      Tab(0).Control(1)=   "DatCadastro"
      Tab(0).Control(2)=   "DatDados(0)"
      Tab(0).Control(3)=   "DBGrid1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Identificação"
      TabPicture(1)   =   "FrmMdfe.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PctFundo1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Destino"
      TabPicture(2)   =   "FrmMdfe.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "PctFundo2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton CmdEncerrar 
         Caption         =   "Encerrar Mdf-e"
         Height          =   330
         Left            =   -62805
         TabIndex        =   40
         Top             =   7020
         Width           =   1725
      End
      Begin VB.Data DatCadastro 
         Caption         =   "Cadastro"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -69690
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5580
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.PictureBox PctFundo1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7155
         Left            =   -74910
         ScaleHeight     =   7155
         ScaleWidth      =   13770
         TabIndex        =   17
         Top             =   360
         Width           =   13770
         Begin VB.ComboBox CmbUnidade 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmMdfe.frx":0060
            Left            =   9540
            List            =   "FrmMdfe.frx":006A
            TabIndex        =   68
            Top             =   390
            Width           =   1440
         End
         Begin VB.ComboBox CmbTipo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmMdfe.frx":0084
            Left            =   7125
            List            =   "FrmMdfe.frx":008E
            TabIndex        =   52
            Top             =   390
            Width           =   2295
         End
         Begin VB.Frame Frame3 
            Caption         =   "Motorista"
            Height          =   2565
            Left            =   45
            TabIndex        =   47
            Top             =   855
            Width           =   10995
            Begin VB.TextBox TxtCiot 
               Appearance      =   0  'Flat
               DataField       =   "Bairro_Entrega"
               Height          =   315
               Left            =   8685
               MaxLength       =   20
               TabIndex        =   102
               Tag             =   "S"
               Top             =   1185
               Width           =   1680
            End
            Begin VB.TextBox TextUfVeiculo 
               Appearance      =   0  'Flat
               DataField       =   "Nro_Entrega"
               Height          =   315
               Left            =   7605
               MaxLength       =   40
               TabIndex        =   66
               Tag             =   "S"
               Top             =   1170
               Width           =   420
            End
            Begin VB.TextBox TxtPlaca 
               Appearance      =   0  'Flat
               DataField       =   "Nro_Entrega"
               Height          =   315
               Left            =   6315
               MaxLength       =   40
               TabIndex        =   63
               Tag             =   "S"
               Top             =   1170
               Width           =   1185
            End
            Begin VB.TextBox TxtTara 
               Appearance      =   0  'Flat
               DataField       =   "Bairro_Entrega"
               Height          =   315
               Left            =   6315
               MaxLength       =   20
               TabIndex        =   62
               Tag             =   "S"
               Top             =   1920
               Width           =   1680
            End
            Begin VB.ComboBox CmbRodado 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "FrmMdfe.frx":00DD
               Left            =   3165
               List            =   "FrmMdfe.frx":00F3
               TabIndex        =   60
               Top             =   1950
               Width           =   2565
            End
            Begin VB.ComboBox CmbCarroceria 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "FrmMdfe.frx":012E
               Left            =   240
               List            =   "FrmMdfe.frx":0144
               TabIndex        =   58
               Top             =   1950
               Width           =   2565
            End
            Begin VB.ComboBox CmbTipoVeiculo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "FrmMdfe.frx":018F
               Left            =   3165
               List            =   "FrmMdfe.frx":0199
               TabIndex        =   56
               Top             =   1185
               Width           =   2565
            End
            Begin VB.ComboBox CmbVeiculo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "FrmMdfe.frx":01AE
               Left            =   240
               List            =   "FrmMdfe.frx":01B8
               TabIndex        =   54
               Top             =   1185
               Width           =   2565
            End
            Begin VB.TextBox Motorista 
               Appearance      =   0  'Flat
               DataField       =   "Endereço_Entrega"
               Height          =   315
               Left            =   240
               MaxLength       =   50
               TabIndex        =   48
               Tag             =   "S"
               Top             =   480
               Width           =   5895
            End
            Begin MSMask.MaskEdBox Cpf 
               Height          =   315
               Index           =   1
               Left            =   6315
               TabIndex        =   49
               Top             =   480
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "###-###-###-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox CpfCiot 
               Height          =   315
               Index           =   0
               Left            =   8685
               TabIndex        =   100
               Top             =   1950
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "###-###-###-##"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Ciot:"
               Height          =   195
               Index           =   13
               Left            =   8745
               TabIndex        =   103
               Top             =   945
               Width           =   315
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cpf resp. Ciot"
               Height          =   195
               Index           =   7
               Left            =   8655
               TabIndex        =   101
               Top             =   1710
               Width           =   945
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Uf:"
               Height          =   195
               Index           =   11
               Left            =   7620
               TabIndex        =   67
               Top             =   945
               Width           =   210
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Placa:"
               Height          =   195
               Index           =   10
               Left            =   6315
               TabIndex        =   65
               Top             =   945
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tara:"
               Height          =   195
               Index           =   9
               Left            =   6315
               TabIndex        =   64
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Rodado"
               Height          =   195
               Index           =   8
               Left            =   3165
               TabIndex        =   61
               Top             =   1755
               Width           =   1155
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Carroceria"
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   59
               Top             =   1710
               Width           =   720
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tipo Veiculo"
               Height          =   195
               Index           =   3
               Left            =   3165
               TabIndex        =   57
               Top             =   945
               Width           =   885
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Veiculo"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   55
               Top             =   945
               Width           =   525
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Nome:"
               Height          =   195
               Index           =   6
               Left            =   240
               TabIndex        =   51
               Top             =   240
               Width           =   465
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cpf:"
               Height          =   195
               Index           =   5
               Left            =   6315
               TabIndex        =   50
               Top             =   240
               Width           =   285
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   ">>>"
            Height          =   330
            Left            =   5445
            TabIndex        =   46
            Top             =   6165
            Width           =   735
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "NomeVendedor"
            Height          =   315
            Index           =   1
            Left            =   45
            MaxLength       =   40
            TabIndex        =   45
            Tag             =   "S"
            Top             =   6165
            Width           =   5325
         End
         Begin VB.ComboBox CmbModo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmMdfe.frx":01CF
            Left            =   4560
            List            =   "FrmMdfe.frx":01DC
            TabIndex        =   0
            Top             =   390
            Width           =   2340
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "CodEntidade"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   0
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "N"
            Top             =   390
            Width           =   1605
         End
         Begin VB.Data DatDados 
            Caption         =   "Convenio"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   4
            Left            =   7830
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   6300
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Data DatDados 
            Caption         =   "Time"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   1
            Left            =   10035
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   6300
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Frame Frame6 
            Caption         =   "Documentos"
            Height          =   2580
            Left            =   45
            TabIndex        =   18
            Top             =   3465
            Width           =   13650
         End
         Begin MSMask.MaskEdBox MebData 
            DataField       =   "DtCadastro"
            Height          =   315
            Left            =   1845
            TabIndex        =   1
            Tag             =   "S"
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox HoraSaida 
            Height          =   315
            Left            =   3285
            TabIndex        =   43
            Top             =   390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Unidade de Medida"
            Height          =   195
            Index           =   24
            Left            =   9525
            TabIndex        =   69
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo Emitente"
            Height          =   195
            Index           =   1
            Left            =   7080
            TabIndex        =   53
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            Height          =   195
            Left            =   3285
            TabIndex        =   44
            Top             =   180
            Width           =   345
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Index           =   61
            Left            =   1875
            TabIndex        =   22
            Top             =   180
            Width           =   390
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Modo de Transporte"
            Height          =   195
            Index           =   62
            Left            =   4545
            TabIndex        =   21
            Top             =   180
            Width           =   1440
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   180
            Width           =   540
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Mdfe"
         Connect         =   "Access"
         DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -69330
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From Mdfe Order by Data Desc"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.PictureBox PctFundo2 
         BorderStyle     =   0  'None
         Height          =   7035
         Left            =   225
         ScaleHeight     =   7035
         ScaleWidth      =   13635
         TabIndex        =   15
         Top             =   405
         Width           =   13635
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            DataField       =   "Obs"
            Height          =   2325
            Index           =   40
            Left            =   7695
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Tag             =   "S"
            Top             =   3825
            Width           =   3765
         End
         Begin VB.Frame Frame4 
            Caption         =   "Uf do Percurso"
            Height          =   2490
            Left            =   90
            TabIndex        =   70
            Top             =   3690
            Width           =   7395
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "CE"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   26
               Left            =   6525
               TabIndex        =   99
               Top             =   1350
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "MA"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   25
               Left            =   6525
               TabIndex        =   98
               Top             =   1710
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "RN"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   24
               Left            =   6525
               TabIndex        =   97
               Top             =   990
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "PB"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   23
               Left            =   6525
               TabIndex        =   96
               Top             =   630
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "PI"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   22
               Left            =   5850
               TabIndex        =   95
               Top             =   1935
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "PE"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   21
               Left            =   5850
               TabIndex        =   94
               Top             =   1620
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "ES"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   20
               Left            =   1395
               TabIndex        =   93
               Top             =   1755
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "SE"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   19
               Left            =   5850
               TabIndex        =   92
               Top             =   990
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "AL"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   18
               Left            =   5850
               TabIndex        =   91
               Top             =   1305
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "DF"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   17
               Left            =   2655
               TabIndex        =   90
               Top             =   1755
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "BA"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   16
               Left            =   5850
               TabIndex        =   89
               Top             =   630
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "AC"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   15
               Left            =   4635
               TabIndex        =   88
               Top             =   1350
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "AM"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   14
               Left            =   3870
               TabIndex        =   87
               Top             =   1710
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "RR"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   13
               Left            =   4635
               TabIndex        =   86
               Top             =   675
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "RO"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   12
               Left            =   4635
               TabIndex        =   85
               Top             =   1035
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "AP"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   11
               Left            =   3870
               TabIndex        =   84
               Top             =   1395
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "PA"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   10
               Left            =   3870
               TabIndex        =   83
               Top             =   1035
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "TO"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   9
               Left            =   3870
               TabIndex        =   82
               Top             =   675
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "GO"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   8
               Left            =   2655
               TabIndex        =   81
               Top             =   1395
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "MT"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   7
               Left            =   2655
               TabIndex        =   80
               Top             =   1035
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "MS"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   6
               Left            =   2655
               TabIndex        =   79
               Top             =   675
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "RJ"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   5
               Left            =   1395
               TabIndex        =   78
               Top             =   1395
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "MS"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   4
               Left            =   1395
               TabIndex        =   77
               Top             =   1035
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "SP"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   3
               Left            =   1395
               TabIndex        =   76
               Top             =   675
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "PR"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   2
               Left            =   225
               TabIndex        =   75
               Top             =   1395
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "SC"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   1
               Left            =   225
               TabIndex        =   74
               Top             =   1035
               Width           =   555
            End
            Begin VB.CheckBox ChkUf 
               Appearance      =   0  'Flat
               Caption         =   "RS"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   225
               TabIndex        =   73
               Top             =   675
               Width           =   555
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "NORDESTE"
               Height          =   195
               Index           =   27
               Left            =   5895
               TabIndex        =   108
               Top             =   315
               Width           =   900
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "NORTE"
               Height          =   195
               Index           =   26
               Left            =   4200
               TabIndex        =   107
               Top             =   315
               Width           =   570
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CENTRO-OESTE"
               Height          =   195
               Index           =   25
               Left            =   2250
               TabIndex        =   106
               Top             =   315
               Width           =   1260
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "SUDESTE"
               Height          =   195
               Index           =   16
               Left            =   1260
               TabIndex        =   105
               Top             =   315
               Width           =   765
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "SUL"
               Height          =   195
               Index           =   14
               Left            =   450
               TabIndex        =   104
               Top             =   315
               Width           =   315
            End
            Begin VB.Line Line4 
               X1              =   5490
               X2              =   5490
               Y1              =   630
               Y2              =   2475
            End
            Begin VB.Line Line3 
               X1              =   3600
               X2              =   3600
               Y1              =   630
               Y2              =   2475
            End
            Begin VB.Line Line2 
               X1              =   2205
               X2              =   2205
               Y1              =   630
               Y2              =   2475
            End
            Begin VB.Line Line1 
               X1              =   1035
               X2              =   1035
               Y1              =   630
               Y2              =   2475
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Destino"
            Height          =   1530
            Left            =   45
            TabIndex        =   30
            Top             =   1800
            Width           =   11400
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Bairro_Cob"
               Height          =   315
               Index           =   2
               Left            =   7680
               MaxLength       =   20
               TabIndex        =   111
               Tag             =   "S"
               Top             =   1080
               Width           =   1365
            End
            Begin VB.CommandButton Command3 
               Caption         =   ">>>"
               Height          =   330
               Left            =   10350
               TabIndex        =   72
               Top             =   1080
               Width           =   555
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Endereço_Cob"
               Height          =   315
               Index           =   31
               Left            =   240
               MaxLength       =   50
               TabIndex        =   8
               Tag             =   "S"
               Top             =   450
               Width           =   5895
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Nro_Cob"
               Height          =   315
               Index           =   30
               Left            =   6330
               MaxLength       =   40
               TabIndex        =   9
               Tag             =   "S"
               Top             =   480
               Width           =   1185
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Bairro_Cob"
               Height          =   315
               Index           =   29
               Left            =   7665
               MaxLength       =   20
               TabIndex        =   10
               Tag             =   "S"
               Top             =   495
               Width           =   3210
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Cidade_Cob"
               Height          =   315
               Index           =   28
               Left            =   225
               MaxLength       =   20
               TabIndex        =   11
               Tag             =   "S"
               Top             =   1080
               Width           =   4920
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Uf_Cob"
               Height          =   315
               Index           =   8
               Left            =   5250
               MaxLength       =   2
               TabIndex        =   12
               Tag             =   "S"
               Top             =   1080
               Width           =   765
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Cob"
               Height          =   315
               Index           =   5
               Left            =   6330
               TabIndex        =   13
               Top             =   1080
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Index           =   29
               Left            =   7695
               TabIndex        =   112
               Top             =   855
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Index           =   50
               Left            =   240
               TabIndex        =   36
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Numero:"
               Height          =   195
               Index           =   49
               Left            =   6330
               TabIndex        =   35
               Top             =   240
               Width           =   600
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Index           =   48
               Left            =   7650
               TabIndex        =   34
               Top             =   255
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Index           =   17
               Left            =   225
               TabIndex        =   33
               Top             =   840
               Width           =   540
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Index           =   15
               Left            =   5220
               TabIndex        =   32
               Top             =   870
               Width           =   255
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Index           =   12
               Left            =   6330
               TabIndex        =   31
               Top             =   840
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Origem"
            Height          =   1530
            Left            =   45
            TabIndex        =   23
            Top             =   135
            Width           =   11400
            Begin VB.CommandButton Command2 
               Caption         =   ">>>"
               Height          =   330
               Left            =   9585
               TabIndex        =   71
               Top             =   1035
               Width           =   555
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Uf_Entrega"
               Height          =   315
               Index           =   9
               Left            =   5280
               MaxLength       =   2
               TabIndex        =   6
               Tag             =   "S"
               Top             =   1080
               Width           =   720
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Cidade_Entrega"
               Height          =   315
               Index           =   10
               Left            =   225
               MaxLength       =   20
               TabIndex        =   5
               Tag             =   "S"
               Top             =   1080
               Width           =   4935
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Bairro_Entrega"
               Height          =   315
               Index           =   11
               Left            =   7575
               MaxLength       =   20
               TabIndex        =   4
               Tag             =   "S"
               Top             =   495
               Width           =   3300
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Nro_Entrega"
               Height          =   315
               Index           =   12
               Left            =   6270
               MaxLength       =   40
               TabIndex        =   3
               Tag             =   "S"
               Top             =   480
               Width           =   1185
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Endereço_Entrega"
               Height          =   315
               Index           =   13
               Left            =   240
               MaxLength       =   50
               TabIndex        =   2
               Tag             =   "S"
               Top             =   480
               Width           =   5895
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Entrega"
               Height          =   315
               Index           =   8
               Left            =   6300
               TabIndex        =   7
               Top             =   1080
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Index           =   18
               Left            =   6300
               TabIndex        =   29
               Top             =   840
               Width           =   315
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Index           =   19
               Left            =   5280
               TabIndex        =   28
               Top             =   840
               Width           =   255
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Index           =   20
               Left            =   225
               TabIndex        =   27
               Top             =   840
               Width           =   540
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Index           =   21
               Left            =   7575
               TabIndex        =   26
               Top             =   255
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Numero:"
               Height          =   195
               Index           =   22
               Left            =   6270
               TabIndex        =   25
               Top             =   240
               Width           =   600
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Index           =   23
               Left            =   240
               TabIndex        =   24
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Obs:"
            Height          =   195
            Index           =   28
            Left            =   7695
            TabIndex        =   110
            Top             =   3600
            Width           =   330
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmMdfe.frx":0204
         Height          =   6330
         Left            =   -74865
         OleObjectBlob   =   "FrmMdfe.frx":021E
         TabIndex        =   16
         Top             =   540
         Width           =   13755
      End
   End
End
Attribute VB_Name = "FrmMdfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset   'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Dim FlagAbertura As Boolean
Dim Cliente As Integer
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

Public SelectFromSQL  As String  ' Representa o Conteudo das Cláusulas Select e From do Sql Utilizado p/ acessar os registros deste formulário
Public WhereSQL As String        ' Representa o Conteudo da Cláusula Where do Sql Utilizado p/ acessar os registros deste formulário
Public FiltroSQL As String       ' Critério imposto pelo usuário
Public OrderBySQL As String      ' Representa a cláusula Order By do Sql do formulário Ativo
Public PesquisaSql As String     ' Representa a pesquisa F2
Public SelectFromFiltro As String  ' Representa a relação de campos que sera exibida na visualização no formulário de filtragem

Public LocalizacaoBasica As String ' Armazena o nome do campo que sera utilizado na ferramenta de localizaco basica
Public VarPrazo As String


Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodEntidade = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
tabCad.Tab = 1 'direcina para o TAB seguinte
If Registros.RecordCount > 0 Then
    MostraDados
    
End If


End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário


Formulário.Left = 100 '(MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = 50 '(MDIPrincipal.Height - Formulário.Height) / 6


tabCad.Tab = 0


End Sub


Private Sub Form_Load()
Dim Componentes As Variant
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
    
Set Registros = Banco.OpenRecordset("SELECT * From Ent Order by Nome")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select * From Mdfe Order by Data Desc "

CmbVeiculo.ListIndex = 0
CmbTipoVeiculo.ListIndex = 0
CmbCarroceria.ListIndex = 1
CmbRodado.ListIndex = 0
CmbModo.ListIndex = 0
CmbTipo.ListIndex = 0
CmbUnidade.ListIndex = 0
MebData = Date
HoraSaida = Time


'DtData1 = Date
'DtData2 = Date

End Sub

Private Sub Txtdados_Change(Index As Integer)
Dim SQL As String
Dim Pago As Currency
Dim Apagar As Currency
Pago = 0
Apagar = 0
If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
     DatDados(5).RecordSource = "SELECT *  FROM Cto WHERE Entidade = " & TxtDados(0) & " Order by CodCto Desc"
    If IsNumeric(TxtDados(0)) = False Then
         DatDados(5).RecordSource = "SELECT *  FROM Cto WHERE Entidade = 0 Order by CodCto Desc"
    Else
         DatDados(5).RecordSource = "SELECT *  FROM Cto WHERE Entidade = " & TxtDados(0) & " Order by CodCto Desc"
    End If
    DatDados(5).Refresh
End If


End Sub



