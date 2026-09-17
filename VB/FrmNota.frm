VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmNota 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   Icon            =   "FrmNota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14055
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000002&
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14040
      TabIndex        =   113
      Top             =   0
      Width           =   14100
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   7800
         TabIndex        =   150
         Top             =   120
         Width           =   5925
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão de Nota Fiscal"
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
         TabIndex        =   114
         Top             =   45
         Width           =   6045
      End
   End
   Begin TabDlg.SSTab SSTabNota 
      Height          =   8130
      Left            =   0
      TabIndex        =   17
      Top             =   585
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   14340
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa Nota"
      TabPicture(0)   =   "FrmNota.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label27"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label30"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbltam"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DtData"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DBGrid"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DatPed"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "BtAtualiza"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "BtNfe(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "BtNfe(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "BtNfe(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtCancela"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtNroNota"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "BtNfe(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Nro1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Nro2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "BtNfe(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "BBtStatus"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtSeq"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Picture1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "BtNfe(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ChkVersao"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "BtNfe(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Comstatus"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Emissão Nota"
      TabPicture(1)   =   "FrmNota.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label32"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "QtdChave"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "MebNfe"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "MebNota"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CodPed"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "DbcCfo"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtCfop"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "DatDados(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame5"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "MebDocumento"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "ChkDoc"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Frame"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TxtChave"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "DatMinutos"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Visualiza"
      TabPicture(2)   =   "FrmNota.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LstNota"
      Tab(2).Control(1)=   "Picture3"
      Tab(2).ControlCount=   2
      Begin MSComDlg.CommonDialog Comstatus 
         Left            =   -67260
         Top             =   6435
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "Baixar Nfe"
         Height          =   330
         Index           =   6
         Left            =   -71175
         TabIndex        =   132
         Top             =   6165
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CheckBox ChkVersao 
         Caption         =   "4.0"
         Height          =   330
         Left            =   -70410
         TabIndex        =   131
         Top             =   7155
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "Gera Pdf"
         Height          =   330
         Index           =   5
         Left            =   -67350
         TabIndex        =   130
         Top             =   7155
         Width           =   1230
      End
      Begin VB.Data DatMinutos 
         Caption         =   "Minutos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5985
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   540
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74955
         ScaleHeight     =   375
         ScaleWidth      =   13965
         TabIndex        =   124
         Top             =   7650
         Width           =   13965
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -75000
         ScaleHeight     =   375
         ScaleWidth      =   13965
         TabIndex        =   123
         Top             =   7605
         Width           =   13965
      End
      Begin VB.TextBox TxtChave 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   180
         TabIndex        =   121
         Top             =   7020
         Width           =   7665
      End
      Begin VB.TextBox TxtSeq 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -70995
         TabIndex        =   117
         Text            =   "1"
         Top             =   7155
         Width           =   510
      End
      Begin VB.CommandButton BBtStatus 
         Caption         =   "Status Nota"
         Height          =   330
         Left            =   -64515
         TabIndex        =   116
         Top             =   6705
         Width           =   1275
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "Carta de Correção"
         Height          =   330
         Index           =   4
         Left            =   -72705
         TabIndex        =   115
         Top             =   7155
         Width           =   1635
      End
      Begin VB.TextBox Nro2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68790
         TabIndex        =   111
         Top             =   6705
         Width           =   780
      End
      Begin VB.TextBox Nro1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -69690
         TabIndex        =   110
         Top             =   6705
         Width           =   780
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "Inutilização de Numeros"
         Height          =   330
         Index           =   3
         Left            =   -69690
         TabIndex        =   109
         Top             =   7155
         Width           =   2085
      End
      Begin VB.TextBox TxtNroNota 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -64515
         TabIndex        =   107
         Top             =   7200
         Width           =   960
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data"
         Height          =   555
         Left            =   -63075
         TabIndex        =   104
         Top             =   6525
         Width           =   1905
         Begin VB.OptionButton Optdata 
            Caption         =   "Nota"
            Height          =   240
            Index           =   1
            Left            =   1080
            TabIndex        =   106
            Top             =   225
            Width           =   735
         End
         Begin VB.OptionButton Optdata 
            Caption         =   "Pedido"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   105
            Top             =   225
            Value           =   -1  'True
            Width           =   870
         End
      End
      Begin VB.TextBox TxtCancela 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74820
         TabIndex        =   98
         Top             =   6705
         Width           =   4920
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "Cancela Nfe já transmitida"
         Height          =   330
         Index           =   2
         Left            =   -74775
         TabIndex        =   95
         Top             =   7155
         Width           =   2040
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "enviar e-mail xml"
         Height          =   330
         Index           =   1
         Left            =   -66135
         TabIndex        =   93
         Top             =   6705
         Width           =   1455
      End
      Begin VB.CommandButton BtNfe 
         Caption         =   "2 via nota eletrônica"
         Height          =   330
         Index           =   0
         Left            =   -67890
         TabIndex        =   90
         Top             =   6705
         Width           =   1635
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   45
         TabIndex        =   88
         Top             =   7605
         Width           =   13965
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Aguarde . . . "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Left            =   240
            TabIndex        =   89
            Top             =   0
            Visible         =   0   'False
            Width           =   13560
         End
      End
      Begin VB.CheckBox ChkDoc 
         Caption         =   "Consumidor Final"
         Height          =   285
         Left            =   6690
         TabIndex        =   82
         Top             =   2550
         Width           =   1605
      End
      Begin MSMask.MaskEdBox MebDocumento 
         Height          =   315
         Left            =   4080
         TabIndex        =   81
         Top             =   2520
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
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
      Begin VB.Frame Frame3 
         Height          =   2430
         Left            =   120
         TabIndex        =   55
         Top             =   4245
         Width           =   8310
         Begin VB.Data DatDados 
            Caption         =   "Transport"
            Connect         =   "Access"
            DatabaseName    =   "C:\Rede\Sistemas\Vendas\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   1
            Left            =   2745
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Select CodEntidade, Nome, Endereco, Nro, Cidade, Uf, Cpf, Cgc, InscrEst  From Ent Where Tipo = 4 or Tipo = 6 Order by Nome"
            Top             =   225
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Frame Frame4 
            Caption         =   "Pago pelo"
            Height          =   600
            Left            =   180
            TabIndex        =   58
            Top             =   540
            Width           =   3075
            Begin VB.OptionButton OptRem 
               Caption         =   "&Remetente"
               Height          =   285
               Left            =   135
               TabIndex        =   61
               Top             =   210
               Value           =   -1  'True
               Width           =   1140
            End
            Begin VB.OptionButton OptDest 
               Caption         =   "&Destinatário"
               Height          =   195
               Left            =   1530
               TabIndex        =   59
               Top             =   255
               Width           =   1185
            End
         End
         Begin VB.CheckBox ChkSomar 
            Caption         =   "&Somar"
            Height          =   285
            Left            =   3375
            TabIndex        =   57
            Top             =   855
            Width           =   960
         End
         Begin VB.CheckBox ChkDestaq 
            Caption         =   "&Destacar"
            Height          =   240
            Left            =   3375
            TabIndex        =   56
            Top             =   645
            Width           =   960
         End
         Begin VB.ComboBox CmbObsPed 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   1395
            Width           =   7425
         End
         Begin VB.ComboBox CmbObsRoda 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            TabIndex        =   14
            Top             =   1935
            Width           =   7440
         End
         Begin MSMask.MaskEdBox Placa 
            Height          =   315
            Left            =   5625
            TabIndex        =   62
            Top             =   645
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Uf 
            Height          =   315
            Left            =   7470
            TabIndex        =   63
            Top             =   645
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo DBCTransp 
            Bindings        =   "FrmNota.frx":0060
            Height          =   315
            Left            =   1485
            TabIndex        =   64
            Top             =   225
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "nome"
            BoundColumn     =   "CodEntidade"
            Text            =   ""
         End
         Begin VB.Label LblFrete 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   119
            Top             =   945
            Width           =   2445
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            Height          =   195
            Left            =   5040
            TabIndex        =   71
            Top             =   645
            Width           =   405
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   7125
            TabIndex        =   70
            Top             =   705
            Width           =   210
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Transportadora"
            Height          =   195
            Left            =   180
            TabIndex        =   69
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Observação no Rodapé"
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   1725
            Width           =   1710
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observação &Corpo da Nota"
            Height          =   195
            Left            =   180
            TabIndex        =   67
            Top             =   1170
            Width           =   1950
         End
         Begin VB.Label QtdCorpo 
            Height          =   285
            Left            =   7650
            TabIndex        =   66
            Top             =   1395
            Width           =   570
         End
         Begin VB.Label QtdRodape 
            Height          =   285
            Left            =   7635
            TabIndex        =   65
            Top             =   1920
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1305
         Left            =   120
         TabIndex        =   38
         Top             =   2955
         Width           =   8295
         Begin MSMask.MaskEdBox Emissao 
            Height          =   315
            Left            =   180
            TabIndex        =   5
            Top             =   405
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Saida 
            Height          =   315
            Left            =   1545
            TabIndex        =   6
            Top             =   405
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox HoraSaida 
            Height          =   315
            Left            =   3105
            TabIndex        =   7
            Top             =   405
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Qtd 
            Height          =   315
            Left            =   4410
            TabIndex        =   8
            Top             =   405
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Especie 
            Height          =   315
            Left            =   5850
            TabIndex        =   9
            Top             =   405
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Marca 
            Height          =   315
            Left            =   180
            TabIndex        =   10
            Top             =   900
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PesoLiquido 
            Height          =   315
            Left            =   3105
            TabIndex        =   11
            Top             =   900
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Bruto 
            Height          =   315
            Left            =   4410
            TabIndex        =   12
            Top             =   900
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Saida"
            Height          =   195
            Left            =   1515
            TabIndex        =   45
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            Height          =   195
            Left            =   3150
            TabIndex        =   44
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bruto"
            Height          =   195
            Index           =   5
            Left            =   4410
            TabIndex        =   43
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Peso Liq."
            Height          =   195
            Index           =   4
            Left            =   3105
            TabIndex        =   42
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Qtd"
            Height          =   195
            Index           =   3
            Left            =   4410
            TabIndex        =   41
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Espécie"
            Height          =   195
            Index           =   1
            Left            =   5850
            TabIndex        =   40
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   39
            Top             =   720
            Width           =   450
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6900
         Left            =   8685
         TabIndex        =   25
         Top             =   495
         Width           =   5100
         Begin TabDlg.SSTab SSTab1 
            Height          =   1725
            Left            =   2700
            TabIndex        =   135
            Top             =   3780
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   3043
            _Version        =   393216
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "SUB."
            TabPicture(0)   =   "FrmNota.frx":007A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label35"
            Tab(0).Control(1)=   "Label34"
            Tab(0).Control(2)=   "txticmssub"
            Tab(0).Control(3)=   "txtbasesub"
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "IPI"
            TabPicture(1)   =   "FrmNota.frx":0096
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label37"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label36"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "VarPIpi"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "VarIpiDev"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "ICMS"
            TabPicture(2)   =   "FrmNota.frx":00B2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "BaseIcmsDif"
            Tab(2).Control(1)=   "ValorIcmDif"
            Tab(2).Control(2)=   "Label39"
            Tab(2).Control(3)=   "Label38"
            Tab(2).ControlCount=   4
            Begin MSMask.MaskEdBox txtbasesub 
               Bindings        =   "FrmNota.frx":00CE
               Height          =   315
               Left            =   -74775
               TabIndex        =   136
               Top             =   630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txticmssub 
               Bindings        =   "FrmNota.frx":00D9
               Height          =   315
               Left            =   -74775
               TabIndex        =   137
               Top             =   1215
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox VarIpiDev 
               Bindings        =   "FrmNota.frx":00E4
               Height          =   315
               Left            =   225
               TabIndex        =   140
               Top             =   1215
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox VarPIpi 
               Bindings        =   "FrmNota.frx":00EF
               Height          =   315
               Left            =   225
               TabIndex        =   141
               Top             =   630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox BaseIcmsDif 
               Bindings        =   "FrmNota.frx":00FA
               Height          =   315
               Left            =   -74775
               TabIndex        =   144
               Top             =   630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox ValorIcmDif 
               Bindings        =   "FrmNota.frx":0105
               Height          =   315
               Left            =   -74775
               TabIndex        =   145
               Top             =   1215
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base Icms Dif."
               Height          =   195
               Left            =   -74775
               TabIndex        =   147
               Top             =   405
               Width           =   1020
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Icms Dif"
               Height          =   195
               Left            =   -74775
               TabIndex        =   146
               Top             =   990
               Width           =   975
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Ipi Devolução"
               Height          =   195
               Left            =   225
               TabIndex        =   143
               Top             =   990
               Width           =   1395
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Percentual Ipi Devolução"
               Height          =   195
               Left            =   225
               TabIndex        =   142
               Top             =   405
               Width           =   1800
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base Icms Sub"
               Height          =   195
               Left            =   -74775
               TabIndex        =   139
               Top             =   405
               Width           =   1065
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Icms Sub"
               Height          =   195
               Left            =   -74775
               TabIndex        =   138
               Top             =   990
               Width           =   1065
            End
         End
         Begin VB.CheckBox ChKBar 
            Caption         =   "Código Barra"
            Height          =   330
            Left            =   180
            TabIndex        =   134
            Top             =   3465
            Width           =   2220
         End
         Begin VB.CommandButton Confirma 
            Caption         =   "Simulação Nota (Nfe)"
            Height          =   390
            Index           =   3
            Left            =   2700
            TabIndex        =   133
            Top             =   5580
            Width           =   2220
         End
         Begin VB.CheckBox ChkEstorno 
            Caption         =   "Nota de &Estorno"
            Height          =   330
            Left            =   180
            TabIndex        =   129
            Top             =   3140
            Width           =   2220
         End
         Begin VB.CheckBox ChkPcl 
            Caption         =   "Imprimir Parcelas"
            Height          =   330
            Left            =   180
            TabIndex        =   125
            Top             =   2490
            Width           =   1770
         End
         Begin VB.CheckBox ChkCfop 
            Caption         =   "&Cfop da Operação"
            Height          =   330
            Left            =   180
            TabIndex        =   120
            Top             =   540
            Width           =   2220
         End
         Begin VB.Frame Frame6 
            Caption         =   "Regime Tributação"
            Height          =   690
            Left            =   2565
            TabIndex        =   100
            Top             =   135
            Width           =   2445
            Begin VB.Label LblRegime 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   135
               TabIndex        =   101
               Top             =   315
               Width           =   2175
            End
         End
         Begin VB.CheckBox ChkIpi 
            Caption         =   "IPI."
            Height          =   255
            Left            =   1980
            TabIndex        =   92
            Top             =   150
            Width           =   600
         End
         Begin VB.CommandButton Confirma 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Nota Eletrônica (Nfe)"
            Height          =   705
            Index           =   2
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3780
            Width           =   2220
         End
         Begin VB.CheckBox ChkSub 
            Caption         =   "&Substituição Tributária"
            Height          =   330
            Left            =   180
            TabIndex        =   83
            Top             =   2815
            Width           =   2220
         End
         Begin VB.CheckBox ChkDesconto 
            Caption         =   "Des&tacar Desconto"
            Height          =   330
            Left            =   180
            TabIndex        =   36
            Top             =   1515
            Width           =   2220
         End
         Begin VB.CheckBox ChkDeducao 
            Caption         =   "&Dedução Base de Calculo"
            Height          =   330
            Left            =   2745
            TabIndex        =   35
            Top             =   810
            Width           =   2220
         End
         Begin VB.CheckBox ChkCompl 
            Caption         =   "Nota &Complementar"
            Height          =   330
            Left            =   180
            TabIndex        =   34
            Top             =   1840
            Width           =   1770
         End
         Begin VB.CheckBox ChkFunRural 
            Caption         =   "FunRural"
            Height          =   330
            Left            =   180
            TabIndex        =   33
            Top             =   1190
            Width           =   2220
         End
         Begin VB.CheckBox ChkSimples 
            Caption         =   "&Nota com base no simples"
            Height          =   330
            Left            =   180
            TabIndex        =   32
            Top             =   865
            Width           =   2265
         End
         Begin VB.CommandButton CmbLimpa 
            Caption         =   "&Limpa Tela"
            Height          =   390
            Left            =   180
            TabIndex        =   31
            Top             =   4620
            Width           =   2220
         End
         Begin VB.CommandButton Confirma 
            Caption         =   " &Imprimir Nota Formulário"
            Enabled         =   0   'False
            Height          =   750
            Index           =   0
            Left            =   4770
            TabIndex        =   16
            Top             =   1215
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton CmbPedido 
            Caption         =   "&Alterar Itens Nota"
            Enabled         =   0   'False
            Height          =   390
            Left            =   180
            TabIndex        =   30
            Top             =   5580
            Width           =   2220
         End
         Begin VB.CommandButton CmbParcelas 
            Caption         =   "&Parcelar Pedido"
            Enabled         =   0   'False
            Height          =   390
            Left            =   180
            TabIndex        =   29
            Top             =   5085
            Width           =   2220
         End
         Begin VB.CheckBox ChkValor 
            Caption         =   "Valor dos Produtos"
            Height          =   330
            Left            =   180
            TabIndex        =   28
            Top             =   2165
            Width           =   1770
         End
         Begin VB.CheckBox ChkIcms 
            Caption         =   "ICMS"
            Height          =   255
            Left            =   990
            TabIndex        =   26
            Top             =   150
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CommandButton Confirma 
            Height          =   390
            Index           =   1
            Left            =   4770
            TabIndex        =   60
            Top             =   2250
            Width           =   195
         End
         Begin MSMask.MaskEdBox MebReducao 
            Bindings        =   "FrmNota.frx":0110
            Height          =   315
            Left            =   2745
            TabIndex        =   27
            Top             =   1170
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebValor 
            DataField       =   "Venda"
            Height          =   435
            Left            =   2655
            TabIndex        =   37
            Top             =   6255
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   767
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
         Begin MSMask.MaskEdBox MebSub 
            Bindings        =   "FrmNota.frx":011B
            Height          =   315
            Left            =   2745
            TabIndex        =   84
            Top             =   1755
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtOutras 
            Bindings        =   "FrmNota.frx":0126
            Height          =   315
            Left            =   2745
            TabIndex        =   102
            Top             =   2295
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtDesconto 
            Bindings        =   "FrmNota.frx":0131
            Height          =   315
            Left            =   2745
            TabIndex        =   127
            Top             =   2835
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox PedRef 
            Bindings        =   "FrmNota.frx":013C
            Height          =   315
            Left            =   2745
            TabIndex        =   148
            Top             =   3420
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "c"
            PromptChar      =   "_"
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ped Referencia"
            Height          =   195
            Left            =   2745
            TabIndex        =   149
            Top             =   3195
            Width           =   1110
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto"
            Height          =   195
            Left            =   2745
            TabIndex        =   128
            Top             =   2610
            Width           =   690
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outras Despesas"
            Height          =   195
            Left            =   2745
            TabIndex        =   103
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4455
            TabIndex        =   97
            Top             =   1845
            Width           =   180
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4455
            TabIndex        =   96
            Top             =   1215
            Width           =   180
         End
         Begin VB.Label Label23 
            Caption         =   "Calcular"
            Height          =   240
            Left            =   180
            TabIndex        =   91
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crédito Presumido"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2745
            TabIndex        =   85
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   5085
            Y1              =   6060
            Y2              =   6060
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Cfo"
         Connect         =   "Access"
         DatabaseName    =   "C:\Rede\Sistemas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   1830
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from cfo order by descricao"
         Top             =   960
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.TextBox TxtCfop 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   4
         Top             =   975
         Width           =   1590
      End
      Begin VB.CommandButton BtAtualiza 
         Caption         =   "&Atualiza"
         Height          =   315
         Left            =   -61995
         TabIndex        =   20
         Top             =   7200
         Width           =   810
      End
      Begin VB.Data DatPed 
         Caption         =   "Pedido"
         Connect         =   "Access"
         DatabaseName    =   "E:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -69735
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3375
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.ListBox LstNota 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6960
         Left            =   -74925
         TabIndex        =   18
         Top             =   450
         Visible         =   0   'False
         Width           =   13860
      End
      Begin MSDBGrid.DBGrid DBGrid 
         Bindings        =   "FrmNota.frx":0147
         Height          =   5820
         Left            =   -74865
         OleObjectBlob   =   "FrmNota.frx":015C
         TabIndex        =   19
         Top             =   585
         Width           =   13770
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   315
         Left            =   -63495
         TabIndex        =   21
         Top             =   7200
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99287041
         CurrentDate     =   37501
      End
      Begin MSDBCtls.DBCombo DbcCfo 
         Bindings        =   "FrmNota.frx":16EE
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   975
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Descricao"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox CodPed 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   615
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebNota 
         Height          =   315
         Left            =   1875
         TabIndex        =   1
         Top             =   615
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1710
         Left            =   135
         TabIndex        =   47
         Top             =   1260
         Width           =   8280
         Begin MSMask.MaskEdBox Entidade 
            DataField       =   "0"
            Height          =   270
            Left            =   1860
            TabIndex        =   48
            Top             =   180
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox EnderecoEntidade 
            Height          =   270
            Left            =   900
            TabIndex        =   49
            Top             =   450
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CidadeEntidade 
            Height          =   270
            Left            =   900
            TabIndex        =   50
            Top             =   990
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebEntidade 
            Height          =   270
            Left            =   900
            TabIndex        =   51
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Nro 
            Height          =   270
            Left            =   6930
            TabIndex        =   72
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   270
            Left            =   900
            TabIndex        =   74
            Top             =   720
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox UfEntidade 
            Height          =   270
            Left            =   6930
            TabIndex        =   76
            Top             =   945
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox IE 
            Height          =   270
            Left            =   900
            TabIndex        =   78
            Top             =   1260
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CepEntidade 
            Height          =   270
            Left            =   6930
            TabIndex        =   86
            Top             =   675
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep:"
            Height          =   195
            Left            =   6435
            TabIndex        =   87
            Top             =   750
            Width           =   330
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cnpj/Cpf"
            Height          =   195
            Left            =   3210
            TabIndex        =   80
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "I.E./R.G.:"
            Height          =   195
            Left            =   105
            TabIndex        =   79
            Top             =   1290
            Width           =   690
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Uf:"
            Height          =   195
            Left            =   6555
            TabIndex        =   77
            Top             =   1005
            Width           =   210
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   345
            TabIndex        =   75
            Top             =   750
            Width           =   450
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nro:"
            Height          =   195
            Left            =   6465
            TabIndex        =   73
            Top             =   495
            Width           =   300
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   255
            TabIndex        =   54
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Entidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   225
            Width           =   675
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   60
            TabIndex        =   52
            Top             =   480
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox MebNfe 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   615
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Label QtdChave 
         Height          =   285
         Left            =   7920
         TabIndex        =   126
         Top             =   7065
         Width           =   300
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Chave nota de origem da devolução (Digitar somente os numeros)"
         Height          =   195
         Left            =   180
         TabIndex        =   122
         Top             =   6750
         Width           =   4875
      End
      Begin VB.Label Lbltam 
         BackStyle       =   0  'Transparent
         Height          =   330
         Left            =   -69555
         TabIndex        =   118
         Top             =   6705
         Width           =   645
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo de Numeros"
         Height          =   195
         Left            =   -69690
         TabIndex        =   112
         Top             =   6480
         Width           =   1515
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Nota"
         Height          =   195
         Left            =   -65550
         TabIndex        =   108
         Top             =   7245
         Width           =   945
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Justificativa Cancelamento, Inutilizacao e Carta de Correção"
         Height          =   195
         Left            =   -74820
         TabIndex        =   99
         Top             =   6480
         Width           =   4245
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota &Eletrônica"
         Height          =   195
         Left            =   3615
         TabIndex        =   94
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pedido"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.F.Op."
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nota Formulario"
         Height          =   195
         Left            =   1860
         TabIndex        =   22
         Top             =   420
         Width           =   1110
      End
   End
End
Attribute VB_Name = "FrmNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Doc As DOMDocument

Public xml As New MSXML2.DOMDocument

Public dadosLoteNfe As MSXML2.IXMLDOMElement
Dim TabSql As Recordset 'tabela que sera iniciada pelo sql que esta no banco
Public pedidos As Recordset
Public Produtos As Recordset
Public Pag As Recordset
Public Dup As Recordset
Public Oper As Byte
Public EndTransp As String
Public CidTransp As String
Public DocTransport As String
Public UfTransp As String
Public IETransp As String
Public VarPacote As Recordset
Public VarIcms As Currency
Public VarRedutor As Currency
Public IcmsNota As Currency
Public VarNroNfe As String
Public VarSerieNfe As String
Public VarNota As Recordset
Public VarSub As Boolean

Public Registro01 As String
Public Registro02 As String
Public Registro03 As String
Public Registro04 As String
Public Registro05 As String
Public Registro06 As String
Public Registro07 As String
Public Registro08 As String
Public Registro09 As String
Public Registro10 As String
Public Registro11 As String
Public Registro12 As String
Public Registro13 As String
Public Registro14 As String
Public Registro15 As String
Public Registro19 As String
Public Registro20 As String
Public Registro99 As String
Public ConfCfg As Recordset
Public ContadorItens As Integer
Public DtAnt As Date
Public VarVolume As String
Public VarSub2 As Boolean
Dim VarPais As String


Private Sub BBtStatus_Click()
Dim VarCnpjEmite As String
Dim VarSerieNfe As String
Dim varNome As String
Dim VarNfe As String
Dim VarNro As Double

If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
    VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
End If

If LerINI("Nfe", "SerieNfe", App.Path & "\config.ini") <> "" Then
    VarSerieNfe = LerINI("Nfe", "SerieNfe", App.Path & "\config.ini")
End If
Comstatus.InitDir = PortaAcbr & "\logs"
Comstatus.filename = "*-nfe.xml"
Comstatus.ShowOpen

varNome = Comstatus.filename
If varNome <> "*-nfe.xml" Then
        VarNfe = Mid$(varNome, Len(varNome) - 51, 44)
        VarNro = Mid$(VarNfe, 26, 9)

        AbreAcbr "NFE.ConsultarNFe(" & varNome & ")"
        If TrataAcbr("Consulta", 1) = 100 Then
            If DBGrid.Columns(3) = 0 Then
                DatPed.Recordset.Edit
                DatPed.Recordset("NroNt") = VarNro
                DatPed.Recordset.Update
                TxtNroNota = VarNro
                
                Set VarNota = Banco.OpenRecordset("Select * from nota where codped = " & DBGrid.Columns(0))
                VarNota.Requery
                    
                If VarNota.RecordCount > 0 Then
                    VarNota.Edit
                    VarNota!InformFisco = VarNfe
                    VarNota.Update
                End If
                
                BtNfe_Click 5
                
                If MsgBox("Avançar um numero na nota ?", vbYesNo) = vbYes Then
                        EscreveINI "Nfe", "NroNfe", CInt(VarNro) + 1, App.Path & "\config.ini"
                End If
                
                BtNfe_Click 0
                                
            End If
        
        End If
End If

'If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\*" & Trim(TxtNroNota) & "?_01-procEventoNFe.xml") <> "" Then
'    MsgBox "Essa nota foi cancelada !  ", vbInformation, App.Title
'Else

'    Shell "ERP.IntegraTxt.exe " & TxtNroNota & " 7"
'End If

End Sub

Private Sub BtNfe_Click(Index As Integer)
Dim VarNfe As Recordset
Dim VarCnpjEmite As String
Dim VarSerieNfe As String
Dim VarInutil As String
Dim VarData As Date
Dim VarNroNfe As String
Dim VarTemp As String
Dim VarResp As String
'DatNfe.RecordSource = "Select * FROM Nfe WHERE CodPed = " & DBGrid.Columns(0)
'DatNfe.Refresh

'If TxtNroNota <> "" Then
    
'End If
'Shell "C:\ACBrMonitorPLUS\PDF\35110400178853000118550020000001061000001060-nfe.PDF"
'AbreAcbr "NFe.ImprimirDANFEPDF(" & App.Path & "\106_NE_00178853000118.xml" & ")"
'TrataAcbr

'Exit Sub

If Index >= 2 And Index <= 4 Then
    If Len(TxtCancela) < 20 Then
        MsgBox "Justificatica de cancelamento deve ter no minimo 20 caracteres !", vbInformation, App.Title
        TxtCancela.SetFocus
        Exit Sub
    End If
    
End If

If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
    VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
End If

If LerINI("Nfe", "SerieNfe", App.Path & "\config.ini") <> "" Then
    VarSerieNfe = LerINI("Nfe", "SerieNfe", App.Path & "\config.ini")
End If

If PortaAcbr <> "" And ChkVersao = 1 Then
    Set VarNota = Banco.OpenRecordset("Select * from nota where codped = " & DBGrid.Columns(0))
    VarNota.Requery
        
    If VarNota.RecordCount > 0 Then
    '    VarNota.Edit
         VarChaveNfe = VarNota!InformFisco
    '    VarNota.Update
    End If
    

    Select Case Index
        Case 0
            If Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") <> "" Then
               If MsgBox("Confirma impressão da segunda via da NFE ? ", vbYesNo, App.Title) = vbYes Then
                    ShellExecute Me.Hwnd, "open", PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf", _
                          vbNullString, vbNullString, vbNormalFocus
               End If
            Else
                MsgBox "Nota não encontrada !", vbInformation, App.Title
            End If
        Case 1
            Set VarNota = Banco.OpenRecordset("SELECT PED.CodPed, ENT.Email FROM PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade WHERE PED.CodPed= " & DBGrid.Columns(0))
            VarNota.Requery
        
            VarEmail = VarNota!Email
            If MsgBox("Deseja enviar e-mail do xml agora ?", vbYesNo, App.Title) = vbYes Then
            
                 AbreAcbr "NFe.EnviarEmail(" & VarEmail & "," & PortaAcbr & "\Logs\" & VarChaveNfe & "-nfe.xml" & ",1," & "Envio de NFe - " & NomeEmpresa & ")"
            End If
        
        
        Case 2
            If MsgBox("Confirma Cancelamento da NFE ? ", vbYesNo, App.Title) = vbYes Then
               AbreAcbr "Nfe.CANCELARNFE(" & VarChaveNfe & "," & TxtCancela & "," & VarCnpjEmite & ")"
               VarResp = TrataAcbr("Cancela", 1)
               If VarResp = 135 Or VarResp = 155 Then
                        MsgBox "Nota cancelada com sucesso !", vbInformation, App.Title
                        DatPed.Recordset.Edit
                        DatPed.Recordset("NroNt") = 0
                        DatPed.Recordset("Cfo") = ""
                        DatPed.Recordset.Update
               End If
            End If
                
        Case 4
        
            If Dir(App.Path & "\Cfe.Ini") <> "" Then
               Kill App.Path & "\Cfe.ini"
            End If
            
            If Dir(PortaAcbr & "\Sai.txt") <> "" Then
                Kill PortaAcbr & "\Sai.txt"
            End If
            If MsgBox("Confirma envio da carta de correção ? ", vbYesNo, App.Title) = vbYes Then
               EscreveINI "EVENTO", "IdLote", "1", App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "ChNfe", VarChaveNfe, App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "cOrgao", "35", App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "Cnpj", VarCnpjEmite, App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "dhEvento", Date & " " & Time, App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "tpEvento", "110110", App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "nSeqEvento", TxtSeq, App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "versaoEvento", "4.00", App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "descEvento", "Carta de Correcao", App.Path & "\cfe.ini"
               EscreveINI "EVENTO001", "xCorrecao", TxtCancela, App.Path & "\cfe.ini"
               
               AbreAcbr "NFE.EnviarEvento(" & App.Path & "\cfe.ini" & ")"
               If TrataAcbr("Carta", psCompString(1, TxtSeq, 3)) = 135 Then
                    MsgBox "Carta transmitida com sucesso !", vbInformation, App.Title
               End If

            End If
        Case 5
            If VarChaveNfe <> "" Then
                AbreAcbr "Nfe.ImprimirDANFEPDF(" & PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfe\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\Nfe\" & VarChaveNfe & "-nfe.xml" & ")"
'                AbreAcbr "Nfe.ImprimirDANFEPDF(" & PortaAcbr & "\Logs\" & VarChaveNfe & "-nfe.xml" & ")"
            Else
                MsgBox "Chave da nota não encontrada !", vbInformation, App.Title
            End If
            
         Case 6
               VarChaveNfe = InputBox("Baixar Nfe", "Digite a Chave da Nota", VarChaveNfe)
               AbreAcbr "NFe.DistribuicaoDFePorChaveNFe(" & "SP," & VarCnpjEmite & ", " & VarChaveNfe & ")"
        
    End Select
Else

    If Index <> 3 Then
    
        If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml") <> "" Then
            Select Case Index
                Case 0
                    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procCancNFe.xml") = "" Then
                        If MsgBox("Confirma impressão da segunda via da NFE ? ", vbYesNo, App.Title) = vbYes Then
                                If Dir("C:\ACBrMonitorPLUS\ACBRMONITOR.EXE") = "" Then
                                      Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml 3"
                                Else
                                      AbreAcbr "NFe.ImprimirDANFE(" & LocalDrive & "\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml" & ")"
                                End If
                        
                        End If
                    Else
                        MsgBox "Essa nota foi cancelada !  ", vbInformation, App.Title
                    End If
                Case 1
                    Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(DBGrid.Columns(3)) & "-procNfe.xml 4"
                Case 2
                    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procCancNFe.xml") = "" Then
                        If MsgBox("Confirma Cancelamento da NFE ? ", vbYesNo, App.Title) = vbYes Then
                            Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml 5 " & """" & Trim(TxtCancela) & """"
                            While Dir("..\MonitorNfe\" & VarCnpjEmite & "\erro\" & Trim(DBGrid.Columns(3)) & "-nfe.xml") = "" And Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\*" & Trim(TxtNroNota) & "?_01-procEventoNFe.xml") = "" And Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-canc.txt") = ""
                                DoEvents
                            Wend
                            If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\*" & Trim(TxtNroNota) & "?_01-procEventoNFe.xml") <> "" Then
                                MsgBox "Nota cancelada com sucesso", vbInformation, App.Title
                                DatPed.Recordset.Edit
                                DatPed.Recordset("NroNt") = 0
                                DatPed.Recordset.Update
                            ElseIf Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-canc.txt") <> "" Then
                                    Open "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-canc.txt" For Input As #6
                                    Line Input #6, VarTemp
                                    Close #6
                                    MsgBox VarTemp, vbInformation
    '                                Shell "notepad.exe ..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-re-ret-env-canc.txt"
                                    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\*" & Trim(TxtNroNota) & "?_110111_01-procEventoNFe.xml") <> "" Then
                                        MsgBox "Nota cancelada com sucesso", vbInformation, App.Title
                                        DatPed.Recordset.Edit
                                        DatPed.Recordset("NroNt") = 0
                                        DatPed.Recordset.Update
                                    End If
                                    
                                Else
                                    MsgBox "Erro no cancelamento da nota !", vbInformation, App.Title
                            End If
                        End If
                    Else
                        MsgBox "Essa nota ja foi cancelada !  ", vbInformation, App.Title
                    End If
                Case 4
                    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procCancNFe.xml") = "" Then
                        If MsgBox("Confirma envio carta de correção da NFE ? ", vbYesNo, App.Title) = vbYes Then
                            Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml 9 " & """" & Trim(TxtCancela) & """" & " " & TxtSeq
                            While Dir("..\MonitorNfe\" & VarCnpjEmite & "\erro\" & Trim(DBGrid.Columns(3)) & "-env-cce.xml") = "" And Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-cce.txt") = ""
                                DoEvents
                            Wend
                            If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-cce.txt") <> "" Then
    '                            MsgBox "Carta Enviada com sucesso !", vbInformation, App.Title
                                    Open "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-cce.txt" For Input As #6
                                    Line Input #6, VarTemp
                                    Close #6
                                    MsgBox VarTemp, vbInformation
    
    '                                Shell "notepad.exe ..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-cce.txt"
                            Else
                                MsgBox "Erro no envio da carta !", vbInformation, App.Title
                                If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-ret-env-cce.err") <> "" Then
                                    Open "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-re-env-cce.err" For Input As #6
                                    Line Input #6, VarTemp
                                    Close #6
                                    MsgBox VarTemp, vbInformation
                                    
    '                                Shell "notepad.exe ..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(TxtNroNota) & "-re-env-cce.err"
                                End If
                            End If
                        End If
                    Else
                        MsgBox "Essa nota ja foi cancelada !  ", vbInformation, App.Title
                    End If
                
            End Select
        Else
            MsgBox "Não tem nota eletrônica gerada ! " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(DBGrid.Columns(3)) & "-procNfe.xml", vbInformation, App.Title
        End If
    Else
        
        If MsgBox("Confirma a inutilizacao da faixa espeficada ? ", vbYesNo, App.Title) = vbYes Then
            
            Open "env_Inutil.txt" For Output As #1
    
            
            VarInutil = "16" & VarCnpjEmite & Trim(Str(Year(DtData))) & "SP"
            VarInutil = VarInutil & "55" & psCompString(3, VarSerieNfe, 3)
            VarInutil = VarInutil & psCompString(1, Nro1, 9) & psCompString(1, Nro2, 9) & Trim(TxtCancela)
            
            Print #1, VarInutil
            Close #1
            
            Shell "ERP.IntegraTxt.exe env_Inutil.txt 8", vbHide
            
            
        End If
            
    End If
    
End If
End Sub

Private Sub ChkDeducao_Click()
If ChkDeducao.Enabled = True Then
    MebReducao.Enabled = True
    MebReducao = Format(LerINI("Nota", "Reducao", App.Path & "\config.ini"), "###,##0.0000")
Else
    MebReducao.Enabled = False
End If
End Sub

Private Sub ChkValor_Click()
If MebNota <> "" Then
    If ChkValor.value = 0 Then
        MebValor = pedidos("Total")
    Else
        MebValor = pedidos("TotalPrd")
    End If
End If
End Sub

Private Sub CmbLimpa_Click()
LimpaCampos
End Sub

Private Sub CmbObsPed_Change()
QtdCorpo = Len(CmbObsPed)
End Sub

Private Sub CmbObsRoda_Change()
QtdRodape = Len(CmbObsRoda)
End Sub

Private Sub CmbParcelas_Click()

If CodPed = "" Then
    MsgBox "Pedido inválido", vbExclamation, App.Title
    CodPed.SetFocus
    Exit Sub
End If

FrmParcelas.CodDoc = CodPed
FrmParcelas.CodEnt = MebEntidade
FrmParcelas.MebVlrPed = Format(MebValor, "#,##0.00")
FrmParcelas.TxtNome = Entidade
FrmParcelas.Show 'vbModal

Set pedidos = Banco.OpenRecordset("SELECT ENT.*, PED.*, PED.CodPed FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.CodPed = " & VarCodPed)
pedidos.Requery

End Sub

Private Sub CmbPedido_Click()

If CodPed = "" Then
    MsgBox "Pedido inválido", vbExclamation, App.Title
    CodPed.SetFocus
    Exit Sub
End If

FormFatura.Show vbModal
Set pedidos = Banco.OpenRecordset("SELECT ENT.*, PED.*, PED.CodPed FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.CodPed = " & VarCodPed)
pedidos.Requery

End Sub

Private Sub CodPed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
    FormConsultaPed.Show
End If
End Sub

Private Sub CodPed_KeyPress(KeyAscii As Integer)
Dim Criterio As String
    If KeyAscii = 13 Then
    
        If LerINI("FIXO", "Deducao", App.Path & "\nfe.ini") <> "" Then
            VarPercDeducao = LerINI("FIXO", "Deducao", App.Path & "\nfe.ini")
        Else
            VarPercDeducao = ""
        End If

        If IsNumeric(CodPed) Then
            VarCodPed = Val(CodPed)
        Else
            VarCodPed = 0
        End If
        
        LimpaCampos
        If LerINI("Nota", "Marca", App.Path & "\config.ini") <> "" Then
            Marca = LerINI("Nota", "Marca", App.Path & "\config.ini")
        End If
        Set pedidos = Banco.OpenRecordset("SELECT ENT.*, PED.*, PED.CodPed FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.CodPed = " & VarCodPed)
        pedidos.Requery
                
        Set Produtos = Banco.OpenRecordset("SELECT Sum(ITP.Qtd) AS SomaDeQtd, Sum([itp].[qtd]*[prd].[pesobruto]) AS SomaBruto, Sum([itp].[qtd]*[prd].[pesoliquido]) AS SomaLiquido, First(PRD.Embalagem) AS PrimeiroDeEmbalagem, ITP.Pedido FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd GROUP BY ITP.Pedido HAVING ITP.Pedido = " & VarCodPed & " ORDER BY First(PRD.Embalagem) ")
        Produtos.Requery
        If VarTranspNt = True Then
            Set VarPacote = Banco.OpenRecordset("SELECT PRD.Embalagem  FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido = " & VarCodPed & " GROUP BY PRD.Embalagem")
            VarPacote.Requery
            If VarVolume = "" Then
                If VarPacote.RecordCount > 0 Then
                    VarPacote.MoveFirst
                    Especie = Especie & VarPacote!Embalagem
                    VarPacote.MoveNext
                    While Not VarPacote.Eof
                        Especie = Especie & "/" & VarPacote!Embalagem
                        VarPacote.MoveNext
                    Wend
                End If
                
                If Produtos!SomaDeQtd <> "" Then
                    Qtd = Produtos!SomaDeQtd
                Else
                    Qtd = 0
                End If
            Else
                Especie = VarVolume
            End If
            PesoLiquido = Produtos!somaliquido
            Bruto = Produtos!somabruto
        End If
        
        If pedidos.RecordCount = 0 Then
            MsgBox "Pedido não encontrado !!!", 16
            LimpaCampos
            CodPed.SelStart = 0
            CodPed.SelLength = 5000
            CodPed.SetFocus
            Exit Sub
        Else
            If pedidos!NroNt <> 0 Then
                If MsgBox("Está pedido já foi emitido nota ! Deseja emitir outra com o mesmo pedido ?", vbYesNo, App.Title) = vbNo Then
                    LimpaCampos
                    CodPed.SelStart = 0
                    CodPed.SelLength = 5000
                    CodPed.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        CodPed.Text = VarCodPed
        MebEntidade.Text = pedidos("CodEntidade")
        Entidade.Text = pedidos("Nome")
        If Not IsNull(pedidos("Endereco")) Then
            EnderecoEntidade.Text = pedidos("Endereco")
        End If
        
        If Not IsNull(pedidos("Nro")) Then
            Nro.Text = pedidos("Nro")
        Else
            Nro.Text = ""
        End If
        
        If Not IsNull(pedidos("Bairro")) Then
            Bairro.Text = pedidos("Bairro")
        Else
            Bairro.Text = ""
        End If
                
        If Not IsNull(pedidos("Cidade")) Then
            CidadeEntidade.Text = pedidos("Cidade")
        Else
            CidadeEntidade.Text = ""
        End If
        
        If Not IsNull(pedidos("Ent.Uf")) Then
            UfEntidade.Text = pedidos("Ent.Uf")
        Else
            UfEntidade.Text = ""
        End If
        
        If Not IsNull(pedidos("Cep")) Then
            CepEntidade.Text = pedidos("Cep")
        Else
            CepEntidade.Text = ""
        End If
        
        LblFrete = "CIF"
                
        If Not IsNull(pedidos("TipoFrete")) Then
            If pedidos("TipoFrete") = "FOB" Then
                LblFrete = "FOB"
                OptDest.value = True
                ChkDestaq.value = 1
                ChkSomar.value = 0
            End If
        End If
        
        If Not IsNull(pedidos("PedTipoSub")) Then
            If pedidos("PedTipoSub") = True Then
                ChkSub.value = 1
                MebSub = pedidos("PedCredito")
            End If
        End If
                                                        
        MebDocumento.Mask = ""
        If pedidos!Classificacao = 2 Then
            ChkDoc.value = 1
            ChkDoc.Enabled = True
            If Not IsNull(pedidos("CGC")) And Not IsNull(pedidos("Cpf")) Then
                If pedidos!CGC <> "" Then
                    MebDocumento = Format(LIMPACampo2(pedidos!CGC), "@@.@@@.@@@/@@@@-@@")
                Else
                    MebDocumento = Format(LIMPACampo2(pedidos!Cpf), "@@@.@@@.@@@-@@")
                End If
            End If
            If Not IsNull(pedidos("InscrEst")) Then
                IE = pedidos!InscrEst
            End If
'            If MebDocumento = "" Or IE = "" Then
'                If MsgBox("Nota para pessoa juridica sem CNPJ ou Inscr.Est. Deseja continuar !", vbYesNo, App.Title) = vbNo Then
'                    Exit Sub
'                End If
'            End If
        ElseIf pedidos!Classificacao = 1 Then
            If Not IsNull(pedidos("Cpf")) Then
                If pedidos("Cpf") <> "" Then
                    ChkDoc.value = 1
                    MebDocumento = Format(LIMPACampo2(pedidos!Cpf), "@@@.@@@.@@@-@@")
                Else
                    MebDocumento.Mask = "###.###.###-##"
                End If
            End If
            If Not IsNull(pedidos("Rg")) Then
                IE = Format(LIMPACampo2(pedidos!Rg), "@@.@@@.@@@")
            End If
        End If
        
        If VarPercDeducao <> "" And UfEntidade = "SP" Then
            ChkDeducao.value = 1
            MebReducao = VarPercDeducao
        End If
        
        
        If ChkValor.value = 0 Then
            MebValor = pedidos("Total")
        Else
            MebValor = pedidos("Total")
        End If
            
        Oper = pedidos("Operacao")
        
        If Left(pedidos("InscrEst"), 1) = "P" And pedidos("Operacao") = 0 Then
              ChkFunRural.value = 1
        Else
              ChkFunRural.value = 0
        End If
        If pedidos!FunRural = True Then
            ChkFunRural.value = 1
        End If
        If Dir(Caminho & "\BitNfe.EXE") <> "" Then
'            Confirma(0).Enabled = False
'            Confirma(1).Enabled = True
        Else
'            Confirma(0).Enabled = True
'            Confirma(1).Enabled = False
        End If
        CmbPedido.Enabled = True
        CmbParcelas.Enabled = True
        If LerINI("NOTA", "Cfop", App.Path & "\Config.ini") <> "" Then
            DbcCfo.Text = LerINI("NOTA", "Cfop", App.Path & "\Config.ini")
            TxtCfop.Text = LerINI("NOTA", "Cfop", App.Path & "\Config.ini")
        Else
            DbcCfo.Text = "5.102"
            TxtCfop.Text = "5.102"
        End If
        
        If Dir(App.Path & "\Obs.Txt") <> "" Then
            Dim NomeObs As String
            On Error GoTo ErroObs
            CmbObsPed.Clear
            CmbObsRoda.Clear
            Open App.Path & "\Obs.Txt" For Input As #1
            Do
                Input #1, NomeObs
                CmbObsPed.AddItem NomeObs
                CmbObsRoda.AddItem NomeObs
            Loop Until NomeObs = ""
            Close #1
ErroObs:
            If Err.Number = 62 Then
                Close #1
            End If
        
        End If
        
        If pedidos("Ped.Obs") <> "" Then
            CmbObsPed.Text = pedidos("Ped.Obs")
        End If
        
        DBCTransp.Text = ""
        
        HoraSaida.Text = Time
        
        Emissao.Text = pedidos("DataEmiss")
        DtAnt = pedidos("DataEmiss")
        
        If Not IsNull(pedidos("DtFat")) Then
            Emissao.Text = pedidos("DtFat")
        End If
        
        If Not IsNull(pedidos("DtSaida")) Then
            Saida.Text = pedidos("DtSaida")
'            Saida.Text = Pedidos("DtFat")
'            DbcCfo.Text = Pedidos("Cfo")
'            TxtCfop.Text = Pedidos("Cfo")
        
        End If
        
        If Dir(Caminho & "\ERP.IntegraTxt.exe") <> "" Then
            Saida.Text = Date
            Emissao.Text = Date
        End If
        
        If IsNull(Configuracao("NroNota")) Then
            MebNota.Text = 1
        Else
            MebNota.Text = Configuracao("NroNota")
'            MebNota.Text = Pedidos("NroNt")
                
        End If
        
        VarNroNfe = ""
        
        If LerINI("Nfe", "NroNfe", App.Path & "\config.ini") <> "" Then
            VarNroNfe = LerINI("Nfe", "NroNfe", App.Path & "\config.ini")
            MebNfe = VarNroNfe
        Else
            MebNfe = MebNota
        End If
                
        Criterio = "CodEntidade = " & pedidos("Ped.Transportadora")
        DatDados(1).Recordset.FindFirst Criterio
        If DatDados(1).Recordset.NoMatch Then
            DBCTransp.Text = "(Nenhum)"
            EndTransp = ""
            CidTransp = ""
            DocTransport = ""
            UfTransp = ""
            IETransp = ""
        Else
            DBCTransp.Text = DatDados(1).Recordset.Fields("Nome")
            If Not IsNull(DatDados(1).Recordset("Endereco")) Then
                EndTransp = Trim(DatDados(1).Recordset("Endereco")) & ", " & DatDados(1).Recordset("Nro")
            Else
                EndTransp = ""
            End If
            If Not IsNull(DatDados(1).Recordset("Cidade")) Then
                CidTransp = Trim(DatDados(1).Recordset("Cidade"))
            Else
                CidTransp = ""
            End If
            
            If Not IsNull(DatDados(1).Recordset("Uf")) Then
                UfTransp = Trim(DatDados(1).Recordset("Uf"))
            Else
                UfTransp = ""
            End If
            
            If (DatDados(1).Recordset("Cpf")) <> "" Then
                DocTransport = DatDados(1).Recordset("Cpf")
            ElseIf (DatDados(1).Recordset("Cgc")) <> "" Then
                    DocTransport = DatDados(1).Recordset("Cgc")
                Else
                    DocTransport = ""
            End If
            If (DatDados(1).Recordset("InscrEst")) <> "" Then
                IETransp = DatDados(1).Recordset("InscrEst")
            Else
                IETransp = ""
            End If
                
        End If
        If VarObsAutomatica Then
            If CmbObsPed.ListCount > 0 Then
                CmbObsPed.ListIndex = 0
            End If
            If CmbObsRoda.ListCount > 0 Then
                CmbObsRoda.ListIndex = 1
            End If
        End If
        
    If pedidos("Ped.Desconto") <> 0 Then
'        If Index = 2 Then
            TxtDesconto = pedidos("Ped.Desconto")
'        End If
        ChkDesconto.value = 1
    Else
        TxtDesconto = 0
        ChkDesconto = 0
    End If

        
        MebNota.SetFocus
   End If
                        
End Sub


Private Sub Consulta_Click()

Open Caminho & "\env_consulta.txt" For Output As #3
Print #3, "13" & Trim(TxtRecibo) & "SPS"
Print #3, "99"
Close #3

Frame.Visible = True
Label22.Visible = True
Label22.Caption = "Aguarde . . . esperando retorno . . ."
DoEvents

While Dir(Caminho & "\env_consulta.txt") = ""
    Frame.Visible = True
    Label22.Visible = True
    Label22.Caption = "Aguarde . . . esperando retorno . . ."
    DoEvents
Wend

'Frame.Visible = False
        Label22.Visible = False
Shell "BitNFE.exe env_consulta.txt 2", vbNormalFocus

End Sub

Private Sub DatPed_Reposition()
 TxtNroNota = Trim(DBGrid.Columns(3))
End Sub

Private Sub DbcCfo_LostFocus()
Dim Criterio As String
If DbcCfo.Text <> "" Then
    If IsNumeric(Left(DbcCfo, 1)) Then
        Criterio = "Codigo = " & """" & DbcCfo & """"
        DatDados(0).Recordset.FindFirst Criterio
        If Not DatDados(0).Recordset.NoMatch Then
            DbcCfo.Text = DatDados(0).Recordset.Fields("Descricao")
        Else
            MsgBox "C.f.Op. inválido !", vbInformation, App.Title
            DbcCfo.SetFocus
        End If
    End If
    If IsNumeric(DbcCfo.BoundText) Then
        TxtCfop.Text = DbcCfo.BoundText
    End If
    If Trim(LIMPACampo2(TxtCfop)) = "5405" Or Trim(LIMPACampo2(TxtCfop)) = "5102" Then
        ChkCfop.value = 0
    Else
        ChkCfop.value = 1
    End If

    
'    If Oper = 0 Or Oper = 9 Or Oper = 10 Then ' Entrada
'        If Left(TxtCfop, 1) > 4 Then
'            MsgBox "C.f.Op inconsistente com a natureza da operação", vbInformation, App.Title
'            DbcCfo.SetFocus
'        End If
'    Else ' Saida
'        If Left(TxtCfop, 1) < 4 Then
'            MsgBox "C.f.Op inconsistente com a natureza da operação", vbInformation, App.Title
'            DbcCfo.SetFocus
'        End If
'    End If
End If
End Sub


Private Sub DBCTransp_LostFocus()
Dim Criterio As String

        Criterio = "CodEntidade = " & DBCTransp.BoundText
        DatDados(1).Recordset.FindFirst Criterio
        If Not DatDados(1).Recordset.NoMatch Then
        End If

End Sub

Private Sub DbGrid_dblClick()

If DatPed.Recordset.RecordCount > 0 Then
    CodPed.Text = DBGrid.Columns(0)
    CodPed.SetFocus
    SSTabNota.Tab = 1
    SendKeys ("{Enter}")
End If

End Sub

Private Sub Form_Activate()
BtAtualiza_Click
If DatPed.Recordset.RecordCount > 0 Then
    If CodPed = "" Then
        SSTabNota.Tab = 0
    Else
        SSTabNota.Tab = 1
    End If
Else
    SSTabNota.Tab = 1
End If


If LerINI("Nota", "Substituicao", App.Path & "\config.ini") <> "" Then
    VarSub = LerINI("Nota", "substituicao", App.Path & "\config.ini")
Else
    VarSub = 0
End If



End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If LstNota.Visible = True Then
        LstNota.Visible = False
    Else
        Unload Me
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

Private Sub LimpaCampos()
MebEntidade.Text = ""
TxtOutras = ""
CepEntidade = ""
Entidade.Text = ""
EnderecoEntidade.Text = ""
CidadeEntidade.Text = ""
Nro.Text = ""
Bairro = ""
UfEntidade = ""
IE = ""
MebDocumento = "___.___.___-__"
DBCTransp.Text = ""
HoraSaida.Text = ""
DbcCfo.Text = ""
Emissao.Text = "__/__/____"
Saida.Text = "__/__/____"
Placa.Text = ""
Uf.Text = ""
Qtd.Text = 0
Especie.Text = ""
Marca.Text = ""
PesoLiquido.Text = 0
Bruto.Text = 0
CodPed.Text = ""
'TxtObs.Text = ""
CmbObsPed = ""
CmbObsRoda = ""
MebNota.Text = ""
MebNfe.Text = ""
'ChkSimples = 0
'ChkFunRural = 0
MebValor = 0
CodPed.SetFocus
End Sub

Private Sub Confirma_Click(Index As Integer)
Dim Conf As Recordset 'recordset da tabela CNT (configuracao formulario)
Dim Classe As Recordset
'Dim ConfCfg As Recordset
Dim Icm As Recordset

'Dim Transporte As Recordset
Dim QtdLinhas As Currency 'variavel onde fica informado a quantidade de linhas
Dim LinhaAtual As Currency 'variavel onde fica informado a linha atual
Dim UltimoItem As Currency 'variavel onde fica informado a ultima linha, para o for next funcionar corretamente
Dim X As Byte   'variavel que conta as linhas
Dim LinhaAnt As Integer
Dim contlinha As Integer
Dim ClasseF(1 To 90) As String
Dim Criterio As String
Dim F As Byte
Dim t As Byte
Dim Det As Integer
Dim TotLinhas As Integer
Dim Controle As Boolean
Dim Dup As Recordset
Dim y As Integer
Dim VarObs As String
Dim ContFat As Boolean
Dim LinhaFat As Integer
Dim TotColunaFat As Integer
Dim TotLinhaFat As Integer
Dim ColunaFat As Integer
Dim VarFunRural As Currency
Dim Valor As Currency
Dim Desconto As Currency
Dim varValor As Currency
Dim VarTemp As Integer
Dim VarColObs As Integer
Dim PortaTemp As String
Dim SQL As String
Dim PortaPedido As String
Dim ModeloTemp As String
Dim VarCnpjEmite As String
Dim ContNfe As Boolean


Desconto = 0
Controle = True
VarIcms = 0
VarRedutor = 0
ContFat = True
Valor = 0
VarColObs = 0
PortaPedido = ""

If LerINI("Nota", "Substituicao", App.Path & "\config.ini") <> "" Then
    VarSub = LerINI("Nota", "substituicao", App.Path & "\config.ini")
Else
    VarSub = False
End If


If Index = 2 Or Index = 3 Then
    Confirma(2).Enabled = False
    If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
        VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
    End If

    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(Date))) & psCompString(1, Month(Date), 2) & "\" & MebNfe & "-procNfe.xml") <> "" Then
        If MsgBox("Essa nota ja esta validada, Deseja Reimprimir ?", vbYesNo, App.Title) = vbYes Then
            If Dir("C:\ACBrMonitorPLUS\ACBRMONITOR.EXE") = "" Then
    '              Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(DBGrid.Columns(7)))) & psCompString(1, Month(DBGrid.Columns(7)), 2) & "\" & Trim(TxtNroNota) & "-procNfe.xml 3"
    '             Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(Emissao))) & psCompString(1, Month(Emissao), 2) & "\" & MebNfe & "-procNfe.xml 3"
                   Shell "ERP.IntegraTxt.exe " & "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(Date))) & psCompString(1, Month(Date), 2) & "\" & MebNfe & "-procNfe.xml 3"
           
            Else
                  AbreAcbr "NFe.ImprimirDANFE(" & LocalDrive & "\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(Emissao))) & psCompString(1, Month(Emissao), 2) & "\" & MebNfe & "-procNfe.xml" & ")"
                  TrataAcbr "Nota", MebNfe
                
                  AbreAcbr "NFe.ImprimirDANFEPDF(" & LocalDrive & "\MonitorNfe\" & VarCnpjEmite & "\Enviado\Autorizados\" & Trim(Str(Year(Emissao))) & psCompString(1, Month(Emissao), 2) & "\" & MebNfe & "-procNfe.xml" & ")"
                  TrataAcbr "Nota", MebNfe
            End If
        End If
        LimpaCampos
        Exit Sub
    End If
    
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Erro\" & MebNfe & "-nfe.xml") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\Erro\" & MebNfe & "-nfe.xml"
    End If
    
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\validar\validado\" & MebNfe & "-nfe.xml") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\validar\validado\" & MebNfe & "-nfe.xml"
    End If
    
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\Enviado\EmProcessamento\" & MebNfe & "-nfe.xml") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\Enviado\EmProcessamento\" & MebNfe & "-nfe.xml"
    End If
        
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-num-lot.xml") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-num-lot.xml"
    End If
        
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-num-lot.txt") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-num-lot.txt"
    End If
        
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-nfe.err") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-Nfe.err"
    End If
        
    If Dir("..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-nfe-ret.xml") <> "" Then
        Kill "..\MonitorNfe\" & VarCnpjEmite & "\retorno\" & Trim(MebNfe) & "-Nfe-ret.xml"
    End If
        
        
        
End If

If DbcCfo.BoundText = "" Then
    MsgBox "C.f.Op inválido !", vbExclamation, App.Title
    DbcCfo.SetFocus
    Confirma(2).Enabled = True
    Exit Sub
End If

If CodPed = "" Then
    MsgBox "Pedido inválido", vbExclamation, App.Title
    CodPed.SetFocus
    Confirma(2).Enabled = True
    Exit Sub
End If

If ChkDeducao = 1 Then
    If Not IsNumeric(MebReducao) Then
        MsgBox "Calculo da dedução invalido", vbInformation, App.Title
        Confirma(2).Enabled = True
        Exit Sub
    End If
End If

If Not IsNumeric(Qtd) Then
    MsgBox "O Campo Quantidade deve ser numerico !", vbExclamation, App.Title
    Qtd.SetFocus
    Confirma(2).Enabled = True
    Exit Sub
End If
    
If pedidos!pais = "" Or pedidos!pais = "BRASIL" Then
    VarPais = "BRASIL"
    If ChkDoc.value = 1 Then
        If pedidos("Classificacao") = 1 Then
            If Not CalculaDacCPF(Trim(MebDocumento)) Then
                MsgBox "CPF Inválido", vbInformation
                Confirma(2).Enabled = True
                Exit Sub
            End If
        Else
            If Not VerificaCGC(Trim(MebDocumento)) = True Then
                MsgBox "CNPJ Inválido.", vbInformation
                Confirma(2).Enabled = True
                Exit Sub
            End If
        End If
    End If
Else
    VarPais = pedidos!pais
    
End If

If Index = 2 Or Index = 3 Then
    If VarPais = "BRASIL" Then
        If IsNull(pedidos!matricula) Then
           MsgBox "É necessario o código do IBGE !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
        End If
        
        If pedidos!matricula = "" Then
           MsgBox "É necessario o código do IBGE !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
        End If
        
        If Len(Trim(pedidos!matricula)) <> 7 Then
           MsgBox "Codigo IBGE incorreto !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
        End If
        
        If Len(Trim(pedidos!Nro)) = "" Then
           MsgBox "Numero do Imovel nao preenchido !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
        End If
        
            
        If Len(Trim(pedidos!Fone)) = 0 Or Len(Trim(pedidos!Fone)) = "" Then
           MsgBox "Telefone do cliente não informado !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
        End If
    End If
End If

If pedidos!pais = "" Or pedidos!pais = "BRASIL" Then

    If Len(Trim(pedidos!Cep)) <> 8 Then
           MsgBox "CEP incorreto !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
    End If
    
    If Len(Trim(pedidos.Fields("Ent.Uf"))) <> 2 Then
           MsgBox "UF incorreto !", vbInformation, App.Title
           Confirma(2).Enabled = True
           Exit Sub
    End If
End If

If MsgBox("Confirma Impressão da Nota ? ", vbYesNo, App.Title) = vbYes Then
'    On Error Resume Next
    pedidos.Edit
    If Index = 2 Or Index = 3 Then
'        pedidos("NroNt") = MebNfe.Text
    Else
        pedidos("NroNt") = MebNota.Text
    End If
    pedidos("Cfo") = TxtCfop
    pedidos("Entidade") = MebEntidade.Text
    pedidos("DataEmiss") = Emissao.Text
    If Saida.Text = "__/__/____" Then
        pedidos("DtSaida") = Null
    Else
        pedidos("DtSaida") = Saida.Text
    End If
    pedidos("DtFat") = Emissao.Text
    If HoraSaida.Text <> "" Then
        pedidos("Hora") = HoraSaida.Text
    Else
        pedidos("Hora") = Null
    End If
    If IsNumeric(DBCTransp.BoundText) Then
        pedidos("Ped.Transportadora") = DBCTransp.BoundText
    End If
    pedidos("PlacaVeiculo") = Placa.Text
    pedidos("Ped.Uf") = Uf.Text
    If Qtd.Text = "" Then Qtd = 0
    pedidos("QtdVolume") = Qtd.Text
    pedidos("Especie") = Especie.Text
    pedidos("Marca") = Marca.Text
    pedidos("PesoLiquido") = PesoLiquido.Text
    pedidos("PesoBruto") = Bruto.Text
    
    pedidos.Update
                
    Variavel(1) = 0 'Valor total da Nota
    Variavel(2) = 0 'Valor total dos Produtos
    Variavel(3) = 0 ' Valor total do Icms
    Variavel(4) = 0 'Base de calculo do Ims
    Variavel(5) = 0 'Valor total do Ipi
    Variavel(6) = 0 'Calculo do do IPI
    Variavel(7) = 0 'Valor total do Ipi
    Variavel(8) = 0 'Valor do Frete
    Variavel(9) = 0 'Tipo do Frete
    Variavel(10) = "" 'Placa
    Variavel(11) = "" 'Uf Placa
    Variavel(12) = ""  'Obs
    Variavel(13) = ""  'Simples
    Variavel(14) = "" ' Classificação Fiscal
    Variavel(15) = "" ' Nome Transportadora
    Variavel(16) = "" ' Pedido / Parcela
    Variavel(17) = "" 'Cnpj / Cpf
    Variavel(18) = 0 ' preco unitario
    Variavel(19) = 0 ' VALOR
    Variavel(20) = "" ' Classificação Fiscal
    Variavel(21) = "" ' Dedução base
    Variavel(22) = "" ' Dedução base
    Variavel(23) = "" ' Dedução base
    Variavel(24) = "" ' Endereco da Transportadora
    Variavel(25) = "" ' Cidade da Transportadora
    Variavel(26) = "" ' RG/Inscr
    Variavel(27) = "" ' Documento Transportadora
    Variavel(28) = "" ' Nota / Parcela
    Variavel(29) = "" ' ie transp.
    Variavel(30) = "" 'Uf Transp.
    Variavel(31) = 0 ' BaseSubst
    Variavel(32) = 0 ' IcmSubst
    Variavel(33) = 0 ' Desconto Sub
    Variavel(34) = "" ' Situação Tributaria
            
    Set Icm = Banco.OpenRecordset("Select * from Icm")
    Icm.Requery
                       
    Set ConfCfg = Banco.OpenRecordset("Select * From Cfg")
    ConfCfg.Requery
    SQL = "SELECT ENT.*, PED.*, ITP.*, PRD.*, PED.CodPed, CFO.* FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN CFO ON PED.Cfo = CFO.Codigo) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE Ped.CodPed = " & VarCodPed & "  ORDER BY ITP.CodItp"
    
    Set TabSql = Banco.OpenRecordset("SELECT ENT.*, PED.*, ITP.*, PRD.*, PED.CodPed, CFO.* FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN CFO ON PED.Cfo = CFO.Codigo) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE Ped.CodPed = " & VarCodPed & "  ORDER BY ITP.CodItp")
    TabSql.Requery
    
    If Index = 5 Then
        ContNfe = False
        If TabSql.RecordCount > 0 Then
            TabSql.MoveFirst
            While Not TabSql.Eof
                If LblRegime.Caption = "3 - Regime Normal" Or LblRegime.Caption = "2 - Regime Diferenciado" Then
                      If TabSql!SitTrib > "100" Or TabSql!SitTrib = "" Then
                            MsgBox "Erro no regime de tributacao", vbInformation, App.Title
                            ContNfe = True
                      End If
                Else
                      If TabSql!SitTrib < "100" Or TabSql!SitTrib = "" Then
                            MsgBox "Erro no regime de tributacao", vbInformation, App.Title
                            ContNfe = True
                      End If
                End If
                If ContNfe = True Then
                    TabSql.MoveLast
                End If
                TabSql.MoveNext
            Wend
            If ContNfe = True Then
                Exit Sub
            End If
            TabSql.MoveFirst
        End If
    End If
    

    If ChkSub = 1 And VarSub2 = True Then
        Dim VarUfIva As Recordset
'                SQL = "Select * from IVA Where Uf = " & """" & Trim(FrmPedidos.Datdados(0).Recordset.Fields("Uf")) & """" & " AND Classe = " & """" & Trim(Datdados(1).Recordset.Fields("Classe")) & """"
        Set VarUfIva = Banco.OpenRecordset("Select * from IVA Where Uf = " & """" & Trim(TabSql("ENT.Uf")) & """" & " AND Ncm = " & """" & Trim(TabSql("Classe")) & """")
        VarUfIva.Requery
        If VarUfIva.RecordCount > 0 Then
            VarIcms = TabSql!Icms
            VarRedutor = TabSql!Icms
        Else
            MsgBox "Estado sem icms cadastrado", vbExclamation, App.Title
            Exit Sub
        End If
         
    Else
        If TabSql("ENT.Uf") <> "SP" Then
            If Icm.RecordCount > 0 Then
                Criterio = "UF = " & """" & TabSql("ENT.Uf") & """"
                Icm.FindFirst Criterio
                If Icm.NoMatch Then
                    MsgBox "Estado sem icms cadastrado", vbExclamation, App.Title
                    Exit Sub
                Else
                    VarIcms = Icm!Icms
                    VarRedutor = Icm!Entrada
                End If
            Else
                MsgBox "Nota fora do estado, tabela Icms vazia !", vbExclamation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    
    
    If TabSql!Classificacao = 2 Then
'        If TabSql!CGC <> "" Then
            Variavel(17) = MebDocumento  'Format(MebDocumento, "@@.@@@.@@@/@@@@-@@")
'        Else
'            Variavel(17) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
'        End If
        Variavel(26) = TabSql!InscrEst
'        If Variavel(17) = "" Or Variavel(26) = "" Then
'           If MsgBox("Nota para pessoa juridica sem CNPJ ou Inscr.Est. Deseja continuar !", vbYesNo, App.Title) = vbNo Then
'                Exit Sub
'            End If
'        End If
    ElseIf TabSql!Classificacao = 1 Then
        Variavel(17) = MebDocumento 'Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
        Variavel(26) = Format(TabSql!Rg, "@@.@@@.@@@")
    End If
    
    If ChkSub = 1 Then
        Variavel(31) = TabSql("PedBaseSub")
        Variavel(32) = TabSql("PedSub")
'                            Variavel(31) = Variavel(31) + (((Valor / 100) * TabSql("Iva")) + Valor)
'                            Variavel(33) = ((((Valor / 100) * TabSql("Iva")) + Valor) * (TabSql("Icm") / 100))
'                            If Not IsNumeric(MebSub) Then
'                                MebSub = 0
'                            End If
'                                    Variavel(32) = Variavel(32) + ((Variavel(33) - ((Valor / 100) * MebSub)))

    End If
     
    If txtbasesub <> "" Then
        Variavel(31) = txtbasesub
    End If
    
    If txticmssub <> "" Then
        Variavel(32) = txticmssub
    End If
    
        
    If TabSql("Ped.Desconto") <> 0 Then
        If Index = 2 Then
            TxtDesconto = TabSql("Ped.Desconto")
        End If
        ChkDesconto.value = 1
    Else
        TxtDesconto = 0
        ChkDesconto = 0
    End If
    
'    Set Transporte = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Tipo, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Uf FROM ENT WHERE (((ENT.Tipo)=4 Or (ENT.Tipo)=6)) ORDER BY ENT.CodEntidade")
    Set Classe = Banco.OpenRecordset("SELECT PRD.ClasseFiscal FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido = " & Val(CodPed.Text) & " GROUP BY PRD.ClasseFiscal ")
    Classe.Requery
    Classe.MoveFirst
    F = 0
    While Not Classe.Eof
        If Not IsNull(Classe!CLASSEFISCAL) Then
            F = F + 1
            ClasseF(F) = Classe!CLASSEFISCAL
        End If
        Classe.MoveNext
    Wend
    Dim varQtdLinhas As Currency 'variavel onde fica informado a quantidade de linhas
    Dim varLinhaAtual As Currency 'variavel onde fica informado a linha atual
    Dim varUltimoItem As Currency 'variavel onde fica informado a ultima linha, para o for next funcionar corretamente
    
    PortaPedido = Porta
    If LerINI("Impressora", "ImprNota", App.Path & "\config.ini") <> "" Then
        PortaTemp = LerINI("Impressora", "ImprNota", App.Path & "\config.ini")
    End If
                                                              
    If LerINI("Nota", "QtdLinhaFatura", App.Path & "\config.ini") <> "" Then
        TotLinhaFat = LerINI("Nota", "QtdLinhaFatura", App.Path & "\config.ini")
    Else
        TotLinhaFat = 2
    End If
                                                              
    If LerINI("Nota", "QtdColunaFatura", App.Path & "\config.ini") <> "" Then
        TotColunaFat = LerINI("Nota", "QtdColunaFatura", App.Path & "\config.ini")
    Else
        TotColunaFat = 3
    End If
    
                                                                                                                            
    If PortaTemp <> "" Then
        Porta = PortaTemp
    End If
    
    If Index >= 1 Then
        PortaTemp = Porta
        Porta = Caminho & "\Nota.txt"
    End If
        
    Iniciar True 'inicia o registro onde sera gravado as informacoes
    
        If ConfCfg!Espacamento = "1/8" Then
            ImprimeLinha 0, 0, Chr(27) & Chr(48), Len(Chr(27) & Chr(48)), False, False
        Else
            ImprimeLinha 0, 0, Chr(27) & Chr(50), Len(Chr(27) & Chr(50)), False, False
        End If
           
' *************   Cabeçalho da Nota
    
    TotLinhas = Int(TabSql.RecordCount / ConfCfg!TotalLinhaDetalhe)
    If (TabSql.RecordCount Mod ConfCfg!TotalLinhaDetalhe) > 0 Then
        TotLinhas = TotLinhas + 1
    End If
    
    ContadorItens = 0
    
    If TxtOutras = "" Then
         TxtOutras = 0
    End If

    
    For Det = 1 To TotLinhas
            Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 1 Order by Codigo")
            Conf.Requery
            Conf.MoveFirst
    
            LinhaAnt = 0
            If Conf.RecordCount <> 0 Then
               While Not Conf.Eof
                     If Conf!Linha <> LinhaAnt Then
                        For X = 1 To Conf!Linha - LinhaAnt
                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                        Next
                     End If
                     LinhaAnt = Conf!Linha
                     If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                            If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                            Else
                                If Mid(Conf("Formato"), 1, 1) = "+" Then
                                      ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                Else
                                      ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                End If
                            End If
                     Else
                            If Conf!Formato = "X" Then
                                   If TabSql(Conf!NomeCampo) = 1 Or TabSql(Conf!NomeCampo) = 2 Or TabSql(Conf!NomeCampo) = 3 Or TabSql(Conf!NomeCampo) = 6 Or TabSql(Conf!NomeCampo) = 8 Or TabSql(Conf!NomeCampo) = 11 Or TabSql(Conf!NomeCampo) = 13 Then
                                           ImprimeLinha 0, Conf!Coluna, "X", Conf!TamanhoCampo, Conf!Condensado, 0
                                   Else
                                           ImprimeLinha 0, Conf!Coluna, Space(Conf!TamanhoCampo - 1) & "X", Conf!TamanhoCampo, Conf!Condensado, 0
                                   End If
                            Else
                                  If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                       If IsNull(TabSql(Conf!NomeCampo)) Then
                                          ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                       Else
                                          ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf("NomeCampo")), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       End If
                                  Else
                                       If IsNull(TabSql(Conf!NomeCampo)) Then
                                          ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                       Else
                                          ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                       End If
                                  End If
                            End If
                     End If
                     Conf.MoveNext
               Wend
            End If
            
            ' *************   Fatura
            If Oper = 0 Then ' Compra
                Set Pag = Banco.OpenRecordset("Select * From Pag Where Pedido = " & CodPed.Text)
                Pag.Requery
                If Pag.RecordCount > 0 Then
                    Pag.MoveFirst
                End If
                Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 6 Order by Codigo")
                Conf.Requery
                contlinha = 0
                LinhaFat = 0
                ColunaFat = 0
                           
                If Conf.RecordCount <> 0 Then
                      Conf.MoveFirst
                      If ContFat = True Then
                       Do While Not Pag.Eof
                             Conf.MoveFirst
                             While Not Conf.Eof
                                    Variavel(16) = Pag!Numero
'                                    If ContFat = True Then
                                        If Conf!Linha <> LinhaAnt Then
                                            For X = 1 To Conf!Linha - LinhaAnt
                                                ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                            Next
                                        End If
'                                        ContFat = False
'                                    End If
                                    LinhaAnt = Conf!Linha
                                    If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                           If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                               ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                           Else
                                               If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                     ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               Else
                                                     ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               End If
                                           End If
                                    Else
                                           If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                                If Conf!NomeCampo <> "ICM" Then
                                                      ImprimeLinha 0, Conf!Coluna, Format(Pag(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                Else
                                                      If VarIcms = 0 Then
                                                          ImprimeLinha 0, Conf!Coluna, Format(Pag(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                      Else
                                                          ImprimeLinha 0, Conf!Coluna, Format(VarIcms, CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
            '                                              ImprimeLinha 0, Conf!Coluna, CStr(VarIcms), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                      End If
                                                End If
                                           Else
                                                If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                   ImprimeLinha 0, Conf!Coluna, CStr(Pag(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                Else
                                                   If IsNull(Pag(Conf!NomeCampo)) Then
                                                      ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   Else
                                                      ImprimeLinha 0, Conf!Coluna, CStr(Pag(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   End If
                                               End If
                                           End If
                                    End If
                                    Conf.MoveNext
                             Wend
                             
                             ColunaFat = ColunaFat + 1
                             If ColunaFat = TotColunaFat Then
                                    ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                    ColunaFat = 0
                                    LinhaFat = LinhaFat + 1
                             End If
                                
'                             If ContFat = True Then
'                                ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
'                                contlinha = contlinha + 1
'                             End If
'                             pag.MoveNext
'                             LinhaFat = LinhaFat + 1
'                             If LinhaFat > TotLinhaFat Then
'                                ContFat = True
'                             End If
                       Loop
                       ContFat = False
                     End If
                       
                End If
                       
            Else ' venda
                Set Dup = Banco.OpenRecordset("SELECT PED.NroNt, Dup.*, Dup.CodDup FROM Dup INNER JOIN PED ON Dup.Pedido = PED.CodPed WHERE Dup.Pedido = " & CodPed.Text & "  ORDER BY Dup.CodDup ")
                Dup.Requery
                If Dup.RecordCount > 0 Then
                    Dup.MoveFirst
                End If
                
                Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 5 Order by Codigo")
                Conf.Requery
                contlinha = 0
                LinhaFat = 0
                ColunaFat = 0
                           
                If Conf.RecordCount <> 0 Then
                      Conf.MoveFirst
                    
                      If ContFat = True Then
                       Do While Not Dup.Eof
                             Conf.MoveFirst
                             While Not Conf.Eof
                                    Variavel(16) = Dup!Pedido & "/" & Trim(CStr(Dup!NroPcl))
                                    Variavel(28) = MebNota & "/" & Trim(CStr(Dup!NroPcl))
                                    
'                                    If ContFat = True Then
                                        If Conf!Linha <> LinhaAnt Then
                                            For X = 1 To Conf!Linha - LinhaAnt
                                                ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                            Next
                                        End If
'                                        ContFat = False
'                                    End If
                                    LinhaAnt = Conf!Linha
                                    If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                           If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                               ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                           Else
                                               If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                     ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               Else
                                                     ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               End If
                                           End If
                                    Else
                                           If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                                If Conf!NomeCampo <> "ICM" Then
                                                      ImprimeLinha 0, Conf!Coluna, Format(Dup(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                Else
                                                      If VarIcms = 0 Then
                                                          ImprimeLinha 0, Conf!Coluna, Format(Dup(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                      Else
                                                          ImprimeLinha 0, Conf!Coluna, Format(VarIcms, CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
            '                                              ImprimeLinha 0, Conf!Coluna, CStr(VarIcms), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                      End If
                                                End If
                                           Else
                                                If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                   ImprimeLinha 0, Conf!Coluna, CStr(Dup(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                Else
                                                   If IsNull(Dup(Conf!NomeCampo)) Then
                                                      ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   Else
                                                      ImprimeLinha 0, Conf!Coluna, CStr(Dup(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   End If
                                               End If
                                           End If
                                    End If
                                    Conf.MoveNext
                             Wend
                             Dup.MoveNext
                             
                             ColunaFat = ColunaFat + 1
                             If ColunaFat = TotColunaFat Then
                                    ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                    ColunaFat = 0
                                    LinhaFat = LinhaFat + 1
                             End If
                       Loop
                       ContFat = False
                     End If
                
                End If
                
             End If
             
            If LinhaFat < TotLinhaFat Then
                For X = 1 To TotLinhaFat - LinhaFat
                    ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                Next
            End If
            LinhaFat = 0
                             
            ' *************   Linha de Detalhe
               
            Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 2 Order by Codigo")
            Conf.Requery
            Conf.MoveFirst
            contlinha = 0
                       
    '        LinhaAnt = Conf!Linha
            If Conf.RecordCount <> 0 Then

                   Do While Not TabSql.Eof
                         If ChkValor.value = 0 Then
                            varValor = TabSql("VlrVdaDia")
                         Else
                            varValor = TabSql("ValorUnit") * TabSql("Qtd")
                         End If
                         
                         If ChkDesconto.value = 0 Then
                             Valor = varValor
                         Else
                             Valor = varValor
'                             Desconto = Desconto + TabSql("Ped.Desconto")
                             Desconto = TabSql("Ped.Desconto")
                         End If
                         If ChkFunRural.value = 1 Then
                            VarFunRural = TabSql("Ped.Desconto")
                         End If
                         
                         If ChkSub = 1 Then
                            Variavel(34) = "060"
                         Else
                            Variavel(34) = TabSql("SitTrib")
                         End If
                         
                         Variavel(19) = Valor
                         Variavel(2) = Variavel(2) + Valor ' total do produtos
                         If TabSql("Qtd") <> 0 Then
                            Variavel(18) = Valor / TabSql("Qtd") ' preco unitario
                         End If
                            
                         If ChkIcms.value = 1 Then
                            If VarIcms = 0 Then
                                If TabSql!SitTrib = "000" Or TabSql!SitTrib = "020" Or TabSql!SitTrib = "101" Or TabSql!SitTrib = "102" Or TabSql!SitTrib = "090" Or TabSql!SitTrib = "900" Then
                                    If TabSql!SitTrib = "900" And ChkDesconto = 1 Then
'                                        Variavel(3) = Variavel(3) + ((Valor - (TxtDesconto / TabSql.RecordCount)) * (TabSql("ICM") / 100))   ' valor icms
                                        Variavel(3) = Variavel(3) + (((Valor - CCur(TxtDesconto)) / 100) * TabSql("ICM"))
                                    Else
                                        Variavel(3) = Variavel(3) + ((Valor / 100) * TabSql("ICM")) ' valor icms
                                    End If
                                Else
                                    Variavel(3) = Variavel(3) + 0
                                End If
                                IcmsNota = TabSql("Icm")
                                If TabSql("ICM") <> 0 Then
                                    If TabSql!SitTrib = "000" Or TabSql!SitTrib = "020" Or TabSql!SitTrib = "101" Or TabSql!SitTrib = "102" Or TabSql!SitTrib = "090" Or TabSql!SitTrib = "900" Then
                                        If TabSql!SitTrib = "900" And ChkDesconto = 1 Then
                                            Variavel(4) = Variavel(4) + Valor ' - CCur(TxtDesconto)   ' valor da base de calculo
                                        Else
                                            Variavel(4) = Variavel(4) + Valor ' valor da base de calculo
                                        End If
                                    Else
                                        Variavel(4) = Variavel(4) + 0
                                    End If
                                End If
                            Else
                               If TabSql!SitTrib = "000" Or TabSql!SitTrib = "020" Or TabSql!SitTrib = "101" Or TabSql!SitTrib = "102" Or TabSql!SitTrib = "090" Or TabSql!SitTrib = "900" Then
                                   If TabSql!SitTrib = "900" And ChkDesconto = 1 Then
'                                        Variavel(3) = Variavel(3) + ((Valor - (TxtDesconto / TabSql.RecordCount)) * (VarIcms / 100))   ' valor icms
                                       Variavel(4) = Variavel(4) + Valor ' - TxtDesconto
                                        Variavel(3) = Variavel(3) + (Variavel(4) * (VarIcms / 100))    ' valor icms
                                    Else
                                        Variavel(3) = Variavel(3) + ((Valor / 100) * VarIcms) ' valor icms
                                        Variavel(4) = Variavel(4) + Valor
                                    End If
                               Else
                                    Variavel(3) = Variavel(3) + 0
                                    Variavel(4) = Variavel(4) + 0
                               
                               End If
                               IcmsNota = VarIcms
                            End If
                         Else
                              Variavel(3) = Variavel(3) + 0
                         End If
                         
                                                     
                         If ChkIpi.value = 1 Then
                            Variavel(6) = ((Valor / 100) * TabSql("Prd.Ipi")) ' valor do ipi
                         Else
                              Variavel(6) = 0
                         End If
                                                  
                         Variavel(7) = Variavel(7) + Variavel(6) ' total do ipi
'                         Sql = " UPDATE ITP SET AliqIcms = " & IcmsNota & ", AliqIpi = " & TabSql("prd.ipi") & ", ValorIpi = " & Variavel(6) & " WHERE Pedido  = " & TabSql("Itp.CodPed")
 '                        Banco.Execute Sql
                         TabSql.Edit
                         TabSql("AliqIcms") = IcmsNota
                         TabSql("AliqIpi") = TabSql("Prd.Ipi")
                         TabSql("ValorIpi") = Variavel(6)
                         TabSql.Update
                         
                         If ChkSub = 1 Then
'                            Variavel(31) = Variavel(31) + (((Valor / 100) * TabSql("Iva")) + Valor)
'                            Variavel(33) = ((((Valor / 100) * TabSql("Iva")) + Valor) * (TabSql("Icm") / 100))
'                            If Not IsNumeric(MebSub) Then
'                                MebSub = 0
'                            End If
'                                    Variavel(32) = Variavel(32) + ((Variavel(33) - ((Valor / 100) * MebSub)))
  
                          End If
                                                      
                         For t = 1 To F
                            If TabSql!CLASSEFISCAL = ClasseF(t) Then
                                Variavel(14) = Chr(64 + t)
                                t = F
                            End If
                         Next
                         Conf.MoveFirst
                         While Not Conf.Eof
                                If Conf!Linha <> LinhaAnt Then
                                   For X = 1 To Conf!Linha - LinhaAnt
                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                   Next
                                End If
                                LinhaAnt = Conf!Linha
'                                If ChkCompl.value = 0 Then
                                   If TabSql("Qtd") <> 0 Then
                                        If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                               If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                                   ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                               Else
                                                   If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                         ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   Else
                                                         ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                   End If
                                               End If
                                        Else
                                               If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                                    If Conf!NomeCampo <> "ICM" Then
                                                          ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                    Else
                                                          If VarIcms = 0 Then
                                                              ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                                          Else
                                                              ImprimeLinha 0, Conf!Coluna, Format(VarIcms, CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                '                                              ImprimeLinha 0, Conf!Coluna, CStr(VarIcms), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                          End If
                                                    End If
                                               Else
                                                    If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                       ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                    Else
                                                       If IsNull(TabSql(Conf!NomeCampo)) Then
                                                          ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                       Else
                                                          If Conf!NomeCampo = "Descrição_Produto" Then
                                                              If Len(Trim(CStr(TabSql(Conf!NomeCampo)))) <= Conf("TamanhoCampo") Then
                                                                    ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                              Else
                                                                    ImprimeLinha 0, Conf!Coluna, Left(CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                                    ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                                                    ImprimeLinha 0, Conf!Coluna + VarTemp, Mid$(CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo") + 1, Len(Trim(CStr(TabSql(Conf!NomeCampo)))) - Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                                    contlinha = contlinha + 1
                                                              End If
                                                          Else
                                                              If Conf!NomeCampo = "CodPrd" Then
                                                                  VarColObs = Conf!Coluna + Conf!TamanhoCampo
                                                              End If
                                                              ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                              VarTemp = CInt(Conf!Coluna + Conf("TamanhoCampo"))
                                                          End If
                                                       End If
                                                   End If
                                               End If
                                        End If
                                    Else
                                        If Conf!NomeCampo = "Descrição_Produto" Then
                                            If IsNull(Conf!Formato) Then
                                              ImprimeLinha 0, Conf!Coluna, TabSql(Conf!NomeCampo), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                            Else
                                              ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                            End If
                                        Else
                                              ImprimeLinha 0, Conf!Coluna, Space(Conf!TamanhoCampo), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                        End If
                                    End If
'                                Else
'                                    If Conf!NomeCampo = "Descrição_Produto" Or Conf!NomeCampo = "VlrVdaDia" Then
'                                        If IsNull(Conf!Formato) Then
'                                          ImprimeLinha 0, Conf!Coluna, TabSql(Conf!NomeCampo), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
'                                        Else
'                                          ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
'                                        End If
'                                    Else
'                                          ImprimeLinha 0, Conf!Coluna, Space(Conf!TamanhoCampo), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
'                                    End If
'                                End If
                                If Not Conf.Eof Then Conf.MoveNext
                         Wend
                 '        VarObs = TabSql("Ped.Obs")
                         contlinha = contlinha + 1
                         ContadorItens = ContadorItens + 1

                         ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                         TabSql.MoveNext
                         If contlinha >= ConfCfg!TotalLinhaDetalhe Then
                                Controle = False
                                Exit Do
                         End If
                   Loop
            End If
            If TabSql.Eof Then
                Controle = True
            End If
            Variavel(4) = Variavel(4) - TxtDesconto
            If VarIcms <> 0 And ChkIcms.value = 1 Then
                Variavel(3) = (Variavel(4) * (VarIcms / 100))     ' valor icms
            End If
            
            
            If Controle = True Then
                    If CmbObsPed <> "" Then
                         CmbObsPed = LimpaString(CmbObsPed)
                         For y = 1 To Int(Len(CmbObsPed) / 50) + 1
                             ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                             If y = 1 Then
                                 ImprimeLinha 0, VarColObs, Mid$(CmbObsPed, 1, 50), 50, True, False
                             Else
                                 ImprimeLinha 0, VarColObs, Mid$(CmbObsPed, (50 * (y - 1)) + 1, 50), 50, True, False
                             End If
                             contlinha = contlinha + 1
                         Next
                    End If
            
                   If ChkFunRural.value = 1 Then
                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                            ImprimeLinha 0, VarColObs, "Desconto 2,3% FunRural - R$", 30, True, False
                            ImprimeLinha 0, 1, Format(VarFunRural, "##,#0.00"), 10, True, False
                            contlinha = contlinha + 1
                   End If
                   If ChkDesconto.value = 1 Then
                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                            If Desconto <> 0 Then
                                ImprimeLinha 0, CInt(VarColObs) + 10, "Valor Total Desconto   " & Format(Desconto, "##,#0.00"), 30, True, False
 '                               Variavel(1) = Variavel(1) - Desconto
                            Else
                                ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                contlinha = contlinha + 1
                            End If
                            contlinha = contlinha + 1
                   End If
                If ConfCfg(10) > contlinha Then
                    For X = 1 To ConfCfg(10) - contlinha
                             ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                    Next
                End If
                ' *************   Transportadora
                TabSql.MoveFirst
                If ChkDestaq.value = 1 Then
                        Variavel(8) = TabSql!ValorFrete
                    Else
                        Variavel(8) = 0
                End If
                If ChkDesconto.value = 1 Then
                    If ChkFunRural.value = 0 Then
                        Variavel(1) = Variavel(1) - Desconto
                    End If
                End If
                                                            
                If pedidos("Cidade") = "Manaus" Or pedidos("Cidade") = "MANAUS" Then
                    Variavel(1) = Variavel(1) - Variavel(3)
                End If
                If OptRem.value = True Then
                    Variavel(9) = 1
                Else
                    Variavel(9) = 2
                End If
'                TxtObs.Text = LimpaString(CmbObsRoda.Text)
                Variavel(10) = Placa.Text
                If Uf.Text <> "" Then
                    Variavel(11) = Uf.Text
                Else
                    Variavel(11) = ""
                End If
                Variavel(15) = DBCTransp.Text
                Variavel(24) = EndTransp
                Variavel(25) = CidTransp
                Variavel(27) = DocTransport
                Variavel(29) = IETransp
                Variavel(30) = UfTransp
                
                Variavel(12) = Left(LimpaString(CmbObsRoda.Text), VarCompObs)
                Variavel(13) = Mid$(LimpaString(CmbObsRoda.Text), VarCompObs + 1, VarCompObs)
                Variavel(21) = Mid$(LimpaString(CmbObsRoda.Text), (VarCompObs * 2) + 1, VarCompObs)
                Variavel(22) = Mid$(LimpaString(CmbObsRoda.Text), (VarCompObs * 3) + 1, VarCompObs)
                Variavel(23) = Mid$(LimpaString(CmbObsRoda.Text), (VarCompObs * 4) + 1, VarCompObs)
                                
                If Dir(App.Path & "\NovaRazao.Txt") <> "" Then
                    Dim Razao As String
                    Open App.Path & "\NovaRazao.Txt" For Input As #1
                    Input #1, Razao
                    Close #1
                    Variavel(23) = Razao
                Else
'                    Variavel(13) = Mid$(TxtObs.Text, 46, 45)
                End If
                If ChkDestaq.value = 1 Then
                    Variavel(1) = (Variavel(1) + Variavel(2)) + Variavel(8)
                Else
                    Variavel(1) = Variavel(1) + Variavel(2)
                End If
               If ChkIcms.value = 1 Then
              
'                    Variavel(4) = Variavel(2)
               End If
                
                If ChkFunRural.value = 1 Then
                    Variavel(1) = Variavel(1) - VarFunRural
                End If

                If IsNumeric(Variavel(7)) Then
                   Variavel(1) = Variavel(1) + Variavel(7)
'                   Variavel(4) = Variavel(4) + Variavel(7)
'                   Variavel(3) = Variavel(4) * (VarIcms / 100)
                End If
                                                   
                If ChkDeducao = 1 Then
                    Variavel(4) = Variavel(2) - (Variavel(2) * (MebReducao / 100))
                    If VarIcms <> 0 Then
                        Variavel(3) = (Variavel(4) * (VarIcms / 100))
                    Else
                        Variavel(3) = (Variavel(3) - (Variavel(3) * (MebReducao / 100)))
                    End If
                    Variavel(13) = "REDUCAO BASE DE"
                    Variavel(21) = "CALCULO ICMS " & MebReducao & "%"
                    Variavel(22) = "CF.ANEXO II Art.3º,"
                    Variavel(23) = "INC.II DECRETO 45.583/01"
                
                End If
                
                If ChkSub Then
                    Variavel(1) = Variavel(1) + Variavel(32)
                End If
                
                If TxtOutras <> "" Then
                    Variavel(1) = Variavel(1) + CCur(TxtOutras)
                    Variavel(4) = Variavel(4) ' + CCur(TxtOutras)
'                    Variavel(3) = txtbase * (12 / 100)
                End If
                                                              
                If ChkSimples = 1 Then
'                    Variavel(13) = "Empresa optante pelo simples lei 9317/96"
                    Variavel(13) = "Doc.emitido ME/EPP optande pelo simples nacional."
                    Variavel(21) = "Nao gera direito a credito de ICMS,ISS e de IPI"
'                    Variavel(22) = ""
                End If
                                                                              
                Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 3 Order by Codigo")
                Conf.Requery
                Conf.MoveFirst
                If Conf.RecordCount <> 0 Then
                         Conf.MoveFirst
                         While Not Conf.Eof
                                If Conf!Linha > LinhaAnt Then
                                   For X = 1 To Conf!Linha - LinhaAnt
                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                   Next
                                End If
                                LinhaAnt = Conf!Linha
                                If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                    If Mid(Conf!NomeCampo, 2, 2) = "20" Then
                                        For t = 1 To F
                                            For X = 1 To 2
                                                If ClasseF(t) <> "" Then
                                                    Variavel(20) = Chr(64 + t) & " - " & ClasseF(t)
                                                    ImprimeLinha 0, Conf!Coluna, CStr(Variavel(20)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                                End If
                                                t = t + 1
                                            Next
                                            t = t - 1
                                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                            LinhaAnt = LinhaAnt + 1
                                        Next
                                    Else
                                       If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                           ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                           If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           Else
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           End If
                                       End If
                                    End If
                                Else
                                       If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                           ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                            If Mid(Conf("Formato"), 1, 1) = "+" Then
                                               ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("NomeCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                            Else
                                               If IsNull(TabSql(Conf!NomeCampo)) Then
                                                  ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               Else
                                                  ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               End If
                                           End If
                                       End If
                                End If
                                Conf.MoveNext
                         Wend
                End If
            Else
                Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 3 Order by Codigo")
                Conf.Requery
                Conf.MoveFirst
                If Conf.RecordCount <> 0 Then
                         Conf.MoveFirst
                         While Not Conf.Eof
                                If Conf!Linha <> LinhaAnt Then
                                   For X = 1 To Conf!Linha - LinhaAnt
                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                   Next
                                End If
                                LinhaAnt = Conf!Linha
                                ImprimeLinha 0, Conf!Coluna, "*******", Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                Conf.MoveNext
                         Wend
                End If
            
            End If
            Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 4 Order by Codigo")
            Conf.Requery
            Conf.MoveFirst
            For X = 0 To Conf!Linha
               ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
            Next
    Next
    ImprimeLinha 0, 0, Chr(27) & Chr(50), Len(Chr(27) & Chr(50)), False, False
    Fim True 'finaliza o arquivo onde foi gravado as informacoes
    If Index = 0 Then
        ConfCfg.Edit
        If IsNull(ConfCfg("NroNota")) Then
            ConfCfg("NroNota") = 2
        Else
            ConfCfg("NroNota") = CCur(ConfCfg("NroNota")) + 1
        End If
        ConfCfg.Update
    End If
'    If MsgBox("Editar o pedido ? ", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
'        FrmPedidos.TxtCodPed = CodPed
'        FrmPedidos.Show
'    End If
        If Index = 2 Then
            ModeloTemp = VarModeloNota
            VarModeloNota = "55"
        End If
        
       

        SQL = "DELETE CodPed FROM NOTA WHERE CodPed = " & CodPed
        Banco.Execute SQL

        SQL = "INSERT INTO NOTA ( CodPed, NaturezaOp,Modelo, SerieNota, TipoOp, IESubst, IMunicipal, Documento, "
        SQL = SQL & " RazaoSocial, Logradouro, Nro, Complemento, Bairro, Cidade, UfNota, Cep, "
        SQL = SQL & " IEDestinatario, BaseIcms, ValorIcms, BaseIcmsSub, ValorIcmsSub, ValorTotalPrd, ValorFrete, "
        SQL = SQL & " ValorSeguro, ValorDesconto, TotIpi, TotOutras, TotNota, TotServicos, AliqIss, ValorIss, "
        SQL = SQL & " TipoFrete, CNPJTransp, RazaoTransp, IETransp, EndTransp, CidadeTransp, UFTransp, "
        SQL = SQL & " PlacaVeiculo, QtdVolume, Especie,UF, Marca, PesoBruto, PesoLiquido, BaseNoSimples, "
        SQL = SQL & " DadosFatura, InformFisco ) VALUES ( "
        
        SQL = SQL & CodPed & ", "
        SQL = SQL & """" & DbcCfo.Text & """" & ", "
        SQL = SQL & """" & VarModeloNota & """" & ", "
        SQL = SQL & """" & 0 & """" & ", "
        SQL = SQL & LIMPACampo2(TxtCfop) & ", "
        SQL = SQL & """" & " " & """" & ", "  ' iesub
        SQL = SQL & """" & " " & """" & ", "  ' imunicipal
        If ChkDoc.value = 1 Then
            SQL = SQL & """" & MebDocumento & """" & ", "
        Else
            SQL = SQL & """" & " " & """" & ", "
        End If
        SQL = SQL & """" & Entidade & """" & ", "
        SQL = SQL & """" & EnderecoEntidade & """" & ", "
        SQL = SQL & """" & Nro & """" & ", "
        SQL = SQL & """" & " " & """" & ", "
        SQL = SQL & """" & Bairro & """" & ", "
        SQL = SQL & """" & CidadeEntidade & """" & ", "
        SQL = SQL & """" & UfEntidade & """" & ", "
        SQL = SQL & """" & CepEntidade & """" & ", "
        If pedidos!Classificacao = 2 Then
            SQL = SQL & """" & IE & """" & ", "
        Else
            SQL = SQL & """" & " " & """" & ", "
        End If
        If ChkIcms.value = 1 Then
            SQL = SQL & Numero(Format(Variavel(4), "###,##0.00"), 2) & ", "
        Else
            SQL = SQL & 0 & ", "
        End If
        SQL = SQL & Numero(Format(Variavel(3), "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(31), "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(32), "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(2), "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(8), "###,##0.00"), 2) & ", "
        SQL = SQL & 0 & ", "
        SQL = SQL & Numero(Format(Desconto, "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(7), "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(TxtOutras, "###,##0.00"), 2) & ", "
        SQL = SQL & Numero(Format(Variavel(1), "###,##0.00"), 2) & ", "
        SQL = SQL & 0 & ", "
        SQL = SQL & 0 & ", "
        SQL = SQL & 0 & ", "
        SQL = SQL & Variavel(9) & ", "
        SQL = SQL & """" & Variavel(27) & """" & ", "
        SQL = SQL & """" & Variavel(15) & """" & ", "
        SQL = SQL & """" & Variavel(29) & """" & ", "
        SQL = SQL & """" & Variavel(24) & """" & ", "
        SQL = SQL & """" & Variavel(25) & """" & ", "
        SQL = SQL & """" & Variavel(30) & """" & ", "
        SQL = SQL & """" & Variavel(10) & """" & ", "
        SQL = SQL & Numero(Format(Qtd, "###,##0.00"), 2) & ", "
        SQL = SQL & """" & Especie & """" & ", "
        SQL = SQL & """" & Variavel(11) & """" & ", "
        SQL = SQL & """" & Marca & """" & ", "
        SQL = SQL & Numero(Format(PesoLiquido, "###,##0.000"), 3) & ", "
        SQL = SQL & Numero(Format(Bruto, "###,##0.000"), 3) & ", "
        If ChkSimples.value = 1 Then
            SQL = SQL & "True" & ", "
        Else
            SQL = SQL & "False" & ", "
        End If
        SQL = SQL & """" & Variavel(12) & Variavel(13) & Variavel(21) & Variavel(22) & Variavel(23) & """" & ", "
        SQL = SQL & """" & "" & """" & ")"
        Banco.Execute SQL
        
        VarModeloNota = ModeloTemp
        
        
'    Criterio = "CodPed = " & CodPed
'    Pedidos.FindFirst Criterio
    If pedidos.Fields("Ped.CodPed") = CodPed Then
        pedidos.Edit
        pedidos("DataEmiss") = DtAnt
        pedidos.Update
    End If
    If Index = 1 Then
        On Error GoTo Erro
        LstNota.Visible = True
        LstNota.Clear
        Porta = PortaTemp
        If MsgBox("Avançar um numero na nota ?", vbYesNo) = vbYes Then
                EscreveINI "Nfe", "NroNfe", CCur(MebNfe) + 1, App.Path & "\config.ini"
        End If
        
        If Dir(Caminho & "\Nota.Txt") <> "" Then
            Open Caminho & "\Nota.Txt" For Input As #1
            Do
                Line Input #1, PortaTemp
                LstNota.AddItem PortaTemp
            Loop Until PortaTemp = "<>"
            Close #1
        End If
        SSTabNota.Tab = 2
                       
    End If
    
    If Index = 2 Or Index = 3 Then
    
'        LinhaNota
'        If PortaAcbr = "" And ChkVersao = 1 Then
'            Nfe
'        Else
            NfeAcbr Index
'        End If
    End If
  
    Porta = PortaPedido
  
End If
LimpaCampos
CodPed.SetFocus

Exit Sub

Erro:
If Err.Number = 62 Then
    Close #1
    SSTabNota.Tab = 2
    Porta = PortaPedido
    
End If


End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Conf As Recordset
Set Formulário = Me

'Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
'Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

    Me.Left = 50
    Me.Top = 30

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
   
DatPed.DatabaseName = Caminho & "\Dados.mdb"
'DatNfe.DatabaseName = Caminho & "\Dados.mdb"

DtData.value = Date
   
' Set Nota = Banco.OpenRecordset("Select * From CNT Order by Linha Desc")
Set Configuracao = Banco.OpenRecordset("Select * From CFG")
    
Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 20 Order by Codigo")
Conf.Requery
If Conf.RecordCount > 0 Then
    If IsNumeric(Conf!NomeCampo) Then
        ChkSimples.value = Left(Conf!NomeCampo, 1)
'        ChkFunRural.Value = Mid(Conf!NomeCampo, 2, 1)
        ChkDeducao.value = Mid(Conf!NomeCampo, 3, 1)
        ChkDesconto.value = Mid(Conf!NomeCampo, 3, 1)
    End If
End If

ChkSimples.value = LerINI("Nota", "Simples", App.Path & "\config.ini")
ChkDeducao.value = LerINI("Nota", "Deducao", App.Path & "\config.ini")
ChkDesconto.value = LerINI("Nota", "Desconto", App.Path & "\config.ini")
ChkIcms.value = LerINI("Nota", "CalculaIcms", App.Path & "\config.ini")
'ChkIpi.value = LerINI("Nota", "CalculaIcms", App.Path & "\config.ini")
ChkPcl.value = IIf(LerINI("Nota", "ImprPcl", App.Path & "\config.ini") = "", 0, LerINI("Nota", "ImprPcl", App.Path & "\config.ini"))
If ChkDeducao.value = 1 Then
'    MebReducao = 33.33
    MebReducao = Format(LerINI("Nota", "Reducao", App.Path & "\config.ini"), "###,##0.0000")
    MebReducao.Enabled = True
Else
    MebReducao = ""
    MebReducao.Enabled = False
End If

'If Dir(Caminho & "\ERP.IntegraTxt.exe") <> "" Then
    Confirma(2).Enabled = True
    BtNfe(0).Enabled = True
    BtNfe(1).Enabled = True
'Else
'    Confirma(2).Enabled = False
'    Confirma(0).Enabled = True
'    BtNfe(0).Enabled = False
'    BtNfe(1).Enabled = False
'End If

If LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini") <> "" Then
    If LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini") = 3 Then
        LblRegime.Caption = "3 - Regime Normal"
    ElseIf LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini") = 2 Then
            LblRegime.Caption = "2 - Regime Diferenciado"
        Else
            LblRegime.Caption = "1 - Simples Nacional"
    End If
End If

If LerINI("EMITENTE", "Fone", App.Path & "\nfe.ini") <> "" Then
    VarTelEmite = LerINI("EMITENTE", "Fone", App.Path & "\nfe.ini")
Else
    VarTelEmite = ""
End If

If LerINI("NOTA", "Volume", App.Path & "\Config.ini") <> "" Then
    VarVolume = LerINI("NOTA", "Volume", App.Path & "\Config.ini")
Else
    VarVolume = ""
End If

AbreAcbr "NFE.CertificadoDataVencimento(" & App.Path & "\cfe.ini" & ")"
TrataAcbr "Certificado", 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub TxtObs_GotFocus()
'FrmNota.KeyPreview = False
End Sub

Private Sub TxtObs_LostFocus()
'FrmNota.KeyPreview = True
End Sub


Private Sub BtAtualiza_Click()
If Optdata(0).value = True Then
    DatPed.RecordSource = "SELECT Ped.Cfo,Ped.DtFat, PED.NroNt, PED.CodPed, PED.DataEmiss, PED.Total, PED.Desconto, ENT.Nome, Ent.Cidade, PED.Operacao FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.DataEmiss =# " & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed DESC"
Else
    DatPed.RecordSource = "SELECT Ped.Cfo,Ped.DtFat, PED.NroNt, PED.CodPed, PED.DataEmiss, PED.Total, PED.Desconto, ENT.Nome, Ent.Cidade, PED.Operacao FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.Dtfat=# " & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed DESC"
End If

DatPed.Refresh
If DatPed.Recordset.RecordCount <> 0 Then
     DatPed.Recordset.MoveFirst
End If

End Sub


Private Sub GridNfe_dblClick()
If DatNfe.Recordset.RecordCount > 0 Then
    TxtChave = GridNfe.Columns(1)
    TxtProtocolo = GridNfe.Columns(3)
    txtStatus = GridNfe.Columns(2)
'    TxtRecibo = GridNfe.Columns(0)
End If
End Sub


Private Sub OptDest_Click()
If OptDest.value = True Then
    LblFrete = "FOB"
Else
    LblFrete = "CIF"
End If

End Sub

Private Sub OptRem_Click()
If OptRem.value = True Then
    LblFrete = "CIF"
Else
    LblFrete = "FOB"
End If
End Sub

Private Sub TxtCancela_Change()
Lbltam.Caption = Len(TxtCancela)
End Sub

Private Sub TxtChave_Change()
    QtdChave.Caption = Len(Trim(TxtChave))
End Sub

' ************************************************************************************
' ************************************************************************************
' ************************************************************************************
' ************************************************************************************
' ************************************************************************************
' ************************************************************************************
' ************************************************************************************

Private Sub NfeAcbr(Index As Integer)
Dim VarCnpjEmite As String
Dim VarUfEmite As String
Dim VarIbgeEmite As String
Dim VarRazaoEmite As String
Dim VarFantasiaEmite As String
Dim VarEndEmite As String
Dim VarNroEmite As String
Dim VarBairroEmite As String
Dim VarCidadeEmite As String
Dim VarInscEmite As String
Dim VarCepEmite As String
Dim VarNota As Recordset
Dim TipoPgto As String
Dim VarNat As String
Dim VarFinalidade As String
Dim VarTipo As String
Dim VarPis As String
Dim VarCofins As String
Dim VarSitTrib As String
Dim VarSitTrib2 As String
Dim VarRecibo As String
Dim VarChave As String
Dim VarStatus As String
Dim VarProtocolo As String
Dim varMensagem As String
Dim VarPrecoBase As String
Dim VarCfop As String
Dim VarSitTribFixo As String
Dim VarTemp As String
Dim VarObsDeducao As String
Dim ObsItem As String
Dim VarIpi As String
Dim VarBaseST As String
Dim VarBaseTemp As String
Dim VarInscrMuni As String
Dim VarCNAE As String
Dim VarRegimeTrib As String
Dim VarControle As Boolean
Dim VarNotaValida As Boolean
Dim VarEmail As String
Dim VarAliqPis  As String
Dim VarAliqCofins As String
Dim VarBarra As Recordset
Dim VarTotCofins As String
Dim VarTotPis As String
Dim BasePisCofins As String
Dim nRec As String
Dim NroLote As String
Dim VarDesonerado As Currency
Dim VarDesc As Currency
Dim VarContPrd As Integer
Dim Criterio As String
Dim VarImposto As Currency
Dim VarImpostoNac As Currency
Dim VarImpostoEst As Currency
Dim VarTotImpostoNac As Currency
Dim VarTotImpostoEst  As Currency
Dim VarObsFinal As String
Dim VarPercIva As Currency
Dim VarPag As String
Dim VarPercDeducao As Currency
Dim VarUfEstado As Recordset
Dim VarIcmsEstado As Integer
Dim CodChave As Double
Dim VarDescTemp As Currency
Dim VarTotIpi As Currency
Dim VarContItem As Integer
Dim VarMonoRet As String
Dim VarTotMonoRet As String
Dim VarTotQtdMono As String
Dim VarTotCbs As String
Dim VarTotIbs As String

VarTotCbs = 0
VarTotIbs = 0
VarMonoRet = 0
VarTotMonoRet = 0
VarTotQtdMono = 0

VarPercDeducao = 0
VarTotIpi = 0

DatMinutos.DatabaseName = Caminho & "\Minutos.mdb"
DatMinutos.RecordSource = "SELECT Avg(Ncm.aliqNac) AS MediaDealiqNac, Avg(Ncm.aliqEst) AS MediaDealiqEst FROM Ncm"
DatMinutos.Refresh
VarMediaImpostoNac = DatMinutos.Recordset.Fields("MediaDealiqNac")
VarMediaImpostoEst = DatMinutos.Recordset.Fields("MediaDealiqEst")

DatMinutos.RecordSource = "Select codigo, AliqNac, AliqEst From Ncm Order By Codigo"
DatMinutos.Refresh
VarTotImpostoNac = 0
VarTotImpostoEst = 0
VarTotPis = 0
VarTotCofins = 0
BasePisCofins = 0
VarDesonerado = 0

VarEmail = ""
VarSerieNfe = ""

If LerINI("Nfe", "SerieNfe", App.Path & "\config.ini") <> "" Then
    VarSerieNfe = LerINI("Nfe", "SerieNfe", App.Path & "\config.ini")
End If

If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
    VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
End If

If LerINI("EMITENTE", "Uf", App.Path & "\nfe.ini") <> "" Then
    VarUfEmite = LerINI("EMITENTE", "Uf", App.Path & "\nfe.ini")
End If

If LerINI("EMITENTE", "CodigoIbge", App.Path & "\nfe.ini") <> "" Then
    VarIbgeEmite = LerINI("EMITENTE", "CodigoIbge", App.Path & "\nfe.ini")
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

If LerINI("EMITENTE", "Pis", App.Path & "\nfe.ini") <> "" Then
    VarPis = LerINI("EMITENTE", "Pis", App.Path & "\nfe.ini")
Else
    VarPis = ""
End If

If LerINI("EMITENTE", "AliqPis", App.Path & "\nfe.ini") <> "" Then
    VarAliqPis = LerINI("EMITENTE", "AliqPis", App.Path & "\nfe.ini")
Else
    VarAliqPis = ""
End If

If LerINI("EMITENTE", "AliqCofins", App.Path & "\nfe.ini") <> "" Then
    VarAliqCofins = LerINI("EMITENTE", "AliqCofins", App.Path & "\nfe.ini")
Else
    VarAliqCofins = ""
End If

If LerINI("EMITENTE", "Cofins", App.Path & "\nfe.ini") <> "" Then
    VarCofins = LerINI("EMITENTE", "Cofins", App.Path & "\nfe.ini")
Else
    VarCofins = ""
End If

If LerINI("EMITENTE", "SitTrib", App.Path & "\nfe.ini") <> "" Then
    VarSitTrib = LerINI("EMITENTE", "SitTrib", App.Path & "\nfe.ini")
Else
    VarSitTrib = ""
End If

If LerINI("EMITENTE", "InscrMuni", App.Path & "\nfe.ini") <> "" Then
    VarInscrMuni = LerINI("EMITENTE", "InscrMuni", App.Path & "\nfe.ini")
Else
    VarInscrMuni = Space(15)
End If

If LerINI("EMITENTE", "Cnae", App.Path & "\nfe.ini") <> "" Then
    VarCNAE = LerINI("EMITENTE", "Cnae", App.Path & "\nfe.ini")
Else
    VarCNAE = Space(7)
End If

If LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini") <> "" Then
    VarRegimeTrib = LerINI("EMITENTE", "RegimeTrib", App.Path & "\nfe.ini")
Else
    VarRegimeTrib = ""
End If

VarSitTribFixo = ""
VarCfop = ""

If LerINI("FIXO", "SitTribFixo", App.Path & "\nfe.ini") <> "" Then
    VarSitTribFixo = LerINI("FIXO", "SitTribFixo", App.Path & "\nfe.ini")
Else
    VarSitTribFixo = ""
End If

If LerINI("FIXO", "Cfop", App.Path & "\nfe.ini") <> "" Then
    VarCfop = LerINI("FIXO", "Cfop", App.Path & "\nfe.ini")
Else
    VarCfop = ""
End If

If LerINI("FIXO", "Ipi", App.Path & "\nfe.ini") <> "" Then
    VarIpi = LerINI("FIXO", "Ipi", App.Path & "\nfe.ini")
Else
    VarIpi = ""
End If

CodChave = Int(Mid(Int(CCur(Mid(Time, 7, 2)) * (Rnd * 100000)), 1, 6))
Set VarNota = Banco.OpenRecordset("SELECT PED.*, ITP.*, NOTA.*, PED.CodPed,itp.Complemento, ITP.CodItp, ENT.Classificacao, ENT.InscrProd, ENT.Matricula, ENT.Fantasia, ENT.Fone, ENT.CPF, ENT.InscrEst, ENT.CGC, ENT.Email, ENT.InscrProd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Icm, PRD.Ipi, PRD.CfOpPrd, PRD.SitTrib, PRD.Classe, PRD.ClasseFiscal, PRD.Iva, Prd.Redutor, Prd.CodBar, Prd.Fator,Prd.CodAbc FROM (PRD INNER JOIN (ENT INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON ENT.CodEntidade = PED.Entidade) ON PRD.CodPrd = ITP.Produto) INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed WHERE PED.CodPed = " & CodPed & "  ORDER BY PED.CodPed, ITP.CodItp")
VarNota.Requery

If VarNota.RecordCount = 0 Then
    MsgBox "Houve alguma falha, arquivo sem registros !", vbInformation
    Confirma(2).Enabled = True
    Exit Sub
End If

LimpaIni

If VarPercDeducao <> 0 Then
    ChkDeducao.value = 1
    MebReducao = VarPercDeducao
End If

VarNota.MoveFirst
Select Case VarNota!Operacao
    Case 17
        TipoPgto = "0"
    Case 2
        TipoPgto = "1"
    Case 4
        TipoPgto = "0"
    Case Else
        TipoPgto = "0"
End Select

If VarNota!Operacao = 0 Or VarNota!Operacao = 9 Or VarNota!Operacao = 10 Or VarNota!Operacao = 12 Then
    VarNat = 0
Else
    VarNat = 1
End If

If ChkCompl.value = 0 Then
    VarFinalidade = 1
Else
    VarFinalidade = 2
End If
If TxtChave <> "" And ChkCompl = 0 Then
    VarFinalidade = 4
End If

If ChkEstorno = 1 Then
    VarFinalidade = 3
End If

If VarSerieNota = "" Then
    VarSerieNota = "0"
End If

If VarSerieNfe <> "" Then
    VarSerieNota = VarSerieNfe
End If
EscreveINI "Identificacao", "natOp", DbcCfo, App.Path & "\cfe.ini"
EscreveINI "Identificacao", "mod", "55", App.Path & "\cfe.ini"
EscreveINI "Identificacao", "serie", VarSerieNota, App.Path & "\cfe.ini"
EscreveINI "Identificacao", "dEmi", Format(Emissao, "dd/mm/yyyy"), App.Path & "\cfe.ini"
EscreveINI "Identificacao", "dSaiEnt", IIf(Saida = "__/__/____", Space(10), Format(Saida, "dd/mm/yyyy")), App.Path & "\cfe.ini"
EscreveINI "Identificacao", "hSaiEnt", Format(HoraSaida, "HH:MM:SS"), App.Path & "\cfe.ini"
EscreveINI "Identificacao", "nNF", psCompString(1, MebNfe, 9), App.Path & "\cfe.ini"
EscreveINI "Identificacao", "cNF", psCompString(1, CodChave, 9), App.Path & "\cfe.ini"
EscreveINI "Identificacao", "tpNF", VarNat, App.Path & "\cfe.ini"
EscreveINI "Identificacao", "indPag", TipoPgto, App.Path & "\cfe.ini"

VarInscri = VarNota!InscrEst

If pedidos!Classificacao = 1 Then
    EscreveINI "Identificacao", "indFinal", 1, App.Path & "\cfe.ini"
ElseIf VarInscri = "ISENTO" Or VarInscri = "" Then
       EscreveINI "Identificacao", "indFinal", 1, App.Path & "\cfe.ini"
    Else
       EscreveINI "Identificacao", "indFinal", 0, App.Path & "\cfe.ini"
End If

If ChkDoc.value = 1 Then
    EscreveINI "Identificacao", "indFinal", 1, App.Path & "\cfe.ini"
End If

EscreveINI "Identificacao", "finNfe", VarFinalidade, App.Path & "\cfe.ini"

If VarPais = "BRASIL" Then
    If TabSql("ENT.Uf") <> "SP" Then
    
            If pedidos!Classificacao = 1 Then
               EscreveINI "Identificacao", "IdDest", "2", App.Path & "\cfe.ini"
            ElseIf VarInscri = "ISENTO" Or VarInscri = "" Then
               EscreveINI "Identificacao", "IdDest", "2", App.Path & "\cfe.ini"
            Else
               EscreveINI "Identificacao", "IdDest", "2", App.Path & "\cfe.ini"
            End If
    Else
         EscreveINI "Identificacao", "IdDest", "1", App.Path & "\cfe.ini"
    End If
Else
    EscreveINI "Identificacao", "IdDest", "3", App.Path & "\cfe.ini"
End If


If TxtChave <> "" Then
     EscreveINI "NFRef001", "refNFE", Trim(TxtChave), App.Path & "\cfe.ini"
End If

EscreveINI "Emitente", "Cnpj", psCompString(1, VarCnpjEmite, 14), App.Path & "\cfe.ini"
EscreveINI "Emitente", "IE", psCompString(3, VarInscEmite, 14), App.Path & "\cfe.ini"
EscreveINI "Emitente", "xNome", psCompString(3, VarRazaoEmite, 60), App.Path & "\cfe.ini"
EscreveINI "Emitente", "xFant", psCompString(3, VarFantasiaEmite, 60), App.Path & "\cfe.ini"
EscreveINI "Emitente", "Fone", psCompString(3, VarTelEmite, 10), App.Path & "\cfe.ini"
EscreveINI "Emitente", "CEP", psCompString(1, VarCepEmite, 8), App.Path & "\cfe.ini"
EscreveINI "Emitente", "xLgr", psCompString(3, VarEndEmite, 60), App.Path & "\cfe.ini"
EscreveINI "Emitente", "nro", psCompString(3, VarNroEmite, 10), App.Path & "\cfe.ini"
EscreveINI "Emitente", "cMun", Trim(VarIbgeEmite), App.Path & "\cfe.ini"
EscreveINI "Emitente", "xMun", psCompString(3, VarCidadeEmite, 60), App.Path & "\cfe.ini"
EscreveINI "Emitente", "xBairro", psCompString(3, VarBairroEmite, 60), App.Path & "\cfe.ini"
EscreveINI "Emitente", "UF", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
EscreveINI "Emitente", "CRT", VarRegimeTrib, App.Path & "\cfe.ini"

Dim varTeste As Currency
varTeste = 0
If VarPais = "BRASIL" Then
    If VarNota!Classificacao = 1 Then
        If TabSql("ENT.Uf") <> "SP" Then
            VarTipo = "9"
        Else
            VarTipo = "2"
        End If
        EscreveINI "Destinatario", "CNPJCPF", psCompString(3, Trim(LIMPACampo2(VarNota!Documento)), 11), App.Path & "\cfe.ini"
    Else
        EscreveINI "Destinatario", "CNPJCPF", psCompString(3, Trim(LIMPACampo2(VarNota!Documento)), 14), App.Path & "\cfe.ini"
    End If
    
    If VarNota!Classificacao = 2 Then
        VarTipo = "1"
        If VarNota!InscrEst = "ISENTO" Then
            EscreveINI "Destinatario", "IE", "ISENTO", App.Path & "\cfe.ini"
        Else
            EscreveINI "Destinatario", "IE", psCompString(3, Trim(Left(LIMPACampo2(VarNota!InscrEst), 14)), 14), App.Path & "\cfe.ini"
        End If
        VarInscri = VarNota!InscrEst
    Else
        If VarNota!InscrProd <> "" Then
            EscreveINI "Destinatario", "IE", psCompString(3, Trim(Left(LIMPACampo2(VarNota!InscrEst), 14)), 14), App.Path & "\cfe.ini"
        Else
            EscreveINI "Destinatario", "IE", Space(14), App.Path & "\cfe.ini"
        End If
    End If
Else
    EscreveINI "Destinatario", "CNPJCPF", "", App.Path & "\cfe.ini"
End If

EscreveINI "Destinatario", "xNome", psCompString(3, VarNota!RazaoSocial, 60), App.Path & "\cfe.ini"
EscreveINI "Destinatario", "XLgr", psCompString(3, VarNota!Logradouro, 60), App.Path & "\cfe.ini"
EscreveINI "Destinatario", "nro", psCompString(3, VarNota!Nro, 10), App.Path & "\cfe.ini"
EscreveINI "Destinatario", "xBairro", psCompString(3, VarNota!Bairro, 60), App.Path & "\cfe.ini"

If VarPais = "BRASIL" Then
    EscreveINI "Destinatario", "xMun", psCompString(3, VarNota!Cidade, 60), App.Path & "\cfe.ini"
    EscreveINI "Destinatario", "cMun", psCompString(1, Left(VarNota!matricula, 7), 7), App.Path & "\cfe.ini"
    EscreveINI "Destinatario", "CEP", psCompString(2, VarNota!Cep, 8), App.Path & "\cfe.ini"
Else
    FuncaoExterior psCompString(3, Trim(LIMPACampo2(VarNota!Documento)), 11)
End If
EscreveINI "Destinatario", "UF", psCompString(3, VarNota.Fields("UfNota"), 2), App.Path & "\cfe.ini"

EscreveINI "Destinatario", "xPais", psCompString(3, VarPais, 60), App.Path & "\cfe.ini"
EscreveINI "Destinatario", "fone", psCompString(3, Trim(Left(LIMPACampo2(VarNota!Fone), 11)), 11), App.Path & "\cfe.ini"
If Not IsNull(VarNota!Email) Then
    EscreveINI "Destinatario", "email", psCompString(3, VarNota!Email, 60), App.Path & "\cfe.ini"
End If

If pedidos!Classificacao = 1 Then
   EscreveINI "Destinatario", "indIEDest", "9", App.Path & "\cfe.ini"
Else
   If TabSql("ENT.Uf") = "SP" Then
        If VarInscri = "ISENTO" Or VarInscri = "" Then
            EscreveINI "Destinatario", "indIEDest", "2", App.Path & "\cfe.ini"
        Else
            EscreveINI "Destinatario", "indIEDest", "1", App.Path & "\cfe.ini"
        End If
   Else
        If VarInscri = "ISENTO" Or VarInscri = "" Then
            EscreveINI "Destinatario", "indIEDest", "9", App.Path & "\cfe.ini"
        Else
            EscreveINI "Destinatario", "indIEDest", "1", App.Path & "\cfe.ini"
        End If
   End If
End If
If Not IsNull(VarNota!Email) Then
    VarEmail = VarNota!Email
End If

VarContPrd = 1
VarContItem = 1
VarNota.MoveFirst
Dim VarUfIva As Recordset

While Not VarNota.Eof
    VarTotIpi = 0
    Set VarBarra = Banco.OpenRecordset("SELECT BAR.*, BAR.Produto FROM BAR WHERE (((BAR.Produto)= " & VarNota!Produto & ") AND ((Len((Trim([bar].[barbarra]))))=13)) ")
    VarBarra.Requery
        
    Criterio = "Codigo = " & """" & Trim(LIMPACampo2(VarNota!Classe)) & """"
    DatMinutos.Recordset.FindFirst Criterio
    If Not DatMinutos.Recordset.NoMatch Then
        VarImpostoNac = ((VarNota!Qtd * VarNota!ValorUnit) * (DatMinutos.Recordset("AliqNac") / 100))
        VarImpostoEst = ((VarNota!Qtd * VarNota!ValorUnit) * (DatMinutos.Recordset("AliqEst") / 100))
        VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
        VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
    Else
        VarImpostoNac = ((VarNota!Qtd * VarNota!ValorUnit) * (VarMediaImpostoNac / 100))
        VarImpostoEst = ((VarNota!Qtd * VarNota!ValorUnit) * (VarMediaImpostoEst / 100))
        VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
        VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
    End If
    
    VarCfop = ""
    If Trim(VarNota.Fields("Itp.Complemento")) = "BONIFICACAO" Then
        If TabSql("ENT.Uf") = "SP" Then
            VarCfop = "5.910"
        Else
            VarCfop = "6.910"
        End If
    Else
        VarSitTrib = ""
    End If
        
    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cProd", VarNota!Produto, App.Path & "\cfe.ini"
    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "xProd", VarNota!Descrição_Produto, App.Path & "\cfe.ini"
    
    If Not IsNull(VarNota!CodBar) And Len(Trim(VarNota!CodBar)) = 13 And Len(Trim(VarNota!CodBar)) <> "" And ChKBar.value = 1 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEan", Trim(VarNota!CodBar), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEan", "SEM GTIN", App.Path & "\cfe.ini"
    End If
    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", psCompString(3, Trim(LIMPACampo2(VarNota!Classe)), 8), App.Path & "\cfe.ini"
    
    If VarNota.Fields("ClasseFiscal") <> "" Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CEST", psCompString(3, Trim(LIMPACampo2(VarNota.Fields("ClasseFiscal"))), 8), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CEST", psCompString(3, "0000000", 8), App.Path & "\cfe.ini"
    End If
    
    If ChkSub.value = 1 Then
        If VarCfop <> "" Then
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(VarCfop)), 4), App.Path & "\cfe.ini"
        Else
            If ChkCfop = 0 Then
                If VarNota!CfopPrd = 0 Then
                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(TxtCfop)), 4), App.Path & "\cfe.ini"

                Else
                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(VarNota!CfopPrd)), 4), App.Path & "\cfe.ini"
                End If
            Else
                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(TxtCfop)), 4), App.Path & "\cfe.ini"
            End If
        End If
    Else
       If VarCfop <> "" Then
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(VarCfop)), 4), App.Path & "\cfe.ini"
       Else
            If ChkCfop = 0 Then    ' Trim(LIMPACampo2(Txtcfop)) = "5405" Or Trim(LIMPACampo2(Txtcfop)) = "5102" Then
                If VarNota!CfopPrd = 0 Then
                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(TxtCfop)), 4), App.Path & "\cfe.ini"
                Else
                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(VarNota!CfopPrd)), 4), App.Path & "\cfe.ini"
                End If
            Else
                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", psCompString(3, Trim(LIMPACampo2(TxtCfop)), 4), App.Path & "\cfe.ini"
            End If
    
       End If
    End If
    If VarNota!Embalagem <> "" Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "uCom", psCompString(3, VarNota!Embalagem, 6), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "uCom", "UN", App.Path & "\cfe.ini"
    End If
    
    If Not IsNull(VarNota!CodBar) And Len(Trim(VarNota!CodBar)) = 13 And Len(Trim(VarNota!CodBar)) <> "" And ChKBar.value = 1 Then
        
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEANTrib", Trim(VarNota!CodBar), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEanTrib", "SEM GTIN", App.Path & "\cfe.ini"
        
    End If
    
    If VarFinalidade = 2 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "qCom", Numero4(0), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "qCom", Numero4(VarNota!Qtd), App.Path & "\cfe.ini"
    End If


    If ChkValor.value = 0 Then
        If VarFinalidade = 2 Then
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", Numero2(0), App.Path & "\cfe.ini"
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vProd", Numero2(0), App.Path & "\cfe.ini"
        Else
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", Numero4(VarNota!VlrVdaDia / VarNota!Qtd), App.Path & "\cfe.ini"
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vProd", Numero2((VarNota!VlrVdaDia / VarNota!Qtd) * VarNota!Qtd), App.Path & "\cfe.ini"
        End If
        VarPrecoBase = (VarNota!VlrVdaDia)
        BasePisCofins = (VarNota!VlrVdaDia)
        VarTotIpi = (VarNota!VlrVdaDia) * (VarNota("Prd.Ipi") / 100)
    Else
        If VarFinalidade = 2 Then
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", Numero2(0), App.Path & "\cfe.ini"
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vProd", Numero2(0), App.Path & "\cfe.ini"
        Else
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", Numero4(VarNota!ValorUnit), App.Path & "\cfe.ini"
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vProd", Numero2(VarNota!ValorUnit * VarNota!Qtd), App.Path & "\cfe.ini"
        End If
        VarPrecoBase = VarNota!ValorUnit * VarNota!Qtd
        BasePisCofins = VarNota!ValorUnit * VarNota!Qtd
        VarTotIpi = (VarNota!ValorUnit * VarNota!Qtd) * (VarNota("Prd.Ipi") / 100)
    End If
    
    If VarIpiDev <> "" Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "pDevol", Numero2(VarPIpi), App.Path & "\cfe.ini"
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vIPIDevol", Numero2((VarNota!ValorUnit * VarNota!Qtd) * (VarPIpi / 100)), App.Path & "\cfe.ini"
    End If
    
    VarBaseTemp = VarPrecoBase
    If ChkDeducao.value = 1 Then
        VarPrecoBase = VarPrecoBase - (VarPrecoBase * (MebReducao / 100))
    End If
        
    If VarNota.Fields("Nota.ValorFrete") <> 0 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vFrete", Numero2(VarNota.Fields("Nota.ValorFrete") / ContadorItens), App.Path & "\cfe.ini"
    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vFrete", Numero2(0), App.Path & "\cfe.ini"
    End If
    
    If VarNota.Fields("TotOutras") <> 0 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vOutro", Numero2(VarNota.Fields("TotOutras") / ContadorItens), App.Path & "\cfe.ini"

    Else
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vOutro", Numero2(0), App.Path & "\cfe.ini"
    End If
    
    If ChkDesconto.value = 0 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vDesc", Numero2(0), App.Path & "\cfe.ini"
    Else
        VarDescTemp = (TxtDesconto / (CCur(MebValor) + TxtDesconto))
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vDesc", Numero2(((VarNota!VlrVdaDia / VarNota!Qtd) * VarNota!Qtd) * VarDescTemp), App.Path & "\cfe.ini"

    End If
        
    ObsItem = Trim(VarNota.Fields("Itp.Complemento")) & " " & Trim(VarNota!Referencia)
    
    If VarNota!NroPedido <> "" Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "xPed", Trim(VarNota!NroPedido), App.Path & "\cfe.ini"
        If VarNota!Referencia <> "" Then
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "nItemPed", Trim(VarNota!Referencia), App.Path & "\cfe.ini"
        Else
            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "nItemPed", psCompString(1, VarContPrd, 6), App.Path & "\cfe.ini"
        End If
    End If
    
    If Len(ObsItem) > 1 Then
        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "infAdProd", Trim(ObsItem), App.Path & "\cfe.ini"
    End If
   
 '   VarSitTrib2 = Left(Trim(VarNota!SitTrib), 3)
            
    VarTemp = VarSitTrib
    If VarSitTrib <> "" Then
        EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"

    Else
        VarSitTrib = Left(Trim(VarNota!SitTrib), 3)

        If ChkSub.value = 1 And VarSitTribFixo <> "" Then
          VarSitTrib = VarSitTribFixo
        End If
 
        
        If LblRegime.Caption = "3 - Regime Normal" Then
            
            If ChkDeducao.value = 1 Then
                If UfEntidade = "SP" Then
                    VarSitTrib = "020"
                    CmbObsRoda = "REDUCAO DA BASE DE CALCULO CONFORME ART.52 ANEXO II DO RICMS"
                End If
            End If
        
            If TxtCfop = "5.901" Or TxtCfop = "5901" Then
                VarSitTrib = "050"
                VarPis = "99"
            End If
            
            If TxtCfop = "5.949" Or TxtCfop = "5949" Then
                VarSitTrib = "041"
                VarCofins = "99"
            End If
        
            If TxtCfop = "5.912" Or TxtCfop = "5912" Then
                VarSitTrib = "041"
                VarCofins = "08"
                VarPis = "08"
            End If
        
            If TxtCfop = "6.912" Or TxtCfop = "6912" Then
                VarSitTrib = "041"
                VarCofins = "08"
                VarPis = "08"
            End If
        
        End If
                        
        If VarIcms = 0 Then
            IcmsNota = VarNota!Icm
        Else
            IcmsNota = VarIcms
        End If
        
        If MebReducao = "" Then
            MebReducao = 0
        End If
        
        If MebSub = "" Then
            MebSub = 0
        End If
        
        If ChkSub = 1 Then
            Set VarUfEstado = Banco.OpenRecordset("Select * from ICM Where Uf = " & """" & Trim(VarNota("UfNota")) & """")
            VarUfEstado.Requery
            If VarUfEstado.RecordCount > 0 Then
                VarIcmsEstado = VarUfEstado!Icms
            End If
            If VarSub And VarSub2 = True Then
                Set VarUfIva = Banco.OpenRecordset("Select * from IVA Where Uf = " & """" & Trim(VarNota("UfNota")) & """" & " AND Ncm = " & """" & Trim(VarNota("Classe")) & """")  '& " AND Cest = " & """" & Trim(VarNota.Fields("ClasseFiscal")) & """")
                VarUfIva.Requery
                If VarUfIva.RecordCount > 0 Then
                    VarPercIva = VarUfIva!Iva
                    IcmsNota = VarUfIva!Icms
                    VarIcmsEstado = VarUfIva!Icms
                Else
                    VarPercIva = 0
                End If
            Else
                VarPercIva = VarNota!Iva
            End If
        Else
        End If
                
        If Trim(LIMPACampo2(VarNota!CfopPrd)) = "5656" Or Trim(LIMPACampo2(VarNota!CfopPrd)) = "5.656" Or Trim(LIMPACampo2(VarNota!CfopPrd)) = "5655" Or Trim(LIMPACampo2(VarNota!CfopPrd)) = "5.655" Or Trim(LIMPACampo2(VarNota!CfopPrd)) = "5661" Or Trim(LIMPACampo2(VarNota!CfopPrd)) = "6661" Then
            If psCompString(3, Trim(LIMPACampo2(VarNota!Classe)), 8) = "27111910" Then
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "210203001", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "GLP", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGLP", "40.00", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNn", "30.00", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNi", "30.00", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "vPart", "5.00", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                
                EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "indImport", "0", App.Path & "\cfe.ini"
                EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "cUFOrig", "35", App.Path & "\cfe.ini"
                EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "pOrig", "100.00", App.Path & "\cfe.ini"
                VarMonoRet = 1.4139
            End If
            
            If psCompString(3, Trim(LIMPACampo2(VarNota!Classe)), 8) = "27101932" Then
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "620505001", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "OUTROS OLEOS LUBRIFICANTES AUTOMOTIVOS", App.Path & "\cfe.ini"
                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                VarMonoRet = 1.3721
            End If
        End If

        If LblRegime.Caption = "3 - Regime Normal" Or LblRegime.Caption = "2 - Regime Diferenciado" Then

            Select Case VarSitTrib
                Case "102"
                    VarSitTrib = "000"
                Case "500"
                    VarSitTrib = "060"
                Case "400"
                    VarSitTrib = "040"
                Case "900"
                    VarSitTrib = "090"
             End Select
        Else
            Select Case VarSitTrib
                Case "000"
                    VarSitTrib = "102"
                Case "060"
                    VarSitTrib = "500"
                Case "040"
                    VarSitTrib = "400"
                Case "090"
                    VarSitTrib = "900"
             End Select
        End If
                
                
    If VarSitTrib = "020" Or VarSitTrib = "040" Or VarSitTrib = "041" Or VarSitTrib = "103" Or VarSitTrib = "050" Then
            If Left(VarNota!CodAbc, 2) = "SP" Then
                 EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cBenef", VarNota!CodAbc, App.Path & "\cfe.ini"
            End If
    End If
                
        Select Case VarSitTrib
            Case "000"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBC", Numero2(VarPrecoBase), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                
                   
            Case "020"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                If VarFinalidade = 2 Then
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "VBC", Numero2(0), App.Path & "\cfe.ini"
                Else
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "VBC", Numero2(VarPrecoBase), App.Path & "\cfe.ini"
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", Numero2(Str(IcmsNota)), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pRedBC", Numero2(MebReducao), App.Path & "\cfe.ini"

'              EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
'                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                VarDesonerado = VarDesonerado + (IcmsNota * VarPrecoBase) / 100
                         
            Case "040"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                VarDesonerado = VarDesonerado + (IcmsNota * VarPrecoBase) / 100
                
            Case "041"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                
            Case "050"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSDeson", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "motDesICMS", 9, App.Path & "\cfe.ini"
            Case "051"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                
            Case "061"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "qBCMonoRet", Numero4(VarNota!Qtd), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "adRemICMSRet", Numero4(VarMonoRet), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSMonoRet", Numero2(VarNota!Qtd * VarMonoRet), App.Path & "\cfe.ini"
                VarTotMonoRet = VarTotMonoRet + VarNota!Qtd * VarMonoRet
                VarTotQtdMono = VarTotQtdMono + VarNota!Qtd

            Case "060"
                If IsNumeric(VarPercIva) Then
                    If VarPercIva > 0 Then
                        VarBaseST = VarPrecoBase + ((VarPrecoBase * VarPercIva) / 100)
                    Else
                        VarBaseST = VarPrecoBase
                    End If
                Else
                        VarBaseST = VarPrecoBase
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 2), App.Path & "\cfe.ini"
                        
            Case "101"
            
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pCredSN", Numero2(MebSub), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vCredICMSSN", Numero2((MebSub * VarPrecoBase) / 100), App.Path & "\cfe.ini"
           
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
                
            Case "102"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
            Case "103"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                
            Case "201"
            
                Dim VarRedST As Currency
                                       
                If VarNota.Fields("Nota.ValorFrete") <> 0 Then
                    VarPrecoBase = VarPrecoBase + (VarNota.Fields("Nota.ValorFrete") / ContadorItens)
                End If
                If VarRedutor = 0 Then
                    If VarNota.Fields("Redutor") <> 0 Then
                        VarRedST = (VarPrecoBase / 100) * 12
                    Else
                        VarRedST = (VarPrecoBase / 100) * IcmsNota
                    End If
                Else
                    VarRedST = (VarPrecoBase / 100) * VarRedutor
                End If
                VarBaseST = VarPrecoBase + ((VarPrecoBase * VarPercIva) / 100)
                
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pMVAST", Str(VarPercIva), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pRedBCST", Numero2(MebReducao), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCST", Numero2(VarBaseST), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSST", Numero2((VarBaseST * (IcmsNota / 100)) - VarRedST), App.Path & "\cfe.ini"  ' Numero2((VarBaseST * (IcmsNota / 100)) - VarRedST), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pCredSN", psCompString(2, Numero2(MebSub), 5), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vCredICMSSN", Numero2((MebSub * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
            
            Case "202"
                If VarNota.Fields("Nota.ValorFrete") <> 0 Then
                    VarPrecoBase = VarPrecoBase + (VarNota.Fields("Nota.ValorFrete") / ContadorItens)
                End If
                VarBaseST = VarPrecoBase + ((VarPrecoBase * VarPercIva) / 100)
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pMVAST", IIf(VarPercIva = "0,01", Numero2(0), Str(VarPercIva)), App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pRedBCST", Numero2(MebReducao), App.Path & "\cfe.ini"
                If VarSub = True Then
                       If VarUfIva!Icms <> 20 Then
                            VarImpSub = (VarBaseSub2 * (VarUfIva!Icms / 100)) - VarIcmsSub
                       Else
                            VarImpSub = (VarBaseSub2 * (VarUfIva!Icms / 100))
                       End If
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(0), App.Path & "\cfe.ini"
                        EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCST", Numero2(VarBaseST), App.Path & "\cfe.ini"
                        EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSST", Numero2((VarBaseST * (Str(IcmsNota) / 100) - (VarPrecoBase * (VarIcmsEstado / 100)))), App.Path & "\cfe.ini"

                    ValorSt = ValorSt + (VarBaseST * (Str(IcmsNota) / 100) - (VarPrecoBase * (VarIcmsEstado / 100)))
                Else
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCST", Numero2(VarBaseST), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSST", Numero2((VarBaseST - VarPrecoBase) * (Str(IcmsNota) / 100)), App.Path & "\cfe.ini"
                    ValorSt = ValorSt + (VarBaseST - VarPrecoBase) * (Str(IcmsNota) / 100)
                End If
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                End If
            Case "300"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                
                
            Case "400"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                
            Case "500"
                If psCompString(3, Trim(LIMPACampo2(VarNota!Classe)), 8) = "27111910" Then
                    
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right("061", 2), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "qBCMonoRet", Numero4(VarNota!Qtd), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "adRemICMSRet", Numero4(VarMonoRet), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSMonoRet", Numero2(VarNota!Qtd * VarMonoRet), App.Path & "\cfe.ini"
                    VarTotMonoRet = VarTotMonoRet + VarNota!Qtd * VarMonoRet
                    VarTotQtdMono = VarTotQtdMono + VarNota!Qtd
                    
                Else
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCSTReT", "0.0001", App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pST", "0.00001", App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSSubstituto", "0.0001", App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSSTRet", "0.0001", App.Path & "\cfe.ini"
                End If
            Case "900"
                If ChkIcms.value = 0 Then
                    VarPrecoBase = 0
                Else
                    VarDescTemp = (TxtDesconto / (CCur(MebValor) + TxtDesconto))
                    VarPrecoBase = CCur(VarPrecoBase) - (((VarNota!VlrVdaDia / VarNota!Qtd) * VarNota!Qtd) * VarDescTemp) ' TxtDesconto
                
                    If ChkSomar.value = 0 <> 0 Then
                           VarPrecoBase = VarPrecoBase + (VarNota.Fields("Nota.ValorFrete") / ContadorItens)
                    End If
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", Left(VarSitTrib, 3), App.Path & "\cfe.ini"

                If BaseIcmsDif <> "" Then
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBC", Numero2(BaseIcmsDif), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2(ValorIcmDif), App.Path & "\cfe.ini"
                Else
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBC", Numero2(VarPrecoBase), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMS", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
                End If
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pRedBC", Numero2(MebReducao), App.Path & "\cfe.ini"
                VarBaseST = 0
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "modBCST", "4", App.Path & "\cfe.ini"
                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pMVAST", Str(VarPercIva), App.Path & "\cfe.ini"
                If txtbasesub <> "" Then
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCST", Numero2(txtbasesub), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(Format(IcmsNota, "##0.00")), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSST", Numero2(txticmssub), App.Path & "\cfe.ini"
                Else
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vBCST", Numero2(VarBaseST), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMSST", Numero2(0), App.Path & "\cfe.ini"
                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSST", Numero2(0), App.Path & "\cfe.ini"
                End If
                
        End Select
        If VarRegimeTrib = 3 Then
           If VarSitTrib = "000" Then
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "CSTIS", Right(VarSitTrib, 3), App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "cClassTribIs", "000001", App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "vBCIS", 0, App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "pIS", 0, App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "pISEspec", 0, App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "uTrib", 0, App.Path & "\cfe.ini"
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "qTrib", 0, App.Path & "\cfe.ini"
        ''         EscreveINI "IS" & psCompString(1, VarContPrd, 3), "vIS", 0, App.Path & "\cfe.ini"
                
        '        EscreveINI "IS" & psCompString(1, VarContPrd, 3), "ValorBase", Numero2(BasePisCofins), App.Path & "\cfe.ini"
                
                EscreveINI "IBSCBS" & psCompString(1, VarContPrd, 3), "CST", Right(VarSitTrib, 3), App.Path & "\cfe.ini"
                If VarNota!fator = "" Or VarNota!fator = 0 Then
                    EscreveINI "IBSCBS" & psCompString(1, VarContPrd, 3), "cClassTrib", "000001", App.Path & "\cfe.ini"
                Else
                    EscreveINI "IBSCBS" & psCompString(1, VarContPrd, 3), "cClassTrib", VarNota!fator, App.Path & "\cfe.ini"
                End If
                
                EscreveINI "gIBSCBS" & psCompString(1, VarContPrd, 3), "vBC", Numero2(VarPrecoBase), App.Path & "\cfe.ini"
                EscreveINI "gIBSCBS" & psCompString(1, VarContPrd, 3), "vIBS", Numero2(VarPrecoBase * (0.1 / 100)), App.Path & "\cfe.ini"
                
                EscreveINI "gIBSUF" & psCompString(1, VarContPrd, 3), "pIBSUF", "0.1000", App.Path & "\cfe.ini"
                EscreveINI "gIBSUF" & psCompString(1, VarContPrd, 3), "vIBSUF", Numero2(Round(VarPrecoBase * (0.1 / 100), 2)), App.Path & "\cfe.ini"
                EscreveINI "gIBSUF" & psCompString(1, VarContPrd, 3), "pRedAliq", "0.00", App.Path & "\cfe.ini"
                EscreveINI "gIBSUF" & psCompString(1, VarContPrd, 3), "pAliqEfet", "0.00", App.Path & "\cfe.ini"
                
                VarTotIbs = VarTotIbs + Round(VarPrecoBase * (0.1 / 100), 2)
                
                EscreveINI "gIBSMun" & psCompString(1, VarContPrd, 3), "pIBSMun", "0.00", App.Path & "\cfe.ini"
                EscreveINI "gIBSMun" & psCompString(1, VarContPrd, 3), "pRedAliq", "0.00", App.Path & "\cfe.ini"
                EscreveINI "gIBSMun" & psCompString(1, VarContPrd, 3), "pAliqEfet", "0.00", App.Path & "\cfe.ini"
                
                EscreveINI "gCBS" & psCompString(1, VarContPrd, 3), "pCBS", "0.9000", App.Path & "\cfe.ini"
                EscreveINI "gCBS" & psCompString(1, VarContPrd, 3), "vCBS", Numero2(Round(VarPrecoBase * (0.9 / 100), 2)), App.Path & "\cfe.ini"
                EscreveINI "gCBS" & psCompString(1, VarContPrd, 3), "pRedAliq", "0.00", App.Path & "\cfe.ini"
                EscreveINI "gCBS" & psCompString(1, VarContPrd, 3), "pAliqEfet", "0.00", App.Path & "\cfe.ini"
                
                VarTotCbs = VarTotCbs + Round(VarPrecoBase * (0.9 / 100), 2)
                
          End If
        End If
        
   
    End If
    If VarPais = "BRASIL" Then
        If TabSql("ENT.Uf") <> "SP" Then
             If pedidos!Classificacao = 1 Then
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pICMSInter", Numero2(Str(VarIcms)), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pICMSInterPart", Numero2("100,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vBCUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pFCPUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vFCPUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vICMSUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vICMSUFRemet", Numero2("0,00"), App.Path & "\cfe.ini"
                
             
             ElseIf VarInscri = "ISENTO" Or VarInscri = "" Then
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pICMSInter", Numero2(Str(VarIcms)), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pICMSInter", Numero2(Str(VarIcms)), App.Path & "\cfe.ini"
                                
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pICMSInterPart", Numero2("100,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vBCUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "pFCPUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vFCPUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vICMSUFDest", Numero2("0,00"), App.Path & "\cfe.ini"
                EscreveINI "ICMSUFDEST" & psCompString(1, VarContPrd, 3), "vICMSUFRemet", Numero2("0,00"), App.Path & "\cfe.ini"
                             
             End If
        
        End If
    End If
            If ChkIpi.value = 1 Then
                If VarNota!TotIpi <> 0 Then
                    EscreveINI "IPI" & psCompString(1, VarContPrd, 3), "CST", "00", App.Path & "\cfe.ini"
                    EscreveINI "IPI" & psCompString(1, VarContPrd, 3), "vBC", Numero2(VarPrecoBase), App.Path & "\cfe.ini"
                    EscreveINI "IPI" & psCompString(1, VarContPrd, 3), "pIPI", Numero2(VarNota.Fields("Prd.Ipi")), App.Path & "\cfe.ini"
                    EscreveINI "IPI" & psCompString(1, VarContPrd, 3), "vIPI", Numero2(VarPrecoBase * (VarNota.Fields("Prd.Ipi") / 100)), App.Path & "\cfe.ini"
                End If
            End If
    If VarTotIpi > 0 Then
    End If
    If VarPis <> "" Then
        If VarPis = "08" Or VarPis = "06" Or VarPis = "07" Or VarPis = "99" Or VarPis = "49" Then
            EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", Left(VarPis, 2), App.Path & "\cfe.ini"
        ElseIf VarPis = "01" Or VarPis = "02" Then
           EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", Left(VarPis, 2), App.Path & "\cfe.ini"
           EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "ValorBase", Numero2(BasePisCofins), App.Path & "\cfe.ini"
           EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "Aliquota", Numero2(VarAliqPis), App.Path & "\cfe.ini"
           EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "Valor", Numero2((VarAliqPis * BasePisCofins) / 100), App.Path & "\cfe.ini"
            VarTotPis = VarTotPis + ((VarAliqPis * BasePisCofins) / 100)
        End If
    End If
    If VarCofins <> "" Then
        If VarCofins = "08" Or VarCofins = "06" Or VarCofins = "07" Or VarCofins = "99" Or VarCofins = "49" Then
            EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", Left(VarPis, 2), App.Path & "\cfe.ini"
            
        ElseIf VarCofins = "01" Or VarCofins = "02" Then
             EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", Left(VarPis, 2), App.Path & "\cfe.ini"
             EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "ValorBase", Numero2(BasePisCofins), App.Path & "\cfe.ini"
             EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "Aliquota", Numero2(VarAliqCofins), App.Path & "\cfe.ini"
             EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "Valor", Numero2((VarAliqCofins * BasePisCofins) / 100), App.Path & "\cfe.ini"
            VarTotCofins = VarTotCofins + ((VarAliqCofins * BasePisCofins) / 100)
        
        End If
    End If
        
   If TxtCfop <> "5.901" Then
        If TxtCfop <> "5901" Then
            VarSitTrib = VarTemp
        End If
   End If
    VarContPrd = VarContPrd + 1
    VarNota.MoveNext
Wend
VarNota.MovePrevious
    If TxtCfop = "5.929" Or TxtCfop = "6.929" Or TxtCfop = "5929" Or TxtCfop = "6929" Then
        
        Dim DatNfe As Recordset

        
        If VarNota!NroNt = 0 Then
            Set DatNfe = Banco.OpenRecordset("SELECT * FROM Nfe WHERE CodNota = " & CodPed & " Order By Codigo ")
            DatNfe.Requery
            If DatNfe.RecordCount > 0 Then
                DatNfe.MoveFirst
                VarContPrd = 1
                While Not DatNfe.Eof
                    EscreveINI "NFRef" & psCompString(1, VarContPrd, 3), "Tipo", "NFE", App.Path & "\cfe.ini"
                    EscreveINI "NFRef" & psCompString(1, VarContPrd, 3), "RefNFE", psCompString(2, DatNfe!NroChave, 6), App.Path & "\cfe.ini"
                    VarContPrd = VarContPrd + 1
                    DatNfe.MoveNext
                Wend
            End If
         Else
                VarContPrd = 1
                EscreveINI "NFRef" & psCompString(1, VarContPrd, 3), "RefNFE", psCompString(2, VarNota!sat, 6), App.Path & "\cfe.ini"
                VarContPrd = VarContPrd + 1
        End If
    
    End If

If BaseIcmsDif <> "" Then
    EscreveINI "TOTAL", "vBC", Numero2(BaseIcmsDif), App.Path & "\cfe.ini"
    EscreveINI "TOTAL", "vICMS", Numero2(ValorIcmDif), App.Path & "\cfe.ini"
Else
    EscreveINI "TOTAL", "vBC", Numero2(VarNota!BaseIcms), App.Path & "\cfe.ini"
    EscreveINI "TOTAL", "vICMS", Numero2(VarNota!ValorIcms), App.Path & "\cfe.ini"
End If

'End If
If VarSitTrib = "050" Then
    EscreveINI "TOTAL", "vICMSDeson", Numero2((IcmsNota * VarPrecoBase) / 100), App.Path & "\cfe.ini"
End If

VarSitTrib = VarTemp

If varTeste = 0 Then
    EscreveINI "TOTAL", "vBCST", Numero2(VarNota!BaseIcmsSub), App.Path & "\cfe.ini"
Else
    EscreveINI "TOTAL", "vBCST", Numero2(0), App.Path & "\cfe.ini"
End If

EscreveINI "TOTAL", "vST", Numero2(VarNota!ValorIcmsSub), App.Path & "\cfe.ini"
If VarFinalidade = 2 Then
   EscreveINI "TOTAL", "vProd", Numero2(0), App.Path & "\cfe.ini"
Else
    EscreveINI "TOTAL", "vProd", Numero2(VarNota!ValorTotalPrd), App.Path & "\cfe.ini"
End If
EscreveINI "TOTAL", "vFrete", Numero2(VarNota.Fields("Nota.ValorFrete")), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vSeg", Numero2(VarNota!ValorSeguro), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vDesc", Numero2(TxtDesconto), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vIpi", Numero2(VarNota!TotIpi), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vPis", Numero2(VarTotPis), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vCofins", Numero2(VarTotCofins), App.Path & "\cfe.ini"
EscreveINI "TOTAL", "vOutro", Numero2(VarNota!TotOutras), App.Path & "\cfe.ini"
If VarTotMonoRet <> 0 Then
    EscreveINI "TOTAL", "vICMSMonoRet", Numero2(VarTotMonoRet), App.Path & "\cfe.ini"
    EscreveINI "TOTAL", "qBCMonoRet", Numero4(VarTotQtdMono), App.Path & "\cfe.ini"
End If


If VarIpiDev <> "" Then
    EscreveINI "TOTAL", "vIPIDevol", Numero2(VarIpiDev), App.Path & "\cfe.ini"
End If

If VarFinalidade = 2 Then
    If txticmssub <> "" Then
        EscreveINI "TOTAL", "vNF", Numero2(txticmssub), App.Path & "\cfe.ini"
    Else
        EscreveINI "TOTAL", "vNF", Numero2(0), App.Path & "\cfe.ini"
    End If
Else
    If txtbasesub <> "" Then
        If VarIpiDev <> "" Then
            EscreveINI "TOTAL", "vNF", Numero2(VarNota!TotNota + txticmssub + VarIpiDev), App.Path & "\cfe.ini"
        Else
            EscreveINI "TOTAL", "vNF", Numero2(VarNota!TotNota + txticmssub), App.Path & "\cfe.ini"
        End If
    Else
        If VarIpiDev <> "" Then
            EscreveINI "TOTAL", "vNF", Numero2(VarNota!TotNota + VarIpiDev), App.Path & "\cfe.ini"
        Else
            EscreveINI "TOTAL", "vNF", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
        End If
    End If
End If


If VarRegimeTrib = 3 And VarTotCbs <> 0 Then

    EscreveINI "ISTot", "vIS", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
    EscreveINI "IBSCBSTot", "vBCIBSCBS", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
    EscreveINI "gIBS", "vIBS", Numero2(VarTotIbs), App.Path & "\cfe.ini"
    EscreveINI "gIBSUFTot", "vIBSUF", Numero2(VarTotIbs), App.Path & "\cfe.ini"
'    EscreveINI "gIBSMunTot", "vIBSMun", "0.00", App.Path & "\cfe.ini"
    
    EscreveINI "gCBSTot", "vCBS", Numero2(VarTotCbs), App.Path & "\cfe.ini"
            
End If

If DBCTransp.Text <> "(Nenhum)" Then
        
    EscreveINI "TRANSPORTADOR", "modFrete", VarNota.Fields("Nota.TipoFrete") - 1, App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "CnpjCpf", psCompString(3, IIf(Len(Trim(LIMPACampo2(VarNota!CnpjTransp))) > 12, Trim(LIMPACampo2(VarNota!CnpjTransp)), ""), 14), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "IE", psCompString(3, VarNota!IETransp, 20), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "xNome", psCompString(3, VarNota!RazaoTransp, 60), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "xEnder", psCompString(3, VarNota!EndTransp, 60), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "xMun", psCompString(3, Trim(VarNota!CidadeTransp), 60), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "UF", psCompString(3, Trim(VarNota!UfTransp), 2), App.Path & "\cfe.ini"
    
    
    EscreveINI "TRANSPORTADOR", "Placa", psCompString(3, Placa.Text, 8), App.Path & "\cfe.ini"
    EscreveINI "TRANSPORTADOR", "UFPlaca", psCompString(3, Uf.Text, 2), App.Path & "\cfe.ini"
Else
    EscreveINI "TRANSPORTADOR", "modFrete", VarNota.Fields("Nota.TipoFrete") - 1, App.Path & "\cfe.ini"
End If

EscreveINI "VOLUME001", "qVol", CInt(Qtd), App.Path & "\cfe.ini"
EscreveINI "VOLUME001", "Esp", psCompString(3, LIMPACampo(Especie), 60), App.Path & "\cfe.ini"
EscreveINI "VOLUME001", "Marca", psCompString(3, Marca, 60), App.Path & "\cfe.ini"
EscreveINI "VOLUME001", "PesoL", Numero3(PesoLiquido), App.Path & "\cfe.ini"
EscreveINI "VOLUME001", "PesoB", Numero3(Bruto), App.Path & "\cfe.ini"

VarPag = "01"

If Oper = 0 Then ' Compra
    Set Pag = Banco.OpenRecordset("Select * From Pag Where Pedido = " & CodPed.Text)
    Pag.Requery
    VarContPrd = 1
    If Pag.RecordCount > 0 Then
        Pag.MoveFirst
        While Not Pag.Eof
           EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "nDup", psCompString(3, Pag!Numero, 60), App.Path & "\cfe.ini"
           EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "dVenc", Format(Pag!DtVenc, "dd/mm/yyyy"), App.Path & "\cfe.ini"
           EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "vDup", Numero2(Pag!Valor), App.Path & "\cfe.ini"
            Pag.MoveNext
            VarContPrd = VarContPrd + 1
        Wend
    End If

Else
    If ChkPcl.value = 1 Then
        Set Dup = Banco.OpenRecordset("SELECT PED.NroNt, Dup.*, Dup.CodDup FROM Dup INNER JOIN PED ON Dup.Pedido = PED.CodPed WHERE Dup.Pedido = " & CodPed.Text & "  ORDER BY Dup.CodDup ")
        Dup.Requery
        VarContPrd = 1
        If Dup.RecordCount > 0 Then
            VarPag = "02"
            EscreveINI "FATURA", "nFat", psCompString(1, Dup.RecordCount, 2), App.Path & "\cfe.ini"
            EscreveINI "FATURA", "vOrig", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
            EscreveINI "FATURA", "vLiq", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
            EscreveINI "FATURA", "vDesc", "0.0001", App.Path & "\cfe.ini"
            
            Dup.MoveFirst
            While Not Dup.Eof
            
                EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "nDup", psCompString(1, Trim(CStr(Dup!NroPcl)), 3), App.Path & "\cfe.ini"
                EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "dVenc", Format(Dup!DtVenc, "dd/mm/yyyy"), App.Path & "\cfe.ini"
                EscreveINI "DUPLICATA" & psCompString(1, VarContPrd, 3), "vDup", Numero2(Dup!Valor), App.Path & "\cfe.ini"
            
                Dup.Edit
                Dup!Documento = MebNfe
                Dup.Update
                Dup.MoveNext
                VarContPrd = VarContPrd + 1
                
            Wend
        End If
    Else
        Set Dup = Banco.OpenRecordset("SELECT PED.NroNt, Dup.*, Dup.CodDup FROM Dup INNER JOIN PED ON Dup.Pedido = PED.CodPed WHERE Dup.Pedido = " & CodPed.Text & "  ORDER BY Dup.CodDup ")
        Dup.Requery
        If Dup.RecordCount > 0 Then
            Dup.MoveFirst
            While Not Dup.Eof
                Dup.Edit
                Dup!Documento = MebNfe
                Dup.Update
                Dup.MoveNext
            Wend
        End If
        
    
    End If
End If

If TxtChave <> "" Or VarNat = 0 Then
    EscreveINI "Pag001", "tPag", psCompString(1, 90, 2), App.Path & "\cfe.ini"
    EscreveINI "Pag001", "vPag", 0, App.Path & "\cfe.ini"
Else
    EscreveINI "Pag001", "tPag", psCompString(1, VarPag, 2), App.Path & "\cfe.ini"
    EscreveINI "Pag001", "vPag", Numero2(VarNota!TotNota), App.Path & "\cfe.ini"
End If
VarContPrd = 1

If ChkDeducao.value = 1 Then
    VarObsDeducao = "Icm diferido sobre " & MebReducao & "% "
Else
    VarObsDeducao = ""
End If
If Trim(CmbObsRoda) <> "" Then
    VarObsFinal = Left(Trim(CmbObsRoda), 500)
End If

If Trim(CmbObsPed) <> "" Then

    VarObsFinal = VarObsFinal & " " & Left(Trim(CmbObsPed), 500)
    VarContPrd = VarContPrd + 1
End If

If VarObsDeducao <> "" Then
    VarObsFinal = VarObsFinal & " " & Left(Trim(VarObsDeducao), 500)
End If

If VarNota!TotNota <> 0 Then
    EscreveINI "DadosAdicionais", "InfCpl", VarObsFinal & " " & Left(Trim(" Trib.R$ " & Format(VarTotImpostoNac, "#,##0.00") & "(" & Format(((VarTotImpostoNac / CCur(VarNota!TotNota)) * 100), "##0.00") & "%) Federal R$ " & Format(VarTotImpostoEst, "#,##0.00") & "(" & Format(((VarTotImpostoEst / CCur(VarNota!TotNota)) * 100), "##0.00") & "%)" & " Estadual. "), 500), App.Path & "\cfe.ini"
    VarContPrd = VarContPrd + 1
Else
    EscreveINI "DadosAdicionais", "InfCpl", VarObsFinal, App.Path & "\cfe.ini"
    VarContPrd = VarContPrd + 1
End If
If Len(Registro10) > 2 Then

End If

Frame.Visible = True
DoEvents

Label22.Visible = False
VarNotaValida = False

If Index = 2 Then
    Frame.Visible = True
    Label22.Visible = True
    Label22.Caption = "Aguarde . . . enviando nota"
        
        Label22.Visible = False
        AbreAcbr "NFE.CriarEnviarNFe(" & App.Path & "\cfe.ini" & ",1,0,1" & ")"
        If TrataAcbr("Nota", MebNfe) = 100 Then
'            VarControle = True
            MsgBox "Nota enviada com sucesso !", vbInformation, App.Title
            VarNotaValida = True
            
        End If
End If

If Index = 3 Then

        AbreAcbr "NFE.CriarNFe(" & App.Path & "\cfe.ini" & ",1" & ")"

        TrataAcbr "Visualiza", MebNfe
        MsgBox "Nota Gerada com sucesso !", vbInformation, App.Title
        AbreAcbr "Nfe.ImprimirDANFEPDF(" & PortaAcbr & "\logs\" & VarChaveNfe & "-nfe.xml" & ")"
        
        While Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") = ""
        Wend
        
        X = 0
        
        If Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") = "" Then
            For X = 1 To 700000
            Next
        Else
            For X = 1 To 100000
            Next
        End If
        
        If Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") <> "" Then
        
            ShellExecute Me.Hwnd, "open", PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf", _
                  vbNullString, vbNullString, vbNormalFocus
        Else
            MsgBox "Arquivo de visualização da nota não gerado. Tente uma segunda via"
        End If
        
End If

VarIpiDev = ""
VarPIpi = ""
If VarNotaValida = False Then
    Confirma(2).Enabled = True
    LimpaCampos
    
    Exit Sub
End If
If VarNroNfe = "" Then
    ConfCfg.Edit
    If IsNull(ConfCfg("NroNota")) Then
        ConfCfg("NroNota") = 2
    Else
        If VarNotaValida = True Then
            pedidos.Edit
            pedidos("NroNt") = CCur(ConfCfg("NroNota"))
            pedidos.Update
            
            ConfCfg("NroNota") = CCur(ConfCfg("NroNota")) + 1
        End If
    End If
    ConfCfg.Update
Else
    If VarNotaValida = True Then
        pedidos.Edit
        pedidos("NroNt") = CCur(MebNfe)
        pedidos.Update
        EscreveINI "Nfe", "NroNfe", CCur(MebNfe) + 1, App.Path & "\config.ini"
    End If
End If

Set VarNota = Banco.OpenRecordset("Select * from nota where codped = " & CodPed)
VarNota.Requery
VarNota.MoveLast
VarNota.MoveFirst

If VarNota.RecordCount > 0 Then
    VarNota.Edit
    VarNota!InformFisco = VarChaveNfe
    VarNota.Update
End If

ImprimiNota VarCnpjEmite, VarEmail
 
Confirma(2).Enabled = True
LimpaCampos

End Sub

Sub FuncaoExterior(Documento As String)

    EscreveINI "Destinatario", "cMun", "9999999", App.Path & "\cfe.ini"
    EscreveINI "Destinatario", "xMun", psCompString(3, "EXTERIOR", 60), App.Path & "\cfe.ini"
    EscreveINI "Destinatario", "idEstrangeiro", Documento, App.Path & "\cfe.ini"
    If VarPais = "BOLIVIA" Then
        EscreveINI "Destinatario", "cPais", "973", App.Path & "\cfe.ini"
    End If
    If VarPais = "CHILE" Then
        EscreveINI "Destinatario", "cPais", "1589", App.Path & "\cfe.ini"
    End If
    
    If VarPais = "ESTADOS UNIDOS" Then
        EscreveINI "Destinatario", "cPais", "2496", App.Path & "\cfe.ini"
    End If

End Sub

Public Sub LimpaIni()
If Dir(App.Path & "\Cfe.Ini") <> "" Then
   Kill App.Path & "\Cfe.ini"
End If

If Dir(PortaAcbr & "\Sai.txt") <> "" Then
    Kill PortaAcbr & "\Sai.txt"
End If

End Sub

Public Sub ImprimiNota(VarCnpjEmite As String, VarEmail As String)

AbreAcbr "Nfe.ImprimirDANFEPDF(" & PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfe\" & Trim(Str(Year(Emissao))) & psCompString(1, Month(Emissao), 2) & "\Nfe\" & VarChaveNfe & "-nfe.xml" & ")"

While Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") = ""

Wend

X = 0

If Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") = "" Then
    For X = 1 To 700000
    Next
Else
    For X = 1 To 100000
    
    Next
End If

If Dir(PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf") <> "" Then

    ShellExecute Me.Hwnd, "open", PortaAcbr & "\PDF\" & VarChaveNfe & "-nfe.pdf", _
          vbNullString, vbNullString, vbNormalFocus
Else
    MsgBox "Arquivo de visualização da nota não gerado. Tente uma segunda via"
End If

If VarEmail <> "" Then
    If MsgBox("Deseja enviar e-mail do xml agora ?", vbYesNo, App.Title) = vbYes Then
         AbreAcbr "NFe.EnviarEmail(" & VarEmail & "," & PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfe\" & Trim(Str(Year(Emissao))) & psCompString(1, Month(Emissao), 2) & "\Nfe\" & VarChaveNfe & "-nfe.xml" & ",1," & "Envio de NFe - " & NomeEmpresa & ")"
    End If
End If


End Sub



Public Sub Teste()

End Sub




