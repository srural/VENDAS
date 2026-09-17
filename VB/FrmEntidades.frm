VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FrmEntidades 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14085
   Icon            =   "FrmEntidades.frx":0000
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
      TabIndex        =   194
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
         TabIndex        =   196
         Top             =   45
         Width           =   8460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção, não apagar os registros com o nome (Nenhum)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8865
         TabIndex        =   195
         Top             =   45
         Width           =   5055
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
      TabIndex        =   156
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
         TabIndex        =   158
         Top             =   90
         Width           =   6270
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Entidades"
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
         TabIndex        =   157
         Top             =   45
         Width           =   5190
      End
   End
   Begin TabDlg.SSTab tabCad 
      Height          =   7545
      Left            =   0
      TabIndex        =   68
      Top             =   570
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   13309
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmEntidades.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ChkAtivos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdPesq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DatCadastro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DatDados(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DBGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cadastro"
      TabPicture(1)   =   "FrmEntidades.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PctFundo(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CmbFina"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CmbCad"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbPromi"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Gerenciais"
      TabPicture(2)   =   "FrmEntidades.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "PctFundo(5)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Contato"
      TabPicture(3)   =   "FrmEntidades.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GridCto"
      Tab(3).Control(1)=   "DatDados(5)"
      Tab(3).Control(2)=   "BtIncluir"
      Tab(3).Control(3)=   "BtExcluir"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Obs"
      TabPicture(4)   =   "FrmEntidades.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TxtDados(40)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Movimento"
      TabPicture(5)   =   "FrmEntidades.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label17"
      Tab(5).Control(1)=   "Label18"
      Tab(5).Control(2)=   "DtData2"
      Tab(5).Control(3)=   "DtData1"
      Tab(5).Control(4)=   "DbgFina"
      Tab(5).Control(5)=   "GridEstat"
      Tab(5).Control(6)=   "DatDados(6)"
      Tab(5).Control(7)=   "ChkPago"
      Tab(5).Control(8)=   "Frame8"
      Tab(5).Control(9)=   "DatDados(7)"
      Tab(5).Control(10)=   "CmdAtu"
      Tab(5).Control(11)=   "Frame9"
      Tab(5).ControlCount=   12
      Begin VB.CheckBox ChkAtivos 
         Caption         =   "Ativos"
         Height          =   285
         Left            =   -74640
         TabIndex        =   210
         Top             =   7020
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CommandButton cmbPromi 
         Caption         =   "Promissoria"
         Height          =   720
         Left            =   -62130
         Picture         =   "FrmEntidades.frx":00B4
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   3690
         Width           =   945
      End
      Begin VB.CommandButton CmdPesq 
         Caption         =   "Pesquisa Externa"
         Height          =   330
         Left            =   -62805
         TabIndex        =   193
         Top             =   7020
         Width           =   1725
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   -65415
         TabIndex        =   174
         Top             =   3870
         Width           =   4290
         Begin VB.CommandButton CmbImprimir 
            Caption         =   "Imprimir"
            Height          =   285
            Left            =   1440
            TabIndex        =   207
            Top             =   3195
            Width           =   1545
         End
         Begin MSMask.MaskEdBox MebPrimeiraCompra 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   180
            Top             =   180
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebUltimaCompra 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   181
            Top             =   510
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebMediaCompra 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   182
            Top             =   1830
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebUltimoPgto 
            Height          =   330
            Left            =   2385
            TabIndex        =   183
            Top             =   2160
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebMediaPgto 
            Height          =   330
            Left            =   2385
            TabIndex        =   184
            Top             =   2490
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDiasAtraso 
            Height          =   330
            Left            =   2385
            TabIndex        =   185
            Top             =   2820
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebVista 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   187
            Top             =   840
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebPrazo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   188
            Top             =   1170
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebOrcamento 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   191
            Top             =   1500
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Orcamentos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   192
            Top             =   1574
            Width           =   1020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Compras vista"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   190
            Top             =   922
            Width           =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Compras a prazo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   189
            Top             =   1248
            Width           =   1425
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Dias Atraso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   186
            Top             =   2880
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Media atraso Pagamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   179
            Top             =   2552
            Width           =   2070
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Pagamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   178
            Top             =   2226
            Width           =   1545
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Media Compra Mes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   177
            Top             =   1900
            Width           =   1605
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ultima Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   176
            Top             =   596
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Primeira Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   175
            Top             =   270
            Width           =   1425
         End
      End
      Begin VB.CommandButton CmdAtu 
         Caption         =   "Atualiza"
         Height          =   330
         Left            =   -74820
         TabIndex        =   173
         Top             =   6975
         Width           =   1320
      End
      Begin VB.Data DatDados 
         Caption         =   "Dup"
         Connect         =   "Access"
         DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   7
         Left            =   -71130
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT * FROM Dup"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Frame Frame8 
         Height          =   780
         Left            =   -69915
         TabIndex        =   165
         Top             =   6525
         Width           =   4425
         Begin MSMask.MaskEdBox MebSubTot 
            DataField       =   "Custo"
            Height          =   360
            Left            =   90
            TabIndex        =   166
            Tag             =   "S"
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebPagos 
            DataField       =   "Custo"
            Height          =   360
            Left            =   1485
            TabIndex        =   167
            Tag             =   "S"
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;- #,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebAPagar 
            DataField       =   "Custo"
            Height          =   330
            Left            =   2955
            TabIndex        =   168
            Tag             =   "S"
            Top             =   405
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   582
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Index           =   45
            Left            =   90
            TabIndex        =   171
            Top             =   180
            Width           =   360
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagas"
            Height          =   195
            Index           =   44
            Left            =   1470
            TabIndex        =   170
            Top             =   180
            Width           =   450
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Abertas"
            Height          =   195
            Index           =   43
            Left            =   2865
            TabIndex        =   169
            Top             =   180
            Width           =   540
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   1395
            X2              =   1395
            Y1              =   90
            Y2              =   765
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   2835
            X2              =   2835
            Y1              =   90
            Y2              =   765
         End
      End
      Begin VB.CheckBox ChkPago 
         Caption         =   "Mostrar os Pagos"
         Height          =   285
         Left            =   -74820
         TabIndex        =   164
         Top             =   6525
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
      Begin VB.CommandButton CmbCad 
         Caption         =   "Ficha"
         Height          =   735
         Left            =   -64110
         Picture         =   "FrmEntidades.frx":04F6
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Ficha Cadastral"
         Top             =   3690
         Width           =   870
      End
      Begin VB.CommandButton CmbFina 
         Caption         =   "Financeiro"
         Height          =   735
         Left            =   -63120
         Picture         =   "FrmEntidades.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Financeiro"
         Top             =   3690
         Width           =   870
      End
      Begin VB.Data DatDados 
         Caption         =   "estatistica"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   6
         Left            =   -69195
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2655
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         DataField       =   "Obs"
         Height          =   6645
         Index           =   40
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   153
         Tag             =   "S"
         Top             =   540
         Width           =   13710
      End
      Begin VB.CommandButton BtExcluir 
         Caption         =   "&Exluir"
         Height          =   285
         Left            =   -73650
         TabIndex        =   151
         Top             =   6930
         Width           =   1050
      End
      Begin VB.CommandButton BtIncluir 
         Caption         =   "&Incluir"
         Height          =   285
         Left            =   -74775
         TabIndex        =   150
         Top             =   6930
         Width           =   1050
      End
      Begin VB.Data DatDados 
         Caption         =   "Contato"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   5
         Left            =   -72030
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6930
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   7155
         Index           =   0
         Left            =   -74775
         ScaleHeight     =   7155
         ScaleWidth      =   13725
         TabIndex        =   91
         Top             =   360
         Width           =   13725
         Begin VB.Data DatPesquisaG 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   11610
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT tabcidades.codcidade, tabcidades.cep, tabcidades.cidade, tabcidades.uf FROM tabcidades ORDER BY tabcidades.cidade"
            Top             =   4590
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Frame FramePesquisaG 
            Height          =   3795
            Left            =   720
            TabIndex        =   208
            Top             =   2925
            Visible         =   0   'False
            Width           =   5460
            Begin MSDBCtls.DBList DblPesquisaG 
               Bindings        =   "FrmEntidades.frx":0D7A
               DataField       =   "Cidade"
               DataSource      =   "DatPesquisaG"
               Height          =   3375
               Left            =   135
               TabIndex        =   209
               Top             =   270
               Width           =   5190
               _ExtentX        =   9155
               _ExtentY        =   5953
               _Version        =   393216
               Appearance      =   0
               ListField       =   "Cidade"
               BoundColumn     =   "Cidade"
            End
         End
         Begin VB.Data DatCidade 
            Caption         =   "Cid"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   11745
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT tabcidades.codcidade, tabcidades.cep, tabcidades.cidade, tabcidades.uf FROM tabcidades ORDER BY tabcidades.cidade"
            Top             =   5040
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Frame FrmCidade 
            Caption         =   "Cidade"
            Height          =   3795
            Left            =   6705
            TabIndex        =   162
            Top             =   2880
            Visible         =   0   'False
            Width           =   5460
            Begin MSDBCtls.DBList DBLCidade 
               Bindings        =   "FrmEntidades.frx":0D95
               DataField       =   "cidade"
               DataSource      =   "DatCidade"
               Height          =   3375
               Left            =   135
               TabIndex        =   163
               Top             =   270
               Width           =   5190
               _ExtentX        =   9155
               _ExtentY        =   5953
               _Version        =   393216
               Appearance      =   0
               ListField       =   "Cidade"
               BoundColumn     =   "codcidade"
            End
         End
         Begin VB.CommandButton BtCidades 
            Caption         =   "Cidade"
            Height          =   300
            Left            =   4275
            TabIndex        =   161
            Top             =   2220
            Width           =   870
         End
         Begin VB.CommandButton BtIbge 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   9270
            TabIndex        =   160
            Top             =   2850
            Width           =   465
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Nome"
            Height          =   315
            Index           =   1
            Left            =   60
            MaxLength       =   60
            TabIndex        =   2
            Tag             =   "S"
            Top             =   990
            Width           =   6435
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Endereco"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   2
            Left            =   60
            MaxLength       =   40
            TabIndex        =   4
            Tag             =   "S"
            Top             =   1620
            Width           =   6015
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Nro"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   3
            Left            =   6225
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "S"
            Top             =   1620
            Width           =   1095
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Bairro"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   4
            Left            =   9315
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "S"
            Top             =   1590
            Width           =   3720
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Cidade"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   5
            Left            =   45
            MaxLength       =   40
            TabIndex        =   8
            Tag             =   "S"
            Top             =   2205
            Width           =   4095
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Uf"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   6
            Left            =   5265
            MaxLength       =   2
            TabIndex        =   9
            Tag             =   "S"
            Top             =   2220
            Width           =   585
         End
         Begin VB.CheckBox ChkDados 
            Caption         =   "Ativo"
            DataField       =   "Ativo"
            Height          =   255
            Index           =   0
            Left            =   9420
            TabIndex        =   110
            Tag             =   "S"
            Top             =   405
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox CmbDados 
            Appearance      =   0  'Flat
            DataField       =   "Tipo"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmEntidades.frx":0DAD
            Left            =   1680
            List            =   "FrmEntidades.frx":0DC9
            TabIndex        =   0
            Tag             =   "SI"
            Top             =   405
            Width           =   2310
         End
         Begin VB.ComboBox CmbDados 
            Appearance      =   0  'Flat
            DataField       =   "Classificacao"
            Height          =   315
            Index           =   4
            ItemData        =   "FrmEntidades.frx":0E58
            Left            =   4245
            List            =   "FrmEntidades.frx":0E62
            TabIndex        =   1
            Tag             =   "SI"
            Top             =   420
            Width           =   2565
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Pais"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   36
            Left            =   6030
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "S"
            Top             =   2220
            Width           =   3210
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
            TabIndex        =   109
            Tag             =   "N"
            Top             =   390
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "Documentos pessoa física"
            Height          =   825
            Left            =   45
            TabIndex        =   106
            Top             =   3285
            Width           =   4110
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "RG"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   10
               Left            =   150
               TabIndex        =   19
               Tag             =   "S"
               Top             =   360
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
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "CPF"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   11
               Left            =   2025
               TabIndex        =   21
               Tag             =   "S"
               Top             =   360
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CPF:"
               Height          =   195
               Index           =   37
               Left            =   2040
               TabIndex        =   108
               Top             =   180
               Width           =   345
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "RG:"
               Height          =   195
               Index           =   29
               Left            =   180
               TabIndex        =   107
               Top             =   180
               Width           =   285
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Documentos pessoa Juridica"
            Height          =   825
            Left            =   4275
            TabIndex        =   103
            Top             =   3285
            Width           =   4545
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "CGC"
               Height          =   315
               Index           =   12
               Left            =   165
               TabIndex        =   22
               Tag             =   "S"
               Top             =   405
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   18
               Mask            =   "##.###.###/####-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "InscrEst"
               Height          =   315
               Index           =   13
               Left            =   2430
               TabIndex        =   23
               Tag             =   "S"
               Top             =   405
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CNPJ"
               Height          =   195
               Index           =   39
               Left            =   180
               TabIndex        =   105
               Top             =   180
               Width           =   405
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Inscr.Estadual"
               Height          =   195
               Index           =   38
               Left            =   2430
               TabIndex        =   104
               Top             =   180
               Width           =   1005
            End
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Fantasia"
            ForeColor       =   &H80000007&
            Height          =   315
            Index           =   26
            Left            =   6615
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "S"
            Top             =   990
            Width           =   6435
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Complemento"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   34
            Left            =   7425
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "S"
            Top             =   1620
            Width           =   1785
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
            Left            =   9135
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3465
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
            Left            =   9135
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3825
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Frame Frame6 
            Caption         =   "Dados Complementares (Pessoa Física)"
            Height          =   2895
            Left            =   45
            TabIndex        =   92
            Top             =   4140
            Width           =   13650
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Email"
               Height          =   315
               Index           =   33
               Left            =   2100
               MaxLength       =   60
               TabIndex        =   34
               Tag             =   "S"
               Top             =   2160
               Width           =   5535
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "HomePage"
               Height          =   315
               Index           =   32
               Left            =   7875
               MaxLength       =   60
               TabIndex        =   35
               Tag             =   "S"
               Top             =   2160
               Width           =   5565
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Uf_Trabalho"
               Height          =   315
               Index           =   20
               Left            =   12765
               MaxLength       =   2
               TabIndex        =   32
               Tag             =   "S"
               Top             =   1575
               Width           =   675
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Endereco_Trabalho"
               Height          =   315
               Index           =   21
               Left            =   4410
               MaxLength       =   40
               TabIndex        =   30
               Tag             =   "S"
               Top             =   1575
               Width           =   4230
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Cidade_Trabalho"
               Height          =   315
               Index           =   18
               Left            =   8775
               MaxLength       =   30
               TabIndex        =   31
               Tag             =   "S"
               Top             =   1575
               Width           =   3825
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Trabalho"
               Height          =   315
               Index           =   17
               Left            =   180
               MaxLength       =   40
               TabIndex        =   29
               Tag             =   "S"
               Top             =   1575
               Width           =   4065
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Profissao"
               Height          =   315
               Index           =   16
               Left            =   7110
               MaxLength       =   30
               TabIndex        =   28
               Tag             =   "S"
               Top             =   1035
               Width           =   4065
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Filiacao"
               Height          =   315
               Index           =   15
               Left            =   180
               MaxLength       =   60
               TabIndex        =   27
               Tag             =   "S"
               Top             =   1035
               Width           =   6615
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Conjuge"
               Height          =   315
               Index           =   14
               Left            =   2055
               MaxLength       =   40
               TabIndex        =   25
               Tag             =   "S"
               Top             =   465
               Width           =   4950
            End
            Begin VB.ComboBox CmbDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "EstadoCivil"
               Height          =   315
               Index           =   1
               ItemData        =   "FrmEntidades.frx":0E78
               Left            =   180
               List            =   "FrmEntidades.frx":0E8B
               TabIndex        =   24
               Tag             =   "S"
               Top             =   450
               Width           =   1755
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "DtCasamento"
               Height          =   315
               Index           =   9
               Left            =   7110
               TabIndex        =   26
               Tag             =   "S"
               Top             =   450
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777152
               MaxLength       =   10
               Format          =   "dd-mmm-yy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "TimeTorce"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   17
               Left            =   8775
               TabIndex        =   198
               Tag             =   "S"
               Top             =   450
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777152
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Fone_Trabalho"
               Height          =   315
               Index           =   18
               Left            =   180
               TabIndex        =   33
               Top             =   2160
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777152
               PromptInclude   =   0   'False
               MaxLength       =   14
               Mask            =   "(##)C###-####C"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fone Trabalho"
               Height          =   195
               Index           =   46
               Left            =   165
               TabIndex        =   211
               Top             =   1935
               Width           =   1035
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "E-mail:"
               Height          =   195
               Index           =   55
               Left            =   2115
               TabIndex        =   202
               Top             =   1935
               Width           =   465
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Home Page"
               Height          =   195
               Index           =   54
               Left            =   7875
               TabIndex        =   201
               Top             =   1935
               Width           =   840
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Data de Casamento"
               Height          =   195
               Index           =   36
               Left            =   7110
               TabIndex        =   102
               Top             =   210
               Width           =   1410
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cpf Conjuge"
               Height          =   195
               Index           =   35
               Left            =   8760
               TabIndex        =   101
               Top             =   225
               Width           =   870
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Index           =   30
               Left            =   12765
               TabIndex        =   100
               Top             =   1350
               Width           =   210
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "Endereço do Trabalho"
               Height          =   195
               Index           =   31
               Left            =   4410
               TabIndex        =   99
               Top             =   1350
               Width           =   1590
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade Onde Trabalha"
               Height          =   195
               Index           =   28
               Left            =   8775
               TabIndex        =   98
               Top             =   1350
               Width           =   1605
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Empresa onde Trababalha"
               Height          =   195
               Index           =   27
               Left            =   180
               TabIndex        =   97
               Top             =   1350
               Width           =   1875
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Profissão"
               Height          =   195
               Index           =   26
               Left            =   7110
               TabIndex        =   96
               Top             =   810
               Width           =   645
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Filiação:"
               Height          =   195
               Index           =   25
               Left            =   180
               TabIndex        =   95
               Top             =   810
               Width           =   585
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conjuge:"
               Height          =   195
               Index           =   24
               Left            =   2055
               TabIndex        =   94
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label2 
               Caption         =   "Estado Civil:"
               Height          =   255
               Left            =   180
               TabIndex        =   93
               Top             =   225
               Width           =   1215
            End
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cep"
            Height          =   315
            Index           =   1
            Left            =   9345
            TabIndex        =   11
            Top             =   2220
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##-###-###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone"
            Height          =   315
            Index           =   2
            Left            =   1815
            TabIndex        =   14
            Top             =   2850
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)CC###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fax"
            Height          =   315
            Index           =   3
            Left            =   3585
            TabIndex        =   15
            Top             =   2850
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)C###-####C"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Celular"
            Height          =   315
            Index           =   4
            Left            =   45
            TabIndex        =   13
            Top             =   2850
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)C###-####C"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "DtNasc"
            Height          =   315
            Index           =   23
            Left            =   10860
            TabIndex        =   12
            Tag             =   "S"
            Top             =   2220
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "DtCadastro"
            Height          =   315
            Index           =   14
            Left            =   10080
            TabIndex        =   18
            Tag             =   "S"
            Top             =   2850
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "InscrProd"
            Height          =   315
            Index           =   15
            Left            =   5370
            TabIndex        =   16
            Tag             =   "S"
            Top             =   2850
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Matricula"
            Height          =   315
            Index           =   16
            Left            =   7575
            TabIndex        =   17
            Tag             =   "S"
            Top             =   2850
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dados básicos nota fiscal"
            Height          =   195
            Left            =   11160
            TabIndex        =   155
            Top             =   450
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   10560
            Top             =   450
            Width           =   375
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   130
            Top             =   765
            Width           =   465
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   129
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Numero:"
            Height          =   195
            Index           =   3
            Left            =   6225
            TabIndex        =   128
            Top             =   1410
            Width           =   600
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   6
            Left            =   9315
            TabIndex        =   127
            Top             =   1395
            Width           =   450
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   126
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Index           =   4
            Left            =   5295
            TabIndex        =   125
            Top             =   1980
            Width           =   255
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Index           =   5
            Left            =   9360
            TabIndex        =   124
            Top             =   1980
            Width           =   315
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fone:"
            Height          =   195
            Index           =   8
            Left            =   1845
            TabIndex        =   123
            Top             =   2610
            Width           =   405
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Index           =   10
            Left            =   3585
            TabIndex        =   122
            Top             =   2610
            Width           =   300
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Index           =   11
            Left            =   45
            TabIndex        =   121
            Top             =   2610
            Width           =   525
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   61
            Left            =   1680
            TabIndex        =   120
            Top             =   180
            Width           =   360
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Classificação"
            Height          =   195
            Index           =   62
            Left            =   4245
            TabIndex        =   119
            Top             =   225
            Width           =   930
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Pais"
            Height          =   195
            Index           =   64
            Left            =   6030
            TabIndex        =   118
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label3 
            Caption         =   "Data Nascimento"
            Height          =   195
            Left            =   10890
            TabIndex        =   117
            Tag             =   "S"
            Top             =   1980
            Width           =   1320
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   116
            Top             =   180
            Width           =   540
         End
         Begin VB.Label Label5 
            Caption         =   "Data &Cadastro"
            Height          =   195
            Left            =   10080
            TabIndex        =   115
            Tag             =   "S"
            Top             =   2610
            Width           =   1320
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Fantasia/Apelido"
            Height          =   195
            Index           =   34
            Left            =   6600
            TabIndex        =   114
            Top             =   765
            Width           =   1200
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscr.Produtor"
            Height          =   225
            Index           =   40
            Left            =   5355
            TabIndex        =   113
            Top             =   2610
            Width           =   990
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cx.Postal"
            Height          =   195
            Index           =   41
            Left            =   7440
            TabIndex        =   112
            Top             =   1410
            Width           =   660
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Codigo I.B.G.E"
            Height          =   195
            Index           =   42
            Left            =   7590
            TabIndex        =   111
            Top             =   2610
            Width           =   1050
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Titulares"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -68745
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4545
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Frame Frame7 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   -74865
         TabIndex        =   82
         Top             =   450
         Width           =   13740
         Begin VB.CheckBox ChkPesq 
            Caption         =   "Qualquer parte do campo"
            Height          =   285
            Left            =   6840
            TabIndex        =   159
            Top             =   270
            Width           =   2220
         End
         Begin VB.TextBox TxtPesquisa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            TabIndex        =   87
            Top             =   270
            Width           =   6525
         End
         Begin VB.OptionButton OptPesq 
            Caption         =   "&Cliente"
            Height          =   285
            Index           =   0
            Left            =   10305
            TabIndex        =   86
            Top             =   315
            Width           =   825
         End
         Begin VB.OptionButton OptPesq 
            Caption         =   "&Fornecedor"
            Height          =   285
            Index           =   1
            Left            =   11250
            TabIndex        =   85
            Top             =   315
            Width           =   1230
         End
         Begin VB.OptionButton OptPesq 
            Caption         =   "&Vendedor"
            Height          =   285
            Index           =   2
            Left            =   12465
            TabIndex        =   84
            Top             =   315
            Width           =   1005
         End
         Begin VB.OptionButton OptPesq 
            Caption         =   "&Todos"
            Height          =   285
            Index           =   3
            Left            =   9360
            TabIndex        =   83
            Top             =   315
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   7035
         Index           =   5
         Left            =   225
         ScaleHeight     =   7035
         ScaleWidth      =   13635
         TabIndex        =   69
         Top             =   405
         Width           =   13635
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PesAut"
            Height          =   315
            Index           =   25
            Left            =   225
            MaxLength       =   50
            TabIndex        =   45
            Tag             =   "S"
            Top             =   2430
            Width           =   9045
         End
         Begin VB.Frame Frame1 
            Caption         =   "Endereço de Cobrança"
            Height          =   1530
            Left            =   135
            TabIndex        =   142
            Top             =   5490
            Width           =   12165
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Endereço_Cob"
               Height          =   315
               Index           =   31
               Left            =   240
               MaxLength       =   50
               TabIndex        =   61
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
               TabIndex        =   62
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
               TabIndex        =   63
               Tag             =   "S"
               Top             =   495
               Width           =   4200
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Cidade_Cob"
               Height          =   315
               Index           =   28
               Left            =   225
               MaxLength       =   20
               TabIndex        =   64
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
               TabIndex        =   65
               Tag             =   "S"
               Top             =   1080
               Width           =   765
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Pais_Cob"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   38
               Left            =   7650
               MaxLength       =   30
               TabIndex        =   67
               Tag             =   "S"
               Top             =   1035
               Width           =   4215
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Cob"
               Height          =   315
               Index           =   5
               Left            =   6330
               TabIndex        =   66
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
               Caption         =   "Endereço:"
               Height          =   195
               Index           =   50
               Left            =   240
               TabIndex        =   149
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
               TabIndex        =   148
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
               TabIndex        =   147
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
               TabIndex        =   146
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
               TabIndex        =   145
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
               TabIndex        =   144
               Top             =   840
               Width           =   315
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "Pais"
               Height          =   195
               Index           =   66
               Left            =   7650
               TabIndex        =   143
               Top             =   870
               Width           =   300
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Endereço de Entrega"
            Height          =   1530
            Left            =   135
            TabIndex        =   134
            Top             =   3960
            Width           =   12165
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Pais_Entrega"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   37
               Left            =   7575
               MaxLength       =   30
               TabIndex        =   60
               Tag             =   "S"
               Top             =   1080
               Width           =   4275
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Uf_Entrega"
               Height          =   315
               Index           =   9
               Left            =   5280
               MaxLength       =   2
               TabIndex        =   58
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
               TabIndex        =   57
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
               TabIndex        =   56
               Tag             =   "S"
               Top             =   495
               Width           =   4245
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               DataField       =   "Nro_Entrega"
               Height          =   315
               Index           =   12
               Left            =   6270
               MaxLength       =   40
               TabIndex        =   55
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
               TabIndex        =   54
               Tag             =   "S"
               Top             =   480
               Width           =   5895
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Entrega"
               Height          =   315
               Index           =   8
               Left            =   6300
               TabIndex        =   59
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
               TabIndex        =   141
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
               TabIndex        =   140
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
               TabIndex        =   139
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
               TabIndex        =   138
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
               TabIndex        =   137
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
               TabIndex        =   136
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LblDados 
               AutoSize        =   -1  'True
               Caption         =   "Pais"
               Height          =   195
               Index           =   65
               Left            =   7575
               TabIndex        =   135
               Top             =   840
               Width           =   300
            End
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "DiaBomPgto"
            Height          =   315
            Index           =   27
            Left            =   4470
            MaxLength       =   40
            TabIndex        =   38
            Tag             =   "S"
            Top             =   285
            Width           =   1050
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "NomeVendedor"
            Height          =   315
            Index           =   19
            Left            =   4410
            MaxLength       =   40
            TabIndex        =   47
            Tag             =   "S"
            Top             =   3015
            Width           =   3525
         End
         Begin VB.Frame Frame5 
            Caption         =   "Referências (Pessoas,Comercias e/ou Bancárias)"
            Height          =   1545
            Left            =   225
            TabIndex        =   70
            Top             =   630
            Width           =   10320
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Refer1"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   22
               Left            =   180
               MaxLength       =   40
               TabIndex        =   42
               Tag             =   "S"
               Top             =   315
               Width           =   9975
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Refer2"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   23
               Left            =   180
               MaxLength       =   40
               TabIndex        =   43
               Tag             =   "S"
               Top             =   675
               Width           =   9975
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "Refer3"
               DataSource      =   "DatDados(1)"
               Height          =   315
               Index           =   24
               Left            =   180
               MaxLength       =   40
               TabIndex        =   44
               Tag             =   "S"
               Top             =   1035
               Width           =   9975
            End
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Comissao"
            Height          =   315
            Index           =   35
            Left            =   12060
            MaxLength       =   40
            TabIndex        =   49
            Tag             =   "S"
            Top             =   3015
            Width           =   1470
         End
         Begin VB.ComboBox CmbDados 
            Appearance      =   0  'Flat
            DataField       =   "Status"
            Height          =   315
            Index           =   3
            ItemData        =   "FrmEntidades.frx":0EBC
            Left            =   225
            List            =   "FrmEntidades.frx":0ED5
            TabIndex        =   36
            Tag             =   "SI"
            Top             =   270
            Width           =   2550
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Prazo"
            Height          =   315
            Index           =   7
            Left            =   5640
            MaxLength       =   40
            TabIndex        =   39
            Tag             =   "S"
            Top             =   270
            Width           =   1680
         End
         Begin VB.Data DatDados 
            Caption         =   "Transport"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   3
            Left            =   11565
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   900
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Data DatDados 
            Caption         =   "Vendedor"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   2
            Left            =   11610
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Select CodEntidade, Nome, Tipo from Ent Where  Tipo = 3 or Tipo = 7 Order By Nome"
            Top             =   1350
            Visible         =   0   'False
            Width           =   1860
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Credito"
            Height          =   315
            Index           =   6
            Left            =   7455
            TabIndex        =   40
            Tag             =   "S"
            Top             =   270
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Juros"
            Height          =   315
            Index           =   0
            Left            =   9270
            TabIndex        =   20
            Tag             =   "S"
            Top             =   270
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Desconto"
            Height          =   315
            Index           =   22
            Left            =   11070
            TabIndex        =   41
            Tag             =   "S"
            Top             =   270
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmEntidades.frx":0F2A
            DataField       =   "Transportadora"
            Height          =   315
            Index           =   1
            Left            =   225
            TabIndex        =   46
            Tag             =   "SI"
            Top             =   3015
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Nome"
            BoundColumn     =   "CodEntidade"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Salario"
            Height          =   315
            Index           =   21
            Left            =   2895
            TabIndex        =   37
            Tag             =   "S"
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmEntidades.frx":0F44
            DataField       =   "Vendedor"
            Height          =   315
            Index           =   0
            Left            =   8055
            TabIndex        =   48
            Tag             =   "SI"
            Top             =   3015
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Nome"
            BoundColumn     =   "CodEntidade"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone2"
            Height          =   315
            Index           =   7
            Left            =   225
            TabIndex        =   50
            Top             =   3585
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)-C###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fax2"
            Height          =   315
            Index           =   19
            Left            =   1755
            TabIndex        =   51
            Top             =   3585
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)-C###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Celular2"
            Height          =   315
            Index           =   20
            Left            =   3300
            TabIndex        =   52
            Top             =   3585
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "(##)-C###-####"
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmEntidades.frx":0F5E
            DataField       =   "Convenio"
            Height          =   315
            Index           =   3
            Left            =   4860
            TabIndex        =   53
            Tag             =   "SI"
            Top             =   3600
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Nome_Conv"
            BoundColumn     =   "CodConv"
            Text            =   ""
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pessoas autorizadas a Comprar no Nome"
            Height          =   195
            Index           =   33
            Left            =   225
            TabIndex        =   200
            Top             =   2205
            Width           =   2910
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Convênio"
            Height          =   195
            Index           =   32
            Left            =   4860
            TabIndex        =   199
            Top             =   3375
            Width           =   675
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fax 2:"
            Height          =   195
            Index           =   51
            Left            =   1755
            TabIndex        =   133
            Top             =   3375
            Width           =   435
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fone 2:"
            Height          =   195
            Index           =   52
            Left            =   225
            TabIndex        =   132
            Top             =   3375
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Celular 2:"
            Height          =   195
            Index           =   53
            Left            =   3300
            TabIndex        =   131
            Top             =   3375
            Width           =   660
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Vendedor/Comprador"
            Height          =   195
            Index           =   16
            Left            =   8055
            TabIndex        =   81
            Top             =   2790
            Width           =   1530
         End
         Begin VB.Label Label4 
            Caption         =   "Contato"
            Height          =   195
            Left            =   4410
            TabIndex        =   80
            Top             =   2790
            Width           =   945
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Salário"
            Height          =   195
            Index           =   67
            Left            =   2895
            TabIndex        =   79
            Top             =   60
            Width           =   480
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comissão Vendedor"
            Height          =   195
            Index           =   60
            Left            =   12060
            TabIndex        =   78
            Top             =   2790
            Width           =   1410
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Transportadora"
            Height          =   195
            Index           =   59
            Left            =   225
            TabIndex        =   77
            Top             =   2790
            Width           =   1080
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descontos"
            Height          =   195
            Index           =   58
            Left            =   11070
            TabIndex        =   76
            Top             =   75
            Width           =   765
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Index           =   7
            Left            =   9270
            TabIndex        =   75
            Top             =   75
            Width           =   375
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Status:"
            Height          =   195
            Index           =   57
            Left            =   255
            TabIndex        =   74
            Top             =   60
            Width           =   495
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dia Bom Pgto"
            Height          =   195
            Index           =   56
            Left            =   4470
            TabIndex        =   73
            Top             =   75
            Width           =   975
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Crédito"
            Height          =   195
            Index           =   13
            Left            =   7470
            TabIndex        =   72
            Top             =   75
            Width           =   495
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Prazo"
            Height          =   195
            Index           =   14
            Left            =   5655
            TabIndex        =   71
            Top             =   75
            Width           =   405
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmEntidades.frx":0F78
         Height          =   5610
         Left            =   -74865
         OleObjectBlob   =   "FrmEntidades.frx":0F92
         TabIndex        =   90
         Top             =   1260
         Width           =   13755
      End
      Begin MSDBGrid.DBGrid GridCto 
         Bindings        =   "FrmEntidades.frx":1E65
         Height          =   6225
         Left            =   -74775
         OleObjectBlob   =   "FrmEntidades.frx":1E7F
         TabIndex        =   152
         Top             =   540
         Width           =   13650
      End
      Begin MSDBGrid.DBGrid GridEstat 
         Bindings        =   "FrmEntidades.frx":2BB2
         Height          =   3390
         Left            =   -74865
         OleObjectBlob   =   "FrmEntidades.frx":2BCC
         TabIndex        =   154
         Top             =   450
         Width           =   13740
      End
      Begin MSDBGrid.DBGrid DbgFina 
         Bindings        =   "FrmEntidades.frx":4319
         Height          =   2505
         Left            =   -74820
         OleObjectBlob   =   "FrmEntidades.frx":4333
         TabIndex        =   172
         Top             =   3960
         Width           =   9315
      End
      Begin MSComCtl2.DTPicker DtData1 
         Height          =   330
         Left            =   -73065
         TabIndex        =   203
         Top             =   6975
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   45481985
         CurrentDate     =   36641
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   330
         Left            =   -71535
         TabIndex        =   204
         Top             =   6975
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   45481985
         CurrentDate     =   36641
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Final"
         Height          =   195
         Left            =   -71535
         TabIndex        =   206
         Top             =   6750
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inicial"
         Height          =   195
         Left            =   -73065
         TabIndex        =   205
         Top             =   6750
         Width           =   405
      End
   End
End
Attribute VB_Name = "FrmEntidades"
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

Private Sub DefineObjetos()
'Define quais objeto foram utilizados neste formulário

CTextBox = True         ' TxtDados
CMaskEditBox = True     ' MebDados
CComboBox = True        ' CmbDados
CListBox = False
CdbCombo = True        ' DbcDados
CDbList = False         '
CCheckBox = True        ' ChkDados
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo
' Titular = Registros.Fields("CodTit")
End Sub

Private Sub BtCidades_Click()
If Dir(App.Path & "\cidades.mdb") <> "" Then
    'CarregaCidade
'    If TxtDados(5).Text <> "" Then
'       PesquisaGeral App.Path & "\cidades.mdb", "SELECT tabcidades.codcidade, tabcidades.cep, tabcidades.cidade, tabcidades.uf FROM tabcidades  where tabcidades.cidade like " & """" & "*" & TxtDados(5).Text & "*" & """" & "  ORDER BY tabcidades.cidade", "Cidade"
'    Else
'        PesquisaGeral App.Path & "\cidades.mdb", "SELECT tabcidades.codcidade, tabcidades.cep, tabcidades.cidade, tabcidades.uf FROM tabcidades ORDER BY tabcidades.cidade", "Cidade"
'    End If
        
    FrmPesquisaCidade.Show vbModal
        
End If

End Sub

Function CarregaCidade()
FrmCidade.Visible = True
DatCidade.DatabaseName = App.Path & "\cidades.mdb"
DatCidade.RecordSource = "SELECT tabcidades.codcidade, tabcidades.cep, tabcidades.cidade, tabcidades.uf FROM tabcidades   ORDER BY tabcidades.cidade"
DatCidade.Refresh
DBLCidade.Refresh


End Function

Private Sub BtExcluir_Click()
If GridCto.Columns(0) <> 0 Then
    If MsgBox("Confirma exlusão", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        If DatDados(5).Recordset.RecordCount <> 0 Then
            DatDados(5).Recordset.Delete
        End If
    End If
End If
End Sub

Private Sub BtIbge_Click()

ShellExecute Me.Hwnd, "open", "https://cidades.ibge.gov.br/", _
               vbNullString, vbNullString, 0

End Sub

Private Sub BtIncluir_Click()

If TxtDados(0) <> 0 Then
    With DatDados(5).Recordset
        .AddNew
        .Fields("Entidade") = TxtDados(0)
        .Fields("Data") = Date
        .Update
        .Bookmark = .LastModified
        .Edit
        .Fields("Protocolo") = "CTO" & Right(Year(Date), 2) & psCompString(1, .Fields("CodCto"), 6)
        .Update
        GridCto.SetFocus
    End With
End If
End Sub

Private Sub BtOK_Click()
DBGrid1_DBLClick
tabCad.Tab = 1
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub


Private Sub ChkAtivos_Click()

End Sub

Private Sub ChkPago_Click()
If ChkPago.value = 0 Then
    DatDados(7).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & TxtDados(0).Text & " AND DUP.Cond = False ORDER BY DUP.DtVenc"
Else
    DatDados(7).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & TxtDados(0).Text & " ORDER BY DUP.DtVenc"
End If
DatDados(7).Refresh
SomaGrid

End Sub

Private Sub CmbCad_Click()
    
 '   If MsgBox("deseja ver a ficha ? ", vbYesNo) = vbYes Then
    
        'MDIPrincipal.ComRel.ShowPrinter
        MDIPrincipal.RptRel.ReportFileName = Caminho & "\Ficha.rpt"
        MDIPrincipal.RptRel.SelectionFormula = "{Ent.CodEntidade} = " & CInt(TxtDados(0))
        Imprimir
  '   End If
End Sub

Private Sub CmbDados_Click(Index As Integer)

'If Index = 4 Then
'    If CmbDados(4).ListIndex = 0 Then
'        Frame4.Enabled = False
'        Frame3.Enabled = True
'    End If
    
'    If CmbDados(4).ListIndex = 1 Then
'        Frame4.Enabled = True
'        Frame3.Enabled = False
'    End If
'End If
'
End Sub

Private Sub CmbDados_LostFocus(Index As Integer)
If Index = 0 Then
    If CmbDados(0).Text = "" Then
        CmbDados(0).ListIndex = 0
    End If
    
End If
'CmbDados_Click 4
If Index = 4 Then
     If CmbDados(4).ListIndex = 0 Then
        MebDados(10).BackColor = &HC0FFFF
        MebDados(11).BackColor = &HC0FFFF
        MebDados(12).BackColor = &H80000005
        MebDados(13).BackColor = &H80000005
     Else
        MebDados(10).BackColor = &H80000005
        MebDados(11).BackColor = &H80000005
        MebDados(12).BackColor = &HC0FFFF
        MebDados(13).BackColor = &HC0FFFF
     End If
 
End If

End Sub

Private Sub CmbFina_Click()
FormFina.Show
End Sub

Private Sub CmdFicha_Click()

End Sub


Private Sub CmbImprimir_Click()
Dim SQL As String

If MsgBox("Somente Notas ?", vbYesNo) = vbYes Then

    If DtData1 = Date And DtData2 = Date Then
        SQL = " SELECT  Ped.Cfo, Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade, Itp.Valor, Ped.Desconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & " and Ped.Cfo <> " & """" & "" & """" & " ORDER BY PED.DataEmiss"
    Else
        SQL = " SELECT  Ped.Cfo, Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade, Itp.Valor, Ped.Desconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & " AND  PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# and PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#  and Ped.Cfo <> " & """" & "" & """" & " ORDER BY PED.DataEmiss"
    End If

    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelConf.Relatorio.Recordset = varPesquisa(1)
    FrmRelConf.rpCabTitulo.Caption = "Relatório de Vendas"
    FrmRelConf.rpCabPeriodo.Caption = LblTitulo
    FrmRelConf.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelVdaCli.flag = ""
    FrmRelConf.Config

Else

    If DtData1 = Date And DtData2 = Date Then
        SQL = " SELECT  Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade, Itp.Valor, Ped.Desconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & "  ORDER BY PED.DataEmiss"
    Else
        SQL = " SELECT  Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade, Itp.Valor, Ped.Desconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & " AND  PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# and PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#  ORDER BY PED.DataEmiss"
    End If

    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelConf.Relatorio.Recordset = varPesquisa(1)
    FrmRelConf.rpCabTitulo.Caption = "Relatório de Vendas"
    FrmRelConf.rpCabPeriodo.Caption = LblTitulo
    FrmRelConf.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelVdaCli.flag = ""
    FrmRelConf.Config


End If

End Sub

Private Sub CmbPromi_Click()
            Dim varValor As Currency
            varValor = InputBox("Digite o Valor da Promissoria")

            If Not IsNumeric(varValor) Then
                varValor = 0
            End If
'             varPesquisa(1).re .Close
'            Set varPesquisa(1) = DatDados(0).Recordset.Clone
'            varPesquisa(1).Requery
 '           varPesquisa(1).MoveFirst
'            Set VarEnt = Banco.OpenRecordset("Select CodEntidade, Nome, Endereco, cpf, Nro, Cidade, Uf FROM Ent")
'            Criterio = "CodEntidade=" & CodEnt
'            VarEnt.FindFirst Criterio
            
'            DatDados(0).Recordset.MoveLast
 '           FrmRelCapa.NroPedido.Caption = Format(CodDoc, "000000")
            FrmRelCapa.ExtensoDia.Caption = Day(Date) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(Date))) & " DO ANO DE " & Year(Date)
            If varValor <> 0 Then
                FrmRelCapa.QuantiaExtenso.Caption = UCase(PassaExtenso(CStr(varValor))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                FrmRelCapa.ValorPromissoria.Caption = Format(varValor, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
                FrmRelCapa.DiaVenc.Caption = Day(Date)
                FrmRelCapa.MesVenc.Caption = UCase(RetornaMes(Month(Date + 30)))
                FrmRelCapa.AnoVenc.Caption = Year(Date)
            End If
            FrmRelCapa.Emitente.Caption = TxtDados(1) ' Left(VarEnt!Nome, 40)
            FrmRelCapa.Emitente2.Caption = TxtDados(1)  'Left(VarEnt!Nome, 40)
            DatDados(0).Recordset.MoveFirst
            If Not IsNull(MebDados(11)) Then
                FrmRelCapa.EmitenteCpf.Caption = MebDados(11)
            End If
            FrmRelCapa.Pagavel.Caption = CidadeEmpresa
            FrmRelCapa.EmitenteEndereco.Caption = TxtDados(2) & " - " & TxtDados(3)
            FrmRelCapa.EmitenteCidade.Caption = TxtDados(5) & " - " & TxtDados(6)
            FrmRelCapa.NomeEmpresa.Caption = NomeInscr
            FrmRelCapa.EmpresaCGCCPF.Caption = Cnpj
            FrmRelCapa.RpDia.Caption = Day(Date)
            FrmRelCapa.RpMes.Caption = UCase(RetornaMes(Month(Date)))
            FrmRelCapa.RpAno.Caption = Year(Date)
                    
'            FrmRelCapa.rpTitulo.Caption = Configura!Cabecalho1
'            FrmRelCapa.RpTitulo2.Caption = Configura!Cabecalho1
            FrmRelCapa.rpNome.Caption = TxtDados(1) ' Left(VarEnt!Nome, 40)
            FrmRelCapa.rpNome2.Caption = TxtDados(1) 'Left(VarEnt!Nome, 40)
 '           FrmRelCapa.RpMens.Caption = Configura!MensagemRodape
 '           FrmRelCapa.RpMens2.Caption = Configura!MensagemRodape
 '           FrmRelCapa.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
 '           FrmRelCapa.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                        
'            Set FrmRelCapa.Relatorio.Recordset = varPesquisa(1)
'            Me.Hide
            FrmRelCapa.Config
'            Me.Show vbModal

End Sub

Private Sub CmdAtu_Click()
TabCad_Click 5
End Sub

Private Sub CmdPesq_Click()
Dim Criterio As String
    
    If Dir(App.Path & "\Cadastro.mdb") <> "" Then
         DatCadastro.DatabaseName = App.Path & "\Cadastro.Mdb"
         DatCadastro.RecordSource = "SELECT ENT.*, CodEntidade FROM ENT ORDER BY CodEntidade"
         DatCadastro.Refresh
         If DatCadastro.Recordset.RecordCount Then
         
            If IsNumeric(TxtPesquisa) Then
                If Len(TxtPesquisa) > 13 Then
                        Criterio = DatCadastro.Recordset.Fields("cgc").Name & " =  " & """" & TxtPesquisa & """"
                Else
                    If Len(TxtPesquisa) > 10 Then
                        Criterio = DatCadastro.Recordset.Fields("Cpf").Name & " =  " & """" & TxtPesquisa & """"
                    ElseIf Len(TxtPesquisa) = 10 Then
                            Criterio = DatCadastro.Recordset.Fields("Celular").Name & " =  " & """" & TxtPesquisa & """" & " Or " & DatDados(0).Recordset.Fields("fone").Name & " =  " & """" & TxtPesquisa & """"
                        Else
                            Criterio = DatCadastro.Recordset.Fields(0).Name & " =  " & TxtPesquisa
                    End If
                End If
            Else
                   Criterio = DatDados(0).Recordset.Fields(3).Name & " LIKE " & """" & TxtPesquisa & "*"""

            End If
            DatCadastro.Recordset.FindFirst Criterio
            If Not DatCadastro.Recordset.NoMatch Then
                If MsgBox(DatCadastro.Recordset.Fields("Nome") & " - " & DatCadastro.Recordset.Fields("Endereco") & " Deseja incluir ? ", vbYesNo, App.Title) = vbYes Then
                    tabCad.Tab = 1
                    Manutencoes 1
'                    PermiteInclusao
                    MostraCadastro
                    MsgBox "Veririque se os dados estão corretos !", vbInformation
'                    PosMovimentacao
                End If
            End If
        
        End If
    End If
End Sub

Private Sub DatDados_Reposition(Index As Integer)
If Index = 0 Then
    If DatDados(0).Recordset.RecordCount > 0 Then
'        FrmEntidades.Caption = "  " & DatDados(0).Recordset("Nome")
        LblTitulo.Caption = DatDados(0).Recordset("Nome")
    End If
    
End If
End Sub

Private Sub DbcDados_LostFocus(Index As Integer)
If Index = 2 Then
End If
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodEntidade = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
tabCad.Tab = 1 'direcina para o TAB seguinte
If Registros.RecordCount > 0 Then
    MostraDados
    
End If
PosMovimentacao

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DBGrid1_DBLClick
    tabCad.Tab = 1

End If
End Sub

Private Sub DblPesquisaG_DblClick()
Dim SQL As String
FramePesquisaG.Visible = False
SQL = "Cidade = " & """" & DblPesquisaG.BoundText & """"
DatPesquisaG.Recordset.FindFirst SQL
TxtDados(5) = DatPesquisaG.Recordset.Fields("Cidade")
TxtDados(6) = DatPesquisaG.Recordset.Fields("Uf")
MebDados(1) = DatPesquisaG.Recordset.Fields("CEP")
MebDados(16) = DatPesquisaG.Recordset.Fields("CodCidade")
End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)

Formulário.Left = 100 '(MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = 50 '(MDIPrincipal.Height - Formulário.Height) / 6

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o banco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
tabCad.Tab = 0
AbilitaBotoes (False)
TxtPesquisa.SetFocus


If Admin = True Then
    CmbDados(3).Enabled = True
    TxtDados(7).Enabled = True
    MebDados(6).Enabled = True
Else
    CmbDados(3).Enabled = False
    TxtDados(7).Enabled = False
    MebDados(6).Enabled = False

End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
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
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
    
Set Registros = Banco.OpenRecordset("SELECT * From Ent Order by Nome")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select * From Ent Order by Nome"


LocalizacaoBasica = "Nome"
DefineObjetos
SelectFromSQL = "Select * From Ent"
SelectFromFiltro = "Select CodEntidade, Nome, Cidade, Uf from Ent"
PesquisaSql = "SELECT CodEntidade , Nome FROM Ent ORDER BY Nome"
FlagAbertura = False

OrderBySQL = "[Ent].[Nome]"

DtData1 = Date
DtData2 = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
   Confirmacoes (15)
End If

Registros.Close
AbilitaBotoes (True)

AtivarMenu (False)

End Sub

Public Sub PosInclusao()
' sera executada logo depois que o usuario iniciar o processo de inclusao addnew

TxtDados(7) = 0
MebDados(6) = 0
TxtDados(27) = 0

TxtDados(35) = 0
MebDados(0) = 0
MebDados(22) = 0
MebDados(21) = 0
CmbDados(0).ListIndex = 0
CmbDados(1).ListIndex = 0
'CmbDados(2).ListIndex = 0

CmbDados(4).ListIndex = 0
ChkDados(0).value = 1
CmbDados(0).SetFocus
tabCad.TabEnabled(0) = False
DbcDados(1).Text = "(Nenhum)"
DbcDados(0).Text = "(Nenhum)"
TxtDados(36).Text = "BRASIL"
MebDados(14).Text = Date
If LerINI("ENTIDADE", "Status", App.Path & "\config.ini") <> "" Then
    CmbDados(3).ListIndex = CInt(LerINI("Entidade", "Status", App.Path & "\config.ini")) - 1
Else
    CmbDados(3).ListIndex = 0
End If

If LerINI("ENTIDADE", "Prazo", App.Path & "\config.ini") <> "" Then
     TxtDados(7) = LerINI("Entidade", "Prazo", App.Path & "\config.ini")
Else
    TxtDados(7) = 0
End If

If LerINI("ENTIDADE", "Juros", App.Path & "\config.ini") <> "" Then
     MebDados(0) = LerINI("Entidade", "Juros", App.Path & "\config.ini")
Else
    MebDados(0) = 0
End If




If VarLimiteNaParcela = True Then
   CmbDados(3).ListIndex = 4
End If


End Sub

Public Function PermiteInclusao() As Boolean
' Será executado antes do processo de inclusao (addnew)
' Aqui deve ser feitos os processo p/ que o usuário pode ou não incluir um novo registro
' MsgBox "Neste formulário não pode ser feitas Inclusão de novos registros", vbInformation, App.Title


PermiteInclusao = True
PosMovimentacao


End Function

Public Sub PosAlteracao()
' sera executada logo depois que o usuario iniciar o processo de alteracao
'TxtDados(0).BackColor = QBColor(8) 'Muda a Cor de fundo do código

CmbDados(0).SetFocus
tabCad.TabEnabled(0) = False
tabCad.TabEnabled(2) = True
If TxtDados(27) = "" Then
        TxtDados(27) = 0
End If
FlagAltera = True

'Atualizar
End Sub

Public Sub PreConfirmacao() 'Será executada antes das rotinas de salvar ou cancelar

End Sub


Public Function PermiteAlteracao() As Boolean
' Será executado antes do processo de Alteracao


If Registros.Fields("Nome") <> "(Nenhum)" Then

    If MebDados(21).Text = "" Then
        MebDados(21) = 0
    End If
    PermiteAlteracao = True
    Open Caminho & "\Altera.txt" For Append As #1
    Print #1, "Entidade; " & Registros.Fields("CodEntidade") & "; " & Registros.Fields("Nome") & "; " & Date & "; " & Time & "; " & Maquina
    Close #1
Else
    MsgBox "Não pode alterar este registro", vbExclamation, App.Title
End If

If VarPermiteEnt = False Then
    PermiteAlteracao = True
Else
    PermiteAlteracao = False
End If

End Function

Public Sub PosMovimentacao()
'FrmEntidades.Caption = "  " & TxtDados(1).Text
LblTitulo.Caption = "  " & TxtDados(1).Text
VarPrazo = TxtDados(7).Text
    If CmbDados(4).Text = "Física" Then
       MebDados(10).BackColor = &HC0FFFF
       MebDados(11).BackColor = &HC0FFFF
       MebDados(12).BackColor = &H80000005
       MebDados(13).BackColor = &H80000005
    Else
       MebDados(10).BackColor = &H80000005
       MebDados(11).BackColor = &H80000005
       MebDados(12).BackColor = &HC0FFFF
       MebDados(13).BackColor = &HC0FFFF
    End If




End Sub

Public Sub PosExclusao()
' sera executada logo depois que o usuario iniciar o processo de exclusao

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
' Será executado antes do processo de exclusao
Dim VarPag As Recordset
Dim VarDup As Recordset
Set VarPag = Banco.OpenRecordset("Select * FROM Pag WHERE Fornecedor = " & TxtDados(0))
VarPag.Requery

Set VarDup = Banco.OpenRecordset("Select * FROM Dup WHERE Entidade = " & TxtDados(0))
VarDup.Requery


If VarPag.RecordCount > 0 Then
    MsgBox "Entidade não pode ser excluida pois já tem duplicatas lançadas !", vbInformation, App.Title
    PermiteExclusao = False
    Exit Function
End If

If VarDup.RecordCount > 0 Then
    MsgBox "Entidade não pode ser excluida pois já tem movimento !", vbInformation, App.Title
    PermiteExclusao = False
    Exit Function
End If

Dim Mensagem As String
If Registros.Fields("Nome") <> "(Nenhum)" Then
    Mensagem = "Confirma Exclusão de " & vbCr & Registros.Fields("Nome")
    
    If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes And VarPermiteEnt = False Then
    ' nesta posicao do codigo podemos incluir os procedimentos que deverao
    ' apagar todos os registro que dependem do registro que está sendo excluido, no caso de relacionamento entre tabelas
    ' ex : poderiamos excluir os itens de um pedido antes de excluir o cabeçalho
        PermiteExclusao = True
        Open Caminho & "\Delete.txt" For Append As #1
        Print #1, "Entidade; " & Registros.Fields("CodEntidade") & "; " & Registros.Fields("Nome") & "; " & Date & "; " & Time & "; " & Maquina
        Close #1
    Else
        PermiteExclusao = False
    End If
Else
    MsgBox "Não pode exluir este registro", vbExclamation, App.Title
End If

'If VarPermiteEnt = False Then
'    PermiteExclusao = True
'Else
'    PermiteExclusao = False
'End If

End Function


Public Sub PosConfirmacao() 'Será executada deposi dos processos de salvar ou cancelar

'TxtDados(0).BackColor = QBColor(7) 'Muda a Cor de fundo do código
TxtDados(0).Locked = False
Atualizar

End Sub

Public Function PermiteCancelamento() As Boolean
' Será executada antes do processo de cancelamento ser executado (cancelupdate)
' podemos aqui impedir o usuario de cancelar uma inclusao ou uma alteracao

PermiteCancelamento = True

End Function

Public Function PermiteSalvar() As Boolean
' Será axecutada antes de atribuir os campos do formulário no bco de dados
' Podemos fazer qualquer tipo de consistencia que não possa ser feita pelo bco de  dados
' Ex: O estoque máximo sempre deverá ser maior que o estoque minimo
If FlagAltera = False Then
        Dim RegCli As Recordset
        Set RegCli = Banco.OpenRecordset("Select Nome From Ent Where Nome = '" & TxtDados(1).Text & "'")
        If RegCli.RecordCount > 0 Then
            If MsgBox("Nome duplicado, deseja continuar ? ", vbYesNo, App.Title) = vbYes Then
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
PermiteContinuarSalvar = True

End Function

Public Function PosSalvamento() As Boolean
'Será executado assim que todos os dados forem gravados no bco
Dim VarSql As String

PosSalvamento = True

tabCad.TabEnabled(0) = True
tabCad.TabEnabled(2) = True

'If VarPrazo <> TxtDados(7) Then
'    VarSql = " UPDATE ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade SET DUP.DtVenc = [dup].[dtemiss]+[ent].[prazo] WHERE ENT.Prazo > '0'  AND DUP.Cond =False  AND ENT.CodEntidade= " & TxtDados(0)
'    Banco.Execute VarSql
'
'End If


End Function

Public Sub PosCancelamento()

'sera executado logo após do processo de cancelamento da aplicaçaõ

tabCad.TabEnabled(0) = True
tabCad.TabEnabled(2) = True

End Sub

Public Sub Atualizar()
'Registros.Requery
DatDados(0).Refresh
DatDados(2).Refresh
End Sub

Private Sub MebDados_LostFocus(Index As Integer)
Dim CpfCorreto As Boolean

Select Case Index
    Case 11
        If MebDados(11) <> "" Then
            
            If Not CalculaDacCPF(Format(MebDados(11), "@@@.@@@.@@@-@@")) Then
                MsgBox "CPF Inválido", vbInformation
                MebDados(11).SetFocus
            End If
        End If
    Case 12
        If MebDados(12) <> "" Then
            If VerificaCGC(Format(MebDados(12), "@@.@@@.@@@/@@@@-@@")) = True Then
'                MsgBox "CNPJ Válido.", vbInformation
            Else
                MsgBox "CNPJ Inválido.", vbInformation
                MebDados(12).SetFocus
            End If
        End If
    Case 21
        If IsNumeric(MebDados(21)) Then
            MebDados(6) = 0
            MebDados(6) = MebDados(21) * (20 / 100)
'            CmbDados(3).ListIndex = 4
        End If

End Select
End Sub

Private Sub Opt1_Click(Index As Integer)

'Ajustar Mascaras

Select Case Index
    Case 0
        MebDados(5).Mask = "##.###.###/####/##"
    Case 1
        MebDados(5).Mask = "###.###.###-##"
End Select

End Sub


Private Sub OptPesq_Click(Index As Integer)
Select Case Index
    Case 0
        DatDados(0).RecordSource = "Select * From Ent Where Tipo = 1 or Tipo = 5 or Tipo = 7  Order by Nome"
    Case 1
        DatDados(0).RecordSource = "Select * From Ent Where Tipo = 2 or Tipo =  5  Order by Nome"
    Case 2
        DatDados(0).RecordSource = "Select * From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
    Case 3
        DatDados(0).RecordSource = "Select * From Ent  Order by Nome"
End Select
DatDados(0).Refresh
End Sub

Private Sub TabCad_Click(previoustab As Integer)
Dim ContPed As Integer
Dim VarTotal As Currency
Dim VarVista As Currency
Dim VarParzo As Currency
Dim VarOrca As Currency

Select Case tabCad.Tab

    Case 0
        Atualizar
        AbilitaBotoes (False)
    Case 5
        If IsNumeric(TxtDados(0)) Then
            ContPed = 0
            VarTotal = 0
            VarVista = 0
            VarPrazo = 0
            VarOrca = 0
            If DtData1 = Date And DtData2 = Date Then
                DatDados(6).RecordSource = " SELECT Itp.Referencia, Ped.NroPedido, Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & " AND Ped.NroPedido <> " & """" & "999999" & """" & "  ORDER BY PED.DataEmiss"
            Else
                DatDados(6).RecordSource = " SELECT Itp.Referencia, Ped.NroPedido, Ped.Operacao, PED.CodPed, PED.DataEmiss, PRD.Descrição_Produto, ITP.Qtd, Itp.ValorUnit, ITP.VlrVdaDia, Itp.Loja, Itp.Entregue, PED.Entidade FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE  PED.Entidade = " & TxtDados(0) & " AND  PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# and PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#  AND Ped.NroPedido <> " & """" & "999999" & """" & "  ORDER BY PED.DataEmiss"
            End If
            DatDados(6).Refresh
            If DatDados(6).Recordset.RecordCount > 0 Then
                DatDados(6).Recordset.MoveFirst
                MebPrimeiraCompra = DatDados(6).Recordset.Fields("DataEmiss")
                While Not DatDados(6).Recordset.Eof
                    If DatDados(6).Recordset("operacao") = 1 Then
                        VarVista = VarVista + (DatDados(6).Recordset("VlrVdaDia"))
                        VarTotal = VarTotal + (DatDados(6).Recordset("VlrVdaDia"))
                    End If
                    If DatDados(6).Recordset("operacao") = 2 Then
                        VarPrazo = VarPrazo + (DatDados(6).Recordset("VlrVdaDia"))
                        VarTotal = VarTotal + (DatDados(6).Recordset("VlrVdaDia"))
                    End If
                     If DatDados(6).Recordset("operacao") = 4 Then
                        VarOrca = VarOrca + (DatDados(6).Recordset("VlrVdaDia"))
                    End If
                    
                    ContPed = ContPed + 1
                    DatDados(6).Recordset.MoveNext
                Wend
                
                DatDados(6).Recordset.MoveLast
                MebUltimaCompra = DatDados(6).Recordset.Fields("DataEmiss")
                If DateDiff("d", CDate(MebPrimeiraCompra), CDate(MebUltimaCompra)) <> 0 Then
                    MebMediaCompra = VarTotal / (DateDiff("d", CDate(MebPrimeiraCompra), CDate(MebUltimaCompra)) / 30)
                End If
                MebPrazo = VarPrazo
                MebVista = VarVista
                MebOrcamento = VarOrca
                
'                DatDados(6).Refresh

                
            End If
            
            DatDados(7).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & TxtDados(0).Text & " ORDER BY DUP.DtVenc"
            DatDados(7).Refresh
            SomaGrid
            
            DatDados(7).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & TxtDados(0).Text & " AND DUP.Cond = False ORDER BY DUP.DtVenc"
            DatDados(7).Refresh
            If DatDados(7).Recordset.RecordCount > 0 Then
                DatDados(7).Recordset.MoveFirst
                MebDiasAtraso = DateDiff("d", DatDados(7).Recordset("DtVenc"), Date)

            End If
            
'            SomaGrid

      
            
        End If
    Case Else
        AbilitaBotoes (True)
End Select

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

Private Sub TxtDados_GotFocus(Index As Integer)
If Index = 40 Then
    Me.KeyPreview = False
End If
End Sub

Private Sub TxtDados_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 5 And PctFundo(0).Enabled = True And KeyAscii = 13 Then
End If

End Sub

Private Sub TxtDados_LostFocus(Index As Integer)

Select Case Index

    Case 25
    tabCad.Tab = 2
    Case 38
        tabCad.Tab = 3
    Case 40
        Me.KeyPreview = True
End Select
End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        If Len(TxtPesquisa) > 13 Then
                Criterio = DatDados(0).Recordset.Fields("cgc").Name & " =  " & """" & TxtPesquisa & """"
        Else
            If Len(TxtPesquisa) >= 11 Then
                Criterio = DatDados(0).Recordset.Fields("Cpf").Name & " =  " & """" & TxtPesquisa & """"
            ElseIf Len(TxtPesquisa) = 9 Then
                    Criterio = DatDados(0).Recordset.Fields("Celular").Name & " like  " & """" & "*" & TxtPesquisa & "*" & """" & " Or " & DatDados(0).Recordset.Fields("fone").Name & " like  " & """" & "*" & TxtPesquisa & "*" & """"
                Else
                    Criterio = DatDados(0).Recordset.Fields(0).Name & " =  " & TxtPesquisa
            End If
        End If
    Else
        If ChkPesq.value = 0 Then
           Criterio = DatDados(0).Recordset.Fields(3).Name & " LIKE " & """" & TxtPesquisa & "*"""
'        Else
'            DatDados(0).RecordSource = "Select * From Ent Where Nome Like *" & TxtPesquisa & "*" & " Order by Nome"
        End If
    End If
    
    If ChkPesq.value = 0 Then
        DatDados(0).Recordset.FindFirst Criterio
        If DatDados(0).Recordset.NoMatch Then
            TxtPesquisa.SetFocus
        Else
            Registros.FindFirst Criterio
            MostraDados
            
            PosMovimentacao
        End If
    End If
End If

End Sub

Private Sub TxtPesquisa_GotFocus()
LblMensagem = "Digite o código / nome / Cpf ou Cnpj / Fone ou Celular da entidade "
End Sub

Private Sub TxtPesquisa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ChkPesq.value = 1 Then
'    Sql = "Select * From Ent Where Nome Like *" & TxtPesquisa & "*" & " Order by Nome"
        DatDados(0).RecordSource = "Select * From Ent Where Nome Like " & """" & "*" & TxtPesquisa & "*" & """" & " Order by Nome"
        DatDados(0).Refresh
        If DatDados(0).Recordset.RecordCount > 0 Then
            MostraDados
            
            PosMovimentacao
        End If
    Else
        DBGrid1_DBLClick
    End If
End If
End Sub

Private Sub TxtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    DBGrid1.SetFocus
End If

End Sub

Private Sub TxtPesquisa_LostFocus()
LblMensagem = ""
End Sub



Public Sub MostraCadastro()


Dim Mascara As String
Dim Componente As Variant
With Me
    If .CTextBox Then
        For Each Componente In .TxtDados
            If Componente.DataField <> "CodEntidade" Then
                Componente.Text = IIf(IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)), "", .DatCadastro.Recordset.Fields(Componente.DataField))
            End If
        Next
     End If
     If .CMaskEditBox Then
        For Each Componente In .MebDados
               If .DatCadastro.Recordset.Fields(Componente.DataField).Type = dbDate Then
                    ' a propriedade type de um campo represente o tipo do campo dentro do banco de dados
                    If IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)) Then
                        Mascara = Componente.Mask
                        Componente.Mask = ""
                        Componente.Text = ""
                        Componente.Mask = Mascara
                    Else
                        Componente.Text = Format(.DatCadastro.Recordset.Fields(Componente.DataField), "dd/mm/yyyy")
                    End If
                Else
                    Componente.Text = IIf(IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)), "", .DatCadastro.Recordset.Fields(Componente.DataField))
                End If
        Next
     End If
     If .CComboBox Then
        For Each Componente In .CmbDados
            If Mid(Componente.Tag, 2, 1) = "I" Then
                Componente.ListIndex = IIf(IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)), -1, .DatCadastro.Recordset.Fields(Componente.DataField)) - 1
            Else
                Componente.Text = IIf(IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)), "", .DatCadastro.Recordset.Fields(Componente.DataField))
            End If
        Next
     End If
     
     If .CdbCombo Then
        For Each Componente In .DbcDados
            If Mid(Componente.Tag, 1, 1) = "S" Then
                Componente.BoundText = IIf(IsNull(.DatCadastro.Recordset.Fields(Componente.DataField)), "", .DatCadastro.Recordset.Fields(Componente.DataField))
                Componente.Refresh
            End If
        Next
     End If
     
     If .CCheckBox Then
        For Each Componente In .ChkDados
            If Componente.Tag = "S" Then
                Componente.value = IIf(.DatCadastro.Recordset.Fields(Componente.DataField) = True, 1, 0)
            Else
                Componente.value = IIf(.DatCadastro.Recordset.Fields(Componente.DataField) = "SIM", 1, 0)
            End If
        Next
     End If
     If .ComandBotom Then
'        For Each Componente In .CmdDados
'            Componente.Enabled = Not .DatCadastro.Recordset.Fields(Componente.Tag)
'        Next
     End If
     
     
     
End With

End Sub


Private Sub SomaGrid()
Dim SomaGrid As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Dim ContPgto As Integer
Dim VarMedia As Double
Pago = 0
Apagar = 0
ContPgto = 0
VarMedia = 0

'Set SomaGrid = DatDados(0).Recordset.Clone
If DatDados(7).Recordset.RecordCount <> 0 Then
    DatDados(7).Recordset.MoveFirst
    Do Until DatDados(7).Recordset.Eof
        If GetInputState() Then DoEvents

        If DatDados(7).Recordset.Fields("Cond") = True Then
            Pago = Pago + DatDados(7).Recordset.Fields("VlrPago")
            If Not IsNull(DatDados(7).Recordset("DtPgto")) Then
                MebUltimoPgto = DatDados(7).Recordset("DtPgto")
                ContPgto = ContPgto + 1
                If Not IsNull(DatDados(7).Recordset("DtVenc")) Then
                    VarMedia = VarMedia + DateDiff("d", CDate(DatDados(7).Recordset("DtVenc")), CDate(DatDados(7).Recordset("DtPgto")))
                End If
                MebMediaPgto = VarMedia / ContPgto

            End If
        Else
            Apagar = Apagar + DatDados(7).Recordset.Fields("Valor")
        End If
       DatDados(7).Recordset.MoveNext
    Loop
'    DatDados(7).Recordset.MovePrevious
    MebPagos.Text = Pago
    MebAPagar.Text = Apagar
    MebSubTot = Pago + Apagar

End If

End Sub










