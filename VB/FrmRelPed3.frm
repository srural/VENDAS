VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPed3 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   9765
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   172.244
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   495
      TabIndex        =   0
      Top             =   5220
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Orientacao      =   2
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   2985
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5265
      Tipo            =   1
      Ordem           =   1
      Begin ReportX.ReportField rpCabInformacao 
         Height          =   240
         Index           =   0
         Left            =   1650
         TabIndex        =   5
         Top             =   720
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   423
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField10 
         Height          =   330
         Left            =   6255
         TabIndex        =   6
         Top             =   1020
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         Caption         =   "PEDIDO"
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   480
         Index           =   0
         Left            =   1650
         TabIndex        =   1
         Top             =   150
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   847
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField23 
         Height          =   225
         Left            =   1185
         TabIndex        =   59
         Top             =   1395
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   397
         Campo           =   "Nome"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField24 
         Height          =   225
         Left            =   1185
         TabIndex        =   60
         Top             =   1899
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   397
         Campo           =   "Cidade"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField25 
         Height          =   225
         Left            =   6045
         TabIndex        =   61
         Top             =   1899
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         Campo           =   "Uf"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField26 
         Height          =   225
         Left            =   1185
         TabIndex        =   62
         Top             =   2151
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   397
         Campo           =   "Fone"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   225
         Left            =   3315
         TabIndex        =   63
         Top             =   2151
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   397
         Campo           =   "Fax"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField28 
         Height          =   225
         Left            =   7575
         TabIndex        =   64
         Top             =   2190
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         Campo           =   "Cep"
         Formato         =   "@@.@@@-@@@"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpCgc 
         Height          =   225
         Left            =   1185
         TabIndex        =   65
         Top             =   2403
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpInscr 
         Height          =   225
         Left            =   4470
         TabIndex        =   66
         Top             =   2403
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   285
         Left            =   9945
         TabIndex        =   67
         Top             =   1380
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         Campo           =   "CodPed"
         Caption         =   "Nome"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField33 
         Height          =   225
         Left            =   9945
         TabIndex        =   68
         Top             =   1680
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   397
         Campo           =   "DataEmiss"
         Caption         =   "Nome"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField38 
         Height          =   225
         Left            =   1185
         TabIndex        =   69
         Top             =   1647
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   397
         Campo           =   "Endereco"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField39 
         Height          =   225
         Left            =   7035
         TabIndex        =   70
         Top             =   1695
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         Campo           =   "Nro"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField40 
         Height          =   210
         Left            =   7275
         TabIndex        =   71
         Top             =   1965
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   370
         Campo           =   "Bairro"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField41 
         Height          =   225
         Left            =   7050
         TabIndex        =   72
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         Campo           =   "CodEntidade"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField42 
         Height          =   225
         Left            =   8145
         TabIndex        =   73
         Top             =   1680
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         Campo           =   "NroPedido"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   225
         Left            =   8565
         TabIndex        =   74
         Top             =   2430
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   397
         Campo           =   "InscrProd"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   225
         Left            =   9855
         TabIndex        =   75
         Top             =   2190
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   397
         Campo           =   "Ent.Complemento"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   225
         Left            =   5430
         TabIndex        =   76
         Top             =   2151
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   397
         Campo           =   "Celular"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   225
         Left            =   1170
         TabIndex        =   199
         Top             =   2655
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   397
         Campo           =   "Fantasia"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpCabInformacao 
         Height          =   240
         Index           =   1
         Left            =   9795
         TabIndex        =   202
         Top             =   735
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   423
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   480
         Index           =   1
         Left            =   9795
         TabIndex        =   203
         Top             =   210
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   847
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Logotipo 
         Height          =   960
         Index           =   1
         Left            =   135
         Stretch         =   -1  'True
         Top             =   90
         Width           =   7950
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fantasia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   198
         Top             =   2655
         Width           =   765
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6615
         TabIndex        =   94
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6645
         TabIndex        =   93
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6645
         TabIndex        =   92
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uf:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5745
         TabIndex        =   91
         Top             =   1950
         Width           =   225
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc/Rg:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   90
         Top             =   2430
         Width           =   675
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7155
         TabIndex        =   89
         Top             =   2190
         Width           =   375
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2895
         TabIndex        =   88
         Top             =   2190
         Width           =   345
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9465
         TabIndex        =   87
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9255
         TabIndex        =   86
         Top             =   1410
         Width           =   630
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cnpj/Cpf:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   85
         Top             =   2430
         Width           =   750
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   84
         Top             =   2190
         Width           =   780
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   83
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   82
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   510
         TabIndex        =   81
         Top             =   1440
         Width           =   540
      End
      Begin VB.Shape Shape9 
         Height          =   1590
         Left            =   90
         Top             =   1350
         Width           =   11220
      End
      Begin VB.Line Line24 
         X1              =   7995
         X2              =   7995
         Y1              =   1350
         Y2              =   1920
      End
      Begin VB.Line Line25 
         X1              =   7995
         X2              =   11295
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8175
         TabIndex        =   80
         Top             =   1410
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cx.Postal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8955
         TabIndex        =   79
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscr.Produtor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7215
         TabIndex        =   78
         Top             =   2430
         Width           =   1230
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cel.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4995
         TabIndex        =   77
         Top             =   2190
         Width           =   360
      End
      Begin VB.Image Logotipo 
         Height          =   960
         Index           =   0
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7950
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   7635
         X2              =   8085
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   120
         X2              =   6210
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   4485
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
      Begin ReportX.ReportField ReportField8 
         Height          =   225
         Left            =   165
         TabIndex        =   13
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   397
         Campo           =   "Produto"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   225
         Left            =   1260
         TabIndex        =   14
         Top             =   0
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   397
         Campo           =   "Descrição_Produto"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpDetalhes 
         Height          =   225
         Index           =   0
         Left            =   7830
         TabIndex        =   15
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   397
         Campo           =   "forEntregue"
         Caption         =   "Nome"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpDetalhes 
         Height          =   225
         Index           =   2
         Left            =   9225
         TabIndex        =   16
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   397
         Campo           =   "forUnitario"
         Formato         =   "##,##0.0000"
         Caption         =   "Nome"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpDetalhes 
         Height          =   225
         Index           =   1
         Left            =   8205
         TabIndex        =   17
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Campo           =   "Qtd"
         Caption         =   "Nome"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpDetalhes 
         Height          =   225
         Index           =   3
         Left            =   10170
         TabIndex        =   18
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         Campo           =   "forVlrVdaDia"
         Formato         =   "Standard"
         Caption         =   "Nome"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   225
         Left            =   5805
         TabIndex        =   32
         Top             =   0
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   397
         Campo           =   "Itp.Complemento"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   225
         Left            =   7455
         TabIndex        =   35
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   397
         Campo           =   "Embalagem"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField29 
         Height          =   225
         Left            =   1260
         TabIndex        =   196
         Top             =   225
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   397
         Campo           =   "Referencia"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line16 
         X1              =   120
         X2              =   120
         Y1              =   -30
         Y2              =   390
      End
      Begin VB.Line Line15 
         X1              =   11340
         X2              =   11340
         Y1              =   -30
         Y2              =   390
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   11340
         Y1              =   225
         Y2              =   225
      End
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   10320
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         DataField       =   "forSoma"
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
         Left            =   180
         TabIndex        =   2
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   1455
      Left            =   0
      Top             =   8775
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2566
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8910
         TabIndex        =   3
         Top             =   1050
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Campo           =   "=Página [Pagina]"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   4
         Top             =   1020
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   370
         Campo           =   "=Emitido em [Hoje] as [Hora]"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField30 
         Height          =   285
         Left            =   10125
         TabIndex        =   200
         Top             =   45
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         Campo           =   "CodPed"
         Caption         =   "Nome"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9450
         TabIndex        =   201
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SidComp - Automação Comercial   (17) 3552-1074     Cel. (17) 8116-0099"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   197
         Top             =   1035
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.Shape Shape5 
         Height          =   645
         Left            =   9945
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E ESTAR DE PLENO ACORDO COM ESTE PEDIDO "
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
         Left            =   4635
         TabIndex        =   195
         Top             =   90
         Width           =   4695
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "carimbo"
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
         Left            =   10395
         TabIndex        =   170
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Data:  ____/____/________"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   169
         Top             =   585
         Width           =   2760
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DECLARO TER  RECEBIDO  AS  MERCADORIAS"
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
         Left            =   15
         TabIndex        =   168
         Top             =   90
         Width           =   4440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Completo - Rg"
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
         Left            =   4725
         TabIndex        =   167
         Top             =   765
         Width           =   1755
      End
      Begin VB.Line Line30 
         X1              =   3495
         X2              =   7485
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line lneRod 
         X1              =   11430
         X2              =   90
         Y1              =   1020
         Y2              =   1020
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   1155
      Index           =   1
      Left            =   0
      Top             =   2985
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2037
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   7410
         TabIndex        =   9
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   370
         Campo           =   "Cidade_cob"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   9990
         TabIndex        =   11
         Top             =   270
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   370
         Campo           =   "Uf_Cob"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   370
         Campo           =   "Endereço_Cob"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   210
         Left            =   4710
         TabIndex        =   30
         Top             =   270
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Campo           =   "Nro_Cob"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   5400
         TabIndex        =   31
         Top             =   270
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   370
         Campo           =   "Bairro_Cob"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   7410
         TabIndex        =   44
         Top             =   855
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   370
         Campo           =   "Cidade_Entrega"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   9990
         TabIndex        =   45
         Top             =   855
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   370
         Campo           =   "Uf_Entrega"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   210
         TabIndex        =   46
         Top             =   855
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   370
         Campo           =   "Endereço_Entrega"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   4710
         TabIndex        =   47
         Top             =   855
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Campo           =   "Nro_Entrega"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   5400
         TabIndex        =   48
         Top             =   855
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   370
         Campo           =   "Bairro_Entrega"
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField20 
         Height          =   210
         Left            =   10320
         TabIndex        =   55
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Cep_Cob"
         Formato         =   "@@.@@@-@@@"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField21 
         Height          =   210
         Left            =   10320
         TabIndex        =   57
         Top             =   870
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Cep_Entrega"
         Formato         =   "@@.@@@-@@@"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   58
         Top             =   690
         Width           =   330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   56
         Top             =   90
         Width           =   330
      End
      Begin VB.Shape Shape1 
         Height          =   525
         Left            =   105
         Top             =   600
         Width           =   11235
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   53
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7380
         TabIndex        =   52
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9960
         TabIndex        =   51
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4710
         TabIndex        =   50
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   49
         Top             =   675
         Width           =   525
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   28
         Top             =   90
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4710
         TabIndex        =   12
         Top             =   75
         Width           =   300
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9990
         TabIndex        =   10
         Top             =   90
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7410
         TabIndex        =   8
         Top             =   90
         Width           =   585
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Cobrança"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   45
         Width           =   1680
      End
      Begin VB.Shape Shape2 
         Height          =   525
         Left            =   105
         Top             =   30
         Width           =   11235
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   2400
      Index           =   1
      Left            =   0
      Top             =   4935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   4233
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField rpGrupo 
         Height          =   240
         Index           =   0
         Left            =   9585
         TabIndex        =   22
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         Campo           =   "forSubTotal"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   1
         Left            =   7110
         TabIndex        =   23
         Top             =   30
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   397
         Campo           =   "ValorFrete"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   270
         Index           =   2
         Left            =   9585
         TabIndex        =   24
         Top             =   960
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   476
         Campo           =   "forTotalPedido"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Obs3 
         Height          =   210
         Left            =   255
         TabIndex        =   21
         Top             =   465
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Obs2 
         Height          =   210
         Left            =   255
         TabIndex        =   20
         Top             =   255
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Obs1 
         Height          =   210
         Left            =   255
         TabIndex        =   19
         Top             =   45
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   240
         Index           =   3
         Left            =   9585
         TabIndex        =   33
         Top             =   285
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         Campo           =   "Desconto"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Obs4 
         Height          =   210
         Left            =   255
         TabIndex        =   54
         Top             =   675
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Transport 
         Height          =   210
         Left            =   1605
         TabIndex        =   171
         Top             =   1740
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   1380
         TabIndex        =   172
         Top             =   990
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   370
         Campo           =   "forTotalItens"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpPrevEntrega 
         Height          =   225
         Left            =   1605
         TabIndex        =   173
         Top             =   1470
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField37 
         Height          =   225
         Left            =   5520
         TabIndex        =   174
         Top             =   1470
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   397
         Campo           =   "TipoFrete"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Vendedor 
         Height          =   225
         Left            =   8205
         TabIndex        =   175
         Top             =   1470
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   397
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpCondpgto 
         Height          =   225
         Left            =   4830
         TabIndex        =   181
         Top             =   2070
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpOperacao 
         Height          =   225
         Left            =   1170
         TabIndex        =   182
         Top             =   2070
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   4
         Left            =   7110
         TabIndex        =   185
         Top             =   285
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   397
         Campo           =   "forIpi"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpContato 
         Height          =   225
         Left            =   8205
         TabIndex        =   187
         Top             =   1740
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField RpFormaPgto 
         Height          =   225
         Left            =   8580
         TabIndex        =   189
         Top             =   2070
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   240
         Index           =   5
         Left            =   9585
         TabIndex        =   192
         Top             =   585
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         Campo           =   "forAcrescimo"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   6
         Left            =   7110
         TabIndex        =   194
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   397
         Campo           =   "PedSub"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line4 
         X1              =   6390
         X2              =   6390
         Y1              =   0
         Y2              =   855
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acresc.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   8700
         TabIndex        =   193
         Top             =   585
         Width           =   705
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subst.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6480
         TabIndex        =   191
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   190
         Top             =   2070
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7350
         TabIndex        =   188
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.p.i.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   6660
         TabIndex        =   186
         Top             =   285
         Width           =   375
      End
      Begin VB.Shape Shape4 
         Height          =   345
         Left            =   120
         Top             =   2010
         Width           =   11235
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pgto:"
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
         Left            =   7245
         TabIndex        =   184
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label LabelCondicao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cond. Pgto:"
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
         Left            =   3645
         TabIndex        =   183
         Top             =   2070
         Width           =   1125
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7200
         TabIndex        =   180
         Top             =   1470
         Width           =   870
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Entrega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4335
         TabIndex        =   179
         Top             =   1470
         Width           =   1110
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Entrega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   178
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Shape Shape3 
         Height          =   645
         Left            =   120
         Top             =   1380
         Width           =   11235
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Itens:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   177
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportadora:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   176
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Line Line29 
         X1              =   6390
         X2              =   11340
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line28 
         X1              =   6390
         X2              =   11340
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line27 
         X1              =   6390
         X2              =   11340
         Y1              =   255
         Y2              =   255
      End
      Begin VB.Line Line14 
         X1              =   8475
         X2              =   8475
         Y1              =   0
         Y2              =   1260
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   90
         X2              =   11310
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape Shape10 
         Height          =   1290
         Left            =   120
         Top             =   0
         Width           =   11235
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   8580
         TabIndex        =   34
         Top             =   285
         Width           =   855
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub.Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   8595
         TabIndex        =   27
         Top             =   30
         Width           =   840
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   6570
         TabIndex        =   26
         Top             =   30
         Width           =   480
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tota.l:"
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
         Index           =   0
         Left            =   8790
         TabIndex        =   25
         Top             =   930
         Width           =   600
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   1440
      Index           =   2
      Left            =   0
      Top             =   7335
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2540
      Tipo            =   6
      Ordem           =   2
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   0
         Left            =   1080
         TabIndex        =   95
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   0
         Left            =   2010
         TabIndex        =   96
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   2
         Left            =   5835
         TabIndex        =   97
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   2
         Left            =   6765
         TabIndex        =   98
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   2
         Left            =   7695
         TabIndex        =   99
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   3
         Left            =   9600
         TabIndex        =   100
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   3
         Left            =   10500
         TabIndex        =   101
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   5
         Left            =   3030
         TabIndex        =   102
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   5
         Left            =   3930
         TabIndex        =   103
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   5
         Left            =   4860
         TabIndex        =   104
         Top             =   285
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   6
         Left            =   6765
         TabIndex        =   105
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   6
         Left            =   7695
         TabIndex        =   106
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   8
         Left            =   165
         TabIndex        =   107
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   8
         Left            =   1080
         TabIndex        =   108
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   8
         Left            =   2010
         TabIndex        =   109
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   9
         Left            =   3930
         TabIndex        =   110
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   9
         Left            =   4860
         TabIndex        =   111
         Top             =   495
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   11
         Left            =   8685
         TabIndex        =   112
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   11
         Left            =   9600
         TabIndex        =   113
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   11
         Left            =   10500
         TabIndex        =   114
         Top             =   495
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   12
         Left            =   1080
         TabIndex        =   115
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   12
         Left            =   2010
         TabIndex        =   116
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   14
         Left            =   5835
         TabIndex        =   117
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   14
         Left            =   6765
         TabIndex        =   118
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   14
         Left            =   7695
         TabIndex        =   119
         Top             =   705
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   15
         Left            =   9600
         TabIndex        =   120
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   15
         Left            =   10500
         TabIndex        =   121
         Top             =   705
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   17
         Left            =   3030
         TabIndex        =   122
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   17
         Left            =   3930
         TabIndex        =   123
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   17
         Left            =   4860
         TabIndex        =   124
         Top             =   915
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   125
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   3
         Left            =   8685
         TabIndex        =   126
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   6
         Left            =   5835
         TabIndex        =   127
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   9
         Left            =   3030
         TabIndex        =   128
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   12
         Left            =   165
         TabIndex        =   129
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   15
         Left            =   8685
         TabIndex        =   130
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   1
         Left            =   3030
         TabIndex        =   131
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   1
         Left            =   3930
         TabIndex        =   132
         Top             =   75
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   1
         Left            =   4860
         TabIndex        =   133
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   4
         Left            =   165
         TabIndex        =   134
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   4
         Left            =   1080
         TabIndex        =   135
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   4
         Left            =   2010
         TabIndex        =   136
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   7
         Left            =   8685
         TabIndex        =   137
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   7
         Left            =   9600
         TabIndex        =   138
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   7
         Left            =   10500
         TabIndex        =   139
         Top             =   285
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   10
         Left            =   5835
         TabIndex        =   140
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   10
         Left            =   6765
         TabIndex        =   141
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   10
         Left            =   7695
         TabIndex        =   142
         Top             =   495
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   13
         Left            =   3030
         TabIndex        =   143
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   13
         Left            =   3930
         TabIndex        =   144
         Top             =   705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   13
         Left            =   4860
         TabIndex        =   145
         Top             =   705
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   16
         Left            =   165
         TabIndex        =   146
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   16
         Left            =   1080
         TabIndex        =   147
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   16
         Left            =   2010
         TabIndex        =   148
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   18
         Left            =   5835
         TabIndex        =   149
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   18
         Left            =   6765
         TabIndex        =   150
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   18
         Left            =   7695
         TabIndex        =   151
         Top             =   915
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   19
         Left            =   8685
         TabIndex        =   152
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   19
         Left            =   9600
         TabIndex        =   153
         Top             =   915
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   19
         Left            =   10500
         TabIndex        =   154
         Top             =   915
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   20
         Left            =   165
         TabIndex        =   155
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   20
         Left            =   1080
         TabIndex        =   156
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   20
         Left            =   2010
         TabIndex        =   157
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   21
         Left            =   3030
         TabIndex        =   158
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   21
         Left            =   3930
         TabIndex        =   159
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   21
         Left            =   4860
         TabIndex        =   160
         Top             =   1125
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   22
         Left            =   5835
         TabIndex        =   161
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   22
         Left            =   6765
         TabIndex        =   162
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   22
         Left            =   7695
         TabIndex        =   163
         Top             =   1125
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Pcl 
         Height          =   210
         Index           =   23
         Left            =   8685
         TabIndex        =   164
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   23
         Left            =   9600
         TabIndex        =   165
         Top             =   1125
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField Valor 
         Height          =   210
         Index           =   23
         Left            =   10500
         TabIndex        =   166
         Top             =   1125
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   0
         Left            =   90
         Top             =   30
         Width           =   2820
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   2
         Left            =   2985
         Top             =   30
         Width           =   2760
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   3
         Left            =   5805
         Top             =   30
         Width           =   2790
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   4
         Left            =   8640
         Top             =   30
         Width           =   2730
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   3
      Left            =   0
      Top             =   4140
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Tipo            =   3
      Ordem           =   2
      Begin VB.Label lblGrupo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entregue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   7845
         TabIndex        =   43
         Top             =   90
         Width           =   750
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7365
         TabIndex        =   42
         Top             =   90
         Width           =   405
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5820
         TabIndex        =   41
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label lblGrupo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   10785
         TabIndex        =   40
         Top             =   90
         Width           =   405
      End
      Begin VB.Label lblGrupo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   9450
         TabIndex        =   39
         Top             =   90
         Width           =   630
      End
      Begin VB.Label lblGrupo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   8790
         TabIndex        =   38
         Top             =   90
         Width           =   330
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1245
         TabIndex        =   37
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Prd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   36
         Top             =   90
         Width           =   750
      End
      Begin VB.Shape Shape11 
         Height          =   285
         Left            =   120
         Top             =   45
         Width           =   11235
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   105
      Index           =   0
      Left            =   0
      Top             =   10230
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   185
      Tipo            =   5
      Mostrar         =   0   'False
      Ordem           =   2
   End
   Begin VB.Shape Shape8 
      Height          =   1245
      Left            =   0
      Top             =   0
      Width           =   10965
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   0
      X2              =   67.204
      Y1              =   -4.763
      Y2              =   -4.763
   End
End
Attribute VB_Name = "FrmRelPed3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalItens As Currency
Dim SubTotal As Currency
Dim TotalIpi As Currency

Public Sub Config()
    Relatorio.Ativar
    Unload Me
End Sub

Private Sub Form_Load()
    If Dir(Caminho & "\Logo.jpg") <> "" Then
        Logotipo = LoadPicture(Caminho & "\Logo.jpg")
        'rpCabTitulo.Alignment = vbLeftJustify
        'rpCabInformacao.Alignment = vbLeftJustify
        rpCabTitulo.Mostrar = False
        rpCabInformacao.Mostrar = False
    Else
        rpCabTitulo.Left = 990
        rpCabInformacao.Left = 990
    End If
      
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    With shpZebra
        .Left = 0
        .Top = 0
        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
        .Height = Det.Height
    End With
    
    TotalItens = 0
    SubTotal = 0
    TotalIpi = 0
        
End Sub


Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
'        shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
'        shpZebra.Visible = False
    End If
    
    
       
End Sub

Private Sub Rpx_MsgErro(Numero As Long)

    Dim Msg$
    
    If Numero < 0 Then
    
        ' Mensagens de erro previstas
        Select Case Numero - vbObjectError
            Case 1001: Msg = "É necessário existir uma impressora instalada no Windows"
            Case 1002: Msg = "Não há registros a imprimir"
            Case 1003: Msg = "Não foi definida a seção de detalhe do relatório"
            Case 1004: Msg = "A configuração das seções de grupos está incorreta"
            Case 1005: Msg = "Foi definido um cursor do tipo Forward-Only para o recordset do relatório."
            Case 1006: Msg = "A página configurada para o relatório não possuí espaço suficiente para a impressão"
            Case 1007: Msg = "Já existe um relatório em andamento"
        End Select
        
        MsgBox Msg, vbInformation, "Impressão"
        
    Else
        
        MsgBox "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impressão"
        
    End If
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        Case "forEntregue"
            Valor = IIf(Relatorio.Recordset("Entregue") = True, "Sim", "Não")
        Case "forUnitario"
            If VarValorPedidoVista = False Then
                Valor = Relatorio.Recordset("VlrVdaDia") / Relatorio.Recordset("Qtd")
            Else
                Valor = Relatorio.Recordset("Valor") / Relatorio.Recordset("Qtd")
            End If
        Case "forVlrVdaDia"
            If VarValorPedidoVista = False Then
                Valor = Relatorio.Recordset("VlrVdaDia")
            Else
                Valor = Relatorio.Recordset("Valor")
            End If
        Case "forTotalItens"
            Valor = TotalItens
        Case "forSubTotal"
            Valor = SubTotal
        Case "forTotalPedido"
            Valor = (SubTotal + Relatorio.Recordset("ValorFrete") + TotalIpi + Relatorio.Recordset("PedSub")) - Relatorio.Recordset("Desconto")
'            TotalItens = 0
            SubTotal = 0
            
        Case "forIpi"
            Valor = TotalIpi
        Case "forAcrescimo"
            Valor = Relatorio.Recordset("ValorFrete") + TotalIpi + Relatorio.Recordset("PedSub")
            
        
    End Select
        
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    ' Indica que o grupo irá quebrar a cada novo codigo
    ' do pedido. Verifica se é formula para o grupo 1 ou
    ' para o grupo 2. Nesse caso o grupo 2 nao tem quebra
    ' apenas o grupo 1.
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("CodPed")
    End If
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)

    TotalItens = TotalItens + Relatorio.Recordset("Qtd")
    If VarValorPedidoVista = False Then
        SubTotal = SubTotal + Relatorio.Recordset("VlrVdaDia")
    Else
        SubTotal = SubTotal + Relatorio.Recordset("Valor")
    End If
    If VarSimples = False Then
'        TotalIpi = TotalIpi + ((Relatorio.Recordset("VlrVdaDia") / 100) * Relatorio.Recordset("Ipi"))
    End If
    TotalIpi = Relatorio.Recordset("Ped.Ipi")
    
'                             Variavel(6) = ((Valor / 100) * TabSql("Prd.Ipi")) ' valor do ipi
'                         Variavel(7) = Variavel(7) + Variavel(6) ' total do ipi

'        If Len(Trim(Relatorio.Recordset("Itp.Complemento"))) <= 15 Then
'            Me.Det.Height = 4
'            Me.Line2.y1 = 190
'            Me.Line2.y2 = 190
'            Me.ReportField16.Mostrar = False
'        Else
'            Me.Det.Height = 7
'            Me.Line2.y1 = 400
'            Me.Line2.y2 = 400
'            Me.ReportField15.Mostrar = False
'            Me.ReportField16.Mostrar = True
'        End If
    
    
    
            
End Sub
