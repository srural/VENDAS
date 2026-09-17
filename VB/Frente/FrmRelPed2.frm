VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPed2 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   119.856
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      MargemSuperior  =   5
      Titulo          =   ""
      Copias          =   2
      LarguraPapel    =   210
      AlturaPapel     =   150
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   2265
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   3995
      Tipo            =   1
      Ordem           =   1
      Begin ReportX.ReportField rpCabInformacao 
         Height          =   240
         Left            =   165
         TabIndex        =   3
         Top             =   360
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   423
         Caption         =   ""
         TipoCampo       =   7
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField10 
         Height          =   330
         Left            =   8865
         TabIndex        =   4
         Top             =   660
         Width           =   1950
         _ExtentX        =   3440
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
         Height          =   285
         Left            =   165
         TabIndex        =   1
         Top             =   60
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   503
         Caption         =   ""
         TipoCampo       =   7
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField23 
         Height          =   225
         Left            =   1185
         TabIndex        =   44
         Top             =   1035
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
         TabIndex        =   45
         Top             =   1710
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
         TabIndex        =   46
         Top             =   1680
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
         TabIndex        =   47
         Top             =   1935
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   3000
         TabIndex        =   48
         Top             =   1935
         Width           =   1350
         _ExtentX        =   2381
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
         TabIndex        =   49
         Top             =   1710
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
         Left            =   7215
         TabIndex        =   50
         Top             =   1935
         Width           =   1680
         _ExtentX        =   2963
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
         Left            =   9645
         TabIndex        =   51
         Top             =   1935
         Width           =   1575
         _ExtentX        =   2778
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
         TabIndex        =   52
         Top             =   975
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
         TabIndex        =   53
         Top             =   1275
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
         TabIndex        =   54
         Top             =   1485
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
         TabIndex        =   55
         Top             =   1290
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
         TabIndex        =   56
         Top             =   1515
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
         TabIndex        =   57
         Top             =   1035
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
         TabIndex        =   58
         Top             =   1230
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
         Left            =   1170
         TabIndex        =   59
         Top             =   1260
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   397
         Campo           =   "Fantasia"
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
         TabIndex        =   60
         Top             =   1710
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
         Left            =   4935
         TabIndex        =   61
         Top             =   1935
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   3195
         TabIndex        =   94
         Top             =   675
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
         Left            =   135
         TabIndex        =   95
         Top             =   675
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
      Begin VB.Line Line2 
         X1              =   90
         X2              =   11340
         Y1              =   2205
         Y2              =   2205
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
         TabIndex        =   79
         Top             =   1035
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
         TabIndex        =   78
         Top             =   1500
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
         TabIndex        =   77
         Top             =   1290
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
         TabIndex        =   76
         Top             =   1680
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
         Left            =   8895
         TabIndex        =   75
         Top             =   1935
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
         TabIndex        =   74
         Top             =   1695
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
         Left            =   2535
         TabIndex        =   73
         Top             =   1920
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
         Height          =   270
         Left            =   9465
         TabIndex        =   72
         Top             =   1275
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
         TabIndex        =   71
         Top             =   1005
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
         Left            =   6435
         TabIndex        =   70
         Top             =   1890
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
         TabIndex        =   69
         Top             =   1920
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
         TabIndex        =   68
         Top             =   1680
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
         TabIndex        =   67
         Top             =   1470
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
         TabIndex        =   66
         Top             =   1035
         Width           =   540
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
         TabIndex        =   65
         Top             =   1005
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
         TabIndex        =   64
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label12 
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
         Left            =   270
         TabIndex        =   63
         Top             =   1260
         Width           =   765
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
         Left            =   4455
         TabIndex        =   62
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image Logotipo 
         Height          =   600
         Left            =   90
         Stretch         =   -1  'True
         Top             =   30
         Width           =   11235
      End
      Begin VB.Line Line1 
         X1              =   10830
         X2              =   11280
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   8820
         Y1              =   945
         Y2              =   945
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   3000
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   165
         TabIndex        =   5
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   370
         Campo           =   "Produto"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   1260
         TabIndex        =   6
         Top             =   0
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   370
         Campo           =   "Descrição_Produto"
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
      Begin ReportX.ReportField rpDetalhes 
         Height          =   210
         Index           =   0
         Left            =   7470
         TabIndex        =   7
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   370
         Campo           =   "forEntregue"
         Caption         =   "Nome"
         Formula         =   -1  'True
         Mostrar         =   0   'False
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
      Begin ReportX.ReportField rpDetalhes 
         Height          =   210
         Index           =   2
         Left            =   9225
         TabIndex        =   8
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   370
         Campo           =   "forUnitario"
         Formato         =   "##,##0.0000"
         Caption         =   "Nome"
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
      Begin ReportX.ReportField rpDetalhes 
         Height          =   210
         Index           =   1
         Left            =   8205
         TabIndex        =   9
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   370
         Campo           =   "Qtd"
         Caption         =   "Nome"
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
      Begin ReportX.ReportField rpDetalhes 
         Height          =   210
         Index           =   3
         Left            =   10170
         TabIndex        =   10
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "VlrVdaDia"
         Formato         =   "Standard"
         Caption         =   "Nome"
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   5040
         TabIndex        =   20
         Top             =   0
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   370
         Campo           =   "Itp.Complemento"
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
      Begin ReportX.ReportField ReportField17 
         Height          =   210
         Left            =   7815
         TabIndex        =   23
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   370
         Campo           =   "Embalagem"
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
      Begin VB.Line Line4 
         X1              =   135
         X2              =   11340
         Y1              =   225
         Y2              =   225
      End
      Begin VB.Line Line16 
         X1              =   120
         X2              =   120
         Y1              =   -30
         Y2              =   270
      End
      Begin VB.Line Line15 
         X1              =   11340
         X2              =   11340
         Y1              =   -30
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
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   435
      Index           =   1
      Left            =   0
      Top             =   2265
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   7410
         TabIndex        =   31
         Top             =   225
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
         TabIndex        =   32
         Top             =   225
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
         TabIndex        =   33
         Top             =   225
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
         TabIndex        =   34
         Top             =   225
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
         TabIndex        =   35
         Top             =   225
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
      Begin ReportX.ReportField ReportField21 
         Height          =   210
         Left            =   10320
         TabIndex        =   42
         Top             =   240
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
         TabIndex        =   43
         Top             =   60
         Width           =   330
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
         TabIndex        =   40
         Top             =   0
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
         TabIndex        =   39
         Top             =   45
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
         TabIndex        =   38
         Top             =   45
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
         TabIndex        =   37
         Top             =   30
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
         TabIndex        =   36
         Top             =   45
         Width           =   525
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   2085
      Index           =   1
      Left            =   0
      Top             =   3270
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   3678
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   0
         Left            =   9585
         TabIndex        =   14
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         Campo           =   "forSubTotal"
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
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   1
         Left            =   7110
         TabIndex        =   15
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
         Height          =   240
         Index           =   2
         Left            =   9585
         TabIndex        =   16
         Top             =   735
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         Campo           =   "forTotalPedido"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
      Begin ReportX.ReportField Obs3 
         Height          =   210
         Left            =   255
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         Height          =   225
         Index           =   3
         Left            =   9585
         TabIndex        =   21
         Top             =   240
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         Campo           =   "Desconto"
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
      Begin ReportX.ReportField Obs4 
         Height          =   210
         Left            =   255
         TabIndex        =   41
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
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   6420
         TabIndex        =   80
         Top             =   900
         Width           =   915
         _ExtentX        =   1614
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
      Begin ReportX.ReportField Vendedor 
         Height          =   225
         Left            =   1230
         TabIndex        =   81
         Top             =   885
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
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
         TabIndex        =   84
         Top             =   240
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
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   5
         Left            =   9585
         TabIndex        =   87
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         Campo           =   "forAcrescimo"
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
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   6
         Left            =   7110
         TabIndex        =   89
         Top             =   450
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
      Begin VB.Line Line30 
         Visible         =   0   'False
         X1              =   2865
         X2              =   6855
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor/Entregador"
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
         Left            =   4050
         TabIndex        =   93
         Top             =   1800
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAVOR CONFERIR A MERCADORIA NO ATO DA ENTREGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   92
         Top             =   1215
         Visible         =   0   'False
         Width           =   6540
      End
      Begin VB.Line Line3 
         Visible         =   0   'False
         X1              =   7335
         X2              =   11175
         Y1              =   1830
         Y2              =   1830
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
         TabIndex        =   91
         Top             =   1620
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assinatura"
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
         Left            =   8910
         TabIndex        =   90
         Top             =   1815
         Visible         =   0   'False
         Width           =   885
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
         TabIndex        =   88
         Top             =   450
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
         TabIndex        =   86
         Top             =   450
         Width           =   585
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
         TabIndex        =   85
         Top             =   240
         Width           =   375
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
         Left            =   270
         TabIndex        =   83
         Top             =   885
         Visible         =   0   'False
         Width           =   870
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
         Left            =   5445
         TabIndex        =   82
         Top             =   900
         Width           =   930
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   120
         X2              =   11340
         Y1              =   -45
         Y2              =   -45
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
         TabIndex        =   22
         Top             =   240
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   -15
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
         TabIndex        =   17
         Top             =   705
         Width           =   600
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   300
      Index           =   3
      Left            =   0
      Top             =   2700
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   3
      Ordem           =   2
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
         Left            =   7725
         TabIndex        =   30
         Top             =   45
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
         Left            =   5055
         TabIndex        =   29
         Top             =   45
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
         TabIndex        =   28
         Top             =   45
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
         TabIndex        =   27
         Top             =   45
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
         Left            =   8880
         TabIndex        =   26
         Top             =   45
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
         TabIndex        =   25
         Top             =   45
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
         TabIndex        =   24
         Top             =   45
         Width           =   750
      End
      Begin VB.Shape Shape11 
         Height          =   240
         Left            =   120
         Top             =   45
         Width           =   11235
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   15
      Index           =   0
      Left            =   0
      Top             =   13875
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   -26
      Tipo            =   5
      Mostrar         =   0   'False
      Ordem           =   2
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   8520
      Left            =   0
      Top             =   5355
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   15028
      Tipo            =   7
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
Attribute VB_Name = "FrmRelPed2"
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
    If dir(LocalBanco & "\Logo.jpg") <> "" Then
        Logotipo = LoadPicture(LocalBanco & "\Logo.jpg")
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
            Valor = Relatorio.Recordset("VlrVdaDia") / Relatorio.Recordset("Qtd")
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
    SubTotal = SubTotal + Relatorio.Recordset("VlrVdaDia")
'    If VarSimples = False Then
'        TotalIpi = TotalIpi + ((Relatorio.Recordset("VlrVdaDia") / 100) * Relatorio.Recordset("Ipi"))
'    End If
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
