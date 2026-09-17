VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelVdaFam 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   70.115
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   210
      TabIndex        =   0
      Top             =   3330
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   1245
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2196
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   582
         Caption         =   ""
         TipoCampo       =   7
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
      Begin ReportX.ReportField rpCabSubTitulo 
         Height          =   270
         Left            =   6300
         TabIndex        =   9
         Top             =   90
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   476
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   1
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField rpCabPeriodo 
         Height          =   225
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   397
         Caption         =   ""
         TipoCampo       =   7
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   660
         TabIndex        =   18
         Top             =   690
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   370
         Caption         =   "Descrição da Familia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   210
         Left            =   90
         TabIndex        =   19
         Top             =   690
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         Caption         =   "Cód."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   180
         TabIndex        =   20
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   370
         Caption         =   "Cód."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   750
         TabIndex        =   21
         Top             =   900
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   370
         Caption         =   "Descrição do Produto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   210
         Left            =   3810
         TabIndex        =   22
         Top             =   900
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Caption         =   "Quantidade"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   210
         Left            =   4770
         TabIndex        =   23
         Top             =   900
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Valor Unit."
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   210
         Left            =   5835
         TabIndex        =   24
         Top             =   900
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Valor"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   210
         Left            =   7905
         TabIndex        =   25
         Top             =   900
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Custo Dia"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField20 
         Height          =   210
         Left            =   6840
         TabIndex        =   26
         Top             =   900
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Vlr.Vda.Dia"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField21 
         Height          =   210
         Left            =   8955
         TabIndex        =   27
         Top             =   900
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Lucro"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   10005
         TabIndex        =   28
         Top             =   900
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Finan."
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField38 
         Height          =   210
         Left            =   10710
         TabIndex        =   44
         Top             =   900
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Est."
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   1140
         Y2              =   1140
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   240
      Index           =   2
      Left            =   0
      Top             =   1680
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   180
         TabIndex        =   4
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   370
         Campo           =   "CodPrd"
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
         Left            =   750
         TabIndex        =   5
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
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
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   3810
         TabIndex        =   10
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Campo           =   "SomaDeQtd"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   4770
         TabIndex        =   11
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   370
         Campo           =   "MédiaDeValor"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   5790
         TabIndex        =   12
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "SomaDeValor"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   7950
         TabIndex        =   13
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "Custo"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField10 
         Height          =   210
         Left            =   6840
         TabIndex        =   14
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "SomaDeVlrVdaDia"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   9045
         TabIndex        =   15
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "Lucro"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   10095
         TabIndex        =   16
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Campo           =   "forFinanceiro"
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
      Begin ReportX.ReportField ReportField37 
         Height          =   210
         Left            =   10755
         TabIndex        =   43
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Campo           =   "Estoque"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   10095
         Top             =   30
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
         TabIndex        =   1
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      Top             =   2610
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   953
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8520
         TabIndex        =   2
         Top             =   60
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
         Left            =   255
         TabIndex        =   3
         Top             =   60
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
      Begin VB.Line lneRod 
         X1              =   10980
         X2              =   240
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   435
      Index           =   0
      Left            =   0
      Top             =   1245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   630
         TabIndex        =   6
         Top             =   60
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   370
         Campo           =   "Descrição_Familia"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         Campo           =   "CodGru"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   30
         Y2              =   30
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   315
      Index           =   0
      Left            =   0
      Top             =   1920
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField23 
         Height          =   210
         Left            =   3810
         TabIndex        =   29
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forSomaDeQtd"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField24 
         Height          =   210
         Left            =   5700
         TabIndex        =   30
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forSomaDeValor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField25 
         Height          =   210
         Left            =   7950
         TabIndex        =   31
         Top             =   30
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "forCusto"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField26 
         Height          =   210
         Left            =   6840
         TabIndex        =   32
         Top             =   30
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "forSomaDeVlrVdaDia"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   210
         Left            =   9045
         TabIndex        =   33
         Top             =   30
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "forLucro"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField28 
         Height          =   210
         Left            =   10095
         TabIndex        =   34
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Campo           =   "forSomaFinanceiro"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField29 
         Height          =   210
         Left            =   2250
         TabIndex        =   35
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   370
         Caption         =   "Totais do Familia"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   11400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   2235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField ReportField30 
         Height          =   210
         Left            =   3810
         TabIndex        =   36
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "totSomaDeQtd"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField31 
         Height          =   210
         Left            =   5670
         TabIndex        =   37
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   370
         Campo           =   "totSomaDeValor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   210
         Left            =   7950
         TabIndex        =   38
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "totCusto"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField33 
         Height          =   210
         Left            =   6840
         TabIndex        =   39
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Campo           =   "totSomaDeVlrVdaDia"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField34 
         Height          =   210
         Left            =   9045
         TabIndex        =   40
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "totLucro"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField35 
         Height          =   210
         Left            =   10095
         TabIndex        =   41
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Campo           =   "totSomaFinanceiro"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField36 
         Height          =   210
         Left            =   2250
         TabIndex        =   42
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   370
         Caption         =   "Totais Geral"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
   End
End
Attribute VB_Name = "FrmRelVdaFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SomaQtd As Currency
Dim SomaValor As Currency
Dim SomaVlrVdaDia As Currency
Dim SomaCusto As Currency
Dim SomaLucro As Currency
Dim SomaFinanceiro As Currency

Dim totQtd As Currency
Dim totValor As Currency
Dim totVlrVdaDia As Currency
Dim TotCusto As Currency
Dim totLucro As Currency
Dim totFinanceiro As Currency

Public Sub Config()

    Relatorio.Ativar
    
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    With shpZebra
    End With
    
    totQtd = 0
    totValor = 0
    totVlrVdaDia = 0
    TotCusto = 0
    totLucro = 0
    totFinanceiro = 0
    
    SomaQtd = 0
    SomaValor = 0
    SomaVlrVdaDia = 0
    SomaCusto = 0
    SomaLucro = 0
    SomaFinanceiro = 0
    
    
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        'shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
        shpZebra.Visible = False
    End If
       
       
End Sub

Private Sub Rpx_MsgErro(Numero As Long)

    Dim Msg$
    
    If Numero < 0 Then
    
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
        Case "forFinanceiro"
            If Relatorio.Recordset("SomaDeVlrVdaDia") - Relatorio.Recordset("SomaDeValor") > 0 Then
                Valor = Relatorio.Recordset("SomaDeVlrVdaDia") - Relatorio.Recordset("SomaDeValor")
                SomaFinanceiro = SomaFinanceiro + Valor
            Else
                Valor = 0
            End If
        Case "forSomaDeQtd"
            Valor = SomaQtd
                SomaQtd = 0
        Case "forSomaDeValor"
            Valor = SomaValor
                SomaValor = 0
        Case "forSomaDeVlrVdaDia"
            Valor = SomaVlrVdaDia
                SomaVlrVdaDia = 0
        Case "forCusto"
            Valor = SomaCusto
                SomaCusto = 0
        Case "forLucro"
            Valor = SomaLucro
                SomaLucro = 0
        Case "forSomaFinanceiro"
            Valor = SomaFinanceiro
                SomaFinanceiro = 0
                
        
        Case "totSomaDeQtd"
            Valor = totQtd
        Case "totSomaDeValor"
            Valor = totValor
        Case "totSomaDeVlrVdaDia"
            Valor = totVlrVdaDia
        Case "totCusto"
            Valor = TotCusto
        Case "totLucro"
            Valor = totLucro
        Case "totSomaFinanceiro"
            Valor = totFinanceiro
    
    End Select
End Sub


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("CodFam")
    End If
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
        
    SomaQtd = SomaQtd + Relatorio.Recordset("SomaDeQtd")
    SomaValor = SomaValor + Relatorio.Recordset("SomaDeValor")
    SomaVlrVdaDia = SomaVlrVdaDia + Relatorio.Recordset("SomaDeVlrVdaDia")
    SomaCusto = SomaCusto + Relatorio.Recordset("Custo")
    SomaLucro = SomaLucro + Relatorio.Recordset("Lucro")
    
    totQtd = totQtd + Relatorio.Recordset("SomaDeQtd")
    totValor = totValor + Relatorio.Recordset("SomaDeValor")
    totVlrVdaDia = totVlrVdaDia + Relatorio.Recordset("SomaDeVlrVdaDia")
    TotCusto = TotCusto + Relatorio.Recordset("Custo")
    totLucro = totLucro + Relatorio.Recordset("Lucro")
    
End Sub
