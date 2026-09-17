VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelVda 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   51.065
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   206.904
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   300
      TabIndex        =   0
      Top             =   2310
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1879
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   30
         Width           =   6885
         _ExtentX        =   12144
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
         TabIndex        =   5
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
         TabIndex        =   6
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
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Caption         =   "Código"
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
         Left            =   1140
         TabIndex        =   13
         Top             =   720
         Width           =   2745
         _ExtentX        =   4842
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
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   6030
         TabIndex        =   14
         Top             =   720
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Caption         =   "Qtd"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   9045
         TabIndex        =   15
         Top             =   720
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Caption         =   "Total Vda"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   210
         Left            =   6585
         TabIndex        =   16
         Top             =   720
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Caption         =   "Méd Vda"
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
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   3975
         TabIndex        =   20
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   370
         Caption         =   "Média Custo"
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
      Begin ReportX.ReportField ReportField13 
         Height          =   210
         Left            =   5100
         TabIndex        =   21
         Top             =   720
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   370
         Caption         =   "Méd(%)"
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
      Begin ReportX.ReportField ReportField23 
         Height          =   210
         Left            =   8145
         TabIndex        =   29
         Top             =   720
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         Caption         =   "Tot Custo"
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
         Left            =   10110
         TabIndex        =   30
         Top             =   720
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   370
         Caption         =   "Lucro"
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
         Left            =   10935
         TabIndex        =   34
         Top             =   720
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   370
         Caption         =   "Estoq"
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
      Begin ReportX.ReportField ReportField30 
         Height          =   210
         Left            =   7425
         TabIndex        =   36
         Top             =   720
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   370
         Caption         =   "Vda Tab"
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
         Y1              =   960
         Y2              =   960
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   210
      Left            =   0
      Top             =   1065
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   370
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "Produto"
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
         Left            =   1140
         TabIndex        =   8
         Top             =   0
         Width           =   3105
         _ExtentX        =   5477
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
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   5820
         TabIndex        =   9
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
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
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   9060
         TabIndex        =   10
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
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
         Left            =   6630
         TabIndex        =   11
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   370
         Campo           =   "MédiaDeValorUnit"
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
         Left            =   4290
         TabIndex        =   19
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Campo           =   "MédiaDeCusto"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   5175
         TabIndex        =   22
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Campo           =   "forPerc"
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
      Begin ReportX.ReportField ReportField21 
         Height          =   210
         Left            =   8190
         TabIndex        =   27
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "SomaDeCusto"
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
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   10020
         TabIndex        =   28
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   210
         Left            =   10980
         TabIndex        =   33
         Top             =   0
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   370
         Campo           =   "Estoque"
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
      Begin ReportX.ReportField ReportField29 
         Height          =   210
         Left            =   7425
         TabIndex        =   35
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   370
         Campo           =   "MediaVendaDia"
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
         Left            =   10155
         Top             =   30
         Width           =   660
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
      Height          =   450
      Left            =   0
      Top             =   1650
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   794
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8865
         TabIndex        =   2
         Top             =   45
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
         X1              =   11370
         X2              =   90
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1275
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   661
      Tipo            =   6
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   2970
         TabIndex        =   17
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   370
         Caption         =   "Total"
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
         Left            =   9060
         TabIndex        =   18
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   370
         Campo           =   "forValor"
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
      Begin ReportX.ReportField ReportField16 
         Height          =   210
         Left            =   6630
         TabIndex        =   23
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   370
         Campo           =   "forVenda"
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
      Begin ReportX.ReportField ReportField18 
         Height          =   210
         Left            =   5820
         TabIndex        =   24
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   370
         Campo           =   "forQtd"
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
      Begin ReportX.ReportField ReportField19 
         Height          =   210
         Left            =   5175
         TabIndex        =   25
         Top             =   60
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Campo           =   "forMedPerc"
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
      Begin ReportX.ReportField ReportField20 
         Height          =   210
         Left            =   4290
         TabIndex        =   26
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
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
      Begin ReportX.ReportField ReportField25 
         Height          =   210
         Left            =   8220
         TabIndex        =   31
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Campo           =   "forTotCusto"
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
         Left            =   10020
         TabIndex        =   32
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Campo           =   "forTotLucro"
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
      Begin ReportX.ReportField ReportField31 
         Height          =   210
         Left            =   7425
         TabIndex        =   37
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   370
         Campo           =   "forVendaTab"
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
      Begin VB.Line Line1 
         X1              =   11430
         X2              =   90
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "FrmRelVda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SomaValor As Currency
Dim VarPerc As Currency
Dim MediaVenda As Currency
Dim MediaCusto As Currency
Dim MediaPerc As Currency
Dim TotItens As Currency
Dim SomaVenda As Currency
Dim SomaQtd As Currency
Dim SomaCusto As Currency
Dim VarReg As Currency
Dim TotCusto As Currency
Dim MediaTab As Currency


Public Sub Config()

    Relatorio.Ativar
    
    Unload Me
    
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

    SomaValor = 0
    VarPerc = 0
    SomaVenda = 0
    SomaQtd = 0
    SomaCusto = 0
    VarReg = 0
    MediaPerc = 0
    TotCusto = 0
    MediaTab = 0
    
    
End Sub



Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
        shpZebra.Visible = False
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
        Case "forValor"
            Valor = SomaValor
        Case "forPerc"
            If Relatorio.Recordset("médiadecusto") <> 0 And Relatorio.Recordset("SomaDeQtd") <> 0 Then
                VarPerc = ((Relatorio.Recordset("médiadevalorunit") / Relatorio.Recordset("médiadecusto")) - 1) * 100
            Else
                VarPerc = 0
            End If
            Valor = VarPerc
            MediaPerc = MediaPerc + VarPerc
        Case "forVenda"
            Valor = SomaVenda / VarReg
        Case "forQtd"
            Valor = SomaQtd
        Case "forCusto"
            Valor = SomaCusto / VarReg
        Case "forMedPerc"
'            Valor = ((SomaVenda / SomaCusto) - 1) * 100
             Valor = MediaPerc / VarReg
        Case "forLucro"
            Valor = Relatorio.Recordset("SomaDeValor") - Relatorio.Recordset("SomaDeCusto")
        Case "forTotCusto"
            Valor = TotCusto
        Case "forTotLucro"
            Valor = SomaValor - TotCusto
        Case "forVendaTab"
            Valor = MediaTab / VarReg
    End Select
End Sub


Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
            
    SomaValor = SomaValor + Relatorio.Recordset("SomaDeValor")
    TotCusto = TotCusto + Relatorio.Recordset("SomaDeCusto")
    If Relatorio.Recordset("SomaDeQtd") <> 0 Then
        MediaVenda = MediaVenda + Relatorio.Recordset("MédiaDeValorUnit")
        SomaVenda = SomaVenda + Relatorio.Recordset("MédiaDeValorUnit")
'        MediaTab = MediaTab + Relatorio.Recordset("MediaVendaDia")
    End If
    MediaCusto = MediaCusto + Relatorio.Recordset("MédiaDeCusto")
    SomaQtd = SomaQtd + Relatorio.Recordset("SomaDeQtd")
    SomaCusto = SomaCusto + Relatorio.Recordset("MédiaDeCusto")
    VarReg = VarReg + 1
    
End Sub

