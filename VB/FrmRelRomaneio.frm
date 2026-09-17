VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelRomaneio 
   ClientHeight    =   6885
   ClientLeft      =   1950
   ClientTop       =   2100
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11880
   Begin ReportX.ReportSection RodTmp 
      Align           =   1  'Align Top
      Height          =   270
      Index           =   0
      Left            =   0
      Top             =   3885
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   0
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   23
         Left            =   6345
         TabIndex        =   20
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   24
         Left            =   4095
         TabIndex        =   21
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Total dos Pedidos"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
   End
   Begin ReportX.ReportSection CabTmp 
      Align           =   1  'Align Top
      Height          =   885
      Index           =   0
      Left            =   0
      Top             =   2730
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1561
      Tipo            =   0
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   12
         Top             =   570
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   397
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   7
         Left            =   5445
         TabIndex        =   13
         Top             =   570
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Caption         =   "Nota"
         Alignment       =   1
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   8
         Left            =   6285
         TabIndex        =   14
         Top             =   570
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Total"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   13
         Left            =   10245
         TabIndex        =   18
         Top             =   570
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Forma de Pgto"
         TipoCampo       =   1
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   12
         Left            =   7740
         TabIndex        =   22
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Data"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   14
         Left            =   8865
         TabIndex        =   23
         Top             =   570
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Pedido"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin VB.Line Line4 
         X1              =   6225
         X2              =   6225
         Y1              =   780
         Y2              =   540
      End
      Begin VB.Line Line3 
         X1              =   5385
         X2              =   5385
         Y1              =   780
         Y2              =   540
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   60
         Top             =   540
         Width           =   11775
      End
   End
   Begin ReportX.ReportSection RodPri 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Top             =   1950
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      Tipo            =   5
      Mostrar         =   0   'False
      Ordem           =   1
   End
   Begin ReportX.ReportSection CabPri 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   675
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
      Tipo            =   3
      Mostrar         =   0   'False
      Ordem           =   1
   End
   Begin ReportX.ReportSection RodGrp 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   1665
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   5
      Mostrar         =   0   'False
      Ordem           =   2
   End
   Begin ReportX.ReportSection CabGrp 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1125
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   3
      Mostrar         =   0   'False
      Ordem           =   2
   End
   Begin ReportX.ReportSection DetTmp 
      Align           =   1  'Align Top
      Height          =   270
      Index           =   0
      Left            =   0
      Top             =   3615
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   0
      Begin ReportX.ReportField rpfDet 
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   370
         Campo           =   "00.Nome"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   1
         Left            =   5490
         TabIndex        =   2
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "00.NroNt"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   2
         Left            =   6345
         TabIndex        =   10
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "00.Total"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   210
         Index           =   22
         Left            =   3330
         TabIndex        =   19
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   370
         Campo           =   "00.cidade"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   15
         Left            =   7740
         TabIndex        =   24
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   397
         Campo           =   "00.DataEmiss"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   16
         Left            =   8730
         TabIndex        =   25
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "00.CodPed"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin VB.Line Line24 
         X1              =   8685
         X2              =   8685
         Y1              =   225
         Y2              =   -45
      End
      Begin VB.Line Line22 
         X1              =   9945
         X2              =   9945
         Y1              =   240
         Y2              =   -30
      End
      Begin VB.Line Line21 
         X1              =   7665
         X2              =   7665
         Y1              =   240
         Y2              =   -30
      End
      Begin VB.Line Line2 
         X1              =   6255
         X2              =   6255
         Y1              =   225
         Y2              =   -45
      End
      Begin VB.Line Line1 
         X1              =   5400
         X2              =   5400
         Y1              =   225
         Y2              =   -45
      End
      Begin VB.Shape Shape2 
         Height          =   285
         Left            =   60
         Top             =   -30
         Width           =   11790
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1395
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   2370
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   7530
         TabIndex        =   5
         Top             =   75
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Campo           =   "=Página [Pagina] de [Paginas]"
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
         Left            =   225
         TabIndex        =   6
         Top             =   75
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
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   165
      TabIndex        =   0
      Top             =   6240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   1200
      Escala          =   1
      Titulo          =   "Relatório Múltiplo"
      HTMLUnico       =   -1  'True
   End
   Begin ReportX.ReportSection DetTmp 
      Align           =   1  'Align Top
      Height          =   270
      Index           =   1
      Left            =   0
      Top             =   4740
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   0
      Begin ReportX.ReportField rpfDet 
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   3
         Top             =   0
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   370
         Campo           =   "01.Descrição_Produto"
         Formato         =   "#,##0.00"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   4
         Left            =   6810
         TabIndex        =   4
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "01.SomaDeQtd"
         Formato         =   "0.00"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   5
         Left            =   7650
         TabIndex        =   11
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         Campo           =   "01.SomaDeVlrVdaDia"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin VB.Shape Shape4 
         Height          =   285
         Left            =   60
         Top             =   -30
         Width           =   8835
      End
      Begin VB.Line Line8 
         X1              =   6750
         X2              =   6750
         Y1              =   240
         Y2              =   -30
      End
      Begin VB.Line Line7 
         X1              =   7590
         X2              =   7590
         Y1              =   240
         Y2              =   -30
      End
   End
   Begin ReportX.ReportSection CabTmp 
      Align           =   1  'Align Top
      Height          =   585
      Index           =   1
      Left            =   0
      Top             =   4155
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1032
      Tipo            =   0
      Ordem           =   1
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   9
         Left            =   90
         TabIndex        =   15
         Top             =   270
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   397
         Caption         =   "Produto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   10
         Left            =   6810
         TabIndex        =   16
         Top             =   270
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Caption         =   "Qtd."
         Alignment       =   1
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   11
         Left            =   7620
         TabIndex        =   17
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Formato         =   "Standard"
         Caption         =   "Total"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin VB.Line Line6 
         X1              =   6750
         X2              =   6750
         Y1              =   480
         Y2              =   240
      End
      Begin VB.Line Line5 
         X1              =   7590
         X2              =   7590
         Y1              =   480
         Y2              =   240
      End
      Begin VB.Shape Shape3 
         Height          =   255
         Left            =   60
         Top             =   240
         Width           =   8835
      End
   End
   Begin ReportX.ReportSection RodTmp 
      Align           =   1  'Align Top
      Height          =   270
      Index           =   1
      Left            =   0
      Top             =   5010
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   0
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1191
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
   End
End
Attribute VB_Name = "FrmRelRomaneio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GrupoAtual As Integer
Private SqlPrincipal As String

Private rds As clsRecordset
Private RSSec() As Recordset
Private strTotalGeral As Currency
Private strTotalVdaDia As Currency
Public VarSomaTotal As Currency

Private Function FiltroString(ByVal Valor As String) As String
    
    FiltroString = "'" & Valor & "'"
    
End Function

Private Function TotalRegistros(Rs As Recordset)
    
    ' Essa função server para contar o numero de registros
    ' do recordset para que seja colocada na propriedade
    ' recordcount
    If Not Rs.Eof Then
        Rs.MoveLast
        TotalRegistros = Rs.RecordCount
        Rs.MoveFirst
    End If
    
End Function

Public Function Config()

    Dim i As Integer
    Dim RSPri As Recordset
    Dim Cn As String
    Dim Num_RS As Integer

    PreparaControles
        
    strTotalGeral = 0
    strTotalVdaDia = 0
    'cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\..\Exemplo.MDB"
    
    Set rds = New clsRecordset
    'Set RSPri = New Recordset
    
    ' O relatorio pode ser baseado em um recordset principal
    ' que controlará as secoes de Grupo CabPri e RodPri
    ' que farão a quebra, se necessário
    ' Foi colocado aqui uma expressão simples para gerar
    ' um recordset de 1 linha apenas para o relatorio possa
    ' ser gerado para 1 grupo. Se existir um recordset principa
    ' pode-se colocar o SQL para ele aqui.
    
    'RSPri.Open "SELECT 1", cn, adOpenStatic
    SqlPrincipal = "Select CodEntidade from Ent"
    
    Set RSPri = Banco.OpenRecordset(SqlPrincipal)
    Set rds.RSPri = RSPri
    
    ' Aqui abre a matriz e cria todos os objetos recordset
    ' vai ter um RS para cada conjunto CabTmp, DetTmp, RodTmp
    Num_RS = DetTmp.UBound
    ReDim RSSec(Num_RS)
    For i = 0 To Num_RS
        'Set RSSec(i) = New Recordset
        'Set RSSec(i) = RSPri.Clone
        Set RSSec(i) = Banco.OpenRecordset(SqlDiv(i + 1))
        rds.AddSec RSSec(i)
    Next
    
    Set Relatorio.Recordset = rds
    Relatorio.Ativar
    
    ' Retira o objeto da memoria que fará com
    ' que todos os RS sejam fechados
    Set rds = Nothing
    Unload Me
    
End Function

Private Sub PreparaControles()

    ' Essa sub irá copiar todos os controles que estão nas seções "indefinido"
    ' para as seções de cabecalho de grupo, detalhe e rodapé.
    ' Isso é necessário pq o ReportX só irá trabalhar com os controles que estiverem
    ' nas seções quando foi iniciado. Controles que são adicionados depois do
    ' inicio do relatorio não sao utilizados.
    
    ' Entao a tecnica consiste em copiar e deixar todos invisiveis e ir mostrando apenas
    ' os controles da seção corrente conforme o relatorio vai sendo executado.
    ' cada vez que mudar de conjunto de dados, as rotinas irão desabilitar os controles
    ' e habilitar somente aqueles que fizerem parte do conjunto que deve ser impresso.
    ' A ordem de impressao das secoes vai ser definida conforme a propriedade
    ' Index de cada secao do tipo indefinido.
    
    ' A secao temporarias chamam-se CabTmp(1 a N) , DetTmp(1 a N) e RodTmp(1 a N)
    
    Dim Ctr As Object
    
    For Each Ctr In Controls
        ' Se o controle estiver em um ReportSection do tipo 0-None, entao
        ' deve passar para dentro de um ReportSection CabGrp, Det ou RodGrp
        If TypeName(Ctr.Container) = "ReportSection" Then
            If Ctr.Container.Tipo = secNone Then
                ' Coloca no TAG o numero do indice do container original
                ' para sabermos de onde veio o controle.
                If TypeName(Ctr) = "Label" Then
                    Ctr.ToolTipText = Ctr.Container.Index
                Else
                    Ctr.Tag = Ctr.Container.Index
                End If
                Select Case Ctr.Container.Name
                    Case "CabTmp":
                        Set Ctr.Container = CabGrp
                    Case "DetTmp":
                        Set Ctr.Container = Det
                    Case "RodTmp":
                        Set Ctr.Container = RodGrp
                End Select
                ' Deixa cada controle invisivel depois de passar para a secao.
                ' sera habilitada conforme o andamento do relatorio
                StatusCtr Ctr, False
            End If
        End If
    Next
    
End Sub

Private Sub StatusCtr(Ctr As Object, Status As Boolean)
    
    ' Muda o status de visibilidade conforme o tipo de controle
    Select Case TypeName(Ctr)
        Case "ReportField": Ctr.Mostrar = Status
        Case "Line", "Label", "Shape", "Image", "PictureBox": Ctr.Visible = Status
    End Select
    
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

strTotalGeral = 0
strTotalVdaDia = 0
VarSomaTotal = 0


End Sub


Private Sub Relatorio_Erro(ByVal Numero As Long)

    Rpx_MsgErro Numero

End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)

    Dim NomeCampo As String
    Dim Grupo As Integer
    
    ' Campos formula no formato "XX.nome_campo" devem
    ' ser tratados nesse IF. Outros campos do tipo formula podem
    ' ser tratados no ELSE, se necessario.
    
    If Mid(Campo, 3, 1) = "." Then
        NomeCampo = Mid(Campo, 4)
        Grupo = Val(Left(Campo, 2))
        
        If Not RSSec(Grupo).Eof Then
            Valor = RSSec(Grupo).Fields(NomeCampo).value
        End If
        
        If NomeCampo = "Total" And Not RSSec(Grupo).Eof Then
            strTotalGeral = strTotalGeral + RSSec(Grupo).Fields(NomeCampo).value
        ElseIf NomeCampo = "SomaDeVlrVdaDia" And Not RSSec(Grupo).Eof Then
            strTotalVdaDia = strTotalVdaDia + RSSec(Grupo).Fields(NomeCampo).value
        End If
        
    Else
        Select Case Campo
            Case "TotalGeral"
                Valor = strTotalGeral
            Case "TotalVdaDia"
                Valor = strTotalGeral
        End Select
        ' Tratamentos de outros campos formula
        ' que nao fazem parte dos grupos
    End If
 'VarSomaTotal = VarSomaTotal + Relatorio.Recordset("Total")
         

 '   If Campo = "SomaTotal" Then
 '       Valor = VarSomaTotal
 '   End If
    
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
'    VarSomaTotal = VarSomaTotal + Relatorio.Recordset("Total")
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)

    If Ordem = 1 Then
        ' Aqui define uma opcional se houver um
        ' grupo de registros ligados aos dados
        ' caso contrario, nunca quebra
        Valor = 1
        ' Valor = Relatorio.Recordset("Campo_Grupo")
    Else
        ' O grupo interno faz a quebra
        ' quando passa ser EOF
        Valor = rds.RSAtual
    End If

End Sub

Private Sub Relatorio_IniciarGrupo(ByVal Ordem As Byte)
    
    ' Ao iniciar o grupo principal define os recordset secundarios
    If Ordem = 1 Then
        ' Abre todos os RS cada um para cada conjunto de dados
        ' Na mesma ordem que estao os indices de DetTmp.
        ' Se quiser pode usar os registros do grupo principal para filtrar
        ' os dados dos recordsets secundários. Basta usar a referencia
        ' Relatorio.Recordset
  
        rds.OpenSec 0, SqlDiv(1)
        rds.OpenSec 1, SqlDiv(2)

        'rds.OpenSec 2, "SELECT * FROM INSS ORDER BY Faixa DESC"
        ' Chama rotina que irá posicionar o grupo atual de acordo com
        ' o recordset que tiver registros
        rds.PosicionaRSAtual
    ElseIf Ordem = 2 Then
        ' Cada vez que um novo grupo secundario iniciar, ativa os controles
        ' de acordo desse grupo
        AtivaControles rds.RSAtual
    End If
    
End Sub

Private Sub AtivaControles(Grupo As Integer)
    
    Dim Ctr As Object
    Dim GrpCtr As Integer
    
    If Grupo > DetTmp.UBound Or Grupo < 0 Then
        
        CabGrp.Mostrar = False
        Det.Mostrar = False
        RodGrp.Mostrar = False
    
    Else
        
        For Each Ctr In Controls
            If TypeName(Ctr.Container) = "ReportSection" Then
                ' apenas a secao de detalhe, CabGrp e RodGrp
                Select Case Ctr.Container.Name
                    Case "Det", "CabGrp", "RodGrp"
                        If TypeName(Ctr) = "Label" Then
                            GrpCtr = Val(Ctr.ToolTipText)
                        Else
                            GrpCtr = Val(Ctr.Tag)
                        End If
                        If GrpCtr = Grupo Then
                            StatusCtr Ctr, True
                        Else
                            StatusCtr Ctr, False
                        End If
                End Select
            End If
        Next
        
        ' Copias as propriedades principais de cada grupo
        ' Se precisar de alguma propriedade que nao esteja
        ' relacionada, basta incluir aqui. Foram colocadas
        ' apenas as mais úteis.
        
        CabGrp.Mostrar = CabTmp(Grupo).Mostrar
        CabGrp.Height = CabTmp(Grupo).Height
        CabGrp.RepetirSecao = CabTmp(Grupo).RepetirSecao
        
        Det.Mostrar = DetTmp(Grupo).Mostrar
        Det.Height = DetTmp(Grupo).Height
        
        RodGrp.Mostrar = RodTmp(Grupo).Mostrar
        RodGrp.Height = RodTmp(Grupo).Height
        RodGrp.QuebraDepois = RodTmp(Grupo).QuebraDepois
            
        ' propriedades adicionais, caso necessario
        Det.AutoEncolher = DetTmp(Grupo).AutoEncolher
        Det.AutoExpandir = DetTmp(Grupo).AutoExpandir
    
    End If
    
End Sub

' Sub para apresentar mensagens de erro para o Visual ReportX
' Utilize sempre uma rotina no evento Erro do componente.
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
        
        ' Mensagens não previstas. Isso pode significar um erro
        ' interno no ReportX. Se isso acontecer, por favor reporte isso
        ' através de e-mail para ser corrigido.
        MsgBox "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impressão"
        
    End If
    
End Sub
