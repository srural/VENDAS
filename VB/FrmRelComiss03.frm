VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelComiss03 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   75.406
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1425
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2514
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   4
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
         Height          =   225
         Left            =   540
         TabIndex        =   7
         Top             =   930
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         Caption         =   "Pedido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   225
         Left            =   2670
         TabIndex        =   8
         Top             =   930
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   397
         Caption         =   "Entidade"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   225
         Left            =   7950
         TabIndex        =   9
         Top             =   930
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Caption         =   "Desconto"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   225
         Left            =   9045
         TabIndex        =   10
         Top             =   930
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Caption         =   "Comiss.Ped"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   225
         Left            =   6855
         TabIndex        =   12
         Top             =   930
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Caption         =   "Valor"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   225
         Left            =   1530
         TabIndex        =   13
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         Caption         =   "Data"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   225
         Left            =   1080
         TabIndex        =   15
         Top             =   690
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   397
         Caption         =   "Vendedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField29 
         Height          =   225
         Left            =   10215
         TabIndex        =   37
         Top             =   930
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Caption         =   "Comiss.Prd"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   2115
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField8 
         Height          =   225
         Left            =   225
         TabIndex        =   16
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Campo           =   "CodPed"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   225
         Left            =   2220
         TabIndex        =   17
         Top             =   0
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   397
         Campo           =   "Ent.Nome"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   225
         Left            =   7995
         TabIndex        =   18
         Top             =   0
         Width           =   990
         _ExtentX        =   1746
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
      Begin ReportX.ReportField ReportField12 
         Height          =   225
         Left            =   9045
         TabIndex        =   19
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "Valor"
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
      Begin ReportX.ReportField ReportField13 
         Height          =   225
         Left            =   6855
         TabIndex        =   20
         Top             =   0
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Campo           =   "Total"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   225
         Left            =   1125
         TabIndex        =   21
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         Campo           =   "DataEmiss"
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
      Begin ReportX.ReportField ReportField28 
         Height          =   225
         Left            =   10260
         TabIndex        =   36
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "SomaDeComissao"
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
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   10230
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
      Top             =   3255
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
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   2880
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Tipo            =   6
      Begin ReportX.ReportField ReportField15 
         Height          =   225
         Left            =   4530
         TabIndex        =   11
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         Caption         =   "Total Geral"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   225
         Left            =   7995
         TabIndex        =   26
         Top             =   60
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "forTotDesconto"
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
      Begin ReportX.ReportField ReportField19 
         Height          =   225
         Left            =   9045
         TabIndex        =   27
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "forTotValor"
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
      Begin ReportX.ReportField ReportField20 
         Height          =   225
         Left            =   6855
         TabIndex        =   28
         Top             =   45
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Campo           =   "forTotTotal"
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
      Begin ReportX.ReportField ReportField32 
         Height          =   225
         Left            =   10260
         TabIndex        =   40
         Top             =   45
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "forTotComiss"
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
      Begin ReportX.ReportField PedTotal 
         Height          =   225
         Left            =   5940
         TabIndex        =   42
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         Campo           =   "PedTotal"
         Formato         =   "0"
         Caption         =   ""
         TipoCampo       =   5
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
      Begin VB.Line Line1 
         X1              =   11430
         X2              =   90
         Y1              =   30
         Y2              =   30
      End
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Top             =   1425
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField16 
         Height          =   225
         Left            =   90
         TabIndex        =   24
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         Campo           =   "Vendedor.CodEntidade"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   225
         Left            =   1080
         TabIndex        =   25
         Top             =   120
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   397
         Campo           =   "Vendedor.Nome"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
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
         TabIndex        =   22
         Top             =   630
         Width           =   45
      End
      Begin VB.Line Line2 
         X1              =   11430
         X2              =   60
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line3 
         X1              =   11430
         X2              =   60
         Y1              =   360
         Y2              =   360
      End
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   2595
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField21 
         Height          =   225
         Index           =   0
         Left            =   4170
         TabIndex        =   29
         Top             =   60
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   397
         Caption         =   "Total do Vendedor"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField22 
         Height          =   225
         Left            =   7995
         TabIndex        =   30
         Top             =   60
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "forSomaDesconto"
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
      Begin ReportX.ReportField ReportField23 
         Height          =   225
         Left            =   9045
         TabIndex        =   31
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "forSomaValor"
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
      Begin ReportX.ReportField ReportField24 
         Height          =   225
         Left            =   6855
         TabIndex        =   32
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Campo           =   "forSomaTotal"
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
      Begin ReportX.ReportField ReportField31 
         Height          =   225
         Left            =   10260
         TabIndex        =   39
         Top             =   45
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "forSomaComiss"
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
      Begin ReportX.ReportField TotPed 
         Height          =   225
         Left            =   5940
         TabIndex        =   41
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         Campo           =   "TotPed"
         Formato         =   "0"
         Caption         =   ""
         TipoCampo       =   5
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
      Begin VB.Line Line4 
         X1              =   11430
         X2              =   90
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label Label2 
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
         TabIndex        =   23
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection CabGrp 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1845
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField ReportField33 
         Height          =   225
         Left            =   1035
         TabIndex        =   43
         Top             =   0
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   397
         Campo           =   "Cidade"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ReportX.ReportSection RodGrp 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   2355
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Tipo            =   5
      Ordem           =   2
      Begin ReportX.ReportField ReportField25 
         Height          =   225
         Left            =   8010
         TabIndex        =   33
         Top             =   0
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "forEntDesconto"
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
      Begin ReportX.ReportField ReportField26 
         Height          =   225
         Left            =   9045
         TabIndex        =   34
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "forEntValor"
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
      Begin ReportX.ReportField ReportField27 
         Height          =   225
         Left            =   6870
         TabIndex        =   35
         Top             =   0
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         Campo           =   "forEntTotal"
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
      Begin ReportX.ReportField ReportField30 
         Height          =   225
         Left            =   10260
         TabIndex        =   38
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "forComiss"
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
      Begin VB.Line Line5 
         X1              =   11340
         X2              =   7020
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "FrmRelComiss03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim somaTotal As Currency
Dim somaDesconto As Currency
Dim SomaValor As Currency

Dim totTotal As Currency
Dim totDesconto As Currency
Dim totValor As Currency

Dim EntValor As Currency
Dim EntDesconto As Currency
Dim EntTotal As Currency

Dim VarComiss As Currency
Dim VarSomaComiss As Currency
Dim VarTotComiss As Currency
Dim VarTotPed As Currency
Dim VarPedTotal  As Currency


'Private Rs As ADODB.Recordset

' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()

    'Dim Cn As String
        
    ' Uso do ADO nesse exemplo
    'Set Rs = New ADODB.Recordset
    
    ' Define a conexão. Pode ser usado uma conexao
    ' ja aberta pelo sistema.
    'Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\..\Exemplo.MDB"

    ' Abre o recordset com os dados
    'Rs.Open "SELECT * FROM ", Cn, adOpenStatic
    
    '  Associa o recordset ao relatorio
    'Set Relatorio.Recordset = Pesquisa(1)
    ' Inicia a geração do relatório.
    Relatorio.Ativar
    
    ' Fecha o recordset
    'Rs.Close
    'Set Rs = Nothing
    
    ' Retira o formulário de relatorio da memória
    Unload Me
    
End Sub

Private Sub Label9_Click()

End Sub


Private Sub Label14_Click()

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    ' O componente trabalha no modo silencioso para
    ' erros. Ele dispara esse evento Erro e sai. Caso
    ' o seu relatório esteja iniciando e saindo sem
    ' apresentar erro, verifique se você colocou algum
    ' código nesse evento.

    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' Configura o tamanho ideal para o Shape que servirá para
    ' fazer um relatório do tipo "zebrado"
    With shpZebra
        .Left = 0
        .Top = 0
        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
        .Height = Det.Height
    End With
    'lneCab.X1 = 0
    'lneCab.X2 = Relatorio.LarguraUtil
    'lneRod.X1 = 0
    'lneRod.X2 = Relatorio.LarguraUtil
    
    somaTotal = 0
    somaDesconto = 0
    SomaValor = 0
    totTotal = 0
    totDesconto = 0
    totValor = 0
    EntValor = 0
    EntDesconto = 0
    EntTotal = 0
    VarComiss = 0
    VarSomaComiss = 0
    VarTotComiss = 0
    VarTotPed = 0
    VarPedTotal = 0
    
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

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        Case "forSomaTotal"
            Valor = somaTotal
            somaTotal = 0
        Case "forSomaDesconto"
            Valor = somaDesconto
            somaDesconto = 0
        Case "forSomaValor"
            Valor = SomaValor
            SomaValor = 0
        
        Case "forTotTotal"
            Valor = totTotal
        Case "forTotDesconto"
            Valor = totDesconto
        Case "forTotValor"
            Valor = totValor
            
        Case "forEntTotal"
            Valor = EntTotal
            EntTotal = 0
        Case "forEntDesconto"
            Valor = EntDesconto
            EntDesconto = 0
        Case "forEntValor"
            Valor = EntValor
            EntValor = 0
        Case "forComiss"
            Valor = VarComiss
            VarComiss = 0
        Case "forSomaComiss"
            Valor = VarSomaComiss
            VarSomaComiss = 0
        Case "forTotComiss"
            Valor = VarTotComiss
'            VarTotComiss = 0
        Case "TotPed"
            Valor = VarTotPed
            VarTotPed = 0
            
        Case "PedTotal"
            Valor = VarPedTotal
            
                
    End Select
End Sub


Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada impressão de registro na linha de detalhe,
    ' soma as qtdes e os totais para apresentar no rodapé
    ' do grupo
        
    somaTotal = somaTotal + Relatorio.Recordset("Total")
    somaDesconto = somaDesconto + Relatorio.Recordset("Desconto")
    SomaValor = SomaValor + Relatorio.Recordset("Valor")
    
    totTotal = totTotal + Relatorio.Recordset("Total")
    totDesconto = totDesconto + Relatorio.Recordset("Desconto")
    totValor = totValor + Relatorio.Recordset("Valor")
    
    EntTotal = EntTotal + Relatorio.Recordset("Total")
    EntDesconto = EntDesconto + Relatorio.Recordset("Desconto")
    EntValor = EntValor + Relatorio.Recordset("Valor")
    
    VarComiss = VarComiss + Relatorio.Recordset("SomaDeComissao")
    VarSomaComiss = VarSomaComiss + Relatorio.Recordset("SomaDeComissao")
    VarTotComiss = VarTotComiss + Relatorio.Recordset("SomaDeComissao")
    
    VarTotPed = VarTotPed + 1
    VarPedTotal = VarPedTotal + 1
    
    
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    ' Indica que o grupo irá quebrar a cada novo codigo
    ' do pedido. Verifica se é formula para o grupo 1 ou
    ' para o grupo 2. Nesse caso o grupo 2 nao tem quebra
    ' apenas o grupo 1.
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("Vendedor.Nome")
        
    End If
    
    
    If Ordem = 2 Then
        Valor = Relatorio.Recordset("Ent.CodEntidade")
    End If
    
End Sub
