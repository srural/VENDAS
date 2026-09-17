VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelVdaFunGru 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   75.935
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   150
      TabIndex        =   0
      Top             =   3330
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1245
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2196
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
         Left            =   1050
         TabIndex        =   7
         Top             =   930
         Width           =   735
         _ExtentX        =   1296
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
      Begin ReportX.ReportField ReportField5 
         Height          =   225
         Left            =   1890
         TabIndex        =   8
         Top             =   945
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   397
         Caption         =   "Descrição do Grupo"
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
         Left            =   5595
         TabIndex        =   9
         Top             =   930
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         Caption         =   "Quantidade"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   225
         Left            =   8070
         TabIndex        =   10
         Top             =   930
         Width           =   930
         _ExtentX        =   1640
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
      Begin ReportX.ReportField ReportField10 
         Height          =   225
         Left            =   6630
         TabIndex        =   11
         Top             =   930
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Caption         =   "Valor Unit."
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
      Begin ReportX.ReportField ReportField11 
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   885
         _ExtentX        =   1561
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
      Begin ReportX.ReportField ReportField12 
         Height          =   225
         Left            =   1050
         TabIndex        =   15
         Top             =   690
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   397
         Caption         =   "Nome do Funcionário"
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
      Begin ReportX.ReportField ReportField20 
         Height          =   225
         Left            =   10080
         TabIndex        =   20
         Top             =   930
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Caption         =   "Comissão"
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
      Begin ReportX.ReportField ReportField16 
         Height          =   225
         Left            =   4335
         TabIndex        =   22
         Top             =   930
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         Caption         =   "Qtd.Pedido"
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
      Begin ReportX.ReportField ReportField23 
         Height          =   225
         Left            =   9090
         TabIndex        =   23
         Top             =   930
         Width           =   930
         _ExtentX        =   1640
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
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   120
         X2              =   11460
         Y1              =   660
         Y2              =   660
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   1965
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Mostrar         =   0   'False
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
      Top             =   3300
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
      Top             =   2925
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Tipo            =   6
      Begin ReportX.ReportField ReportField15 
         Height          =   225
         Left            =   6315
         TabIndex        =   12
         Top             =   90
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
         Left            =   7710
         TabIndex        =   13
         Top             =   105
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField22 
         Height          =   225
         Left            =   10305
         TabIndex        =   21
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   225
         Left            =   9135
         TabIndex        =   24
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField29 
         Height          =   225
         Left            =   4770
         TabIndex        =   39
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
         Campo           =   "TotalContador"
         Caption         =   ""
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
      Height          =   360
      Left            =   0
      Top             =   1245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField13 
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Campo           =   "CodEntidade"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   225
         Left            =   1050
         TabIndex        =   17
         Top             =   60
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   397
         Campo           =   "Nome"
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
      Begin VB.Line Line3 
         X1              =   11430
         X2              =   60
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   11430
         X2              =   60
         Y1              =   30
         Y2              =   30
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
         TabIndex        =   18
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   2565
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField21 
         Height          =   225
         Left            =   10290
         TabIndex        =   35
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField25 
         Height          =   225
         Left            =   7695
         TabIndex        =   36
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         Campo           =   "forValor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField26 
         Height          =   225
         Left            =   9075
         TabIndex        =   37
         Top             =   105
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   397
         Campo           =   "forDesconto"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField28 
         Height          =   225
         Left            =   4770
         TabIndex        =   38
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
         Campo           =   "forTotContador"
         Caption         =   ""
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
         X1              =   11460
         X2              =   120
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
         TabIndex        =   19
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection ReportSection4 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   1605
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      Tipo            =   3
      Mostrar         =   0   'False
      Ordem           =   2
      Begin VB.Label Label3 
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
         TabIndex        =   25
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection ReportSection5 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   2205
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      Tipo            =   5
      Ordem           =   2
      Begin ReportX.ReportField ReportField1 
         Height          =   225
         Left            =   1080
         TabIndex        =   27
         Top             =   90
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   397
         Campo           =   "ForCodGru"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField2 
         Height          =   225
         Left            =   1875
         TabIndex        =   28
         Top             =   90
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   397
         Campo           =   "ForDEscricao"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField3 
         Height          =   225
         Left            =   5625
         TabIndex        =   29
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         Campo           =   "ForQtd"
         Caption         =   ""
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
      Begin ReportX.ReportField ReportField7 
         Height          =   225
         Left            =   7875
         TabIndex        =   30
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Campo           =   "forSomadeValor"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   225
         Left            =   6705
         TabIndex        =   31
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         Campo           =   "ForMedia"
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
         Left            =   10275
         TabIndex        =   32
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         Campo           =   "ForComissao"
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
      Begin ReportX.ReportField ReportField18 
         Height          =   225
         Left            =   4770
         TabIndex        =   33
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
         Campo           =   "forContador"
         Caption         =   ""
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
         Left            =   9165
         TabIndex        =   34
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   397
         Campo           =   "SomaDesconto"
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
      Begin VB.Label Label4 
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
         TabIndex        =   26
         Top             =   630
         Width           =   45
      End
      Begin VB.Line Line5 
         X1              =   11460
         X2              =   120
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "FrmRelVdaFunGru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SomaValor As Currency
Dim ValorComiss As Currency
Dim TotComiss As Currency
Dim totValor As Currency
Dim somaDesconto As Currency
Dim totDesconto As Currency
Dim ValorDet As Currency
Dim DescontoDet As Currency
Dim VarDesc  As String
Dim VarCodGru As String
Dim VarContador As Currency
Dim VarQtd As Currency
Dim VarComissao As Currency
Dim VarMedia As Currency
Dim VarTotContador As Currency
Dim TotalContador As Currency

Public VarTotDesc As Currency
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
    
    VarCodGru = ""
    VarMedia = 0
    VarContador = 0
    VarQtd = 0
    VarComissao = 0
    SomaValor = 0
    ValorComiss = 0
    TotComiss = 0
    totValor = 0
    somaDesconto = 0
    totDesconto = 0
    VarTotDesc = 0
    ValorDet = 0
    DescontoDet = 0
    VarDesc = ""
    VarTotContador = 0
    TotalContador = 0
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
            Valor = SomaValor - VarTotDesc
            SomaValor = 0
        Case "forTotValor"
            Valor = totValor - totDesconto
        Case "forDesconto"
            Valor = VarTotDesc
            VarTotDesc = 0
        Case "forComiss"
            Valor = ValorComiss
            ValorComiss = 0
        Case "forTotComiss"
            Valor = TotComiss
        Case "forSomaDesconto"
            Valor = somaDesconto
            somaDesconto = 0
        Case "forTotDesconto"
            Valor = totDesconto
        Case "SomaDesconto"
            Valor = somaDesconto
            somaDesconto = 0
        Case "forSomadeValor"
            Valor = ValorDet - DescontoDet
            ValorDet = 0
            DescontoDet = 0
        Case "ForDEscricao"
            Valor = VarDesc
        Case "ForCodGru"
            Valor = VarCodGru
        Case "forContador"
            Valor = VarContador
            VarContador = 0

        Case "ForQtd"
            Valor = VarQtd
            VarQtd = 0
        Case "ForMedia"
            Valor = VarMedia
        Case "ForComissao"
            Valor = VarComissao
        Case "forTotContador"
            Valor = VarTotContador
            VarTotContador = 0
        Case "TotalContador"
            Valor = TotalContador
            
            
    End Select
End Sub


Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)

On Error Resume Next
    VarDesc = Relatorio.Recordset("Descrição_Grupo")
    VarCodGru = Relatorio.Recordset("CodGru")
    VarContador = VarContador + Relatorio.Recordset("ContarDeCodPed")
    VarQtd = VarQtd + Relatorio.Recordset("SomadeQTd")
    VarComissao = Relatorio.Recordset("SomaDeComissao")
    VarMedia = Relatorio.Recordset("MédiaDeValor")
    SomaValor = SomaValor + Relatorio.Recordset("SomaDeValor")
    totValor = totValor + Relatorio.Recordset("SomaDeValor")
    ValorDet = ValorDet + Relatorio.Recordset("SomaDeValor")
    DescontoDet = DescontoDet + Relatorio.Recordset("Desconto")
    totDesconto = totDesconto + Relatorio.Recordset("Desconto")
    somaDesconto = somaDesconto + Relatorio.Recordset("Desconto")
    VarTotDesc = VarTotDesc + Relatorio.Recordset("Desconto")
    VarTotContador = VarTotContador + Relatorio.Recordset("ContarDeCodPed")
    TotalContador = TotalContador + Relatorio.Recordset("ContarDeCodPed")
    
    If Not IsNull(Relatorio.Recordset("SomaDeComissao")) Then
        ValorComiss = ValorComiss + Relatorio.Recordset("SomaDeComissao")
        TotComiss = TotComiss + Relatorio.Recordset("SomaDeComissao")
    Else
        ValorComiss = ValorComiss + 0
        TotComiss = TotComiss + 0
    End If
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("CodEntidade")
    End If
    If Ordem = 2 Then
        Valor = Relatorio.Recordset("CodGru")
    End If
    
    
End Sub
