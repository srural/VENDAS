VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelCarnePromi 
   Caption         =   "Relatório de Contas Financeiras"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   164.571
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   45
      TabIndex        =   0
      Top             =   8325
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   4980
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   2
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   3165
      Left            =   0
      Top             =   5250
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5583
      Begin ReportX.ReportField rpTipo2 
         Height          =   240
         Left            =   6165
         TabIndex        =   12
         Top             =   1155
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
         Campo           =   "Tipo_Dup"
         Caption         =   ""
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
      Begin ReportX.ReportField rpTipo 
         Height          =   240
         Left            =   540
         TabIndex        =   13
         Top             =   1155
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   423
         Campo           =   "Tipo_Dup"
         Caption         =   ""
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
      Begin ReportX.ReportField RpVenc 
         Height          =   240
         Left            =   4365
         TabIndex        =   14
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         Campo           =   "DtVenc"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
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
      Begin ReportX.ReportField rpTitulo 
         Height          =   240
         Left            =   540
         TabIndex        =   15
         Top             =   165
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField rpPedido 
         Height          =   240
         Left            =   1290
         TabIndex        =   16
         Top             =   1155
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   423
         Campo           =   "Pedido"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField RpNome 
         Height          =   240
         Left            =   540
         TabIndex        =   17
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField RpMens 
         Height          =   240
         Left            =   540
         TabIndex        =   18
         Top             =   2565
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField RpMens2 
         Height          =   240
         Left            =   6150
         TabIndex        =   19
         Top             =   2565
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField RpValor 
         Height          =   240
         Left            =   2910
         TabIndex        =   20
         Top             =   1650
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   423
         Campo           =   "Valor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField RpTitulo2 
         Height          =   240
         Left            =   6150
         TabIndex        =   31
         Top             =   165
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField RpNome2 
         Height          =   240
         Left            =   6195
         TabIndex        =   32
         Top             =   675
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField RpEmissao 
         Height          =   240
         Left            =   2895
         TabIndex        =   33
         Top             =   1155
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   423
         Campo           =   "DtEmiss"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
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
      Begin ReportX.ReportField RpEmissao2 
         Height          =   240
         Left            =   8505
         TabIndex        =   34
         Top             =   1155
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Campo           =   "DtEmiss"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
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
      Begin ReportX.ReportField RpVenc2 
         Height          =   240
         Left            =   9975
         TabIndex        =   35
         Top             =   1155
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   423
         Campo           =   "DtVenc"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
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
      Begin ReportX.ReportField RpPedido2 
         Height          =   240
         Left            =   6945
         TabIndex        =   36
         Top             =   1155
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   423
         Campo           =   "Pedido"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField RpValor2 
         Height          =   240
         Left            =   8505
         TabIndex        =   37
         Top             =   1650
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Campo           =   "Valor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField1 
         Height          =   240
         Left            =   2070
         TabIndex        =   38
         Top             =   1155
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Campo           =   "NroPcl"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField ReportField2 
         Height          =   240
         Left            =   7695
         TabIndex        =   39
         Top             =   1155
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Campo           =   "NroPcl"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField RpTotPcl 
         Height          =   240
         Left            =   2400
         TabIndex        =   44
         Top             =   1155
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField RpTotPcl2 
         Height          =   240
         Left            =   8025
         TabIndex        =   45
         Top             =   1155
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   ""
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7965
         TabIndex        =   43
         Top             =   1155
         Width           =   60
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2340
         TabIndex        =   42
         Top             =   1155
         Width           =   60
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7635
         TabIndex        =   41
         Top             =   1170
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2010
         TabIndex        =   40
         Top             =   1155
         Width           =   60
      End
      Begin VB.Line Line22 
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   5910
         X2              =   5910
         Y1              =   3060
         Y2              =   -90
      End
      Begin VB.Line Line21 
         X1              =   6060
         X2              =   11400
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line20 
         X1              =   8370
         X2              =   8370
         Y1              =   1425
         Y2              =   2520
      End
      Begin VB.Line Line19 
         X1              =   9840
         X2              =   9840
         Y1              =   1425
         Y2              =   2520
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ped/Pcl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6960
         TabIndex        =   30
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6135
         TabIndex        =   29
         Top             =   945
         Width           =   300
      End
      Begin VB.Line Line18 
         X1              =   6060
         X2              =   11400
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8445
         TabIndex        =   28
         Top             =   1425
         Width           =   390
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8505
         TabIndex        =   27
         Top             =   945
         Width           =   600
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6180
         TabIndex        =   26
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9960
         TabIndex        =   25
         Top             =   945
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   2790
         Left            =   6060
         Top             =   105
         Width           =   5355
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6135
         TabIndex        =   24
         Top             =   1455
         Width           =   375
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros/Desc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9915
         TabIndex        =   23
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Pgto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8445
         TabIndex        =   22
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9915
         TabIndex        =   21
         Top             =   1935
         Width           =   795
      End
      Begin VB.Line Line17 
         X1              =   11400
         X2              =   6060
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Line Line16 
         X1              =   11400
         X2              =   8370
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line15 
         X1              =   9840
         X2              =   9840
         Y1              =   945
         Y2              =   1425
      End
      Begin VB.Line Line14 
         X1              =   8370
         X2              =   8370
         Y1              =   945
         Y2              =   1455
      End
      Begin VB.Line Line13 
         X1              =   6840
         X2              =   6840
         Y1              =   945
         Y2              =   1425
      End
      Begin VB.Line Line12 
         X1              =   11400
         X2              =   6060
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Line Line11 
         X1              =   5790
         X2              =   450
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Line Line10 
         X1              =   1200
         X2              =   1200
         Y1              =   945
         Y2              =   1425
      End
      Begin VB.Line Line5 
         X1              =   2760
         X2              =   2760
         Y1              =   945
         Y2              =   1455
      End
      Begin VB.Line Line4 
         X1              =   4230
         X2              =   4230
         Y1              =   945
         Y2              =   1425
      End
      Begin VB.Line Line8 
         X1              =   5790
         X2              =   2760
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line7 
         X1              =   5790
         X2              =   450
         Y1              =   1425
         Y2              =   1425
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
         TabIndex        =   11
         Top             =   630
         Width           =   45
      End
      Begin VB.Line Line9 
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   45
         X2              =   11505
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4305
         TabIndex        =   10
         Top             =   1935
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Pgto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2835
         TabIndex        =   9
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros/Desc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4305
         TabIndex        =   8
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   525
         TabIndex        =   7
         Top             =   1455
         Width           =   375
      End
      Begin VB.Shape Shape2 
         Height          =   2790
         Left            =   450
         Top             =   105
         Width           =   5355
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4350
         TabIndex        =   6
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   570
         TabIndex        =   5
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2895
         TabIndex        =   4
         Top             =   945
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2835
         TabIndex        =   3
         Top             =   1425
         Width           =   390
      End
      Begin VB.Line Line6 
         X1              =   450
         X2              =   5790
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   525
         TabIndex        =   2
         Top             =   945
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ped/Pcl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1275
         TabIndex        =   1
         Top             =   945
         Width           =   525
      End
      Begin VB.Line Line3 
         X1              =   4230
         X2              =   4230
         Y1              =   1425
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2760
         Y1              =   1425
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   5790
         Y1              =   480
         Y2              =   480
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
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Top             =   8415
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   7
   End
   Begin ReportX.ReportSection SecPromissoria 
      Align           =   1  'Align Top
      Height          =   4980
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8784
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField DiaVenc 
         Height          =   225
         Left            =   8220
         TabIndex        =   46
         Top             =   510
         Width           =   435
         _ExtentX        =   767
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
      Begin ReportX.ReportField AnoVenc 
         Height          =   225
         Left            =   10740
         TabIndex        =   47
         Top             =   510
         Width           =   555
         _ExtentX        =   979
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
      Begin ReportX.ReportField MesVenc 
         Height          =   225
         Left            =   9120
         TabIndex        =   48
         Top             =   510
         Width           =   1215
         _ExtentX        =   2143
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
      Begin ReportX.ReportField NroPedido 
         Height          =   285
         Left            =   2940
         TabIndex        =   49
         Top             =   1080
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   2
         Mostrar         =   0   'False
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
      Begin ReportX.ReportField ValorPromissoria 
         Height          =   285
         Left            =   9660
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         Caption         =   ""
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
      Begin ReportX.ReportField ExtensoDia 
         Height          =   240
         Left            =   1005
         TabIndex        =   51
         Top             =   1635
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField Emitente 
         Height          =   240
         Left            =   840
         TabIndex        =   52
         Top             =   2025
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField NomeEmpresa 
         Height          =   240
         Left            =   690
         TabIndex        =   53
         Top             =   2445
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField EmpresaCGCCPF 
         Height          =   240
         Left            =   8580
         TabIndex        =   54
         Top             =   2445
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField QuantiaExtenso 
         Height          =   480
         Left            =   3180
         TabIndex        =   55
         Top             =   2835
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   847
         Linhas          =   2
         Caption         =   ""
         TipoCampo       =   7
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
      Begin ReportX.ReportField Pagavel 
         Height          =   240
         Left            =   1680
         TabIndex        =   56
         Top             =   3450
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField Emitente2 
         Height          =   240
         Left            =   1680
         TabIndex        =   57
         Top             =   3712
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField EmitenteCpf 
         Height          =   240
         Left            =   1680
         TabIndex        =   58
         Top             =   3974
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField EmitenteEndereco 
         Height          =   240
         Left            =   1680
         TabIndex        =   59
         Top             =   4236
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField RpDia 
         Height          =   225
         Left            =   7920
         TabIndex        =   60
         Top             =   3570
         Width           =   435
         _ExtentX        =   767
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
      Begin ReportX.ReportField RpAno 
         Height          =   225
         Left            =   10440
         TabIndex        =   61
         Top             =   3570
         Width           =   555
         _ExtentX        =   979
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
      Begin ReportX.ReportField RpMes 
         Height          =   225
         Left            =   8820
         TabIndex        =   62
         Top             =   3570
         Width           =   1215
         _ExtentX        =   2143
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
      Begin ReportX.ReportField EmitenteCidade 
         Height          =   240
         Left            =   1680
         TabIndex        =   82
         Top             =   4500
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   423
         Caption         =   ""
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
      Begin ReportX.ReportField NroPedido1 
         Height          =   240
         Left            =   9600
         TabIndex        =   84
         Top             =   2055
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   423
         Caption         =   ""
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
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pagar(ei)  por essa promissoria numero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   85
         Top             =   2040
         Width           =   3390
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   83
         Top             =   4500
         Width           =   600
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTA PROMISSÓRIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   81
         Top             =   270
         Width           =   3705
      End
      Begin VB.Line Line30 
         BorderWidth     =   4
         X1              =   630
         X2              =   4320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line29 
         BorderWidth     =   4
         X1              =   4320
         X2              =   4320
         Y1              =   420
         Y2              =   690
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
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
         Left            =   6990
         TabIndex        =   80
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
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
         Left            =   8790
         TabIndex        =   79
         Top             =   510
         Width           =   210
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
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
         Left            =   10410
         TabIndex        =   78
         Top             =   510
         Width           =   210
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
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
         Left            =   600
         TabIndex        =   77
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Shape Shape6 
         Height          =   405
         Left            =   1200
         Top             =   1020
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9135
         TabIndex        =   76
         Top             =   1080
         Width           =   315
      End
      Begin VB.Shape Shape7 
         Height          =   405
         Left            =   9570
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Line Line28 
         X1              =   11310
         X2              =   6990
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ao(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   495
         TabIndex        =   75
         Top             =   1635
         Width           =   465
      End
      Begin VB.Line Line27 
         X1              =   510
         X2              =   11400
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   74
         Top             =   2445
         Width           =   105
      End
      Begin VB.Line Line26 
         X1              =   510
         X2              =   11400
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.N.P.J."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7590
         TabIndex        =   73
         Top             =   2445
         Width           =   735
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ou a sua ordem a quantia de"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   72
         Top             =   2835
         Width           =   2475
      End
      Begin VB.Line Line25 
         X1              =   510
         X2              =   11400
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "em moeda corrente deste pais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   71
         Top             =   3075
         Width           =   2640
      End
      Begin VB.Line Line24 
         X1              =   510
         X2              =   11400
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagável em"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   70
         Top             =   3435
         Width           =   1005
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emitente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   69
         Top             =   3705
         Width           =   780
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cpf / Cnpj"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   68
         Top             =   3975
         Width           =   855
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   67
         Top             =   4245
         Width           =   825
      End
      Begin VB.Line Line23 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   7035
         X2              =   11295
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÚNICA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   66
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
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
         Left            =   10110
         TabIndex        =   65
         Top             =   3570
         Width           =   210
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
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
         Left            =   8490
         TabIndex        =   64
         Top             =   3570
         Width           =   210
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   63
         Top             =   2010
         Width           =   240
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4650
         Left            =   390
         Top             =   195
         Width           =   11115
      End
   End
End
Attribute VB_Name = "FrmRelCarnePromi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
'    With shpZebra
'        .Left = 0
'        .Top = 0
'        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
'        .Height = Det.Height
'    End With
    'lneCab.X1 = 0
    'lneCab.X2 = Relatorio.LarguraUtil
    'lneRod.X1 = 0
    'lneRod.X2 = Relatorio.LarguraUtil
    
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
'Dim Completa As String
'Completa = " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"

 '   Select Case Campo
 '       Case "ForParcela"
 '           Valor = Relatorio.Recordset("Valor") + Relatorio.Recordset("Juros")
 '       Case "ForExtenso"
 '           Valor = Format(PassaExtenso(Relatorio.Recordset("Valor") + Relatorio.Recordset("Juros")), ">") & Completa
 '   End Select
End Sub

