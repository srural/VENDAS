VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelCarne 
   Caption         =   "Relatório de Contas Financeiras"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   79.904
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   3780
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   2
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   3075
      Left            =   0
      Top             =   270
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5424
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
      Top             =   3345
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   7
   End
End
Attribute VB_Name = "FrmRelCarne"
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

