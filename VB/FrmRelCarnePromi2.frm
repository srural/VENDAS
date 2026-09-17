VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelCarnePromi2 
   Caption         =   "Relatório de Contas Financeiras"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   104.775
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   270
      TabIndex        =   0
      Top             =   9810
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
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Tipo            =   2
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   4830
      Left            =   0
      Top             =   270
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8520
      Begin ReportX.ReportField rpTipo2 
         Height          =   240
         Left            =   360
         TabIndex        =   6
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
      Begin ReportX.ReportField RpVenc 
         Height          =   240
         Left            =   1800
         TabIndex        =   7
         Top             =   1740
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
         Left            =   360
         TabIndex        =   8
         Top             =   135
         Width           =   2835
         _ExtentX        =   5001
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
         Left            =   1800
         TabIndex        =   9
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
         Left            =   360
         TabIndex        =   10
         Top             =   630
         Width           =   2835
         _ExtentX        =   5001
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
      Begin ReportX.ReportField RpValor 
         Height          =   240
         Left            =   1785
         TabIndex        =   11
         Top             =   2370
         Width           =   1335
         _ExtentX        =   2355
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
      Begin ReportX.ReportField RpEmissao 
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1740
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
      Begin ReportX.ReportField RpPedido2 
         Height          =   240
         Left            =   4245
         TabIndex        =   14
         Top             =   210
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
      Begin ReportX.ReportField ReportField1 
         Height          =   240
         Left            =   2610
         TabIndex        =   15
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
         Left            =   4995
         TabIndex        =   16
         Top             =   210
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
         Left            =   2940
         TabIndex        =   21
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
         Left            =   5325
         TabIndex        =   22
         Top             =   210
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
      Begin ReportX.ReportField DiaVenc 
         Height          =   225
         Left            =   8340
         TabIndex        =   24
         Top             =   180
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   397
         Campo           =   "forDia"
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
      Begin ReportX.ReportField AnoVenc 
         Height          =   225
         Left            =   10860
         TabIndex        =   25
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
         Campo           =   "forAno"
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
      Begin ReportX.ReportField MesVenc 
         Height          =   225
         Left            =   9240
         TabIndex        =   26
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Campo           =   "forMes"
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
      Begin ReportX.ReportField ValorPromissoria 
         Height          =   285
         Left            =   9765
         TabIndex        =   27
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         Campo           =   "Valor"
         Formato         =   "Standard"
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
         Left            =   4275
         TabIndex        =   32
         Top             =   1080
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   423
         Campo           =   "forData"
         Caption         =   ""
         Formula         =   -1  'True
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
         Left            =   4035
         TabIndex        =   33
         Top             =   1380
         Width           =   5655
         _ExtentX        =   9975
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
         Left            =   6840
         TabIndex        =   39
         Top             =   1650
         Width           =   4560
         _ExtentX        =   8043
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
         Left            =   4545
         TabIndex        =   41
         Top             =   1980
         Width           =   2340
         _ExtentX        =   4128
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
      Begin ReportX.ReportField valorextenso 
         Height          =   450
         Left            =   3750
         TabIndex        =   43
         Top             =   2295
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   794
         Linhas          =   2
         Campo           =   "forextenso"
         Caption         =   ""
         TipoCampo       =   7
         Formula         =   -1  'True
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField Emitente2 
         Height          =   240
         Left            =   4725
         TabIndex        =   45
         Top             =   2835
         Width           =   6435
         _ExtentX        =   11351
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
         Left            =   4725
         TabIndex        =   46
         Top             =   3585
         Width           =   2160
         _ExtentX        =   3810
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
         Left            =   4725
         TabIndex        =   47
         Top             =   3085
         Width           =   6435
         _ExtentX        =   11351
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
         Left            =   8460
         TabIndex        =   51
         Top             =   3465
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
         Left            =   10755
         TabIndex        =   52
         Top             =   3465
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
         Left            =   9225
         TabIndex        =   53
         Top             =   3465
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
      Begin ReportX.ReportField RpMens 
         Height          =   210
         Left            =   360
         TabIndex        =   62
         Top             =   4050
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   370
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField EmitenteCidade 
         Height          =   240
         Left            =   4725
         TabIndex        =   64
         Top             =   3335
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
      Begin ReportX.ReportField Pagavel 
         Height          =   240
         Left            =   4860
         TabIndex        =   67
         Top             =   3960
         Width           =   2025
         _ExtentX        =   3572
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagavel em:"
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
         Left            =   3780
         TabIndex        =   66
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Left            =   3780
         TabIndex        =   65
         Top             =   3315
         Width           =   600
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00EBEBEB&
         FillStyle       =   0  'Solid
         Height          =   465
         Left            =   3735
         Top             =   2295
         Width           =   7740
      End
      Begin VB.Label Label23 
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
         Left            =   4500
         TabIndex        =   61
         Top             =   645
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Shape Shape3 
         Height          =   405
         Left            =   4380
         Top             =   585
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label14 
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
         Left            =   3780
         TabIndex        =   60
         Top             =   645
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de mora e outras despesas."
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
         TabIndex        =   59
         Top             =   3690
         Width           =   2430
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "vencimento, não dispensar juros"
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
         Left            =   360
         TabIndex        =   58
         Top             =   3465
         Width           =   2790
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Na falta de pagamento no"
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
         Left            =   630
         TabIndex        =   57
         Top             =   3240
         Width           =   2220
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATENÇÃO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   56
         Top             =   2835
         Width           =   1395
      End
      Begin VB.Line Line23 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   7020
         X2              =   11280
         Y1              =   4185
         Y2              =   4185
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
         Left            =   10485
         TabIndex        =   55
         Top             =   3465
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
         Left            =   8940
         TabIndex        =   54
         Top             =   3465
         Width           =   210
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
         Left            =   3780
         TabIndex        =   50
         Top             =   3075
         Width           =   825
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
         Left            =   3780
         TabIndex        =   49
         Top             =   3555
         Width           =   855
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
         Left            =   3780
         TabIndex        =   48
         Top             =   2835
         Width           =   780
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ou a sua ordem a quantia de em moeda corrente deste pais"
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
         Left            =   6975
         TabIndex        =   44
         Top             =   1980
         Width           =   4290
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
         Left            =   3735
         TabIndex        =   42
         Top             =   1980
         Width           =   735
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
         Left            =   6660
         TabIndex        =   40
         Top             =   1650
         Width           =   105
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTA PROMISSÓRIA"
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
         Left            =   4815
         TabIndex        =   38
         Top             =   1665
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "única via de"
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
         Left            =   3735
         TabIndex        =   37
         Top             =   1620
         Width           =   1020
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
         Left            =   3735
         TabIndex        =   36
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pagar(ei)  por essa"
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
         Left            =   9765
         TabIndex        =   35
         Top             =   1380
         Width           =   1620
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
         Left            =   3735
         TabIndex        =   34
         Top             =   1080
         Width           =   465
      End
      Begin VB.Line Line28 
         X1              =   11430
         X2              =   7110
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Shape Shape7 
         Height          =   405
         Left            =   9690
         Top             =   600
         Width           =   1755
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
         Left            =   9255
         TabIndex        =   31
         Top             =   660
         Width           =   315
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
         Left            =   10530
         TabIndex        =   30
         Top             =   180
         Width           =   210
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
         Left            =   8910
         TabIndex        =   29
         Top             =   180
         Width           =   210
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
         Left            =   7110
         TabIndex        =   28
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3780
         TabIndex        =   23
         Top             =   135
         Width           =   315
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
         Left            =   5265
         TabIndex        =   20
         Top             =   210
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
         Left            =   2880
         TabIndex        =   19
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
         Left            =   4935
         TabIndex        =   18
         Top             =   225
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
         Left            =   2550
         TabIndex        =   17
         Top             =   1155
         Width           =   60
      End
      Begin VB.Line Line22 
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   3420
         X2              =   3435
         Y1              =   4815
         Y2              =   -90
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
         Left            =   375
         TabIndex        =   12
         Top             =   945
         Width           =   300
      End
      Begin VB.Shape Shape1 
         Height          =   4275
         Left            =   3615
         Top             =   105
         Width           =   7935
      End
      Begin VB.Line Line11 
         X1              =   3285
         X2              =   270
         Y1              =   2730
         Y2              =   2730
      End
      Begin VB.Line Line5 
         X1              =   1680
         X2              =   1680
         Y1              =   900
         Y2              =   1485
      End
      Begin VB.Line Line8 
         X1              =   3270
         X2              =   240
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line7 
         X1              =   3285
         X2              =   270
         Y1              =   1470
         Y2              =   1470
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
         TabIndex        =   5
         Top             =   630
         Width           =   45
      End
      Begin VB.Line Line9 
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   45
         X2              =   11505
         Y1              =   4590
         Y2              =   4590
      End
      Begin VB.Shape Shape2 
         Height          =   4275
         Left            =   270
         Top             =   105
         Width           =   3015
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1485
         Width           =   855
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
         Left            =   360
         TabIndex        =   3
         Top             =   1485
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
         Left            =   1800
         TabIndex        =   2
         Top             =   2145
         Width           =   390
      End
      Begin VB.Line Line6 
         X1              =   270
         X2              =   3285
         Y1              =   900
         Y2              =   900
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
         Left            =   1800
         TabIndex        =   1
         Top             =   945
         Width           =   525
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   1680
         Y1              =   1485
         Y2              =   2070
      End
      Begin VB.Line Line1 
         X1              =   270
         X2              =   3285
         Y1              =   435
         Y2              =   435
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
      Top             =   5100
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   7
   End
   Begin VB.Label Label19 
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
      Left            =   0
      TabIndex        =   63
      Top             =   765
      Width           =   825
   End
End
Attribute VB_Name = "FrmRelCarnePromi2"
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

    Select Case Campo
        Case "forDia"
            Valor = Day(Relatorio.Recordset("DtVenc"))
        Case "forMes"
            Valor = UCase(RetornaMes(Month(Relatorio.Recordset("DtVenc"))))
        Case "forAno"
            Valor = Year(Relatorio.Recordset("DtVenc"))
        Case "forData"
            Valor = ModPrincipal.valorextenso(Day(Relatorio.Recordset("DtVenc"))) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(Relatorio.Recordset("DtVenc")))) & " DO ANO DE " & ModPrincipal.valorextenso(Year(Relatorio.Recordset("DtVenc")))
        Case "forextenso"
            Valor = UCase(PassaExtenso(CStr(Relatorio.Recordset("valor")))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"

    End Select
End Sub


Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
Dim somaTotal As Integer
    ' A cada impressão de registro na linha de detalhe,
    ' soma as qtdes e os totais para apresentar no rodapé
    ' do grupo
        
'    somatotal = somatotal + Relatorio.Recordset("valor")
'    SomaDesconto = SomaDesconto + Relatorio.Recordset("Desconto")
'    SomaValor = SomaValor + Relatorio.Recordset("Valor")
    
End Sub

