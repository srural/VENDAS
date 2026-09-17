VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPedPromi 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   11595
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   204.523
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
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
         Left            =   7950
         TabIndex        =   30
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
         Left            =   10470
         TabIndex        =   31
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
         Left            =   8850
         TabIndex        =   32
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
      Begin ReportX.ReportField ValorPromissoria 
         Height          =   285
         Left            =   9390
         TabIndex        =   37
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
         Left            =   735
         TabIndex        =   39
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
         Left            =   570
         TabIndex        =   40
         Top             =   2025
         Width           =   5055
         _ExtentX        =   8916
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
         Left            =   420
         TabIndex        =   42
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
         Left            =   8310
         TabIndex        =   44
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
         Left            =   2910
         TabIndex        =   46
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
         Left            =   1410
         TabIndex        =   49
         Top             =   3405
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
         Left            =   1410
         TabIndex        =   51
         Top             =   3645
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
         Left            =   1410
         TabIndex        =   53
         Top             =   3885
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
         Left            =   1410
         TabIndex        =   55
         Top             =   4125
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
         Left            =   7650
         TabIndex        =   63
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
         Left            =   10170
         TabIndex        =   64
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
         Left            =   8550
         TabIndex        =   65
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
         Left            =   1410
         TabIndex        =   222
         Top             =   4365
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
         Left            =   9360
         TabIndex        =   224
         Top             =   2040
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
      Begin VB.Label Label18 
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
         Left            =   270
         TabIndex        =   223
         Top             =   4365
         Width           =   600
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
         Left            =   240
         TabIndex        =   71
         Top             =   2010
         Width           =   240
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
         Left            =   8220
         TabIndex        =   67
         Top             =   3570
         Width           =   210
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
         Left            =   9840
         TabIndex        =   66
         Top             =   3570
         Width           =   210
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
         Left            =   1050
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Line Line12 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   6765
         X2              =   11025
         Y1              =   4440
         Y2              =   4440
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
         Left            =   270
         TabIndex        =   56
         Top             =   4110
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
         Left            =   270
         TabIndex        =   54
         Top             =   3840
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
         Left            =   270
         TabIndex        =   52
         Top             =   3615
         Width           =   780
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
         Left            =   270
         TabIndex        =   50
         Top             =   3390
         Width           =   1005
      End
      Begin VB.Line Line11 
         X1              =   240
         X2              =   11130
         Y1              =   3375
         Y2              =   3375
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
         Left            =   240
         TabIndex        =   48
         Top             =   3075
         Width           =   2640
      End
      Begin VB.Line Line10 
         X1              =   240
         X2              =   11130
         Y1              =   2760
         Y2              =   2760
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
         Left            =   240
         TabIndex        =   47
         Top             =   2835
         Width           =   2475
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
         Left            =   7320
         TabIndex        =   45
         Top             =   2445
         Width           =   735
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   11130
         Y1              =   2340
         Y2              =   2340
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
         Left            =   240
         TabIndex        =   43
         Top             =   2445
         Width           =   105
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
         Left            =   5760
         TabIndex        =   41
         Top             =   2025
         Width           =   3390
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   11130
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Label30 
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
         Left            =   225
         TabIndex        =   38
         Top             =   1635
         Width           =   465
      End
      Begin VB.Line Line6 
         X1              =   11040
         X2              =   6720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Shape Shape7 
         Height          =   405
         Left            =   9300
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label Label29 
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
         Left            =   8865
         TabIndex        =   36
         Top             =   1080
         Width           =   315
      End
      Begin VB.Shape Shape6 
         Height          =   405
         Left            =   930
         Top             =   1020
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label28 
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
         Left            =   330
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label24 
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
         Left            =   10140
         TabIndex        =   34
         Top             =   510
         Width           =   210
      End
      Begin VB.Label Label23 
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
         Left            =   8520
         TabIndex        =   33
         Top             =   510
         Width           =   210
      End
      Begin VB.Label Label22 
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
         Left            =   6720
         TabIndex        =   29
         Top             =   510
         Width           =   1065
      End
      Begin VB.Line Line5 
         BorderWidth     =   4
         X1              =   4050
         X2              =   4050
         Y1              =   420
         Y2              =   690
      End
      Begin VB.Line Line4 
         BorderWidth     =   4
         X1              =   360
         X2              =   4050
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label21 
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
         Left            =   270
         TabIndex        =   28
         Top             =   270
         Width           =   3660
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4650
         Left            =   120
         Top             =   60
         Width           =   11115
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   210
      TabIndex        =   0
      Top             =   6660
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   2790
      Left            =   0
      Top             =   4980
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   4921
      Tipo            =   1
      Ordem           =   1
      Begin ReportX.ReportField rpCabInformacao 
         Height          =   240
         Left            =   1650
         TabIndex        =   5
         Top             =   720
         Width           =   9585
         _ExtentX        =   16907
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
         Left            =   9360
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
         Height          =   570
         Left            =   1650
         TabIndex        =   1
         Top             =   150
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   1005
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   2
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField23 
         Height          =   225
         Left            =   1185
         TabIndex        =   182
         Top             =   1440
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
         TabIndex        =   183
         Top             =   1950
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
         TabIndex        =   184
         Top             =   1950
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
         TabIndex        =   185
         Top             =   2190
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
         TabIndex        =   186
         Top             =   2190
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
         TabIndex        =   187
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
         TabIndex        =   188
         Top             =   2430
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
         TabIndex        =   189
         Top             =   2430
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
         TabIndex        =   190
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
         TabIndex        =   191
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
         TabIndex        =   192
         Top             =   1695
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
         TabIndex        =   193
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
         TabIndex        =   194
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
         TabIndex        =   195
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
         TabIndex        =   196
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
         TabIndex        =   197
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
         TabIndex        =   198
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
         TabIndex        =   199
         Top             =   2190
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
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   210
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
         TabIndex        =   209
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
         TabIndex        =   208
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
         TabIndex        =   207
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
         TabIndex        =   206
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
         TabIndex        =   205
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
         TabIndex        =   204
         Top             =   1440
         Width           =   540
      End
      Begin VB.Shape Shape9 
         Height          =   1365
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
         TabIndex        =   203
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
         TabIndex        =   202
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
         TabIndex        =   201
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
         TabIndex        =   200
         Top             =   2190
         Width           =   360
      End
      Begin VB.Image Logotipo 
         Height          =   960
         Left            =   90
         Stretch         =   -1  'True
         Top             =   75
         Width           =   11235
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   10830
         X2              =   11280
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   120
         X2              =   9300
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   9690
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   165
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
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
         Left            =   7710
         TabIndex        =   15
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   370
         Campo           =   "forEntregue"
         Caption         =   "Nome"
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
      End
      Begin ReportX.ReportField rpDetalhes 
         Height          =   210
         Index           =   2
         Left            =   9075
         TabIndex        =   16
         Top             =   0
         Width           =   1020
         _ExtentX        =   1799
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
         Left            =   8145
         TabIndex        =   17
         Top             =   0
         Width           =   870
         _ExtentX        =   1535
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
         TabIndex        =   18
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "ForVlrVdaDia"
         Formato         =   "Standard"
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
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   5535
         TabIndex        =   61
         Top             =   0
         Width           =   1650
         _ExtentX        =   2910
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
         Left            =   7245
         TabIndex        =   70
         Top             =   0
         Width           =   420
         _ExtentX        =   741
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
      Begin ReportX.ReportField ReportField29 
         Height          =   210
         Left            =   1260
         TabIndex        =   177
         Top             =   165
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   370
         Campo           =   "Referencia"
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
      Height          =   420
      Left            =   0
      Top             =   7770
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8910
         TabIndex        =   3
         Top             =   90
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
         Top             =   90
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
         X1              =   11430
         X2              =   90
         Y1              =   60
         Y2              =   60
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   1155
      Index           =   1
      Left            =   0
      Top             =   8190
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2037
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField2 
         Height          =   225
         Left            =   7230
         TabIndex        =   9
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   397
         Campo           =   "Cidade_cob"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   225
         Left            =   9810
         TabIndex        =   11
         Top             =   270
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
         Campo           =   "Uf_Cob"
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
      Begin ReportX.ReportField ReportField1 
         Height          =   225
         Left            =   180
         TabIndex        =   58
         Top             =   270
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   397
         Campo           =   "Endereço_Cob"
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
      Begin ReportX.ReportField ReportField19 
         Height          =   225
         Left            =   4440
         TabIndex        =   59
         Top             =   270
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "Nro_Cob"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   5250
         TabIndex        =   60
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
         Height          =   225
         Left            =   7230
         TabIndex        =   80
         Top             =   855
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   397
         Campo           =   "Cidade_Entrega"
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
      Begin ReportX.ReportField ReportField5 
         Height          =   225
         Left            =   9810
         TabIndex        =   81
         Top             =   855
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         Campo           =   "Uf_Entrega"
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
      Begin ReportX.ReportField ReportField6 
         Height          =   225
         Left            =   210
         TabIndex        =   82
         Top             =   855
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   397
         Campo           =   "Endereço_Entrega"
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
      Begin ReportX.ReportField ReportField7 
         Height          =   225
         Left            =   4440
         TabIndex        =   83
         Top             =   855
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "Nro_Entrega"
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
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   5220
         TabIndex        =   84
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
         Height          =   225
         Left            =   10110
         TabIndex        =   178
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         Campo           =   "Cep_Cob"
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
      Begin ReportX.ReportField ReportField21 
         Height          =   225
         Left            =   10080
         TabIndex        =   180
         Top             =   870
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         Campo           =   "Cep_Entrega"
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
         Left            =   10080
         TabIndex        =   181
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
         Left            =   10110
         TabIndex        =   179
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
         TabIndex        =   89
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
         Left            =   7230
         TabIndex        =   88
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
         Left            =   9810
         TabIndex        =   87
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
         Left            =   4470
         TabIndex        =   86
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
         Left            =   5220
         TabIndex        =   85
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
         Left            =   5250
         TabIndex        =   57
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
         Left            =   4440
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
         Left            =   9810
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
         Left            =   7230
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
      Height          =   4440
      Index           =   1
      Left            =   0
      Top             =   10140
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   7832
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField rpGrupo 
         Height          =   225
         Index           =   0
         Left            =   9585
         TabIndex        =   22
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
         Left            =   9585
         TabIndex        =   23
         Top             =   585
         Width           =   1680
         _ExtentX        =   2963
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
         Height          =   225
         Index           =   2
         Left            =   9585
         TabIndex        =   24
         Top             =   855
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   397
         Campo           =   "forTotalPedido"
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
      Begin ReportX.ReportField Obs3 
         Height          =   210
         Left            =   390
         TabIndex        =   21
         Top             =   570
         Width           =   7275
         _ExtentX        =   12832
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
         Left            =   390
         TabIndex        =   20
         Top             =   330
         Width           =   7275
         _ExtentX        =   12832
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
         Left            =   390
         TabIndex        =   19
         Top             =   90
         Width           =   7275
         _ExtentX        =   12832
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
         TabIndex        =   68
         Top             =   315
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
      Begin ReportX.ReportField Transport 
         Height          =   210
         Left            =   1650
         TabIndex        =   90
         Top             =   3255
         Width           =   5655
         _ExtentX        =   9975
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
         Left            =   8565
         TabIndex        =   91
         Top             =   3255
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
      Begin ReportX.ReportField Venc 
         Height          =   210
         Index           =   0
         Left            =   1110
         TabIndex        =   92
         Top             =   1545
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
         Left            =   2040
         TabIndex        =   93
         Top             =   1545
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5895
         TabIndex        =   94
         Top             =   1545
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
         Left            =   6825
         TabIndex        =   95
         Top             =   1545
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
         Left            =   7755
         TabIndex        =   96
         Top             =   1545
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   9660
         TabIndex        =   97
         Top             =   1545
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
         Left            =   10590
         TabIndex        =   98
         Top             =   1545
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3090
         TabIndex        =   99
         Top             =   1755
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
         Left            =   3990
         TabIndex        =   100
         Top             =   1755
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
         Left            =   4920
         TabIndex        =   101
         Top             =   1755
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   6825
         TabIndex        =   102
         Top             =   1755
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
         Left            =   7755
         TabIndex        =   103
         Top             =   1755
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   195
         TabIndex        =   104
         Top             =   1965
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
         Left            =   1110
         TabIndex        =   105
         Top             =   1965
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
         Left            =   2040
         TabIndex        =   106
         Top             =   1965
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3990
         TabIndex        =   107
         Top             =   1965
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
         Left            =   4920
         TabIndex        =   108
         Top             =   1965
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   8745
         TabIndex        =   109
         Top             =   1965
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
         Left            =   9660
         TabIndex        =   110
         Top             =   1965
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
         Left            =   10590
         TabIndex        =   111
         Top             =   1965
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   1110
         TabIndex        =   112
         Top             =   2175
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
         Left            =   2040
         TabIndex        =   113
         Top             =   2175
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5895
         TabIndex        =   114
         Top             =   2175
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
         Left            =   6825
         TabIndex        =   115
         Top             =   2175
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
         Left            =   7755
         TabIndex        =   116
         Top             =   2175
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   9660
         TabIndex        =   117
         Top             =   2175
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
         Left            =   10590
         TabIndex        =   118
         Top             =   2175
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3090
         TabIndex        =   119
         Top             =   2385
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
         Left            =   3990
         TabIndex        =   120
         Top             =   2385
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
         Left            =   4920
         TabIndex        =   121
         Top             =   2385
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   195
         TabIndex        =   122
         Top             =   1545
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
         Left            =   8745
         TabIndex        =   123
         Top             =   1545
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
         Left            =   5895
         TabIndex        =   124
         Top             =   1755
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
         Left            =   3090
         TabIndex        =   125
         Top             =   1965
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
         Left            =   195
         TabIndex        =   126
         Top             =   2175
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
         Left            =   8745
         TabIndex        =   127
         Top             =   2175
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
         Left            =   3090
         TabIndex        =   128
         Top             =   1545
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
         Left            =   3990
         TabIndex        =   129
         Top             =   1545
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
         Left            =   4920
         TabIndex        =   130
         Top             =   1545
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   195
         TabIndex        =   131
         Top             =   1755
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
         Left            =   1110
         TabIndex        =   132
         Top             =   1755
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
         Left            =   2040
         TabIndex        =   133
         Top             =   1755
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   8745
         TabIndex        =   134
         Top             =   1755
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
         Left            =   9660
         TabIndex        =   135
         Top             =   1755
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
         Left            =   10590
         TabIndex        =   136
         Top             =   1755
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5895
         TabIndex        =   137
         Top             =   1965
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
         Left            =   6825
         TabIndex        =   138
         Top             =   1965
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
         Left            =   7755
         TabIndex        =   139
         Top             =   1965
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3090
         TabIndex        =   140
         Top             =   2175
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
         Left            =   3990
         TabIndex        =   141
         Top             =   2175
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
         Left            =   4920
         TabIndex        =   142
         Top             =   2175
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   195
         TabIndex        =   143
         Top             =   2385
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
         Left            =   1110
         TabIndex        =   144
         Top             =   2385
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
         Left            =   2040
         TabIndex        =   145
         Top             =   2385
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5895
         TabIndex        =   146
         Top             =   2385
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
         Left            =   6825
         TabIndex        =   147
         Top             =   2385
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
         Left            =   7755
         TabIndex        =   148
         Top             =   2385
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   8745
         TabIndex        =   149
         Top             =   2385
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
         Left            =   9660
         TabIndex        =   150
         Top             =   2385
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
         Left            =   10590
         TabIndex        =   151
         Top             =   2385
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   195
         TabIndex        =   152
         Top             =   2595
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
         Left            =   1110
         TabIndex        =   153
         Top             =   2595
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
         Left            =   2040
         TabIndex        =   154
         Top             =   2595
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3090
         TabIndex        =   155
         Top             =   2595
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
         Left            =   3990
         TabIndex        =   156
         Top             =   2595
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
         Left            =   4920
         TabIndex        =   157
         Top             =   2595
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5895
         TabIndex        =   158
         Top             =   2595
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
         Left            =   6825
         TabIndex        =   159
         Top             =   2595
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
         Left            =   7755
         TabIndex        =   160
         Top             =   2595
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   8745
         TabIndex        =   161
         Top             =   2595
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
         Left            =   9660
         TabIndex        =   162
         Top             =   2595
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
         Left            =   10590
         TabIndex        =   163
         Top             =   2595
         Width           =   735
         _ExtentX        =   1296
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
      Begin ReportX.ReportField RpCondpgto 
         Height          =   225
         Left            =   2625
         TabIndex        =   164
         Top             =   1260
         Width           =   2340
         _ExtentX        =   4128
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
         Left            =   7245
         TabIndex        =   165
         Top             =   1260
         Width           =   2340
         _ExtentX        =   4128
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
      Begin ReportX.ReportField RpPrevEntrega 
         Height          =   225
         Left            =   1650
         TabIndex        =   166
         Top             =   2985
         Width           =   2820
         _ExtentX        =   4974
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
         Left            =   5685
         TabIndex        =   167
         Top             =   2985
         Width           =   1620
         _ExtentX        =   2858
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
         Left            =   8565
         TabIndex        =   168
         Top             =   2985
         Width           =   2715
         _ExtentX        =   4789
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
      Begin ReportX.ReportField Obs4 
         Height          =   210
         Left            =   390
         TabIndex        =   176
         Top             =   810
         Width           =   7275
         _ExtentX        =   12832
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
         Left            =   8970
         TabIndex        =   221
         Top             =   4200
         Width           =   885
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
         Left            =   180
         TabIndex        =   220
         Top             =   4005
         Width           =   2760
      End
      Begin VB.Line Line3 
         X1              =   7395
         X2              =   11235
         Y1              =   4215
         Y2              =   4215
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DECLARO ESTAR DE PLENO ACORDO COM ESTE PEDIDO"
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
         Left            =   150
         TabIndex        =   219
         Top             =   3690
         Width           =   5430
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   4470
         TabIndex        =   218
         Top             =   4185
         Width           =   795
      End
      Begin VB.Line Line30 
         X1              =   2925
         X2              =   6915
         Y1              =   4200
         Y2              =   4200
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
         Left            =   7605
         TabIndex        =   175
         Top             =   2985
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
         Left            =   4515
         TabIndex        =   174
         Top             =   2985
         Width           =   1110
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prev. Entrega:"
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
         Left            =   255
         TabIndex        =   173
         Top             =   2955
         Width           =   1185
      End
      Begin VB.Shape Shape3 
         Height          =   645
         Left            =   165
         Top             =   2895
         Width           =   11205
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pagamento:"
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
         Left            =   5250
         TabIndex        =   172
         Top             =   1260
         Width           =   1845
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   4
         Left            =   8700
         Top             =   1500
         Width           =   2670
      End
      Begin VB.Shape Shape 
         Height          =   285
         Index           =   1
         Left            =   150
         Top             =   1230
         Width           =   11220
      End
      Begin VB.Label LabelCondicao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condições Pagamento:"
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
         Left            =   285
         TabIndex        =   171
         Top             =   1260
         Width           =   2235
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   3
         Left            =   5865
         Top             =   1500
         Width           =   2670
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   2
         Left            =   3045
         Top             =   1500
         Width           =   2670
      End
      Begin VB.Shape Shape 
         Height          =   1335
         Index           =   0
         Left            =   150
         Top             =   1500
         Width           =   2670
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
         Left            =   7545
         TabIndex        =   170
         Top             =   3255
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
         Left            =   210
         TabIndex        =   169
         Top             =   3255
         Width           =   1380
      End
      Begin VB.Line Line29 
         X1              =   8070
         X2              =   11340
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line28 
         X1              =   8070
         X2              =   11340
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line27 
         X1              =   8070
         X2              =   11340
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Line Line14 
         X1              =   8070
         X2              =   8070
         Y1              =   0
         Y2              =   1140
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   120
         X2              =   11340
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape Shape10 
         Height          =   1140
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
         TabIndex        =   69
         Top             =   315
         Width           =   855
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal:"
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
         Left            =   8640
         TabIndex        =   27
         Top             =   30
         Width           =   795
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Frete:"
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
         Left            =   8460
         TabIndex        =   26
         Top             =   585
         Width           =   975
      End
      Begin VB.Label lblInfGrupo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedido:"
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
         Index           =   0
         Left            =   8340
         TabIndex        =   25
         Top             =   855
         Width           =   1095
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   3
      Left            =   0
      Top             =   9345
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
         Left            =   7695
         TabIndex        =   79
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
         Left            =   7185
         TabIndex        =   78
         Top             =   90
         Width           =   405
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cor/Acab./Tecido"
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
         Left            =   5580
         TabIndex        =   77
         Top             =   90
         Width           =   1395
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         Left            =   8610
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
      Top             =   14580
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
Attribute VB_Name = "FrmRelPedPromi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalItens As Currency
Dim SubTotal As Currency

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
            Valor = (SubTotal + Relatorio.Recordset("ValorFrete")) - Relatorio.Recordset("Desconto")
'            TotalItens = 0
            SubTotal = 0
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
'    SubTotal = SubTotal + Relatorio.Recordset("VlrVdaDia")
            
     If VarValorPedidoVista = False Then
        SubTotal = SubTotal + Relatorio.Recordset("VlrVdaDia")
    Else
        SubTotal = SubTotal + Relatorio.Recordset("Valor")
    End If
           
End Sub
