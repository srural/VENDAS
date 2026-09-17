VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelCorrigeEst 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   75.671
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   2610
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1879
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   14
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
         TabIndex        =   15
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contagem"
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
         TabIndex        =   29
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN13"
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
         Left            =   5625
         TabIndex        =   27
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
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
         Left            =   8190
         TabIndex        =   25
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença R$"
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
         TabIndex        =   16
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição "
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
         Left            =   1110
         TabIndex        =   11
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   90
         TabIndex        =   10
         Top             =   450
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
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
         Left            =   9030
         TabIndex        =   6
         Top             =   660
         Width           =   825
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Width           =   585
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do Produto"
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
         Left            =   1425
         TabIndex        =   1
         Top             =   660
         Width           =   1845
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      Top             =   1665
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   397
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   370
         Campo           =   "CodPrd"
         Caption         =   ""
         MostrarSeRepetir=   0   'False
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
         Left            =   1215
         TabIndex        =   8
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   8160
         TabIndex        =   9
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
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
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   9120
         TabIndex        =   17
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Campo           =   "CodBarLoja"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   10065
         TabIndex        =   18
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "forDiferenca"
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
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   7200
         TabIndex        =   28
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   370
         Campo           =   "Est_Ideal"
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
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   5625
         TabIndex        =   30
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   370
         Campo           =   "CodBar"
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
      Begin VB.Line Line9 
         BorderStyle     =   6  'Inside Solid
         X1              =   0
         X2              =   11340
         Y1              =   195
         Y2              =   195
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
         TabIndex        =   3
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      Top             =   2790
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1588
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8520
         TabIndex        =   4
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
         TabIndex        =   5
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
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   1065
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField rpNome 
         Height          =   225
         Left            =   1110
         TabIndex        =   12
         Top             =   60
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   397
         Campo           =   "Descrição_Grupo"
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
      Begin ReportX.ReportField rpCodigo 
         Height          =   225
         Left            =   45
         TabIndex        =   13
         Top             =   60
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "CodGru"
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
      Height          =   300
      Index           =   0
      Left            =   0
      Top             =   2190
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField10 
         Height          =   225
         Index           =   0
         Left            =   6855
         TabIndex        =   19
         Top             =   30
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Campo           =   "forSubEstoque"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   225
         Index           =   0
         Left            =   10005
         TabIndex        =   20
         Top             =   30
         Width           =   1005
         _ExtentX        =   1773
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   225
         Index           =   1
         Left            =   8955
         TabIndex        =   31
         Top             =   45
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         Campo           =   "forDif"
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
      Begin VB.Label Label7 
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
         Index           =   0
         Left            =   5505
         TabIndex        =   21
         Top             =   30
         Width           =   945
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   300
      Index           =   1
      Left            =   0
      Top             =   2490
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField ReportField12 
         Height          =   225
         Left            =   6900
         TabIndex        =   22
         Top             =   30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "forTotalEstoque"
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
      Begin ReportX.ReportField ReportField13 
         Height          =   225
         Left            =   10005
         TabIndex        =   23
         Top             =   30
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "forTotalGeral"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   225
         Left            =   8685
         TabIndex        =   32
         Top             =   45
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Campo           =   "forTotalDif"
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   5850
         TabIndex        =   24
         Top             =   30
         Width           =   585
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   225
      Index           =   1
      Left            =   0
      Top             =   1440
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   397
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField ReportField14 
         Height          =   225
         Left            =   1215
         TabIndex        =   26
         Top             =   0
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   397
         Campo           =   "Descrição_Grupo"
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
         Borda           =   8
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   1890
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      Tipo            =   5
      Mostrar         =   0   'False
      Ordem           =   2
   End
End
Attribute VB_Name = "FrmRelCorrigeEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varSubTotal As Currency
Dim varTotalGeral As Currency
Dim varSubEstoque As Currency
Dim varTotalEstoque As Currency
Dim varEstAnt As Currency
Dim varDiferenca As Currency
Dim VarDif As Currency
Dim VarTotDif As Currency

Public Sub Config()
    Relatorio.Ativar
    Unload Me
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

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
    
    varSubTotal = 0
    varTotalGeral = 0
    varSubEstoque = 0
    varTotalEstoque = 0
    varEstAnt = 0
    VarDif = 0
    VarTotDif = 0
    varDiferenca = 0 ' (Relatorio.Recordset("CodBarLoja") * Relatorio.Recordset("Venda"))

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
        Case "forSubEstoque"
            Valor = varSubEstoque
            varSubEstoque = 0
        Case "forSubTotal"
            Valor = varSubTotal
            varSubTotal = 0
        Case "forTotalGeral"
            Valor = varTotalGeral
            varTotalGeral = 0
        Case "forTotalEstoque"
            Valor = varTotalEstoque
            varTotalEstoque = 0
        Case "forEstAnt"
            Valor = varEstAnt
        Case "forDiferenca"
            Valor = (Relatorio.Recordset("CodBarLoja") * Relatorio.Recordset("Custo"))
        Case "forDif"
            Valor = VarDif
        Case "forTotalDif"
            Valor = VarTotDif
    End Select
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    ' Indica que o grupo irá quebrar a cada novo codigo
    ' do pedido. Verifica se é formula para o grupo 1 ou
    ' para o grupo 2. Nesse caso o grupo 2 nao tem quebra
    ' apenas o grupo 1.
    
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset(rpCodigo.Campo)
    End If
    
    If Ordem = 2 Then
        Valor = Relatorio.Recordset("Descrição_Grupo")
    End If
    
End Sub


Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada impressão de registro na linha de detalhe,
    ' soma as qtdes e os totais para apresentar no rodapé
    ' do grupo
'    If VarLojaUnica = False Then
        If Not IsNull(Relatorio.Recordset("SubTotal")) Then
            varSubTotal = varSubTotal + Relatorio.Recordset("SubTotal")
            varTotalGeral = varTotalGeral + Relatorio.Recordset("SubTotal")
        End If
'    End If
    If Not IsNull(Relatorio.Recordset("Est_Ideal")) Then
        varSubEstoque = varSubEstoque + Relatorio.Recordset("Est_Ideal")
        varTotalEstoque = varTotalEstoque + Relatorio.Recordset("Est_Ideal")
'        varDiferenca = (Relatorio.Recordset("CodBarLoja") * Relatorio.Recordset("Venda"))
        VarDif = VarDif + Relatorio.Recordset("CodBarLoja")
        VarTotDif = VarTotDif + Relatorio.Recordset("CodBarLoja")
    End If
    
'    If FrmRelEst.Tag <> "" Then
'        varEstAnt = varEstAnt + Relatorio.Recordset("SubEstAnt")
'    End If
    
End Sub
