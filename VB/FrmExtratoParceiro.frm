VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmExtratoParceiro 
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
      Left            =   180
      TabIndex        =   0
      Top             =   6000
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection CabRel 
      Align           =   1  'Align Top
      Height          =   1380
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2434
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   90
         TabIndex        =   16
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
         TabIndex        =   17
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
      Begin ReportX.ReportField RptParceiro 
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   990
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
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
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissao"
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
         Left            =   7080
         TabIndex        =   18
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Parceiro"
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
         Left            =   1020
         TabIndex        =   15
         Top             =   450
         Width           =   1530
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
         TabIndex        =   14
         Top             =   450
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   6405
         TabIndex        =   8
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
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
         Left            =   5535
         TabIndex        =   7
         Top             =   690
         Width           =   300
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
         Caption         =   "Data"
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
         TabIndex        =   3
         Top             =   675
         Width           =   390
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
         Left            =   1605
         TabIndex        =   2
         Top             =   690
         Width           =   1845
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   8040
         TabIndex        =   1
         Top             =   690
         Width           =   585
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   1485
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   90
         TabIndex        =   9
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "DataEmiss"
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
         Left            =   1590
         TabIndex        =   10
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
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
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   5085
         TabIndex        =   11
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Campo           =   "PrimeiroDeQtd"
         Formato         =   "Standard"
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
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   6045
         TabIndex        =   12
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "PrimeiroDeValor"
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
      Begin ReportX.ReportField RpEstoque 
         Height          =   210
         Left            =   7080
         TabIndex        =   13
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "PrimeiroDeComissao"
         Formato         =   "Standard"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   8055
         TabIndex        =   19
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   370
         Campo           =   "Nome"
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
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   90
         X2              =   11430
         Y1              =   210
         Y2              =   210
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
         TabIndex        =   4
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   1305
      Left            =   0
      Top             =   2010
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2302
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8520
         TabIndex        =   5
         Top             =   1005
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
         TabIndex        =   6
         Top             =   1005
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
      Begin ReportX.ReportSection RodGrupo 
         Height          =   885
         Index           =   1
         Left            =   -45
         Top             =   0
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   1561
         Tipo            =   6
         Ordem           =   1
         Begin ReportX.ReportField ReportField35 
            Height          =   225
            Left            =   4035
            TabIndex        =   24
            Top             =   30
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   397
            Caption         =   "Total"
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
            Height          =   195
            Left            =   5130
            TabIndex        =   25
            Top             =   45
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   344
            Campo           =   "forTotQtd"
            Formato         =   "Standard"
            Caption         =   ""
            Formula         =   -1  'True
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField10 
            Height          =   195
            Left            =   6075
            TabIndex        =   26
            Top             =   45
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   344
            Campo           =   "forTotValor"
            Formato         =   "Standard"
            Caption         =   ""
            Formula         =   -1  'True
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField11 
            Height          =   195
            Left            =   7155
            TabIndex        =   27
            Top             =   45
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   344
            Campo           =   "forTotComissao"
            Formato         =   "Standard"
            Caption         =   ""
            Formula         =   -1  'True
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField RptTaxa 
            Height          =   195
            Left            =   7155
            TabIndex        =   28
            Top             =   270
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   344
            Formato         =   "Standard"
            Caption         =   "Taxa"
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField RptTotal 
            Height          =   195
            Left            =   6975
            TabIndex        =   29
            Top             =   540
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   344
            Formato         =   "Standard"
            Caption         =   "Total"
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField12 
            Height          =   225
            Left            =   6075
            TabIndex        =   30
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   397
            Caption         =   "Tx Cartão"
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
            Left            =   6075
            TabIndex        =   31
            Top             =   540
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   397
            Caption         =   "Comissão"
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
         Begin VB.Line Line6 
            BorderStyle     =   6  'Inside Solid
            DrawMode        =   14  'Copy Pen
            X1              =   150
            X2              =   11490
            Y1              =   0
            Y2              =   0
         End
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   105
      Index           =   0
      Left            =   0
      Top             =   1380
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   185
      Tipo            =   3
      Ordem           =   1
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   45
         Y2              =   45
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   1725
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   5085
         TabIndex        =   21
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
         Campo           =   "forQtd"
         Formato         =   "Standard"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   6030
         TabIndex        =   22
         Top             =   45
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   344
         Campo           =   "forValor"
         Formato         =   "Standard"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   7110
         TabIndex        =   23
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
         Campo           =   "forComissao"
         Formato         =   "Standard"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmExtratoParceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VarQtd As Currency
Public VarValor As Currency
Public VarComiss As Currency

Public VarTotQtd As Currency
Public VarTotValor As Currency
Public VarTotComiss As Currency

Public Sub Config()

    Relatorio.Ativar
    
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)

    VarQtd = VarQtd + Relatorio.Recordset("PrimeiroDeQtd")
    VarValor = VarValor + Relatorio.Recordset("PrimeiroDeValor")
    VarComiss = VarComiss + (Relatorio.Recordset("PrimeiroDeValor") * (Relatorio.Recordset("PrimeiroDeComissao") / 100))
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
VarQtd = 0
VarValor = 0
VarComiss = 0

VarTotQtd = 0
VarTotValor = 0
VarTotComiss = 0

    With shpZebra
        .Left = 0
        .Top = 0
        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
        .Height = Det.Height
    End With
    
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
    If Campo = "forQtd" Then
        Valor = VarQtd
        VarTotQtd = VarTotQtd + VarQtd
        VarQtd = 0
    End If
    
    If Campo = "forValor" Then
        Valor = VarValor
        VarTotValor = VarTotValor + VarValor
        VarValor = 0
    End If
    
    If Campo = "forComissao" Then
        Valor = VarComiss
        VarTotComiss = VarTotComiss + VarComiss
        VarComiss = 0
    End If
    
    If Campo = "forTotQtd" Then
        Valor = VarTotQtd
    End If
    
    If Campo = "forTotValor" Then
        Valor = VarTotValor
    End If
    
    If Campo = "forTotComissao" Then
        Valor = VarTotComiss
    End If
    
End Sub


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("DataEmiss")
    End If
    
End Sub
