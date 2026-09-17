VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelVend 
   Caption         =   "Relatório de Vendedores"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   130.44
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   90
      TabIndex        =   0
      Top             =   5340
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
         Left            =   90
         TabIndex        =   24
         Top             =   60
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
         TabIndex        =   25
         Top             =   120
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
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   90
         TabIndex        =   11
         Top             =   450
         Width           =   675
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
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
         Left            =   900
         TabIndex        =   10
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
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
         Left            =   5520
         TabIndex        =   9
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro"
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
         Left            =   10290
         TabIndex        =   8
         Top             =   450
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
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
         Left            =   900
         TabIndex        =   7
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
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
         Left            =   9525
         TabIndex        =   5
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
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
         Left            =   900
         TabIndex        =   4
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
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
         Left            =   2610
         TabIndex        =   3
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.N.P.J./C.P.F."
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
         Left            =   6150
         TabIndex        =   2
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição/R.G."
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
         Left            =   9000
         TabIndex        =   1
         Top             =   900
         Width           =   1320
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      Top             =   1935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1323
      Begin ReportX.ReportField ReportField7 
         Height          =   240
         Left            =   900
         TabIndex        =   12
         Top             =   450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   423
         Campo           =   "Fone"
         Formato         =   "(@@)@@@@-@@@@"
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
      Begin ReportX.ReportField rpSoma 
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   -30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   423
         Campo           =   "Ent.CodEntidade"
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
      Begin ReportX.ReportField ReportField1 
         Height          =   240
         Left            =   900
         TabIndex        =   14
         Top             =   -30
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         Campo           =   "Ent.Nome"
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
      Begin ReportX.ReportField ReportField2 
         Height          =   240
         Left            =   5520
         TabIndex        =   15
         Top             =   -30
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   423
         Campo           =   "Endereco"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   240
         Left            =   10290
         TabIndex        =   16
         Top             =   -30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Campo           =   "Nro"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   240
         Left            =   900
         TabIndex        =   17
         Top             =   210
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         Campo           =   "Bairro"
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
      Begin ReportX.ReportField ReportField5 
         Height          =   240
         Left            =   5520
         TabIndex        =   18
         Top             =   210
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   423
         Campo           =   "Cidade"
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
      Begin ReportX.ReportField ReportField6 
         Height          =   240
         Left            =   9525
         TabIndex        =   19
         Top             =   210
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   423
         Campo           =   "Cep"
         Formato         =   "@@.@@@-@@"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   240
         Left            =   2580
         TabIndex        =   20
         Top             =   450
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   423
         Campo           =   "NomeVendedor"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   240
         Left            =   6150
         TabIndex        =   28
         Top             =   450
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Campo           =   "forCNPJ"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField12 
         Height          =   240
         Left            =   9000
         TabIndex        =   29
         Top             =   450
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   423
         Campo           =   "forInscricao"
         Caption         =   ""
         Formula         =   -1  'True
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
         DrawMode        =   14  'Copy Pen
         X1              =   90
         X2              =   11550
         Y1              =   690
         Y2              =   690
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
         TabIndex        =   21
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   3165
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8520
         TabIndex        =   22
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
         TabIndex        =   23
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
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   1590
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField ReportField9 
         Height          =   240
         Left            =   330
         TabIndex        =   27
         Top             =   30
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         Campo           =   "Cidade"
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
      Begin VB.Line Line4 
         X1              =   60
         X2              =   11400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   11400
         Y1              =   270
         Y2              =   270
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   2685
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      Tipo            =   5
      Ordem           =   2
      Begin ReportX.ReportField ReportField14 
         Height          =   240
         Left            =   9900
         TabIndex        =   30
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   423
         Campo           =   "forTotalCidade"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total por Cidade"
         Height          =   240
         Left            =   8370
         TabIndex        =   31
         Top             =   0
         Width           =   1410
      End
   End
   Begin ReportX.ReportSection CabGrupo 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   1
      Left            =   0
      Top             =   1245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField8 
         Height          =   240
         Left            =   60
         TabIndex        =   26
         Top             =   90
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         Campo           =   "EntVend.Nome"
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
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   30
         Y2              =   30
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   225
      Index           =   1
      Left            =   0
      Top             =   2940
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   397
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField13 
         Height          =   240
         Left            =   9900
         TabIndex        =   32
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   423
         Campo           =   "forTotalGeral"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Geral"
         Height          =   240
         Left            =   8370
         TabIndex        =   33
         Top             =   0
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmRelVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VarContador As Double
Public VarContadorGeral As Double

Public Sub Config()

    Relatorio.Ativar
    Unload Me
    
End Sub

Private Sub Label9_Click()

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
    'lneCab.X1 = 0
    'lneCab.X2 = Relatorio.LarguraUtil
    'lneRod.X1 = 0
    'lneRod.X2 = Relatorio.LarguraUtil
    VarContador = 0
    VarContadorGeral = 0
    
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

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 2 Then
        Valor = Relatorio.Recordset("Cidade")
        
    ElseIf Ordem = 1 Then
        Valor = Relatorio.Recordset("EntVend.Nome")
    End If
   
    
End Sub
Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    VarContador = VarContador + 1
    VarContadorGeral = VarContadorGeral + 1

End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        Case "forCNPJ"
            If Relatorio.Recordset("Classificacao") = 1 Then
                Valor = Format(Relatorio.Recordset("CPF"), "@@@.@@@.@@@-@@")
            ElseIf Relatorio.Recordset("Classificacao") = 2 Then
                Valor = Format(Relatorio.Recordset("CGC"), "@@.@@@.@@@/@@@@-@@")
            End If
                
        Case "forInscricao"
            If Relatorio.Recordset("Classificacao") = 1 Then
                Valor = Format(Relatorio.Recordset("RG"), "@@.@@@.@@@-@")
            ElseIf Relatorio.Recordset("Classificacao") = 2 Then
                Valor = Relatorio.Recordset("InscrEst")
            End If
                
        Case "forTotalCidade"
              Valor = VarContador
              VarContador = 0
        
        Case "forTotalGeral"
              Valor = VarContadorGeral
              
    
    End Select
End Sub

