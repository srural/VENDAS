VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelPclData 
   Caption         =   "Relatório de Contas a Receber / Recebidas"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   85.196
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   90
      TabIndex        =   0
      Top             =   3285
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   1095
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1931
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
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   5340
         TabIndex        =   6
         Top             =   720
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Caption         =   "Pedido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   7920
         TabIndex        =   7
         Top             =   720
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Valor"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   210
         Left            =   9120
         TabIndex        =   8
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Condicao"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Top             =   330
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
      Begin ReportX.ReportField ReportField17 
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   720
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   "Cod. Ent."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   210
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   370
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   6030
         TabIndex        =   12
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Caption         =   "Nro. Pcl."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField24 
         Height          =   210
         Left            =   6840
         TabIndex        =   13
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Caption         =   "Dt. Pgto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField25 
         Height          =   210
         Left            =   10260
         TabIndex        =   14
         Top             =   720
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "Valor Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   990
         Y2              =   990
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   240
      Index           =   2
      Left            =   0
      Top             =   1500
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   5370
         TabIndex        =   20
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Campo           =   "Pedido"
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
         Left            =   7950
         TabIndex        =   21
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   370
         Campo           =   "Valor"
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
         Left            =   10350
         TabIndex        =   22
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   370
         Campo           =   "VlrPago"
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
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   990
         TabIndex        =   23
         Top             =   0
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   370
         Campo           =   "Ent.Nome"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   6060
         TabIndex        =   24
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Campo           =   "NroPcl"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   6870
         TabIndex        =   25
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "DtPgto"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   9150
         TabIndex        =   26
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "Condicao"
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
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   60
         TabIndex        =   27
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "CodEntidade"
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
      Top             =   2430
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
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   0
      Left            =   0
      Top             =   1095
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField12 
         Height          =   240
         Left            =   60
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "DtVenc"
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
      Begin VB.Line Line3 
         X1              =   60
         X2              =   11400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   11400
         Y1              =   90
         Y2              =   90
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   1740
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField27 
         Height          =   210
         Left            =   7980
         TabIndex        =   16
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forTotal"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   210
         Left            =   9180
         TabIndex        =   17
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forTotalPgto"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField36 
         Height          =   210
         Left            =   6930
         TabIndex        =   19
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Caption         =   "SubTotal:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line4 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   7950
         X2              =   9060
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line9 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   9180
         X2              =   10290
         Y1              =   30
         Y2              =   30
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   1
      Left            =   0
      Top             =   2025
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField ReportField35 
         Height          =   210
         Left            =   6930
         TabIndex        =   18
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Caption         =   "Total Geral:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   210
         Left            =   7980
         TabIndex        =   28
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forTotalG"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   9180
         TabIndex        =   29
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forTotalPgtoG"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         X1              =   60
         X2              =   11400
         Y1              =   270
         Y2              =   270
      End
   End
End
Attribute VB_Name = "FrmRelPclData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalValor As Currency
Dim TotalValorG As Currency
Dim TotalValorPgto As Currency
Dim TotalValorPgtoG As Currency

Public Sub Config()

    Relatorio.Ativar
    
    ' Retira o formulário de relatorio da memória
    Unload Me
    
End Sub

Private Sub Form_Load()

TotalValor = 0
TotalValorG = 0
TotalValorPgto = 0
TotalValorPgtoG = 0

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
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
    Select Case Campo
        Case "forTotal"
            Valor = TotalValor
            TotalValor = 0
        
        Case "forTotalG"
            Valor = TotalValorG
        
        Case "forTotalPgto"
            Valor = TotalValorPgto
            TotalValorPgto = 0
    
        Case "forTotalPgtoG"
            Valor = TotalValorPgtoG
    End Select

        Debug.Print Campo
        
End Sub


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    ' Indica que o grupo irá quebrar a cada novo codigo
    ' do pedido. Verifica se é formula para o grupo 1 ou
    ' para o grupo 2. Nesse caso o grupo 2 nao tem quebra
    ' apenas o grupo 1.
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("DtVenc")
    End If
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada impressão de registro na linha de detalhe,
    ' soma as qtdes e os totais para apresentar no rodapé
    ' do grupo
        
    TotalValor = TotalValor + Relatorio.Recordset("Valor")
    TotalValorG = TotalValorG + Relatorio.Recordset("Valor")
    TotalValorPgto = TotalValorPgto + Relatorio.Recordset("VlrPago")
    TotalValorPgtoG = TotalValorPgtoG + Relatorio.Recordset("VlrPago")
    
End Sub
