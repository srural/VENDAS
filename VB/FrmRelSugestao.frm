VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelSugestao 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   57.944
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   105
      Left            =   0
      Top             =   1395
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   185
      Tipo            =   5
      Ordem           =   1
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   840
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField9 
         Height          =   225
         Left            =   45
         TabIndex        =   24
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Campo           =   "Fabricante"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   270
         Left            =   750
         TabIndex        =   25
         Top             =   0
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   476
         Campo           =   "Nome"
         Caption         =   ""
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
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   6210
         Y1              =   270
         Y2              =   270
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   210
      TabIndex        =   0
      Top             =   2130
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1482
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   8
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
         Left            =   6270
         TabIndex        =   9
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodFabric."
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
         Left            =   4950
         TabIndex        =   27
         Top             =   495
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   10800
         TabIndex        =   21
         Top             =   480
         Width           =   510
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   420
         Y2              =   420
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
         Top             =   480
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
         Left            =   1020
         TabIndex        =   1
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embal."
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
         Left            =   6120
         TabIndex        =   20
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo"
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
         Left            =   9765
         TabIndex        =   19
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sugestão"
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
         Left            =   8535
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
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
         Left            =   7710
         TabIndex        =   17
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mínimo"
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
         Left            =   6870
         TabIndex        =   16
         Top             =   480
         Width           =   705
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1125
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Begin ReportX.ReportField ReportField7 
         Height          =   225
         Left            =   45
         TabIndex        =   6
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Campo           =   "CodPrd"
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
         Left            =   750
         TabIndex        =   7
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   397
         Campo           =   "Descrição_Produto"
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
      Begin ReportX.ReportField ReportField2 
         Height          =   240
         Left            =   10380
         TabIndex        =   10
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Campo           =   "forValor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField3 
         Height          =   240
         Left            =   9420
         TabIndex        =   11
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   423
         Campo           =   "Custo"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField4 
         Height          =   240
         Left            =   8535
         TabIndex        =   12
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   423
         Campo           =   "forSugestao"
         Formato         =   "Fixed"
         Caption         =   ""
         TipoCampo       =   6
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
      Begin ReportX.ReportField ReportField5 
         Height          =   240
         Left            =   7560
         TabIndex        =   13
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   423
         Campo           =   "Estoque"
         Formato         =   "Fixed"
         Caption         =   ""
         TipoCampo       =   6
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
      Begin ReportX.ReportField ReportField6 
         Height          =   240
         Left            =   6825
         TabIndex        =   14
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         Campo           =   "Mínimo"
         Formato         =   "Fixed"
         Caption         =   ""
         TipoCampo       =   6
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
      Begin ReportX.ReportField ReportField8 
         Height          =   240
         Left            =   6120
         TabIndex        =   15
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   423
         Campo           =   "Embalagem"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
         Height          =   225
         Left            =   4995
         TabIndex        =   26
         Top             =   0
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   397
         Campo           =   "CodBar"
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
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   10095
         Top             =   30
         Width           =   795
      End
      Begin VB.Line Line9 
         BorderStyle     =   3  'Dot
         DrawMode        =   14  'Copy Pen
         X1              =   90
         X2              =   11430
         Y1              =   240
         Y2              =   240
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
      Height          =   780
      Left            =   0
      Top             =   1500
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1376
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8910
         TabIndex        =   4
         Top             =   450
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
         Left            =   120
         TabIndex        =   5
         Top             =   420
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
         Height          =   345
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   609
         Tipo            =   6
         Ordem           =   1
         Begin ReportX.ReportField ReportField30 
            Height          =   240
            Left            =   9930
            TabIndex        =   22
            Top             =   60
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   423
            Campo           =   "forTotal"
            Caption         =   ""
            Formula         =   -1  'True
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
         Begin ReportX.ReportField ReportField36 
            Height          =   240
            Left            =   8100
            TabIndex        =   23
            Top             =   60
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   423
            Caption         =   "Total Compra"
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
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   60
            X2              =   11400
            Y1              =   30
            Y2              =   30
         End
      End
      Begin VB.Line lneRod 
         X1              =   10980
         X2              =   240
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "FrmRelSugestao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sugestao As Currency
Dim somaTotal As Currency

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

somaTotal = 0
    
    With shpZebra
        .Left = 0
        .Top = 0
        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
        .Height = Det.Height
    End With
    Sugestao = 0
    
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
        Case "forSugestao"
            Sugestao = ((Relatorio.Recordset("Mínimo")) - Relatorio.Recordset("Estoque"))
            If Sugestao > 0 Then
                Valor = Sugestao
            Else
                Valor = 0
            End If
        Case "forValor"
            Valor = Relatorio.Recordset("Custo") * Sugestao
        Case "forTotal"
            Valor = somaTotal
    End Select
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
     somaTotal = somaTotal + (Relatorio.Recordset("Custo") * Sugestao)
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("Fabricante")
    End If
    
End Sub

