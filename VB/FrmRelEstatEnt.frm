VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelEstatEnt 
   Caption         =   "Relatório de Clientes mais frequentes"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   46.567
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   216.165
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   420
      TabIndex        =   0
      Top             =   1950
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1799
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   7
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
         Left            =   6600
         TabIndex        =   8
         Top             =   90
         Width           =   4815
         _ExtentX        =   8493
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
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   705
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Caption         =   "Código"
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
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   1035
         TabIndex        =   11
         Top             =   705
         Width           =   2565
         _ExtentX        =   4524
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   10035
         TabIndex        =   12
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Ultima Compra"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   210
         Left            =   8730
         TabIndex        =   13
         Top             =   705
         Width           =   1200
         _ExtentX        =   2117
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
      Begin ReportX.ReportField rpCabPeriodo 
         Height          =   225
         Left            =   60
         TabIndex        =   14
         Top             =   345
         Width           =   5415
         _ExtentX        =   9551
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
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   4500
         TabIndex        =   16
         Top             =   720
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Caption         =   "Cidade"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   6840
         TabIndex        =   17
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   370
         Caption         =   "UF"
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
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   7425
         TabIndex        =   18
         Top             =   720
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Caption         =   "Fone"
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
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   960
         Y2              =   960
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1020
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Campo           =   "Entidade"
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
         Left            =   1035
         TabIndex        =   5
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
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   10140
         TabIndex        =   6
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "ÚltimoDeDataEmiss"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
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
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   8715
         TabIndex        =   9
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "SomaDeTotal"
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
         Left            =   4500
         TabIndex        =   15
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   370
         Campo           =   "Cidade"
         Caption         =   "Cidade"
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
      Begin ReportX.ReportField ReportField13 
         Height          =   210
         Left            =   6840
         TabIndex        =   19
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   370
         Campo           =   "UF"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   7425
         TabIndex        =   20
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "Fone"
         Caption         =   "Cidade"
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
         Left            =   10950
         Top             =   30
         Width           =   75
      End
      Begin VB.Line Line9 
         DrawMode        =   14  'Copy Pen
         X1              =   90
         X2              =   11430
         Y1              =   225
         Y2              =   225
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
      Top             =   1275
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   953
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8880
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
End
Attribute VB_Name = "FrmRelEstatEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Config()

    Relatorio.Ativar
    
    Unload Me
    
End Sub


Private Sub Label29_Click()
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
    If Campo = "forPrazo" Then
        Valor = Relatorio.Recordset("Venda") + (Relatorio.Recordset("Venda") * (FormRelPrd.VarIndice / 100))
    End If
End Sub
