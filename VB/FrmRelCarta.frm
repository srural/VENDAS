VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelCarta 
   ClientHeight    =   12810
   ClientLeft      =   1065
   ClientTop       =   6090
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   225.954
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   215.636
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   165
      TabIndex        =   0
      Top             =   1680
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Regua           =   -1  'True
      Titulo          =   "Domingo"
      TipoColunas     =   1
      HTMLUnico       =   -1  'True
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      Top             =   10200
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   1164
      Tipo            =   2
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   9720
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   847
      Tipo            =   7
      Mostrar         =   0   'False
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   9720
      Left            =   0
      Top             =   0
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   17145
      AutoEncolher    =   -1  'True
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpfDet 
         Height          =   9660
         Index           =   0
         Left            =   990
         TabIndex        =   1
         Top             =   4335
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   17039
         Linhas          =   28
         Campo           =   "Linha"
         Caption         =   ""
         TipoCampo       =   7
         Formula         =   -1  'True
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         Justificar      =   -1  'True
      End
      Begin ReportX.ReportField rpfArq 
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Top             =   2010
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   450
         Caption         =   "Ao(a)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfArq 
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   3
         Top             =   2355
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   450
         Caption         =   "Ilmo(a) Sr(a)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   360
         Left            =   4095
         TabIndex        =   4
         Top             =   1140
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   635
         Campo           =   "forData"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   270
         Left            =   975
         TabIndex        =   5
         Top             =   2835
         Width           =   6450
         _ExtentX        =   11377
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
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   270
         Left            =   975
         TabIndex        =   6
         Top             =   3445
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   476
         Campo           =   "Bairro"
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
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   270
         Left            =   975
         TabIndex        =   7
         Top             =   3140
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   476
         Campo           =   "Endereco"
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
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   270
         Left            =   5940
         TabIndex        =   8
         Top             =   3140
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         Campo           =   "Nro"
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
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   270
         Left            =   975
         TabIndex        =   9
         Top             =   3750
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   476
         Campo           =   "Cidade"
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
         ForeColor       =   -2147483630
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   13965
         Left            =   690
         Top             =   660
         Width           =   10455
      End
   End
End
Attribute VB_Name = "FrmRelCarta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sugestao As Currency
Dim LocArq As String
Dim CanalArq As String
Public CobJuros As String
Public VarCorrecao As Boolean

Public Sub Config(Arquivo As String)

    LocArq = Arquivo
    Relatorio.Ativar
    Unload Me
    
End Sub

Private Sub Form_Load()
Sugestao = 0
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    'With shpZebra
    '    .Left = 0
    '    .Top = 0
    '    .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
    '    .Height = Det.Height
    'End With
    
Sugestao = 0
LocArq = 0
CanalArq = 0
    
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        'shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
        'shpZebra.Visible = False
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
Dim strValor As String
Dim sLinha As String

    Select Case Campo
        Case "Linha"
        
            Open (Caminho & "\Aniversariante.txt") For Input As #1
            
            Do
              Line Input #1, sLinha
              strValor = strValor & Chr(10) & sLinha
            
            Loop Until Eof(1)
            Close #1
    
            Valor = (strValor) ', "@forJuros@", CobJuros)
    
        Case "forData"
            Valor = Day(Relatorio.Recordset("DtNasc")) & " de " & TraduzMes(Month(Relatorio.Recordset("DtNasc"))) & " de " & Year(Date)

    End Select
End Sub



