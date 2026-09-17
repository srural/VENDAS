VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelIdVolume 
   Caption         =   "Identificador de Volume"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   104.511
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
      Tipo            =   2
      Mostrar         =   0   'False
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   3900
      Left            =   0
      Top             =   495
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   6879
      Begin ReportX.ReportField varNome 
         Height          =   255
         Left            =   1770
         TabIndex        =   2
         Top             =   1890
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   255
         Left            =   1470
         TabIndex        =   7
         Top             =   3120
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   450
         Campo           =   "Preco"
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   3120
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   450
         Campo           =   "Qtd"
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   255
         Left            =   9060
         TabIndex        =   9
         Top             =   3150
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   450
         Campo           =   "Produto"
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField VarCodNome 
         Height          =   255
         Left            =   9600
         TabIndex        =   10
         Top             =   1890
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField NomePrd 
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   450
         Campo           =   "Obs"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   2460
         Width           =   1605
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   1800
         X2              =   11010
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Image Logotipo 
         Height          =   1515
         Left            =   270
         Stretch         =   -1  'True
         Top             =   240
         Width           =   10875
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5340
         X2              =   7140
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   1500
         X2              =   3300
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   9000
         X2              =   11010
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1770
         X2              =   10980
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CODPRD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7590
         TabIndex        =   6
         Top             =   3150
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "VOLUME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3750
         TabIndex        =   5
         Top             =   3150
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº N.F."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   390
         TabIndex        =   4
         Top             =   3150
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1830
         Width           =   1605
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   4
         Height          =   3645
         Left            =   120
         Top             =   120
         Width           =   11265
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
      Height          =   420
      Left            =   0
      Top             =   4395
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      Tipo            =   7
      Mostrar         =   0   'False
   End
End
Attribute VB_Name = "FrmRelIdVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sugestao As Currency

Public Sub Config()

    Relatorio.Ativar
    
    Unload Me
    
End Sub




Private Sub Form_Load()
'On Error Resume Next

If Dir(App.Path & "\Logo.jpg") <> "" Then
    Logotipo = LoadPicture(App.Path & "\Logo.jpg")
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
    End Select
End Sub

