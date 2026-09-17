VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelRecibo2 
   Caption         =   "Recibo"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   194.204
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   150
      TabIndex        =   0
      Top             =   6150
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   9030
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Ordem           =   1
      Begin ReportX.ReportField ReportField14 
         Height          =   240
         Left            =   2370
         TabIndex        =   1
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "CodPed"
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
      Begin ReportX.ReportField ReportField15 
         Height          =   240
         Left            =   4230
         TabIndex        =   2
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "DtFat"
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
      Begin ReportX.ReportField ReportField16 
         Height          =   240
         Left            =   6210
         TabIndex        =   3
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "calTotal"
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
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   795
      Index           =   0
      Left            =   0
      Top             =   9375
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1402
      Tipo            =   7
      Ordem           =   1
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2064
      Tipo            =   3
      Ordem           =   1
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   7860
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   13864
      Tipo            =   2
      Begin ReportX.ReportField rpNome 
         Height          =   285
         Left            =   3225
         TabIndex        =   4
         Top             =   2760
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField rpTitulo 
         Height          =   450
         Left            =   90
         TabIndex        =   5
         Top             =   675
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   794
         Caption         =   "Folha de Pagamento Individual (Recibo)"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpcomissao 
         Height          =   270
         Left            =   8535
         TabIndex        =   6
         Top             =   3705
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField rpTotal 
         Height          =   270
         Left            =   8550
         TabIndex        =   7
         Top             =   6090
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   270
         Left            =   750
         TabIndex        =   8
         Top             =   3345
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   476
         Caption         =   "C R É D I T O S"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   270
         Left            =   1200
         TabIndex        =   9
         Top             =   3705
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   476
         Caption         =   "Comissão e Serviços"
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
         Height          =   270
         Left            =   750
         TabIndex        =   10
         Top             =   4110
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   476
         Caption         =   "Total dos Créditos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   270
         Left            =   750
         TabIndex        =   11
         Top             =   4740
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   476
         Campo           =   "D E S C O N T O S"
         Caption         =   "Incr. Est."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   255
         Index           =   0
         Left            =   795
         TabIndex        =   12
         Top             =   7035
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   450
         Caption         =   "Recebi em   ____/_______/________"
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
      Begin ReportX.ReportField rpVale 
         Height          =   270
         Left            =   8535
         TabIndex        =   13
         Top             =   5145
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField ReportField8 
         Height          =   270
         Left            =   1200
         TabIndex        =   14
         Top             =   5145
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   476
         Caption         =   "Vale"
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
      Begin ReportX.ReportField rpTotalDeb 
         Height          =   270
         Left            =   8535
         TabIndex        =   15
         Top             =   5550
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   270
         Left            =   750
         TabIndex        =   16
         Top             =   5550
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   476
         Caption         =   "Total dos Débitos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   750
         TabIndex        =   17
         Top             =   6090
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   476
         Caption         =   "T O T A L   A   R E C E B E R"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
      End
      Begin ReportX.ReportField rpPeriodo 
         Height          =   255
         Left            =   6825
         TabIndex        =   18
         Top             =   1365
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   450
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField rpTotalCred 
         Height          =   270
         Left            =   8550
         TabIndex        =   19
         Top             =   4110
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin VB.Line Line1 
         X1              =   6015
         X2              =   10650
         Y1              =   7260
         Y2              =   7260
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   750
         X2              =   10605
         Y1              =   3165
         Y2              =   3165
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   4350
         X2              =   11460
         Y1              =   1185
         Y2              =   1185
      End
   End
End
Attribute VB_Name = "FrmRelRecibo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SomaValor As Currency
Public FunRural As Boolean
Public SqlCondicao As String
Public Tabela As String
Dim NomeAnt As String

Public Sub Config()

    
'    rpNome.Campo = Tabela & ".Nome"
'    rpNome2.Campo = "for" & Tabela & ".Nome"
    
    Relatorio.Ativar
        
    Unload Me
    
End Sub

Private Sub Form_Load()

SomaValor = 0

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    SomaValor = 0
    
    
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

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
        
    SomaValor = SomaValor + Relatorio.Recordset("Total")
    rpTotal.Caption = "R$ " & Format(SomaValor, "##,#0.00")
    
End Sub


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset(Tabela & ".Nome")
    End If
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
Dim Completa As String
Dim valorTemp As Currency
Completa = " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
    Select Case Campo
        Case "forExtenso"
            NomeAnt = Relatorio.Recordset(Tabela & ".Nome")
            varPesquisa(2).MoveFirst
            varPesquisa(2).FindFirst IIf((Tabela = "Ent"), "Entidade", "Vendedor") & " = " & Relatorio.Recordset(Tabela & ".CodEntidade")
            
            If FunRural = True Then
                valorTemp = varPesquisa(2)("SomaDeTotal") - (varPesquisa(2)("SomaDeTotal") * "0,023")
            Else
                valorTemp = varPesquisa(2)("SomaDeTotal")
            End If
            
            Valor = "((" & Format(PassaExtenso(CStr(valorTemp)), ">") & "))" & Completa
        Case "forCpf"
            Valor = Format(Relatorio.Recordset(Tabela & ".Cpf"), "@@@.@@@.@@@-@@")
        Case "forEmpresa"
            Valor = NomeEmpresa
        Case "forTotal"
            varPesquisa(2).MoveFirst
            varPesquisa(2).FindFirst IIf((Tabela = "Ent"), "Entidade", "Vendedor") & " = " & Relatorio.Recordset(Tabela & ".CodEntidade")
            
            If FunRural = True Then
                valorTemp = varPesquisa(2)("SomaDeTotal") - (varPesquisa(2)("SomaDeTotal") * "0,023")
            Else
                valorTemp = varPesquisa(2)("SomaDeTotal")
            End If
            
            If varPesquisa(2).NoMatch = False Then
                Valor = "R$ " & Format(valorTemp, "#,#0.00")
            End If
        Case "forDataExtenso"
            Valor = DataExtenso(Date)
        Case "calTotal"
            If FunRural = True Then
                Valor = Relatorio.Recordset("Total") - (Relatorio.Recordset("Total") * "0,023")
            Else
                Valor = Relatorio.Recordset("Total")
            End If
        Case ("for" & Tabela & ".Nome")
            Valor = NomeAnt
    End Select
End Sub

