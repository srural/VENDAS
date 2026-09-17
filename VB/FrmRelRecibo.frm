VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelRecibo 
   Caption         =   "Relatório de CFO"
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
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   5400
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   9525
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField rpExtenso 
         Height          =   270
         Left            =   150
         TabIndex        =   4
         Top             =   4080
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   476
         Campo           =   "forExtenso"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField rpNome 
         Height          =   270
         Left            =   1650
         TabIndex        =   3
         Top             =   1860
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpTotal 
         Height          =   330
         Left            =   8940
         TabIndex        =   2
         Top             =   1350
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         Campo           =   "forTotal"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
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
      Begin ReportX.ReportField ReportField10 
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   3780
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         Caption         =   "Recebi da Empresa acima descrita a importância de:"
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
      Begin ReportX.ReportField rpTitulo 
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   390
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   1085
         Caption         =   "RECIBO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField36 
         Height          =   270
         Left            =   570
         TabIndex        =   9
         Top             =   4950
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Nro. da Nota"
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
      Begin ReportX.ReportField ReportField37 
         Height          =   270
         Left            =   2370
         TabIndex        =   10
         Top             =   4950
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Cod. Pedido"
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
      Begin ReportX.ReportField ReportField38 
         Height          =   270
         Left            =   4230
         TabIndex        =   11
         Top             =   4950
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Data"
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
      Begin ReportX.ReportField ReportField39 
         Height          =   270
         Left            =   6240
         TabIndex        =   12
         Top             =   4950
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Total"
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
      Begin ReportX.ReportField rpEndereco 
         Height          =   270
         Left            =   1650
         TabIndex        =   14
         Top             =   2190
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpCidade 
         Height          =   270
         Left            =   1650
         TabIndex        =   15
         Top             =   2520
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpEstado 
         Height          =   270
         Left            =   6240
         TabIndex        =   16
         Top             =   2520
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpCep 
         Height          =   270
         Left            =   1650
         TabIndex        =   17
         Top             =   2850
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpCgc 
         Height          =   270
         Left            =   1650
         TabIndex        =   18
         Top             =   3180
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   476
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
      Begin ReportX.ReportField rpCnpj 
         Height          =   270
         Left            =   4830
         TabIndex        =   19
         Top             =   3180
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   476
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
      Begin ReportX.ReportField ReportField1 
         Height          =   270
         Left            =   150
         TabIndex        =   20
         Top             =   1860
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Nome"
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
         Height          =   270
         Left            =   150
         TabIndex        =   21
         Top             =   2190
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Endereço"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   270
         Left            =   150
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Cidade"
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
         Left            =   150
         TabIndex        =   23
         Top             =   2850
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Cep"
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
      Begin ReportX.ReportField ReportField5 
         Height          =   270
         Left            =   150
         TabIndex        =   24
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
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
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   270
         Left            =   4110
         TabIndex        =   25
         Top             =   3180
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   476
         Caption         =   "Cnpj"
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
      Begin ReportX.ReportField ReportField7 
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   4410
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         Caption         =   "Referente a:"
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
      Begin VB.Line Line10 
         BorderStyle     =   2  'Dash
         X1              =   150
         X2              =   11280
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   150
         X2              =   11280
         Y1              =   5220
         Y2              =   5220
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillColor       =   &H00CECECE&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   8820
         Top             =   1230
         Width           =   2625
      End
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   5400
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Ordem           =   1
      Begin ReportX.ReportField ReportField14 
         Height          =   240
         Left            =   2370
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
      Top             =   5745
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1402
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField rpNome2 
         Height          =   270
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   2
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
      Begin VB.Line Line2 
         X1              =   60
         X2              =   5550
         Y1              =   210
         Y2              =   210
      End
   End
End
Attribute VB_Name = "FrmRelRecibo"
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

    
    rpNome.Campo = Tabela & ".Nome"
    rpNome2.Campo = "for" & Tabela & ".Nome"
    
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

