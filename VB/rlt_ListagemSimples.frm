VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.3#0"; "REPORTX.OCX"
Begin VB.Form rlt_ListagemSimples 
   ClientHeight    =   2925
   ClientLeft      =   510
   ClientTop       =   3855
   ClientWidth     =   11880
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
   ScaleHeight     =   51.594
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      Top             =   1245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1270
      Tipo            =   7
      Begin ReportX.ReportSection ReportSection1 
         Height          =   465
         Left            =   45
         Top             =   -90
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   820
         Tipo            =   6
         Ordem           =   1
      End
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   7560
         TabIndex        =   6
         Top             =   60
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Campo           =   "=Página [Pagina] de [Paginas]"
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
         TabIndex        =   7
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
         X1              =   10200
         X2              =   240
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   975
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      Begin ReportX.ReportField Rpfdet 
         Height          =   195
         Index           =   2
         Left            =   6300
         TabIndex        =   8
         Top             =   0
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   344
         Campo           =   "cgc"
         Formato         =   "##,###,###/####-##"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Rpfdet 
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   15
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Campo           =   "CodEntidade"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Rpfdet 
         Height          =   225
         Index           =   1
         Left            =   1530
         TabIndex        =   5
         Top             =   0
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   397
         Campo           =   "Nome"
         Formato         =   "0.000"
         Caption         =   ""
         TipoCampo       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8595
         Y1              =   225
         Y2              =   225
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
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1720
      Tipo            =   2
      Begin ReportX.ReportField rpfCab 
         Height          =   345
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   180
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         Caption         =   "Relatório de Entidades"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   1
         Left            =   225
         TabIndex        =   2
         Top             =   675
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   397
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   4
         Left            =   6255
         TabIndex        =   9
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   397
         Caption         =   "Cnpj"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line lneCab 
         X1              =   10200
         X2              =   240
         Y1              =   915
         Y2              =   915
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   300
      TabIndex        =   0
      Top             =   2160
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Regua           =   -1  'True
      Titulo          =   "Índices Bovespa"
      LarguraPapel    =   200
      AlturaPapel     =   150
      Registrado      =   0   'False
      HTMLUnico       =   -1  'True
   End
End
Attribute VB_Name = "rlt_ListagemSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs As ADODB.Recordset

' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()

    Dim Cn As String
        
    ' Uso do ADO nesse exemplo
    Set Rs = New ADODB.Recordset
    
    ' Define a conexão. Pode ser usado uma conexao
    ' ja aberta pelo sistema.
    Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\dados.mdb"

    ' Abre o recordset com os dados
    Rs.Open "SELECT * FROM Ent ORDER BY nome", Cn, adOpenStatic
    
    '  Associa o recordset ao relatorio
    Set Relatorio.Recordset = Rs
    ' Inicia a geração do relatório.
        
    Relatorio.Ativar
    
    ' Fecha o recordset
    Rs.Close
    Set Rs = Nothing
    
    ' Retira o formulário de relatorio da memória
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    ' O componente trabalha no modo silencioso para
    ' erros. Ele dispara esse evento Erro e sai. Caso
    ' o seu relatório esteja iniciando e saindo sem
    ' apresentar erro, verifique se você colocou algum
    ' código nesse evento.

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
    lneCab.X1 = 0
    lneCab.X2 = Relatorio.LarguraUtil
    lneRod.X1 = 0
    lneRod.X2 = Relatorio.LarguraUtil
    
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
'        shpZebra.Visible = Not shpZebra.Visible
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


