VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmRelGrafico 
   ClientHeight    =   7050
   ClientLeft      =   2475
   ClientTop       =   645
   ClientWidth     =   11985
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
   ScaleHeight     =   12.435
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   21.14
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4605
      Left            =   1380
      OleObjectBlob   =   "Rel Grafico.frx":0000
      TabIndex        =   9
      Top             =   2025
      Width           =   6945
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1852
      Tipo            =   2
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   1
         Left            =   7845
         TabIndex        =   4
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   397
         Caption         =   "Data"
         Alignment       =   2
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
         Left            =   9390
         TabIndex        =   5
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Caption         =   "Abertura"
         Alignment       =   1
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
         Index           =   6
         Left            =   10395
         TabIndex        =   6
         Top             =   780
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Caption         =   "Fechamento"
         Alignment       =   1
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
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   195
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   609
         Caption         =   "Índices Bovespa - Janeiro de 1993"
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
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8685
         TabIndex        =   8
         Top             =   375
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Campo           =   "=Página [Pagina] de [Paginas]"
         Caption         =   ""
         TipoCampo       =   64
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
      Begin VB.Image Image1 
         Height          =   4065
         Left            =   210
         Stretch         =   -1  'True
         Top             =   810
         Width           =   7215
      End
      Begin VB.Line Line1 
         X1              =   11820
         X2              =   7665
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   1320
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   847
      Tipo            =   7
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   "Gráfico"
      Copias          =   2
      HTMLUnico       =   -1  'True
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   1050
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   476
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   0
         Left            =   7830
         TabIndex        =   1
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   397
         Campo           =   "Data"
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   1
         Left            =   9390
         TabIndex        =   2
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Campo           =   "Abertura"
         Formato         =   "0.000"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
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
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   5
         Left            =   10395
         TabIndex        =   3
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         Campo           =   "Fechamento"
         Formato         =   "0.000"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
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
   End
End
Attribute VB_Name = "FrmRelGrafico"
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
    Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\..\Exemplo.MDB"

    ' Abre o recordset com os dados
    Rs.Open "SELECT Data, Avg(Abertura) AS Abertura, Avg(Fechamento) AS Fechamento FROM IndicesBovespa WHERE Data BETWEEN #01/01/1993# AND #01/31/1993# GROUP BY Data ORDER BY Data", Cn, adOpenStatic
    
    ' Preenche os dados do Chart com os valores
    Rs.MoveLast
    Rs.MoveFirst
    
    Dim Col As Integer, Row As Integer
    
    With MSChart1
        .ColumnCount = 2
        For Col = 1 To .ColumnCount
            .Column = Col
            .ColumnLabel = Rs.Fields(Col).Name
        Next
        .RowCount = Rs.RecordCount
        For Row = 1 To .RowCount
            For Col = 1 To .ColumnCount
                .Column = Col
                .Row = Row
                .RowLabel = Format(Rs.Fields(0).value, "dd")
                .Data = Rs.Fields(Col).value
            Next
            Rs.MoveNext
        Next
        .Refresh
    End With
        
    Rs.MoveFirst
    
    ' Nesse exemplo pode se colocar um tamanho fixo no Image
    ' para que seja a metade do tamanho do objeto do
    ' MSChart. O ideal é que sempre o objeto do MSChart seja
    ' maior para não prejudicar a impressão pq na impressora
    ' a resolução é sempre maior que na tela.
    ' Se quiser utilize as rotinas abaixo e poderá fixar a
    ' largura ou a altura do Image e a rotina irá calcular
    ' a proporção correta.
        
    ' Caso utilize essas rotinas coloque a escala sempre em Twips.
    Relatorio.Escala = scmTwips
    AcertaTamanho Image1, MSChart1
    
    ' Copia os graficos para os objetos Image
    MSChart1.EditCopy
    Image1.Picture = Clipboard.GetData(vbCFBitmap)
    
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

Private Sub Form_Load()

End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    ' O componente trabalha no modo silencioso para
    ' erros. Ele dispara esse evento Erro e sai. Caso
    ' o seu relatório esteja iniciando e saindo sem9i9
    ' apresentar erro, verifique se você colocou algum
    ' código nesse evento.

    Rpx_MsgErro Numero
    
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


Private Sub AcertaTamanho(Destino As Object, Referencia As Object, Optional AncoraLargura As Boolean = True)

    Dim Proporcao As Double
    
    Proporcao = Referencia.Width / Referencia.Height
    
    If AncoraLargura Then
        Destino.Height = Destino.Width / Proporcao
    Else
        Destino.Width = Destino.Height * Proporcao
    End If
    
End Sub
