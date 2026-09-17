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
      Left            =   180
      OleObjectBlob   =   "FrmRelGrafico.frx":0000
      TabIndex        =   9
      Top             =   1980
      Width           =   6135
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   2170
      Tipo            =   2
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   1
         Left            =   6450
         TabIndex        =   4
         Top             =   780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   "Data"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Left            =   9210
         TabIndex        =   5
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Caption         =   "Serv.Dia"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Left            =   10215
         TabIndex        =   6
         Top             =   780
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         Caption         =   "Serv. Mês"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   609
         Caption         =   "Estatisitica"
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
      Begin ReportX.ReportField data1 
         Height          =   225
         Left            =   5490
         TabIndex        =   10
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
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
      Begin ReportX.ReportField Data2 
         Height          =   225
         Left            =   7155
         TabIndex        =   11
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         Caption         =   " Data"
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
      Begin ReportX.ReportField CabCli 
         Height          =   225
         Left            =   450
         TabIndex        =   13
         Top             =   630
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   397
         Caption         =   "Total Cliente"
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
      Begin ReportX.ReportField CabTicket 
         Height          =   225
         Left            =   450
         TabIndex        =   14
         Top             =   945
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   397
         Caption         =   "Ticket"
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
      Begin ReportX.ReportField rpCabSubTitulo 
         Height          =   270
         Left            =   1980
         TabIndex        =   15
         Top             =   270
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   476
         Caption         =   ""
         TipoCampo       =   7
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
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   225
         Index           =   3
         Left            =   7380
         TabIndex        =   16
         Top             =   765
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   397
         Caption         =   "Atend. Dia"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Left            =   8280
         TabIndex        =   18
         Top             =   765
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Caption         =   "Atend. Mês"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6930
         TabIndex        =   12
         Top             =   270
         Width           =   330
      End
      Begin VB.Image Image1 
         Height          =   3975
         Left            =   180
         Stretch         =   -1  'True
         Top             =   900
         Width           =   6135
      End
      Begin VB.Line Line1 
         X1              =   11340
         X2              =   6480
         Y1              =   990
         Y2              =   990
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   1500
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
      Top             =   1230
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   476
      Begin ReportX.ReportField rpfDet 
         Height          =   225
         Index           =   0
         Left            =   6435
         TabIndex        =   1
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Campo           =   "Data"
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Left            =   8265
         TabIndex        =   2
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         Campo           =   "MAtend"
         Formato         =   "##0"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Left            =   10215
         TabIndex        =   3
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         Campo           =   "MServicos"
         Formato         =   "##0"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Index           =   2
         Left            =   7380
         TabIndex        =   17
         Top             =   0
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   397
         Campo           =   "DAtend"
         Formato         =   "##0"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
         Index           =   3
         Left            =   9225
         TabIndex        =   19
         Top             =   0
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         Campo           =   "DServicos"
         Formato         =   "##0"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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

Public Rsql As String
Private Rs As Recordset





' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()

rpCabSubTitulo.Caption = NomeEmpresa


'    Dim Cn As String
        
    ' Uso do ADO nesse exemplo
'    Set Rs = New ADODB.Recordset
    
    ' Define a conexão. Pode ser usado uma conexao
    ' ja aberta pelo sistema.
'    Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\..\Exemplo.MDB"

    ' Abre o recordset com os dados
     Set Rs = Banco.OpenRecordset("SELECT ESTAT.Data, ESTAT.DAtend, ESTAT.MAtend,ESTAT.DServicos, ESTAT.MServicos,  ESTAT.TotCli,   ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(data1.Caption, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(Data2.Caption, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data")
    
    ' Preenche os dados do Chart com os valores
    Rs.MoveLast
    Rs.MoveFirst
    
    Dim Col As Integer, Row As Integer
    
    With MSChart1
        .ColumnCount = 4
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
