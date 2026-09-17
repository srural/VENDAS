VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelTXT2 
   ClientHeight    =   2160
   ClientLeft      =   1065
   ClientTop       =   6090
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
   ScaleHeight     =   3.81
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   21.14
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1111
      Tipo            =   2
      Begin ReportX.ReportField rpfNome 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   300
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   450
         Caption         =   "NomeArquivo"
         Alignment       =   2
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         Borda           =   14
         BackColor       =   12648447
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   840
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   423
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   13680
         TabIndex        =   2
         Top             =   0
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
      Begin VB.Line lneRod 
         X1              =   16110
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   210
      Left            =   0
      Top             =   630
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   370
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpfDet 
         Height          =   240
         Index           =   0
         Left            =   780
         TabIndex        =   1
         Top             =   0
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   423
         Campo           =   "Linha"
         Caption         =   ""
         Formula         =   -1  'True
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         TamanhoAuto     =   -1  'True
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   210
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   370
         Campo           =   "NumLinha"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         WordWrap        =   -1  'True
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1275
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   2
      Regua           =   -1  'True
      Escala          =   7
      Orientacao      =   2
      MargemEsquerda  =   1
      MargemDireita   =   1
      Titulo          =   "Arquivo TXT"
      HTMLUnico       =   -1  'True
   End
End
Attribute VB_Name = "FrmRelTXT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CanalArq As Integer
Private pLinha As String
Private pLin As Long

' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config(Arquivo As String)

    Dim Lin As Long
    
    ' Verifica se o arquivo existe.
    If dir(Arquivo) = "" Then Exit Sub
        
    rpfNome.Caption = Arquivo
    
    ' Conta o numero de linhas do arquivo para associar
    ' a propriedade NumeroRegistros que indicará o numero
    ' de linhas de detalhe a serem impressas.
    CanalArq = FreeFile
    pLin = 0
    Open Arquivo For Input As #CanalArq
    While Not EOF(CanalArq)
        Line Input #CanalArq, pLinha
        pLin = pLin + 1
    Wend
        
    Relatorio.NumeroRegistros = pLin
    
    ' Inicia a geração do relatório.
    Relatorio.Ativar
    
    ' Fecha o arquivo
    Close #CanalArq
    
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
    
    ' Mostra a linha do arquivo
    Select Case Campo
    
        Case "Linha": Valor = pLinha
        Case "NumLinha": Valor = pLin
        
    End Select
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)

    ' Posiciona no inicio do arquivo
    If Not Impressora Then
        pLin = 0
        Seek #CanalArq, 1
    End If
    
End Sub

Private Sub Relatorio_ValidarRegistro(Invalido As Boolean, Cancelar As Boolean)
    
    '  Faz a leitura de uma linha do arquivo texto.
    Line Input #CanalArq, pLinha
    pLin = pLin + 1
    
End Sub
