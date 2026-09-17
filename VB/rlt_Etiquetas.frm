VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form rlt_Etiquetas 
   ClientHeight    =   2190
   ClientLeft      =   1725
   ClientTop       =   4755
   ClientWidth     =   5790
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
   ScaleHeight     =   38.629
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   102.129
   Begin ReportX.ReportSection Det 
      Height          =   1440
      Left            =   90
      Top             =   105
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2540
      Begin ReportX.ReportField rpfDet 
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   503
         Campo           =   "Descrição_Produto"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   503
         Campo           =   "Descrição_Grupo"
         Caption         =   ""
         TipoCampo       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfDet 
         Height          =   285
         Index           =   2
         Left            =   4005
         TabIndex        =   3
         Top             =   630
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         Campo           =   "CodPrd"
         Caption         =   ""
         TipoCampo       =   6
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Label LblDia 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   4
         Top             =   945
         Width           =   1590
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   105
      TabIndex        =   0
      Top             =   1605
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Regua           =   -1  'True
      MargemEsquerda  =   5
      MargemDireita   =   5
      MargemSuperior  =   5
      MargemInferior  =   5
      Titulo          =   "Etiquetas"
      PaginaReal      =   -1  'True
      HTMLUnico       =   -1  'True
   End
End
Attribute VB_Name = "rlt_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs As ADODB.Recordset
Private Cn As String

Private pLinhaIni As Integer
Private pColunaIni As Integer
Private pInicio As Integer

Public Property Let Colunas(Valor As Integer)
    Relatorio.NumeroColunas = Valor
    Relatorio.TipoColunas = tpcHorizontal
End Property

Public Property Let LinhaInicio(Valor As Integer)
    ' permite que a impressao começe a partir de
    ' uma certa linha do formulario
    pLinhaIni = Valor
End Property

Public Property Let ColunaInicio(Valor As Integer)
    ' permite que a impressao começe a partir de
    ' uma certa coluna do formulario
    pColunaIni = Valor
End Property

Public Property Let Largura(Valor As Double)
    Det.Width = Valor
End Property

Public Property Let Altura(Valor As Double)
    Det.Height = Valor
End Property

Public Property Let MargemEsquerda(Valor As Double)
    Relatorio.MargemEsquerda = Valor
End Property

Public Property Let MargemSuperior(Valor As Double)
    Relatorio.MargemSuperior = Valor
End Property

Public Property Let MargemInferior(Valor As Double)
    Relatorio.MargemInferior = Valor
End Property

Private Sub ConfiguraCampos()

'    Dim RsE As ADODB.Recordset
'    Dim Ctr As ReportField
    
'    Set RsE = New ADODB.Recordset
    
'    RsE.CursorLocation = adUseClient
'    RsE.Open "SELECT * FROM CamposEtiqueta", Cn, adOpenStatic
    
'    While Not RsE.Eof
'        Set Ctr = rpfDet(RsE("Index"))
'        Ctr.Campo = RsE("Campo")
'        Ctr.Top = ScaleY(RsE("Top"), vbMillimeters, vbTwips)
'        Ctr.Left = ScaleX(RsE("Left"), vbMillimeters, vbTwips)
'        Ctr.Width = ScaleX(RsE("Width"), vbMillimeters, vbTwips)
'        If Ctr.Left + Ctr.Width > Det.Width Then
'            Ctr.Width = Det.Width - Ctr.Left
'        End If
'        Ctr.Formato = RsE("Formato") & ""
'        RsE.MoveNext
'    Wend
'
'    RsE.Close
'    Set RsE = Nothing
    
End Sub
' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()

    ' Uso do ADO nesse exemplo
    Set Rs = New ADODB.Recordset
    
    ' Define a conexão. Pode ser usado uma conexao
    ' ja aberta pelo sistema.
    Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Caminho & "\dados.MDB"
    
    ' Configura algumas propriedade do relatorio que são
    ' importantes para impressão desse tipo
    With Relatorio
        .PaginaReal = True
        .Escala = scmTwips
    End With
    
    ' Configura os campos na etiqueta
    ConfiguraCampos
    
    ' Abre o recordset com os dados, filtra apenas
    ' os municipios de Santa Catarina
    Rs.Open "SELECT PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo FROM (ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd) INNER JOIN GRU ON PRD.Grupo = GRU.CodGru", Cn, adOpenStatic
    
    ' Faz o calculo para saber quantas etiquetas deve pular antes de
    ' iniciar a impressao
    If pColunaIni > Relatorio.NumeroColunas Then pColunaIni = Relatorio.NumeroColunas
    pInicio = (pLinhaIni - 1) * Relatorio.NumeroColunas + (pColunaIni - 1)
    
    ' Conta o numero de registros a serem impressos porque
    ' irá colocar o relatorio em modo UNBOUND e somar o total
    ' de registros ao numero de etiquetas em branco que deverá
    ' pular antes de iniciar o relatorio.
    If Not Rs.Eof Then
        Rs.MoveLast
        Rs.MoveFirst
        Relatorio.NumeroRegistros = Rs.RecordCount + pInicio
    End If
    
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

Private Sub Form_Activate()
    LblDia.Caption = Date

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

    ' Verifica se deve ou não imprimir o valor do campo
    If pInicio > 0 Then
        Valor = Null
    Else
        Valor = Rs(Campo)
    End If
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    
    ' A cada etiqueta impressa em branco, diminui o contador
    If pInicio > 0 Then
        pInicio = pInicio - 1
    Else
        ' Como está em modo UNBOUND deve mover para o proximo
        ' registro manualmente. Pode-se usar a referencia do Recordset
        ' que foi associada ao Relatorio ou a propria variavel do Recordset.
        Rs.MoveNext
    End If
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)

    ' Quando se utiliza o Modo UNBOUND e vai usar um
    '  recordset associado, é preciso posiciona-lo no inicio
    ' nesse evento dessa forma, pq toda vez que o usuario
    ' selecionar outra impressora ou tamanho de papel
    ' será necessário gerar o relatorio novamente.
    If Not Impressora Then
        Rs.MoveFirst
    End If
    LblDia.Caption = Date
    
End Sub
