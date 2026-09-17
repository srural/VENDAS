VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form rlt_FlexGrid 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   510
   ClientTop       =   2715
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4.339
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   18.203
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      Top             =   1185
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   926
      Tipo            =   7
      Begin ReportX.ReportField rpf_Exemplo 
         Height          =   210
         Left            =   0
         TabIndex        =   4
         Top             =   225
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   370
         Caption         =   "Exemplo de Geração de Relatório a partir de um MSFlexGrid"
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      Top             =   750
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   767
      AutoEncolher    =   -1  'True
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpf_Det 
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   423
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
         ForeColor       =   -2147483630
         Borda           =   15
      End
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1323
      Tipo            =   2
      Begin ReportX.ReportField rpf_Cab 
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   495
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   423
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         Borda           =   15
      End
      Begin ReportX.ReportField rpf_Titulo 
         Height          =   360
         Left            =   2655
         TabIndex        =   3
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         Caption         =   "Titulo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         Borda           =   15
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   150
      TabIndex        =   1
      Top             =   1890
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   2
      Regua           =   -1  'True
      Escala          =   7
      MargemInferior  =   1
      Titulo          =   ""
      HTMLUnico       =   -1  'True
   End
End
Attribute VB_Name = "rlt_FlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Formulário modelo para impressao de MSFlexGrid
'  para uso com o Visual ReportX
'  © 2003 Vagner Pagotti
'
Option Explicit

Private grd As MSFlexGrid
Private LinhaAtual As Long

Public Titulo As String

Private Sub Form_Unload(Cancel As Integer)
    Set rlt_FlexGrid = Nothing
    Set grd = Nothing
End Sub

Public Property Set Grid(value As MSFlexGrid)
    Set grd = value
End Property

Public Sub Executar()

    Dim i As Integer
    Dim c As Integer
    Dim l As Integer
      
    If (grd Is Nothing) Then Exit Sub
    
    ' Desativa o grid
    grd.Redraw = False
    
    ' Define o numero de registros a serem impressos
    Relatorio.NumeroRegistros = grd.Rows - grd.FixedRows
      
    Det.AutoEncolher = grd.WordWrap
    Det.AutoExpandir = grd.WordWrap
    
    Det.AutoEncolher = True
    Det.AutoExpandir = True
    Cab.AutoEncolher = True
    Cab.AutoExpandir = True
    
    ' pega como sendo cabeçalho os dados
    ' na parte fixa do grid.
    ' ** Apenas a 1a. linha fixa.
    ' ** Para mais linhas é necessario fazer adaptação
    
    i = 0
    grd.row = 0
    
    For c = 0 To grd.Cols - 1
    
        grd.col = c
        
        ' Configura os campos do cabecalho
        If i > rpf_Cab.UBound Then
            Load rpf_Cab(i)
'            Load lne_Cab(i)
        End If
        With rpf_Cab(i)
            Set .Container = Cab
            .AlturaLivre = True
            .Visible = True
            .Mostrar = False
            SetFieldProps rpf_Cab(i), True
            .WordWrap = grd.WordWrap
            .TamanhoAuto = grd.WordWrap
            ' Tem que usar o alinhamento definido na coluna pq
            ' o FlexGrid mostra um texto na celula como sendo a direta
            ' mas mantem o valor da prop. de alinhamento da celula como
            ' sendo LEFT.
            FlexToAlignment rpf_Cab(i), grd.FixedAlignment(c)
            ' O reportField não aceita width < (2 pixels) ou negativo,
            ' entao se a coluna for invisivel (Width = 0) verificamos
            ' outra forma de informar as rotinas que o ReportField não
            ' ficará visivel.
            If grd.CellWidth <= 0 Then
                .Tag = "0"
            Else
                .Tag = ""
                .Width = grd.ColWidth(grd.col)
            End If
            .Caption = grd.TextMatrix(0, c)
            If i = grd.Cols - 1 Then
                .Borda = brdRetangulo
            Else
                .Borda = brdTopEsqBase
            End If
            .BordaCor = grd.GridColorFixed
            .BordaLargura = grd.GridLineWidth
        End With
          
          ' Configura os campos do detalhe
        If i > rpf_Det.UBound Then
            Load rpf_Det(i)
'            Load lne_Det(i)
        End If
        With rpf_Det(i)
            Set .Container = Det
            .Visible = True
            .Campo = c & ""
            .Formula = True
            .Mostrar = False
            SetFieldProps rpf_Det(i)
            .WordWrap = grd.WordWrap
            .TamanhoAuto = grd.WordWrap
            FlexToAlignment rpf_Det(i), grd.ColAlignment(c)
            If grd.CellWidth <= 0 Then
                .Tag = "0"
            Else
                .Tag = ""
                .Width = grd.ColWidth(grd.col)
            End If
            If i = grd.Cols - 1 Then
                .Borda = brdEsqDirBase
            Else
                .Borda = brdEsqBase
            End If
            .BordaCor = grd.GridColor
            .BordaLargura = grd.GridLineWidth
        End With
          
        i = i + 1
      
    Next
    
    ' Acerta o titulo
    Relatorio.Titulo = Titulo
    rpf_Titulo.Caption = Relatorio.Titulo
    
    ' Mostra o relatorio
    Relatorio.Ativar
    
    ' Reativa o grid
    grd.Redraw = True
    
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)

    ' Para os campos que ficam no detalhe, serao todos
    ' do tipo FORMULA = TRUE
    ' entao devolve o conteudo que estiver na celula do grid
    
    Valor = grd.TextMatrix(grd.row, Val(Campo))
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada linha impressa acerta o apontador de
    ' linha do grid.
    
    LinhaAtual = LinhaAtual + 1
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)

    Dim i As Integer
    Dim c As Integer
    Dim l As Integer
    
    Dim Posicao As Long
    Dim TopoCab As Long
    Dim Largura_Pagina As Long
    Dim FimColunas As Boolean
    
    If Not Impressora Then
        
        TopoCab = rpf_Cab(0).Top
        LinhaAtual = grd.FixedRows
        Largura_Pagina = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
            
        ' Acerta altura do cabecalho
        ' Se o grid tiver mais de uma linha fixa, configura apenas
        ' a primeira.
        ' ** A FAZER : Permitir mais de uma linha fixa do grid
        ' ** Atualmente campos a mais são cortados.
                                
        Posicao = 0
        grd.row = 0
        FimColunas = False
        
        For c = 0 To grd.Cols - 1
                    
            grd.col = c
            
            rpf_Cab(c).Mostrar = False
            rpf_Det(c).Mostrar = False
            
            With rpf_Det(c)
                If Not FimColunas And (Posicao + .Width < Largura_Pagina) Then
                    .Left = Posicao
                    If .Tag = "" Then
                        Posicao = Posicao + .Width
                        .Mostrar = True
                    End If
                Else
                    FimColunas = True
                End If
            End With
            
            With rpf_Cab(c)
                .Top = TopoCab
                .Left = rpf_Det(c).Left
                .Mostrar = (grd.FixedRows > 0) And rpf_Det(c).Mostrar
            End With
            
        Next
        
        ' centraliza o titulo
        rpf_Titulo.Move 0, 0, Posicao
        
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

Private Sub Relatorio_ValidarRegistro(Invalido As Boolean, Cancelar As Boolean)

    Dim c As Long
    
    grd.row = LinhaAtual
    
    ' Configura todas os campos de acordo com as colunas do grid.
    For c = 0 To grd.Cols - 1
        grd.col = c
        SetFieldProps rpf_Det(c)
    Next
        
End Sub

Private Sub FlexToAlignment(fld As ReportField, GridAligment As Long, Optional value As String = "")

    ' Funcao para traduzir alihamentos de texto
    ' do Grid para os alinhamentos suportados pelo ReportField
    Select Case GridAligment
        Case flexAlignLeftTop, flexAlignLeftCenter, flexAlignLeftBottom
            fld.Alignment = vbLeftJustify
        Case flexAlignCenterTop, flexAlignCenterCenter, flexAlignCenterBottom
            fld.Alignment = vbCenter
        Case flexAlignRightTop, flexAlignRightCenter, flexAlignRightBottom
            fld.Alignment = vbRightJustify
        Case flexAlignGeneral
            If IsNumeric(value) Or IsDate(value) Then
                fld.Alignment = vbRightJustify
            Else
                fld.Alignment = vbLeftJustify
            End If
    End Select
    
    Select Case GridAligment
        Case flexAlignGeneral, flexAlignLeftTop, flexAlignCenterTop, flexAlignRightTop
            fld.AlinhamentoVertical = alvTopo
        Case flexAlignLeftCenter, flexAlignCenterCenter, flexAlignRightCenter
            fld.AlinhamentoVertical = alvCentro
        Case flexAlignLeftBottom, flexAlignCenterBottom, flexAlignRightBottom
            fld.AlinhamentoVertical = alvBase
    End Select
    
End Sub

Private Sub SetFieldProps(fld As ReportField, Optional isFixedCell As Boolean = False)

    ' Copia as propriedades de cada celula
    ' para o ReportField poder apresentar
    ' igual
    
    On Error Resume Next
    
    fld.Font.Name = grd.CellFontName
    fld.Font.Size = grd.CellFontSize
    fld.Font.Bold = grd.CellFontBold
    fld.Font.Italic = grd.CellFontItalic
    fld.Font.Strikethrough = grd.CellFontStrikeThrough
    fld.Font.Underline = grd.CellFontUnderline
    
    If grd.CellForeColor = 0 Then
'        If isFixedCell Then
'            fld.ForeColor = grd.ForeColorFixed
        
'        Else
'            fld.ForeColor = grd.ForeColor
'        End If
    Else
        fld.ForeColor = grd.CellForeColor
    End If
    
    If grd.CellBackColor = 0 Then
        If isFixedCell Then
'            fld.BackColor = grd.BackColorFixed
'            fld.BackColor = grd.BackColor
        
        Else
'            fld.BackColor = grd.BackColor
        End If
    Else
'        fld.BackColor = grd.BackColor
    End If
    
    If Not grd.WordWrap Then fld.Height = grd.CellHeight
    
    On Error GoTo 0
    
End Sub

