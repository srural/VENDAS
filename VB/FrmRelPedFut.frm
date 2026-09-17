VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPedFut 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   510
   ClientTop       =   2715
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.864
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   19.711
   Begin ReportX.ReportSection SecCabec 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      Top             =   2655
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   556
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField rpf_Cab 
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   370
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
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   10770
         X2              =   30
         Y1              =   270
         Y2              =   270
      End
   End
   Begin ReportX.ReportSection SecTitulo 
      Align           =   1  'Align Top
      Height          =   1395
      Left            =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2461
      Tipo            =   1
      Begin ReportX.ReportField TituloRelatorio 
         Height          =   270
         Left            =   5145
         TabIndex        =   15
         Top             =   1080
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpCabSubTitulo 
         Height          =   270
         Left            =   90
         TabIndex        =   17
         Top             =   1080
         Width           =   4920
         _ExtentX        =   8678
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
      End
      Begin VB.Image Logotipo 
         Height          =   960
         Left            =   30
         Stretch         =   -1  'True
         Top             =   60
         Width           =   11055
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      Top             =   3240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2143
      Tipo            =   7
      Begin ReportX.ReportSection SecTotal 
         Height          =   405
         Left            =   0
         Top             =   0
         Width           =   19080
         _ExtentX        =   33655
         _ExtentY        =   714
         Tipo            =   6
         AutoEncolher    =   -1  'True
         AutoExpandir    =   -1  'True
         Begin ReportX.ReportField Total 
            Height          =   240
            Left            =   7380
            TabIndex        =   10
            Top             =   60
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   423
            Caption         =   ""
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
         Begin ReportX.ReportField TotPed 
            Height          =   240
            Left            =   4410
            TabIndex        =   11
            Top             =   60
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   423
            Caption         =   ""
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Restante"
            Height          =   195
            Left            =   6240
            TabIndex        =   13
            Top             =   90
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Pedido"
            Height          =   195
            Left            =   3420
            TabIndex        =   12
            Top             =   90
            Width           =   900
         End
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   210
         Index           =   4
         Left            =   45
         TabIndex        =   14
         Tag             =   "N"
         Top             =   900
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
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   2970
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   476
      AutoEncolher    =   -1  'True
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpf_Det 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   370
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
      End
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   1290
         Top             =   30
         Width           =   1575
      End
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1260
      Left            =   0
      Top             =   1395
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2223
      Tipo            =   2
      Begin ReportX.ReportField Cabec1 
         Height          =   240
         Left            =   1050
         TabIndex        =   2
         Top             =   90
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   423
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
      Begin ReportX.ReportField Cabec2 
         Height          =   240
         Left            =   1050
         TabIndex        =   3
         Top             =   330
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   423
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
      Begin ReportX.ReportField Cabec3 
         Height          =   240
         Left            =   1050
         TabIndex        =   4
         Top             =   570
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   423
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
      Begin ReportX.ReportField Cabec4 
         Height          =   240
         Left            =   1050
         TabIndex        =   5
         Top             =   810
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   423
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.N.P.J"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   150
         Width           =   420
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   60
         Top             =   30
         Width           =   10725
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   150
      TabIndex        =   0
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
Attribute VB_Name = "FrmRelPedFut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Formulário modelo para impressao de ListView
'  para uso com o Visual ReportX
'  © 2004 Vagner Pagotti
'
Option Explicit

Private grd As ListView
Private LinhaAtual As Long

Public Titulo As String
Public GapColunas As Integer

Private Sub Form_Load()
    If Dir(Caminho & "\Logo.jpg") <> "" Then
        Logotipo = LoadPicture(Caminho & "\Logo.jpg")
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmRelPedFut = Nothing
    Set grd = Nothing
End Sub

Public Property Set Grid(value As ListView)
    Set grd = value
End Property

Public Sub Executar()

    Dim i As Integer
    Dim c As Integer
      
    If (grd Is Nothing) Then Exit Sub
    
    If GapColunas = 0 Then GapColunas = 60
    
    ' Define o numero de registros a serem impressos
    Relatorio.NumeroRegistros = grd.ListItems.Count
    
    i = 0
            
        Det.AutoEncolher = True
        Det.AutoExpandir = True
        Cab.AutoEncolher = True
        Cab.AutoExpandir = True
        
        
        For c = 0 To grd.ColumnHeaders.Count - 1
        
            ' Configura os campos do cabecalho
            If i > rpf_Cab.UBound Then
                Load rpf_Cab(i)
            End If
            With rpf_Cab(i)
                Set .Container = Cab
                .Visible = True
                .Mostrar = False
                CopyFont rpf_Cab(i), grd.Font
                .Font.Bold = True
                .WordWrap = False
                .TamanhoAuto = False
                ListViewToAlignment rpf_Cab(i), grd.ColumnHeaders.Item(c + 1).Alignment
                ' O reportField não aceita width < (2 pixels) ou negativo,
                ' entao se a coluna for invisivel (Width = 0) verificamos
                ' outra forma de informar as rotinas que o ReportField não
                ' ficará visivel.
                If grd.ColumnHeaders.Item(c + 1).Width <= 0 Then
                    .Tag = "0"
                Else
                    .Tag = ""
                    .Width = grd.ColumnHeaders.Item(c + 1).Width
                End If
                .Caption = grd.ColumnHeaders.Item(c + 1).Text
            End With
              
              ' Configura os campos do detalhe
            If i > rpf_Det.UBound Then
                 Load rpf_Det(i)
            End If
            With rpf_Det(i)
                Set .Container = Det
                .Visible = True
                .Campo = c & ""
                .Formula = True
                .Mostrar = False
                CopyFont rpf_Det(i), grd.Font
                .WordWrap = False
                .TamanhoAuto = False
                ListViewToAlignment rpf_Det(i), grd.ColumnHeaders.Item(c + 1).Alignment
                If grd.ColumnHeaders.Item(c + 1).Width <= 0 Then
                    .Tag = "0"
                Else
                    .Tag = ""
                    .Width = grd.ColumnHeaders.Item(c + 1).Width
                End If
            End With
              
            i = i + 1
          
        Next
    ' Acerta o titulo
'    Relatorio.Titulo = Titulo
'    rpf_Titulo.Caption = Relatorio.Titulo
    
    ' Mostra o relatorio
    Relatorio.Ativar
    
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
         On Error Resume Next

    ' Para os campos que ficam no detalhe, serao todos
    ' do tipo FORMULA = TRUE
    ' entao devolve o conteudo que estiver na coluna
    ' A primeira coluna fica no ListItem e as restantes no
    ' SubItens, entao verifica antes de devolver o valor
        If grd.ListItems.Item(LinhaAtual).Checked = True Then
        
            If Campo = "0" Then
                Valor = grd.ListItems.Item(LinhaAtual).Text
            Else
                Valor = grd.ListItems.Item(LinhaAtual).SubItems(Val(Campo))
            End If
        End If
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada linha impressa acerta o apontador de
    ' linha do listview
    
    LinhaAtual = LinhaAtual + 1
    
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        shpZebra.Visible = False
    End If
       
End Sub


Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)

    Dim c As Integer
    
    Dim Posicao As Long
    Dim TopoCab As Long
    Dim Largura_Pagina As Long
    Dim FimColunas As Boolean
    
    With shpZebra
        .Left = 0
        .Top = 0
        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
        .Height = Det.Height
    End With
    
    
    If Not Impressora Then
        
        TopoCab = rpf_Cab(0).Top
        LinhaAtual = 1
        Largura_Pagina = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
            
        ' Acerta altura do cabecalho
        ' ** A FAZER : Permitir mais de uma linha de cabeçalho
        ' ** Atualmente campos a mais são cortados.
                                
        Posicao = 0
        FimColunas = False
        
        For c = 0 To grd.ColumnHeaders.Count - 1
                    
            rpf_Cab(c).Mostrar = False
            rpf_Det(c).Mostrar = False
            
            With rpf_Det(c)
                If Not FimColunas And (Posicao + .Width < Largura_Pagina) Then
                    .Left = Posicao
                    If .Tag = "" Then
                        Posicao = Posicao + .Width + GapColunas
                        .Mostrar = True
                    End If
                Else
                    FimColunas = True
                End If
            End With
            
            With rpf_Cab(c)
                .Top = TopoCab
                .Left = rpf_Det(c).Left
                .Mostrar = rpf_Det(c).Mostrar
            End With
        Next
        
        ' centraliza o titulo
 '       rpf_Titulo.Move 0, 0, Posicao
        
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

Private Sub ListViewToAlignment(fld As ReportField, GridAligment As Long)

    ' Funcao para traduzir alihamentos de texto
    ' para os alinhamentos suportados pelo ReportField
    Select Case GridAligment
        Case lvwColumnLeft
            fld.Alignment = vbLeftJustify
        Case lvwColumnCenter
            fld.Alignment = vbCenter
        Case lvwColumnRight
            fld.Alignment = vbRightJustify
    End Select
    
End Sub

Private Function CopyFont(rpf As ReportField, ByVal fontSource As StdFont)
    
    With rpf.Font
        .Name = fontSource.Name
        .Size = fontSource.Size
        .Bold = fontSource.Bold
        .Italic = fontSource.Italic
        .Underline = fontSource.Underline
        .Strikethrough = fontSource.Strikethrough
    End With
    
End Function

