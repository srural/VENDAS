VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelConfEst 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   510
   ClientTop       =   2715
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   11.589
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   19.711
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      Top             =   1485
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1429
      Tipo            =   7
      Begin ReportX.ReportSection ReportSection1 
         Height          =   210
         Left            =   0
         Top             =   0
         Width           =   19080
         _ExtentX        =   33655
         _ExtentY        =   370
         Tipo            =   6
         AutoEncolher    =   -1  'True
         AutoExpandir    =   -1  'True
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   210
         Index           =   4
         Left            =   45
         TabIndex        =   5
         Tag             =   "N"
         Top             =   450
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
      Height          =   300
      Left            =   0
      Top             =   1185
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   529
      AutoEncolher    =   -1  'True
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpf_Det 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   2
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
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10770
         Y1              =   240
         Y2              =   240
      End
   End
   Begin ReportX.ReportSection Cab 
      Align           =   1  'Align Top
      Height          =   1185
      Left            =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2090
      Tipo            =   2
      Begin ReportX.ReportField rpf_Cab 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   870
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
      Begin ReportX.ReportField TituloRelatorio 
         Height          =   270
         Left            =   585
         TabIndex        =   3
         Top             =   90
         Width           =   10215
         _ExtentX        =   18018
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
      Begin ReportX.ReportField Cabec1 
         Height          =   240
         Left            =   585
         TabIndex        =   4
         Top             =   405
         Width           =   10200
         _ExtentX        =   17992
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
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   10785
         X2              =   45
         Y1              =   765
         Y2              =   765
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
Attribute VB_Name = "FrmRelConfEst"
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

Private Grd As ListView
Private LinhaAtual As Long

Public Titulo As String
Public GapColunas As Integer

Private Sub Form_Load()
'    If Dir(Caminho & "\Logo.jpg") <> "" Then
'        Logotipo = LoadPicture(Caminho & "\Logo.jpg")
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmRelPedFut = Nothing
    Set Grd = Nothing
End Sub

Public Property Set Grid(Value As ListView)
    Set Grd = Value
End Property

Public Sub Executar()

    Dim I As Integer
    Dim c As Integer
      
    If (Grd Is Nothing) Then Exit Sub
    
    If GapColunas = 0 Then GapColunas = 60
    
    ' Define o numero de registros a serem impressos
    Relatorio.NumeroRegistros = Grd.ListItems.Count
    
    I = 0
            
        Det.AutoEncolher = True
        Det.AutoExpandir = True
        Cab.AutoEncolher = True
        Cab.AutoExpandir = True
        
        
        For c = 0 To Grd.ColumnHeaders.Count - 1
        
            ' Configura os campos do cabecalho
            If I > rpf_Cab.UBound Then
                Load rpf_Cab(I)
            End If
            With rpf_Cab(I)
                Set .Container = Cab
                .Visible = True
                .Mostrar = False
                CopyFont rpf_Cab(I), Grd.Font
                .Font.Bold = True
                .WordWrap = False
                .TamanhoAuto = False
                ListViewToAlignment rpf_Cab(I), Grd.ColumnHeaders.Item(c + 1).Alignment
                ' O reportField não aceita width < (2 pixels) ou negativo,
                ' entao se a coluna for invisivel (Width = 0) verificamos
                ' outra forma de informar as rotinas que o ReportField não
                ' ficará visivel.
                If Grd.ColumnHeaders.Item(c + 1).Width <= 0 Then
                    .Tag = "0"
                Else
                    .Tag = ""
                    .Width = Grd.ColumnHeaders.Item(c + 1).Width
                End If
                .Caption = Grd.ColumnHeaders.Item(c + 1).Text
            End With
              
              ' Configura os campos do detalhe
            If I > rpf_Det.UBound Then
                 Load rpf_Det(I)
            End If
            With rpf_Det(I)
                Set .Container = Det
                .Visible = True
                .Campo = c & ""
                .Formula = True
                .Mostrar = False
                CopyFont rpf_Det(I), Grd.Font
                .WordWrap = False
                .TamanhoAuto = False
                ListViewToAlignment rpf_Det(I), Grd.ColumnHeaders.Item(c + 1).Alignment
                If Grd.ColumnHeaders.Item(c + 1).Width <= 0 Then
                    .Tag = "0"
                Else
                    .Tag = ""
                    .Width = Grd.ColumnHeaders.Item(c + 1).Width
                End If
            End With
              
            I = I + 1
          
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
        If Grd.ListItems.Item(LinhaAtual).Checked = True Then
        
            If Campo = "0" Then
                Valor = Grd.ListItems.Item(LinhaAtual).Text
            Else
                Valor = Grd.ListItems.Item(LinhaAtual).SubItems(Val(Campo))
            End If
        End If
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada linha impressa acerta o apontador de
    ' linha do listview
    
    LinhaAtual = LinhaAtual + 1
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)

    Dim c As Integer
    
    Dim Posicao As Long
    Dim TopoCab As Long
    Dim Largura_Pagina As Long
    Dim FimColunas As Boolean
    
    If Not Impressora Then
        
        TopoCab = rpf_Cab(0).Top
        LinhaAtual = 1
        Largura_Pagina = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
            
        ' Acerta altura do cabecalho
        ' ** A FAZER : Permitir mais de uma linha de cabeçalho
        ' ** Atualmente campos a mais são cortados.
                                
        Posicao = 0
        FimColunas = False
        
        For c = 0 To Grd.ColumnHeaders.Count - 1
                    
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

