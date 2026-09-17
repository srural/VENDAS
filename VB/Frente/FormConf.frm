VERSION 5.00
Begin VB.Form FormConfigura 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   360
      ScaleWidth      =   11355
      TabIndex        =   39
      Top             =   5355
      Width           =   11355
   End
   Begin VB.CheckBox ChkGelado 
      Caption         =   "Gelado"
      Height          =   375
      Left            =   6030
      TabIndex        =   38
      Top             =   4680
      Width           =   1365
   End
   Begin VB.CheckBox ChkPreco 
      Caption         =   "Preço"
      Height          =   375
      Left            =   4500
      TabIndex        =   37
      Top             =   4680
      Width           =   1365
   End
   Begin VB.CheckBox ChkCaixa 
      Caption         =   "Caixa"
      Height          =   375
      Left            =   2700
      TabIndex        =   36
      Top             =   4680
      Width           =   1230
   End
   Begin VB.CheckBox ChkEst 
      Caption         =   "Estoque"
      Height          =   375
      Left            =   270
      TabIndex        =   35
      Top             =   4680
      Width           =   1950
   End
   Begin VB.TextBox TxtBarra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7965
      TabIndex        =   10
      Top             =   2925
      Width           =   1200
   End
   Begin VB.TextBox TxtDet 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6210
      TabIndex        =   9
      Top             =   2925
      Width           =   1335
   End
   Begin VB.CheckBox ChkDesconto 
      Caption         =   "Permitir Desconto Final"
      Height          =   375
      Left            =   270
      TabIndex        =   16
      Top             =   4275
      Width           =   1950
   End
   Begin VB.CheckBox ChkLanchonete 
      Caption         =   "Lanchonete"
      Height          =   375
      Left            =   2700
      TabIndex        =   14
      Top             =   3870
      Width           =   1590
   End
   Begin VB.CheckBox ChkConveniencia 
      Caption         =   "Conveniencia"
      Height          =   375
      Left            =   270
      TabIndex        =   13
      Top             =   3870
      Width           =   1815
   End
   Begin VB.CheckBox Chk4000 
      Caption         =   "MP-4000"
      Height          =   375
      Left            =   2700
      TabIndex        =   17
      Top             =   4275
      Width           =   1125
   End
   Begin VB.ComboBox TipoImpressora 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormConf.frx":0000
      Left            =   6210
      List            =   "FormConf.frx":0010
      TabIndex        =   6
      Top             =   2250
      Width           =   2985
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10140
      TabIndex        =   30
      Top             =   0
      Width           =   10140
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Configuração do PDV"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   135
         TabIndex        =   31
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.CheckBox ChkCusto 
      Caption         =   "Mostra Custo"
      Height          =   375
      Left            =   2700
      TabIndex        =   11
      Top             =   3465
      Width           =   1575
   End
   Begin VB.CheckBox ChkEstoque 
      Caption         =   "Baixa Automática de Estoque"
      Height          =   375
      Left            =   4500
      TabIndex        =   12
      Top             =   3465
      Width           =   2715
   End
   Begin VB.CheckBox ChkCompl 
      Caption         =   "Complemento do Produto"
      Height          =   375
      Left            =   270
      TabIndex        =   20
      Top             =   3465
      Width           =   2205
   End
   Begin VB.CheckBox ChkMeia 
      Caption         =   "1/2 Folha"
      Height          =   375
      Left            =   4500
      TabIndex        =   18
      Top             =   4275
      Width           =   1125
   End
   Begin VB.TextBox TxtLocalEst 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3420
      TabIndex        =   8
      Top             =   2925
      Width           =   2235
   End
   Begin VB.TextBox TxtLinha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3405
      TabIndex        =   5
      Top             =   2295
      Width           =   1740
   End
   Begin VB.CheckBox ChkDesc 
      Caption         =   "Permitir Desconto no Item"
      Height          =   375
      Left            =   4500
      TabIndex        =   15
      Top             =   3870
      Width           =   2310
   End
   Begin VB.ComboBox CmbImpressora 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormConf.frx":0032
      Left            =   270
      List            =   "FormConf.frx":0054
      TabIndex        =   1
      Top             =   1575
      Width           =   2625
   End
   Begin VB.TextBox TxtConsulta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   7
      Top             =   2940
      Width           =   2910
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Sair"
      Height          =   420
      Left            =   8715
      TabIndex        =   21
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Confirma"
      Height          =   420
      Left            =   7470
      TabIndex        =   19
      Top             =   4050
      Width           =   915
   End
   Begin VB.ComboBox CmbLeitor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormConf.frx":00D3
      Left            =   270
      List            =   "FormConf.frx":00E6
      TabIndex        =   4
      Top             =   2250
      Width           =   2580
   End
   Begin VB.ComboBox CmbNF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormConf.frx":0117
      Left            =   6210
      List            =   "FormConf.frx":0130
      TabIndex        =   3
      Top             =   1575
      Width           =   3435
   End
   Begin VB.ComboBox CmbPorta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormConf.frx":0161
      Left            =   3375
      List            =   "FormConf.frx":0174
      TabIndex        =   2
      Top             =   1575
      Width           =   2355
   End
   Begin VB.TextBox TxtBanco 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   0
      Top             =   945
      Width           =   7185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamanho Barra"
      Height          =   195
      Index           =   10
      Left            =   7965
      TabIndex        =   34
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Linha Detalhe"
      Height          =   195
      Index           =   9
      Left            =   6210
      TabIndex        =   33
      Top             =   2700
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Impressora"
      Height          =   195
      Index           =   8
      Left            =   6210
      TabIndex        =   32
      Top             =   2025
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local de Estoque"
      Height          =   195
      Index           =   7
      Left            =   3420
      TabIndex        =   29
      Top             =   2700
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Linhas Final Cupom"
      Height          =   195
      Index           =   6
      Left            =   3420
      TabIndex        =   28
      Top             =   2070
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local do &Programa Consulta"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   27
      Top             =   2700
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Leitor de Código de Barras"
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   26
      Top             =   1980
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Local Impressora NF"
      Height          =   195
      Index           =   4
      Left            =   6210
      TabIndex        =   25
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Porta"
      Height          =   195
      Index           =   3
      Left            =   3375
      TabIndex        =   24
      Top             =   1305
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Modelo da Impressora Fiscal"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   23
      Top             =   1350
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local do &Banco de Dados"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   22
      Top             =   720
      Width           =   1860
   End
End
Attribute VB_Name = "FormConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtSair_Click()
Unload Me
End Sub

'Open App.Path & "\Config.txt" For Output As #1
'Print #1, Remessa

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub CmdOk_Click()

EscreveINI "Config", "LocalBco", TxtBanco, App.Path & "\config.ini"
EscreveINI "Config", "Ecf", CmbImpressora.ListIndex + 1, App.Path & "\config.ini"
EscreveINI "Config", "Porta", CmbPorta, App.Path & "\config.ini"
EscreveINI "Config", "PortaNf", CmbNF, App.Path & "\config.ini"
EscreveINI "Config", "Leitor", CmbLeitor, App.Path & "\config.ini"
EscreveINI "Config", "Consulta", TxtConsulta, App.Path & "\config.ini"
EscreveINI "Config", "CfgDesc", ChkDesc, App.Path & "\config.ini"
EscreveINI "Config", "QtdLinhas", TxtLinha, App.Path & "\config.ini"
EscreveINI "Config", "LocalEst", TxtLocalEst, App.Path & "\config.ini"
EscreveINI "Config", "MeiaFolha", ChkMeia, App.Path & "\config.ini"
EscreveINI "Config", "Complemento", ChkCompl, App.Path & "\config.ini"
EscreveINI "Config", "BaixaAuto", ChkEstoque, App.Path & "\config.ini"
EscreveINI "Config", "MostraCusto", ChkCusto, App.Path & "\config.ini"
EscreveINI "Config", "ImprNf", TipoImpressora.ListIndex + 1, App.Path & "\config.ini"
EscreveINI "Config", "MP4000", Chk4000.Value, App.Path & "\config.ini"

EscreveINI "Config", "Conveniencia", ChkConveniencia.Value, App.Path & "\config.ini"
EscreveINI "Config", "Desconto", ChkDesconto.Value, App.Path & "\config.ini"
EscreveINI "Config", "Lanchonete", ChkLanchonete.Value, App.Path & "\config.ini"
EscreveINI "Config", "TamBarra", TxtBarra.Text, App.Path & "\config.ini"
EscreveINI "Config", "LinhaDet", TxtDet.Text, App.Path & "\config.ini"

EscreveINI "Config", "Preco", ChkPreco, App.Path & "\config.ini"
EscreveINI "Config", "Caixa", ChkCaixa, App.Path & "\config.ini"
EscreveINI "Config", "MostraEstoque", ChkEst, App.Path & "\config.ini"
EscreveINI "Config", "Gelado", ChkGelado, App.Path & "\config.ini"
            
If dir(App.Path & "\Config.Txt") <> "" Then
'     Kill App.Path & "\Config.Txt"
End If

Unload Me
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next

If dir(App.Path & "\Config.ini") <> "" Then
            
    TxtBanco = LerINI("Config", "LocalBco", App.Path & "\config.ini")
    CmbImpressora.ListIndex = LerINI("Config", "Ecf", App.Path & "\config.ini") - 1
    CmbPorta = LerINI("Config", "Porta", App.Path & "\config.ini")
    CmbLeitor = LerINI("Config", "Leitor", App.Path & "\config.ini")
    TxtConsulta = LerINI("Config", "Consulta", App.Path & "\config.ini")
    Chk4000 = LerINI("Config", "MP4000", App.Path & "\config.ini")
    CmbNF = LerINI("Config", "PortaNf", App.Path & "\config.ini")
            
    ChkDesc = LerINI("Config", "CfgDesc", App.Path & "\config.ini")
    TxtLinha = LerINI("Config", "QtdLinhas", App.Path & "\config.ini")
    TxtLocalEst = LerINI("Config", "LocalEst", App.Path & "\config.ini")
    ChkMeia = LerINI("Config", "MeiaFolha", App.Path & "\config.ini")
    ChkCompl = LerINI("Config", "Complemento", App.Path & "\config.ini")
    ChkEstoque = LerINI("Config", "BaixaAuto", App.Path & "\config.ini")
    ChkCusto = LerINI("Config", "MostraCusto", App.Path & "\config.ini")
    TipoImpressora.ListIndex = LerINI("Config", "ImprNf", App.Path & "\config.ini") - 1
    
    ChkConveniencia = LerINI("Config", "Conveniencia", App.Path & "\config.ini")
    ChkLanchonete = LerINI("Config", "Lanchonete", App.Path & "\config.ini")
    ChkDesconto = LerINI("Config", "Desconto", App.Path & "\config.ini")
    TxtDet = LerINI("Config", "LinhaDet", App.Path & "\config.ini")
    TxtBarra = LerINI("Config", "TamBarra", App.Path & "\config.ini")
                    
    ChkPreco = LerINI("Config", "Preco", App.Path & "\config.ini")
    ChkCaixa = LerINI("Config", "Caixa", App.Path & "\config.ini")
    ChkEst = LerINI("Config", "MostraEstoque", App.Path & "\config.ini")
    ChkGelado = LerINI("Config", "Gelado", App.Path & "\config.ini")
                    
    
ElseIf dir(App.Path & "\Config.Txt") <> "" Then

    Open App.Path & "\Config.txt" For Input As #1
    Input #1, LocalBanco
    Input #1, Impressora
    Input #1, PortaF
    Input #1, PortaNF
    Input #1, Leitor
    Input #1, ConsultaCh
    Input #1, CfgDesc
    Input #1, CfgImpressora
    Input #1, VarLinha
    Input #1, VarLocalEst
    Input #1, VarMeiaFolha
    Input #1, VarComplemento
    Input #1, VarBaixaAuto
    Input #1, VarMostraCusto
    
    Close #1
    If CfgDesc = "Sim" Then
        ChkDesc.Value = 1
    Else
        ChkDesc.Value = 0
    End If
    If CfgImpressora = "Sim" Then
        ChkImpressora.Value = 1
    Else
        ChkImpressora.Value = 0
    End If
    
    TxtBanco = LocalBanco
    CmbImpressora.ListIndex = Impressora - 1
    CmbPorta = PortaF
    CmbNF = PortaNF
    CmbLeitor = Leitor
    TxtConsulta = ConsultaCh
    TxtLinha = VarLinha
    TxtLocalEst = VarLocalEst
    
    If VarMeiaFolha = "Sim" Then
        ChkMeia.Value = 1
    Else
        ChkMeia.Value = 0
    End If
    
    If VarComplemento = "Sim" Then
        ChkCompl.Value = 1
    Else
        ChkCompl.Value = 0
    End If
    
    If VarBaixaAuto = "Sim" Then
        ChkEstoque.Value = 1
    Else
        ChkEstoque.Value = 0
    End If
    
    If VarMostraCusto = "Sim" Then
        ChkCusto.Value = 1
    Else
        ChkCusto.Value = 0
    End If
        
Else
    TxtBanco = "\\servidor\Vendas"
    CmbImpressora.ListIndex = 0
    CmbPorta.ListIndex = 0
    CmbNF.ListIndex = 0
    CmbLeitor.ListIndex = 0
    TipoImpressora.ListIndex = 0
    TxtConsulta = App.Path
    TxtLinha = 7
    VarLocalEst = "LOJA"
End If
End Sub

