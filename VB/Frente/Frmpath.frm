VERSION 5.00
Begin VB.Form FrmPath 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SISTEMA FRENTE DE CAIXA"
   ClientHeight    =   8895
   ClientLeft      =   3420
   ClientTop       =   2700
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "Frmpath.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Aviso 
      BackColor       =   &H0080FFFF&
      Caption         =   "AGUARDE . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6300
      Visible         =   0   'False
      Width           =   11040
   End
   Begin VB.TextBox TxtSenha 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      IMEMode         =   3  'DISABLE
      Left            =   7965
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   5
      Top             =   5670
      Width           =   420
   End
   Begin VB.Label LblTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7920
      TabIndex        =   3
      Top             =   2565
      Width           =   3390
   End
   Begin VB.Label LblData 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7920
      TabIndex        =   2
      Top             =   1710
      Width           =   3390
   End
   Begin VB.Label LblLocal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   675
      TabIndex        =   1
      Top             =   6885
      Width           =   10770
   End
End
Attribute VB_Name = "FrmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
VarBaixaAuto = "Nao"
VarMostraCusto = "Nao"

On Error Resume Next

If dir(App.Path & "\Config.Txt") <> "" Then
    Open App.Path & "\Config.txt" For Input As #1
    Input #1, LocalBanco
    Input #1, Impressora
    Input #1, PortaF
    Input #1, PortaNF
    Input #1, Leitor
    Input #1, ConsultaCh
    Input #1, CfgDesc
    Input #1, Impr40
    Input #1, NroLinha
    Input #1, VarLocalEst
    Input #1, VarMeiaFolha
    Input #1, VarComplemento
    Input #1, VarBaixaAuto
    Input #1, VarMostraCusto
    Close #1
    LblLocal.Caption = LocalBanco
Else
    FormConfigura.Show vbModal
End If

If Impr40 = "Nao" Then
    Nor = Chr(18)
    Cond = Chr(15)
Else
    Nor = Chr(18)
    Cond = Chr(15)
End If
Gde = Chr(14)
Negi = Chr(27) + Chr(69)
NegF = Chr(27) + Chr(70)
SaltoPag = Chr(12)

LimiteDesc = LerINI("DescontoFrente", "Valor", LocalBanco & "\config.ini")
EditarFrente = LerINI("DescontoFrente", "Editar", LocalBanco & "\config.ini")
Receber = LerINI("DescontoFrente", "Receber", LocalBanco & "\config.ini")
VarPcoPrazo = LerINI("Pedido", "PzoProduto", LocalBanco & "\config.ini")
VarReajConta = LerINI("Pedido", "ReajustaConta", LocalBanco & "\config.ini")

If Impr40 = "Nao" Then
    VarEmpresa = LerINI("Config", "Empresa", LocalBanco & "\config.ini")
    VarCGC = LerINI("Config", "Cnpj", LocalBanco & "\config.ini")
    VarInscEst = LerINI("Config", "Inscr", LocalBanco & "\config.ini")
    VarEnd = LerINI("Config", "Endereco", LocalBanco & "\config.ini")
    VarCidade = LerINI("Config", "CidadeEmpresa", LocalBanco & "\config.ini")
End If

VarSenhaBco = LerINI("LOCALBCO", "SenhaBco", LocalBanco & "\config.ini")

LblData.Caption = Date
LblTime.Caption = Time
TxtSenha = ""
Aviso.Visible = False

Dim Sql As String

End Sub

Private Sub Form_Load()

If App.PrevInstance Then
    MsgBox "Existe uma aplicação do " & App.ProductName & " aberta nesta máquina, um ou mais programas abertos pode causar lentidão no sistema!", 16
    End
End If

End Sub

Private Sub TxtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Status As String
Dim Senha  As String

If KeyCode = 13 Then
   
    If Not IsNumeric(TxtSenha) Then
        TxtSenha.Text = ""
        TxtSenha.SetFocus
        Exit Sub
    End If
            
    TxtSenha = Val(TxtSenha)
    
    If TxtSenha < 10 Then
    
        If dir(LocalBanco & "\Dados.Mdb") = "" Then
             If MsgBox("Servidor não encontrado ! Deseja entrar no modo de emergencia ?", vbYesNo, App.Path) = vbYes Then
                    LblLocal.Caption = App.Path
                    LocalBanco = LblLocal.Caption
             End If
        End If
    End If
    
    Select Case TxtSenha
        Case 1
            Aviso.Visible = True
            DoEvents
            MsgBox "Certifique-se de que a impressora fiscal está ligada.", vbInformation
            FrmFrente.TxtDados(2).Text = TxtSenha.Text
            FrmFrente.Show vbModal
        Case 2
            Aviso.Visible = True
            DoEvents
            Impressora = 99
            FrmFrente.Show
            Unload Me
        Case 10
            FormData.Show vbModal
        Case 20
            Senha = InputBox("Senha")
            If Senha = "CONFIG" Then
                FormConfigura.Show vbModal
            End If
        Case 30
            Shell "C:\WINDOWS\EXPLORER.EXE /n,/e,C:\", vbMaximizedFocus
        Case 99
'            Shell "C:\Windows\Rundll.exe user.exe,exitwindows"
        Case 100
            End
        Case Else
            TxtSenha.Text = ""
            TxtSenha.SetFocus
    End Select
'    TxtSenha = ""
'    TxtSenha.SetFocus
End If

End Sub
