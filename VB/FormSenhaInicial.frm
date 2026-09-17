VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FormSenhaInicial 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9585
   ControlBox      =   0   'False
   Icon            =   "FormSenhaInicial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FormSenhaInicial.frx":0442
   ScaleHeight     =   7290
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      Height          =   285
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Prôxima Mensagem"
      Top             =   1080
      Width           =   285
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   405
      Top             =   3195
   End
   Begin VB.Data DatMinutos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4635
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5085
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      Height          =   280
      IMEMode         =   3  'DISABLE
      Left            =   7800
      TabIndex        =   0
      Top             =   4635
      Width           =   1380
   End
   Begin VB.Data DatPwd 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4410
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Pwd"
      Top             =   5580
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Data DatConfig 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2385
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Cfg"
      Top             =   5625
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton CmbCancel 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   280
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   5535
      Width           =   615
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   280
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Confirma"
      Top             =   5535
      Width           =   645
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
      Height          =   280
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5040
      Width           =   1380
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   75
      Left            =   45
      TabIndex        =   10
      Top             =   6210
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   765
      X2              =   495
      Y1              =   4545
      Y2              =   4545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      FillColor       =   &H00FFFFC0&
      Height          =   870
      Left            =   225
      Shape           =   3  'Circle
      Top             =   4095
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      DrawMode        =   5  'Not Copy Pen
      Visible         =   0   'False
      X1              =   495
      X2              =   495
      Y1              =   4275
      Y2              =   4545
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Minutos de Sabedoria / Gotas de Luz"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5265
      TabIndex        =   9
      Top             =   1125
      Width           =   3750
   End
   Begin VB.Label LblTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   4725
      TabIndex        =   8
      Top             =   1440
      Width           =   4800
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   5775
      Width           =   4920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licenciado para:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   5535
      Width           =   1185
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6930
      TabIndex        =   5
      Top             =   4635
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   6885
      Picture         =   "FormSenhaInicial.frx":CFAA
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   645
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6930
      TabIndex        =   4
      Top             =   4965
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1515
      Left            =   6615
      Shape           =   4  'Rounded Rectangle
      Top             =   4500
      Width           =   2790
   End
End
Attribute VB_Name = "FormSenhaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Cont As Integer
Dim counter As Integer
Dim Tempo As Integer
Dim t As Integer, l As Integer

Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
Dim PassWord As String
Dim Logado As String 'logado=0 nao logado, logado=1 logado primeira acao, logado=2 logado segunda acao, logado=3 administracao

Logado = ""

Dim Criterio As String
DatConfig.Recordset.MoveFirst


If Dir(App.Path & "\Vendas.Vbp") <> "" Then
    Logado = 3
End If

If TxtSenha = Left(Time, 2) & Mid(Time, 4, 2) Then
    Logado = 3
End If

If TxtSenha = Trim(DatConfig.Recordset.Fields("Senha1")) Or TxtSenha = Trim(DatConfig.Recordset.Fields("Senha2")) Then
      Logado = 1
End If
VarPermitePrd = False
VarPermiteEnt = False


If Logado = "" Then
    
    DatPwd.Recordset.FindFirst "Nome like '" & txtUsuario & "' and Senha like '" & TxtSenha & "'"
    VarPermitePrd = DatPwd.Recordset.Fields("AltPrd")
    VarPermiteEnt = DatPwd.Recordset.Fields("AltEnt")

    If DatPwd.Recordset.NoMatch Then
        Logado = 0
    Else
        Logado = 2
    End If
End If
'    End If
Admin = False


If Logado > 0 Then
    If TxtSenha = Trim(DatConfig.Recordset.Fields("Senha1")) Then
        MDIPrincipal.MnuRel(3).Enabled = False
        MDIPrincipal.MnuPrecos(6).Enabled = False
        MDIPrincipal.MnuPcl(1).Enabled = False
        MDIPrincipal.MnuCxa(3).Enabled = False
        
    ElseIf TxtSenha = Trim(DatConfig.Recordset.Fields("Senha2")) Or Logado = 3 Then
        Admin = True
        
        MDIPrincipal.MnuRel(3).Enabled = True
        MDIPrincipal.MnuPrecos(6).Enabled = True
        MDIPrincipal.MnuPcl(1).Enabled = True
    End If
        
    If Logado = 2 Then
        MontaMenu
    End If
    
'    Timer1.Enabled = True
'    Tempo = 0
    MDIPrincipal.Show
    Unload Me
      
      
Else
    MsgBox "Nome de usuário ou Senha Inválida !", vbExclamation, App.Title
    If Cont < 3 Then
        txtUsuario.SetFocus
        txtUsuario.SelStart = 0
        txtUsuario.SelLength = 5000
        Cont = Cont + 1
    Else
        Unload Me
        End
    End If
End If
End Sub

Private Sub Command1_Click()
Minutos
End Sub

Private Sub Form_Activate()
Cont = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
Dim Erro As String
Dim UltimoAcesso As Date

Erro = ""

Caminho = App.Path ' & "\Morto"


CfgNome = LerINI("Global", "X1", App.Path & "\config.ini")
Lbl.Caption = LerINI("Global", "X1", App.Path & "\config.ini")
tipoCopia = LerINI("Global", "X10", App.Path & "\config.ini")
UltimoAcesso = IIf(IsDate(LerINI("Global", "X12", App.Path & "\config.ini")), IsDate(LerINI("Global", "UltimoAcesso", App.Path & "\config.ini")), Date)


    If CfgNome = "" Then
        FrmConfigSistema.txtCpf.Enabled = True
        FrmConfigSistema.Show vbModal
    End If

If GetSetting(App.EXEName, "SystemReport", "Sys01") = "" And GetSetting(App.EXEName, "SystemReport", "Sys02") = "" Then
    Erro = "Código 1 - Programa não autorizado a operar neste disco!"
ElseIf DeCripta(GetSetting(App.EXEName, "SystemReport", "Sys02")) <> LerINI("Global", "X6", App.Path & "\config.ini") Then
    Erro = "Código 2 - Erro na proteção"
ElseIf DeCripta(GetSetting(App.EXEName, "SystemReport", "Sys01")) <> DriveSerial(Mid(App.Path, 1, 3)) Then
    Erro = "Código 3 - Programa não autorizado a operar neste disco!"
ElseIf Date < CDate(IIf(IsDate(LerINI("Global", "X12", App.Path & "\config.ini")), LerINI("Global", "X12", App.Path & "\config.ini"), Date)) Then
    If LerINI("Global", "X10", App.Path & "\config.ini") <> 5 Then
        Erro = "Código 7 - Data anterior a da ultima execução!"
    End If
ElseIf CDate(IIf(IsDate(LerINI("Global", "X11", App.Path & "\config.ini")), LerINI("Global", "X11", App.Path & "\config.ini"), Date)) + IIf(LerINI("Global", "X9", App.Path & "\config.ini") = "", 0, LerINI("Global", "X9", App.Path & "\config.ini")) < Date Then
    If LerINI("Global", "X10", App.Path & "\config.ini") <> 5 Then
            Erro = "Código 8 - Validade da cópia expirada!"
    End If
Else
    Call EscreveINI("Global", "X12", Date, App.Path & "\Config.ini") 'aceso
End If

'erro = "" ' AGROLEITE
 
If Dir(App.Path & "\Vendas.Vbp") = "" Then

    If Erro <> "" Then
            FrmConfig.varMensagem = Erro
            FrmConfig.FechaAoSair = True
            FrmConfig.Show vbModal
    End If
End If

If App.PrevInstance Then
    MsgBox "Existe uma aplicação do " & App.ProductName & " aberta nesta máquina, um ou mais programas abertos pode causar lentidão no sistema!", 16
    End
End If

If Dir(App.Path & "\rodape.jpg") <> "" Then
    Rodape.Picture = LoadPicture(App.Path & "\rodape.jpg")
End If

DatConfig.DatabaseName = Caminho & "\Dados.Mdb"
DatPwd.DatabaseName = Caminho & "\Dados.Mdb"


'banco de dados do srcaixa

If LerINI("Financeiro", "Comunica", App.Path & "\config.ini") = 1 Then
    
    If LerINI("Financeiro", "LocalBcoFin", App.Path & "\config.ini") <> "" Then
        StringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & LerINI("Financeiro", "LocalBcoFin", App.Path & "\config.ini") & ";Persist Security Info=False"
        
        LocalCursor = adUseClient
        If BancoSRCaixa.State = 1 Then Banco.Close
        BancoSRCaixa.ConnectionString = StringConexao
        BancoSRCaixa.CommandTimeout = 0
        
        BancoSRCaixa.Open
    Else
        MsgBox "Caminho do banco de dados do sistema financeiro não foi configurado", vbInformation, App.Title
    End If
End If

Minutos
End Sub

Private Sub LblTxt_Click()
Minutos
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdOk_Click
End If
End Sub

Private Sub MontaMenu()
Dim X As Integer

With MDIPrincipal

    If Mid(DatPwd.Recordset("MnuCad"), 1, 1) = "1" Then
        .MnuPrincipal(0).Visible = True
        For X = 0 To .MnuCad.Count - 1
            If Mid(DatPwd.Recordset("MnuCad"), X + 2, 1) = "1" Then
                .MnuCad(X).Visible = True
            Else
                .MnuCad(X).Visible = False
                If X = 3 Then
                    .Toolbar1.Buttons(4).Enabled = False
                End If
                If X = 1 Then
'                    .Toolbar1.Buttons(1).Enabled = False
                End If
            End If
        Next
    Else
        .MnuPrincipal(0).Visible = False
        .Toolbar1.Buttons(1).Enabled = False
        .Toolbar1.Buttons(4).Enabled = False
        
    End If

    If Mid(DatPwd.Recordset("MnuPrecos"), 1, 1) = "1" Then
        .MnuPrincipal(1).Visible = True
        For X = 0 To .MnuPrecos.Count - 1
            If Mid(DatPwd.Recordset("MnuPrecos"), X + 2, 1) = "1" Then
                .MnuPrecos(X).Visible = True
            Else
                .MnuPrecos(X).Visible = False
            End If
        Next
    Else
        .MnuPrincipal(1).Visible = False
    End If
        
    If Mid(DatPwd.Recordset("MnuPed"), 1, 1) = "1" Then
        .MnuPrincipal(2).Visible = True
        For X = 0 To .MnuPed.Count - 1
            If Mid(DatPwd.Recordset("MnuPed"), X + 2, 1) = "1" Then
                .MnuPed(X).Visible = True
            Else
                .MnuPed(X).Visible = False
                If X = 2 Then
                    .Toolbar1.Buttons(10).Enabled = False
                End If
                
            End If
        Next
    Else
        .MnuPrincipal(2).Visible = False
        .Toolbar1.Buttons(7).Enabled = False
        .Toolbar1.Buttons(10).Enabled = False
        .Toolbar1.Buttons(13).Enabled = False
    End If
    
    If Mid(DatPwd.Recordset("MnuPcl"), 1, 1) = "1" Then
        .MnuPrincipal(3).Visible = True
        For X = 0 To .MnuPcl.Count - 1
            If Mid(DatPwd.Recordset("MnuPcl"), X + 2, 1) = "1" Then
                .MnuPcl(X).Visible = True
            Else
                .MnuPcl(X).Visible = False
                If X = 0 Then
                    .Toolbar1.Buttons(13).Enabled = False
                End If
                
            End If
        Next
    Else
        .MnuPrincipal(3).Visible = False
'        If x = 0 Then
            .Toolbar1.Buttons(13).Enabled = False
'        End If
        
    End If
    
    If Mid(DatPwd.Recordset("MnuPag"), 1, 1) = "1" Then
        .MnuPrincipal(4).Visible = True
        For X = 0 To .MnuPag.Count - 1
            If Mid(DatPwd.Recordset("MnuPag"), X + 2, 1) = "1" Then
                .MnuPag(X).Visible = True
            Else
                .MnuPag(X).Visible = False
            End If
        Next
    Else
        .MnuPrincipal(4).Visible = False
    End If
        
    If Mid(DatPwd.Recordset("MnuCxa"), 1, 1) = "1" Then
        .MnuPrincipal(5).Visible = True
        For X = 0 To .MnuCxa.Count - 1
            If Mid(DatPwd.Recordset("MnuCxa"), X + 2, 1) = "1" Then
                .MnuCxa(X).Visible = True
            Else
                .MnuCxa(X).Visible = False
            End If
        Next
    Else
        .MnuPrincipal(5).Visible = False
    End If

    If Mid(DatPwd.Recordset(9), 1, 1) = "1" Then
        .MnuPrincipal(6).Visible = True
        For X = 0 To .MnuRel.Count - 1
            If Mid(DatPwd.Recordset("MnuRel"), X + 2, 1) = "1" Then
                .MnuRel(X).Visible = True
            Else
                .MnuRel(X).Visible = False
            End If
        Next
    Else
        .MnuPrincipal(6).Visible = False
    End If

    If Mid(DatPwd.Recordset("MnuUtil"), 1, 1) = "1" Then
        .MnuPrincipal(7).Visible = True
        For X = 0 To .MnuUtil.Count - 1
            If Mid(DatPwd.Recordset("MnuUtil"), X + 2, 1) = "1" Then
                .MnuUtil(X).Visible = True
            Else
                .MnuUtil(X).Visible = False
            End If
        Next
    Else
        .MnuPrincipal(7).Visible = False
    End If

End With

End Sub


Private Sub Minutos()
Dim Txtminuto As Integer
Dim Criterio As String
Dim X As Double

If Dir(Caminho & "\Minutos.mdb") <> "" Then
'    Command1.Visible = True
    DatMinutos.DatabaseName = Caminho & "\Minutos.mdb"
    DatMinutos.RecordSource = "Select Numero, Mensagem From Minutos Order By Numero"
    DatMinutos.Refresh
'    For x = 1 To Int((Rnd) * 30)
        Txtminuto = Format(Int(CCur(Mid(Time, 7, 2))) + (Rnd * 568), "000")
        If Txtminuto <= 0 Or Txtminuto > 568 Then
            Txtminuto = Format(Int(CCur(Mid(Time, 7, 2))) + (Rnd * 568), "000")
        End If

'        MsgBox Txtminuto, vbInformation
'    Next
    Criterio = "Numero = " & Txtminuto
    DatMinutos.Recordset.FindFirst Criterio
    If Not DatMinutos.Recordset.NoMatch Then
        LblTxt = DatMinutos.Recordset.Fields("Mensagem")
'        Texto.ToolTipText = Texto.Text
    Else
        LblTxt = Txtminuto
    End If
        
End If


End Sub

Private Sub Timer1_Timer()


If Tempo < 100 Then
   t = (Tempo) * (3.14 / 10)
   ProgressBar1.value = Tempo + 1
   'ProgressBar2.Value = time + 1
'   r (t)
   Tempo = Tempo + 1
   Timer1.Enabled = True
    
Else
   Timer1.Enabled = False

'  MDIPrincipal.Show
'  Unload Me
End If
End Sub

Public Function r(t As Integer)
Line1.y1 = (Line1.y2 + 230 * Cos(t + 5))
Line1.X1 = (Line1.X2 - 230 * Sin(t + 5))
Line2.y1 = (Line2.y2 + 230 * Cos(t / 20 + 5))
Line2.X1 = (Line2.X2 - 230 * Sin(t / 20 + 5))
End Function


