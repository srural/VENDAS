VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormAFiscais 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   4545
   ClientTop       =   3675
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7575
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   2070
      Width           =   5145
   End
   Begin VB.ComboBox CmbMes 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormAFiscais.frx":0000
      Left            =   360
      List            =   "FormAFiscais.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1110
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   7
      Top             =   2970
      Width           =   14370
   End
   Begin ComctlLib.ProgressBar BarraProgresso 
      Height          =   285
      Left            =   2295
      TabIndex        =   3
      Top             =   4950
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9960
      TabIndex        =   5
      Top             =   0
      Width           =   9960
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos Fiscais"
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
         TabIndex        =   6
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.TextBox txtZip 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1980
      TabIndex        =   1
      Top             =   5445
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.CommandButton CmdCopia 
      Caption         =   "&Inicia a Cópia"
      Height          =   420
      Left            =   3285
      TabIndex        =   0
      Top             =   1035
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   450
      TabIndex        =   11
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mês Referente"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   855
      Width           =   1545
   End
   Begin VB.Label LblProgresso 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4635
      TabIndex        =   4
      Top             =   4635
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progresso"
      Height          =   195
      Left            =   1455
      TabIndex        =   2
      Top             =   4995
      Width           =   705
   End
End
Attribute VB_Name = "FormAFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdCopia_Click()
Dim ArqCompactadoNfe As String
Dim ArqCompactadoSat As String
Dim ArqCompactadoCancel As String
Dim ArqOrigem As String
Dim ArqSatVendas As String
Dim ArqSatCancel As String
Dim VarTemp As String
Dim AcbrSat As String


If CmbMes.Text = "" Then
    MsgBox "Selecione o Mês", vbInformation
    Exit Sub
End If

If LerINI("EMITENTE", "Cnpj", App.Path & "\Nfe.ini") <> "" Then
    VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
End If

If LerINI("CONFIG", "PastaSat", App.Path & "\Config.ini") <> "" Then
    AcbrSat = LerINI("CONFIG", "PastaSat", App.Path & "\Condig.ini")
Else
    AcbrSat = ""
End If

BarraProgresso.Min = 0
BarraProgresso.Max = 100
BarraProgresso.value = 0

If AcbrSat <> "" Then
    ArqSatVendas = AcbrSat & "\Arqs\Sat\Vendas\" & VarCnpjEmite & "\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "\*.*"
    ArqSatCancel = AcbrSat & "\Arqs\Sat\Cancelamentos\" & VarCnpjEmite & "\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "\*.*"
End If

ArqCompactadoNfe = PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfe\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "-Nfe.zip"
ArqCompactadoSat = PortaAcbr & "\Arqs\Sat\Cancelamentos\" & VarCnpjEmite & "\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "-Sat.zip"
ArqCompactadoCancel = PortaAcbr & "\Arqs\Sat\Cancelamentos\" & VarCnpjEmite & "\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "-SatCancelado.zip"

ArqOrigem = PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfe\" & Trim(Str(Year(Date))) & psCompString(1, CmbMes.ListIndex + 1, 2) & "\*.*"



'If Dir(ArqCompactado) <> "" Then
'    Set Fs = CreateObject("Scripting.FileSystemObject")
'    Set Arq = Fs.GetFile(ArqCompactado)
'    Arq.Delete
'End If

'CmdCopia.Enabled = False

If Dir("C:\Arquivos de Programas\WinRAR\*.*") <> "" Then
    Shell "C:\Arquivos de Programas\WinRAR\winrar a " & ArqCompactadoNfe & " " & ArqOrigem
    Shell "C:\Arquivos de Programas\WinRAR\winrar a " & ArqCompactadoSat & " " & ArqSatVendas
    Shell "C:\Arquivos de Programas\WinRAR\winrar a " & ArqCompactadoCancel & " " & ArqSatCancel
Else
    Shell "C:\Program Files\WinRAR\winrar a " & ArqCompactadoNfe & " " & ArqOrigem
    Shell "C:\Program Files\WinRAR\winrar a " & ArqCompactadoSat & " " & ArqSatVendas
    Shell "C:\Program Files\WinRAR\winrar a " & ArqCompactadoCancel & " " & ArqSatCancel
End If

If Dir(ArqCompactadoNfe) = "" Then

    While FileLen(ArqCompactado) = 0

    Wend

End If

If Dir(App.Path & "\email.ini") <> "" Then
    Kill App.Path & "\email.ini"
End If

EscreveINI "Email", "NomeExibicao", "SidComp Automação", App.Path & "\Email.ini"
EscreveINI "Email", "Endereco", "sidcomp.urupes@gmail.com", App.Path & "\Email.ini"
EscreveINI "Email", "Email", "smtp.gmail.com", App.Path & "\Email.ini"
EscreveINI "Email", "Usuario", "sidcomp.urupes@gmail.com", App.Path & "\Email.ini"
EscreveINI "Email", "Senha", "sidgmail1004", App.Path & "\Email.ini"
EscreveINI "Email", "Porta", "25", App.Path & "\Email.ini"
EscreveINI "Email", "ExigeSSL", "0", App.Path & "\Email.ini"
EscreveINI "Email", "ExigeTLS", "1", App.Path & "\Email.ini"

AbreAcbr "EMail.Novo(" & App.Path & "\Email.ini" & ")"

While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.AdicionaPara(" & TxtEmail & "," & NomeEmpresa & ")"

While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.Assunto(" & "Arquivos Fiscais -  " & NomeEmpresa & ")"
While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.TextoMensagem(" & "Segue os Arquivos sat e nfe -  att SidComp " & ")"
While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.AdicionaAnexo(" & ArqCompactadoNfe & "," & "Arquivos fiscais - " & NomeEmpresa & ")"
While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.AdicionaAnexo(" & ArqCompactadoSat & "," & "Arquivos fiscais - " & NomeEmpresa & ")"
While Dir(PortaAcbr & "\Sai.txt") = ""
Wend

AbreAcbr "EMail.AdicionaAnexo(" & ArqCompactadoCancel & "," & "Arquivos fiscais - " & NomeEmpresa & ")"
While Dir(PortaAcbr & "\Sai.txt") = ""
Wend


AbreAcbr "EMail.Enviar"

While Dir(PortaAcbr & "\Sai.txt") = ""

Wend

Open PortaAcbr & "\Sai.txt" For Input As #6

Line Input #6, VarTemp

MsgBox VarTemp, vbInformation

Close #6
'Compacta ArqCompactado, ArqOrigem

End Sub

Private Sub Form_Load()
Dim Formulario As Form
InicializaZip Me, txtZip

Set Formulario = Me

Formulario.Left = 400
Formulario.Top = 500

If LerINI("Config", "e-mail", App.Path & "\config.ini") <> "" Then
    TxtEmail = LerINI("Config", "E-mail", App.Path & "\config.ini")
End If

CmbMes.ListIndex = Month(Date) - 2


End Sub

Private Sub txtZip_Change()
' lblProgresso = TipoAção(Val(GetAction(txtZip))) & " "
' lblProgresso = lblProgresso & GetFileName(txtZip) & "->"
' lblProgresso = lblProgresso & GetPercentComplete(txtZip) & "%"

BarraProgresso.value = GetPercentComplete(txtZip)
LblProgresso = GetPercentComplete(txtZip) & " %"
DoEvents
If GetPercentComplete(txtZip) >= 99 Then
    CmdCopia.Enabled = True
    LblProgresso = ""
End If
End Sub
