VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConfigSistema 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "FrmConfigSistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   17
      Top             =   3375
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14655
      TabIndex        =   15
      Top             =   0
      Width           =   14655
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Configurações do Sistema"
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
         TabIndex        =   16
         Top             =   45
         Width           =   7125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Empresa"
      Height          =   1710
      Left            =   1080
      TabIndex        =   7
      Top             =   870
      Width           =   5535
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1425
         TabIndex        =   13
         Top             =   1245
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   285
         Left            =   3945
         TabIndex        =   4
         Top             =   1230
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4695
         TabIndex        =   3
         Top             =   915
         Width           =   675
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1425
         TabIndex        =   2
         Top             =   915
         Width           =   2820
      End
      Begin VB.TextBox txtResponsavel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   3945
      End
      Begin VB.TextBox txtEscritorio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cpf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3600
         TabIndex        =   12
         Top             =   1245
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4320
         TabIndex        =   11
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   225
         TabIndex        =   10
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   210
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton BtCancelar 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5625
      Picture         =   "FrmConfigSistema.frx":0442
      TabIndex        =   6
      Top             =   2715
      Width           =   1000
   End
   Begin VB.CommandButton BtOk 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   345
      Left            =   4425
      Picture         =   "FrmConfigSistema.frx":0B84
      TabIndex        =   5
      Top             =   2715
      Width           =   1000
   End
End
Attribute VB_Name = "FrmConfigSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtCancelar_Click()
    Unload Me
End Sub

Private Sub BtOK_Click()
    If txtEscritorio.Text = "" Then
        MsgBox "Informe o Escritório!", 16
        txtUsuario.SetFocus
        Exit Sub
'    ElseIf txtCaminhoImpr.Text = "" Then
'        MsgBox "Informe um caminho válido para a impressora matricial", 16
'        txtCaminhoImpr.SetFocus
'        Exit Sub
    End If
    
    Call EscreveINI("Global", "X1", txtEscritorio.Text, App.Path & "\Config.ini")
    Call EscreveINI("Global", "X2", txtResponsavel.Text, App.Path & "\Config.ini")
    Call EscreveINI("Global", "X3", txtEndereco.Text, App.Path & "\Config.ini")
    Call EscreveINI("Global", "X4", TxtNro.Text, App.Path & "\Config.ini")
    Call EscreveINI("Global", "X5", txtCidade.Text, App.Path & "\Config.ini")
    Call EscreveINI("Global", "X6", txtCpf.Text, App.Path & "\Config.ini")
    

'    Call EscreveINI("Global", "Pesquisa", ChPesquisa.Value, App.Path & "\Config.ini")
'    Call EscreveINI(Trim(Maquina), "Impressora", txtCaminhoImpr.Text, App.Path & "\Config.ini")

    CfgNome = txtEscritorio.Text
'    CfgPesquisa = ChPesquisa.Value
'    CfgCaminhoImpr = txtCaminhoImpr.Text
    
    Unload Me

End Sub

Private Sub btReparar_Click()
On Error GoTo Linha

Dim listaTabelas As String
    
    SQL = "Show Tables from " & MyDb
    
    Set Consulta = New ADODB.Recordset
    
    Consulta.CursorLocation = LocalCursor
    Consulta.CacheSize = 30
    Consulta.Open SQL, StringConexao, adOpenKeyset, adLockReadOnly
    
    Do While Not Consulta.Eof
        If listaTabelas = "" Then
            listaTabelas = Consulta(0)
        Else
            listaTabelas = listaTabelas & "," & Consulta(0)
        End If
        
        Consulta.MoveNext
        
    Loop

Linha:
        If Err.Number = 0 Then
            Banco.Execute "Repair Table " & listaTabelas
            If Err.Number = 0 Then MsgBox "Base de Dados reparada com sucesso!", vbInformation
        Else
            MsgBox "Erro ao Reparar Base de Dados, Contate o SRural", 16
        End If

End Sub

Private Sub Form_DblClick()
    If InputBox("Digite a senha", "Senha para mudar cpf") = Cripta(Mid(txtCpf, 1, 3) & Mid(Time, 4, 2), 1) Then
        FrmConfigSistema.Show
        txtCpf.Enabled = True
        txtCpf.SetFocus
        txtCpf.SelLength = 5000
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SelecionaCampo
    End If
End Sub

Private Sub Form_Load()
    'txtEscritorio.Text = CfgNome
    
    txtEscritorio.Text = LerINI("Global", "X1", App.Path & "\Config.ini")
    txtResponsavel.Text = LerINI("Global", "X2", App.Path & "\Config.ini")
    txtEndereco.Text = LerINI("Global", "X3", App.Path & "\Config.ini")
    TxtNro.Text = LerINI("Global", "X4", App.Path & "\Config.ini")
    txtCidade.Text = LerINI("Global", "X5", App.Path & "\Config.ini")
    txtCpf.Text = LerINI("Global", "X6", App.Path & "\Config.ini")
    
    'Lbl.Caption = LerINI("Global", "Escritorio", App.Path & "\config.ini")
        
'    If IsNumeric(CfgPesquisa) Then ChPesquisa.Value = CfgPesquisa
'    txtCaminhoImpr.Text = LerINI(Maquina, "Impressora", App.Path & "\Config.ini")
    
    txtEscritorio.SelLength = 5000
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtCpf.Text = "" Then End
End Sub
