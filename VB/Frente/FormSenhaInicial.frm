VERSION 5.00
Begin VB.Form FormSenhaInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Senha"
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6855
   Icon            =   "FormSenhaInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormSenhaInicial.frx":0442
   ScaleHeight     =   3660
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4725
      Picture         =   "FormSenhaInicial.frx":3E01
      ScaleHeight     =   600
      ScaleWidth      =   2040
      TabIndex        =   6
      ToolTipText     =   "Duplo click para chamar o suporte"
      Top             =   2925
      Width           =   2040
   End
   Begin VB.TextBox Usuario 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      TabIndex        =   0
      Top             =   810
      Width           =   1725
   End
   Begin VB.CommandButton CmbCancel 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   3915
      Picture         =   "FormSenhaInicial.frx":739C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   1710
      Width           =   555
   End
   Begin VB.CommandButton CmdOk 
      Height          =   510
      Left            =   3150
      Picture         =   "FormSenhaInicial.frx":77DE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Confirma"
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox TxtSenha 
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1260
      Width           =   1725
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
      Left            =   2070
      TabIndex        =   5
      Top             =   1140
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   1980
      Picture         =   "FormSenhaInicial.frx":7C20
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   645
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
      Left            =   2070
      TabIndex        =   4
      Top             =   810
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1695
      Left            =   1845
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   2970
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "FormSenhaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cont As Integer
Private Sub CmbCancel_Click()
Unload Me
formasenha = False
End Sub

Private Sub CmdOk_Click()
' [QUANTIDADE] * [VALOR_UNITARIO]
Dim PassWord As String
Dim Criterio As String

FormSenha = False

Criterio = "Nome = " & """" & Trim(Usuario) & """" & " and Senha = " & """" & Trim(TxtSenha) & """"
FrmFrente.DatSenha.FindFirst Criterio
If Not FrmFrente.DatSenha.NoMatch Then
    FormSenha = True
    Unload Me
Else
    MsgBox "Senha Inválida !", vbExclamation, App.Title
    If Cont < 2 Then
        TxtSenha = ""
        Usuario.SetFocus
        Cont = Cont + 1
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Activate()
Cont = 1
FormSenha = False

End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdOk_Click
End If
End Sub

Private Sub Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtSenha.SetFocus
End If
End Sub
