VERSION 5.00
Begin VB.Form FormSenha2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   Icon            =   "FormSenha2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmbCancel 
      Cancel          =   -1  'True
      Height          =   600
      Left            =   1485
      Picture         =   "FormSenha2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar"
      Top             =   765
      Width           =   645
   End
   Begin VB.CommandButton CmdOk 
      Height          =   600
      Left            =   495
      Picture         =   "FormSenha2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Confirma"
      Top             =   765
      Width           =   645
   End
   Begin VB.TextBox TxtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   135
      Width           =   2220
   End
End
Attribute VB_Name = "FormSenha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cont As Integer
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
Dim PassWord As String
Dim Criterio As String

FormSenha = False

If TxtSenha = VarTelaSenha Then
    FormSenha = True
Else
    If VarTelaSenha = "USUARIO" Then
        Criterio = "Senha = " & """" & Trim(TxtSenha) & """"
        FrmFrente.RecSenha.FindFirst Criterio
        If Not FrmFrente.RecSenha.NoMatch Then
            FormSenha = True
        End If
    End If
End If
If FormSenha Then
    Open FrmFrente.Caminho & "\Senha.txt" For Append As #1
    Print #1, "FRENTE; Cliente; " & FrmFrente.Clientes!CodEntidade & "; " & Date & "; " & Time & "; " & FrmFrente.Maquina & " - " & FrmFrente.RecSenha!Nome
    Close #1
    Unload Me
Else
    MsgBox "Senha Inválida !", vbExclamation, App.Title
    If Cont < 2 Then
        TxtSenha = ""
        TxtSenha.SetFocus
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
