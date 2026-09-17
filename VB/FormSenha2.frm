VERSION 5.00
Begin VB.Form FormSenha2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "FormSenha2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmbCancel 
      Cancel          =   -1  'True
      Height          =   600
      Left            =   1755
      Picture         =   "FormSenha2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar"
      Top             =   900
      Width           =   645
   End
   Begin VB.CommandButton CmdOk 
      Height          =   600
      Left            =   765
      Picture         =   "FormSenha2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Confirma"
      Top             =   900
      Width           =   645
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
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
      Left            =   450
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   270
      Width           =   2220
   End
End
Attribute VB_Name = "FormSenha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cont As Integer
Public Tela As String
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
Dim PassWord As String
Dim VARPASS As Recordset
Dim Criterio As String

If Tela = "" Then
    Set VARPASS = Banco.OpenRecordset("SELECT * FROM PWD WHERE NOME = 'ALTPEDIDO'")
    VARPASS.Requery
    If VARPASS.RecordCount = 1 Then
        PassWord = Trim(VARPASS!Senha)
    Else
        PassWord = Trim(Configuracao.Fields("Senha2"))
    End If
    
    If TxtSenha = Left(Time, 2) & Mid(Time, 4, 2) Then
        PassWord = TxtSenha
    End If

    FrmPedidos.Senha = False
    If TxtSenha = PassWord Then
        FrmPedidos.Senha = True
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

Else
    Set VARPASS = Banco.OpenRecordset("SELECT * FROM PWD WHERE NOME = 'EXCPEDIDO'")
    VARPASS.Requery
    If VARPASS.RecordCount = 1 Then
        PassWord = Trim(VARPASS!Senha)
        
    
    Else
        PassWord = Trim(Configuracao.Fields("Senha2"))
    End If

 '   FrmParcelas.Senha = False
    
    If TxtSenha = PassWord Then
        FrmPedidos.Senha = True
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
    
    

End If

End Sub

Private Sub Form_Activate()
Cont = 1
FrmPedidos.Senha = False

End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdOk_Click
End If
End Sub
