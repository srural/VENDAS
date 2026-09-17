VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConfig 
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "FrmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   11
      Top             =   3600
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
      TabIndex        =   9
      Top             =   0
      Width           =   14655
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Registro do Sistema"
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
         TabIndex        =   10
         Top             =   45
         Width           =   4920
      End
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "Confirma"
      Height          =   375
      Left            =   3165
      TabIndex        =   1
      Top             =   3075
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informe esses valores"
      Enabled         =   0   'False
      Height          =   930
      Left            =   765
      TabIndex        =   3
      Top             =   1230
      Width           =   3750
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   285
         Left            =   1455
         TabIndex        =   5
         Top             =   195
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtComplemento 
         Height          =   285
         Left            =   1455
         TabIndex        =   7
         Top             =   525
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   -2147483641
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   285
         TabIndex        =   8
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cpf:"
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   195
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chave"
      Height          =   660
      Left            =   795
      TabIndex        =   2
      Top             =   2325
      Width           =   3690
      Begin MSMask.MaskEdBox Chave 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   210
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   21
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC.CC.CC.CC.CC.CC-CCC"
         PromptChar      =   "_"
      End
   End
   Begin VB.Label txtMsg 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   750
      TabIndex        =   4
      Top             =   765
      Width           =   3795
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechaAoSair As Boolean
Public varMensagem As String
Dim x As Integer
'Tipo da Copia
'0=Manut
'1=Anual
'2=Teste
'3=Demo

Private Sub BtNroSerial_Click()
Dim lNSerial As Long
lNSerial = DriveSerial(App.Path & "\")

MsgBox "O número seirial do HD é: " & lNSerial

End Sub

Private Sub BtOK_Click()
On Error Resume Next
'Dim NumChave As String
Dim Num As String
Dim NumCompl As String
Dim Erro As Integer

Erro = 0

    If Chave.Text = "" Then
        Erro = 1
    Else
        NumCompl = GeraComplemento(txtCpf.Text, txtComplemento)
    
        Num = Cripta(Chave.Text, 2)
        If GeraDigito(CCur(Mid(Num, 1, Len(Num) - 2))) = Mid(Num, Len(Num) - 1, 2) And Mid(Num, 1, 4) = NumCompl Then
            'txtLimite.Text = CCur(Mid(Num, 5, 4))
            'cmbCopia.ListIndex = CCur(Mid(Num, 9, 1)) - 1
            
            SaveSetting App.EXEName, "SystemReport", "Sys01", EnCripta(DriveSerial(Mid(App.Path, 1, 3)))  'nro serie hd
            Call EscreveINI("Global", "X9", CCur(Mid(Num, 5, 4)), App.Path & "\Config.ini") 'dias
            Call EscreveINI("Global", "X10", CCur(Mid(Num, 9, 1)), App.Path & "\Config.ini") 'tipo da copia
            Call EscreveINI("Global", "X11", Date, App.Path & "\Config.ini") 'aceso
            Call EscreveINI("Global", "X12", Date, App.Path & "\Config.ini") 'aceso
            SaveSetting App.EXEName, "SystemReport", "Sys02", EnCripta(txtCpf.Text) 'cpf
            
            'Unload Me
            
        Else
            Erro = 1
        End If
    End If
   
If Err.Number <> 0 Or Erro = 1 Then
    MsgBox "Código de Reset Inválido!", 16
    Unload Me
Else
    MsgBox "Reset executado com sucesso!", vbInformation
    FechaAoSair = False
    Unload Me
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    'Me.Top = ((Screen.Height) / 3 - Me.Height / 2)
    'Me.Left = Screen.Width / 2 - Me.Width / 2

txtMsg.Caption = varMensagem
txtCpf.Text = LerINI("Global", "X6", App.Path & "\Config.ini")

LinhaInicio:
    txtComplemento = Format(Mid(Int(CCur(Mid(Time, 7, 2)) * (Rnd * 1000)), 1, 4), "0000")
    If txtComplemento = "0000" Then GoTo LinhaInicio
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If FechaAoSair = True Then
        End
     End If
End Sub
