VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmCompac 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7185
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -225
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   6
      Top             =   2430
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
      TabIndex        =   4
      Top             =   0
      Width           =   14655
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Compactar / Reparar  Banco de Dados"
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
         Left            =   90
         TabIndex        =   5
         Top             =   45
         Width           =   6630
      End
   End
   Begin ComctlLib.ProgressBar Progresso 
      Height          =   255
      Left            =   1395
      TabIndex        =   0
      Top             =   1380
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Max             =   11
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Iniciar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde. . . Reparando Banco de dados"
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
      Left            =   1665
      TabIndex        =   1
      Top             =   990
      Width           =   4200
   End
End
Attribute VB_Name = "FrmCompac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "&Iniciar" Then
    Command1.Enabled = False
    Label1.Caption = "Aguarde. . . Reparando Banco de dados"
    CompactarReparar
    Command1.Enabled = True
    Command1.SetFocus
    Command1.Caption = "OK"
    Label1.Caption = "Banco de Dados Reparado com Sucesso"
Else
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
FrmCompac.Left = 2500
FrmCompac.Top = 1000
Label1.Caption = "Certifique-se de que nunhum usuário " & vbCr & " está acessando o Banco de Dados"
End Sub

