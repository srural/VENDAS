VERSION 5.00
Begin VB.Form FrmLogo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   6120
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   2040
      Picture         =   "FrmLogo.frx":0000
      ScaleHeight     =   5055
      ScaleWidth      =   8055
      TabIndex        =   1
      Top             =   840
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   11160
      TabIndex        =   0
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "CAIXA  LIVRE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   6720
      Width           =   9480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "CAIXA  LIVRE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1665
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   6720
      Width           =   9480
   End
End
Attribute VB_Name = "FrmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Command1.SetFocus
End Sub

