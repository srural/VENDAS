VERSION 5.00
Begin VB.Form FormObs 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -540
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   5
      Top             =   4140
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
      ScaleWidth      =   13155
      TabIndex        =   3
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   135
         TabIndex        =   4
         Top             =   45
         Width           =   5955
      End
   End
   Begin VB.CommandButton BtCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      Height          =   330
      Left            =   4140
      TabIndex        =   2
      Top             =   3645
      Width           =   1140
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   5490
      TabIndex        =   1
      Top             =   3645
      Width           =   1140
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   2715
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   765
      Width           =   6450
   End
End
Attribute VB_Name = "FormObs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtCancela_Click()
Unload Me
End Sub

Private Sub BtOK_Click()
FrmPedidos.VarObs = TxtObs
Unload Me
End Sub

