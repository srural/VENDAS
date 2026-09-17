VERSION 5.00
Begin VB.Form FormFicha 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5580
      TabIndex        =   5
      Top             =   990
      Width           =   3120
   End
   Begin VB.TextBox TxtEvento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1035
      Width           =   4515
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9510
      TabIndex        =   1
      Top             =   0
      Width           =   9510
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Fichas"
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
         TabIndex        =   2
         Top             =   90
         Width           =   5325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   9480
      TabIndex        =   0
      Top             =   7515
      Width           =   9510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Evento"
      Height          =   330
      Left            =   315
      TabIndex        =   4
      Top             =   810
      Width           =   1410
   End
End
Attribute VB_Name = "FormFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

TxtData.Text = Date & " - " & Time

End Sub
