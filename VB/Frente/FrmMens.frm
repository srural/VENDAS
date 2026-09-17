VERSION 5.00
Begin VB.Form FrmMens 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1680
      Top             =   120
   End
   Begin VB.Label Mens2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Mens 
      AutoSize        =   -1  'True
      Caption         =   "Erro Código : "
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "FrmMens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
