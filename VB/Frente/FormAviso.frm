VERSION 5.00
Begin VB.Form FormAviso 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   90
      Top             =   90
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   6105
   End
End
Attribute VB_Name = "FormAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
