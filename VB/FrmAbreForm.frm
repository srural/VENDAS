VERSION 5.00
Begin VB.Form FrmAbreForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Index           =   0
      Interval        =   200
      Left            =   4770
      Top             =   1200
   End
   Begin VB.Line Line4 
      X1              =   405
      X2              =   3855
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line3 
      X1              =   375
      X2              =   3825
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   375
      X2              =   3825
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   405
      X2              =   3855
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde ... Carregando Formulário ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   240
      Width           =   3225
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde ... Carregando Formulário ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   1
      Top             =   240
      Width           =   3225
   End
End
Attribute VB_Name = "FrmAbreForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormX As Form

Private Sub Form_Load()
    Set Formulário = Me
    CentralizarForm Me, 1
End Sub

Private Sub Timer_Timer(Index As Integer)
Select Case Index
    Case 0
        CarregaForm
        Unload Me
End Select
End Sub

Sub CarregaForm()
    Timer(0).Interval = 0
    FormX.Show
End Sub
