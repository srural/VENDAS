VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmLivre 
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
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   90
      Top             =   7470
   End
   Begin MSMask.MaskEdBox MebPreco 
      Height          =   2055
      Left            =   1920
      TabIndex        =   4
      Top             =   6240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   -2147483641
      ForeColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   72
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "R$  #,##0.00;(R$ #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtProduto 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   3015
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FrmLivre.frx":0000
      Top             =   2610
      Width           =   11055
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   120
      Top             =   8400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   11160
      TabIndex        =   0
      Top             =   8520
      Width           =   735
   End
   Begin VB.Image Picture2 
      Height          =   7455
      Left            =   720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   10575
   End
   Begin VB.Image Picture1 
      Height          =   4935
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   960
      Width           =   8655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "CAIXA  LIVRE"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   1
      Top             =   6720
      Width           =   9480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "CAIXA  LIVRE"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   2
      Top             =   6720
      Width           =   9480
   End
End
Attribute VB_Name = "FrmLivre"
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

Private Sub Form_Load()

Picture1.Picture = LoadPicture(FrmFrente.Caminho & "\LIVRE.JPG")
Picture2.Picture = LoadPicture(FrmFrente.Caminho & "\LIVRE2.JPG")

Picture2.Visible = False

TxtProduto.Visible = False
MebPreco.Visible = False
'Label1(4).Visible = False

Picture1.Visible = True
Label1(0).Visible = True
Label1(1).Visible = True

FrmFrente.TxtDados(1).Text = ""
'FrmFrente.MebTotal.Text = ""

End Sub

Private Sub Timer1_Timer()
Dim Qtd As Long
If Picture1.Visible = True Then
    Picture1.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Picture2.Visible = True
ElseIf Picture2.Visible = True Then
        Picture2.Visible = False
        TxtProduto.Visible = True
        TxtProduto = "SEMPRE OS MELHORES PREÇOS"
        MebPreco = ""
'        Label1(4).Visible = True
    Else
        TxtProduto.Visible = False
        MebPreco.Visible = False
'        Label1(4).Visible = False
        Picture1.Visible = True
        Label1(0).Visible = True
        Label1(1).Visible = True
End If
End Sub

