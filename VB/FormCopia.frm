VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormCopia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   4545
   ClientTop       =   3675
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7320
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   9
      Top             =   3420
      Width           =   14370
   End
   Begin ComctlLib.ProgressBar BarraProgresso 
      Height          =   285
      Left            =   1395
      TabIndex        =   5
      Top             =   2025
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7350
      TabIndex        =   7
      Top             =   0
      Width           =   7350
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Copia de Segurança"
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
         TabIndex        =   8
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.TextBox txtZip 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   540
      TabIndex        =   3
      Top             =   2610
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.CommandButton CmdCopia 
      Caption         =   "&Inicia a Cópia"
      Height          =   420
      Left            =   3780
      TabIndex        =   2
      Top             =   945
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   1035
      Width           =   2265
   End
   Begin VB.Label LblProgresso 
      AutoSize        =   -1  'True
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
      Left            =   3735
      TabIndex        =   6
      Top             =   1710
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progresso"
      Height          =   195
      Left            =   555
      TabIndex        =   4
      Top             =   2070
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe o drive destino"
      Height          =   195
      Left            =   1170
      TabIndex        =   1
      Top             =   720
      Width           =   1605
   End
End
Attribute VB_Name = "FormCopia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdCopia_Click()
Dim ArqCompactado As String
Dim ArqOrigem As String

BarraProgresso.Min = 0
BarraProgresso.Max = 100
BarraProgresso.value = 0

ArqCompactado = Drive1.Drive & "\CopiaBco.zip"
ArqOrigem = Caminho & "\dados.mdb"


If Dir(ArqCompactado) <> "" Then
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set Arq = Fs.GetFile(ArqCompactado)
    Arq.Delete
End If

CmdCopia.Enabled = False

Compacta ArqCompactado, ArqOrigem

End Sub

Private Sub Form_Load()
Dim Formulario As Form
InicializaZip Me, txtZip

Set Formulario = Me

Formulario.Left = 400
Formulario.Top = 500

End Sub

Private Sub txtZip_Change()
' lblProgresso = TipoAção(Val(GetAction(txtZip))) & " "
' lblProgresso = lblProgresso & GetFileName(txtZip) & "->"
' lblProgresso = lblProgresso & GetPercentComplete(txtZip) & "%"

BarraProgresso.value = GetPercentComplete(txtZip)
LblProgresso = GetPercentComplete(txtZip) & " %"
DoEvents
If GetPercentComplete(txtZip) >= 99 Then
    CmdCopia.Enabled = True
    LblProgresso = ""
End If
End Sub
