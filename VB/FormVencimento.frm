VERSION 5.00
Begin VB.Form FormVencimento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   6
      Top             =   4005
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
      ScaleWidth      =   14430
      TabIndex        =   4
      Top             =   0
      Width           =   14430
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Vencimentos"
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
         Width           =   6360
      End
   End
   Begin VB.TextBox TxtVenc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   345
      TabIndex        =   0
      Tag             =   "N"
      Top             =   1080
      Width           =   7815
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      ItemData        =   "FormVencimento.frx":0000
      Left            =   345
      List            =   "FormVencimento.frx":0002
      TabIndex        =   1
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Os vencimentos tem que estar separados por / Ex: 30/60/90"
      Height          =   195
      Left            =   345
      TabIndex        =   3
      Top             =   3495
      Width           =   4680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venciamento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   345
      TabIndex        =   2
      Top             =   840
      Width           =   1230
   End
End
Attribute VB_Name = "FormVencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim VarTexto As String

        If Dir(App.Path & "\Vencimentos.Txt") <> "" Then
            On Error GoTo ErroObs
            Lista.Clear
            Lista.Clear
            Open App.Path & "\vencimentos.Txt" For Input As #1
            Do
                Input #1, VarTexto
                Lista.AddItem VarTexto
            Loop Until VarTexto = ""
            Close #1
ErroObs:
            If Err.Number = 62 Then
                Close #1
            End If
        End If

End Sub


Private Sub TxtVenc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Lista.AddItem TxtVenc
    Open App.Path & "\vencimentos.Txt" For Output As #1
    For x = 0 To Lista.ListCount - 1
        Print #1, Lista.List(x)
    Next
    Close #1
    TxtVenc = ""
    TxtVenc.SetFocus
End If

End Sub

