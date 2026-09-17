VERSION 5.00
Begin VB.Form FormData 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   12
      Top             =   2925
      Width           =   11355
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4920
      TabIndex        =   10
      Top             =   0
      Width           =   4920
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Data / Hora"
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
         TabIndex        =   11
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.TextBox TxtMin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1545
      TabIndex        =   7
      Top             =   2070
      Width           =   915
   End
   Begin VB.TextBox TxtHora 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   6
      Top             =   2070
      Width           =   915
   End
   Begin VB.TextBox TxtAno 
      Appearance      =   0  'Flat
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
      Left            =   2655
      TabIndex        =   3
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox TxtMes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1500
      TabIndex        =   2
      Top             =   1080
      Width           =   825
   End
   Begin VB.TextBox TxtDia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   360
      TabIndex        =   9
      Top             =   1755
      Width           =   600
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1350
      TabIndex        =   8
      Top             =   2025
      Width           =   165
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2430
      TabIndex        =   5
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1260
      TabIndex        =   4
      Top             =   1080
      Width           =   165
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   765
      Width           =   600
   End
End
Attribute VB_Name = "FormData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo erro
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If

erro:

If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If
End Sub


Private Sub Form_Load()
TxtDia = Day(Date)
TxtMes = Month(Date)
TxtAno = Year(Date)

TxtHora = Mid(Time, 1, 2)
TxtMin = Mid(Time, 4, 2)
End Sub


Private Sub TxtHora_LostFocus()
If TxtHora <> "" Then
    If TxtHora > 24 Then
        TxtHora.SetFocus
    End If
Else
    TxtHora.SetFocus
End If
End Sub

Private Sub TxtMin_LostFocus()
If TxtMin <> "" Then
    If TxtMin < 59 Then
        If MsgBox("Confirma data / hora ? ", vbYesNo, App.Title) = vbYes Then
            If IsDate(Format(TxtDia & "/" & TxtMes & "/" & TxtAno, "dd/mm/yyyy")) Then
                Date = Format(TxtDia & "/" & TxtMes & "/" & TxtAno, "dd/mm/yyyy")
                Time = Format(TxtHora & ":" & TxtMin, "hh:mm")
                Unload Me
            Else
                MsgBox "Data Inválida", vbExclamation, App.Title
                TxtDia.SetFocus
            End If
        Else
            Unload Me
        End If
    Else
        TxtMin.SetFocus
    End If
Else
    TxtHora.SetFocus
End If

End Sub
