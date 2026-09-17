VERSION 5.00
Begin VB.Form FormBoleto2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7110
   Begin VB.CommandButton Retorno 
      Caption         =   "Lê Retorno"
      Height          =   420
      Left            =   4770
      TabIndex        =   5
      Top             =   1755
      Width           =   1905
   End
   Begin VB.CommandButton Remessa 
      Caption         =   "Gera Remessa"
      Height          =   420
      Left            =   2565
      TabIndex        =   4
      Top             =   1755
      Width           =   1905
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "Limpa Lista"
      Height          =   420
      Left            =   450
      TabIndex        =   3
      Top             =   1755
      Width           =   1905
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   7860
      TabIndex        =   2
      Top             =   2610
      Width           =   7890
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Boletos"
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
         Left            =   180
         TabIndex        =   1
         Top             =   90
         Width           =   4650
      End
   End
   Begin VB.Label LblLista 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4410
      TabIndex        =   7
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Boletos na Lista:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   495
      TabIndex        =   6
      Top             =   945
      Width           =   3885
   End
End
Attribute VB_Name = "FormBoleto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLimpa_Click()
  If MsgBox("Deseja Limpar a Lista ?", vbYesNo) = vbYes Then
      AbreAcbr "Boleto.LimparLista"
      LeLista
      
  End If
End Sub

Private Sub Form_Load()

LeLista

End Sub

Private Sub Remessa_Click()
  If MsgBox("Deseja Gerar a Remessa ?", vbYesNo) = vbYes Then
      AbreAcbr "Boleto.GerarRemessa(" & "c:\SidComp\Remessa\" & ")"
      If MsgBox("Deseja Limpar a Lista ?", vbYesNo) = vbYes Then
            AbreAcbr "Boleto.LimparLista"
            LeLista
      End If
  
  
  End If


End Sub

Private Sub LeLista()
      If Dir("c:\SidComp\ACBrMonitorPLUS\Sai.txt") <> "" Then
            Kill "c:\SidComp\ACBrMonitorPLUS\Sai.txt"
      End If
      AbreAcbr "Boleto.TotalTitulosLista"
      LblLista = Mid$(TrataAcbr("Boleto", "1"), 4, 2)

End Sub
