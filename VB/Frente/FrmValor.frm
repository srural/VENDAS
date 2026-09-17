VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmValor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   3930
   ClientTop       =   2925
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   3975
      TabIndex        =   4
      Top             =   0
      Width           =   3975
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Unitário"
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
         TabIndex        =   5
         Top             =   90
         Width           =   2490
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   3990
      TabIndex        =   3
      Top             =   2610
      Width           =   4020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   1395
      TabIndex        =   1
      Top             =   1845
      Width           =   1290
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Custo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   435
      Left            =   675
      TabIndex        =   2
      Tag             =   "S"
      Top             =   1170
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o Valor Unitário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   645
      TabIndex        =   0
      Top             =   840
      Width           =   2670
   End
End
Attribute VB_Name = "FrmValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MebValor = "" Then
    MebValor = ""
    MebValor.SetFocus
    
Else
    ValorUnit = CStr(MebValor)
    If MebValor > 1000 Then
        If MsgBox("Valor maior que  R$ 1.000,00 , confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            FrmFrente.MebUnit = Format(ValorUnit, "#,##0.00;- #,##0.00")
            Unload Me
        Else
            MebValor = ""
            MebValor.SetFocus
        End If
    Else
        FrmFrente.MebUnit = Format(ValorUnit, "#,##0.00;- #,##0.00")
        Unload Me
    End If
End If
End Sub

Private Sub Form_Activate()
MebValor = ""
MebValor.SetFocus
End Sub


Private Sub MebValor_KeyPress(KeyAscii As Integer)
If KeyAscii = Chr(13) Then
    Command1_Click
End If
End Sub
