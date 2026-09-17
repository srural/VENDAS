VERSION 5.00
Begin VB.Form FrmConsulta 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta Preços"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmConsulta.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtBarra 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   4050
      Width           =   2805
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   255
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4050
      Width           =   1095
   End
   Begin VB.Label LblCodigo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   360
      TabIndex        =   4
      Top             =   2745
      Width           =   4380
   End
   Begin VB.Label LblValor 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5175
      TabIndex        =   3
      Top             =   2700
      Width           =   3840
   End
   Begin VB.Label LblDesc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   405
      TabIndex        =   2
      Top             =   630
      Width           =   8565
   End
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub Form_Activate()
TxtBarra.SetFocus
End Sub

Private Sub Form_Load()
Flag = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Flag = 0
'FrmFrente.TxtDados(0).SetFocus
End Sub

Private Sub TxtBarra_KeyPress(KeyAscii As Integer)
Dim Terminação As String
Terminação = """"
If KeyAscii = 13 Then
     If IsNumeric(TxtBarra) Then
        Criterio = FrmFrente.Registros.Fields("BarBarra").Name & " = "
        Criterio = Criterio & Terminação & TxtBarra & Terminação
        FrmFrente.Registros.FindFirst Criterio
        If FrmFrente.Registros.NoMatch Then
              TxtBarra.Text = ""
              TxtBarra.SetFocus
              MsgBox "Informação não Localizada", vbExclamation, App.Title
              SendKeys "{Tab}"

        Else
            LblValor = Format(FrmFrente.Registros.Fields("Preco").Value, "#,##0.00")
            LblDesc = FrmFrente.Registros.Fields(1).Value
            LblCodigo = TxtBarra.Text
        End If
     End If
     CmdOk.SetFocus
End If
End Sub
