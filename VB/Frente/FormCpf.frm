VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormCpf 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   5595
      Picture         =   "FormCpf.frx":0000
      ScaleHeight     =   1170
      ScaleWidth      =   1725
      TabIndex        =   2
      Top             =   330
      Width           =   1725
   End
   Begin MSMask.MaskEdBox MebDoc 
      Height          =   465
      Left            =   270
      TabIndex        =   0
      Top             =   810
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   820
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   4
      Height          =   1905
      Left            =   0
      Top             =   0
      Width           =   7755
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.F./C.N.P.J.  (somente numeros)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   270
      TabIndex        =   1
      Top             =   435
      Width           =   3315
   End
End
Attribute VB_Name = "FormCpf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Form_Activate()
'If Not IsNull(FrmFrente.Clientes.Fields("CgcCpf")) Then
'    MebDoc = FrmFrente.Clientes.Fields("CgcCpf")
'End If
'End Sub

Private Sub MebDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MebDoc <> "" Then
        If IsNumeric(LIMPACampo2(MebDoc)) Then
            If Len(Trim(LIMPACampo2(MebDoc))) = 11 Then
                MebDoc = Format(Trim(LIMPACampo2(MebDoc)), "@@@.@@@.@@@-@@")
                If Not CalculaDacCPF(MebDoc) Then
                    MsgBox "CPF Inválido", vbInformation
                    MebDoc.SetFocus
                Else
                    FormConsultaPed.VarDocumento = MebDoc
                    Unload Me
                End If
            ElseIf Len(Trim(LIMPACampo2(MebDoc))) = 14 Then
                MebDoc = Format(Trim(LIMPACampo2(MebDoc)), "@@.@@@.@@@/@@@@-@@")
                If Not VerificaCGC(MebDoc) Then
                    MsgBox "CNPJ Inválido.", vbInformation
                    MebDoc.SetFocus
                Else
                    FormConsultaPed.VarDocumento = MebDoc
                    Unload Me
                End If
            Else
                MsgBox "Documento invalido", vbInformation
                MebDoc.SetFocus
            End If
        End If
    Else
        Unload Me
    End If
ElseIf KeyAscii = 27 Then
    Unload Me
End If
End Sub
