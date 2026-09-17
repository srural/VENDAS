VERSION 5.00
Begin VB.Form FormIdentificador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   10
      Top             =   6840
      Width           =   11535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7935
      TabIndex        =   8
      Top             =   0
      Width           =   7935
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão de Etiqueta de Identificação"
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
         TabIndex        =   9
         Top             =   45
         Width           =   7170
      End
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6030
      TabIndex        =   5
      Top             =   900
      Width           =   1590
   End
   Begin VB.TextBox txtnome 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   225
      TabIndex        =   4
      Top             =   900
      Width           =   5685
   End
   Begin VB.CommandButton CmbDesmarcar 
      Caption         =   "&Desmarcar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   3
      Top             =   6165
      Width           =   1770
   End
   Begin VB.CommandButton CmbImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6300
      TabIndex        =   2
      Top             =   6165
      Width           =   1320
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   225
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1530
      Width           =   7395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6030
      TabIndex        =   7
      Top             =   675
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   675
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os produtos a serem impressos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   1
      Top             =   1260
      Width           =   2910
   End
End
Attribute VB_Name = "FormIdentificador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Etq As Recordset

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbDesmarcar_Click()
Dim y As Integer
    For y = 0 To List.ListCount - 1
          List.Selected(y) = False
    Next
End Sub

Private Sub CmbImprimir_Click()
Dim y As Integer
Dim Zebra As Boolean
Dim Formula As String
Dim Etq As Recordset
Dim Busca As String
Dim Fs, Arq
Dim VarArq As String
Dim VarPorta As String
Set Etq = Banco.OpenRecordset("SELECT ETQ.*, PRD.Descrição_Produto, ETQ.Codigo FROM ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd ORDER BY ETQ.Codigo")

For y = 0 To List.ListCount - 1
     If List.Selected(y) = False Then
         Busca = "Etq.Codigo = " & Left(List.List(y), 6)
         Etq.FindFirst Busca
         If Not Etq.NoMatch Then
             Etq.Delete
         End If
     End If
Next

FrmRelIdVolume.varNome.Caption = TxtNome
FrmRelIdVolume.VarCodNome.Caption = txtCod

Set varPesquisa(1) = Banco.OpenRecordset("SELECT ETQ.*, PRD.Descrição_Produto, ETQ.Codigo FROM ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd ORDER BY ETQ.Codigo")
Set FrmRelIdVolume.Relatorio.Recordset = varPesquisa(1)
FrmRelIdVolume.Config
Banco.Execute "Delete * From Etq"

Unload Me


'Exit Sub

'erro:
'    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub Form_Load()

Set Etq = Banco.OpenRecordset("SELECT ETQ.*, PRD.Descrição_Produto, ETQ.Codigo FROM ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd ORDER BY ETQ.Codigo")

Etq.Requery
If Etq.RecordCount <> 0 Then
    Etq.MoveFirst
    List.Clear
    While Not Etq.Eof
        List.AddItem psCompString(1, Etq.Fields("Etq.Codigo"), 6) & " - " & psCompString(1, Etq!Produto, 6) & " - " & Etq!Obs
        List.Selected(List.ListCount - 1) = True
        Etq.MoveNext
    Wend
End If

End Sub
