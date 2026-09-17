VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormPesquisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa"
   ClientHeight    =   4170
   ClientLeft      =   4680
   ClientTop       =   3675
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6360
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT Cli.Nome FROM Cli ORDER BY Cli.Nome"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdLocalizar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton CmdLocalizar 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin MSDBCtls.DBCombo DBCDados 
      Bindings        =   "FormPesquisa.frx":0000
      DataSource      =   "DatDados"
      Height          =   2910
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5133
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin VB.Label LblCampo 
      Caption         =   "Campo"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FormPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
With Formulário
    DatDados.DatabaseName = App.Path & "\Cadastro.mdb"
    DatDados.RecordSource = .PesquisaSql
    DatDados.Refresh
    DoEvents
    DbcDados.ListField = .LocalizacaoBasica
    DbcDados.BoundColumn = DatDados.Recordset.Fields(0).Name
    LblCampo.Caption = .LocalizacaoBasica
    DbcDados.SetFocus
   
End With
End Sub

Private Sub CmdLocalizar_Click(Index As Integer)
Dim Criterio As String
If Index = 0 Then
    Criterio = DatDados.Recordset.Fields(0).Name & " = " & DbcDados.BoundText
    DatDados.Recordset.FindFirst Criterio
    If Formulário.Registros.NoMatch Then
          MsgBox "Informação não Localizada", vbExclamation, App.Title
    Else
          MostraDados
          Unload Me
    End If
Else
    Unload Me
End If
End Sub

Private Sub Form_Load()
With Formulário
End With
End Sub
