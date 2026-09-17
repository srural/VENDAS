VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormFamilia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   98.954
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   120.386
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   4
      Top             =   5265
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
      ScaleWidth      =   6855
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Familia de Produto"
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
         TabIndex        =   3
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.CommandButton CmbExclui 
      Caption         =   "&Exclui"
      Height          =   330
      Left            =   5220
      TabIndex        =   1
      Top             =   4725
      Width           =   1275
   End
   Begin VB.Data Datdados 
      Caption         =   "Familia"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Sistemas\Agroleite\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * FROM Fam  ORDER BY CodFam"
      Top             =   2970
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgIndices 
      Bindings        =   "FormFamilia.frx":0000
      Height          =   3795
      Left            =   315
      OleObjectBlob   =   "FormFamilia.frx":001A
      TabIndex        =   0
      Top             =   765
      Width           =   6135
   End
End
Attribute VB_Name = "FormFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbExclui_Click()
If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(0).Recordset.RecordCount <> 0 Then
        DatDados(0).Recordset.Delete
    End If
End If
End Sub

Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

