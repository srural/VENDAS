VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormIndices 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5235
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -630
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   4
      Top             =   5040
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
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   0
      Width           =   5235
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Indices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   3
         Top             =   45
         Width           =   4020
      End
   End
   Begin VB.CommandButton CmbExclui 
      Caption         =   "&Exclui"
      Height          =   375
      Left            =   3825
      TabIndex        =   1
      Top             =   4410
      Width           =   1140
   End
   Begin VB.Data Datdados 
      Caption         =   "Indices"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistemas em Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2025
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Ind.* FROM Ind  ORDER BY NroIndice"
      Top             =   2250
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgIndices 
      Bindings        =   "FormIndices.frx":0000
      Height          =   3525
      Left            =   180
      OleObjectBlob   =   "FormIndices.frx":001A
      TabIndex        =   0
      Top             =   720
      Width           =   4785
   End
End
Attribute VB_Name = "FormIndices"
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

Private Sub DbgIndices_AfterColUpdate(ByVal ColIndex As Integer)
If DbgIndices.Columns(2) <> "" Then
    DbgIndices.Columns(3) = ((CCur(DbgIndices.Columns(2)) + 100) / CCur(DbgIndices.Columns(1))) / 100
End If
End Sub

Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 ' (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

