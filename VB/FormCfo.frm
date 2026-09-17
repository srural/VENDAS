VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormCfo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   98.69
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   154.517
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   4
      Top             =   5220
      Width           =   14370
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8790
      TabIndex        =   2
      Top             =   0
      Width           =   8790
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de C.f.Op"
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
         Width           =   4020
      End
   End
   Begin VB.CommandButton CmbExclui 
      Caption         =   "&Exclui"
      Height          =   285
      Left            =   7290
      TabIndex        =   1
      Top             =   4815
      Width           =   1275
   End
   Begin VB.Data Datdados 
      Caption         =   "CFOp"
      Connect         =   "Access"
      DatabaseName    =   "C:\Rede\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * FROM Cfo  ORDER BY Codigo"
      Top             =   3150
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgIndices 
      Bindings        =   "FormCfo.frx":0000
      Height          =   4020
      Left            =   135
      OleObjectBlob   =   "FormCfo.frx":001A
      TabIndex        =   0
      Top             =   675
      Width           =   8430
   End
End
Attribute VB_Name = "FormCfo"
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

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300  '(MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

