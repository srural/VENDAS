VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormIva 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   115.623
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   119.856
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -720
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   3
      Top             =   6210
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
      ScaleWidth      =   10005
      TabIndex        =   1
      Top             =   0
      Width           =   10005
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastra de I.V.A."
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
         TabIndex        =   2
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.Data Datdados 
      Caption         =   "Iva"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * FROM iva ORDER BY Ncm"
      Top             =   3735
      Visible         =   0   'False
      Width           =   2355
   End
   Begin MSDBGrid.DBGrid DbgIndices 
      Bindings        =   "FormIva.frx":0000
      Height          =   5190
      Left            =   135
      OleObjectBlob   =   "FormIva.frx":001A
      TabIndex        =   0
      Top             =   675
      Width           =   6270
   End
End
Attribute VB_Name = "FormIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbExclui_Click()
If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
    If Datdados(0).Recordset.RecordCount <> 0 Then
        Datdados(0).Recordset.Delete
    End If
End If
End Sub

Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

Me.Left = 300 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 '(MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

