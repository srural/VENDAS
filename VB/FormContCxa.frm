VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormContCxa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11220
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   4
      Top             =   6570
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11940
      TabIndex        =   2
      Top             =   0
      Width           =   11940
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle do Caixa"
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
         Width           =   4695
      End
   End
   Begin VB.CommandButton BtExclui 
      Caption         =   "&Exclui"
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Top             =   5985
      Width           =   1335
   End
   Begin VB.Data DatDados 
      Caption         =   "Caixa"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   3645
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormContCxa.frx":0000
      Height          =   5010
      Left            =   225
      OleObjectBlob   =   "FormContCxa.frx":001A
      TabIndex        =   0
      Top             =   810
      Width           =   10500
   End
End
Attribute VB_Name = "FormContCxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtExclui_Click()
If MsgBox("Confirma Exlusão  ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(2).Recordset.RecordCount <> 0 Then
        
        Open Caminho & "\Delete.txt" For Append As #1
        Print #1, "Caixa; " & DatDados(2).Recordset.Fields(1) & "; " & DatDados(2).Recordset.Fields(2) & "; " & DatDados(2).Recordset.Fields(3) & "; " & DatDados(2).Recordset.Fields(4) & "; " & DatDados(2).Recordset.Fields(5) & "; " & DatDados(2).Recordset.Fields(6) & "; " & Date & "; " & Time & "; " & Maquina
        Close #1
        
        DatDados(2).Recordset.Delete
        DatDados(2).Refresh
        
        
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

DatDados(2).RecordSource = "SELECT * FROM CXA  ORDER BY DATA DESC, HORA DESC"
DatDados(2).Refresh


End Sub

