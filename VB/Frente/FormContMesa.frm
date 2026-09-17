VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormContMesa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisa"
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   990
      Width           =   1140
   End
   Begin VB.TextBox TxtMesa 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   225
      TabIndex        =   4
      Top             =   945
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   3
      Top             =   7785
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
      ScaleWidth      =   12930
      TabIndex        =   1
      Top             =   0
      Width           =   12930
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Movimentação de Mesa"
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
         Width           =   4695
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Caixa"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
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
      Bindings        =   "FormContMesa.frx":0000
      Height          =   6135
      Left            =   225
      OleObjectBlob   =   "FormContMesa.frx":001A
      TabIndex        =   0
      Top             =   1395
      Width           =   12390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa"
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   720
      Width           =   555
   End
End
Attribute VB_Name = "FormContMesa"
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

Private Sub Command1_Click()
If TxtMesa <> "" Then
    DatDados(2).RecordSource = "SELECT * FROM conv  WHERE Cliente = " & TxtMesa & " ORDER BY CodVda"
    DatDados(2).Refresh
    DatDados(2).Recordset.MoveLast
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

'Formcontmes.Left = (MDIPrincipal.Width - Me.Width) / 2
'Form.Top = (MDIPrincipal.Height - Me.Height) / 6

DatDados(2).RecordSource = "SELECT * FROM conv  ORDER BY CodVda"
DatDados(2).Refresh
DatDados(2).Recordset.MoveLast

End Sub

