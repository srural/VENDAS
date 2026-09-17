VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormFatura 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   13770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13770
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -405
      ScaleHeight     =   345
      ScaleWidth      =   14160
      TabIndex        =   12
      Top             =   7920
      Width           =   14190
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13785
      TabIndex        =   10
      Top             =   0
      Width           =   13785
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Faturamento"
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
         TabIndex        =   11
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   180
      TabIndex        =   7
      Tag             =   "N"
      Top             =   855
      Width           =   1095
   End
   Begin VB.TextBox TxtPedido 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   6255
      TabIndex        =   6
      Tag             =   "N"
      Top             =   855
      Width           =   1140
   End
   Begin VB.TextBox TxtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   1350
      TabIndex        =   5
      Tag             =   "N"
      Top             =   855
      Width           =   4830
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Confirma"
      Height          =   330
      Left            =   9585
      TabIndex        =   0
      Top             =   855
      Width           =   1320
   End
   Begin VB.Data DatDados 
      Caption         =   "Itp"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3915
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2655
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormFatura.frx":0000
      Height          =   6060
      Left            =   180
      OleObjectBlob   =   "FormFatura.frx":001A
      TabIndex        =   1
      Top             =   1395
      Width           =   13380
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Venda"
      Height          =   360
      Left            =   7605
      TabIndex        =   8
      Top             =   855
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label LblItens 
      Height          =   285
      Left            =   540
      TabIndex        =   13
      Top             =   7560
      Width           =   2805
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7605
      TabIndex        =   9
      Top             =   630
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   630
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6300
      TabIndex        =   3
      Top             =   630
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1395
      TabIndex        =   2
      Top             =   630
      Width           =   555
   End
End
Attribute VB_Name = "FormFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arq As Recordset
Public VarItem As Double

Private Sub BtOK_Click()
If Datdados(1).Recordset.RecordCount > 0 Then
    Datdados(1).Recordset.Edit
    Datdados(1).Recordset("Total") = MebValor
    Datdados(1).Recordset.Update
    FrmNota.MebValor = MebValor
End If
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub DatDados_Reposition(Index As Integer)
SomaGrid
LblItens.Caption = "Itens: " & Round((Datdados(1).Recordset.RecordCount * Datdados(1).Recordset.PercentPosition) / 100, 4) + 1 & "/" & Datdados(1).Recordset.RecordCount

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String

For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

TxtNome.Text = FrmNota.Entidade
TxtCodigo.Text = FrmNota.MebEntidade
TxtPedido.Text = FrmNota.CodPed
Datdados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.* FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & TxtPedido & " Order by Itp.CodItp"
Datdados(1).Refresh

If Datdados(1).Recordset.RecordCount > 0 Then
    Datdados(1).Recordset.MoveFirst
    VarItem = Datdados(1).Recordset("CodItp")
    MebValor = Datdados(1).Recordset.Fields("Total")
End If

LblItens.Caption = "Itens: " & 1 & "/" & Datdados(1).Recordset.RecordCount
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub SomaGrid()
Set Arq = Datdados(1).Recordset.Clone
MebValor = 0
If Arq.RecordCount > 0 Then
    Arq.MoveFirst
    While Not Arq.Eof
        MebValor = CCur(MebValor) + CCur(Arq!VlrVdaDia)
        Arq.MoveNext
    Wend
End If

End Sub

