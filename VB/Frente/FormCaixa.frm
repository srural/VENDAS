VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormCaixa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   15
      Top             =   3240
      Width           =   11355
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6630
      TabIndex        =   12
      Top             =   0
      Width           =   6630
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Movimentação Caixa"
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
         TabIndex        =   13
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Movientação"
      Height          =   645
      Left            =   225
      TabIndex        =   6
      Top             =   2340
      Width           =   2445
      Begin VB.OptionButton Opt 
         Caption         =   "&Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1305
         TabIndex        =   4
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Hist"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2025
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT CXA.Historico FROM CXA GROUP BY CXA.Historico ORDER BY CXA.Historico"
      Top             =   1755
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   2520
      Width           =   1110
   End
   Begin MSMask.MaskEdBox MebData 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1110
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebValor 
      Height          =   375
      Left            =   2115
      TabIndex        =   1
      Top             =   1125
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSDBCtls.DBCombo DbcHist 
      Bindings        =   "FormCaixa.frx":0000
      Height          =   360
      Left            =   225
      TabIndex        =   2
      Top             =   1755
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Historico"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblMaquina 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   360
      Left            =   5490
      TabIndex        =   14
      Top             =   2700
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4095
      TabIndex        =   11
      Top             =   855
      Width           =   645
   End
   Begin VB.Label LblCaixa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Left            =   4950
      TabIndex        =   10
      Top             =   855
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   9
      Top             =   1890
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   8
      Top             =   855
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   855
      Width           =   435
   End
End
Attribute VB_Name = "FormCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
Dim sql As String
Dim Tipo As String

If Opt(1).Value = True Then
        Tipo = "C"
    Else
        Tipo = "D"
End If

If MebValor = "" Then
    MebValor.SetFocus
    Exit Sub
End If

If MsgBox("Confirma Lançamento ?", vbYesNo, App.Title) = vbYes Then
    sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    sql = sql + """" & LblMaquina & """" & ", "
    sql = sql & """" & Format(DbcHist, ">") & """" & " , #"
    sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    sql = sql & Format(Time, "hh:mm:ss") & "# , "
    sql = sql & Numero(MebValor.Text, 2) & " ,"
    sql = sql & LblCaixa & " ,"
    sql = sql & """" & Tipo & """" & ")"
    FrmFrente.Banco.Execute sql
    MebValor.Text = ""
    DbcHist.Text = ""
    DatDados(0).Refresh
    MebData.SetFocus
End If

End Sub

Private Sub DbcHist_LostFocus()
DbcHist.Text = Format(DbcHist, ">")
End Sub

Private Sub Form_Activate()
Dim caxia As Recordset
MebData.Text = Date
' DbcCaixa.Text = DatDados(1).Recordset.Fields("Maquina")

Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa WHERE Maquina =  " & """" & Trim(LimpaCampo(FrmFrente.Maquina)) & """" & "  Order by  CodControle")
Caixa.Requery

'If Caixa.RecordCount > 0 Then
'    Caixa.MoveLast
'    LblCaixa = Caixa!CodControle
'Else
    Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa  Order by  CodControle")
    Caixa.Requery
    If Caixa.RecordCount > 0 Then
        Caixa.MoveLast
        LblCaixa = Caixa!CodControle
    End If
'End If
LblMaquina = Trim(LimpaCampo(FrmFrente.Maquina))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

'Me.Left = 300
'Me.Top = 300

End Sub

Private Sub MebValor_KeyPress(KeyAscii As Integer)
If keyascci = 27 Then
    Unload Me
End If
End Sub
