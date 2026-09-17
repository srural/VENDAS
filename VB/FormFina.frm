VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormFina 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5955
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9840
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -405
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   13
      Top             =   5580
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
      ScaleWidth      =   9870
      TabIndex        =   11
      Top             =   0
      Width           =   9870
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Financeiro da Entidade"
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
         TabIndex        =   12
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.CheckBox ChkPago 
      Caption         =   "Mostrar os Pagos"
      Height          =   285
      Left            =   6570
      TabIndex        =   9
      Top             =   900
      Width           =   1725
   End
   Begin VB.TextBox TxtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   8
      Tag             =   "N"
      Top             =   825
      Width           =   6045
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   5220
      TabIndex        =   1
      Top             =   4545
      Width           =   4425
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   360
         Left            =   90
         TabIndex        =   2
         Tag             =   "S"
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPagos 
         DataField       =   "Custo"
         Height          =   360
         Left            =   1485
         TabIndex        =   3
         Tag             =   "S"
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebAPagar 
         DataField       =   "Custo"
         Height          =   330
         Left            =   2955
         TabIndex        =   4
         Tag             =   "S"
         Top             =   405
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   2835
         X2              =   2835
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1395
         X2              =   1395
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Abertas"
         Height          =   195
         Index           =   32
         Left            =   2865
         TabIndex        =   7
         Top             =   180
         Width           =   540
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Pagas"
         Height          =   195
         Index           =   33
         Left            =   1470
         TabIndex        =   6
         Top             =   180
         Width           =   450
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   34
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Dup"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   3735
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Dup"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgFina 
      Bindings        =   "FormFina.frx":0000
      Height          =   3180
      Left            =   135
      OleObjectBlob   =   "FormFina.frx":001A
      TabIndex        =   0
      Top             =   1305
      Width           =   9540
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   630
      Width           =   420
   End
End
Attribute VB_Name = "FormFina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbSair_Click()
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub ChkPago_Click()
If ChkPago.value = 0 Then
    DatDados(1).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & FrmEntidades.TxtDados(0).Text & " AND DUP.Cond = False ORDER BY DUP.DtVenc"
Else
    DatDados(1).RecordSource = "SELECT DUP.* FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade WHERE DUP.Entidade =" & FrmEntidades.TxtDados(0).Text & " ORDER BY DUP.DtVenc"
End If
DatDados(1).Refresh
SomaGrid
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

TxtNome.Text = FrmEntidades.TxtDados(1)
ChkPago_Click
Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6


End Sub
Private Sub SomaGrid()
Dim SomaGrid As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Pago = 0
Apagar = 0

'Set SomaGrid = DatDados(0).Recordset.Clone
If DatDados(1).Recordset.RecordCount <> 0 Then
    DatDados(1).Recordset.MoveFirst
    Do Until DatDados(1).Recordset.Eof
        If GetInputState() Then DoEvents

        If DatDados(1).Recordset.Fields("Cond") = True Then
            Pago = Pago + DatDados(1).Recordset.Fields("VlrPago")
        Else
            Apagar = Apagar + DatDados(1).Recordset.Fields("Valor")
        End If
       DatDados(1).Recordset.MoveNext
    Loop
    MebPagos.Text = Pago
    MebAPagar.Text = Apagar
    MebSubTot = Pago + Apagar

End If

End Sub

