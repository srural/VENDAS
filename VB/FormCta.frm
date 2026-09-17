VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormCta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   111.39
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   134.144
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   11
      Top             =   5940
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
      ScaleWidth      =   7620
      TabIndex        =   9
      Top             =   0
      Width           =   7620
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro do Plano de Conta"
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
         TabIndex        =   10
         Top             =   45
         Width           =   5505
      End
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5940
      TabIndex        =   8
      Top             =   5265
      Width           =   1275
   End
   Begin VB.CommandButton BtOk 
      Height          =   375
      Left            =   6840
      Picture         =   "FormCta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   765
      Width           =   375
   End
   Begin VB.CheckBox ChkTot 
      Caption         =   "&Totalizar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2070
      TabIndex        =   2
      Top             =   5310
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox TxtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      MaxLength       =   40
      TabIndex        =   1
      Top             =   855
      Width           =   4290
   End
   Begin MSMask.MaskEdBox MebCta 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##.###.###"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmbExclui 
      Caption         =   "&Exclui"
      Height          =   375
      Left            =   225
      TabIndex        =   5
      Top             =   5265
      Width           =   1275
   End
   Begin VB.Data Datdados 
      Caption         =   "Cta"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2745
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * FROM Cta  ORDER BY CodCta"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgCta 
      Bindings        =   "FormCta.frx":0532
      Height          =   3795
      Left            =   180
      OleObjectBlob   =   "FormCta.frx":054C
      TabIndex        =   4
      Top             =   1260
      Width           =   7080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1980
      TabIndex        =   7
      Top             =   675
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código Conta"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   675
      Width           =   960
   End
End
Attribute VB_Name = "FormCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtImprimir_Click()
MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelCta.rpt")
MDIPrincipal.RptRel.WindowTitle = "Relatório do Plano de Contas"
'MDIPrincipal.RptRel.SelectionFormula = "{Fam.CodFam} = " & DbcPesquisa.BoundText
MDIPrincipal.RptRel.WindowState = crptMaximized
Imprimir

End Sub

Private Sub BtOK_Click()
With DatDados(0).Recordset
    .AddNew
    .Fields("CodCta") = MebCta
    .Fields("Descric") = TxtDesc
    .Fields("Totaliza") = ChkTot.value
    .Update
    .Requery
    .MoveLast
    TxtDesc.Text = ""
    MebCta.SetFocus
End With
End Sub

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

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300  '(MDIPrincipal.Height - Me.Height) / 6

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

