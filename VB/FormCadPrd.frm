VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormCadPrd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Simples de Produto"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   91.281
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   131.233
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6030
      TabIndex        =   8
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton BtOk 
      Height          =   375
      Left            =   6840
      Picture         =   "FormCadPrd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   375
   End
   Begin VB.CheckBox ChkTot 
      Caption         =   "&Totalizar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox TxtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      MaxLength       =   40
      TabIndex        =   1
      Top             =   270
      Width           =   4290
   End
   Begin MSMask.MaskEdBox MebCta 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   270
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
      Top             =   4635
      Width           =   1275
   End
   Begin VB.Data Datdados 
      Caption         =   "Prd"
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
      RecordSource    =   "Select * FROM Prd  ORDER BY CodPrd"
      Top             =   2295
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DbgCta 
      Bindings        =   "FormCadPrd.frx":0532
      Height          =   3795
      Left            =   180
      OleObjectBlob   =   "FormCadPrd.frx":054C
      TabIndex        =   4
      Top             =   675
      Width           =   7080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1980
      TabIndex        =   7
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Conta"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "FormCadPrd"
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
With Datdados(0).Recordset
    .AddNew
    .Fields("CodCta") = MebCta
    .Fields("Descric") = TxtDesc
    .Fields("Totaliza") = ChkTot.Value
    .Update
    .Requery
    .MoveLast
    TxtDesc.Text = ""
    MebCta.SetFocus
End With
End Sub

Private Sub CmbExclui_Click()
If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
    If Datdados(0).Recordset.RecordCount <> 0 Then
        Datdados(0).Recordset.Delete
    End If
End If
End Sub

Private Sub Form_Activate()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

