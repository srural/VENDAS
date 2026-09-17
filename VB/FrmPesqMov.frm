VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmPesqMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Movimento"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10815
   Begin MSMask.MaskEdBox MebDebito 
      Height          =   330
      Left            =   7357
      TabIndex        =   2
      Top             =   6615
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Data DatDados 
      Caption         =   "Conta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT  CodConta, Descricao FROM Conta ORDER BY Descricao"
      Top             =   135
      Visible         =   0   'False
      Width           =   1995
   End
   Begin MSDBCtls.DBCombo DbcConta 
      Bindings        =   "FrmPesqMov.frx":0000
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descricao"
      BoundColumn     =   "CodConta"
      Text            =   ""
   End
   Begin VB.Data DatDados 
      Caption         =   "Movimento"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   7695
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FrmPesqMov.frx":001A
      Top             =   4455
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmPesqMov.frx":015E
      Height          =   5640
      Left            =   135
      OleObjectBlob   =   "FrmPesqMov.frx":0178
      TabIndex        =   1
      Top             =   675
      Width           =   10545
   End
   Begin MSMask.MaskEdBox MebCredito 
      Height          =   330
      Left            =   5625
      TabIndex        =   3
      Top             =   6615
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebSaldo 
      Height          =   330
      Left            =   9090
      TabIndex        =   4
      Top             =   6615
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      Height          =   195
      Left            =   9090
      TabIndex        =   7
      Top             =   6390
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Débito"
      Height          =   195
      Left            =   7335
      TabIndex        =   6
      Top             =   6390
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Crédito"
      Height          =   195
      Left            =   5625
      TabIndex        =   5
      Top             =   6390
      Width           =   495
   End
End
Attribute VB_Name = "FrmPesqMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DbcConta_Change()
Dim SomaGrid As Recordset
Dim Credito As Currency
Dim Debito As Currency
If DbcConta.Text = "" Then
    DatDados(0).RecordSource = "SELECT Conta.Descricao, Cta.Descricao,Mov.Operacao, Mov.Historico, Mov.DtOper, Mov.Documento, Mov.VlrOper, Rem.Nome, Mov.Conta FROM Rem INNER JOIN (Cta INNER JOIN (Conta INNER JOIN Mov ON Conta.CodConta = Mov.Conta) ON Cta.CodCta = Mov.Operacao) ON Rem.CodRem = Mov.RemDest ORDER BY Mov.DtOper"
ElseIf IsNumeric(DbcConta.BoundText) Then
    DatDados(0).RecordSource = "SELECT Conta.Descricao, Cta.Descricao,Mov.Operacao, Mov.Historico, Mov.DtOper, Mov.Documento, Mov.VlrOper, Rem.Nome, Mov.Conta FROM Rem INNER JOIN (Cta INNER JOIN (Conta INNER JOIN Mov ON Conta.CodConta = Mov.Conta) ON Cta.CodCta = Mov.Operacao) ON Rem.CodRem = Mov.RemDest WHERE (((Mov.Conta)= " & DbcConta.BoundText & " )) ORDER BY Mov.DtOper"
End If
    DatDados(0).Refresh
    Set SomaGrid = DatDados(0).Recordset.Clone
    Do Until SomaGrid.EOF
         If Left(SomaGrid.Fields("Operacao"), 1) = "1" Then
             Credito = Credito + SomaGrid.Fields("VlrOper")
         Else
             Debito = Debito + SomaGrid.Fields("VlrOper")
         End If
        SomaGrid.MoveNext
    Loop
    MebCredito.Text = Credito
    MebDebito.Text = Debito
    MebSaldo = Credito - Debito

End Sub

Private Sub Form_Activate()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next
DatDados(1).Refresh
DbcConta_Change

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

End Sub

Private Sub Form_Load()
Set Formulário = Me
End Sub
