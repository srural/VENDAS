VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormAlteraComissao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   8070
   Icon            =   "FormAlteraComissao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   8070
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   2
      Left            =   1620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAlteraComissao.frx":000C
      Top             =   5760
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   8040
      TabIndex        =   11
      Top             =   8595
      Width           =   8070
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   8040
      TabIndex        =   9
      Top             =   0
      Width           =   8070
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Comissão"
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
         Width           =   6630
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAlteraComissao.frx":00CC
      Top             =   4365
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1485
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade,Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
      Top             =   3420
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Index           =   0
      Left            =   150
      ScaleHeight     =   1365
      ScaleWidth      =   7875
      TabIndex        =   4
      Top             =   615
      Width           =   7875
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         Height          =   1095
         Left            =   4980
         TabIndex        =   6
         Top             =   180
         Width           =   1365
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Descricão"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   8
            Top             =   675
            Width           =   1110
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Código"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6435
         TabIndex        =   1
         Top             =   855
         Width           =   1065
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormAlteraComissao.frx":018C
         DataField       =   "Nome_Fornecedor"
         Height          =   945
         Left            =   45
         TabIndex        =   0
         Top             =   270
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   1667
         _Version        =   393216
         Appearance      =   0
         Style           =   1
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Parceiro"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   45
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAlteraComissao.frx":01A6
      Height          =   6330
      Left            =   180
      OleObjectBlob   =   "FormAlteraComissao.frx":01C0
      TabIndex        =   2
      Top             =   1980
      Width           =   7650
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   3
      Top             =   735
      Width           =   645
   End
End
Attribute VB_Name = "FormAlteraComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbSugestao_Click()
Dim SQL As String

SQL = "SELECT * FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd "
SQL = SQL + " Where PRD.Estoque < PRD.Mínimo and PRD.Mínimo <> 0 and Ent.CodEntidade = " & DbcPesquisa.BoundText & " Order By Descrição_Produto "

Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelSugestao.Relatorio.Recordset = varPesquisa(1)
FrmRelSugestao.rpCabTitulo.Caption = "Relatório de Sugestão de Compras por Fornecedor - " & DbcPesquisa.Text
FrmRelSugestao.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelSugestao.Config

End Sub

Private Sub CmdImprimir_Click()
Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim SQL As String

'MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelCotPrd.rpt")
'MDIPrincipal.RptRel.WindowTitle = "Relatório de Produtos por Fornecedor"
'MDIPrincipal.RptRel.SelectionFormula = "{Ent.CodEntidade} = " & DbcPesquisa.BoundText
'MDIPrincipal.RptRel.WindowState = crptMaximized
'Imprimir


    'CamposSql = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, ENT.Nome, ENT.CodEntidade FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, ENT.Nome, ENT.CodEntidade "
    'Ordenado = " ORDER BY Ent.Nome, PRD.Descrição_Produto "
    'If IsNumeric(DbcPesquisa.BoundText) Then CondicaoWhere = "Having Ent.CodEntidade = " & DbcPesquisa.BoundText
    
        'If ChkEstoqueZero.Value = 0 Then
            'Sql = CamposSql & CondicaoWhere & Ordenado
        'Else
            'Sql = CamposSql & Ordenado
        'End If
    If DbcPesquisa.Text = "" Then
        SQL = "SELECT Prd.CodBar, Prd.DtCompra, Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, Cot.Preço, Cot.Data FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & Space(0) & """" & ")) ORDER BY PRD.Descrição_Produto"
    Else
        SQL = "SELECT Prd.CodBar, PRD.DtCompra, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, COT.Preço, COT.Data, ENT.CodEntidade, ENT.Nome FROM ENT INNER JOIN (PRD INNER JOIN COT ON PRD.CodPrd = COT.CodPrd) ON ENT.CodEntidade = COT.CodFor GROUP BY PRD.DtCompra, PRD.CodPrd, PRD.Descrição_Produto,Prd.CodBar, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, COT.Preço, COT.Data, ENT.CodEntidade, ENT.Nome Having (((ENT.CodEntidade)=" & DbcPesquisa.BoundText & ")) ORDER BY PRD.Descrição_Produto"
'         Sql = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda,  Cot.Preço, Cot.Data, Ent.CodEntidade FROM PRD INNER JOIN (ENT INNER JOIN COT ON Ent.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & DbcPesquisa.Text & """" & ")) ORDER BY PRD.Descrição_Produto"
   
    End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    'Set varPesquisa(1) = DatDados(0).Recordset.Clone
     Set FrmRelPrdFor_Altera.Relatorio.Recordset = varPesquisa(1)
'    Set FrmRelPrdFor_Altera.Relatorio.Recordset = DatDados(0).Recordset.Clone
    
    FrmRelPrdFor_Altera.rpCabTitulo.Caption = "Relatório de Produtos por Fornecedor"
    FrmRelPrdFor_Altera.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPrdFor_Altera.Config

End Sub

Private Sub DbcPesquisa_Change()
Dim SQL As String
Dim Ordem As String
Dim VarSql As Recordset
Dim Criterio As String

DatDados(0).RecordSource = "SELECT * FROM COMIS"
DatDados(0).Refresh

DatDados(2).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, FAM.Descrição_Familia FROM FAM INNER JOIN (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) ON FAM.CodFam=GRU.Familia WHERE (((FAM.Descrição_Familia)= " & """" & "SERVICOS" & """" & "))"
DatDados(2).Refresh
If DatDados(2).Recordset.RecordCount > 0 And IsNumeric(DbcPesquisa.BoundText) Then
    DatDados(2).Recordset.MoveFirst
    While Not DatDados(2).Recordset.Eof
        Criterio = "CodFun = " & DbcPesquisa.BoundText & " and CodPrd = " & DatDados(2).Recordset("CodPrd")
        DatDados(0).Recordset.FindFirst Criterio
        If DatDados(0).Recordset.NoMatch Then
                DatDados(0).Recordset.AddNew
                DatDados(0).Recordset("CodPrd") = DatDados(2).Recordset("CodPrd")
                DatDados(0).Recordset("CodFun") = DbcPesquisa.BoundText
                DatDados(0).Recordset.Update
        End If
        DatDados(2).Recordset.MoveNext
    Wend
End If

If Not IsNumeric(DbcPesquisa.BoundText) Then
    SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, COMIS.Comissao, COMIS.CodFun, COMIS.CodPrd FROM PRD INNER JOIN COMIS ON PRD.CodPrd = COMIS.CodPrd WHERE (((COMIS.CodFun)=1.1)) "
Else
    SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, COMIS.Comissao, COMIS.CodFun, COMIS.CodPrd FROM PRD INNER JOIN COMIS ON PRD.CodPrd = COMIS.CodPrd WHERE (((COMIS.CodFun)= " & DbcPesquisa.BoundText & ")) "
End If
If OptOrdem(0).value = True Then
    Ordem = " ORDER BY PRD.CodPrd"
Else
    Ordem = " ORDER BY PRD.Descrição_Produto"
End If

DatDados(0).RecordSource = SQL + Ordem
DatDados(0).Refresh


If DatDados(0).Recordset.RecordCount > 0 Then
'    CmdImprimir.Enabled = True
Else
    CmdImprimir.Enabled = False
End If


End Sub

Private Sub DbcPesquisa_LostFocus()
Dim Criterio As String
If IsNumeric(DbcPesquisa.Text) Then
   Criterio = "CodEntidade = " & DbcPesquisa.Text
   DatDados(1).Recordset.FindFirst Criterio
   If DatDados(1).Recordset.NoMatch Then
       MsgBox "Informação não Localizada", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
        DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
        DBGrid1.SetFocus
   End If
End If

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then SendKeys "{Tab}"

End Sub


Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 300
Formulário.Top = 0

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
DbcPesquisa.SetFocus
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub


Private Sub OptOrdem_Click(Index As Integer)
DbcPesquisa_Change
End Sub
