VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormAlteraPreços 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14715
   Icon            =   "FormAlteraPrecos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   14715
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14700
      TabIndex        =   19
      Top             =   8595
      Width           =   14730
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   14745
      TabIndex        =   17
      Top             =   0
      Width           =   14775
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Alteração de Preços por Fornecedor"
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
         TabIndex        =   18
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   6435
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAlteraPrecos.frx":000C
      Top             =   4770
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2880
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
      Top             =   4770
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
      ScaleWidth      =   14355
      TabIndex        =   6
      Top             =   615
      Width           =   14355
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         Height          =   1095
         Left            =   6690
         TabIndex        =   14
         Top             =   180
         Width           =   1365
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Descricão"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   16
            Top             =   675
            Width           =   1110
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Código"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   315
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.CommandButton CmbSugestao 
         Caption         =   "&Sugestão Compra"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12735
         TabIndex        =   13
         Top             =   330
         Width           =   1425
      End
      Begin VB.Frame Frame2 
         Caption         =   "Altera Preços"
         Height          =   1095
         Left            =   10755
         TabIndex        =   9
         Top             =   180
         Width           =   1605
         Begin VB.CommandButton BtOk 
            Caption         =   "&OK"
            Height          =   270
            Left            =   180
            TabIndex        =   11
            Top             =   750
            Width           =   1140
         End
         Begin MSMask.MaskEdBox MebPerc 
            Height          =   300
            Left            =   150
            TabIndex        =   10
            Top             =   420
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual"
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   225
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Calcular"
         Height          =   1095
         Left            =   8310
         TabIndex        =   8
         Top             =   180
         Width           =   1245
         Begin VB.OptionButton Opt 
            Caption         =   "&Venda"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   1
            Top             =   315
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Margem"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   675
            Width           =   960
         End
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12735
         TabIndex        =   3
         Top             =   855
         Width           =   1425
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormAlteraPrecos.frx":012D
         DataField       =   "Nome_Fornecedor"
         Height          =   945
         Left            =   45
         TabIndex        =   0
         Top             =   270
         Width           =   6390
         _ExtentX        =   11271
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
         Caption         =   "&Nome do Fornecedor"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   45
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAlteraPrecos.frx":0147
      Height          =   6330
      Left            =   180
      OleObjectBlob   =   "FormAlteraPrecos.frx":0161
      TabIndex        =   4
      Top             =   1980
      Width           =   14355
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   5
      Top             =   735
      Width           =   645
   End
End
Attribute VB_Name = "FormAlteraPreços"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub BtOK_Click()
If MebPerc <> "" Then
    If IsNumeric(MebPerc) And DatDados(0).Recordset.RecordCount > 0 Then
        If MsgBox("Atenção esta operação altera todos os preços deste fornecedor. Confirma ?", vbYesNo, App.Title) = vbYes Then
            With DatDados(0).Recordset
                .MoveFirst
                While Not .Eof
                    .Edit
                    .Fields("Custo") = .Fields("Custo") + (.Fields("Custo") * (MebPerc / 100))
                    .Fields("Venda") = .Fields("Venda") + (.Fields("Venda") * (MebPerc / 100))
                    If VarPzoProduto = True Then
                        .Fields("PcoPrazo") = .Fields("PcoPrazo") + (.Fields("PcoPrazo") * (MebPerc / 100))
                    End If
                    .Update
                    .MoveNext
                Wend
                MebPerc = ""
                MsgBox "Atualzação de preços concluída !", vbInformation, App.Title
            End With
        End If
    End If
End If
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbSugestao_Click()
Dim Sql As String

Sql = "SELECT * FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd "
Sql = Sql + " Where PRD.Estoque < PRD.Mínimo and PRD.Mínimo <> 0 and Ent.CodEntidade = " & DbcPesquisa.BoundText & " Order By Descrição_Produto "

Set varPesquisa(1) = Banco.OpenRecordset(Sql)
Set FrmRelSugestao.Relatorio.Recordset = varPesquisa(1)
FrmRelSugestao.rpCabTitulo.Caption = "Relatório de Sugestão de Compras por Fornecedor - " & DbcPesquisa.Text
FrmRelSugestao.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelSugestao.Config

End Sub

Private Sub CmdImprimir_Click()
Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim Sql As String

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
        Sql = "SELECT Prd.CodBar, Prd.DtCompra, Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, Cot.Preço, Cot.Data FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & Space(0) & """" & ")) ORDER BY PRD.Descrição_Produto"
    Else
        Sql = "SELECT Prd.CodBar, PRD.DtCompra, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, COT.Preço, COT.Data, ENT.CodEntidade, ENT.Nome FROM ENT INNER JOIN (PRD INNER JOIN COT ON PRD.CodPrd = COT.CodPrd) ON ENT.CodEntidade = COT.CodFor GROUP BY PRD.DtCompra, PRD.CodPrd, PRD.Descrição_Produto,Prd.CodBar, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, COT.Preço, COT.Data, ENT.CodEntidade, ENT.Nome Having (((ENT.CodEntidade)=" & DbcPesquisa.BoundText & ")) ORDER BY PRD.Descrição_Produto"
'         Sql = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda,  Cot.Preço, Cot.Data, Ent.CodEntidade FROM PRD INNER JOIN (ENT INNER JOIN COT ON Ent.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & DbcPesquisa.Text & """" & ")) ORDER BY PRD.Descrição_Produto"
   
    End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
    'Set varPesquisa(1) = DatDados(0).Recordset.Clone
     Set FrmRelPrdFor_Altera.Relatorio.Recordset = varPesquisa(1)
'    Set FrmRelPrdFor_Altera.Relatorio.Recordset = DatDados(0).Recordset.Clone
    
    FrmRelPrdFor_Altera.rpCabTitulo.Caption = "Relatório de Produtos por Fornecedor"
    FrmRelPrdFor_Altera.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPrdFor_Altera.Config

End Sub

Private Sub DbcPesquisa_Change()
Dim Sql As String
Dim Ordem As String

If DbcPesquisa.Text = "" Then
    Sql = "SELECT  Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, Prd.PcoPrazo, PRD.Venda, Cot.Preço, Cot.Data FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & Space(0) & """" & ")) "
Else
    Sql = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Custo, PRD.Margem, PRD.Venda, Prd.PcoPrazo,  Cot.Preço, Cot.Data FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd WHERE (((ENT.Nome)=" & """" & DbcPesquisa.Text & """" & ")) "
End If
If OptOrdem(0).value = True Then
    Ordem = " ORDER BY PRD.CodPrd"
Else
    Ordem = " ORDER BY PRD.Descrição_Produto"
End If
DatDados(0).RecordSource = Sql + Ordem

DatDados(0).Refresh

If DatDados(0).Recordset.RecordCount > 0 Then
    CmbSugestao.Enabled = True
    CmdImprimir.Enabled = True
Else
    CmbSugestao.Enabled = False
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

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)

If Not IsDate(DBGrid1.Columns(6)) Then
    DBGrid1.Columns(6) = Date
'    Exit Sub
End If

If DBGrid1.Columns(3) <> "" And DBGrid1.Columns(4) <> "" And DBGrid1.Columns(5) <> "" Then

    If OPt(1).value = True Then
        DBGrid1.Columns(5) = DBGrid1.Columns(3) + DBGrid1.Columns(3) * ((DBGrid1.Columns(4) / 100))
    Else
        DBGrid1.Columns(4) = (((DBGrid1.Columns(5) / DBGrid1.Columns(3)) - 1) * 100)
    End If
    If DBGrid1.Columns(3) <> DatDados(0).Recordset.Fields("Preço") Then
        DatDados(0).Recordset.Edit
        DatDados(0).Recordset.Fields("Preço") = DBGrid1.Columns(3)
        DatDados(0).Recordset.Fields("Data") = Date
    '    DatDados(0).Recordset.Fields("DtCompra") = CDate(dbgrid1.Columns(6))
        DatDados(0).Recordset.Update
    End If
End If
End Sub


Private Sub DBGrid1_DBLClick()
DBGrid1.Columns(6) = Date
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

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
Dim Sql As String

Sql = "UPDATE PRD SET PRD.DtAlt = #" & Format(Date, "mm/dd/yyyy") & "# WHERE (((PRD.Venda)<> prd.preco))"
'Banco.Execute Sql

If VarPzoProduto = True Then
    Sql = "UPDATE PRD SET Prd.PcoPrazo = prd.venda + (Prd.Venda * (Prd.MargemPrazo / 100)) WHERE (((PRD.Venda)<>[prd].[preco]))"
'    Banco.Execute Sql
End If


End Sub

Private Sub OptOrdem_Click(Index As Integer)
DbcPesquisa_Change
End Sub
