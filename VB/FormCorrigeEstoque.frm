VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormCorrigeEstoque 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7890
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   13500
   Icon            =   "FormCorrigeEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   13500
   Begin VB.CommandButton AtuEstoque 
      Caption         =   "Atualiza Estoque"
      Height          =   375
      Left            =   2025
      TabIndex        =   16
      Top             =   7065
      Width           =   2760
   End
   Begin VB.CommandButton BtLaser 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   225
      TabIndex        =   15
      Top             =   7065
      Width           =   1410
   End
   Begin VB.CommandButton CmdExclui 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11565
      TabIndex        =   14
      Top             =   7065
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   14340
      TabIndex        =   10
      Top             =   7560
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
      ScaleWidth      =   13515
      TabIndex        =   8
      Top             =   0
      Width           =   13515
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Correção de Estoque"
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
         TabIndex        =   9
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.Data Datdados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas novo Estoque\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3915
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormCorrigeEstoque.frx":000C
      Top             =   4905
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Datdados 
      Caption         =   "Grupo"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   3960
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Gru.CodGru, Gru.Descrição_Grupo  FROM gru ORDER BY gru.descrição_grupo"
      Top             =   3915
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormCorrigeEstoque.frx":00EE
      DataField       =   "Descrição_grupo"
      Height          =   945
      Left            =   225
      TabIndex        =   0
      Top             =   900
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   1667
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   "Descrição_grupo"
      BoundColumn     =   "codGru"
      Text            =   ""
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Index           =   0
      Left            =   180
      ScaleHeight     =   1275
      ScaleWidth      =   12840
      TabIndex        =   3
      Top             =   675
      Width           =   12840
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   4860
         TabIndex        =   17
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "Adiciona"
         Height          =   375
         Left            =   9810
         TabIndex        =   13
         Top             =   810
         Width           =   915
      End
      Begin VB.CheckBox ChkPrd 
         Caption         =   "Produto"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4860
         TabIndex        =   12
         Top             =   540
         Width           =   1590
      End
      Begin VB.Data Datdados 
         Caption         =   "Produtos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   2
         Left            =   8055
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   90
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         Height          =   945
         Left            =   10890
         TabIndex        =   5
         Top             =   225
         Width           =   1365
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Código"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   255
            Width           =   915
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Descricão"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   6
            Top             =   615
            Value           =   -1  'True
            Width           =   1110
         End
      End
      Begin MSDBCtls.DBCombo DbcProdutos 
         Bindings        =   "FormCorrigeEstoque.frx":0108
         Height          =   360
         Left            =   4860
         TabIndex        =   11
         Top             =   810
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   "descrição_Produto"
         BoundColumn     =   "CodPrd"
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
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Grupo"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormCorrigeEstoque.frx":0122
      Height          =   4950
      Left            =   180
      OleObjectBlob   =   "FormCorrigeEstoque.frx":013C
      TabIndex        =   1
      Top             =   2070
      Width           =   12825
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   270
      TabIndex        =   2
      Top             =   780
      Width           =   645
   End
End
Attribute VB_Name = "FormCorrigeEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public VarLoja As Recordset

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub AtuEstoque_Click()

If DatDados(0).Recordset.RecordCount > 0 Then
    DatDados(0).Recordset.MoveFirst
    While Not DatDados(0).Recordset.Eof
        Set VarLoja = Banco.OpenRecordset("SELECT LOJ.Produto, Sum(LOJ.Estoque) AS SomaDeEstoque FROM LOJ GROUP BY LOJ.Produto HAVING LOJ.Produto = " & DatDados(0).Recordset("CodPrd"))
        VarLoja.Requery
        DatDados(0).Recordset.Edit
        If VarLoja.RecordCount = 1 Then
            DatDados(0).Recordset("Prd.Estoque") = DatDados(0).Recordset("Est_Ideal")
            DatDados(0).Recordset("Loj.Estoque") = DatDados(0).Recordset("Est_Ideal")
        Else
            DatDados(0).Recordset("Prd.Estoque") = VarLoja!somadeestoque
            DatDados(0).Recordset("Loj.Estoque") = DatDados(0).Recordset("Est)Ideal")
        End If
        DatDados(0).Recordset.Update
        DatDados(0).Recordset.MoveNext
    Wend
End If

End Sub

Private Sub BtLaser_Click()
Dim SQL As String

SQL = "SELECT Loj.CodBarLoja, Prd.Local, Prd.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, PRD.Venda, PRD.Margem, Prd.Est_Ideal, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[CodBarLOja] AS SubTotal, [PRd].[Custo]*[Prd].[Est_Ideal] AS SubEstAnt, PRD.Fabricante FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto GROUP BY Prd.Local, Prd.Est_Ideal, Prd.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, Prd.Venda, PRD.Venda, PRD.Margem, Loj.CodBarLoja, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque], PRD.Fabricante  Having Prd.Ativo = True And CodGru =   " & DbcPesquisa.BoundText & "  ORDER BY Descrição_Produto"

Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelCorrigeEst.Relatorio.Recordset = varPesquisa(1)
FrmRelCorrigeEst.rpCabSubTitulo.Caption = NomeEmpresa
'FrmRelEst.Tag = "*"
FrmRelCorrigeEst.Config



'Dim rptGrid As FrmRelConfEst

'If Lista.ListItems.Count > 0 Then
'    Set rptGrid = New FrmRelConfEst
'    Set rptGrid.Grid = Lista
'    rptGrid.Relatorio.Escala = scmCentimetros
'    rptGrid.Relatorio.MargemEsquerda = 1
'    rptGrid.Relatorio.MargemDireita = 1
'    rptGrid.TituloRelatorio.Caption = "Conferência de Estoque"
'    rptGrid.Cabec1.Caption = "Periodo:" & DtData1 & "  até  " & DtData2
'    rptGrid.Executar
'End If
End Sub

Private Sub ChkPrd_Click()

If ChkPrd.value = 1 Then

    DatDados(0).RecordSource = " SELECT Loj.CodBarLoja, PRD.Validade, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Mínimo, PRD.Estoque, LOJ.DescricaoLoja, LOJ.Estoque, Len([bar].[BarBarra]) AS Expr1 FROM (PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto) INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((PRD.Validade)=99) AND ((Len([bar].[BarBarra]))=13)) ORDER BY PRD.Descrição_Produto"
    DatDados(0).Refresh

End If

End Sub

Private Sub ChkTodos_Click()
Dim SQL As String
Dim Ordem As String

If ChkTodos.value = 1 Then
    
    SQL = "SELECT Loj.CodBarLoja, Prd.Est_Ideal, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Mínimo, PRD.Estoque, LOJ.DescricaoLoja, LOJ.Estoque FROM PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto "

    If OptOrdem(0).value = True Then
        Ordem = " ORDER BY PRD.CodPrd"
    Else
        Ordem = " ORDER BY PRD.Descrição_Produto"
    End If
    DatDados(0).RecordSource = SQL + Ordem
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount > 0 Then
        DatDados(0).Recordset.MoveFirst
    End If
        
End If
End Sub

Private Sub CmdExclui_Click()
If DatDados(0).Recordset.RecordCount > 0 Then
    DatDados(0).Recordset.Edit
    DatDados(0).Recordset("Validade") = 0
    DatDados(0).Recordset.Update
    DatDados(0).Refresh
End If
End Sub

Private Sub CmdOk_Click()
If IsNumeric(DbcProdutos.BoundText) Then
    DatDados(2).Recordset.Edit
    DatDados(2).Recordset("Validade") = 99
    DatDados(2).Recordset.Update
    DatDados(0).Refresh
End If
End Sub

Private Sub DbcPesquisa_Change()
Dim SQL As String
Dim Ordem As String
Dim SqlEst As String

If IsNumeric(DbcPesquisa.BoundText) Then
     SQL = "SELECT Loj.CodBarLoja, Prd.Est_Ideal, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Mínimo, PRD.Estoque, LOJ.DescricaoLoja, LOJ.Estoque FROM PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto  WHERE (((PRD.Grupo)= " & DbcPesquisa.BoundText & ")) "
End If
If OptOrdem(0).value = True Then
    Ordem = " ORDER BY PRD.CodPrd"
Else
    Ordem = " ORDER BY PRD.Descrição_Produto"
End If

DatDados(0).RecordSource = SQL + Ordem
If SQL <> "" Then
'    SqlEst = "UPDATE PRD SET PRD.Est_Ideal = [prd].[estoque] WHERE PRD.Grupo = " & DbcPesquisa.BoundText
'    Banco.Execute SqlEst
    
'    SqlEst = "UPDATE PRD SET PRD.Est_Ideal = [prd].[estoque] WHERE PRD.Grupo = " & DbcPesquisa.BoundText
     SqlEst = "UPDATE PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto SET LOJ.CodBarLoja = [prd].[est_ideal]-[prd].[estoque] WHERE PRD.Grupo= " & DbcPesquisa.BoundText
    Banco.Execute SqlEst
  
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount > 0 Then
        DatDados(0).Recordset.MoveFirst
    End If
'    Open App.Path & "\ChkEst.txt" For Output As #1
'    While Not DatDados(0).Recordset.Eof
'        Print #1, psCompString(2, Numero2(DatDados(0).Recordset.Fields("loj.Estoque")), 10) & psCompString(1, DatDados(0).Recordset.Fields("CodPrd"), 8) & psCompString(3, Trim(DatDados(0).Recordset.Fields("DescricaoLoja")), 20)
'        DatDados(0).Recordset.MoveNext
'    Wend
'    Close #1
    
'    DatDados(0).Recordset.MoveFirst
'    DatDados(0).Refresh
End If

End Sub


Private Sub DbcProdutos_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If DbcProdutos <> "" And Not IsNumeric(DbcProdutos.Text) Then
    Criterio = "Prd.descrição_produto Like " & """" & DbcProdutos & "*"""
    DatDados(2).Recordset.FindFirst Criterio
    If DatDados(2).Recordset.NoMatch Then
        Beep
'        DbcProdutos.SetFocus
    Else
        If Len(DbcProdutos) >= 1 Then
             DbcProdutos.Text = DatDados(2).Recordset("Descrição_Produto")
        End If
    End If
End If

End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)

If DBGrid1.Columns(9) <> "" Then
'    Set VarLoja = Banco.OpenRecordset("SELECT LOJ.Produto, Sum(LOJ.Estoque) AS SomaDeEstoque FROM LOJ GROUP BY LOJ.Produto HAVING LOJ.Produto = " & DBGrid1.Columns(0))

   DBGrid1.Columns(10) = Format(DBGrid1.Columns(9) - DBGrid1.Columns(8), "##,##0.00")
    
End If


End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 300
Formulário.Top = 200

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
'DbcProdutos.SetFocus
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

'Data.DatabaseName = Caminho & "\Estoque.mdb"


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatDados(2).RecordSource = "SELECT Prd.Validade, PRD.Essenciais, PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PRD.MargemPrazo, PRD.PcoPrazo, BAR.BarBarra FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((Len([bar].[barbarra]))=13)) ORDER BY PRD.Descrição_Produto"

DatDados(0).RecordSource = "SELECT Loj.CodBarLoja, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Mínimo, PRD.Estoque, LOJ.DescricaoLoja, LOJ.Estoque FROM PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto where Grupo = 0 ORDER BY PRD.Descrição_Produto"


End Sub



Private Sub OptOrdem_Click(Index As Integer)
DbcProdutos_Change
End Sub

