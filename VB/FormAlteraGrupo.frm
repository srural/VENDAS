VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormAlteraGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9060
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14730
   Icon            =   "FormAlteraGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   14730
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -405
      ScaleHeight     =   345
      ScaleWidth      =   15150
      TabIndex        =   23
      Top             =   8730
      Width           =   15180
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   14745
      TabIndex        =   21
      Top             =   0
      Width           =   14775
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Alteração de Preços por Grupo"
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
         TabIndex        =   22
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   810
      ItemData        =   "FormAlteraGrupo.frx":000C
      Left            =   12195
      List            =   "FormAlteraGrupo.frx":001C
      TabIndex        =   17
      Top             =   1095
      Width           =   1245
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   1005
      ItemData        =   "FormAlteraGrupo.frx":0058
      Left            =   13500
      List            =   "FormAlteraGrupo.frx":006B
      TabIndex        =   16
      Top             =   885
      Width           =   990
   End
   Begin VB.CommandButton CmbSugestao 
      Caption         =   "&Sug.Compra"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10785
      TabIndex        =   12
      Top             =   900
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Altera Preços"
      Height          =   1185
      Left            =   9135
      TabIndex        =   8
      Top             =   720
      Width           =   1485
      Begin MSMask.MaskEdBox MebPerc 
         Height          =   300
         Left            =   90
         TabIndex        =   11
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton BtOk 
         Caption         =   "&OK"
         Height          =   270
         Left            =   255
         TabIndex        =   9
         Top             =   825
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Percentual"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Relatório"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10785
      TabIndex        =   4
      Top             =   1455
      Width           =   1170
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   6615
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAlteraGrupo.frx":009B
      Top             =   3510
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Grupo"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4815
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Gru.CodGru, Gru.Descrição_Grupo  FROM gru ORDER BY gru.descrição_grupo"
      Top             =   3915
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormAlteraGrupo.frx":01BF
      DataField       =   "Descrição_grupo"
      Height          =   945
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   3945
      _ExtentX        =   6959
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
      Left            =   135
      ScaleHeight     =   1275
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   675
      Width           =   6675
      Begin VB.Frame Frame3 
         Caption         =   "Ordem"
         Height          =   945
         Left            =   4170
         TabIndex        =   13
         Top             =   180
         Width           =   1230
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Código"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   15
            Top             =   255
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "&Descricão"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   615
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Calcular"
         Height          =   960
         Left            =   5490
         TabIndex        =   5
         Top             =   180
         Width           =   1110
         Begin VB.OptionButton Opt 
            Caption         =   "&Margem"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   7
            Top             =   570
            Width           =   960
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Venda"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   6
            Top             =   225
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Grupo"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAlteraGrupo.frx":01D9
      Height          =   6480
      Left            =   135
      OleObjectBlob   =   "FormAlteraGrupo.frx":01F3
      TabIndex        =   1
      Top             =   1980
      Width           =   14400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tributo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12195
      TabIndex        =   19
      Top             =   855
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13500
      TabIndex        =   18
      Top             =   645
      Width           =   870
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   645
   End
End
Attribute VB_Name = "FormAlteraGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub BtOK_Click()
If MebPerc <> "" Then
    If IsNumeric(MebPerc) And DatDados(0).Recordset.RecordCount > 0 Then
        If MsgBox("Atenção esta operação altera todos os preços deste grupo. Confirma ?", vbYesNo, App.Title) = vbYes Then
            With DatDados(0).Recordset
                .MoveFirst
                While Not .Eof
                    .Edit
                    If OPt(1).value = True Then
                        .Fields("Custo") = .Fields("Custo") + (.Fields("Custo") * (MebPerc / 100))
                        .Fields("Venda") = .Fields("Venda") + (.Fields("Venda") * (MebPerc / 100))
                    Else
                        .Fields("Venda") = .Fields("Custo") + (.Fields("Custo") * ((.Fields("Margem") + MebPerc) / 100))
                        .Fields("Margem") = .Fields("Margem") + MebPerc
                    End If
                    
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
Dim SQL As String

SQL = "Select * From Prd Where PRD.Estoque < PRD.Mínimo and PRD.Mínimo <> 0 and Prd.Grupo = " & DbcPesquisa.BoundText & " Order By Descrição_Produto "

Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelSugestao.Relatorio.Recordset = varPesquisa(1)
FrmRelSugestao.rpCabTitulo.Caption = "Relatório de Sugestão de Compras por Grupo - " & DbcPesquisa.Text
FrmRelSugestao.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelSugestao.Config

End Sub

Private Sub CmdImprimir_Click()
'MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelPrdGru.rpt")
'MDIPrincipal.RptRel.WindowTitle = "Relatório de Produtos por Grupo"
'MDIPrincipal.RptRel.SelectionFormula = "{Gru.CodGru} = " & DbcPesquisa.BoundText
'MDIPrincipal.RptRel.WindowState = crptMaximized
'Imprimir
Dim Componentes As Variant
Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim SQL As String

    'CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, PRD.DtAlt, GRU.CodGru, GRU.Descrição_Grupo FROM GRU INNER JOIN (ENT INNER JOIN (PRD INNER JOIN COT ON PRD.CodPrd = COT.CodPrd) ON ENT.CodEntidade = COT.CodFor) ON GRU.CodGru = PRD.Grupo GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, PRD.DtAlt, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo "
    'Ordenado = " ORDER BY GRU.Descrição_Grupo, PRD.Descrição_Produto "
    'If IsNumeric(DbcPesquisa.BoundText) Then CondicaoWhere = "Having CodGru = " & DbcPesquisa.BoundText
    
    
        'If ChkEstoqueZero.Value = 0 Then
            SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, PRD.DtAlt, PRD.Grupo, GRU.CodGru, GRU.Descrição_Grupo FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, PRD.DtAlt, PRD.Grupo, GRU.CodGru, GRU.Descrição_Grupo Having (((PRD.Grupo)= " & DbcPesquisa.BoundText & ")) ORDER BY Descrição_Produto"
        'Else
            'Sql = CamposSql & Ordenado
        'End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
'    Set varPesquisa(1) = Sql
    Set FrmRelPrdGru_Altera.Relatorio.Recordset = varPesquisa(1)
            
    FrmRelPrdGru_Altera.rpCabTitulo.Caption = "Relatório de Produtos por Grupo"
    FrmRelPrdGru_Altera.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPrdGru_Altera.Config

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DbcPesquisa_Change()
Dim SQL As String
Dim Ordem As String

If IsNumeric(DbcPesquisa.BoundText) Then
     SQL = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem,Prd.PcoPrazo, PRD.Venda,Prd.Tributo, Prd.Aliquota, Prd.DtAlt, Prd.SitTrib, Prd.Classe, Prd.CfopPrd, Prd.Iva FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE (((PRD.Grupo)= " & DbcPesquisa.BoundText & ")) "
End If
If OptOrdem(0).value = True Then
    Ordem = " ORDER BY PRD.CodPrd"
Else
    Ordem = " ORDER BY PRD.Descrição_Produto"
End If

DatDados(0).RecordSource = SQL + Ordem
If SQL <> "" Then
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount > 0 Then
        CmbSugestao.Enabled = True
        CmdImprimir.Enabled = True
    Else
        CmbSugestao.Enabled = False
        CmdImprimir.Enabled = False
    End If

End If

End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)

If DBGrid1.Columns(3) <> "" And DBGrid1.Columns(4) <> "" And DBGrid1.Columns(5) <> "" Then
    If OPt(1).value = True Then
        DBGrid1.Columns(5) = DBGrid1.Columns(3) + DBGrid1.Columns(3) * ((DBGrid1.Columns(4) / 100))
    Else
        DBGrid1.Columns(4) = (((DBGrid1.Columns(5) / DBGrid1.Columns(3)) - 1) * 100)
    End If
'    If VarPzoProduto = True Then
'          Datdados(0).Fields("PcoPrazo") = Datdados(0).Fields("Venda") + (Datdados(0).Fields("Venda") * (Datdados(0).Fields("MargemPrazo") / 100))
'    End If
                
End If
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
Dim SQL As String

SQL = "UPDATE PRD SET PRD.DtAlt = #" & Format(Date, "mm/dd/yyyy") & "# WHERE (((PRD.Venda)<>[prd].[preco]))"
'Banco.Execute Sql

If VarPzoProduto = True Then
    SQL = "UPDATE PRD SET Prd.PcoPrazo = prd.venda + (Prd.Venda * (Prd.MargemPrazo / 100)) WHERE (((PRD.Venda)<>[prd].[preco]))"
'    Banco.Execute Sql
End If


End Sub


Private Sub OptOrdem_Click(Index As Integer)
DbcPesquisa_Change
End Sub
