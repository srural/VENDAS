VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormAlteraFamilia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8805
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14730
   Icon            =   "FormAlteraFamilia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14730
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -495
      ScaleHeight     =   345
      ScaleWidth      =   15195
      TabIndex        =   15
      Top             =   8460
      Width           =   15225
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   14745
      TabIndex        =   13
      Top             =   0
      Width           =   14775
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Alteração de Preços por Família"
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
         TabIndex        =   14
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   9810
      TabIndex        =   7
      Top             =   1215
      Width           =   1170
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   1005
      ItemData        =   "FormAlteraFamilia.frx":000C
      Left            =   13365
      List            =   "FormAlteraFamilia.frx":001F
      TabIndex        =   4
      Top             =   900
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   810
      ItemData        =   "FormAlteraFamilia.frx":0053
      Left            =   11925
      List            =   "FormAlteraFamilia.frx":0063
      TabIndex        =   3
      Top             =   885
      Width           =   1335
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
      Left            =   7380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAlteraFamilia.frx":009F
      Top             =   4590
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Familia"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   9585
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *  FROM Fam ORDER BY Fam.descrição_Familia"
      Top             =   720
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormAlteraFamilia.frx":01B8
      DataField       =   "Descrição_familia"
      Height          =   945
      Left            =   135
      TabIndex        =   0
      Top             =   900
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   1667
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ForeColor       =   -2147483630
      ListField       =   "Descrição_Familia"
      BoundColumn     =   "codFam"
      Text            =   ""
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   0
      Left            =   105
      ScaleHeight     =   1320
      ScaleWidth      =   11505
      TabIndex        =   2
      Top             =   660
      Width           =   11505
      Begin VB.Frame Frame1 
         Caption         =   "Calcular"
         Height          =   960
         Left            =   7560
         TabIndex        =   8
         Top             =   180
         Width           =   1320
         Begin VB.OptionButton Opt 
            Caption         =   "&Venda"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   10
            Top             =   225
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Margem"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   540
            Width           =   960
         End
      End
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Família"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAlteraFamilia.frx":01D2
      Height          =   6210
      Left            =   135
      OleObjectBlob   =   "FormAlteraFamilia.frx":01EC
      TabIndex        =   11
      Top             =   1980
      Width           =   14460
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
      Left            =   13410
      TabIndex        =   6
      Top             =   630
      Width           =   870
   End
   Begin VB.Label Label1 
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
      Left            =   11970
      TabIndex        =   5
      Top             =   645
      Width           =   750
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   1
      Top             =   780
      Width           =   645
   End
End
Attribute VB_Name = "FormAlteraFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub BtSair_Click()
Unload Me
End Sub

'Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim Sql As String
'If Opt(1).Value = True Then
'     Sql = "UPDATE (FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo SET PRD.Venda = [PRD].[custo]+([PRD].[custo]*([prd].[margem]/100)) WHERE (((FAM.CodFam)= " & DbcPesquisa.BoundText & "))"
'Else
'     Sql = "UPDATE (FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo SET PRD.Margem = ((([prd].[Venda] / [prd].[custo] ) - 1 ) * 100) WHERE (((FAM.CodFam)= " & DbcPesquisa.BoundText & "))"
'End If
'Banco.Execute Sql
'DBGrid1.Refresh
'End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
If OPt(1).value = True Then
    DBGrid1.Columns(5) = DBGrid1.Columns(3) + DBGrid1.Columns(3) * ((DBGrid1.Columns(4) / 100))
Else
    DBGrid1.Columns(4) = (((DBGrid1.Columns(5) / DBGrid1.Columns(3)) - 1) * 100)
End If

End Sub

Private Sub CmdImprimir_Click()
'MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelGruFan.rpt")
'MDIPrincipal.RptRel.WindowTitle = "Relatório de Produtos por Familia"
'MDIPrincipal.RptRel.SelectionFormula = "{Fam.CodFam} = " & DbcPesquisa.BoundText
'MDIPrincipal.RptRel.WindowState = crptMaximized
'Imprimir

Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim Sql As String

    'CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia FROM ((FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia "
    'Ordenado = " ORDER BY FAM.Descrição_Familia, GRU.Descrição_Grupo, PRD.Descrição_Produto "
    'If IsNumeric(DbcPesquisa.BoundText) Then CondicaoWhere = " Having CodFam = " & DbcPesquisa.BoundText
    
        'If ChkEstoqueZero.Value = 0 Then
            Sql = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, FAM.CodFam, FAM.Descrição_Familia, GRU.CodGru, GRU.Descrição_Grupo FROM (FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, FAM.CodFam, FAM.Descrição_Familia, GRU.CodGru, GRU.Descrição_Grupo Having (((FAM.CodFam)= " & DbcPesquisa.BoundText & ")) ORDER BY PRD.Descrição_Produto"

        'Else
            'Sql = CamposSql & Ordenado
        'End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
'    Set FrmRelPrdFam_Altera.Relatorio.Recordset = DatDados(0).Recordset.Clone
    
    FrmRelPrdFam_Altera.rpCabTitulo.Caption = "Relatório de Produtos por Familia"
    FrmRelPrdFam_Altera.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPrdFam_Altera.Config


End Sub

Private Sub DbcPesquisa_Change()
If IsNumeric(DbcPesquisa.BoundText) Then
    DatDados(0).RecordSource = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, prd.margem, PRD.Venda, PRD.Tributo, PRD.Aliquota, FAM.CodFam FROM (FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE (((FAM.CodFam)= " & DbcPesquisa.BoundText & ")) ORDER BY PRD.Descrição_Produto"

End If
DatDados(0).Refresh

End Sub


Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 200
Formulário.Top = 100

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
DbcPesquisa.SetFocus
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
'Set Registros = Banco.OpenRecordset("SELECT *  FROM FOR  ORDER BY FOR.NOME_FORNECEDOR")


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
End Sub



Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória

Dim Sql As String

Sql = "UPDATE PRD SET PRD.DtAlt = #" & Format(Date, "mm/dd/yyyy") & "# WHERE (((PRD.Venda)<>[prd].[preco]))"

'Banco.Execute Sql


End Sub


