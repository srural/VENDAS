VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormCotacao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7920
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   11460
   Icon            =   "FormCotacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11460
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   8
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
      ScaleWidth      =   11490
      TabIndex        =   6
      Top             =   0
      Width           =   11490
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotação de Preços por Fornecedor"
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
         TabIndex        =   7
         Top             =   45
         Width           =   6675
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "Caminho & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   8100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4620
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   "Caminho & ""Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4410
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
      Top             =   3870
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   180
      ScaleHeight     =   735
      ScaleWidth      =   10725
      TabIndex        =   4
      Top             =   660
      Width           =   10725
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   8685
         TabIndex        =   1
         Top             =   225
         Width           =   1425
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormCotacao.frx":000C
         DataField       =   "Nome"
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   225
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Fornecedor"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormCotacao.frx":0026
      Height          =   5790
      Left            =   180
      OleObjectBlob   =   "FormCotacao.frx":0040
      TabIndex        =   2
      Top             =   1425
      Width           =   10995
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   405
      TabIndex        =   3
      Top             =   780
      Width           =   645
   End
End
Attribute VB_Name = "FormCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Dim FlagAbertura As Boolean

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim SQL As String

    CamposSql = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo FROM GRU INNER JOIN (ENT INNER JOIN (PRD INNER JOIN COT ON PRD.CodPrd = COT.CodPrd) ON ENT.CodEntidade = COT.CodFor) ON GRU.CodGru = PRD.Grupo GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo "
    Ordenado = "ORDER BY Ent.Nome, PRD.Descrição_Produto"
    CondicaoWhere = "Having Ent.Nome Like '" & DbcPesquisa.Text & "'"
    
    SQL = CamposSql & CondicaoWhere & Ordenado
    
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPrdFor.Relatorio.Recordset = varPesquisa(1)
    
    FrmRelPrdFor.rpCabTitulo.Caption = "Relatório de Produtos por Fornecedor"
    FrmRelPrdFor.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPrdFor.Config

End Sub

Private Sub DbcPesquisa_Change()
If DbcPesquisa.Text = "" Then
    DatDados(0).RecordSource = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, COT.Preço, COT.Data, COT.Condições FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd  WHERE (((ENT.Nome)=" & """" & Space(0) & """" & ")) ORDER BY PRD.Descrição_Produto"
Else
    DatDados(0).RecordSource = "SELECT Prd.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, COT.Preço, COT.Data, COT.Condições FROM PRD INNER JOIN (ENT INNER JOIN COT ON ENT.CodEntidade = COT.CodFor) ON PRD.CodPrd = COT.CodPrd  WHERE (((ENT.Nome)=" & """" & DbcPesquisa.Text & """" & ")) ORDER BY PRD.Descrição_Produto"
End If
DatDados(0).Refresh
End Sub

Private Sub DBGrid1_Change()
DBGrid1.Columns(3) = Date
End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
'    Inicio
End If

DbcPesquisa.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

FlagAbertura = False

End Sub



Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
AtivarMenu (False)

End Sub


Public Sub Atualizar()
' Atualiza a variavel registro incluindo os novos registros cadastrados e excluindo os que foram deletados
Registros.Requery
DatDados(0).Refresh
DatDados(1).Refresh

End Sub


Private Sub DbcPesquisa_LostFocus()
Dim Criterio As String
If IsNumeric(DbcPesquisa.Text) Then
   Criterio = "CodFor = " & DbcPesquisa.Text
   DatDados(1).Recordset.FindFirst Criterio
   If DatDados(1).Recordset.NoMatch Then
       MsgBox "Informação não Localizada", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
        DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome_Fornecedor")
        DBGrid1.SetFocus
   End If
End If
End Sub

