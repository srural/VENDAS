VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormConsistencia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8070
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   12135
   Icon            =   "FormConsistencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12135
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   12
      Top             =   7740
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
      ScaleWidth      =   12165
      TabIndex        =   10
      Top             =   0
      Width           =   12165
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Consistência de Produtos"
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
         TabIndex        =   11
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7020
      TabIndex        =   8
      Top             =   1260
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calcular"
      Height          =   1050
      Left            =   5220
      TabIndex        =   4
      Top             =   720
      Width           =   1410
      Begin VB.OptionButton Opt 
         Caption         =   "&Venda"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Margem"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   585
         Width           =   960
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   1005
      ItemData        =   "FormConsistencia.frx":000C
      Left            =   10485
      List            =   "FormConsistencia.frx":001F
      TabIndex        =   1
      Top             =   1035
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   810
      ItemData        =   "FormConsistencia.frx":0053
      Left            =   8925
      List            =   "FormConsistencia.frx":0063
      TabIndex        =   0
      Top             =   1020
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
      Left            =   7515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormConsistencia.frx":009F
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormConsistencia.frx":013F
      Height          =   5130
      Left            =   180
      OleObjectBlob   =   "FormConsistencia.frx":0159
      TabIndex        =   9
      Top             =   2205
      Width           =   11715
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Está pesquisa mostra os produtos que estão sem preços ou sem aliquotas de impostos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   180
      TabIndex        =   7
      Top             =   765
      Width           =   4065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota"
      Height          =   195
      Left            =   10485
      TabIndex        =   3
      Top             =   795
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tributo"
      Height          =   195
      Left            =   8925
      TabIndex        =   2
      Top             =   780
      Width           =   495
   End
End
Attribute VB_Name = "FormConsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelGruPrd.rpt")
MDIPrincipal.RptRel.WindowTitle = "Relatório de Produtos por Grupo"
MDIPrincipal.RptRel.SelectionFormula = "{Gru.Descrição_Grupo} = """ & DbcPesquisa.Text & """"
MDIPrincipal.RptRel.WindowState = crptMaximized
Imprimir

End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim SQL As String
If Opt(1).value = True Then
     SQL = "UPDATE PRD SET PRD.Venda = [PRD].[custo]+([PRD].[custo]*([prd].[margem]/100)) WHERE (((PRD.CodPrd)=" & DBGrid1.Columns(0) & "))"
Else
     SQL = "UPDATE PRD SET PRD.Margem = ((([Prd].[Venda] / [prd].[custo] ) - 1 ) * 100) WHERE (((PRD.CodPrd)=" & DBGrid1.Columns(0) & "))"
End If
Banco.Execute SQL
DBGrid1.Refresh

End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 300
Formulário.Top = 200

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
'    Inicio
End If

DatDados(0).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Margem, PRD.Custo, PRD.Venda, PRD.Tributo, PRD.Aliquota FROM PRD WHERE (((PRD.Custo)=0) or ((PRD.Venda)=0) OR ((PRD.Tributo)=0) OR ((PRD.Aliquota)=0)) ORDER BY Prd.Descrição_Produto"
DatDados(0).Refresh

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
FlagAbertura = False
End Sub



