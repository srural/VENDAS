VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   11070
   Icon            =   "FrmGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   11070
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11280
      TabIndex        =   14
      Top             =   3870
      Width           =   11310
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11085
      TabIndex        =   12
      Top             =   0
      Width           =   11085
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Grupo de Produto"
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
         TabIndex        =   13
         Top             =   45
         Width           =   6540
      End
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   765
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Gru Order by Descrição_Grupo"
      Top             =   3285
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox TxtPesquisa 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   5055
   End
   Begin VB.Data DatDados 
      Caption         =   "Cidades"
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   3420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Gru Order by Descrição_Grupo"
      Top             =   3285
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data DatDados 
      Caption         =   "Grupo"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5580
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Fam Order by Descrição_Familia"
      Top             =   3285
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   5310
      TabIndex        =   1
      Top             =   810
      Width           =   5460
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1335
         Index           =   0
         Left            =   135
         ScaleHeight     =   1335
         ScaleWidth      =   5055
         TabIndex        =   6
         Top             =   675
         Width           =   5055
         Begin VB.CheckBox ChkDados 
            Caption         =   "Baixa de Estoque Automática"
            DataField       =   "BaixarEstoque"
            Height          =   375
            Index           =   0
            Left            =   765
            TabIndex        =   5
            Tag             =   "S"
            Top             =   945
            Width           =   2535
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Descrição_Grupo"
            Height          =   315
            Index           =   1
            Left            =   765
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "S"
            Top             =   105
            Width           =   4095
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmGrupo.frx":000C
            DataField       =   "Familia"
            Height          =   315
            Index           =   0
            Left            =   765
            TabIndex        =   4
            Tag             =   "SI"
            Top             =   525
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "Descrição_Familia"
            BoundColumn     =   "CodFam"
            Text            =   ""
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Família:"
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   8
            Top             =   540
            Width           =   555
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   105
            Width           =   465
         End
      End
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         DataField       =   "CodGru"
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Tag             =   "S"
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   390
         Width           =   540
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmGrupo.frx":0026
      Height          =   1815
      Left            =   135
      OleObjectBlob   =   "FrmGrupo.frx":0040
      TabIndex        =   10
      Top             =   1305
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Grupo"
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FrmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset   'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Dim FlagAbertura As Boolean
Dim Cliente As Integer

Public CTextBox As Boolean      'Caixa de Texto
Public CMaskEditBox As Boolean  'Mascara
Public CComboBox As Boolean     'Combo
Public CListBox As Boolean      'Lista
Public CdbCombo As Boolean      'DBCombo
Public CDbList As Boolean       'DBLista
Public CCheckBox As Boolean     'Caixa de Checagem
Public CDbgrid As Boolean       'Grade
Public ComandBotom As Boolean   'Botão

Public SelectFromSQL  As String  ' Representa o Conteudo das Cláusulas Select e From do Sql Utilizado p/ acessar os registros deste formulário
Public WhereSQL As String        ' Representa o Conteudo da Cláusula Where do Sql Utilizado p/ acessar os registros deste formulário
Public FiltroSQL As String       ' Critério imposto pelo usuário
Public OrderBySQL As String      ' Representa a cláusula Order By do Sql do formulário Ativo
Public PesquisaSql As String     ' Representa a pesquisa F2
Public SelectFromFiltro As String  ' Representa a relação de campos que sera exibida na visualização no formulário de filtragem

Public LocalizacaoBasica As String ' Armazena o nome do campo que sera utilizado na ferramenta de localizaco basica

Private Sub DefineObjetos()
'Define quais objeto foram utilizados neste formulário

CTextBox = True         ' TxtDados
CMaskEditBox = False    ' MebDados
CComboBox = False       ' CmbDados
CListBox = False
CdbCombo = True        ' DbcDados
CDbList = False         '
CCheckBox = True       ' ChkDados
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo
' Titular = Registros.Fields("CodTit")
End Sub



Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodGru = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
'TabCad.Tab = 1 'direcina para o TAB seguinte
If Registros.RecordCount > 0 Then
    MostraDados
End If
End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)
Formulário.Left = 200
Formulário.Top = 300

'Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
'Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)

' testa se o banco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
'TabCad.Tab = 0
AbilitaBotoes (True)
TxtPesquisa.SetFocus

'PosMovimentação
' DbcPesquisa.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
    
Set Registros = Banco.OpenRecordset("SELECT GRU.CodGru, GRU.Descrição_Grupo, GRU.Familia, Fam.CodFam, FAM.Descrição_Familia, Gru.BaixarEstoque  FROM FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia ORDER BY GRU.Descrição_Grupo")
Registros.Requery

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select * From Gru Order by Descrição_Grupo"
DatDados(0).Refresh

LocalizacaoBasica = "Descrição_Grupo"
DefineObjetos
SelectFromSQL = "Select * From Gru"
SelectFromFiltro = "Select CodGru, Descrição_Grupo from Gru"
PesquisaSql = "SELECT CodGru , Descrição_Grupo FROM Gru ORDER BY Descrição_Grupo"
FlagAbertura = False

OrderBySQL = "[Gru].[Descrição_Grupo]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
   Confirmacoes (15)
End If

Registros.Close

AbilitaBotoes (True)
AtivarMenu (False)

End Sub

Public Sub PosInclusao()
Dim REG As Recordset
' sera executada logo depois que o usuario iniciar o processo de inclusao addnew

Set REG = Banco.OpenRecordset("Select *  From Gru Order By CodGru")
If REG.RecordCount > 0 Then
    REG.Requery
    REG.MoveFirst
    REG.MoveLast
    TxtDados(0).Text = CDec(REG!codgru) + 1
    ChkDados(0).value = 1
    DbcDados(0).Text = "DIVERSOS"
Else
    TxtDados(0).Text = 1
End If


TxtDados(0).SetFocus
'TabCad.TabEnabled(0) = False


End Sub

Public Function PermiteInclusao() As Boolean
' Será executado antes do processo de inclusao (addnew)
' Aqui deve ser feitos os processo p/ que o usuário pode ou não incluir um novo registro
' MsgBox "Neste formulário não pode ser feitas Inclusão de novos registros", vbInformation, App.Title

PermiteInclusao = True

End Function

Public Sub PosAlteracao()
' sera executada logo depois que o usuario iniciar o processo de alteracao
'TxtDados(0).BackColor = QBColor(8) 'Muda a Cor de fundo do código
TxtDados(0).SetFocus
'TabCad.TabEnabled(0) = False
'Atualizar
End Sub

Public Function PermiteAlteracao() As Boolean
' Será executado antes do processo de Alteracao

PermiteAlteracao = True

End Function

Public Sub PosExclusao()
' sera executada logo depois que o usuario iniciar o processo de exclusao

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
' Será executado antes do processo de exclusao
Dim Mensagem As String
Dim VarPrd As Recordset

Set VarPrd = Banco.OpenRecordset("Select CodPrd, Grupo From Prd Where Grupo= " & TxtDados(0))
VarPrd.Requery

If VarPrd.RecordCount > 0 Then
    MsgBox "Grupo não pode ser apagado pois já tem produto vinculado !", vbInformation, App.Title
    PermiteExclusao = False
    Exit Function
End If


Mensagem = "Confirma Exclusão do Grupo " & vbCr & Registros.Fields("Descrição_Grupo")

If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
' nesta posicao do codigo podemos incluir os procedimentos que deverao
' apagar todos os registro que dependem do registro que está sendo excluido, no caso de relacionamento entre tabelas
' ex : poderiamos excluir os itens de um pedido antes de excluir o cabeçalho
    PermiteExclusao = True
Else
    PermiteExclusao = False
End If

End Function

Public Sub PreConfirmacao() 'Será executada antes das rotinas de salvar ou cancelar

End Sub

Public Sub PosConfirmacao() 'Será executada deposi dos processos de salvar ou cancelar

'TxtDados(0).BackColor = QBColor(7) 'Muda a Cor de fundo do código
'Txtdados(0).Locked = False
Atualizar
End Sub

Public Function PermiteCancelamento() As Boolean
' Será executada antes do processo de cancelamento ser executado (cancelupdate)
' podemos aqui impedir o usuario de cancelar uma inclusao ou uma alteracao

PermiteCancelamento = True

End Function

Public Function PermiteSalvar() As Boolean
' Será axecutada antes de atribuir os campos do formulário no bco de dados
' Podemos fazer qualquer tipo de consistencia que não possa ser feita pelo bco de  dados
' Ex: O estoque máximo sempre deverá ser maior que o estoque minimo

PermiteSalvar = True

End Function
Public Function PermiteContinuarMostrando() As Boolean

End Function

Public Function PermiteContinuarSalvar() As Boolean
' Será executada logo depois de atribuir os campos ao bco de dados
' Possui as mesma função do permitir salvar, Inclusive podendo atribuir valores a campos do bco de dados que não aparecem na tela
' Ex: Aqui vc pode atribuir um valor ao bco que não está no formulario
PermiteContinuarSalvar = True

End Function

Public Function PosSalvamento() As Boolean
'Será executado assim que todos os dados forem gravados no bco

PosSalvamento = True

'TabCad.TabEnabled(0) = True

End Function

Public Sub PosCancelamento()

'sera executado logo após do processo de cancelamento da aplicaçaõ

'TabCad.TabEnabled(0) = True

End Sub

Public Sub Atualizar()
'Registros.Requery
DatDados(0).Refresh
DatDados(1).Refresh
End Sub

'Private Sub TabCad_Click(previoustab As Integer)
'Select Case TabCad.Tab
'    Case 0
'        Atualizar
'        AbilitaBotoes (False)
'    Case 1
'        AbilitaBotoes (True)
'    Case 2
'        AbilitaBotoes (True)
'End Select

'End Sub

Private Sub Txtdados_Change(Index As Integer)
Dim SomaGrid As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Pago = 0
Apagar = 0
'If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
'    If IsNumeric(TxtDados(0)) = False Then
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = 0 ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=0)) ORDER BY Bol.Data"

'    Else
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = " & TxtDados(0) & " ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=" & TxtDados(0) & " )) ORDER BY Bol.Data"
'    End If
'    DatDados(2).Refresh
'    DatDados(3).Refresh
'End If

End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        Criterio = DatDados(0).Recordset.Fields(0).Name & " =  " & TxtPesquisa
    Else
        Criterio = DatDados(0).Recordset.Fields(1).Name & " LIKE " & """" & TxtPesquisa & "*"""
    End If
    
    DatDados(0).Recordset.FindFirst Criterio
    If DatDados(0).Recordset.NoMatch Then
        TxtPesquisa.SetFocus
    Else
        Registros.FindFirst Criterio
        MostraDados
    End If
End If


End Sub

Private Sub TxtPesquisa_LostFocus()
    If TxtPesquisa.Text <> "" Then
        DBGrid1_DBLClick
    End If
End Sub
Public Sub PosMovimentacao()

End Sub

