VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmFamilias 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   11115
   Icon            =   "FrmFamilias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   11115
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -405
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   12
      Top             =   3420
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11130
      TabIndex        =   10
      Top             =   0
      Width           =   11130
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Familia de Produto"
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
         Width           =   6045
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   5040
      TabIndex        =   5
      Top             =   1350
      Width           =   5865
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "CodFam"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "N"
         Top             =   450
         Width           =   975
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   0
         Left            =   315
         ScaleHeight     =   750
         ScaleWidth      =   5325
         TabIndex        =   6
         Top             =   765
         Width           =   5325
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            DataField       =   "Descrição_Familia"
            Height          =   315
            Index           =   1
            Left            =   765
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "S"
            Top             =   90
            Width           =   4365
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   7
            Top             =   135
            Width           =   465
         End
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   9
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Automático"
         Height          =   195
         Left            =   2115
         TabIndex        =   8
         Top             =   495
         Width           =   795
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Cidades"
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Fam Order by Descrição_Familia"
      Top             =   765
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtPesquisa 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   1035
      Width           =   4695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmFamilias.frx":000C
      Height          =   1785
      Left            =   270
      OleObjectBlob   =   "FrmFamilias.frx":0026
      TabIndex        =   1
      Top             =   1395
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Familia"
      Height          =   240
      Left            =   270
      TabIndex        =   4
      Top             =   810
      Width           =   1215
   End
End
Attribute VB_Name = "FrmFamilias"
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
CdbCombo = False        ' DbcDados 'xxxx
CDbList = False         '
CCheckBox = False       ' ChkDados
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo
' Titular = Registros.Fields("CodTit")
End Sub



Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodFam = " & DBGrid1.Columns(0)
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
    
Set Registros = Banco.OpenRecordset("SELECT * FROM Fam ORDER BY FAM.Descrição_Familia ")


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select * From Fam Order by Descrição_Familia"

LocalizacaoBasica = "Descrição_Familia"
DefineObjetos
SelectFromSQL = "Select * From Fam"
SelectFromFiltro = "Select CodFam, Descrição_Familia from Fam"
PesquisaSql = "SELECT CodFam , Descrição_Familia FROM Fam ORDER BY Descrição_Familia"
FlagAbertura = False

OrderBySQL = "[Fam].[Descrição_Familia]"
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
' sera executada logo depois que o usuario iniciar o processo de inclusao addnew
TxtDados(1).SetFocus
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
TxtDados(1).SetFocus
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
Mensagem = "Confirma Exclusão da Familia " & vbCr & Registros.Fields("Descrição_Familia")

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
TxtDados(0).Locked = False
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
'DatDados(1).Refresh
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
If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
    If IsNumeric(TxtDados(0)) = False Then
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = 0 ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=0)) ORDER BY Bol.Data"

    Else
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = " & TxtDados(0) & " ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=" & TxtDados(0) & " )) ORDER BY Bol.Data"
    End If
'    DatDados(2).Refresh
'    DatDados(3).Refresh
End If

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

