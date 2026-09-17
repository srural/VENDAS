VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmCtaMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contas de Movimento"
   ClientHeight    =   5370
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9720
   Icon            =   "FrmCtaMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9720
   Begin TabDlg.SSTab TabCad 
      Height          =   5130
      Left            =   135
      TabIndex        =   16
      Top             =   135
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9049
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmCtaMov.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtPesquisa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DBGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DatDados(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cadastro"
      TabPicture(1)   =   "FrmCtaMov.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDados(0)"
      Tab(1).Control(1)=   "PctFundo(0)"
      Tab(1).Control(2)=   "TxtDados(0)"
      Tab(1).Control(3)=   "DatDados(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Contas"
      TabPicture(2)   =   "FrmCtaMov.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Data DatDados 
         Caption         =   "Cidades"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -67845
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Tit.Cidade FROM Tit GROUP BY Tit.Cidade ORDER BY Tit.Cidade"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data DatDados 
         Caption         =   "Titulares"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Index           =   0
         Left            =   7245
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodTit, Nome_Titular, Cidade, Uf from Tit ORDER BY Nome_Titular"
         Top             =   585
         Visible         =   0   'False
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmCtaMov.frx":0060
         Height          =   3885
         Left            =   270
         OleObjectBlob   =   "FrmCtaMov.frx":007A
         TabIndex        =   37
         Top             =   1080
         Width           =   8925
      End
      Begin VB.TextBox TxtPesquisa 
         Height          =   330
         Left            =   270
         TabIndex        =   0
         Top             =   630
         Width           =   5415
      End
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "CodTit"
         Height          =   315
         Index           =   0
         Left            =   -73245
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "N"
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3720
         Index           =   0
         Left            =   -74820
         ScaleHeight     =   3720
         ScaleWidth      =   9120
         TabIndex        =   18
         Top             =   1125
         Width           =   9120
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmCtaMov.frx":0DA9
            DataField       =   "Cidade"
            Height          =   315
            Index           =   0
            Left            =   1575
            TabIndex        =   38
            Tag             =   "S"
            Top             =   1485
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Cidade"
            Text            =   ""
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Nome_Titular"
            Height          =   315
            Index           =   1
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   2
            Tag             =   "S"
            Top             =   135
            Width           =   4815
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "Uf"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmCtaMov.frx":0DC3
            Left            =   6165
            List            =   "FrmCtaMov.frx":0E1B
            Sorted          =   -1  'True
            TabIndex        =   6
            Tag             =   "S"
            Top             =   1485
            Width           =   855
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Nro"
            Height          =   315
            Index           =   6
            Left            =   6660
            MaxLength       =   5
            TabIndex        =   4
            Tag             =   "S"
            Top             =   585
            Width           =   855
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Bairro"
            Height          =   315
            Index           =   7
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   5
            Tag             =   "S"
            Top             =   1035
            Width           =   2880
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "EMail"
            Height          =   315
            Index           =   5
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "S"
            Top             =   3150
            Width           =   6180
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1440
            TabIndex        =   8
            Top             =   1890
            Width           =   2400
            Begin VB.OptionButton Opt1 
               Caption         =   "C.N.P.J."
               Height          =   195
               Index           =   1
               Left            =   1035
               TabIndex        =   22
               Top             =   135
               Width           =   1230
            End
            Begin VB.OptionButton Opt1 
               Caption         =   "C.P.F."
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   21
               Top             =   135
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   4365
            TabIndex        =   10
            Top             =   1890
            Width           =   2415
            Begin VB.OptionButton Opt2 
               Caption         =   "R.G."
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton Opt2 
               Caption         =   "Insc. Est."
               Height          =   195
               Index           =   1
               Left            =   945
               TabIndex        =   19
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Endereco"
            Height          =   315
            Index           =   2
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   3
            Tag             =   "S"
            Top             =   585
            Width           =   4275
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CgcCpf"
            Height          =   315
            Index           =   1
            Left            =   1575
            TabIndex        =   9
            Tag             =   "S"
            Top             =   2295
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cep"
            Height          =   315
            Index           =   0
            Left            =   7650
            TabIndex        =   7
            Tag             =   "S"
            Top             =   1485
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone"
            Height          =   315
            Index           =   2
            Left            =   1575
            TabIndex        =   12
            Tag             =   "S"
            Top             =   2722
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###) 9###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fax"
            Height          =   315
            Index           =   6
            Left            =   3870
            TabIndex        =   13
            Tag             =   "S"
            Top             =   2745
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###) 9###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Celular"
            Height          =   315
            Index           =   7
            Left            =   6165
            TabIndex        =   14
            Tag             =   "S"
            Top             =   2745
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###) 9###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Insc_Rg"
            Height          =   315
            Index           =   3
            Left            =   4500
            TabIndex        =   11
            Tag             =   "S"
            Top             =   2250
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Documentos"
            Height          =   195
            Left            =   495
            TabIndex        =   36
            Top             =   2250
            Width           =   900
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   255
            Index           =   3
            Left            =   900
            TabIndex        =   33
            Top             =   1485
            Width           =   495
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   255
            Index           =   2
            Left            =   660
            TabIndex        =   32
            Top             =   540
            Width           =   735
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "C.E.P"
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   31
            Top             =   1485
            Width           =   495
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Uf"
            Height          =   255
            Index           =   4
            Left            =   5850
            TabIndex        =   30
            Top             =   1485
            Width           =   255
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nome do Titular"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   29
            Top             =   90
            Width           =   1125
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   14
            Left            =   6030
            TabIndex        =   28
            Top             =   585
            Width           =   555
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Bairro"
            Height          =   195
            Index           =   16
            Left            =   990
            TabIndex        =   27
            Top             =   1035
            Width           =   405
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fone"
            Height          =   195
            Index           =   8
            Left            =   1035
            TabIndex        =   26
            Top             =   2745
            Width           =   360
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   17
            Left            =   3510
            TabIndex        =   25
            Top             =   2745
            Width           =   255
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Celular"
            Height          =   195
            Index           =   18
            Left            =   5625
            TabIndex        =   24
            Top             =   2745
            Width           =   480
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   11
            Left            =   960
            TabIndex        =   23
            Top             =   3150
            Width           =   435
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Titular"
         Height          =   870
         Left            =   -74730
         TabIndex        =   17
         Top             =   495
         Width           =   8565
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   -73920
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome da Conta"
         Height          =   240
         Left            =   270
         TabIndex        =   34
         Top             =   405
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmCtaMov"
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

CTextBox = True
CMaskEditBox = True
CComboBox = True
CListBox = False
CdbCombo = True
CDbList = False
CCheckBox = False
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo
' Titular = Registros.Fields("CodTit")
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodTit = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte

MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)
Formulário.Left = 700
Formulário.Top = 0

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)

' testa se o banco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
TabCad.Tab = 0
AbilitaBotoes (False)
TxtPesquisa.SetFocus

'PosMovimentação
' DbcPesquisa.SetFocus

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
    
Set Registros = Banco.OpenRecordset("SELECT * FROM Tit ORDER BY Nome_Titular")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select CodTit, Nome_Titular, Cidade, Uf from Tit ORDER BY Nome_Titular"

LocalizacaoBasica = "Nome_Titular"
DefineObjetos
SelectFromSQL = "Select * From Tit"
SelectFromFiltro = "Select CodTit, Nome_Titular, Cidade, Uf from Tit"
PesquisaSql = "SELECT CodTit , Nome_Titular FROM Tit ORDER BY Nome_Titular"
FlagAbertura = False

OrderBySQL = "[Tit].[Nome_Titular]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
   Confirmacoes (14)
End If

Registros.Close

AtivarMenu (False)

End Sub

Public Sub PosInclusao()
' sera executada logo depois que o usuario iniciar o processo de inclusao addnew
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False
TabCad.TabEnabled(2) = False

End Sub

Public Function PermiteInclusao() As Boolean
' Será executado antes do processo de inclusao (addnew)
' Aqui deve ser feitos os processo p/ que o usuário pode ou não incluir um novo registro
' MsgBox "Neste formulário não pode ser feitas Inclusão de novos registros", vbInformation, App.Title

PermiteInclusao = True
PosMovimentacao

End Function

Public Sub PosAlteracao()
' sera executada logo depois que o usuario iniciar o processo de alteracao
'TxtDados(0).BackColor = QBColor(8) 'Muda a Cor de fundo do código
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False
TabCad.TabEnabled(2) = False
Atualizar
End Sub

Public Function PermiteAlteracao() As Boolean
' Será executado antes do processo de Alteracao

PermiteAlteracao = True

End Function

Public Sub PosMovimentacao()
If Registros.RecordCount = 0 Then
        MsgBox "Não existem dados cadastrados", vbInformation, App.Title
Else
    If Len(Registros.Fields("CgcCpf")) <= 11 Then
        Opt1(0).Value = True
        MebDados(1).Mask = "###.###.###-##"
    Else
        Opt1(1).Value = True
        MebDados(1).Mask = "##.###.###/####/##"
    End If
    If Len(Registros.Fields("Insc_Rg")) <= 8 Then
        Opt2(0).Value = True
        MebDados(3).Mask = "##.###.###"
    Else
        Opt2(1).Value = True
        MebDados(3).Mask = "&&&&&&&&&&&&&&&"
    End If
End If
End Sub

Public Sub PosExclusao()
' sera executada logo depois que o usuario iniciar o processo de exclusao

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
' Será executado antes do processo de exclusao
Dim Mensagem As String
Mensagem = "Confirma Exclusão do Titular " & vbCr & Registros.Fields("Nome_Titular")

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

TabCad.TabEnabled(0) = True
TabCad.TabEnabled(2) = True

End Function

Public Sub PosCancelamento()

'sera executado logo após do processo de cancelamento da aplicaçaõ

TabCad.TabEnabled(0) = True
TabCad.TabEnabled(2) = True

End Sub

Public Sub Atualizar()
' Atualiza a variavel registro incluindo os novos registros cadastrados e excluindo os que foram deletados
'Registro.Requery
DatDados(0).Refresh
DatDados(1).Refresh
End Sub

Private Sub Opt1_Click(Index As Integer)
Select Case Index
    Case 0
        MebDados(1).Mask = "###.###.###-##"
    Case 1
        MebDados(1).Mask = "##.###.###/####/##"
End Select
MebDados(1).SetFocus
End Sub

Private Sub Opt2_Click(Index As Integer)
Select Case Index
    Case 0
        MebDados(3).Mask = "##.###.###"
    Case 1
        MebDados(3).Mask = ""
        MebDados(3).Text = ""
End Select
MebDados(3).SetFocus
End Sub


Private Sub TabCad_Click(previoustab As Integer)
Select Case TabCad.Tab
    Case 0
        Atualizar
        AbilitaBotoes (False)
        CmdPesquisar.Enabled = True
    Case 1
        AbilitaBotoes (True)
    Case 2
        AbilitaBotoes (False)
End Select

End Sub

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
        PosMovimentacao
    End If
End If


End Sub

Private Sub TxtPesquisa_LostFocus()
    DBGrid1_DBLClick
End Sub
