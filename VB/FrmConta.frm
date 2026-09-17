VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmConta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contas de Movimento"
   ClientHeight    =   5880
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9720
   Icon            =   "FrmConta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9720
   Begin TabDlg.SSTab TabCad 
      Height          =   5580
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9843
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmConta.frx":000C
      Tab(0).ControlEnabled=   0   'False
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
      TabPicture(1)   =   "FrmConta.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LblDados(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "PctFundo(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtDados(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DatDados(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FrmSocios"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame FrmSocios 
         Caption         =   "Sócios"
         Height          =   2760
         Left            =   405
         TabIndex        =   17
         Top             =   2610
         Width           =   8745
         Begin VB.CommandButton CmdExluir 
            Caption         =   "&Excluir"
            Height          =   315
            Left            =   7065
            TabIndex        =   25
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Data DatDados 
            Caption         =   "Titulares"
            Connect         =   "Access"
            DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   3
            Left            =   4050
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT CodTit, Nome_Titular  FROM Tit  ORDER BY Nome_Titular"
            Top             =   2025
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Data DatDados 
            Caption         =   "Sócios"
            Connect         =   "Access"
            DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   2
            Left            =   855
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   $"FrmConta.frx":0044
            Top             =   1845
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton CmdLancar 
            Caption         =   "&Lançar"
            Height          =   315
            Left            =   7065
            TabIndex        =   21
            Top             =   450
            Width           =   1365
         End
         Begin MSMask.MaskEdBox MebPerc 
            Height          =   315
            Left            =   5175
            TabIndex        =   20
            Top             =   450
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo PesqTit 
            Bindings        =   "FrmConta.frx":0126
            Height          =   315
            Left            =   225
            TabIndex        =   19
            Top             =   450
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome_Titular"
            BoundColumn     =   "CodTit"
            Text            =   ""
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "FrmConta.frx":0140
            Height          =   1725
            Left            =   225
            OleObjectBlob   =   "FrmConta.frx":015A
            TabIndex        =   18
            Top             =   900
            Width           =   6135
         End
         Begin VB.Label Label5 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6435
            TabIndex        =   24
            Top             =   495
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Sócio"
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   225
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Participação (%)"
            Height          =   195
            Left            =   5175
            TabIndex        =   22
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.Data DatDados 
         Caption         =   "Banco"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   6840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT CodBco, Nome  FROM Bco  ORDER BY NOme"
         Top             =   855
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data DatDados 
         Caption         =   "Contas"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -69375
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"FrmConta.frx":0D01
         Top             =   3150
         Visible         =   0   'False
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmConta.frx":0D90
         Height          =   4200
         Left            =   -74730
         OleObjectBlob   =   "FrmConta.frx":0DAA
         TabIndex        =   2
         Top             =   1080
         Width           =   8925
      End
      Begin VB.TextBox TxtPesquisa 
         Height          =   330
         Left            =   -74730
         TabIndex        =   0
         Top             =   630
         Width           =   5415
      End
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "CodConta"
         Height          =   315
         Index           =   0
         Left            =   1755
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
         Height          =   1560
         Index           =   0
         Left            =   180
         ScaleHeight     =   1560
         ScaleWidth      =   9120
         TabIndex        =   9
         Top             =   1125
         Width           =   9120
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmConta.frx":1AE1
            DataField       =   "Banco"
            Height          =   315
            Index           =   0
            Left            =   4995
            TabIndex        =   5
            Tag             =   "SI"
            Top             =   585
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nome"
            BoundColumn     =   "CodBco"
            Text            =   ""
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Descricao"
            Height          =   315
            Index           =   1
            Left            =   1575
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "S"
            Top             =   135
            Width           =   3870
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Tipo"
            Height          =   315
            Index           =   2
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   4
            Tag             =   "S"
            Top             =   585
            Width           =   2520
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "SaldoInicial"
            Height          =   315
            Index           =   1
            Left            =   4365
            TabIndex        =   7
            Tag             =   "S"
            Top             =   1035
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Data"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1575
            TabIndex        =   6
            Tag             =   "S"
            Top             =   1035
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Saldo Inicial"
            Height          =   195
            Left            =   3345
            TabIndex        =   16
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Banco"
            Height          =   195
            Index           =   3
            Left            =   4410
            TabIndex        =   13
            Top             =   585
            Width           =   465
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Tipo"
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   12
            Top             =   585
            Width           =   315
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Data de &Abertura"
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   11
            Top             =   1035
            Width           =   1215
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Descrição"
            Height          =   195
            Index           =   1
            Left            =   675
            TabIndex        =   10
            Top             =   135
            Width           =   720
         End
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Código"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome da Conta"
         Height          =   240
         Left            =   -74730
         TabIndex        =   14
         Top             =   405
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmConta"
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

CTextBox = True
CMaskEditBox = True
CComboBox = False
CListBox = False
CdbCombo = True
CDbList = False
CCheckBox = False
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
End Sub

Private Sub CmdExluir_Click()
If MsgBox("Confirma Exclusão do Sócio ?", vbYesNo, App.Title) = vbYes Then
    DatDados(2).Recordset.Delete
End If
End Sub

Private Sub CmdLancar_Click()
If PesqTit <> "" And MebPerc <> "" Then
    DatDados(2).Recordset.AddNew
    DatDados(2).Recordset.Fields("CodCont") = TxtDados(0)
    DatDados(2).Recordset.Fields("CodTit") = PesqTit.BoundText
    DatDados(2).Recordset.Fields("Participacao") = MebPerc
    DatDados(2).Recordset.Update
    PesqTit = ""
    MebPerc = ""
Else
    MsgBox "Sócio ou participação inválidos ", vbInformation, App.Title
End If
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodConta = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte

MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)

If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
TabCad.Tab = 0
AbilitaBotoes (False)
TxtPesquisa.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_Load()
Dim Componentes As Variant
Set Registros = Banco.OpenRecordset("SELECT Conta.CodConta, Conta.Descricao, Conta.Tipo, Conta.Banco, Conta.Data, Conta.SaldoInicial, Bco.Nome FROM Bco INNER JOIN Conta ON Bco.CodBco = Conta.Banco ORDER BY Conta.Descricao")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "SELECT Conta.CodConta, Conta.Descricao, Conta.Tipo, Bco.Nome FROM Bco INNER JOIN Conta ON Bco.CodBco = Conta.Banco ORDER BY Conta.Descricao"

LocalizacaoBasica = "Descricao"
DefineObjetos
SelectFromSQL = "Select * From Conta"
SelectFromFiltro = "Select CodConta, Descricao, Tipo from Conta"
PesquisaSql = "SELECT CodConta, Descricao, Tipo from Conta ORDER BY Descricao"
FlagAbertura = False

OrderBySQL = "[Conta].[Descricao]"


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
   Confirmacoes (14)
End If

Registros.Close

AtivarMenu (False)

End Sub

Public Sub PosInclusao()
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False
FrmSocios.Enabled = False
End Sub

Public Function PermiteInclusao() As Boolean
PermiteInclusao = True
PosMovimentacao

End Function

Public Sub PosAlteracao()
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False
Atualizar
FrmSocios.Enabled = False
End Sub

Public Function PermiteAlteracao() As Boolean

PermiteAlteracao = True

End Function

Public Sub PosMovimentacao()
If Registros.RecordCount <> 0 Then

End If

End Sub

Public Sub PosExclusao()
Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
Dim Mensagem As String
Mensagem = "Confirma Exclusão da Conta " & vbCr & Registros.Fields("Descricao")

If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
    PermiteExclusao = True
Else
    PermiteExclusao = False
End If

End Function

Public Sub PreConfirmacao() 'Será executada antes das rotinas de salvar ou cancelar

End Sub

Public Sub PosConfirmacao() 'Será executada deposi dos processos de salvar ou cancelar

TxtDados(0).Locked = False
Atualizar
FrmSocios.Enabled = True

End Sub

Public Function PermiteCancelamento() As Boolean

PermiteCancelamento = True

End Function

Public Function PermiteSalvar() As Boolean

PermiteSalvar = True

End Function
Public Function PermiteContinuarMostrando() As Boolean

End Function

Public Function PermiteContinuarSalvar() As Boolean
PermiteContinuarSalvar = True

End Function

Public Function PosSalvamento() As Boolean

PosSalvamento = True

TabCad.TabEnabled(0) = True
FrmSocios.Enabled = True

End Function

Public Sub PosCancelamento()

TabCad.TabEnabled(0) = True
FrmSocios.Enabled = True

End Sub

Private Sub Txtdados_Change(Index As Integer)

If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
    If IsNumeric(TxtDados(0)) = False Then
        DatDados(2).RecordSource = "SELECT Soc.CodCont,Soc.CodTit, Tit.Nome_Titular, Soc.Participacao FROM Tit INNER JOIN (Conta INNER JOIN Soc ON Conta.CodConta = Soc.CodCont) ON Tit.CodTit = Soc.CodTit WHERE (((Soc.CodCont)=0)) ORDER BY Tit.Nome_Titular"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=0)) ORDER BY Bol.Data"

    Else
        DatDados(2).RecordSource = "SELECT Soc.CodCont, Soc.CodTit, Tit.Nome_Titular, Soc.Participacao FROM Tit INNER JOIN (Conta INNER JOIN Soc ON Conta.CodConta = Soc.CodCont) ON Tit.CodTit = Soc.CodTit WHERE (((Soc.CodCont)= " & TxtDados(0) & ")) ORDER BY Tit.Nome_Titular"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=" & TxtDados(0) & " )) ORDER BY Bol.Data"
    End If
    DatDados(2).Refresh
End If

End Sub


Public Sub Atualizar()
' Atualiza a variavel registro incluindo os novos registros cadastrados e excluindo os que foram deletados
'Registro.Requery
DatDados(0).Refresh
End Sub

Private Sub TabCad_Click(previoustab As Integer)
Select Case TabCad.Tab
    Case 0
        Atualizar
        AbilitaBotoes (False)
    Case 1
        AbilitaBotoes (True)
End Select

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
