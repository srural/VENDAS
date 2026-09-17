VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmBco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Bancos"
   ClientHeight    =   4080
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9720
   Icon            =   "FrmBco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9720
   Begin TabDlg.SSTab TabCad 
      Height          =   3780
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   6668
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmBco.frx":000C
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
      TabPicture(1)   =   "FrmBco.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LblDados(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "PctFundo(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtDados(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Data DatDados 
         Caption         =   "Bancos"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -68340
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodBco , Nome, Agencia, Contato from  Bco ORDER BY Nome"
         Top             =   1845
         Visible         =   0   'False
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmBco.frx":0044
         Height          =   2400
         Left            =   -74730
         OleObjectBlob   =   "FrmBco.frx":005E
         TabIndex        =   17
         Top             =   1125
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
         DataField       =   "CodBco"
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
         Height          =   2460
         Index           =   0
         Left            =   180
         ScaleHeight     =   2460
         ScaleWidth      =   9120
         TabIndex        =   9
         Top             =   1125
         Width           =   9120
         Begin VB.TextBox TxtDados 
            DataField       =   "Nome"
            Height          =   315
            Index           =   1
            Left            =   1575
            MaxLength       =   25
            TabIndex        =   2
            Tag             =   "S"
            Top             =   135
            Width           =   4815
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Contato"
            Height          =   315
            Index           =   7
            Left            =   1575
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "S"
            Top             =   1035
            Width           =   4095
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   4365
            TabIndex        =   5
            Top             =   1890
            Width           =   2415
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Agencia"
            Height          =   315
            Index           =   2
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "S"
            Top             =   585
            Width           =   2880
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone"
            Height          =   315
            Index           =   2
            Left            =   1575
            TabIndex        =   6
            Tag             =   "S"
            Top             =   1455
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
            Left            =   1575
            TabIndex        =   7
            Tag             =   "S"
            Top             =   1890
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###) 9###-####"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Agência"
            Height          =   195
            Index           =   2
            Left            =   810
            TabIndex        =   14
            Top             =   540
            Width           =   585
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Nome do Banco"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   13
            Top             =   90
            Width           =   1155
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Contato"
            Height          =   195
            Index           =   16
            Left            =   855
            TabIndex        =   12
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Fone"
            Height          =   195
            Index           =   8
            Left            =   1035
            TabIndex        =   11
            Top             =   1485
            Width           =   360
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   17
            Left            =   1125
            TabIndex        =   10
            Top             =   1890
            Width           =   255
         End
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Código"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Banco"
         Height          =   240
         Left            =   -74730
         TabIndex        =   15
         Top             =   405
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmBco"
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
CdbCombo = False
CDbList = False
CCheckBox = False
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodBco = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte
MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
Set Formulário = Me

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

Set Registros = Banco.OpenRecordset("SELECT * FROM Bco ORDER BY Nome")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select CodBco, Nome, Agencia, Contato from Bco ORDER BY Nome"

LocalizacaoBasica = "Nome"
DefineObjetos
SelectFromSQL = "Select * From Bco"
SelectFromFiltro = "Select CodBco , Nome, Agencia, Contato from Bco"
PesquisaSql = "SELECT CodBco , Nome FROM Bco ORDER BY Nome"
FlagAbertura = False

OrderBySQL = "[Bco].[Nome]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Registros.EditMode <> dbEditNone Then
   Confirmacoes (14)
End If

Registros.Close

AtivarMenu (False)

End Sub

Public Sub PosInclusao()

TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False


End Sub

Public Function PermiteInclusao() As Boolean

PermiteInclusao = True
PosMovimentacao

End Function

Public Sub PosAlteracao()
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False

Atualizar
End Sub

Public Function PermiteAlteracao() As Boolean
PermiteAlteracao = True

End Function

Public Sub PosMovimentacao()

End Sub

Public Sub PosExclusao()

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
Dim Mensagem As String
Mensagem = "Confirma Exclusão do Banco " & vbCr & Registros.Fields("Nome")

If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
    PermiteExclusao = True
Else
    PermiteExclusao = False
End If

End Function

Public Sub PreConfirmacao()

End Sub

Public Sub PosConfirmacao()

TxtDados(0).Locked = False
Atualizar
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

End Function

Public Sub PosCancelamento()

TabCad.TabEnabled(0) = True

End Sub

Public Sub Atualizar()

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
