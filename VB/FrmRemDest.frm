VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmRemDest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Remetentes e Destinatários"
   ClientHeight    =   5370
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9720
   Icon            =   "FrmRemDest.frx":0000
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
      TabIndex        =   21
      Top             =   135
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9049
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmRemDest.frx":000C
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
      TabPicture(1)   =   "FrmRemDest.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDados(0)"
      Tab(1).Control(1)=   "PctFundo(0)"
      Tab(1).Control(2)=   "TxtDados(0)"
      Tab(1).Control(3)=   "DatDados(1)"
      Tab(1).ControlCount=   4
      Begin VB.Data DatDados 
         Caption         =   "Cidades"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -67710
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Rem.Cidade FROM Rem GROUP BY Rem.Cidade ORDER BY Rem.Cidade"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data DatDados 
         Caption         =   "Remetentes"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   6705
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodRem, Nome, Cidade, Uf from Rem ORDER BY Nome"
         Top             =   1890
         Visible         =   0   'False
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmRemDest.frx":0044
         Height          =   3885
         Left            =   270
         OleObjectBlob   =   "FrmRemDest.frx":005E
         TabIndex        =   2
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
         DataField       =   "CodRem"
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
         TabIndex        =   22
         Top             =   1125
         Width           =   9120
         Begin MSDBCtls.DBCombo DbcDados 
            Bindings        =   "FrmRemDest.frx":0D81
            DataField       =   "Cidade"
            Height          =   315
            Index           =   0
            Left            =   1575
            TabIndex        =   7
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
            DataField       =   "Nome"
            Height          =   315
            Index           =   1
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "S"
            Top             =   135
            Width           =   4815
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "Uf"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmRemDest.frx":0D9B
            Left            =   6165
            List            =   "FrmRemDest.frx":0DF3
            Sorted          =   -1  'True
            TabIndex        =   8
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
            TabIndex        =   5
            Tag             =   "S"
            Top             =   585
            Width           =   855
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Bairro"
            Height          =   315
            Index           =   7
            Left            =   1575
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "S"
            Top             =   1035
            Width           =   2880
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "EMail"
            Height          =   315
            Index           =   5
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   17
            Tag             =   "S"
            Top             =   3150
            Width           =   6180
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1440
            TabIndex        =   10
            Top             =   1890
            Width           =   2400
            Begin VB.OptionButton Opt1 
               Caption         =   "C.N.P.J."
               Height          =   195
               Index           =   1
               Left            =   1035
               TabIndex        =   19
               Top             =   135
               Width           =   1230
            End
            Begin VB.OptionButton Opt1 
               Caption         =   "C.P.F."
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   18
               Top             =   135
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   4365
            TabIndex        =   12
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
               TabIndex        =   23
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Endereco"
            Height          =   315
            Index           =   2
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   4
            Tag             =   "S"
            Top             =   585
            Width           =   4275
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CgcCpf"
            Height          =   315
            Index           =   1
            Left            =   1575
            TabIndex        =   11
            Tag             =   "S"
            Top             =   2295
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cep"
            Height          =   315
            Index           =   0
            Left            =   7650
            TabIndex        =   9
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   13
            Tag             =   "S"
            Top             =   2250
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##.###.###"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Documentos"
            Height          =   195
            Left            =   495
            TabIndex        =   37
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
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   540
            Width           =   735
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "C.E.P"
            Height          =   255
            Index           =   5
            Left            =   7200
            TabIndex        =   32
            Top             =   1485
            Width           =   495
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Uf"
            Height          =   255
            Index           =   4
            Left            =   5850
            TabIndex        =   31
            Top             =   1485
            Width           =   255
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Index           =   1
            Left            =   975
            TabIndex        =   30
            Top             =   90
            Width           =   420
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Index           =   14
            Left            =   6030
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   2745
            Width           =   360
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   17
            Left            =   3510
            TabIndex        =   26
            Top             =   2745
            Width           =   255
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "Celular"
            Height          =   195
            Index           =   18
            Left            =   5625
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   3150
            Width           =   435
         End
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   -73920
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   240
         Left            =   270
         TabIndex        =   35
         Top             =   405
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmRemDest"
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
CComboBox = True
CListBox = False
CdbCombo = True
CDbList = False
CCheckBox = False
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodRem = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte

MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)
'Formulário.Left = 700
'Formulário.Top = 0

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
Set Registros = Banco.OpenRecordset("SELECT * FROM Rem ORDER BY Nome")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select CodRem, Nome, Cidade, Uf from Rem ORDER BY Nome"

LocalizacaoBasica = "Nome"
DefineObjetos
SelectFromSQL = "Select * From Rem"
SelectFromFiltro = "Select CodRem, Nome, Cidade, Uf from Rem"
PesquisaSql = "SELECT CodRem , Nome FROM Rem ORDER BY Nome"
FlagAbertura = False

OrderBySQL = "[Rem].[Nome]"
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
If Registros.RecordCount <> 0 Then
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
Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
Dim Mensagem As String
Mensagem = "Confirma Exclusão do Remetente " & vbCr & Registros.Fields("Nome")

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
' Atualiza a variavel registro incluindo os novos registros cadastrados e excluindo os que foram deletados
'Registro.Requery
DatDados(0).Refresh
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
