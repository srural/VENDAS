VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro do Plano de Contas"
   ClientHeight    =   4815
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   8085
   Icon            =   "FrmCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8085
   Begin TabDlg.SSTab TabCad 
      Height          =   4545
      Left            =   135
      TabIndex        =   14
      Top             =   135
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   8017
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmCta.frx":000C
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
      TabPicture(1)   =   "FrmCta.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PctFundo(0)"
      Tab(1).ControlCount=   1
      Begin VB.Data DatDados 
         Caption         =   "Cta"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Index           =   0
         Left            =   5580
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodCta, Descricao  from  Cta ORDER BY Descricao"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmCta.frx":0044
         Height          =   3255
         Left            =   270
         OleObjectBlob   =   "FrmCta.frx":005E
         TabIndex        =   1
         Top             =   1080
         Width           =   7260
      End
      Begin VB.TextBox TxtPesquisa 
         Height          =   330
         Left            =   270
         TabIndex        =   0
         Top             =   630
         Width           =   5325
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3945
         Index           =   0
         Left            =   -74820
         ScaleHeight     =   3945
         ScaleWidth      =   7275
         TabIndex        =   15
         Top             =   450
         Width           =   7275
         Begin VB.Data DatDados 
            Caption         =   "Cta"
            Connect         =   "Access"
            DatabaseName    =   "C:\AATESTE\SRCaixa\Dados.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Index           =   1
            Left            =   5670
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Select CodCta, Descricao  from  Cta ORDER BY Descricao"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            Height          =   2445
            Left            =   1170
            TabIndex        =   18
            Top             =   1260
            Width           =   4290
            Begin VB.CheckBox ChkDados 
               Caption         =   "&IR"
               DataField       =   "Ir"
               Height          =   195
               Index           =   5
               Left            =   540
               TabIndex        =   12
               Tag             =   "S"
               Top             =   2025
               Width           =   1005
            End
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Simples"
               DataField       =   "Simples"
               Height          =   195
               Index           =   4
               Left            =   540
               TabIndex        =   10
               Tag             =   "S"
               Top             =   1620
               Width           =   1005
            End
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Cofins"
               DataField       =   "Cofins"
               Height          =   195
               Index           =   3
               Left            =   540
               TabIndex        =   8
               Tag             =   "S"
               Top             =   1185
               Width           =   825
            End
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Pis"
               DataField       =   "Pis"
               Height          =   195
               Index           =   2
               Left            =   540
               TabIndex        =   6
               Tag             =   "S"
               Top             =   750
               Width           =   825
            End
            Begin VB.CheckBox ChkDados 
               Caption         =   "&Icms"
               DataField       =   "Icms"
               Height          =   195
               Index           =   1
               Left            =   540
               TabIndex        =   4
               Tag             =   "S"
               Top             =   315
               Width           =   825
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "AliqIcms"
               Height          =   315
               Index           =   2
               Left            =   1755
               TabIndex        =   7
               Tag             =   "S"
               Top             =   750
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "AliqPis"
               Height          =   315
               Index           =   4
               Left            =   1755
               TabIndex        =   11
               Tag             =   "S"
               Top             =   1620
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "AliqCofins"
               Height          =   315
               Index           =   1
               Left            =   1755
               TabIndex        =   5
               Tag             =   "S"
               Top             =   315
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "AliqSimples"
               Height          =   315
               Index           =   3
               Left            =   1755
               TabIndex        =   9
               Tag             =   "S"
               Top             =   1185
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "AliqIr"
               Height          =   315
               Index           =   5
               Left            =   1755
               TabIndex        =   13
               Tag             =   "S"
               Top             =   2025
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Descricao"
            Height          =   315
            Index           =   1
            Left            =   1170
            MaxLength       =   25
            TabIndex        =   3
            Tag             =   "S"
            Top             =   810
            Width           =   4815
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CodCta"
            Height          =   330
            Index           =   0
            Left            =   1170
            TabIndex        =   2
            Tag             =   "S"
            Top             =   315
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "#.##.###"
            PromptChar      =   "_"
         End
         Begin VB.Label LblTitulo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   2925
            TabIndex        =   20
            Top             =   315
            Width           =   4110
         End
         Begin VB.Label LblDados 
            AutoSize        =   -1  'True
            Caption         =   "&Código"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   19
            Top             =   315
            Width           =   495
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Descrição"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   16
            Top             =   765
            Width           =   720
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Conta"
         Height          =   240
         Left            =   270
         TabIndex        =   17
         Top             =   405
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmCta"
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

Public flag As Boolean

Private Sub DefineObjetos()
CTextBox = True
CMaskEditBox = True
CComboBox = False
CListBox = False
CdbCombo = False
CDbList = False
CCheckBox = True
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
End Sub

Private Sub ChkDados_Click(Index As Integer)
If ChkDados(Index) = 1 Then
    MebDados(Index).Enabled = True
Else
    MebDados(Index).Enabled = False
End If
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodCta = " & """" & DBGrid1.Columns(0) & """"
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte
MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
Set Formulário = Me

AtivarMenu (True)
Formulário.Left = 700
Formulário.Top = 0

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)

If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
TabCad.Tab = 0
AbilitaBotoes (False)

PosMovimentacao

TxtPesquisa.SetFocus

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

Set Registros = Banco.OpenRecordset("SELECT * FROM Cta ORDER BY CodCta")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select CodCta, Descricao from Cta ORDER BY CodCta"
DatDados(1).RecordSource = "Select CodCta, Descricao from Cta ORDER BY Descricao"

LocalizacaoBasica = "Descricao"
DefineObjetos
SelectFromSQL = "Select * From Cta"
SelectFromFiltro = "Select CodCta , Descricao from Cta"
PesquisaSql = "SELECT CodCta , Descricao FROM Cta ORDER BY Descricao"
FlagAbertura = False

OrderBySQL = "[Bco].[Nome]"

flag = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Registros.EditMode <> dbEditNone Then
   Confirmacoes (14)
End If

Registros.Close

AtivarMenu (False)

End Sub

Public Sub PosInclusao()

MebDados(0).SetFocus
TabCad.TabEnabled(0) = False
MebDados(1).Text = 0
MebDados(2).Text = 0
MebDados(3).Text = 0
MebDados(4).Text = 0
MebDados(5).Text = 0
flag = True

End Sub

Public Function PermiteInclusao() As Boolean

PermiteInclusao = True
PosMovimentacao

End Function

Public Sub PosAlteracao()
MebDados(0).Enabled = False
TxtDados(1).SetFocus
TabCad.TabEnabled(0) = False

Atualizar
End Sub

Public Function PermiteAlteracao() As Boolean
PermiteAlteracao = True

End Function

Public Sub PosMovimentacao()
Dim x As Byte

If Registros.RecordCount <> 0 Then
    For x = 1 To 5
        If ChkDados(x).Value = 1 Then
            MebDados(x).Enabled = True
        Else
            MebDados(x).Enabled = False
        End If
    Next
End If

End Sub

Public Sub PosExclusao()

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
Dim Mensagem As String
If Mid(MebDados(0), 2, 2) = "00" Or Mid(MebDados(0), 4, 3) = "000" Then
    MsgBox "Esta é uma conta de grupo que não pode ser excluída", vbInformation, App.Title
Else
    Mensagem = "Confirma Exclusão do Conta " & vbCr & Registros.Fields("Descricao")
    
    If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        PermiteExclusao = True
    Else
        PermiteExclusao = False
    End If
End If
End Function

Public Sub PreConfirmacao()

End Sub

Public Sub PosConfirmacao()

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
MebDados(0).Enabled = True
TabCad.TabEnabled(0) = True
LblTitulo1.Caption = ""
flag = False

End Function

Public Sub PosCancelamento()
MebDados(0).Enabled = True
TabCad.TabEnabled(0) = True
LblTitulo1.Caption = ""

flag = False
End Sub

Public Sub Atualizar()

'Registro.Requery
DatDados(0).Refresh

End Sub

Private Sub MebDados_Change(Index As Integer)
Dim Criterio As String
If flag Then
    If Index = 0 Then
        If Len(MebDados(0)) = 6 Then
              Criterio = "CodCta = " & """" & MebDados(0) & """"
              DatDados(1).Recordset.FindFirst Criterio
              If Not DatDados(1).Recordset.NoMatch Then
                 LblTitulo1.Caption = DatDados(1).Recordset.Fields("Descricao")
              End If
        End If
        
    End If
End If
End Sub

Private Sub MebDados_LostFocus(Index As Integer)
Dim Criterio As String
If flag Then
    If Len(MebDados(0)) = 6 Then
          Criterio = "CodCta = " & """" & MebDados(0) & """"
          DatDados(1).Recordset.FindFirst Criterio
          If DatDados(1).Recordset.NoMatch Then
              TxtDados(1).SetFocus
          Else
              TxtDados(1).Text = DatDados(1).Recordset.Fields("Descricao")
              If MsgBox("Conta já Cadastrada ! Deseja Alterar ?", vbYesNo, App.Title) = vbYes Then
                    Confirmacoes 15
                    Registros.FindFirst Criterio
                    MostraDados
                    Manutencoes 2
              Else
                    Confirmacoes 15
              End If
          End If
    End If
End If
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
If TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        Criterio = DatDados(0).Recordset.Fields(0).Name & " =  " & """" & TxtPesquisa & """"
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
