VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormLista 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   13185
   Icon            =   "FormLista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13185
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   13215
      TabIndex        =   11
      Top             =   7875
      Width           =   13245
   End
   Begin VB.CommandButton BtMarca 
      Caption         =   "Marca/Desmarca"
      Height          =   330
      Left            =   135
      TabIndex        =   10
      Top             =   7425
      Width           =   1635
   End
   Begin VB.Data DatDados 
      Caption         =   "Hist"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome FROM  HISTEST GROUP BY Nome ORDER BY Nome"
      Top             =   3060
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton CmdLer 
      Caption         =   "&Carrega Lista"
      Height          =   375
      Left            =   10305
      TabIndex        =   8
      Top             =   1035
      Width           =   1095
   End
   Begin VB.ListBox LstPrd 
      Appearance      =   0  'Flat
      Height          =   6330
      ItemData        =   "FormLista.frx":000C
      Left            =   135
      List            =   "FormLista.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   990
      Width           =   4380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13200
      TabIndex        =   4
      Top             =   0
      Width           =   13200
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Produtos "
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
         Left            =   225
         TabIndex        =   5
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.CommandButton CmbMonta 
      Caption         =   "Monta Lista >>"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   990
      Width           =   1395
   End
   Begin VB.CommandButton CmdApaga 
      Caption         =   "&Apaga Lista"
      Height          =   375
      Left            =   11745
      TabIndex        =   2
      Top             =   1035
      Width           =   1170
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   7380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4095
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Lista"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6165
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodPrd, Descrição_Produto From Prd Order By Descrição_Produto"
      Top             =   5355
      Visible         =   0   'False
      Width           =   2430
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormLista.frx":0010
      Height          =   6075
      Left            =   4680
      OleObjectBlob   =   "FormLista.frx":002A
      TabIndex        =   0
      Top             =   1530
      Width           =   8280
   End
   Begin MSDBCtls.DBCombo TxtNome 
      Bindings        =   "FormLista.frx":0F21
      Height          =   360
      Left            =   6210
      TabIndex        =   9
      Top             =   990
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Lista"
      Height          =   195
      Left            =   6300
      TabIndex        =   7
      Top             =   765
      Width           =   1065
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   1
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "FormLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas


Private Sub BtSair_Click()
Unload Me
End Sub


Private Sub BtMarca_Click()
Dim y As Integer
    For y = 0 To LstPrd.ListCount - 1
        If LstPrd.Selected(y) = False Then
            LstPrd.Selected(y) = True
        Else
            LstPrd.Selected(y) = False
        End If
    Next
End Sub


Private Sub CmbMonta_Click()
Dim y As Integer
If TxtNome.Text <> "" Then
    For y = 0 To LstPrd.ListCount - 1
        If LstPrd.Selected(y) = True Then
            DatDados(0).Recordset.AddNew
            DatDados(0).Recordset.Fields("Nome") = TxtNome
            DatDados(0).Recordset.Fields("CodPrd") = Left(LstPrd.List(y), 6)
            DatDados(0).Recordset.Fields("Valor") = 0
            DatDados(0).Recordset.Update
        End If
    Next
    DatDados(0).Recordset.MoveFirst
    DatDados(0).Refresh
Else
    MsgBox "Define um nome para a lista", vbInformation
    TxtNome.SetFocus
End If
End Sub



Private Sub CmdApaga_Click()
Dim SQL As String
If DatDados(0).Recordset.RecordCount > 1 Then
    If MsgBox("Deseja limpar a lista ? ", vbYesNo) = vbYes Then
        SQL = "DELETE HISTEST.*, HISTEST.Nome FROM HISTEST WHERE Nome = " & """" & TxtNome & """"
        Banco.Execute SQL
        DatDados(0).Refresh
    End If
End If
End Sub

Private Sub CmdLer_Click()
If TxtNome.Text <> "" Then
    DatDados(0).RecordSource = "SELECT HISTEST.Nome, HISTEST.CodPrd, PRD.Descrição_Produto, HISTEST.Valor, HISTEST.Ordem FROM HISTEST INNER JOIN PRD ON HISTEST.CodPrd = PRD.CodPrd where nome = " & """" & TxtNome & """" & " Order By HISTEST.Ordem"
    DatDados(0).Refresh
Else
    MsgBox "Digite o nome da lista", vbInformation
    TxtNome.SetFocus
End If

End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 30
Formulário.Top = 0

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
'DbcPesquisa.SetFocus
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatDados(0).RecordSource = "SELECT HISTEST.Nome, HISTEST.CodPrd, PRD.Descrição_Produto, HISTEST.Valor FROM HISTEST INNER JOIN PRD ON HISTEST.CodPrd = PRD.CodPrd where nome = " & """" & TxtNome & """" & " Order By HISTEST.CodPrd"
DatDados(0).Refresh



DatDados(1).RecordSource = "Select  CodPrd, Descrição_Produto From Prd  Order By CodPrd "
DatDados(1).Refresh

If DatDados(1).Recordset.RecordCount > 0 Then
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.Eof
        LstPrd.AddItem psCompString(1, DatDados(1).Recordset.Fields("CodPrd"), 6) & " - " & DatDados(1).Recordset.Fields("Descrição_Produto")
        LstPrd.Selected(LstPrd.ListCount - 1) = True
        DatDados(1).Recordset.MoveNext
    Wend
    LstPrd.Refresh
End If

End Sub


