VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesquisaTab 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de Produtos"
   ClientHeight    =   5250
   ClientLeft      =   720
   ClientTop       =   2355
   ClientWidth     =   10530
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10530
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT TABELA.MED_ABC, TABELA.MED_DES, TABELA.MED_APR, TABELA.LAB_NOM FROM TABELA ORDER BY TABELA.MED_DES"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   9135
      TabIndex        =   3
      Top             =   135
      Width           =   1155
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7875
      TabIndex        =   2
      Top             =   135
      Width           =   1155
   End
   Begin VB.TextBox TxtPesquisa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   6105
   End
   Begin MSDBGrid.DBGrid DBGProdutos 
      Bindings        =   "FrmPesquisaTab.frx":0000
      Height          =   4515
      Left            =   180
      OleObjectBlob   =   "FrmPesquisaTab.frx":0017
      TabIndex        =   1
      Top             =   585
      Width           =   10170
   End
End
Attribute VB_Name = "FrmPesquisaTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbBotao_Click(Index As Integer)
If Index = 0 Then
    FormProdutos.TxtDados(10) = DatDados.Recordset.Fields(0)
End If

Unload Me
End Sub

Private Sub DBGProdutos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Unload Me
End If
End Sub

Private Sub DBGProdutos_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    txtPesquisa = ""
    txtPesquisa.SetFocus
End If

End Sub

Private Sub Form_Load()

DatDados.DatabaseName = Caminho & "\Dados.mdb"
DatDados.RecordSource = "SELECT TABELA.MED_ABC, TABELA.MED_DES, TABELA.MED_APR, TABELA.LAB_NOM FROM TABELA ORDER BY TABELA.MED_DES"

DatDados.Refresh

End Sub
Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If txtPesquisa <> "" Then
    Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & txtPesquisa & "*"""
    DatDados.Recordset.FindFirst Criterio
    If DatDados.Recordset.NoMatch Then
'        MsgBox "Informação não Localizada", vbExclamation, App.Title
        Beep
        txtPesquisa.SetFocus
        SendKeys "{BACKSPACE}"
    Else
        txtPesquisa.SetFocus
    End If
End If

End Sub

Private Sub TxtPesquisa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub TxtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    DBGProdutos.SetFocus
End If
End Sub

