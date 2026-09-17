VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesqCli 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa"
   ClientHeight    =   4320
   ClientLeft      =   4050
   ClientTop       =   4020
   ClientWidth     =   7320
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7320
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Cli.CodCli,Cli.Nome_Cliente,Cli.Endereço From Cli Order By Cli.Nome_Cliente"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmbBotao 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox TxtPesquisa 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin MSDBGrid.DBGrid DBGProdutos 
      Bindings        =   "FrmPesqCli.frx":0000
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "FrmPesqCli.frx":0017
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
End
Attribute VB_Name = "FrmPesqCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmbBotao_Click(Index As Integer)
If Index = 0 Then
    FormPedido.TxtDados(0) = DatDados.Recordset.Fields(0)
    Unload Me
Else
    Unload Me
End If
End Sub

Private Sub DBGProdutos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Unload Me
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Activate()
    
DatDados.DatabaseName = Caminho & "\Produtos.mdb"
DatDados.RecordSource = "Select Cli.CodCli,Cli.Nome_Cliente,Cli.Endereço From Cli Order By Cli.Nome_Cliente"
DatDados.Refresh

flag = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

flag = 0
End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" Then
    Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & TxtPesquisa & "*"""
    DatDados.Recordset.FindFirst Criterio
    If DatDados.Recordset.NoMatch Then
'        MsgBox "Informação não Localizada", vbExclamation, App.Title
        Beep
        TxtPesquisa.SetFocus
        SendKeys "{BACKSPACE}"
    Else
        TxtPesquisa.SetFocus
    End If
End If

End Sub

Private Sub TxtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    DBGProdutos.SetFocus
End If
End Sub
