VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesquisaCest 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de CEST"
   ClientHeight    =   5295
   ClientLeft      =   720
   ClientTop       =   2355
   ClientWidth     =   11685
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   11685
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Minutos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT CEST.Descricao, CEST.Segmento, CEST.CEST, CEST.[NCM/SH] FROM CEST ORDER BY CEST.Descricao"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   10305
      TabIndex        =   3
      Top             =   135
      Width           =   1155
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   9045
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
      Bindings        =   "FrmPesquisaCest.frx":0000
      Height          =   4515
      Left            =   180
      OleObjectBlob   =   "FrmPesquisaCest.frx":0017
      TabIndex        =   1
      Top             =   585
      Width           =   11205
   End
End
Attribute VB_Name = "FrmPesquisaCest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbBotao_Click(Index As Integer)
If Index = 0 Then
'    FormProdutos.TxtDados(10) = DatDados.Recordset.Fields(0)
'    If Datdados.Recordset.Fields("NCM") <> "" Then
        FormProdutos.TxtDados(12) = DatDados.Recordset.Fields("CEST")
'    Else
'        FormProdutos.txtDados(7) = Left(Datdados.Recordset.Fields("Descricao"), 8)
'    End If

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
    TxtPesquisa = ""
    TxtPesquisa.SetFocus
End If

End Sub

Private Sub Form_Load()

DatDados.DatabaseName = Caminho & "\Minutos.mdb"
DatDados.RecordSource = "SELECT CEST.Descricao, CEST.Segmento, CEST.CEST, CEST.[NCM/SH] FROM CEST ORDER BY CEST.Descricao"

DatDados.Refresh

End Sub
Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" Then
'    Datdados.RecordSource = "SELECT CadNcm.Descricao, CadNcm.Categoria, CadNcm.NCM, CadNcm.Inicio, CadNcm.Fim FROM  CadNcm WHERE  Descricao like " & """" & "*" & TxtPesquisa & "*" & """" & "ORDER BY CadNcm.Descricao"
    DatDados.RecordSource = "SELECT CEST.Descricao, CEST.Segmento, CEST.CEST, CEST.[NCM/SH] FROM CEST  WHERE  Descricao like " & """" & "*" & TxtPesquisa & "*" & """" & "ORDER BY CEST.Descricao"
    DatDados.Refresh

'    Criterio = Datdados.Recordset.Fields(0).Name & " Like " & """" & "*" & TxtPesquisa & "*" & "*"""
'    Datdados.Recordset.FindFirst Criterio
'    If Datdados.Recordset.NoMatch Then
'        Beep
'        TxtPesquisa.SetFocus
'        SendKeys "{BACKSPACE}"
'    Else
'        TxtPesquisa.SetFocus
'    End If

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

