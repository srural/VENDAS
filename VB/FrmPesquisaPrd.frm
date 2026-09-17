VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesquisaPrd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   1005
   ClientTop       =   2925
   ClientWidth     =   10890
   FillColor       =   &H00404040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -540
      ScaleHeight     =   345
      ScaleWidth      =   11415
      TabIndex        =   19
      Top             =   6660
      Width           =   11445
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11055
      TabIndex        =   17
      Top             =   0
      Width           =   11085
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa de Produtos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   18
         Top             =   45
         Width           =   6045
      End
   End
   Begin VB.CommandButton CmbFab 
      Caption         =   "&Cod. Fabricante"
      Height          =   375
      Left            =   5805
      TabIndex        =   16
      Top             =   855
      Width           =   1425
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5130
      Width           =   7365
   End
   Begin VB.CommandButton BtCont 
      Height          =   375
      Left            =   9735
      TabIndex        =   4
      ToolTipText     =   "Clique aqui para limpar o filtro"
      Top             =   855
      Width           =   780
   End
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistemas em Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3690
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FrmPesquisaPrd.frx":0000
      Top             =   2730
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   8655
      TabIndex        =   3
      Top             =   855
      Width           =   975
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Top             =   855
      Width           =   975
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
      Height          =   420
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "Coloque * antes e/ou depois da palavra para filtrar"
      Top             =   855
      Width           =   5580
   End
   Begin MSDBGrid.DBGrid DBGProdutos 
      Bindings        =   "FrmPesquisaPrd.frx":010B
      Height          =   3585
      Left            =   135
      OleObjectBlob   =   "FrmPesquisaPrd.frx":0122
      TabIndex        =   1
      Top             =   1305
      Width           =   10395
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor/Fabricante"
      Height          =   195
      Left            =   7605
      TabIndex        =   15
      Top             =   5895
      Width           =   1635
   End
   Begin VB.Label LblFornec 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7695
      TabIndex        =   14
      Top             =   6105
      Width           =   2760
   End
   Begin VB.Label LblMed 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   13
      Top             =   6120
      Width           =   7350
   End
   Begin VB.Label LblPromocao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9000
      TabIndex        =   12
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Label LblCusto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7605
      TabIndex        =   11
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o Código / Nome / EAN 13 do Produto"
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   660
      Width           =   3210
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo"
      Height          =   165
      Left            =   7605
      TabIndex        =   9
      Top             =   5175
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promoção"
      Height          =   195
      Left            =   8985
      TabIndex        =   8
      Top             =   5205
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observação do Produto"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela Vinculada"
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   5865
      Width           =   1245
   End
End
Attribute VB_Name = "FrmPesquisaPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Tabela As Recordset
Public Ind As Currency
Public VarFornec As Recordset


Private Sub BtCont_Click()
If VarPcoPrazo = True Then
    DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Margem, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True ORDER BY PRD.Descrição_Produto"
Else
    DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Margem, Prd.Custo,Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True ORDER BY PRD.Descrição_Produto"
End If
DatDados.Refresh
TxtPesquisa = ""
BtCont.Caption = "Limpa"

TxtPesquisa.SetFocus

'BtCont.Caption = DatDados.Recordset.RecordCount

End Sub

Private Sub CmbBotao_Click(Index As Integer)
If DatDados.Recordset.RecordCount > 0 Then
    FormProdutos.DbcPesquisa = DatDados.Recordset.Fields(1)
    Unload Me
    FormProdutos.DbcPesquisa.SetFocus
End If
End Sub

Private Sub CmbFab_Click()
Dim Criterio As String
Dim Terminação As String
Terminacao = """"

'If KeyAscii = 13 Then

'    Criterio = DatDados.Recordset.Fields("complemento") & " like " & Terminacao & TxtPesquisa & Terminacao
'    DatDados.Recordset.FindFirst Criterio
'    If DatDados.Recordset.NoMatch Then
'         MsgBox "Informação não Localizada", vbExclamation, App.Title
'         TxtPesquisa = ""
'         TxtPesquisa.SetFocus
'    Else
        If VarPcoPrazo = True Then
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.complemento LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
        Else
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.complemento LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
        End If
        DatDados.Refresh
        BtCont.Caption = DatDados.Recordset.RecordCount
 '   End If

'End If

End Sub

Private Sub DatDados_Reposition()
Dim Criterio As String
On Error Resume Next

If Tabela.RecordCount > 0 Then
    If DatDados.Recordset.RecordCount > 0 Then
        Criterio = "Med_Abc = " & DatDados.Recordset("CodAbc")
        Tabela.FindFirst Criterio
        If Not Tabela.NoMatch Then
            LblMed = Trim(Tabela!MED_DES) & " - " & Tabela!MED_APR & "  -  " & Tabela!lab_Nom
        Else
            LblMed = ""
        End If
    Else
        LblMed = ""
    End If
End If
If DatDados.Recordset.RecordCount > 0 Then
    If Not IsNull(DatDados.Recordset("Obs")) Then
        TxtObs.Text = DatDados.Recordset("Obs")
    End If
    
    If Not IsNull(DatDados.Recordset("Custo")) Then
        LblCusto = Format(DatDados.Recordset("Custo"), "#,##0.00")
    End If
        
    If VarPcoPrazo = True Then
        If DatDados.Recordset("Promocao") = True Then
            If Not IsNull(DatDados.Recordset("Venda_Promocao")) Then
                If IsNumeric(DatDados.Recordset("Venda_Promocao")) Then
                    LblPromocao = Format(DatDados.Recordset("Venda_Promocao"), "#,##0.00")
                Else
                    LblPromocao = ""
                End If
            End If
        Else
            LblPromocao = ""
        End If
    End If
Else
    CmbBotao(0).Default = False
End If

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
'If FrmPedidos.Indices.RecordCount <> 0 Then
'    Ind = FrmPedidos.Indices!Perc
'Else
    Ind = 0
'End If

Set VarFornec = Banco.OpenRecordset("Select CodEntidade, Nome From Ent Order By CodEntidade")

Set Tabela = Banco.OpenRecordset("SELECT TABELA.MED_ABC, TABELA.MED_DES, TABELA.MED_APR, TABELA.LAB_NOM FROM TABELA ORDER BY TABELA.MED_ABC")

DatDados.DatabaseName = Caminho & "\Dados.mdb"
'    DatDados.RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Ativo, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.Classe FROM PRD WHERE Prd.Ativo = True  ORDER BY PRD.Descrição_Produto"

If VarPcoPrazo = True Then
    DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Custo,Prd.Margem, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
Else
    DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Custo,Prd.Margem, Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
End If

DatDados.Refresh

BtCont.Caption = "Limpa"

flag = 1
End Sub

'Private Sub CmdLocalizar_Click(Index As Integer)
'Dim Criterio As String
'Dim Terminação As String
'Terminacao = """"
'If Index = 0 Then
'    Criterio = DatDados.Recordset.Fields(0).Name & " = " & Terminacao & DBCDados.BoundText & Terminacao
'    DatDados.Recordset.FindFirst Criterio
'    If DatDados.Recordset.NoMatch Then
'          MsgBox "Informação não Localizada", vbExclamation, App.Title
'    Else
'          Unload Me
'    End If
'Else
'    Unload Me
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmpedidos.MnuTotal.Item(2).Enabled = False
flag = 0
End Sub


Private Sub TxtFab_keyPress(KeyAscii As Integer)

End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim SQL As String
Dim Terminacao As String
Terminacao = """"
Dim X As Boolean
X = False
'If Datdados.Recordset.RecordCount >= 0 And x = True Then
    If IsNumeric(TxtPesquisa) Then
        Criterio = "CodPrd = " & TxtPesquisa
        DatDados.Recordset.FindFirst Criterio
    Else
        If TxtPesquisa <> "" Then
            Criterio = DatDados.Recordset.Fields(2).Name & " Like " & "*" & """" & TxtPesquisa & "*"""
            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then
                If Len(TxtPesquisa) > 1 Then
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
                    End If
                    DatDados.Refresh
                    BtCont.Caption = DatDados.Recordset.RecordCount
                End If
            Else
                DatDados.Recordset.FindFirst Criterio
                BtCont.Caption = "Limpa"
           End If
            If DatDados.Recordset.NoMatch And DatDados.Recordset.RecordCount > 0 Then
        '        MsgBox "Informação não Localizada", vbExclamation, App.Title
                Beep
                TxtPesquisa.SetFocus
                SendKeys "{BACKSPACE}"
            Else
                TxtPesquisa.SetFocus
            End If
        Else
            If VarPcoPrazo = True Then
                DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            Else
                DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda,Prd.Custo, Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            End If
            DatDados.Refresh
        End If
    End If
'End If
End Sub

Private Sub TxtPesquisa_GotFocus()
Dim VarBarra As Recordset

CmbBotao(0).Default = False

If TxtPesquisa = "" Then
    If VarPcoPrazo = True Then
        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
    Else
        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Custo, Prd.Margem,Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
    End If
    DatDados.Refresh
Else
    If IsNumeric(TxtPesquisa) Then
        If Len(Trim(TxtPesquisa)) > 6 Then
        
            Set VarBarra = Banco.OpenRecordset("select * From Bar Where BarBarra = " & """" & TxtPesquisa & """")
            If VarBarra.RecordCount > 0 Then
                Criterio = "CodPrd = " & VarBarra!Produto
            Else
                Criterio = "CodPrd = " & TxtPesquisa
            End If
'            Criterio = FrmPedidos.DatDados(1).Recordset.Fields("BarBarra").Name & " = "
'            Criterio = Criterio & Terminacao & txtPesquisa & Terminacao
            FrmPedidos.DatDados(1).Recordset.FindFirst Criterio
            If FrmPedidos.DatDados(1).Recordset.NoMatch Then
                MsgBox "Informação não Localizada", vbExclamation, App.Title
                Exit Sub
            Else
                TxtPesquisa = FrmPedidos.DatDados(1).Recordset.Fields("CodPrd").value
            End If
        End If
        Criterio = "CodPrd = " & TxtPesquisa
        DatDados.Recordset.FindFirst Criterio
        If VarPcoPrazo = True Then
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        Else
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        End If
        DatDados.Refresh
        BtCont.Caption = DatDados.Recordset.RecordCount
        If BtCont.Caption = "0" Then
 '           BtCont_Click
        End If
        
    Else
        If TxtPesquisa <> "" Then
'            Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & TxtPesquisa & "*"""
'            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then
                If Len(TxtPesquisa) >= 1 Then
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc ,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
                    End If
                    DatDados.Refresh
                    BtCont.Caption = DatDados.Recordset.RecordCount
                    If BtCont.Caption = "0" Then
'                        BtCont_Click
                    End If
                End If
            Else
'                DatDados.Recordset.FindFirst Criterio
                BtCont.Caption = "Limpa"
            End If
            If DatDados.Recordset.NoMatch Then
        '        MsgBox "Informação não Localizada", vbExclamation, App.Title
                Beep
                TxtPesquisa.SetFocus
                SendKeys "{BACKSPACE}"
            Else
                TxtPesquisa.SetFocus
            End If
'        End If
    End If
'    If frmpedidos.CodCli <> 0 Then
        CmbBotao(0).Default = True
'    End If

End If
End Sub

Private Sub TxtPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Terminacao As String
Dim VarBarra As Recordset
Terminacao = """"

If KeyCode = 13 Then
    If IsNumeric(TxtPesquisa) Then
        If Len(Trim(TxtPesquisa)) >= 6 Then
            Set VarBarra = Banco.OpenRecordset("select * From Bar Where BarBarra = " & """" & TxtPesquisa & """")
            If VarBarra.RecordCount > 0 Then
                Criterio = "CodPrd = " & VarBarra!Produto
            Else
                Criterio = "CodPrd = " & TxtPesquisa
            End If

'            Criterio = Criterio & Terminacao & txtPesquisa & Terminacao
            FrmPedidos.DatDados(1).Recordset.FindFirst Criterio
            If FrmPedidos.DatDados(1).Recordset.NoMatch Then
                MsgBox "Informação não Localizada", vbExclamation, App.Title
                Exit Sub
            Else
                TxtPesquisa = FrmPedidos.DatDados(1).Recordset.Fields("CodPrd").value
            End If
        End If
        Criterio = "CodPrd = " & TxtPesquisa
        DatDados.Recordset.FindFirst Criterio
        If VarPcoPrazo = True Then
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        Else
            DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        End If
        DatDados.Refresh
        BtCont.Caption = DatDados.Recordset.RecordCount
        If BtCont.Caption = "0" Then
 '           BtCont_Click
        End If
        
    Else
        If TxtPesquisa <> "" Then
'            Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & TxtPesquisa & "*"""
'            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then

                If Len(TxtPesquisa) >= 1 Then
                    TxtPesquisa = "*" & TxtPesquisa & "*"

                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc ,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT PRD.complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
                    End If
                    DatDados.Refresh
                    BtCont.Caption = DatDados.Recordset.RecordCount
                    If BtCont.Caption = "0" Then
'                        BtCont_Click
                    End If
                End If
            Else
'                DatDados.Recordset.FindFirst Criterio
                BtCont.Caption = "Limpa"
            End If
            If DatDados.Recordset.NoMatch Then
        '        MsgBox "Informação não Localizada", vbExclamation, App.Title
                Beep
                TxtPesquisa.SetFocus
                SendKeys "{BACKSPACE}"
            Else
                TxtPesquisa.SetFocus
            End If
'        End If
    End If
'    If frmpedidos.CodCli <> 0 Then
        CmbBotao(0).Default = True
'    End If
End If

If KeyCode = 27 Then
    Unload Me
End If


End Sub

Private Sub TxtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    DBGProdutos.SetFocus
End If
End Sub

Private Sub TxtPesquisa_LostFocus()
'If frmpedidos.CodCli <> 0 Then
    CmbBotao(0).Default = True
'End If
End Sub
