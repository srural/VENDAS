VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPesquisa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8145
   ClientLeft      =   1005
   ClientTop       =   2925
   ClientWidth     =   15405
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   15405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   9360
      TabIndex        =   33
      Top             =   5850
      Width           =   3615
      Begin VB.Image ImgFoto 
         Height          =   2085
         Left            =   45
         Stretch         =   -1  'True
         Top             =   135
         Width           =   3480
      End
   End
   Begin VB.CheckBox ChkPesq 
      Caption         =   "Qualquer parte do nome"
      Height          =   285
      Left            =   5130
      TabIndex        =   29
      Top             =   630
      Width           =   2355
   End
   Begin VB.TextBox TxtComplemento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5625
      TabIndex        =   28
      ToolTipText     =   "Coloque * antes e/ou depois da palavra para filtrar"
      Top             =   945
      Width           =   2910
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15450
      TabIndex        =   26
      Top             =   0
      Width           =   15450
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa Produtos"
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
         Left            =   135
         TabIndex        =   27
         Top             =   45
         Width           =   5325
      End
   End
   Begin VB.Frame FrmPreco 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   1830
      TabIndex        =   13
      Top             =   2970
      Visible         =   0   'False
      Width           =   10320
      Begin MSMask.MaskEdBox MebCusto 
         Height          =   420
         Left            =   360
         TabIndex        =   17
         Top             =   765
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPerc 
         Height          =   420
         Left            =   2190
         TabIndex        =   18
         Top             =   765
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebVenda 
         Height          =   420
         Left            =   4020
         TabIndex        =   19
         Top             =   765
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDesc 
         Height          =   420
         Left            =   5850
         TabIndex        =   20
         Top             =   765
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   7785
         TabIndex        =   22
         Top             =   630
         Width           =   2205
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto (%)"
         Height          =   195
         Left            =   5850
         TabIndex        =   21
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda"
         Height          =   195
         Left            =   4005
         TabIndex        =   16
         Top             =   540
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margem (%)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2205
         TabIndex        =   15
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   540
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   4830
      ScaleHeight     =   1140
      ScaleWidth      =   4290
      TabIndex        =   10
      Top             =   6120
      Width           =   4320
      Begin VB.TextBox TxtObsPrd 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   90
         Width           =   4065
      End
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
      Height          =   1170
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   6120
      Width           =   4260
   End
   Begin VB.CommandButton BtCont 
      Height          =   375
      Left            =   11385
      TabIndex        =   4
      ToolTipText     =   "Clique aqui para limpar o filtro"
      Top             =   945
      Width           =   1005
   End
   Begin VB.Data DatDados 
      Caption         =   "Pesquisa"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FrmPesquisa.frx":0000
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   10035
      TabIndex        =   3
      Top             =   930
      Width           =   1200
   End
   Begin VB.CommandButton CmbBotao 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   8700
      TabIndex        =   2
      Top             =   930
      Width           =   1155
   End
   Begin VB.TextBox TxtPesquisa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   270
      TabIndex        =   0
      ToolTipText     =   "Coloque * antes e/ou depois da palavra para filtrar"
      Top             =   945
      Width           =   5295
   End
   Begin MSDBGrid.DBGrid DBGProdutos 
      Bindings        =   "FrmPesquisa.frx":010B
      Height          =   4395
      Left            =   360
      OleObjectBlob   =   "FrmPesquisa.frx":0122
      TabIndex        =   1
      Top             =   1440
      Width           =   14865
   End
   Begin MSMask.MaskEdBox MskData 
      Height          =   330
      Left            =   13455
      TabIndex        =   32
      Top             =   945
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atacado"
      Height          =   195
      Left            =   13275
      TabIndex        =   31
      Top             =   7110
      Width           =   600
   End
   Begin VB.Label LblAtacado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   13290
      TabIndex        =   30
      Top             =   7335
      Width           =   1950
   End
   Begin VB.Shape SHEst 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12510
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LblMed 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   25
      Top             =   7695
      Width           =   9555
   End
   Begin VB.Label LblPromocao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   13275
      TabIndex        =   24
      Top             =   6660
      Width           =   1950
   End
   Begin VB.Label LblCusto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   13725
      TabIndex        =   23
      Top             =   5895
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o Código / Nome / EAN 13 do Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   12
      Top             =   660
      Width           =   4605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo"
      Height          =   195
      Left            =   13185
      TabIndex        =   9
      Top             =   5940
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Promoção/Gelado"
      Height          =   195
      Left            =   13260
      TabIndex        =   8
      Top             =   6435
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observação do Produto"
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   5865
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela Vinculada"
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   7485
      Width           =   1245
   End
End
Attribute VB_Name = "FrmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
            
Public Tabela As Recordset
Public Ind As Currency
Public VarFlag As Boolean


Private Sub BtCont_Click()
If VarPcoPrazo = True Then
    DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia, Prd.CodBar,PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Margem, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True ORDER BY PRD.Descrição_Produto"
Else
    DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia, Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Margem, Prd.Custo,Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True ORDER BY PRD.Descrição_Produto"
End If
DatDados.Refresh
TxtPesquisa = ""
BtCont.Caption = "Limpa"

TxtPesquisa.SetFocus

'BtCont.Caption = DatDados.Recordset.RecordCount

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbBotao_Click(Index As Integer)
'If Index = 0 Then
    If FrmFrente.CodCli <> 0 Then
        If DatDados.Recordset.RecordCount > 0 Then
            FrmFrente.TxtDados(2) = DatDados.Recordset.Fields("CodPrd")
            Unload Me
            If VarComplemento = "Sim" Then
                FrmFrente.txtcomplemento.SetFocus
            Else
                FrmFrente.TxtDados(0).SetFocus
            End If
        End If
    Else
        If DatDados.Recordset.RecordCount > 0 And VarMercado = 0 And Tabela.RecordCount > 0 Then
            FrmPreco.Visible = True
            LblPreco.Caption = DatDados.Recordset.Fields("Venda")
            FrmPreco.Caption = DatDados.Recordset.Fields("Descrição_Produto")
            MebCusto = DatDados.Recordset.Fields("Custo")
            MebPerc = DatDados.Recordset.Fields("Margem")
            MebVenda = DatDados.Recordset.Fields("Venda")
            MebDesc_Change
            MebDesc.SetFocus
        Else
            If TxtPesquisa <> "" Then
                TxtPesquisa_KeyPress 13
            End If
        End If
    End If

'End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DatDados_Reposition()
Dim Criterio As String
If Tabela.RecordCount > 0 Then
    If DatDados.Recordset.RecordCount > 0 Then
        If Not IsNull(DatDados.Recordset("Codabc")) Then
            Criterio = "Med_Barra = " & """" & DatDados.Recordset("CodBar") & """"
            Tabela.FindFirst Criterio
            If Not Tabela.NoMatch Then
                LblMed = Trim(Tabela!MED_DES) & " - " & Tabela!MED_APR & "  -  " & Tabela!lab_Nom
            Else
                LblMed = ""
            End If
        Else
            LblMed = ""
        End If
    Else
        LblMed = ""
    End If
End If
If DatDados.Recordset.RecordCount > 0 Then

    If DatDados.Recordset("Estoque") <= 0 Then
        SHEst.BackColor = &HFF&
    Else
        SHEst.BackColor = &HFF0000
    End If
    
    If Not IsNull(DatDados.Recordset("DTAlt")) Then
         MskData.Text = DatDados.Recordset("DTAlt")
    Else
        MskData.Text = ""
    End If
    

    If Not IsNull(DatDados.Recordset("Obs")) Then
        TxtObs.Text = DatDados.Recordset("Obs")
    Else
        TxtObsPrd = ""
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
    If LblPromocao = "" Then
        If Not IsNull(DatDados.Recordset("Venda_Concorrencia")) Then
            If IsNumeric(DatDados.Recordset("Venda_Concorrencia")) Then
                LblPromocao = Format(DatDados.Recordset("Venda_Concorrencia"), "#,##0.00")
            Else
                LblPromocao = ""
            End If
        End If
    
    End If
    
    If Not IsNull(DatDados.Recordset("Margem_Concorrencia")) Then
        If IsNumeric(DatDados.Recordset("Margem_Concorrencia")) Then
            LblAtacado = Format(DatDados.Recordset("Margem_Concorrencia"), "#,##0.00")
        Else
            LblAtacado = ""
        End If
    End If
    
'    If Registros.Fields("Foto") Then
        If dir(FrmFrente.Caminho & "\fotos\" & DatDados.Recordset("CodPrd") & ".jpg") <> "" Then
            ImgFoto.Visible = True
'            LblFoto.Visible = False
    '        ImgFoto.Picture = LoadPicture(App.Path & "\fotos\" & Trim(CStr(Registros.Fields("CodPrd"))) & ".jpg")
            ImgFoto.Picture = LoadPicture(FrmFrente.Caminho & "\fotos\" & DatDados.Recordset("CodPrd") & ".jpg")
        Else
            ImgFoto.Visible = False
'            LblFoto.Visible = True
        End If
'    Else
'        ImgFoto.Visible = False
'        LblFoto.Visible = True
'    End If
         
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

Private Sub Form_Activate()
VarFlag = True
End Sub

Private Sub Form_Load()
If FrmFrente.Indices.RecordCount <> 0 Then
    Ind = FrmFrente.Indices!Perc
Else
    Ind = 0
End If

Set Tabela = FrmFrente.Banco.OpenRecordset("SELECT TABELA.MED_ABC, TABELA.MED_DES, TABELA.MED_APR, TABELA.MED_BARRA, TABELA.LAB_NOM FROM TABELA ORDER BY TABELA.MED_ABC")

DatDados.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
'    DatDados.RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Ativo, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.Classe FROM PRD WHERE Prd.Ativo = True  ORDER BY PRD.Descrição_Produto"

If VarPcoPrazo = True Then
    DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, Prd.Custo,Prd.Margem, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
Else
    DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia, Prd.CodBar,PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Custo,Prd.Margem, Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
End If

DatDados.Refresh

If dir(LocalBanco & "\ObsPrd.Txt") <> "" Then
    Dim NomeObs
    TxtObsPrd = ""
    Open LocalBanco & "\ObsPrd.Txt" For Input As #1
    Input #1, NomeObs
    TxtObsPrd.Text = NomeObs
    Close #1
End If

BtCont.Caption = "Limpa"

If VarMostraCusto = "Nao" Then
    Label4.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    LblCusto.Visible = False
    MebCusto.Visible = False
    MebPerc.Visible = False
End If

If VarMostraEstoque = False Then
    DBGProdutos.Columns(6).Visible = False
    SHEst.Visible = False
End If

Flag = 1

If LerINI("Pedido", "PesqQualquerCampo", LocalBanco & "\config.ini") <> "" Then
    ChkPesq = LerINI("Pedido", "PesqQualquerCampo", LocalBanco & "\config.ini")
Else
    ChkPesq = 0
End If


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
'FrmFrente.MnuTotal.Item(2).Enabled = False
Flag = 0
End Sub


Private Sub MebDesc_Change()
If MebDesc <> "" Then
    If IsNumeric(MebDesc) Then
        LblPreco = Format(MebVenda - (MebVenda * (MebDesc / 100)), "#,##0.00")
    End If
End If
End Sub

Private Sub MebDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    FrmPreco.Visible = False
End If

End Sub

Private Sub MebDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FrmPreco.Visible = False
End If
End Sub

Private Sub TxtComplemento_Change()
Dim Criterio As String
Dim sql As String
Dim Terminacao As String
Terminacao = """"
If DatDados.Recordset.RecordCount >= 0 Then
        If txtcomplemento <> "" Then
            Criterio = DatDados.Recordset.Fields(0).Name & " Like " & """" & txtcomplemento & "*"""
            If Left(txtcomplemento, 1) = "*" Or Right(txtcomplemento, 1) = "*" Then
                If Len(TxtPesquisa) > 1 Then
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Complemento LIKE  " & """" & txtcomplemento & """" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Complemento LIKE  " & """" & txtcomplemento & """" & " ORDER BY PRD.Descrição_Produto"
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
                txtcomplemento.SetFocus
                SendKeys "{BACKSPACE}"
            Else
                txtcomplemento.SetFocus
            End If
        Else
            If VarPcoPrazo = True Then
                DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            Else
                DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda,Prd.Custo, Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            End If
            DatDados.Refresh
        End If
End If
End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim sql As String
Dim Terminacao As String
Terminacao = """"
 On Error Resume Next
If DatDados.Recordset.RecordCount >= 0 And TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        Criterio = "CodPrd = " & Trim(TxtPesquisa)
        DatDados.Recordset.FindFirst Criterio
    Else
        If TxtPesquisa <> "" Then
            If ChkPesq.Value = 1 Then
                Criterio = DatDados.Recordset.Fields(2).Name & " Like " & """" & TxtPesquisa & "*"""
            Else
                Criterio = DatDados.Recordset.Fields(2).Name & " Like " & """" & TxtPesquisa & "*"""
            End If
            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then
                If Len(TxtPesquisa) > 1 Then
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & """" & " ORDER BY PRD.Descrição_Produto"
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
'                Beep
'                TxtPesquisa.SetFocus
'                SendKeys "{BACKSPACE}"
            Else
'                TxtPesquisa.SetFocus
            End If
        Else
            If VarPcoPrazo = True Then
                DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            Else
                DatDados.RecordSource = "SELECT Venda_Concorrencia,  Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, Prd.Margem,PRD.Venda,Prd.Custo, Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
            End If
            DatDados.Refresh
        End If
    End If
End If
End Sub

Private Sub TxtPesquisa_GotFocus()

CmbBotao(0).Default = False
If TxtPesquisa = "" Then
    If VarPcoPrazo = True Then
        DatDados.RecordSource = "SELECT Venda_Concorrencia,  Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Ativo, Prd.Obs, Prd.PcoPrazo as expr1, PRD.Estoque , Prd.CodAbc, Prd.Promocao,  Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
    Else
        DatDados.RecordSource = "SELECT Venda_Concorrencia,  Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Custo, Prd.Margem,Prd.Ativo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND Prd.CodPrd = 0 ORDER BY PRD.Descrição_Produto"
    End If
    DatDados.Refresh
End If
End Sub

Private Sub TxtPesquisa_KeyPress(KeyAscii As Integer)
Dim Terminacao As String
Terminacao = """"
 On Error Resume Next

If KeyAscii = 13 And TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        If Len(Trim(TxtPesquisa)) >= 6 Then
            Criterio = FrmFrente.Registros.Fields("BarBarra").Name & " = "
            Criterio = Criterio & Terminacao & (TxtPesquisa) & Terminacao
            FrmFrente.Registros.FindFirst Criterio
            If FrmFrente.Registros.NoMatch Then
                MsgBox "Informação não Localizada", vbExclamation, App.Title
                Exit Sub
            Else
                TxtPesquisa = FrmFrente.Registros.Fields("CodPrd").Value
            End If
        End If
        Criterio = "CodPrd = " & TxtPesquisa
        DatDados.Recordset.FindFirst Criterio
        If VarPcoPrazo = True Then
            DatDados.RecordSource = "SELECT  Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1 , PRD.Estoque , Prd.CodAbc, Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        Else
            DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda,Prd.Margem, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.CodPrd =  " & TxtPesquisa & " ORDER BY PRD.Descrição_Produto"
        End If
        DatDados.Refresh
        BtCont.Caption = DatDados.Recordset.RecordCount
        If BtCont.Caption = "0" Then
'            BtCont_Click
        End If
        TxtPesquisa = ""
    Else
        If TxtPesquisa <> "" Then
'            Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & TxtPesquisa & "*"""
'            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then
                If Len(TxtPesquisa) >= 1 Then
                    If ChkPesq.Value = 1 Then
                        TxtPesquisa = "*" & TxtPesquisa & "*"
                    End If
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc ,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT Venda_Concorrencia, Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.Descrição_Produto LIKE  " & """" & TxtPesquisa & "*""" & " ORDER BY PRD.Descrição_Produto"
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
            TxtPesquisa = ""
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
'    If FrmFrente.CodCli <> 0 Then
        CmbBotao(0).Default = True
'    End If
Else
    TxtPesquisa.SetFocus
End If

If KeyAscii = 27 Then
    Unload Me
End If

End Sub

Private Sub TxtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    DBGProdutos.SetFocus
End If
End Sub

Private Sub TxtPesquisa_LostFocus()
'If FrmFrente.CodCli <> 0 Then
    CmbBotao(0).Default = True
'End If
End Sub

Private Sub TxtComplemento_KeyPress(KeyAscii As Integer)
Dim Terminacao As String
Terminacao = """"

If KeyAscii = 13 Then
        If txtcomplemento <> "" Then
'            Criterio = DatDados.Recordset.Fields(1).Name & " Like " & """" & TxtPesquisa & "*"""
'            If Left(TxtPesquisa, 1) = "*" Or Right(TxtPesquisa, 1) = "*" Then
                If Len(txtcomplemento) >= 1 Then
                    If VarPcoPrazo = True Then
                        DatDados.RecordSource = "SELECT Venda_Concorrencia,  Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, Prd.PcoPrazo as Expr1, PRD.Estoque , Prd.CodAbc ,Prd.Promocao, Prd.Venda_Promocao FROM PRD WHERE Prd.Ativo = True AND PRD.Complemento LIKE  " & """" & txtcomplemento & "*""" & " ORDER BY PRD.Descrição_Produto"
                    Else
                        DatDados.RecordSource = "SELECT Venda_Concorrencia,  Prd.DtAlt, Prd.Margem_Concorrencia,  Prd.CodBar, PRD.Complemento, PRD.CodPrd, PRD.Descrição_Produto,Prd.Margem, PRD.Venda, Prd.Custo, Prd.Obs, [prd].[venda]+([prd].[venda]*(" & Replace(Ind, ",", ".") & "/100)) AS Expr1, PRD.Estoque , Prd.CodAbc FROM PRD WHERE Prd.Ativo = True AND PRD.complemento LIKE  " & """" & txtcomplemento & "*""" & " ORDER BY PRD.Descrição_Produto"
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
                txtcomplemento.SetFocus
            End If
'        End If
'    If FrmFrente.CodCli <> 0 Then
        CmbBotao(0).Default = True
'    End If
End If

If KeyAscii = 27 Then
    Unload Me
End If

End Sub


