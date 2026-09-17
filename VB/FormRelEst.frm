VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelEst 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9900
   Begin VB.CheckBox ChkQtd 
      Caption         =   "Mostrar Quantidade"
      Height          =   345
      Left            =   570
      TabIndex        =   28
      Top             =   4446
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.TextBox TxtLocal 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1575
      TabIndex        =   26
      Top             =   6120
      Width           =   2400
   End
   Begin VB.TextBox TxtMarca 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1575
      TabIndex        =   24
      Top             =   5670
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   23
      Top             =   6525
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   21
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Estoque"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   90
         TabIndex        =   22
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4635
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4005
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CheckBox ChkLinha 
      Caption         =   "Mostrar Linha entre produtos"
      Height          =   345
      Left            =   570
      TabIndex        =   18
      Top             =   4725
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox ChkCusto 
      Caption         =   "Mostrar Custo do Produto"
      Height          =   345
      Left            =   570
      TabIndex        =   17
      Top             =   4164
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox ChkPerc 
      Caption         =   "Mostrar Percentual Lucro (%)"
      Height          =   345
      Left            =   540
      TabIndex        =   16
      Top             =   3870
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox ChkZero 
      Caption         =   "Listar Produtos Estoque Zero"
      Height          =   345
      Left            =   570
      TabIndex        =   14
      Top             =   3600
      Width           =   2505
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   5850
      TabIndex        =   13
      Top             =   5430
      Width           =   1590
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5430
      Width           =   1590
   End
   Begin VB.TextBox TxtLoja 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1575
      TabIndex        =   2
      Top             =   5235
      Width           =   2400
   End
   Begin VB.Frame Frm1 
      Caption         =   "&Tipo"
      Height          =   1635
      Left            =   540
      TabIndex        =   0
      Top             =   660
      Width           =   2520
      Begin VB.OptionButton Opt1 
         Caption         =   "&Fornecedor Todos"
         Height          =   375
         Index           =   4
         Left            =   195
         TabIndex        =   12
         Top             =   915
         Width           =   1815
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Forne&cedor Específico"
         Height          =   375
         Index           =   5
         Left            =   195
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Grupo Todos"
         Height          =   375
         Index           =   2
         Left            =   195
         TabIndex        =   8
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Grupo &Específico"
         Height          =   375
         Index           =   3
         Left            =   195
         TabIndex        =   7
         Top             =   510
         Width           =   1635
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   2460
         Y1              =   885
         Y2              =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Ordem"
      Height          =   1125
      Left            =   540
      TabIndex        =   3
      Top             =   2370
      Width           =   2520
      Begin VB.OptionButton Opt2 
         Caption         =   "&Tamanho"
         Height          =   375
         Index           =   2
         Left            =   195
         TabIndex        =   15
         Tag             =   "Tamanho"
         Top             =   720
         Width           =   1140
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "&Alfabética"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Tag             =   "Descrição_Produto"
         Top             =   195
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "&Numérica"
         Height          =   375
         Index           =   1
         Left            =   195
         TabIndex        =   4
         Tag             =   "CodPrd"
         Top             =   465
         Width           =   1095
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Grupos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4860
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1575
      Visible         =   0   'False
      Width           =   2025
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "FormRelEst.frx":0000
      Height          =   2910
      Left            =   3240
      TabIndex        =   1
      Top             =   660
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   5106
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      Style           =   1
      ListField       =   "Descrição_Grupo"
      BoundColumn     =   "CODGru"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "FormRelEst.frx":001A
      Height          =   315
      Left            =   3285
      TabIndex        =   19
      Top             =   4005
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Fisico"
      Height          =   315
      Left            =   540
      TabIndex        =   27
      Top             =   6150
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      Height          =   315
      Left            =   900
      TabIndex        =   25
      Top             =   5700
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   3285
      TabIndex        =   20
      Top             =   3825
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loja/Estoque"
      Height          =   315
      Left            =   540
      TabIndex        =   9
      Top             =   5265
      Width           =   960
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   3270
      TabIndex        =   6
      Top             =   645
      Width           =   3420
   End
End
Attribute VB_Name = "FormRelEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub Form_Activate()

DatDados(0).RecordSource = "SELECT Gru.CodGru, Gru.Descrição_Grupo   FROM Gru ORDER BY Gru.Descrição_Grupo"
DatDados(0).Refresh

'DatDados(1).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"

DBCombo2.Enabled = False
Label1.Enabled = False
End Sub

Private Sub BtImprimir_Click()
Dim x As Variant
Dim Op1 As Integer
Dim Op2 As String
Dim Formula As String
Dim CamposSql As String
Dim Ordenado As String
Dim SQL As String

Formula = ""
CamposSql = ""
Ordenado = ""
SQL = ""

'MDIPrincipal.ComRel.ShowPrinter
                        
For Each x In Opt1
    If x.value Then
        Op1 = x.Index
    End If
Next

For Each x In Opt2
    If x.value Then
        Op2 = x.Tag
    End If
Next

If Op2 = "CodPrd" And Op1 <= 3 Then Op2 = "CodGru, CodPrd"
If Op2 = "Descrição_Produto" And Op1 <= 3 Then Op2 = "CodGru, Descrição_Produto"

If Op2 = "CodPrd" And Op1 >= 4 Then Op2 = "CodEntidade, Descrição_Grupo , CodPrd"
If Op2 = "Descrição_Produto" And Op1 >= 4 Then Op2 = "CodEntidade, Descrição_Grupo, Descrição_Produto"

If Op2 = "Tamanho" And Op1 <= 3 Then Op2 = "CodGru, CodBar"
If Op2 = "Tamanho" And Op1 >= 4 Then Op2 = "CodEntidade, Descrição_Grupo, CodBar"


If Op1 = 2 Or Op1 = 3 Then
    
'    CamposSql = "SELECT GRU.CodGru, GRU.Descrição_Grupo, Prd.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, Prd.Classe, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] AS SubTotal FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto GROUP BY GRU.CodGru, GRU.Descrição_Grupo , Prd.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, Prd.Classe, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] "
'    If VarLojaUnica = False Then
        CamposSql = "SELECT Prd.Local, Prd.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] AS SubTotal, [PRd].[Venda]*[Loj].[Estoque] AS SubTotalVda, PRD.Fabricante FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto GROUP BY Prd.Local, Prd.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque], [PRd].[Venda]*[Loj].[Estoque] , PRD.Fabricante "
'    Else
'        CamposSql = "SELECT PRD.Local, PRD.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, PRD.Venda, PRD.Margem, PRD.Fabricante, PRD.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo GROUP BY PRD.Local, PRD.Marca, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Classe, PRD.Custo, PRD.Venda, PRD.Margem, PRD.Fabricante, PRD.Estoque"
'    End If
    
    FrmRelEst.rpCodigo.Campo = "CodGru"
    FrmRelEst.rpNome.Campo = "Descrição_Grupo"
    FrmRelEst.rpCabTitulo.Caption = "Relatório de Estoques / Por Grupos"
    
    If Op1 = 3 And IsNumeric(DBCombo1.BoundText) Then
        If DBCombo2.Text = "" Then
            Formula = " Having Prd.Ativo = True And CodGru =  " & DBCombo1.BoundText & " "
        Else
            Formula = " Having Prd.Ativo = True And CodGru =  " & DBCombo1.BoundText & " And Prd.Fabricante =  " & DBCombo2.BoundText & " "
        End If
    Else
        Formula = " Having Prd.Ativo = True "
    End If
    
    FrmRelEst.CabGrupo(1).Mostrar = False
    
Else
'        CamposSql = "SELECT Prd.CodBar, Prd.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] AS SubTotal, ENT.CodEntidade, ENT.Nome FROM ENT INNER JOIN (PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto) ON ENT.CodEntidade = PRD.Fabricante GROUP BY Prd.CodBar, Prd.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque], ENT.CodEntidade, ENT.Nome "
'    If VarLojaUnica = False Then
        CamposSql = "SELECT Prd.Local, Prd.Marca, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] AS SubTotal, [PRd].[Venda]*[Loj].[Estoque] AS SubTotalVda, ENT.CodEntidade, ENT.Nome, GRU.Descrição_Grupo FROM GRU INNER JOIN (ENT INNER JOIN (PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto) ON ENT.CodEntidade = PRD.Fabricante) ON GRU.CodGru = PRD.Grupo GROUP BY Prd.Local, Prd.Marca, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque], [PRd].[Venda]*[Loj].[Estoque], ENT.CodEntidade, ENT.Nome, GRU.Descrição_Grupo "
'    CamposSql = "SELECT PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque] AS SubTotal, ENT.CodEntidade, ENT.Nome, First(GRU.Descrição_Grupo) AS PrimeiroDeDescrição_Grupo FROM GRU INNER JOIN (ENT INNER JOIN (PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto) ON ENT.CodEntidade = PRD.Fabricante) ON GRU.CodGru = PRD.Grupo GROUP BY PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, LOJ.DescricaoLoja, LOJ.Estoque, [PRd].[Custo]*[Loj].[Estoque], ENT.CodEntidade, ENT.Nome "
'    Else
'        CamposSql = " SELECT Prd.Estoque, PRD.Local, PRD.Marca, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, ENT.CodEntidade,ENT.Nome, GRU.Descrição_Grupo FROM GRU INNER JOIN (ENT INNER JOIN PRD ON ENT.CodEntidade = PRD.Fabricante) ON GRU.CodGru = PRD.Grupo GROUP BY Prd.Estoque, PRD.Local, PRD.Marca, PRD.CodBar, PRD.Ativo, PRD.Codigo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Margem, ENT.CodEntidade,ENT.Nome, GRU.Descrição_Grupo "
'    End If


    FrmRelEst.rpCodigo.Campo = "CodEntidade"
    FrmRelEst.rpNome.Campo = "Nome"
    FrmRelEst.rpCabTitulo.Caption = "Relatório de Estoques / Por Fornecedores"
    
    If Op1 = 5 And IsNumeric(DBCombo1.BoundText) Then
        Formula = " Having Prd.Ativo = True And CodEntidade =  " & DBCombo1.BoundText & " "
    End If
    FrmRelEst.CabGrupo(1).Mostrar = True

End If

If TxtLoja <> "" Then
    If Formula <> "" Then
        Formula = Formula & " and DescricaoLoja Like '" & TxtLoja & "'"
    Else
        Formula = " Having DescricaoLoja Like '" & TxtLoja & "'"
    End If
End If

If TxtMarca <> "" Then
    If Formula <> "" Then
        Formula = Formula & " and Marca Like '" & TxtMarca & "'"
    Else
        Formula = " Having Marca Like '" & TxtMarca & "'"
    End If
End If

If TxtLocal <> "" Then
    If Formula <> "" Then
        Formula = Formula & " and Local Like '" & TxtLocal & "'"
    Else
        Formula = " Having Local Like '" & TxtLocal & "'"
    End If
End If


If ChkZero.value = 0 Then
    If Formula <> "" Then
        Formula = Formula & " and Loj.Estoque > 0 "
    Else
        Formula = " Having Loj.Estoque <> 0 "
    End If
Else
    If MsgBox("Só estoque menor que zero ?", vbYesNo) = vbYes Then
        If Formula <> "" Then
            Formula = Formula & " and Loj.Estoque < 0 "
        Else
            Formula = " Having Loj.Estoque < 0 "
        End If
    End If

End If
    
    Ordenado = "ORDER BY " & Op2
    
    SQL = CamposSql & Formula & Ordenado
    
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelEst.Relatorio.Recordset = varPesquisa(1)
    FrmRelEst.rpCabSubTitulo.Caption = NomeEmpresa
    If ChkPerc.value = 1 Then
        FrmRelEst.ReportField6.Mostrar = True
        FrmRelEst.Label9.Visible = True
        FrmRelEst.Label10.Visible = True
    Else
        FrmRelEst.ReportField6.Mostrar = False
        FrmRelEst.Label9.Visible = False
'        FrmRelEst.Label10.Visible = False
    End If
    
    If ChkCusto.value = 1 Then
        FrmRelEst.ReportField8.Mostrar = True
        FrmRelEst.Label29.Visible = True
        FrmRelEst.ReportField3.Mostrar = True
        FrmRelEst.Label1.Visible = True
        
    Else
        FrmRelEst.ReportField8.Mostrar = False
        FrmRelEst.Label29.Visible = False
        FrmRelEst.ReportField3.Mostrar = False
        FrmRelEst.Label1.Visible = False
    
    End If
    
'    If ChkCusto.value = 1 Then
'        FrmRelEst.ReportField2.Mostrar = True
'        FrmRelEst.Label2.Visible = True
'    Else
'        FrmRelEst.ReportField2.Mostrar = False
'        FrmRelEst.Label2.Visible = False
'    End If
    
    
    
    If ChkLinha.value = 1 Then
        FrmRelEst.Line9.Visible = True
    Else
        FrmRelEst.Line9.Visible = False
    End If
    
    FrmRelEst.Config
    
End Sub

Private Sub Imprimir_Click()

End Sub

Private Sub Opt1_Click(Index As Integer)

DBCombo2.Enabled = False
Label1.Enabled = False


Select Case Index
    Case 2
        DBCombo1.Enabled = False
        'DBCombo1.Enabled = False
    Case 3
        'DBCombo1.Enabled = False
        DBCombo1.Enabled = True
        DBCombo1.SetFocus
        DBCombo2.Enabled = True
        Label1.Enabled = True
        
        DatDados(1).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
        DatDados(1).Refresh
                
    Case 4
        DBCombo1.Enabled = False
        'DBCombo1.Enabled = False
    Case 5
        DBCombo1.Enabled = True
        'DBCombo1.Enabled = True
        'DBCombo1.SetFocus
End Select
    
Select Case Index
    Case 2, 3
        DatDados(0).RecordSource = "SELECT Gru.CodGru, Gru.Descrição_Grupo   FROM Gru ORDER BY Gru.Descrição_Grupo"
        DatDados(0).Refresh
        lblTipo.Caption = "Grupos"
        DBCombo1.BoundColumn = "CodGru"
        DBCombo1.ListField = "Descrição_Grupo"
    Case 4, 5
        DatDados(0).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
        DatDados(0).Refresh
        lblTipo.Caption = "Fornecedores"
        DBCombo1.BoundColumn = "CodEntidade"
        DBCombo1.ListField = "Nome"
        
End Select

End Sub

Private Sub Form_Load()

Set Formulário = Me

Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

CentralizarForm Me, 1

End Sub
