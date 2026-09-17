VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelPrd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6330
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8520
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -225
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   25
      Top             =   6030
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
      TabIndex        =   23
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Produtos"
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
         TabIndex        =   24
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.CheckBox ChkPromocao 
      Caption         =   "Produtos em Promoção"
      Height          =   240
      Left            =   390
      TabIndex        =   22
      Top             =   4470
      Width           =   2385
   End
   Begin VB.CheckBox ChkEstoque 
      Caption         =   "Mostrar Estoque"
      Height          =   240
      Left            =   390
      TabIndex        =   21
      Top             =   5025
      Width           =   2385
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   345
      Left            =   6570
      TabIndex        =   20
      Top             =   4785
      Width           =   1395
   End
   Begin VB.CheckBox ChkQuebra 
      Caption         =   "&Grupos em Páginas Separadas"
      Height          =   240
      Left            =   390
      TabIndex        =   19
      Top             =   5325
      Width           =   2580
   End
   Begin VB.CheckBox ChkEstoqueZero 
      Caption         =   "Listar Produtos Estoque Zero"
      Height          =   240
      Left            =   390
      TabIndex        =   13
      Top             =   4755
      Width           =   2385
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   210
      Left            =   2925
      TabIndex        =   12
      Top             =   4725
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox ChkPrazo 
      Caption         =   "Listar Preço a Prazo"
      Height          =   240
      Left            =   390
      TabIndex        =   11
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Ordem"
      Height          =   960
      Left            =   360
      TabIndex        =   3
      Top             =   3450
      Width           =   2415
      Begin VB.OptionButton OptOrdem 
         Caption         =   "&Alfabética"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "&Numérica"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   555
         Width           =   1455
      End
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4410
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1575
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frm1 
      Caption         =   "&Tipo"
      Height          =   2730
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   2415
      Begin VB.OptionButton Opt1 
         Caption         =   "F&amilia Específica"
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   18
         Top             =   2385
         Width           =   2055
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Familia Todos"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   60
         TabIndex        =   16
         Top             =   1950
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   60
         TabIndex        =   15
         Top             =   1215
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   60
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Todos Produtos"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Grupo Todos"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   570
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Grupo &Específico"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   8
         Top             =   855
         Width           =   1815
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Fornecedor Todos"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   7
         Top             =   1380
         Width           =   1815
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Forne&cedor Específico"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   6
         Top             =   1665
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6570
      TabIndex        =   0
      Top             =   5280
      Width           =   1395
   End
   Begin MSDBCtls.DBCombo Combo 
      Bindings        =   "FormRelPrd.frx":0000
      DataSource      =   "DatDados(0)"
      Height          =   3690
      Left            =   2880
      TabIndex        =   2
      Top             =   690
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6482
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      Style           =   1
      ListField       =   "Descrição_Grupo"
      BoundColumn     =   "CODGru"
      Text            =   ""
   End
End
Attribute VB_Name = "FormRelPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VarIndice As Currency
Public PcoPzo As Byte


Private Sub ChkPrazo_Click()
Dim Ind As Recordset
Dim Prod As Recordset
Dim Rel As Recordset
Dim Sql As String
Dim X As Integer

If chkprazo.value = 1 Then
    
    Command1.Enabled = False
    
    Barra.Visible = True
    Barra.Min = 0

    Barra.value = 0
    
    Set Ind = Banco.OpenRecordset("Select * from  Ind")
    Ind.Requery
    
    Set Prod = Banco.OpenRecordset("Select * From Prd")
    Set Rel = Banco.OpenRecordset("Select * From RelPrd")
    Sql = "Delete * from RelPrd"
    Banco.Execute Sql
    
    Prod.MoveLast
    Barra.Max = Prod.RecordCount
    Prod.MoveFirst
    While Not Prod.Eof
            Rel.AddNew
            Rel("produto") = Prod("CodPrd")
            Ind.MoveFirst
            X = 1
            While Not Ind.Eof
                If X <= 5 Then
                        Rel("Prazo" & X) = Prod("Venda") * Ind("Indice")
                End If
                Ind.MoveNext
                X = X + 1
            Wend
            Rel.Update
            Barra.value = Prod.AbsolutePosition
            Prod.MoveNext
    Wend
    Command1.Enabled = True
    Barra.Visible = False
End If
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub Command1_Click()
Dim X As Variant
Dim Op1 As Integer
Dim Op2 As Integer
Dim Sql As String
Dim CamposSql As String
Dim CondicaoWhere As String
Dim Ordenado As String
Dim Ind As Recordset
'With RptRel
'    .DataFiles(0) = "c:\aateste\cadastro\cadastro.mdb" 'somente p/ access
'    .ReportFileName = "c:\aateste\cadastro\RelCli.rpt"
'    .Destination = crptToWindow
       
'    .SelectionFormula = "UPPERCASE({CLI.CIDADE}) =""" & UCase(Cidade) & """"
        
    Set Ind = Banco.OpenRecordset("Select * from Ind Where NroIndice = 3 ")
    Ind.Requery
    If Ind.RecordCount > 0 Then
        VarIndice = Ind!Perc
    End If
        
    PcoPzo = LerINI("Pedido", "PcoPrazo", App.Path & "\config.ini")
    If PcoPzo = 0 Then
        VarIndice = 0
    End If
        
    For Each X In Opt1
        If X.value Then
            Op1 = X.Index
        End If
    Next
   
    Select Case Op1
        Case 1
            If chkprazo.value = 1 Then
                CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, RELPRD.Prazo1, RELPRD.Prazo2, RELPRD.Prazo3, RELPRD.Prazo4, RELPRD.Prazo5 FROM GRU, RELPRD INNER JOIN PRD ON RELPRD.Produto = PRD.CodPrd GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, RELPRD.Prazo1, RELPRD.Prazo2, RELPRD.Prazo3, RELPRD.Prazo4, RELPRD.Prazo5 "
                CondicaoWhere = " HAVING Estoque > 0 and Ativo = True "
                Ordenado = " ORDER BY " & IIf((OptOrdem(0).value = True), "Prd.Descrição_Produto", "Prd.CodPrd")
                Sql = CamposSql & Ordenado
                
                Set varPesquisa(1) = Banco.OpenRecordset(Sql)
                Set FrmRelPrdPzo.Relatorio.Recordset = varPesquisa(1)
                FrmRelPrdPzo.rpCabTitulo.Caption = "Relatório de Produtos"
                FrmRelPrdPzo.rpCabSubTitulo.Caption = NomeEmpresa
                FrmRelPrdPzo.Config
            Else
                CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda, Prd.PcoPrazo, Prd.Validade, GRU.Descrição_Grupo, PRD.Embalagem, PRD.CodBar, Prd.Promocao, Prd.Venda_Promocao FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, Prd.Validade, PRD.Estoque, PRD.Venda, Prd.PcoPrazo, GRU.Descrição_Grupo, PRD.Embalagem, PRD.CodBar , Prd.Promocao, Prd.Venda_Promocao "
                CondicaoWhere = " HAVING  Ativo = True "
                Ordenado = "ORDER BY " & IIf((OptOrdem(0).value = True), "Descrição_Produto", "CodPrd")
                
                If ChkPromocao.value = 1 Then
                     CondicaoWhere = CondicaoWhere & " and  Promocao = true "
                End If
                
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & " and  Estoque > 0 " & Ordenado
                Else
'                    If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
'                        Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
'                    Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
'                    End If
                End If
                                                               
                Set varPesquisa(1) = Banco.OpenRecordset(Sql)
                Set FrmRelPrd.Relatorio.Recordset = varPesquisa(1)
                FrmRelPrd.rpCabTitulo.Caption = "Relatório de Produtos"
                FrmRelPrd.rpCabSubTitulo.Caption = NomeEmpresa
                If ChkEstoque.value = 0 Then
                    FrmRelPrd.RpEstoque.Mostrar = False
                End If
                FrmRelPrd.Config
            End If

        Case 2
          
            CamposSql = " SELECT PRD.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, GRU.CodGru, ENT.CodEntidade, ENT.Nome FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade GROUP BY PRD.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, GRU.CodGru, ENT.CodEntidade, ENT.Nome "
            Ordenado = " ORDER BY " & IIf((OptOrdem(0).value = True), "GRU.CodGru, PRD.Descrição_Produto", "GRU.CodGru, PRD.CodPrd")
            CondicaoWhere = " Having  Prd.Ativo = True "
            
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & " and Prd.Estoque > 0 " & Ordenado
                Else
'                    If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
'                        Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
'                    Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
'                    End If
                End If
            
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdGru.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdGru.RodGrupo(0).QuebraDepois = False
            End If
            
            Set FrmRelPrdGru.Relatorio.Recordset = varPesquisa(1)
            FrmRelPrdGru.rpCabTitulo.Caption = "Relatório de Produtos"
            FrmRelPrdGru.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdGru.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdGru.Config
            
        Case 3
            CamposSql = " SELECT Prd.Margem_Concorrencia,  Prd.Venda_Concorrencia, Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, GRU.CodGru  FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo)  GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, GRU.CodGru, Prd.Margem_Concorrencia,  Prd.Venda_Concorrencia  "
            Ordenado = "ORDER BY " & IIf((OptOrdem(0).value = True), "GRU.CodGru, PRD.Descrição_Produto", "GRU.CodGru, PRD.CodPrd")
            CondicaoWhere = " Having Prd.Ativo = True AND Gru.Descrição_Grupo Like '" & Combo.Text & "'"
            
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & " and Prd.Estoque  > 0  " & Ordenado
                Else
 '                   If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
 '                       Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
 '                   Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
 '                  End If
                End If
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdGru.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdGru.RodGrupo(0).QuebraDepois = False
            End If
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            Set FrmRelPrdGru.Relatorio.Recordset = varPesquisa(1)
            FrmRelPrdGru.rpCabTitulo.Caption = "Relatório de Produtos"
            FrmRelPrdGru.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdGru.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdGru.Config
            
        Case 4

            CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo "
            Ordenado = "ORDER BY " & IIf((OptOrdem(0).value = True), "Ent.Nome, PRD.Descrição_Produto", "Ent.Nome, PRD.CodPrd")
            CondicaoWhere = "Having  Prd.Ativo = True "
            
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & " and Prd.Estoque  > 0 " & Ordenado
                Else
 '                   If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
 '                       Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
 '                   Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
 '                   End If
                End If
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdFor.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdFor.RodGrupo(0).QuebraDepois = False
            End If
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            Set FrmRelPrdFor.Relatorio.Recordset = varPesquisa(1)
            
            FrmRelPrdFor.rpCabTitulo.Caption = "Relatório de Produtos"
            FrmRelPrdFor.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdFor.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdFor.Config
            
        Case 5
            
            CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  ENT.Nome, ENT.CodEntidade, GRU.CodGru, GRU.Descrição_Grupo "
            Ordenado = "ORDER BY " & IIf((OptOrdem(0).value = True), " Ent.Nome, PRD.Descrição_Produto", "Ent.Nome, PRD.CodPrd")
            CondicaoWhere = "Having Prd.Ativo = True and Ent.Nome Like '" & Combo.Text & "'"
            
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & " and  Prd.Estoque  > 0 " & Ordenado
                Else
  '                  If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
  '                      Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
  '                  Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
  '                  End If
                End If
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdFor.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdFor.RodGrupo(0).QuebraDepois = False
            End If
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            Set FrmRelPrdFor.Relatorio.Recordset = varPesquisa(1)
            
            FrmRelPrdFor.rpCabTitulo.Caption = "Relatório de Produtos por Fornecedor"
            FrmRelPrdFor.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdFor.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdFor.Config
            
        Case 0

            CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia FROM ((FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade GROUP BY Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia "
            Ordenado = "ORDER BY " & IIf((OptOrdem(0).value = True), "FAM.Descrição_Familia, GRU.Descrição_Grupo, PRD.Descrição_Produto", "FAM.Descrição_Familia, GRU.Descrição_Grupo, PRD.CodPrd")
            CondicaoWhere = " Having Prd.Estoque  > 0 and Prd.Ativo = True "
            
                If ChkEstoqueZero.value = 0 Then
                    Sql = CamposSql & CondicaoWhere & Ordenado
                Else
 '                   If MsgBox("Só estoque zero ?", vbYesNo) = vbYes Then
 '                       Sql = CamposSql & CondicaoWhere & " and  Estoque = 0 " & Ordenado
 '                   Else
                        Sql = CamposSql & CondicaoWhere & Ordenado
 '                   End If
                End If
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdFam.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdFam.RodGrupo(0).QuebraDepois = False
            End If
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            Set FrmRelPrdFam.Relatorio.Recordset = varPesquisa(1)
            
            FrmRelPrdFam.rpCabTitulo.Caption = "Relatório de Produtos"
            FrmRelPrdFam.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdFam.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdFam.Config
            
        Case 6
            
            CamposSql = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia FROM ((FAM INNER JOIN GRU ON FAM.CodFam = GRU.Familia) INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade GROUP BY Prd.Ativo,PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque, PRD.Venda,  Prd.PcoPrazo,  GRU.Descrição_Grupo, ENT.Nome, GRU.CodGru, ENT.CodEntidade, FAM.CodFam, FAM.Descrição_Familia "
            Ordenado = "ORDER BY FAM.Descrição_Familia, GRU.Descrição_Grupo, PRD.Descrição_Produto "
            CondicaoWhere = "Having Prd.Ativo = True And  Fam.Descrição_Familia Like '" & Combo.Text & "'"
            
 '               If ChkEstoqueZero.value = 0 Then
 '                   Sql = CamposSql & CondicaoWhere & " and Prd.Estoque > 0 " & Ordenado
 '               Else
                    Sql = CamposSql & CondicaoWhere & Ordenado
 '               End If
            
            If ChkQuebra.value = 1 Then
                FrmRelPrdFam.RodGrupo(0).QuebraDepois = True
            Else
                FrmRelPrdFam.RodGrupo(0).QuebraDepois = False
            End If
            
            Set varPesquisa(1) = Banco.OpenRecordset(Sql)
            Set FrmRelPrdFam.Relatorio.Recordset = varPesquisa(1)
            FrmRelPrdFam.rpCabTitulo.Caption = "Relatório de Produtos"
            FrmRelPrdFam.rpCabSubTitulo.Caption = NomeEmpresa
            If ChkEstoque.value = 0 Then
                FrmRelPrdFam.RpEstoque.Mostrar = False
            End If
            
            FrmRelPrdFam.Config
            
    End Select
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Activate()

DatDados(0).RecordSource = "SELECT Gru.CodGru, Gru.Descrição_Grupo   FROM Gru ORDER BY Gru.Descrição_Grupo"
'DatDados(1).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
DatDados(0).Refresh
'DatDados(1).Refresh
End Sub

Private Sub Opt1_Click(Index As Integer)
Select Case Index
    
    Case 2, 3
        DatDados(0).RecordSource = "SELECT Gru.CodGru, Gru.Descrição_Grupo   FROM Gru ORDER BY Gru.Descrição_Grupo"
        DatDados(0).Refresh
        
        Combo.BoundColumn = "CODGru"
        Combo.ListField = "Descrição_Grupo"
        Combo.Refresh
        
        If Index = 3 Then
            Combo.Enabled = True
        Else
            Combo.Enabled = False
        End If
        
    Case 4, 5
        DatDados(0).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 ORDER BY Ent.nome"
        DatDados(0).Refresh
        
        Combo.BoundColumn = "CodEntidade"
        Combo.ListField = "Nome"
        Combo.Refresh

        If Index = 5 Then
            Combo.Enabled = True
        Else
            Combo.Enabled = False
        End If
    Case 0, 6
        DatDados(0).RecordSource = "SELECT Fam.CodFam,Fam.Descrição_Familia  FROM Fam ORDER BY Fam.Descrição_Familia"
        DatDados(0).Refresh
        
        Combo.BoundColumn = "CodFam"
        Combo.ListField = "Descrição_Familia"
        Combo.Refresh

        If Index = 6 Then
            Combo.Enabled = True
        Else
            Combo.Enabled = False
        End If
        
End Select

    If Index = 2 Or Index = 4 Or Index = 0 Then
        ChkQuebra.value = 0
        ChkQuebra.Enabled = True
    Else
        ChkQuebra.value = 0
        ChkQuebra.Enabled = False
    End If

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub
