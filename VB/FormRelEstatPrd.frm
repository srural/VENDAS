VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelEstatPrd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7710
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   35
      Top             =   5085
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
      TabIndex        =   33
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatística de Produtos"
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
         Left            =   135
         TabIndex        =   34
         Top             =   45
         Width           =   5955
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fabricante/Fornecedor"
      Height          =   735
      Left            =   2520
      TabIndex        =   31
      Top             =   3465
      Visible         =   0   'False
      Width           =   4515
      Begin VB.Data DatDados 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   1485
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select codentidade, nome From Ent Where Tipo = 2 or Tipo = 5 order by Nome"
         Top             =   225
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FormRelEstatPrd.frx":0000
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   32
         Top             =   270
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
   End
   Begin VB.Frame Frame6 
      Height          =   630
      Left            =   630
      TabIndex        =   26
      Top             =   1350
      Width           =   6465
      Begin VB.OptionButton OpTipo 
         Caption         =   "Agrupado Grupo"
         Height          =   225
         Index           =   4
         Left            =   4590
         TabIndex        =   36
         Top             =   225
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton OpTipo 
         Caption         =   "Todos"
         Height          =   225
         Index           =   2
         Left            =   450
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton OpTipo 
         Caption         =   "Separado por Fornecedor"
         Height          =   225
         Index           =   3
         Left            =   1890
         TabIndex        =   27
         Top             =   225
         Width           =   2370
      End
   End
   Begin VB.Frame FrameBotao 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4275
      TabIndex        =   23
      Top             =   4185
      Width           =   2865
      Begin VB.CommandButton CmdConfImpr 
         Caption         =   "&Configura"
         Height          =   330
         Left            =   45
         TabIndex        =   30
         Top             =   180
         Width           =   1300
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   330
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   1300
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo do Relatório"
      Height          =   600
      Left            =   630
      TabIndex        =   12
      Top             =   750
      Width           =   6465
      Begin VB.OptionButton OpTipo 
         Caption         =   "Média de Produtos"
         Height          =   225
         Index           =   1
         Left            =   3285
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OpTipo 
         Caption         =   "Produtos não Vendidos"
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2910
      Index           =   1
      Left            =   585
      TabIndex        =   13
      Top             =   2070
      Visible         =   0   'False
      Width           =   6510
      Begin VB.Frame Frame1 
         Caption         =   "Periodo"
         Height          =   825
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   6450
         Begin MSComCtl2.DTPicker DtData1 
            Height          =   330
            Left            =   1230
            TabIndex        =   15
            Top             =   315
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   36641
         End
         Begin MSComCtl2.DTPicker DtData2 
            Height          =   330
            Left            =   4350
            TabIndex        =   16
            Top             =   315
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   36641
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Inicial"
            Height          =   255
            Left            =   135
            TabIndex        =   18
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Final"
            Height          =   255
            Left            =   3435
            TabIndex        =   17
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ordem"
         Height          =   1770
         Left            =   60
         TabIndex        =   19
         Top             =   1035
         Width           =   1815
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Média"
            Height          =   225
            Index           =   7
            Left            =   270
            TabIndex        =   25
            Tag             =   "Media"
            Top             =   1440
            Width           =   1305
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Valor"
            Height          =   225
            Index           =   6
            Left            =   270
            TabIndex        =   24
            Tag             =   "Sum(ITP.VlrVdaDia)"
            Top             =   1140
            Width           =   1305
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Quantidade"
            Height          =   225
            Index           =   5
            Left            =   270
            TabIndex        =   22
            Tag             =   "Sum(ITP.Qtd)"
            Top             =   840
            Width           =   1485
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Descrição"
            Height          =   225
            Index           =   4
            Left            =   270
            TabIndex        =   21
            Tag             =   "Descrição_Produto"
            Top             =   555
            Width           =   1305
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Código"
            Height          =   225
            Index           =   3
            Left            =   270
            TabIndex        =   20
            Tag             =   "CodPrd"
            Top             =   255
            Value           =   -1  'True
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   6510
      Begin VB.Frame Frame2 
         Caption         =   "Ordem"
         Height          =   1545
         Left            =   60
         TabIndex        =   8
         Top             =   1035
         Width           =   1845
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Produto"
            Height          =   225
            Index           =   0
            Left            =   270
            TabIndex        =   3
            Top             =   390
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Data Venda"
            Height          =   225
            Index           =   1
            Left            =   270
            TabIndex        =   4
            Top             =   780
            Width           =   1305
         End
         Begin VB.OptionButton OptOrdem 
            Caption         =   "Estoque"
            Height          =   225
            Index           =   2
            Left            =   270
            TabIndex        =   5
            Top             =   1140
            Width           =   1305
         End
      End
      Begin VB.CheckBox ChkZero 
         Caption         =   "Considerar apenas produtos em estoque"
         Height          =   285
         Left            =   2010
         TabIndex        =   6
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo"
         Height          =   825
         Left            =   60
         TabIndex        =   9
         Top             =   75
         Width           =   6435
         Begin VB.TextBox TxtDias 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            TabIndex        =   10
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Produtos não vendidos a mais de "
            Height          =   375
            Index           =   0
            Left            =   270
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   2745
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dias"
            Height          =   195
            Left            =   4380
            TabIndex        =   11
            Top             =   360
            Width           =   285
         End
      End
   End
End
Attribute VB_Name = "FormRelEstatPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CamposSql As String
Dim CondicaoWhere As String
Dim SqlGroup As String
Dim Ordenado As String
Dim SQL As String
Dim Dif As Currency
Public VarTotal As Currency

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim VarBco As Recordset



    If OpTipo(0).value = True Then
        If IsNull(TxtDias) Then
            TxtDias = 0
        End If
        
        If TxtDias = "" Or Not IsNumeric(TxtDias) Then
            TxtDias = 0
        End If
        
        If OpTipo(3).value = True Then
            If IsNumeric(DbcDados(0).BoundText) Then
                CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Estoque, PRD.DtVenda, PRD.Margem, ENT.CodEntidade, Prd.DtCompra FROM ENT INNER JOIN PRD ON ENT.CodEntidade = PRD.Fabricante "
'                condicapwhere = " WHERE (((PRD.Estoque)>0) AND ((PRD.DtVenda)<=#" & Format(Date - TxtDias, "mm/dd/yyyy") & "#)) "
'                Ordenado = " ORDER BY PRD.Descrição_Produto"
            Else
                MsgBox "Fabricante/Fornecedor não informado"
                Exit Sub
            End If
        ElseIf OpTipo(4).value = True Then
                CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Estoque, PRD.DtVenda, PRD.Margem, PRD.Grupo FROM PRD ORDER BY PRD.Descrição_Produto, PRD.Grupo"
            Else
                CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, PRD.Custo, PRD.Venda, PRD.Estoque, PRD.DtVenda, PRD.Margem FROM PRD  "
                
        End If
        If OptOrdem(0).value = True Then
            Ordenado = " Order by PRD.Descrição_Produto "
        ElseIf OptOrdem(1).value = True Then
                Ordenado = " Order by PRD.DtVenda "
            Else
                Ordenado = " Order by PRD.Estoque Desc "
        End If
        
        If OpTipo(4).value = True Then
            Ordenado = " ORDER BY PRD.Descrição_Produto, PRD.Grupo "
        End If
        
        CondicaoWhere = " Where Prd.DtVenda <= #" & Format(Date - TxtDias, "mm/dd/yyyy") & "# AND Prd.DtCompra <= #" & Format(Date - TxtDias, "mm/dd/yyyy") & "#"
        
        If ChkZero.value = 1 Then
            CondicaoWhere = CondicaoWhere & " and   Prd.Estoque > 0 "
        End If
        
        If OpTipo(3).value = True Then
            CondicaoWhere = CondicaoWhere & " and Ent.CodEntidade = " & DbcDados(0).BoundText
        End If
        
        SQL = CamposSql & CondicaoWhere & Ordenado

        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelEstatPrd.Relatorio.Recordset = varPesquisa(1)
        FrmRelEstatPrd.rpCabTitulo.Caption = "Relatório de Produtos não Vendidos"
        FrmRelEstatPrd.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelEstatPrd.Config
    Else
    
        Dif = CDate(DtData2) - CDate(DtData1)
        If OpTipo(3).value = True Then
            If IsNumeric(DbcDados(0).BoundText) Then
                CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.VlrVdaDia) AS SomaDeValor, Sum(([Qtd]/" & Dif & ")*30) AS Media, Prd.Fabricante FROM PED LEFT JOIN (PRD RIGHT JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido  "
                CondicaoWhere = " WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3) AND ((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) "
                SqlGroup = " GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Fabricante HAVING (((PRD.Fabricante)= " & DbcDados(0).BoundText & "  )) "
            Else
                MsgBox "Fabricante/Fornecedor não informado"
                Exit Sub
            End If
        Else
            CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.VlrVdaDia) AS SomaDeValor, Sum(([Qtd]/" & Dif & ")*30) AS Media FROM PED LEFT JOIN (PRD RIGHT JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido  "
            CondicaoWhere = " WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3) AND ((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) "
            SqlGroup = " GROUP BY PRD.CodPrd, PRD.Descrição_Produto  "
        End If
        
        For Each Componentes In OptOrdem
            If Componentes.value = True Then
                If Componentes.Index >= 5 And Componentes.Index <> 7 Then
                    Ordenado = " order by " & Componentes.Tag & " desc "
                ElseIf Componentes.Index = 7 Then
                    Ordenado = " order by Sum(([Qtd]/" & Dif & ")*30) desc"
                Else
                    Ordenado = " order by " & Componentes.Tag
                End If
            End If
        Next
           
        SQL = CamposSql & CondicaoWhere & SqlGroup & Ordenado
        
        VarTotal = 0
        Set VarBco = Banco.OpenRecordset(SQL)
        If VarBco.RecordCount > 0 Then
            VarBco.MoveFirst
            VarTotal = 0
            While Not VarBco.Eof
                If Not IsNull(VarBco("SomaDeValor")) Then
                    VarTotal = VarTotal + CCur(VarBco("SomaDeValor"))
                End If
                VarBco.MoveNext
            Wend
        End If
        
        
        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelPrdMedia.Relatorio.Recordset = varPesquisa(1)
        FrmRelPrdMedia.rpCabTitulo.Caption = "Relatório de Média Mensal de Vendas"
        FrmRelPrdMedia.rpCabPeriodo.Caption = "Periodo de " & Dif & " dias "
        FrmRelPrdMedia.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelPrdMedia.Rpfabricante.Caption = DbcDados(0).Text
        FrmRelPrdMedia.Config
        
    End If

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DtData1.value = Date - 30
DtData2.value = Date

DatDados(0).DatabaseName = Caminho & "\Dados.mdb"

End Sub

Private Sub opTipo_Click(Index As Integer)
    If OpTipo(0).value = True Then
        Frame(0).Visible = True
        Frame(1).Visible = False
    Else
        Frame(1).Visible = True
        Frame(0).Visible = False
    End If
    
    If OpTipo(3) = True Then
            Frame7.Visible = True
        Else
            Frame7.Visible = False
    End If
End Sub
