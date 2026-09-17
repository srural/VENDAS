VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormAtuCxa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   HelpContextID   =   140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13215
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13215
      TabIndex        =   17
      Top             =   7875
      Width           =   13245
   End
   Begin VB.CheckBox ChkPromo 
      Caption         =   "Só as Promoções"
      Height          =   285
      Left            =   5715
      TabIndex        =   16
      Top             =   900
      Width           =   1770
   End
   Begin VB.Data DatDados 
      Caption         =   "Itp"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Index           =   1
      Left            =   11070
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2925
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton AtuValores 
      Caption         =   "Atualiza Preços"
      Height          =   330
      Left            =   8235
      TabIndex        =   15
      Top             =   810
      Width           =   1950
   End
   Begin VB.CheckBox ChkEntradas 
      Caption         =   "&Entradas"
      Height          =   285
      Left            =   7110
      TabIndex        =   13
      Top             =   675
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13245
      TabIndex        =   11
      Top             =   0
      Width           =   13245
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Atualiza os Caixas"
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
         TabIndex        =   12
         Top             =   45
         Width           =   5370
      End
   End
   Begin ComctlLib.ProgressBar PbCopiaArquivos 
      Height          =   285
      Left            =   8280
      TabIndex        =   10
      Top             =   7380
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton CmdEtiqueta 
      Caption         =   "&Imprimir Etiqueta"
      Height          =   330
      Left            =   11520
      TabIndex        =   9
      Top             =   810
      Width           =   1455
   End
   Begin VB.CheckBox ChkTodos 
      Caption         =   "&Todos"
      Height          =   285
      Left            =   7290
      TabIndex        =   8
      Top             =   810
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   645
      Left            =   135
      TabIndex        =   4
      Top             =   630
      Width           =   5370
      Begin VB.CheckBox ChkAtual 
         Caption         =   "Preços Atualizados"
         Height          =   420
         Left            =   3465
         TabIndex        =   7
         Top             =   180
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   52953089
         CurrentDate     =   36926
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   330
         Index           =   1
         Left            =   1755
         TabIndex        =   6
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   52953089
         CurrentDate     =   36926
      End
   End
   Begin ComctlLib.ProgressBar Bar 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   7380
      Visible         =   0   'False
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir Relatório"
      Height          =   330
      Left            =   9045
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton CmdAtualiza 
      Caption         =   "&Atualiza os Caixas"
      Height          =   330
      Left            =   8280
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4140
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2430
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAtuCxa.frx":0000
      Height          =   2610
      Left            =   180
      OleObjectBlob   =   "FormAtuCxa.frx":001A
      TabIndex        =   0
      Top             =   1350
      Width           =   12780
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FormAtuCxa.frx":10FD
      Height          =   3135
      Left            =   180
      OleObjectBlob   =   "FormAtuCxa.frx":1117
      TabIndex        =   14
      ToolTipText     =   "Duplo Click p/ Alterar Entrega"
      Top             =   4095
      Width           =   12795
   End
End
Attribute VB_Name = "FormAtuCxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SqlPrinter As String
Public Data1 As String
Public Data2 As String
Public Caixa As Recordset
Dim CamposSql As String
Dim Ordenado As String
Dim SQL As String

Private Sub AtuValores_Click()

Dim VarVenda As Currency

DatDados(1).Recordset.MoveFirst
While Not DatDados(1).Recordset.Eof

    DatDados(1).Recordset.Edit
    DatDados(1).Recordset.Fields("Custo") = DatDados(1).Recordset.Fields("ValorUnit")
  '  If OptCompra(0).value = True Then
        VarVenda = DatDados(1).Recordset.Fields("ValorUnit") + (DatDados(1).Recordset.Fields("ValorUnit") * (DatDados(1).Recordset.Fields("Margem") / 100))
        DatDados(1).Recordset.Fields("Venda") = VarVenda
        If VarPzoProduto = True Then
            DatDados(1).Recordset.Fields("PcoPrazo") = VarVenda + (VarVenda * (DatDados(1).Recordset.Fields("MargemPrazo") / 100))
        End If
  '  Else
  '      DatDados(1).Recordset.Fields("Margem") = ((DatDados(1).Recordset.Fields("Venda") / (DatDados(1).Recordset.Fields("ValorUnit")) - 1) * 100)
  '  End If
    DatDados(1).Recordset.Fields("DtAlt") = Date
    DatDados(1).Recordset.Update
    DatDados(1).Recordset.MoveNext
Wend
DatDados(0).Recordset.Requery


End Sub

Private Sub BtImprimir_Click()

'FormCopia.CommonDialog.ShowPrinter
            
'MDIPrincipal.RptRel.SelectionFormula = SqlPrinter

CamposSql = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, GRU.CodGru, GRU.Descrição_Grupo, PRD.Promocao, PRD.Venda_Promocao, PRD.Venda , PRD.Preco FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo"
Ordenado = " ORDER BY Prd.Descrição_Produto"
SQL = CamposSql & SqlPrinter & Ordenado

Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelPrd.Relatorio.Recordset = varPesquisa(1)
FrmRelPrd.rpCabTitulo.Caption = "Relatório de Produtos"
FrmRelPrd.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelPrd.Config

End Sub

Private Sub ChkAtual_Click()
AtualizaLista
End Sub

Private Sub ChkPromo_Click()
AtualizaLista
End Sub

Private Sub ChkTodos_Click()
Dim SQL As String

If ChkTodos.value = 1 Then
    DatDados(0).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, PRD.Venda_Promocao FROM PRD ORDER BY PRD.Descrição_Produto"
    SqlPrinter = ""
    Frame1.Enabled = False
    DatDados(0).Refresh
Else
    AtualizaLista
    Frame1.Enabled = True
End If

End Sub

Private Sub CmdAtualiza_Click()
Dim Criterio As String
Dim DbTerminal As Recordset
Dim x As Long
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveLast
    Bar.Min = 0
    Bar.Max = DatDados(0).Recordset.RecordCount
    
    DatDados(0).Recordset.MoveFirst
    While Not DatDados(0).Recordset.Eof
'        Criterio = "CodPrd = " & DatDados(0).Recordset.Fields("CodPrd")
'        Registros.FindFirst Criterio
'        If Not Registros.NoMatch Then
            Registros.Edit
            If DatDados(0).Recordset.Fields("Promocao") = -1 Then
                Registros.Fields("Preco") = DatDados(0).Recordset.Fields("Venda_Promocao")
'            ElseIf DatDados(0).Recordset.Fields("Concorrencia") = -1 Then
'                Registros.Fields("Preco") = DatDados(0).Recordset.Fields("Venda_Concorrencia")
            Else
                Registros.Fields("Preco") = DatDados(0).Recordset.Fields("Venda")
            End If
            Registros.Fields("Atual") = True
            Registros.Update
            x = x + 1
            Bar.value = x
'        End If
        DoEvents
        DatDados(0).Recordset.MoveNext
    Wend
    Caixa.MoveFirst
    While Not Caixa.Eof
        If Caixa!Maquina <> "" Then
            Open Caminho & "\" & Caixa!Maquina For Output As #1
            Print #1, atualizado
            Close #1
        End If
        Caixa.MoveNext
    Wend
        
    If VarTerminalConsulta <> "" Then
        Set DbTerminal = Banco.OpenRecordset("SELECT BAR.BarBarra, PRD.Descrição_Produto, PRD.Preco FROM BAR INNER JOIN PRD ON BAR.Produto = PRD.CodPrd WHERE (((BAR.BarBarra)<> ' ')) ORDER BY BAR.BarBarra ")
        DbTerminal.Requery
        If DbTerminal.RecordCount > 0 Then
            DbTerminal.MoveFirst
            If Dir(VarTerminalConsulta & "\produtos.txt") <> "" Then
'                 Kill VarTerminalConsulta & "\produtos.txt"
            End If
            
            Open VarTerminalConsulta & "\produtos.txt" For Output As #1
            While Not DbTerminal.Eof
'                Print #1, CDec(Trim(DbTerminal!BARBARRA)) & ";" & Left(Trim(DbTerminal!Descrição_Produto), 35) & ";" & Trim(CStr(Int(DbTerminal!Preco) & "." & Right(Format(DbTerminal!Preco, "###,##0.00"), 2)))
                Print #1, CStr(Trim(DbTerminal!BarBarra)) & "|" & Left(Trim(DbTerminal!Descrição_Produto), 35) & "|" & Trim(Format(CStr(DbTerminal!Preco), "#,##0.00"))
                DbTerminal.MoveNext
            Wend
            Close #1
        End If
    End If
    
    If VarEmergencia <> "" Then
        If Dir(VarEmergencia & "\PRODUTOS.MDB") <> "" Then
            Kill VarEmergencia & "\produtos.mdb"
        End If
'        CopiarArquivo App.Path & "\produtos.mdb", VarEmergencia & "\produtos.mdb"
    End If
        
    MsgBox "Caixas Atualizados", vbInformation, App.Title
End If
DatDados(0).Refresh
DBGrid1.Refresh
End Sub

Private Sub CmdEtiqueta_Click()
FormZebra.Show vbModal
End Sub

Private Sub DtData_Change(Index As Integer)
    AtualizaLista
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

'Set Formulário = Me

FormAtuCxa.Left = 200
FormAtuCxa.Top = 50

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

Set Registros = Banco.OpenRecordset("SELECT  PRD.CodPrd, PRD.Venda, PRD.DtAlt, PRD.Venda_Promocao FROM PRD ORDER BY PRD.CodPrd")
'Set Caixa = Banco.OpenRecordset("SELECT *  FROM AtuCxa ORDER BY Codigo")
'Caixa.Requery

DtData(0).value = Date
DtData(1).value = Date

DtData_Change 0

DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.* FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where (((Ped.DataEmiss)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (Ped.DataEmiss) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#)) Order by Itp.CodItp"
DatDados(1).Refresh
SQL = "SELECT PED.*, ITP.*, PRD.* FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where (((Ped.DataEmiss)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (Ped.DataEmiss) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#)) Order by Itp.CodItp"


End Sub

Private Sub AtualizaLista()

'Data1 = Format(DtData(0), "mm/dd/yyyy")
'Data2 = Format(DtData(1), "mm/dd/yyyy")

If ChkPromo.value = 0 Then
        DatDados(0).RecordSource = "SELECT PRD.PROMOCAO, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda , PRD.DtAlt, PRD.Venda_Promocao FROM PRD WHERE (((PRD.DtAlt)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (PRD.DtAlt) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#))  ORDER BY PRD.Descrição_Produto"
        SqlPrinter = " WHERE (((PRD.DtAlt)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (PRD.DtAlt) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#)) AND Prd.Atual <> True"
Else
        DatDados(0).RecordSource = "SELECT PRD.PROMOCAO,  PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda , PRD.DtAlt, PRD.Venda_Promocao FROM PRD WHERE  PRD.PROMOCAO = TRUE  ORDER BY PRD.Descrição_Produto"
'        DatDados(0).RecordSource = "SELECT prd.Atual, PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, PRD.Promocao, PRD.Venda_Promocao, PRD.Concorrencia, PRD.Venda_Concorrencia, PRD.DtAlt, Prd.Preco, Prd.Chefe, Prd.Chefe_Grupo FROM PRD WHERE (((PRD.DtAlt)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (PRD.DtAlt) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#)) AND Prd.Atual = True  ORDER BY PRD.Descrição_Produto"
'        SqlPrinter = " WHERE (((PRD.DtAlt)>= #" & Format(DtData(0), "mm/dd/yyyy") & "# and (PRD.DtAlt) <= #" & Format(DtData(1), "mm/dd/yyyy") & "#)) AND Prd.Atual = True"
End If

' SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Venda, PRD.DtAlt, PRD.Venda_Promocao, PED.DataEmiss FROM PRD INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE (((PED.DataEmiss)>=#1/10/2014# And (PED.DataEmiss)<=#10/10/2014#)) ORDER BY PRD.Descrição_Produto


DatDados(0).Refresh

End Sub


Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
If DBGrid2.Columns(9) <> "" Then
    If IsNumeric(DBGrid2.Columns(9)) Then
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset.Fields("Margem") = ((DBGrid2.Columns(9) / DBGrid2.Columns(4)) - 1) * 100
        DatDados(1).Recordset.Update
        DatDados(0).Recordset.Requery
    End If
    
End If

End Sub





