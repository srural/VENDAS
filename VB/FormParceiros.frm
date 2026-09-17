VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormParceiros 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   14145
   Icon            =   "FormParceiros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   14145
   Begin VB.CommandButton CmdFolha 
      Caption         =   "Fecha a Folha"
      Height          =   375
      Left            =   12375
      TabIndex        =   33
      Top             =   4905
      Width           =   1425
   End
   Begin MSMask.MaskEdBox MebVale 
      Height          =   375
      Left            =   12330
      TabIndex        =   31
      Top             =   3825
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmbSugestao 
      Caption         =   "Recibo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12375
      TabIndex        =   30
      Top             =   4410
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissão"
      Enabled         =   0   'False
      Height          =   780
      Left            =   6975
      TabIndex        =   23
      Top             =   7695
      Width           =   5055
      Begin MSMask.MaskEdBox MebTotal 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   24
         Tag             =   "S"
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebComissao 
         DataField       =   "Custo"
         Height          =   270
         Left            =   3360
         TabIndex        =   25
         Tag             =   "S"
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTaxa 
         DataField       =   "Custo"
         Height          =   270
         Left            =   1740
         TabIndex        =   26
         Tag             =   "S"
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   3240
         X2              =   3240
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   1620
         X2              =   1620
         Y1              =   75
         Y2              =   765
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "¨Taxa Cartão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   7
         Left            =   1680
         TabIndex        =   29
         Top             =   225
         Width           =   810
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   6
         Left            =   3285
         TabIndex        =   28
         Top             =   225
         Width           =   630
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   27
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   2
      Left            =   6570
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormParceiros.frx":000C
      Top             =   3645
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Enabled         =   0   'False
      Height          =   780
      Left            =   225
      TabIndex        =   14
      Top             =   7695
      Width           =   6585
      Begin MSMask.MaskEdBox MebVista 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   15
         Tag             =   "S"
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebCartao 
         DataField       =   "Custo"
         Height          =   270
         Left            =   3360
         TabIndex        =   16
         Tag             =   "S"
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPrazo 
         DataField       =   "Custo"
         Height          =   270
         Left            =   1740
         TabIndex        =   17
         Tag             =   "S"
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebCortesia 
         DataField       =   "Custo"
         Height          =   270
         Left            =   5055
         TabIndex        =   21
         Tag             =   "S"
         Top             =   405
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   4860
         X2              =   4860
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Cartesia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   3
         Left            =   4995
         TabIndex        =   22
         Top             =   225
         Width           =   525
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Vista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   20
         Top             =   225
         Width           =   315
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Cartão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   1
         Left            =   3300
         TabIndex        =   19
         Top             =   225
         Width           =   420
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   10
         Left            =   1680
         TabIndex        =   18
         Top             =   225
         Width           =   345
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1620
         X2              =   1620
         Y1              =   75
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   3240
         X2              =   3240
         Y1              =   90
         Y2              =   765
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   16410
      TabIndex        =   9
      Top             =   8595
      Width           =   16440
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   16365
      TabIndex        =   7
      Top             =   0
      Width           =   16395
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle Comissão de Parceiros"
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
         TabIndex        =   8
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   6435
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormParceiros.frx":024A
      Top             =   4770
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4140
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where (Tipo = 3 Or Tipo = 7) and Ativo = true  ORDER BY Ent.nome"
      Top             =   4770
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Index           =   0
      Left            =   150
      ScaleHeight     =   1365
      ScaleWidth      =   16245
      TabIndex        =   4
      Top             =   615
      Width           =   16245
      Begin VB.Frame Frame3 
         Caption         =   "Periodo"
         Height          =   1050
         Left            =   6570
         TabIndex        =   6
         Top             =   180
         Width           =   3795
         Begin MSComCtl2.DTPicker DtData1 
            Height          =   330
            Left            =   180
            TabIndex        =   10
            Top             =   525
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   393216
            Format          =   49741825
            CurrentDate     =   36641
         End
         Begin MSComCtl2.DTPicker DtData2 
            Height          =   330
            Left            =   1980
            TabIndex        =   11
            Top             =   525
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   393216
            Format          =   49741825
            CurrentDate     =   36641
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Final"
            Height          =   195
            Left            =   1980
            TabIndex        =   13
            Top             =   315
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data &Inicial"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   315
            Width           =   795
         End
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         TabIndex        =   1
         Top             =   720
         Width           =   1020
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormParceiros.frx":0488
         DataField       =   "Nome_Fornecedor"
         Height          =   945
         Left            =   45
         TabIndex        =   0
         Top             =   270
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   1667
         _Version        =   393216
         Appearance      =   0
         Style           =   1
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin VB.Label LblDados 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Parceiro"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   45
         Width           =   1935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormParceiros.frx":04A2
      Height          =   5610
      Left            =   180
      OleObjectBlob   =   "FormParceiros.frx":04BC
      TabIndex        =   2
      Top             =   1980
      Width           =   11835
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vale"
      Height          =   195
      Left            =   12330
      TabIndex        =   32
      Top             =   3555
      Width           =   690
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   3
      Top             =   735
      Width           =   645
   End
End
Attribute VB_Name = "FormParceiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public lstPesq As ListItem


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbSugestao_Click()
Dim SQL As String


'Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelSugestao.Relatorio.Recordset = varPesquisa(1)
'FrmRelSugestao.rpCabTitulo.Caption = "Relatório de Sugestão de Compras por Fornecedor - " & DbcPesquisa.Text
'FrmRelSugestao.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelRecibo2.rpPeriodo.Caption = "Período de " & Format(DtData1, "dd/mm/yyyy") & " a " & Format(DtData2, "dd/mm/yyyy")
FrmRelRecibo2.rpcomissao.Caption = Format(MebComissao, "#,##0.00")
FrmRelRecibo2.rpTotalCred.Caption = Format(MebComissao, "#,##0.00")
FrmRelRecibo2.rpVale.Caption = Format(MebVale, "#,##0.00")
FrmRelRecibo2.rpTotalDeb.Caption = Format(MebVale, "#,##0.00")
FrmRelRecibo2.rpTotal.Caption = Format(MebComissao - MebVale, "#,##0.00")
FrmRelRecibo2.rpNome.Caption = DbcPesquisa.Text

FrmRelRecibo2.Config

End Sub

Private Sub CmdFolha_Click()
Dim SQL As String
If MsgBox("Deseja descontar o vale ?", vbYesNo) = vbYes Then
    If IsNumeric(DbcPesquisa.BoundText) Then
        SQL = "INSERT INTO DUP ( Entidade, DtEmiss, DtVenc , Valor,  Pedido, Tipo_Dup, NroPcl, Cond ) VALUES ( "
        SQL = SQL & DbcPesquisa.BoundText & " , #"
        SQL = SQL & Format(Date, "mm/dd/yyyy") & "# , #"
        SQL = SQL & Format(Date, "mm/dd/yyyy") & "# , "
        SQL = SQL & Numero(MebVale.Text, 2) & " ,"
        SQL = SQL & Numero(0, 2) & " ,"
        SQL = SQL & """" & "CRE" & """" & " ,"
        SQL = SQL & 1 & " ,"
        SQL = SQL & 0 & ")"
        Banco.Execute SQL
        MebVale = 0
        MsgBox "Vale lançado com sucesso !", vbInformation, App.Title
    Else
        MsgBox "Parceiro não selecionado !", vbInformation, App.Title
    End If

End If
End Sub

Private Sub CmdImprimir_Click()
Dim CamposSql As String
Dim Ordenado As String
Dim CondicaoWhere As String
Dim SQL As String

If Not IsNumeric(DbcPesquisa.BoundText) Then
    SQL = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= 999999)) ORDER BY PED.DataEmiss "
Else
    SQL = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING  (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= " & DbcPesquisa.BoundText & "  )) ORDER BY PED.DataEmiss  "
End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    'Set varPesquisa(1) = DatDados(0).Recordset.Clone
     Set FrmExtratoParceiro.Relatorio.Recordset = varPesquisa(1)
'    Set FrmRelPrdFor_Altera.Relatorio.Recordset = DatDados(0).Recordset.Clone
    
    FrmExtratoParceiro.rpCabTitulo.Caption = "Extrato Parceiro"
    FrmExtratoParceiro.rpCabSubTitulo.Caption = "Periodo de " & DtData1 & " a " & DtData2
    FrmExtratoParceiro.RptParceiro.Caption = DbcPesquisa.Text
    FrmExtratoParceiro.RptTaxa.Caption = (MebCartao * (MebTaxa / 100))
    FrmExtratoParceiro.RptTotal.Caption = MebComissao
    FrmExtratoParceiro.rpCabSubTitulo.Caption = NomeEmpresa
    
    FrmExtratoParceiro.Config

End Sub

Private Sub DbcPesquisa_Change()
Dim SQL As String
Dim sql2 As String
Dim Ordem As String
MebVista = 0
MebPrazo = 0
MebCartao = 0
MebCortesia = 0
Dim CodAnt As Double
Dim ProdutoAnt As String
Dim totQtd As Double
Dim ContList As Double
Dim VarComissao As Currency
Dim varConta As Recordset

VarComissao = 0

totQtd = 0
ContList = 0


If Not IsNumeric(DbcPesquisa.BoundText) Then
    SQL = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= 999999)) ORDER BY PED.DataEmiss "
    sql2 = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= 999999)) ORDER BY PED.DataEmiss "
Else
    SQL = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING  (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= " & DbcPesquisa.BoundText & "  )) ORDER BY PED.DataEmiss  "
    sql2 = "SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING  (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((ITP.Vendedor)= " & DbcPesquisa.BoundText & "  )) ORDER BY PED.DataEmiss  "
End If

'DatDados(0).RecordSource = sql2

'DatDados(0).Refresh
'On Error Resume Next
'If DatDados(0).Recordset.RecordCount > 0 Then
 '   DatDados(0).Recordset.MoveFirst
    
 '   ListVda.ListItems.Clear
    
 '   CodAnt = DatDados(0).Recordset("CodPrd")
    
 '   While Not DatDados(0).Recordset.Eof
    
 '       DatDados(2).RecordSource = ("SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)>=#11/1/2017# And (PED.DataEmiss)<=#11/21/2017#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor HAVING (((PRD.CodPrd)=" & DatDados(0).Recordset("CodPrd") & ") AND ((ITP.Vendedor)= " & DbcPesquisa.BoundText & ")) ORDER BY PRD.CodPrd")

        
'        ("SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, First(ITP.Qtd) AS PrimeiroDeQtd, First(ITP.Valor) AS PrimeiroDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=4)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ITP.Vendedor, ENT.Nome,  PED.CondPgto, CodItp HAVING  (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND  ((PRD.CodPrd)=" & DatDados(0).Recordset("CodPrd") & ") AND ((ITP.Vendedor)= " & DbcPesquisa.BoundText & "  )) ORDER BY PED.DataEmiss ")
'        DatDados(2).Refresh
'        If DatDados(2).Recordset.RecordCount > 0 Then
'            While Not DatDados(2).Recordset.Eof
'                totQtd = totQtd + DatDados(2).Recordset("PrimeiroDeQtd")
'                ProdutoAnt = DatDados(2).Recordset("Descrição_Produto")
'                DatDados(2).Recordset.MoveNext
'                DatDados(0).Recordset.MoveNext
'            Wend
'        End If
'        If Not DatDados(0).Recordset.Eof Then
'            CodAnt = DatDados(0).Recordset("CodPrd")
'            ContList = ContList + 1
'            Set lstPesq = ListVda.ListItems.Add(, , ContList)
            
'            lstPesq.SubItems(1) = ProdutoAnt   'SeCampoNulo(Consulta, "Agencia")
'            lstPesq.SubItems(2) = totQtd   'SeCampoNulo(Consulta, "Agencia")
'
'            totQtd = 0
'        End If
'    Wend
'End If

Datdados(0).RecordSource = SQL

Datdados(0).Refresh

If Datdados(0).Recordset.RecordCount > 0 Then
    Datdados(0).Recordset.MoveFirst
    
    While Not Datdados(0).Recordset.Eof
                
        VarComissao = VarComissao + (Datdados(0).Recordset("PrimeiroDeValor") * (Datdados(0).Recordset("PrimeiroDeComissao") / 100))
                
        Select Case Datdados(0).Recordset("CondPgto")
        Case "DINHEIRO"
            MebVista = MebVista + Datdados(0).Recordset("PrimeiroDeValor")
        Case "A PRAZO"
            MebPrazo = MebPrazo + Datdados(0).Recordset("PrimeiroDeValor")
        Case "CARTAO DEBITO"
            MebCartao = MebCartao + Datdados(0).Recordset("PrimeiroDeValor")
        Case "CARTAO CREDITO"
            MebCartao = MebCartao + Datdados(0).Recordset("PrimeiroDeValor")
        Case "CORTESIA"
            MebCortesia = MebCortesia + Datdados(0).Recordset("PrimeiroDeValor")
        End Select
        
        Datdados(0).Recordset.MoveNext
    Wend
    Datdados(0).Recordset.MoveFirst
    DBGrid1.Refresh
End If

MebTotal = VarComissao
'MebTaxa = 2.5
MebComissao = VarComissao - (MebCartao * (MebTaxa / 100))
If Datdados(0).Recordset.RecordCount > 0 Then
    CmbSugestao.Enabled = True
    CmdImprimir.Enabled = True
Else
    CmbSugestao.Enabled = False
    CmdImprimir.Enabled = False
End If

If IsNumeric(DbcPesquisa.BoundText) Then

'SQL = "SELECT DUP.Tipo_Dup, DUP.DtVenc, DUP.Entidade, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Tipo_Dup, DUP.DtVenc, DUP.Entidade HAVING (((DUP.Tipo_Dup)='DEB') AND (((DUP.DtVenc)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((Dup.Entidade)= " & DbcPesquisa.BoundText & "))"
    Set varConta = Banco.OpenRecordset("SELECT DUP.Tipo_Dup, DUP.DtVenc, DUP.Entidade, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Tipo_Dup, DUP.DtVenc, DUP.Entidade HAVING ((DUP.Tipo_Dup)='DEB') AND (((DUP.DtVenc)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((Dup.Entidade)= " & DbcPesquisa.BoundText & "))")
    varConta.Requery
    If varConta.RecordCount > 0 Then
        MebVale = varConta!SomaDeValor
    End If
End If


End Sub

Private Sub DbcPesquisa_LostFocus()
Dim Criterio As String
If IsNumeric(DbcPesquisa.Text) Then
   Criterio = "CodEntidade = " & DbcPesquisa.Text
   Datdados(1).Recordset.FindFirst Criterio
   If Datdados(1).Recordset.NoMatch Then
       MsgBox "Informação não Localizada", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
        DbcPesquisa.Text = Datdados(1).Recordset.Fields("Nome")
        DBGrid1.SetFocus
   End If
End If

End Sub


Private Sub DtData1_Change()
DbcPesquisa_Change
End Sub

Private Sub DtData2_Change()
DbcPesquisa_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then SendKeys "{Tab}"

End Sub


Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
Formulário.Left = 300
Formulário.Top = 0

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
DbcPesquisa.SetFocus
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DtData1 = Date
DtData2 = Date

If LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini") <> "" Then
    MebTaxa = LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini")
Else
'    TxtEvento = ""
End If

MebVale = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
Dim SQL As String

SQL = "UPDATE PRD SET PRD.DtAlt = #" & Format(Date, "mm/dd/yyyy") & "# WHERE (((PRD.Venda)<> prd.preco))"
Banco.Execute SQL

If VarPzoProduto = True Then
    SQL = "UPDATE PRD SET Prd.PcoPrazo = prd.venda + (Prd.Venda * (Prd.MargemPrazo / 100)) WHERE (((PRD.Venda)<>[prd].[preco]))"
    Banco.Execute SQL
End If


End Sub

Private Sub OptOrdem_Click(Index As Integer)
DbcPesquisa_Change
End Sub

