VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FormConsultaPed 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCarne 
      Caption         =   "&Carnê"
      Height          =   360
      Left            =   12465
      TabIndex        =   46
      Top             =   6930
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   645
      Left            =   9090
      TabIndex        =   39
      Top             =   3735
      Width           =   4830
      Begin MSMask.MaskEdBox MebVista 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   40
         Tag             =   "S"
         Top             =   315
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
      Begin MSMask.MaskEdBox MebOrcamento 
         DataField       =   "Custo"
         Height          =   270
         Left            =   3360
         TabIndex        =   41
         Tag             =   "S"
         Top             =   315
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
         TabIndex        =   42
         Tag             =   "S"
         Top             =   315
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
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   3240
         X2              =   3240
         Y1              =   90
         Y2              =   630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1620
         X2              =   1620
         Y1              =   75
         Y2              =   630
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
         TabIndex        =   45
         Top             =   135
         Width           =   345
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Orçamento"
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
         Index           =   9
         Left            =   3300
         TabIndex        =   44
         Top             =   135
         Width           =   690
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
         Index           =   0
         Left            =   75
         TabIndex        =   43
         Top             =   135
         Width           =   315
      End
   End
   Begin VB.CommandButton CmdSat 
      Caption         =   "Cancela Nfc-e"
      Height          =   330
      Left            =   10665
      TabIndex        =   38
      Top             =   4950
      Width           =   1815
   End
   Begin VB.TextBox txtsat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   9090
      TabIndex        =   37
      Top             =   4545
      Width           =   4740
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   14160
      TabIndex        =   36
      Top             =   8730
      Width           =   14190
   End
   Begin VB.CheckBox ChkCartao 
      Caption         =   "Cartão"
      Height          =   285
      Left            =   8010
      TabIndex        =   35
      Top             =   8325
      Width           =   960
   End
   Begin VB.CommandButton BtReinpressao 
      Caption         =   "Impressão Orçamento"
      Height          =   375
      Index           =   2
      Left            =   7110
      TabIndex        =   34
      Top             =   6795
      Width           =   1860
   End
   Begin VB.TextBox Txtped 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6930
      TabIndex        =   31
      Top             =   7830
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9045
      TabIndex        =   29
      Top             =   6795
      Width           =   3345
      Begin MSMask.MaskEdBox MebTotal 
         Height          =   600
         Left            =   270
         TabIndex        =   30
         Top             =   855
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   1058
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483640
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   270
         TabIndex        =   32
         Top             =   450
         Width           =   645
      End
   End
   Begin VB.CommandButton Receb 
      Caption         =   "Confirma Recebimento"
      Height          =   645
      Left            =   12555
      TabIndex        =   33
      Top             =   7965
      Width           =   1410
   End
   Begin VB.TextBox TxtLinhaDigitavel 
      Height          =   330
      Left            =   945
      TabIndex        =   28
      Top             =   4995
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.TextBox TxtCodBarra 
      Height          =   330
      Left            =   945
      TabIndex        =   27
      Top             =   4590
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Data DatPcl 
      Caption         =   "Parcelas"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10530
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.Valor, DUP.Cond FROM DUP WHERE (((DUP.Pedido)=500)) ORDER BY DUP.NroPcl"
      Top             =   5580
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1395
      Left            =   270
      TabIndex        =   9
      Top             =   6750
      Width           =   6525
      Begin MSMask.MaskEdBox Entidade 
         DataField       =   "0"
         Height          =   270
         Left            =   1815
         TabIndex        =   10
         Top             =   180
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox EnderecoEntidade 
         Height          =   270
         Left            =   900
         TabIndex        =   11
         Top             =   465
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CidadeEntidade 
         Height          =   270
         Left            =   900
         TabIndex        =   12
         Top             =   1035
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebEntidade 
         Height          =   270
         Left            =   900
         TabIndex        =   13
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nro 
         Height          =   270
         Left            =   5445
         TabIndex        =   14
         Top             =   465
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Bairro 
         Height          =   270
         Left            =   900
         TabIndex        =   15
         Top             =   750
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox UfEntidade 
         Height          =   270
         Left            =   5445
         TabIndex        =   16
         Top             =   1035
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CepEntidade 
         Height          =   270
         Left            =   5445
         TabIndex        =   17
         Top             =   750
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Entidade"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         Height          =   195
         Left            =   5070
         TabIndex        =   21
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   345
         TabIndex        =   20
         Top             =   750
         Width           =   405
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Uf"
         Height          =   195
         Left            =   5190
         TabIndex        =   19
         Top             =   1095
         Width           =   165
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   5055
         TabIndex        =   18
         Top             =   795
         Width           =   285
      End
   End
   Begin VB.CommandButton Boleto 
      Caption         =   "&Boleto"
      Height          =   360
      Left            =   12510
      TabIndex        =   8
      Top             =   7515
      Width           =   1455
   End
   Begin VB.CommandButton BtReinpressao 
      Caption         =   "&Impressão do Cupom/Nfc-e"
      Height          =   360
      Index           =   1
      Left            =   5220
      TabIndex        =   7
      Top             =   8280
      Width           =   2265
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14190
      TabIndex        =   5
      Top             =   0
      Width           =   14190
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta Pedido"
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
         TabIndex        =   6
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.Data DatItens 
      Caption         =   "Itens"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3375
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4275
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton BtEditar 
      Caption         =   "&Editar Pedido"
      Height          =   360
      Left            =   7605
      TabIndex        =   3
      Top             =   7425
      Width           =   1365
   End
   Begin VB.CommandButton BtReinpressao 
      Caption         =   "&Reimpressão do Pedido"
      Height          =   360
      Index           =   0
      Left            =   3150
      TabIndex        =   2
      Top             =   8280
      Width           =   1950
   End
   Begin VB.CommandButton BtAtualiza 
      Caption         =   "&Atualiza Lista"
      Height          =   360
      Left            =   1905
      TabIndex        =   1
      Top             =   8280
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   8280
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   97583105
      CurrentDate     =   37501
   End
   Begin VB.Data DatPed 
      Caption         =   "Pedido"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1980
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSDBGrid.DBGrid Sit 
      Bindings        =   "FormConsultaPed.frx":0000
      Height          =   2880
      Left            =   270
      OleObjectBlob   =   "FormConsultaPed.frx":0017
      TabIndex        =   4
      Top             =   3825
      Width           =   8700
   End
   Begin MSDBGrid.DBGrid DBPcl 
      Bindings        =   "FormConsultaPed.frx":1AA6
      Height          =   1350
      Left            =   9045
      OleObjectBlob   =   "FormConsultaPed.frx":1ABB
      TabIndex        =   26
      Top             =   5355
      Width           =   4875
   End
   Begin Crystal.CrystalReport RptRel 
      Left            =   6930
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormConsultaPed.frx":29A0
      Height          =   2925
      Left            =   270
      OleObjectBlob   =   "FormConsultaPed.frx":29B5
      TabIndex        =   47
      Top             =   720
      Width           =   13650
   End
   Begin VB.Label Tipo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   25
      Top             =   6300
      Width           =   2310
   End
End
Attribute VB_Name = "FormConsultaPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NossoNumero As String
Dim CodigoBarras As String
Dim LinhaDigitavel As String
Public VarDocumento As String

Public VarImpostoNac As Currency
Public VarImpostoEst As Currency
Public VarTotImpostoNac As Currency
Public VarTotImpostoEst As Currency
Public VarCMun As String



Private Sub BtAtualiza_Click()
MebVista = 0
MebPrazo = 0
MebOrcamento = 0
             
If VarCaixaSeparado = False Then
              
    DatPed.RecordSource = "SELECT Ent.Matricula, Ped.DataEmiss, Ped.Obs, Ent.CodEntidade, Ped.CondPgto, PED.Cfo, PED.PrevEntrega, PED.Entidade, ENT.Nome, ENT.Endereco, ENT.Cep, ENT.Bairro, ENT.Fone, ENT.Cidade, ENT.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, PED.NroNt, PED.Operacao, ENT.Nro, ENT.Uf, Ped.Sat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.Operacao > 0 AND PED.Operacao <> 12 AND PED.DataEmiss =# " & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed DESC"
Else
'    If VarLanchonete = True Then
'        DatPed.RecordSource = "SELECT Ent.CodEntidade, Ped.CondPgto, PED.Cfo, PED.PrevEntrega, PED.Entidade, ENT.Nome, ENT.Endereco, ENT.Cep, ENT.Bairro, ENT.Fone, ENT.Cidade, ENT.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, PED.NroNt, PED.Operacao, ENT.Nro, ENT.Uf, Ped.Sat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.Operacao > 0 AND PED.Operacao <> 12 AND PED.DataEmiss =# " & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed DESC"
'    Else
        DatPed.RecordSource = "SELECT Ped.DataEmiss,  Ped.Obs, Ent.CodEntidade, Ped.CondPgto, PED.Cfo, PED.PrevEntrega, PED.Entidade, ENT.Nome, ENT.Endereco, ENT.Cep, ENT.Bairro, ENT.Fone, ENT.Cidade, ENT.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, PED.NroNt, PED.Operacao, ENT.Nro, ENT.Uf, Ped.Sat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.Operacao > 0 AND PED.Operacao <> 12 AND PED.DataEmiss =# " & Format(DtData, "mm/dd/yyyy") & "# and Ped.Contato = " & """" & FrmFrente.Maquina & """" & " ORDER BY PED.CodPed DESC"
'    End If
End If

DatPed.Refresh
If DatPed.Recordset.RecordCount <> 0 Then
     DatPed.Recordset.MoveFirst
     If VarCaixa Then
        While Not DatPed.Recordset.EOF
           Select Case DatPed.Recordset("Operacao")
           Case 1
                MebVista = MebVista + DatPed.Recordset("Total")
           Case 2
                MebPrazo = MebPrazo + DatPed.Recordset("Total")
           Case 4
                MebOrcamento = MebOrcamento + DatPed.Recordset("Total")
           End Select
           DatPed.Recordset.MoveNext
        Wend
    End If
    DatPed.Recordset.MoveFirst
End If


End Sub

Private Sub BtEditar_Click()
Dim Criterio As String
Dim VarDup As Recordset
Dim sql As String


If DBGrid.Columns(9) = "CANCELADO" Then
    MsgBox "Esse pedido não pode ser editado !!", vbInformation, App.Path
    Exit Sub
End If

Set VarDup = FrmFrente.Banco.OpenRecordset("Select * FROM Dup WHERE Pedido = " & DBGrid.Columns(0))
VarDup.Requery

Set FrmFrente.VarLista1 = FrmFrente.Banco.OpenRecordset("SELECT HISTEST.Nome, HISTEST.CodPrd, HISTEST.Valor, HISTEST.Qtd, HISTEST.Ordem FROM HISTEST WHERE HISTEST.Nome = " & """" & MebEntidade & """" & " ORDER BY HISTEST.CodPrd")
FrmFrente.VarLista1.Requery


If MsgBox("Confirma a edição desse pedido ", vbYesNo) = vbYes Then
    
   With FrmFrente
    .ContList = 0
    .DbcCliente.Enabled = True
       .DatCompra.RecordSource = "SELECT Prd.CodBar,  Itp.Entregue, Itp.Produto, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp , Itp.Loja FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & DBGrid.Columns(0) & " ORDER BY ITP.CodItp"
       .DatCompra.Refresh
        
        Open LocalBanco & "\Altera.txt" For Append As #1
        Print #1, "Frente; " & DBGrid.Columns(0) & "; " & DBGrid.Columns(4) & "; " & Date & "; " & Time & "; " & FrmFrente.Maquina & "; " & DBGrid.Columns(5) & "; " & Alinhar(DBGrid.Columns(2), 10, 2)
        Close #1
       
'       VarOrcamento = False
       If .DatCompra.Recordset.RecordCount > 0 Then
            
            Criterio = "CodPed = " & DBGrid.Columns(0)
            .Pedido.FindFirst Criterio
            If .Pedido.NoMatch Then
                MsgBox "Erro ao localizar pedido ! Sai do programa e edita de novo ", vbInformation, App.Title
                Exit Sub
            End If
            
            
            .MebTotal = 0
            If dir(App.Path & "\" & Trim(DBGrid.Columns(0)) & ".txt") <> "" Then
                Kill App.Path & "\" & Trim(DBGrid.Columns(0)) & ".txt"
            End If
            
            While Not .DatCompra.Recordset.EOF
                .MebTotal = .MebTotal + (CCur(.DatCompra.Recordset("Qtd")) * CCur(.DatCompra.Recordset("ValorUnit")))
                .MebTotal = Format(.MebTotal, "##,##0.00")
                If .DatCompra.Recordset("Entregue") = True Then
                    If .Pedido.Fields("Operacao") = 10 Then
'                        BaixarEstoque .DatCompra.Recordset.Fields("Qtd"), .DatCompra.Recordset.Fields("Produto"), .DatCompra.Recordset.Fields("Loja"), True
                        Open App.Path & "\" & DBGrid.Columns(0) & ".txt" For Append As #1
                        Print #1, psCompString(2, Numero2(.DatCompra.Recordset.Fields("Qtd")), 10) & psCompString(1, .DatCompra.Recordset.Fields("Produto"), 8) & psCompString(3, Trim(.DatCompra.Recordset.Fields("Loja")), 20) & "1"
                        Close #1
                        VarOrcaPrazo = False
                    Else
'                        BaixarEstoque .DatCompra.Recordset.Fields("Qtd"), .DatCompra.Recordset.Fields("Produto"), .DatCompra.Recordset.Fields("loja"), False
                        Open App.Path & "\" & DBGrid.Columns(0) & ".txt" For Append As #1
                        Print #1, psCompString(2, Numero2(.DatCompra.Recordset.Fields("Qtd")), 10) & psCompString(1, .DatCompra.Recordset.Fields("Produto"), 8) & psCompString(3, Trim(.DatCompra.Recordset.Fields("Loja")), 20) & "0"
                        Close #1
                        VarOrcaPrazo = True
                    
                    End If
                
                End If

                Set .LstPesq = .ListVda.ListItems.Add(, , .DatCompra.Recordset("CodItp"))
                .LstPesq.SubItems(1) = Format(.DatCompra.Recordset.Fields("Qtd"), " ##,##0.000")  'SeCampoNulo(Consulta, "Agencia")
                .LstPesq.SubItems(2) = .DatCompra.Recordset.Fields("Produto")
                .LstPesq.SubItems(3) = .DatCompra.Recordset.Fields("CodBar")
                .LstPesq.SubItems(4) = .DatCompra.Recordset("Descrição_Produto")    'SeCampoNulo(Consulta, "Agencia")
                .LstPesq.SubItems(5) = Format(.DatCompra.Recordset("ValorUnit"), " ##,##0.00")  'SeCampoNulo(Consulta, "Conta")
                .LstPesq.SubItems(6) = Format(Format(CCur(.DatCompra.Recordset("Qtd")), "##,##0.000") * Format(CCur(.DatCompra.Recordset("ValorUnit")), "##,##0.000"), "##,##0.00")  'SeCampoNulo(Consulta, "DataOper")
                .LstPesq.SubItems(7) = .DatCompra.Recordset("CodItp")
                .ContList = ContList + 1
                .DatCompra.Recordset.MoveNext
            Wend
            
            VarDtVenc = ""
            .Soma = .MebTotal
            .DatCompra.Recordset.MovePrevious
'            Criterio = "CodPed = " & DBGrid.Columns(0)
'            .Pedido.FindFirst Criterio
'            If Not .Pedido.NoMatch Then
            
'            End If
            
            .CodTotalizador = .DatCompra.Recordset("Pedido")
            .LblPedido = .CodTotalizador
            .LblItens = .DatCompra.Recordset("CodItp")
            .LblData = .Pedido("DataEmiss")
            .LblMaquina = .Maquina
            .CodCli = .Pedido("Entidade")
            Criterio = "CodEntidade = " & .Pedido("Entidade")
            .Clientes.FindFirst Criterio
'            .DtData = .Pedido("DataEmiss")
            If Not .Clientes.NoMatch Then
            End If
'            If .Clientes.RecordCount > 0 Then
                .LblCodigo = .CodCli
                .LblNome = .Clientes.Fields("Nome")
                Criterio = "CodEntidade = " & .Pedido("Entidade")
                .DatCli.Recordset.FindFirst Criterio
                .DbcCliente.Text = .DatCli.Recordset("Nome")
'            End If
            Criterio = "CodEntidade = " & .Pedido!Vendedor
            .DatVend.Recordset.FindFirst Criterio
            If Not .DatVend.Recordset.NoMatch Then
            End If
            
            If .Pedido.Fields("Operacao") = 1 Then
                Vista = True
            ElseIf .Pedido.Fields("Operacao") = 2 Then
                Vista = False
                     If .LblPedido <> "" Then
                        If .LblPedido > 0 Then
                            sql = "DELETE Dup.*, DUP.PEDIDO FROM DUP WHERE DUP.PEDIDO  = " & .LblPedido
                            .Banco.Execute sql
                         End If
                     End If

            End If
            
            If DBGrid.Columns(0) <> "" Then
                sql = "DELETE REC.*, REC.CodDupli FROM REC WHERE REC.CodDupli= " & DBGrid.Columns(0)
                FrmFrente.Banco.Execute sql
            End If
       End If
    End With
    Unload Me
    FrmFrente.TxtDados(2).Text = ""
    FrmFrente.TxtDados(0).SetFocus

End If
End Sub

Private Sub BtReinpressao_Click(Index As Integer)
Dim Retorno As Integer
Dim retorno2 As String
Dim sql As String
Dim Cupom As String
Dim Troco As Currency
Dim erro As Integer
Dim pag As String
Dim buf_ret As String * 100
Dim FormaPgto As String
Dim VarDupli As Recordset
Dim Contador As Integer

Dim Linha1 As String
Dim Linha2 As String
Dim Linha3 As String
Dim Linha4 As String
Dim Linha5 As String
Dim Linha6 As String
Dim Linha7 As String
Dim MensagemPromo As String
Dim Resp  As String
Dim Status As String
Dim TotalCupom As String
Dim Criterio As String
Dim Pago As Integer
Dim ConvertQtd As String
Dim ConvertVr As String
Dim CodigoBarra As String
Dim Dias As Integer
Dim x As Byte
Dim VarDesc As Currency
Dim VarValorNeg As Currency
Dim VarPedido As String
Dim NroCupom2 As String
Dim Varimpr As Byte

Dim varSubTotal As String
Dim VarTotal As Currency

Dim varSubTotal2 As String

Dim VarDtVenc As String
Dim Doc1 As String
Dim Doc2 As String
Dim VarOperacao As String
Dim ValorCupom As Currency
Dim MebNfe As String

Dim Linhas As Integer
Dim VarNroNota As String
Dim TxtOperacao
Dim VarDesconto As Currency
Dim VarDesconto2 As String
Dim TotalCusto As Currency
'Dim VarPedido As Double
Dim PercDesc As Double
Dim VarValordesc As Currency
'Dim VarDesc As Currency
Dim VarTotalGeral As String
Dim VarMonoRet As String
Dim VarTotMonoRet As String
Dim VarTotQtdMono As String

VarMonoRet = 0
VarTotMonoRet = 0
VarTotQtdMono = 0



VarTotalGeral = 0
Varimpr = Impressora

TotalCusto = 0

'On Error GoTo erro2

If Index = 0 Or Index = 2 Then
    Impressora = 99
End If

VarNroNota = 0

varSubTotal = 0
VarTotal = 0
Linhas = 0
varSubTotal2 = 0

Cupom = "0"

If Index = 1 And Impressora = 5 Then

     If MsgBox("Essa operação gera um novo cupom ! Deseja continuar ?", vbYesNo, App.Title) = vbYes Then
     Else
        Exit Sub
     End If
End If

If Index = 1 And Impressora = 5 And DBGrid.Columns(5) = "4" Then

     MsgBox "Operação não aceita", vbInformation
     Exit Sub
End If


If txtsat <> "" And Index = 1 Then

        If MsgBox("Confirma impressão ?", vbYesNo, App.Title) = vbYes Then
            AbreAcbr "Nfe.ImprimirDanfe(" & PortaAcbr & "\Arqs\" & VarCnpjEmite & "\Nfce\" & Year(DBGrid.Columns(1)) & psCompString(1, Month(DBGrid.Columns(1)), 2) & "\Nfce\" & txtsat & "-nfe.xml" & ")"
'            TrataAcbr
        End If

    Exit Sub
End If

    If Impressora <> 99 Then
        
'        If DBGrid.Columns(5) <> 1 And DBGrid.Columns(5) <> 2 Then
'            MsgBox "Operação não aceita na impressora fiscal", vbInformation, App.Title
'            Exit Sub
'        End If
            
    End If
    Select Case DBGrid.Columns(5)
        Case 1
            VarOperacao = DBGrid.Columns(5)
            TxtOperacao = "A VISTA"
        Case 2
            VarOperacao = 6
            TxtOperacao = "A PRAZO"
        Case 4
            VarOperacao = 4
            TxtOperacao = "ORCAMENTO"
        Case 13
            VarOperacao = 13
            TxtOperacao = "PEDIDO ENTREGA"
        Case Else
            TxtOperacao = "OUTROS"
    End Select
        
    VarPedido = DBGrid.Columns(0)
    If VarPedido = "" Then
         MsgBox "Nao tem pedido selecionado, clica em atualiza !", vbInformation
         Exit Sub
    End If
        
    If MsgBox("Confirma esta operação ?", vbYesNo) = vbYes Then
'        If MsgBox("Impressora pronta ?", vbYesNo, App.Title) = vbYes Then
'            DatPed.RecordSource = "SELECT PED.Cfo, PED.PrevEntrega, PED.Entidade, ENT.Nome, ENT.Endereco, ENT.Cep, ENT.Bairro, ENT.Fone, ENT.Cidade, ENT.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, PED.NroNt, PED.Operacao, ENT.Nro, ENT.Uf FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade  WHERE PED.CodPed =  " & Txtped
'           DatPed.RecordSource = "SELECT Ped.Cfo, Ped.PrevEntrega,PED.Entidade, ENT.Nome, ENT.Endereco, Ent.Cep, Ent.Bairro, Ent.Fone, ENT.Cidade, Ent.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, Ped.NroNt, ITP.Produto, ITP.Qtd, ITP.ValorUnit, Itp.Valor,  ITP.Complemento,PRD.Descrição_Produto, PRD.CodBar, PRD.Embalagem, PRD.Tributo, PRD.Aliquota, Ped.Operacao, Itp.CodItp, Ent.Nro, Ent.UF FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE PED.CodPed =  " & DBGrid.Columns(0) & " ORDER BY Itp.CodItp"
'MsgBox "ok 1 " & DatPed.Recordset.RecordCount, vbInformation
'           DatPed.Refresh
'MsgBox "ok 2 " & DatPed.Recordset.RecordCount, , vbInformation

'           sql = "SELECT Ped.Cfo, Ped.PrevEntrega,PED.Entidade, ENT.Nome, ENT.Endereco, Ent.Cep, Ent.Bairro, Ent.Fone, ENT.Cidade, Ent.Classificacao, ENT.RG, ENT.CPF, ENT.InscrEst, ENT.CGC, PED.DataEmiss, PED.Total, PED.Desconto, PED.CodPed, Ped.NroNt, ITP.Produto, ITP.Qtd, ITP.ValorUnit, Itp.Valor,  ITP.Complemento,PRD.Descrição_Produto, PRD.CodBar, PRD.Embalagem, PRD.Tributo, PRD.Aliquota, Ped.Operacao, Itp.CodItp, Ent.Nro, Ent.UF FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE PED.CodPed =  " & DBGrid.Columns(0) & " ORDER BY Itp.CodItp"
           If DatPed.Recordset.RecordCount < 0 Or DatItens.Recordset.RecordCount < 0 Then
                Exit Sub
           Else
'                DatPed.Recordset.MoveFirst
           End If
           
           VarPedido = DBGrid.Columns(0)
           
            VarDocumento = ""
            If Impressora <= 9 Then
                If MsgBox("Identificar cliente para nota fiscal paulista ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                    If DatPed.Recordset("Cpf") <> "" Then
                        FormCpf.MebDoc = DatPed.Recordset("Cpf")
                    End If
                    FormCpf.Show vbModal
                End If
           End If
           FrmFrente.CodCli = DatPed.Recordset("Entidade")
                                 
'MsgBox "ok 1", vbInformation
       If Index <= 1 Then
           Select Case Impressora
                Case 1
                Case 5
    '                    If VarDocumento = "" Then
                            Retorno = Bematech_FI_AbreCupom(VarDocumento)
    '                    Else
    '                    End If
                        'Mens.Caption = Retorno
                        NroCupom2 = Space(6)
                        Retorno = Bematech_FI_NumeroCupom(NroCupom2)
                        NroCupom = NroCupom2
                        VarTotImpostoNac = 0
                        VarTotImpostoEst = 0
                        
                                        
                    '    DatPed.Recordset.MoveFirst
                        
                        DatItens.Recordset.MoveFirst
                        While Not DatItens.Recordset.EOF
                        
                            Criterio = " Codigo = " & """" & Trim(LIMPACampo2(DatItens.Recordset.Fields("Classe"))) & """"
                            FrmFrente.DatMinutos.Recordset.FindFirst Criterio
                            If Not FrmFrente.DatMinutos.Recordset.NoMatch Then
                                VarImpostoNac = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqNac") / 100))
                                VarImpostoEst = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqEst") / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            Else
                                VarImpostoNac = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (VarMediaImpostoNac / 100))
                                VarImpostoEst = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (VarMediaImpostoes / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            End If
                                            
                            ConvertQtd = DatItens.Recordset.Fields("Qtd")
                            quantidade = DatItens.Recordset.Fields("Qtd")
                            
                            ConvertVr = Format(DatItens.Recordset.Fields("ValorUnit"), "##,##0.00")
                                                        
                            If Right(ConvertQtd, 3) = "000" Then
                                TipoQtd = "F"
                                ConvertQtd = Format(quantidade, "0000.000")
                            Else
                                TipoQtd = "F"
                                ConvertQtd = Format(quantidade, "0000.000")
                            End If
                            CodigoBarra = DatItens.Recordset("Produto")
                            
                            If VarMp4000 = True Then
                                ConvertVr = Format(ConvertVr, "0000.000")
                                Retorno = Bematech_FI_VendeItemArredondamentoMFD(psCompString(3, CodigoBarra, 13), Left(LimpaPrd, 29), _
                                      CodTrib(DatItens.Recordset.Fields("Tributo"), DatItens.Recordset("Aliquota")), Left(DatItens.Recordset.Fields("Embalagem"), 2), psCompString(1, ConvertQtd, 7), _
                                      psCompString(1, ConvertVr, 8), "0000", "0000", True)
                                    
                            
                            Else
                                ConvertVr = Format(ConvertVr, "00000.00")
                                Retorno = Bematech_FI_VendeItem(psCompString(3, CodigoBarra, 13), Left(LimpaPrd, 29), _
                                      CodTrib(DatItens.Recordset.Fields("Tributo"), DatItens.Recordset("Aliquota")), TipoQtd, psCompString(1, ConvertQtd, 7), 2, _
                                      psCompString(1, ConvertVr, 8), "%", "0000")
                            End If
                                  
    '                        Mens.Caption = Retorno
                                                                            
                            DatItens.Recordset.MoveNext
                                            
                        Wend
               '         DatPed.Recordset.MovePrevious
                        
                        VarDesconto = Format(DatPed.Recordset("Desconto"), "##,##0.00")
                        
                        VarDesconto = Format(VarDesconto, "00000.00")
                     
                        Retorno = Bematech_FI_IniciaFechamentoCupom("D", "$", psCompString(1, VarDesconto, 14))
                        FrmFrente.Mens = Retorno
                        Select Case VarOperacao
                            Case 1
                                FormaPgto = "Dinheiro"
                            Case 2
                                FormaPgto = "Cheque"
                            Case 3
                                FormaPgto = "Cheque Pré"
                            Case 4
                                FormaPgto = "Cartão"
                            Case 5
                                FormaPgto = "Ticket"
                            Case 6
                                FormaPgto = "A Prazo"
                        End Select
                        If ChkCartao = 1 Then
                            FormaPgto = "Cartão"
                        End If
                        varSubTotal = Format(DatPed.Recordset("Total"), "##,##0.00")
                        varSubTotal = Format(varSubTotal, "00000.00")
                        Retorno = Bematech_FI_EfetuaFormaPagamento(FormaPgto, psCompString(1, varSubTotal, 14))
                        FrmFrente.Mens = Retorno
                   
               Case Is < 10
                        If Impressora = 8 Then
                            MebNfe = LerINI("Nfce", "NroNfce", App.Path & "\config.ini")
                        End If

                        
                        ' AcbrMonitor
                        Dim VarContPrd As Integer
                        VarContPrd = 1
                        ' Venda do Item
                        If dir(App.Path & "\" & Trim(Left(LimpaCampo(FrmFrente.Maquina), 8)) & ".TXT") <> "" Then
                           Kill App.Path & "\" & Trim(Left(LimpaCampo(FrmFrente.Maquina), 8)) & ".TXT"
                        End If
                        If dir(App.Path & "\Cfe.Ini") <> "" Then
                           Kill App.Path & "\Cfe.ini"
                        End If
                        
                        If dir(PortaAcbr & "\Sai.txt") <> "" Then
                            Kill PortaAcbr & "\Sai.txt"
                        End If
                        VarImprSat = ""
                        VarTotImpostoNac = 0
                        VarTotImpostoEst = 0

                        
                        
                        
    '                    Open App.Path & "\Cfe.Ini" For Output As #1
    '                   InfCfe
                        If Impressora = 9 Then
                                EscreveINI "InfCfe", "Versao", VarVersao, App.Path & "\cfe.ini"
            '                   Identificacao
                                EscreveINI "Identificacao", "Cnpj", VarCnpjDesenvolvedor, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "signAC", VarChave, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "NumeroCaixa", VarCaixaSat, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "NCfe", DBGrid.Columns(0), App.Path & "\cfe.ini"
            '                   Emitente
                                EscreveINI "Emitente", "Cnpj", VarCnpjEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IE", VarInscEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IM", VarInscrMuni, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "indRatISSQN", "S", App.Path & "\cfe.ini"
                        Else
                        
                                Dim CodChave As Double
                                CodChave = Int(Mid(Int(CCur(Mid(Time, 7, 2)) * (Rnd * 100000)), 1, 6))
            '                   Identificacao
                                EscreveINI "Identificacao", "natOp", "Venda de Mercadoria", App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "mod", "65", App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "serie", VarSerieNfce, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "dEmi", Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS"), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "nNF", psCompString(1, MebNfe, 9), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "cNF", psCompString(1, CodChave, 9), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "tpNF", 1, App.Path & "\cfe.ini"
        '                        EscreveINI "Identificacao", "indPag", TipoPgto, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "indFinal", 1, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "indPres", 1, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "tpImp", 4, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "finNfe", 1, App.Path & "\cfe.ini"
            
            '                   Emitente
                                EscreveINI "Emitente", "Cnpj", VarCnpjEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IE", VarInscEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xNome", psCompString(3, VarRazaoEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xFant", psCompString(3, VarFantasiaEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "Fone", psCompString(3, VarTelEmite, 10), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "CEP", psCompString(1, VarCepEmite, 8), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xLgr", psCompString(3, VarEndEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "nro", psCompString(3, VarNroEmite, 10), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "cMun", Trim(VarIbgeEmite), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xMun", psCompString(3, VarCidadeEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xBairro", psCompString(3, VarBairroEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "UF", Left(VarUfEmite, 2), App.Path & "\cfe.ini"

                                
                                EscreveINI "Emitente", "CRT", 1, App.Path & "\cfe.ini"
                                                
                        End If
    '                   Destinatario
                        EscreveINI "Destinatario", "CNPJCPF", LIMPACampo2(VarDocumento), App.Path & "\cfe.ini"
                        EscreveINI "Destinatario", "XNome", Left(Entidade, 40), App.Path & "\cfe.ini"
                            EscreveINI "Destinatario", "indIEDest", "9", App.Path & "\cfe.ini"
                        
                        If Trim(Left(FrmFrente.Clientes.Fields("Nome"), 40)) <> "CONSUMIDOR" Then

        '                   Entrega
                            If EnderecoEntidade <> "" Then
                                EscreveINI "Entrega", "XLgr", RTrim(EnderecoEntidade), App.Path & "\cfe.ini"
    '                            EscreveINI "Entrega", "XLgr", RTrim(Nro), App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Entrega", "XLgr", "", App.Path & "\cfe.ini"
    '                            EscreveINI "Entrega", "XLgr", "", App.Path & "\cfe.ini"
                            End If
                            
                            If RTrim(Nro) <> "" Then
                                EscreveINI "Entrega", "nro", RTrim(Nro), App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Entrega", "nro", "", App.Path & "\cfe.ini"
                            End If
                            
                            
                            EscreveINI "Entrega", "XCpl", "", App.Path & "\cfe.ini"
                            If Bairro <> "" Then
                                EscreveINI "Entrega", "XBairro", RTrim(Bairro), App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Entrega", "XBairro", "", App.Path & "\cfe.ini"
                            End If
                            If CidadeEntidade <> "" Then
                                EscreveINI "Entrega", "XMun", RTrim(CidadeEntidade), App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Entrega", "XMun", "", App.Path & "\cfe.ini"
                            End If
                            If UfEntidade <> "" Then
                                EscreveINI "Entrega", "UF", RTrim(UfEntidade), App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Entrega", "UF", "", App.Path & "\cfe.ini"
                            End If
                            
                            If VarCMun <> "" Then
                                EscreveINI "Entrega", "cMun", RTrim(VarCMun), App.Path & "\cfe.ini"
                            Else
    '                            EscreveINI "Entrega", "UF", "", App.Path & "\cfe.ini"
                            End If
                            
                        End If
    '                   Produto
                        VarTotImposto = 0
                        
                        DatItens.Recordset.MoveFirst
                        While Not DatItens.Recordset.EOF
                            Criterio = "Produto = " & DatItens.Recordset("Produto")
                            FrmFrente.Registros.FindFirst Criterio
                            
                            Criterio = " Codigo = " & """" & Trim(LIMPACampo2(DatItens.Recordset.Fields("Classe"))) & """"
                            FrmFrente.DatMinutos.Recordset.FindFirst Criterio
                            If Not FrmFrente.DatMinutos.Recordset.NoMatch Then
                                VarImpostoNac = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqNac") / 100))
                                VarImpostoEst = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqEst") / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            Else
                                VarImpostoNac = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (VarMediaImpostoNac / 100))
                                VarImpostoEst = ((DatItens.Recordset.Fields("Qtd") * DatItens.Recordset.Fields("ValorUnit")) * (VarMediaImpostoEst / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            End If
                                                                                    
                            ConvertQtd = Format(DatItens.Recordset.Fields("Qtd"), "0.000")
                            quantidade = DatItens.Recordset.Fields("Qtd")
                            
                            ConvertVr = Format(DatItens.Recordset.Fields("ValorUnit"), "0.000")
                            varSubTotal2 = varSubTotal2 + Arredondamento((ConvertQtd * ConvertVr), 2)
                            If IsNull(FrmFrente.Registros("BarBarra")) Then
                                CodigoBarra = 1
                            Else
                                CodigoBarra = Left(FrmFrente.Registros("BarBarra"), 13)
                            End If
                            
'                                ConvertVr = Format(ConvertVr, "00000.00")
'                                Retorno = Bematech_FI_VendeItem(psCompString(3, CodigoBarra, 13), Left(LimpaPrd, 29), _
'                                      CodTrib(FrmFrente.DatCompra.Recordset.Fields("Tributo"), FrmFrente.DatCompra.Recordset("Aliquota")), TipoQtd, psCompString(1, ConvertQtd, 7), 2, _
'                                      psCompString(1, ConvertVr, 8), "%", "0000")

                           If DBGrid.Columns(3) <> 0 Then
                                PercDesc = (CCur(DBGrid.Columns(3)) / CCur(CCur(DBGrid.Columns(3)) + CCur(DBGrid.Columns(2))))
                                VarValordesc = ConvertVr * ConvertQtd
                                VarDesc = Format((VarValordesc * PercDesc), "##,##0.0000")
                            Else
                                VarDesc = 0
                            End If

                                  
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cProd", DatItens.Recordset("Produto"), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "xProd", Left(DatItens.Recordset("Descrição_Produto"), 29), App.Path & "\cfe.ini"
'                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEan", IIf(Len(CodigoBarra) >= 13, psCompString(3, CodigoBarra, 13), ""), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEan", "", App.Path & "\cfe.ini"
                            
'                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", "", App.Path & "\cfe.ini"
                            
                            If DatItens.Recordset("Classe") = "" Or Len(Trim(DatItens.Recordset("Classe"))) < 8 Then
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", "54023390", App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", Trim(LimpaNro(DatItens.Recordset("Classe"))), App.Path & "\cfe.ini"
                            End If
                            
                           
'                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", DatItens.Recordset("CfopPrd"), App.Path & "\cfe.ini"
                            
                            If DatItens.Recordset("CfopPrd") <> "5405" Then
                                If DatItens.Recordset("CfopPrd") <> "5102" Then
                                    If DatItens.Recordset("CfopPrd") <> "0" Then
                                        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", DatItens.Recordset("CfopPrd"), App.Path & "\cfe.ini"
                                    Else
                                        EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", "5102", App.Path & "\cfe.ini"
                                    End If
                                Else
                                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", DatItens.Recordset("CfopPrd"), App.Path & "\cfe.ini"
                                End If
                            Else
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", DatItens.Recordset("CfopPrd"), App.Path & "\cfe.ini"
                            End If
                                                        
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "uCom", IIf(DatItens.Recordset("Embalagem") = "", "UN", DatItens.Recordset("Embalagem")), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "Combustivel", 0, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "qCom", ConvertQtd, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", ConvertVr, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "indRegra", "A", App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vDesc", CCur(VarDesc), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vOutro", 0, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vItem23741", CCur(VarImpostoNac + VarImpostoEst), App.Path & "\cfe.ini"
                            If Impressora <> 9 Then
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vprod", ConvertVr * ConvertQtd, App.Path & "\cfe.ini"
                            End If
                            
                             If DatItens.Recordset("CfopPrd") = "5656" Or DatItens.Recordset("CfopPrd") = "5.656" Or DatItens.Recordset("CfopPrd") = "5655" Or DatItens.Recordset("CfopPrd") = "5.655" Or DatItens.Recordset("CfopPrd") = "5661" Or DatItens.Recordset("CfopPrd") = "6661" Then
                                If psCompString(3, Trim(DatItens.Recordset("Classe")), 8) = "27111910" Then
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "210203001", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "GLP", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGLP", "40.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNn", "30.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNi", "30.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "vPart", "5.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                                    VarMonoRet = 1.4139
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right("061", 2), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "qBCMonoRet", Numero2(DatItens.Recordset.Fields("Qtd")), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "adRemICMSRet", Numero2(VarMonoRet), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSMonoRet", Numero2(DatItens.Recordset.Fields("Qtd") * VarMonoRet), App.Path & "\cfe.ini"
                                    VarTotMonoRet = VarTotMonoRet + DatItens.Recordset.Fields("Qtd") * VarMonoRet
                                    VarTotQtdMono = VarTotQtdMono + DatItens.Recordset.Fields("Qtd")
                                                     
                                    
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "indImport", "0", App.Path & "\cfe.ini"
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "cUFOrig", "35", App.Path & "\cfe.ini"
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "pOrig", "100.00", App.Path & "\cfe.ini"
                                    
                                    
                                    
                                End If
                                
                                If psCompString(3, Trim(DatItens.Recordset("Classe")), 8) = "27101932" Then
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "620505001", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "OUTROS OLEOS LUBRIFICANTES AUTOMOTIVOS", App.Path & "\cfe.ini"
                    '                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGLP", "40.00", App.Path & "\cfe.ini"
                    '                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNn", "30.00", App.Path & "\cfe.ini"
                    '                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNi", "30.00", App.Path & "\cfe.ini"
                    '                EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "vPart", "5.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                                End If
                           
                            
                            End If
    '                       ObsFisco
                            EscreveINI "ObsFiscoDet" & psCompString(1, VarContPrd, 3) & psCompString(1, VarContPrd, 3), "xCampoDet", "", App.Path & "\cfe.ini"
                            EscreveINI "ObsFiscoDet" & psCompString(1, VarContPrd, 3) & psCompString(1, VarContPrd, 3), "xTextoDet", "", App.Path & "\cfe.ini"
    '                       Icms
                            EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "Orig", 0, App.Path & "\cfe.ini"
'                            If DatItens.Recordset("SitTrib") = "000" Then
'                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "00", App.Path & "\cfe.ini"
'                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", DatItens.Recordset("Icms"), App.Path & "\cfe.ini"
'                            Else
'                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "60", App.Path & "\cfe.ini"
'                            End If


                            If VarRegime = 3 Then
                                If DatItens.Recordset("SitTrib") = "000" Then
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "00", App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", DatItens.Recordset("Icm"), App.Path & "\cfe.ini"
                                
                                Else
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "60", App.Path & "\cfe.ini"
    '                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", "500", App.Path & "\cfe.ini"
                                End If
                                
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", "01", App.Path & "\cfe.ini"
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "vBC", ConvertVr, App.Path & "\cfe.ini"
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "pPIS", VarAliqPis / 100, App.Path & "\cfe.ini"
        '                       Cofins
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", "01", App.Path & "\cfe.ini"
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "vBC", ConvertVr, App.Path & "\cfe.ini"
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "pCOFINS", VarAliqCofins / 100, App.Path & "\cfe.ini"
                            
                            
                            
                            Else
                                If DatItens.Recordset("SitTrib") = "102" Or DatItens.Recordset("SitTrib") = "000" Then
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", "102", App.Path & "\cfe.ini"
    '                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", DatItens.Recordset("Icms"), App.Path & "\cfe.ini"
                                Else
    '                                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "60", App.Path & "\cfe.ini"
                                    
                                    
                                    If psCompString(3, Trim(DatItens.Recordset("Classe")), 8) = "27111910" Then
                                    
   ''                                     EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right("061", 2), App.Path & "\cfe.ini"
    ''                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "qBCMonoRet", Numero2(0), App.Path & "\cfe.ini"
   ''                                     EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "adRemICMSRet", Numero2(0), App.Path & "\cfe.ini"
   ''                                     EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSMonoRet", Numero2(0), App.Path & "\cfe.ini"
                                    Else
                                        EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", DatItens.Recordset("SitTrib"), App.Path & "\cfe.ini"
                                    End If
                                End If
                            
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
        '                       Cofins
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
                            
                            
                            End If
                            

'            If DatItens.Recordset("SitTrib") = "102" Then
'                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", "102", App.Path & "\cfe.ini"
'            Else
'                EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", "500", App.Path & "\cfe.ini"
'            End If

                                       
'            EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
'            EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
                                         
                            VarTotalGeral = VarTotalGeral + (ConvertQtd * ConvertVr)
        
                            DatItens.Recordset.MoveNext
                            VarContPrd = VarContPrd + 1
                        Wend
    
                        
'                        MebDesconto = Format(MebDesconto, "00000.00")
                     
                        Select Case VarOperacao
                            Case 1
                                FormaPgto = "Dinheiro"
                            Case 2
                                FormaPgto = "Cheque"
                            Case 3
                                FormaPgto = "Cheque Pré"
                            Case 4
                                FormaPgto = "Cartão"
                            Case 5
                                FormaPgto = "Ticket"
                            Case 6
                                FormaPgto = "A Prazo"
                        End Select
                        
                        If ChkCartao = 1 Then
                            FormaPgto = "Cartão"
                        End If
                        
                        varSubTotal = Format(DatPed.Recordset("Total"), "##,##0.00")
                                                                                        
    '                    Total
                         EscreveINI "Total", "vCFeLei12741", CCur(VarTotImpostoNac + VarTotImpostoEst), App.Path & "\cfe.ini"
                         EscreveINI "DescAcrEntr", "VDescSubTotal", DBGrid.Columns(3), App.Path & "\cfe.ini"
                                 
                         If Impressora <> 9 Then
                            EscreveINI "TOTAL", "vProd", VarTotalGeral, App.Path & "\cfe.ini"
                            EscreveINI "TOTAL", "vDesc", DBGrid.Columns(3), App.Path & "\cfe.ini"
                            EscreveINI "TOTAL", "vnf", VarTotalGeral - DBGrid.Columns(3), App.Path & "\cfe.ini"
                            EscreveINI "TRANSPORTADOR", "modFrete", 9, App.Path & "\cfe.ini"
                            If VarTotMonoRet <> 0 Then
                                EscreveINI "TOTAL", "vICMSMonoRet", Numero2(VarTotMonoRet), App.Path & "\cfe.ini"
                                EscreveINI "TOTAL", "qBCMonoRet", Numero2(VarTotQtdMono), App.Path & "\cfe.ini"
                            End If
                            

                         End If
                                
                         If Impressora <> 9 Then
                         
                            EscreveINI "Pag001", "tPag", psCompString(1, VarOperacao, 2), App.Path & "\cfe.ini"
                            If VarOperacao = 3 Or VarOperacao = 4 Then
                                EscreveINI "Pag001", "tpIntegra", psCompString(1, "2", 1), App.Path & "\cfe.ini"
                            End If


                            If TotPgto < varSubTotal2 Then
                               EscreveINI "Pag001", "vPag", Format(VarTotalGeral - DBGrid.Columns(3), "####00.00"), App.Path & "\cfe.ini"
                            Else
                               EscreveINI "Pag001", "vPag", Format(VarTotalGeral - DBGrid.Columns(3), "####00.00"), App.Path & "\cfe.ini"
                            End If
                                                
                                
        '                        TxtImposto = "Trib Aprox R$ " & VarTotImpostoNac & " Federal e R$ " & Format(VarTotImpostoEst, "##0.00") & " Estadual Fonte:IBPT"
                         Else
                             EscreveINI "Pagto001", "cMP", psCompString(1, VarOperacao, 2), App.Path & "\cfe.ini"
                         
                            If TotPgto < varSubTotal2 Then
                               EscreveINI "Pagto001", "vMP", Format(VarTotalGeral, "####00.00"), App.Path & "\cfe.ini"
                            Else
                               EscreveINI "Pagto001", "vMp", Format(VarTotalGeral, "####00.00"), App.Path & "\cfe.ini"
                            End If
                            
                             
 '                            If MebComprove <> "" Then
 '                               EscreveINI "Pag002", "tPag", psCompString(1, 6, 2), App.Path & "\cfe.ini"
 ''                              EscreveINI "Pag002", "vPag", Format(MebComprove, "####00.00"), App.Path & "\cfe.ini"
  '                          End If
                        
                         End If
                         
                         TxtImposto = "Trib Aprox R$ " & VarTotImpostoNac & " Federal e R$ " & Format(VarTotImpostoEst, "##0.00") & " Estadual Fonte:IBPT"
                         
                         EscreveINI "DadosAdicionais", "InfCpl", TxtImposto & " " & Trim(TxtObs) & " " & Trim(TxtObs2), App.Path & "\cfe.ini"
                         EscreveINI "ObsFiscoDet001", "xCampo", "", App.Path & "\cfe.ini"
                         EscreveINI "ObsFiscoDet001", "xTexto", "", App.Path & "\cfe.ini"
                         
                         
                                                     
            '        AbreAcbr "SAT.ImprimirExtratoVenda(" & VarImprSat & ")"
            '        TrataAcbr
                        If Impressora = 9 Then
                                AbreAcbr "SAT.CriarEnviarCfe(" & App.Path & "\cfe.ini" & ")"
                                TrataAcbr
                        
                                If ResultSat = "6000" Then
                                    FrmTotal.ImprimirSat
                                End If
                        Else
                               If MsgBox("Imprimir Cupom ?", vbYesNo, App.Title) = vbYes Then
                                    AbreAcbr "NFE.CriarEnviarNFe(" & App.Path & "\cfe.ini" & " , 1, 1, 1,,0)"
                               Else
                                    AbreAcbr "NFE.CriarEnviarNFe(" & App.Path & "\cfe.ini" & " , 1, 0, 1,,0)"
                               End If
            
                               If TrataAcbr2("Nota", MebNfe) = 100 Then
                                   MsgBox "NFc-e enviada com sucesso !", vbInformation, App.Title
                                   
                                   sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.Cfo = 'Nfc-e', Ped.PrevEntrega = ' ', Ped.NroNt = " & CCur(MebNfe) & "  WHERE PED.CodPed = " & psCompString(1, DatPed.Recordset("CodPed"), 6)
                                   FrmFrente.Banco.Execute sql

                                   EscreveINI "Nfce", "NroNfce", CCur(MebNfe) + 1, App.Path & "\config.ini"

                               Else
                                   MsgBox "ATENÇÃO !!!! ERRO NO ENVIO DA NFC-e !", vbInformation, App.Title
                                   
                                                              
                               End If
                        
                        
                        End If
                                                                                                                       
                
                Case 99
                    
                    Set Empresa = FrmFrente.Banco.OpenRecordset("CFG")
                    Empresa.MoveFirst
                    Iniciar (False)
                    varSubTotal = 0
                    
                    If Left(PortaNF, 3) <> "COM" Then
                        If VarImpressoraNf = "TERMICA" Then
                             Print #1, ""
'                             Fim (False)
                             Print #1, "" '+ Cond
                             Print #1, Centralizar(FrmFrente.Empresa("Cabecalho1"), 40)
                             Print #1, Centralizar(FrmFrente.Empresa("Cabecalho2"), 40)
                             Print #1, "========================================"
                             Print #1, "Data:" & DatPed.Recordset("DataEmiss") & " Hora:" & Time; Tab(31); "Ped:"; psCompString(1, DatPed.Recordset("CodPed"), 6)
                             Print #1, Tab(1); psCompString(1, DatPed.Recordset("Entidade"), 6); "-"; Left(DatPed.Recordset("Nome"), 30)
                             If Trim(Left(DatPed.Recordset("Endereco"), 35)) <> "" Then
                                Print #1, Tab(1); Trim(Left(DatPed.Recordset("Endereco"), 35)) & ", " & DatPed.Recordset("Nro")
                             End If
                             If Trim(Left(DatPed.Recordset("Bairro"), 35)) <> "" Then
                                Print #1, Tab(1); Trim(Left(DatPed.Recordset("Bairro"), 35))
                             End If
                             If Trim(Left(DatPed.Recordset("Cidade"), 30)) <> "" Then
                                Print #1, Tab(1); Trim(Left(DatPed.Recordset("Cidade"), 30))
                             End If
                             Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
                             Print #1, Tab(1); "  QTD"; Tab(20); "UNIT."; Tab(35); "TOTAL"
                             Print #1, "----------------------------------------"
                             
                       '      DatPed.Recordset.MoveFirst
                             VarNroNota = DatPed.Recordset("NroNt")
                             DatItens.Recordset.MoveFirst
                             While Not DatItens.Recordset.EOF
                                If Len(Trim(DatItens.Recordset("Complemento"))) > 0 Then
                                    Print #1, Tab(1); psCompString(1, DatItens.Recordset("Produto"), 6); "-"; Left(DatItens.Recordset("Descrição_Produto") & "-" & Trim(DatItens.Recordset("Complemento")), 33)
                                Else
                                    Print #1, Tab(1); psCompString(1, DatItens.Recordset("Produto"), 6); "-"; Left(DatItens.Recordset("Descrição_Produto"), 33)
                                End If
                                Print #1, Tab(2); Alinhar(Format(DatItens.Recordset("Qtd"), " ##,##0.000"), 7, 3); Tab(13); "X"; Tab(16); Alinhar(Format(DatItens.Recordset("ValorUnit"), "##,##0.00"), 9, 2); Tab(31); Alinhar(Format(DatItens.Recordset("Qtd") * DatItens.Recordset("ValorUnit"), "##,##0.00"), 10, 2)
                                varSubTotal = varSubTotal + DatItens.Recordset("Valor")

                                DatItens.Recordset.MoveNext
                             Wend
              '               DatPed.Recordset.MovePrevious
                             
'                             Print #1, ""

                             Print #1, Tab(18); "SUB-TOTAL =>"; Tab(31); Alinhar(CStr(varSubTotal), 10, 2)
                             Print #1, Tab(18); "DESCONTO  =>"; Tab(31); Alinhar(CStr(DatPed.Recordset("Desconto")), 10, 2)
                             Print #1, Tab(18); "TOTAL     =>"; Tab(31); Alinhar(CStr(varSubTotal - DatPed.Recordset("Desconto")), 10, 2)
 '                            Print #1, ""
                             
                             If VarOperacao = "1" Then
                                    Print #1, Tab(18); "RECEBIDO ==>"; Tab(31); Alinhar(CStr(DatPed.Recordset("Total")), 10, 2)
'                                    Print #1, ""
                             End If
                             Print #1, Tab(1); TxtOperacao
                             
                            If Not IsNull(DatPed.Recordset("Obs")) Then
                                Print #1, ""
                                Print #1, Tab(1); Left(DatPed.Recordset("Obs"), 35)
                                If Len(DatPed.Recordset("Obs")) > 35 Then
                                    Print #1, Tab(1); Mid$(DatPed.Recordset("Obs"), 35, 35)
                                End If
                                
                            End If
'                            If Not IsNull(TxtObs2) And TxtObs <> "" Then
'                                Print #1, Tab(1); Left(TxtObs2, 35)
'                            End If
                             
                             
 '                            Print #1, ""
                             Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 40)
                             Print #1, ""

                             
                        End If
                        
                        If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Then
                             Print #1, "" & Cond
                             Print #1, ""
                             Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
                             Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 48)
                             Print #1, "================================================"
                             Print #1, "Data:" & Date & " Hora:" & Time; Tab(33); "Pedido.:"; psCompString(1, DatPed.Recordset("CodPed"), 6)
                             Print #1, Tab(1); psCompString(1, DatPed.Recordset("Entidade"), 6); "-"; Left(DatPed.Recordset("Nome"), 30)
                             Print #1, Tab(1); Trim(Left(DatPed.Recordset("Endereco"), 35)) & ", " & DatPed.Recordset("Nro")
                             Print #1, Tab(1); Trim(Left(DatPed.Recordset("Bairro"), 35))
                             Print #1, Tab(1); Trim(Left(DatPed.Recordset("Cidade"), 30))
                             
                             Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
                             Print #1, Tab(1); "QTD"; Tab(20); "UNIT."; Tab(40); "TOTAL"
                             Print #1, "------------------------------------------------"
                             
                          '   DatPed.Recordset.MoveFirst
                             VarNroNota = DatPed.Recordset("NroNt")
                             DatItens.Recordset.MoveFirst
                             Contador = 0
                             
                             While Not DatItens.Recordset.EOF
                                If Len(Trim(DatItens.Recordset("Complemento"))) > 0 Then
                             
                                    Print #1, Tab(1); psCompString(1, DatItens.Recordset("Produto"), 6); "-"; Left(DatItens.Recordset("Descrição_Produto") & "-" & Trim(DatItens.Recordset("Complemento")), 39)
                                 Else
                                    Print #1, Tab(1); psCompString(1, DatItens.Recordset("Produto"), 6); "-"; Left(DatItens.Recordset("Descrição_Produto"), 39)
                                 End If
                                 
                                 Print #1, Tab(3); Alinhar(Format(DatItens.Recordset("Qtd"), " ##,##0.000"), 7, 3); Tab(13); "X"; Tab(18); Alinhar(Format(DatItens.Recordset("ValorUnit"), "##,##0.00"), 9, 2); Tab(37); Alinhar(Format(DatItens.Recordset("Qtd") * DatItens.Recordset("ValorUnit"), "##,##0.00"), 10, 2)
                                 varSubTotal = varSubTotal + DatItens.Recordset("Valor")

                                 DatItens.Recordset.MoveNext
                                 Contador = Contador + 1
                             Wend
'                             DatPed.Recordset.MovePrevious
                                                          
                             Print #1, ""
                             Print #1, Tab(5); Contador & "  Itens"; Tab(20); "SUB-TOTAL =>"; Tab(37); Alinhar(CStr(varSubTotal), 10, 2)
                             Print #1, Tab(20); "DESCONTO  =>"; Tab(37); Alinhar(CStr(DatPed.Recordset("Desconto")), 10, 2)
                             Print #1, Tab(20); "TOTAL     =>"; Tab(37); Alinhar(CStr(varSubTotal - DatPed.Recordset("Desconto")), 10, 2)
                             Print #1, ""
                             
                             If VarOperacao = "1" Then
                                    Print #1, Tab(20); "RECEBIDO ==>"; Tab(37); Alinhar(CStr(DatPed.Recordset("Total")), 10, 2)
                                    Print #1, ""
                             End If
                             Print #1, Tab(1); TxtOperacao
                             
                             Print #1, ""
                             Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 48)
                             Print #1, "" & Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
                             Print #1, Chr(29) & Chr(102) & Chr(0) & vbCrLf
                             Print #1, ""

                            
                        End If
                        
                        
                        If VarImpressoraNf = "EPSON" Then
                        
                            If VarMeiaFolha = "Sim" Then
                                If PortaNF = "TELA" Then
                                     Fim (False)
                                
                                    Dim RegPed As Recordset
                                    Dim Sql2 As String
    
                             '       DatPed.Recordset.MoveFirst
                                    
    '                                 Set Pedidos = Banco.OpenRecordset("SELECT * From Ped Where CodPed = " & LblTot)
    '                                 Set Parcelas = Banco.OpenRecordset("SELECT * From Dup Where Pedido = " & TxtDados(0))
                                     Sql2 = "SELECT Ped.Ipi, Ped.PedSub,Ent.CodEntidade, ENT.Nome, Ent.Fantasia, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor, Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob , PED.ValorFrete, Ped.Desconto, ITP.CodItp, Ped.NroPedido, Prd.complemento FROM PRD INNER JOIN (ENT INNER JOIN (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON ENT.CodEntidade = PED.Entidade) ON PRD.CodPrd = ITP.Produto "
                                     Sql2 = Sql2 & " GROUP BY Ped.ipi, Ped.PedSub, Ent.CodEntidade, ENT.Nome, Ent.Fantasia, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor , Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob ,PED.ValorFrete,Ped.Desconto,PED.NroPedido,ITP.CodItp, Prd.complemento "
                                     Sql2 = Sql2 & " HAVING (((PED.CodPed)= " & DatPed.Recordset("CodPed") & ")) ORDER BY ITP.CodItp"
                                     Set varPesquisa(1) = FrmFrente.Banco.OpenRecordset(Sql2)
                                     
                                                     
                                     
    '                                 Criterio = "CodPed = " & TxtDados(0).Text
    '                                 Pedidos.FindFirst Criterio
                                     
                                     Set FrmRelPed.Relatorio.Recordset = varPesquisa(1)
                                     
                                     
                                     FrmRelPed.rpCabTitulo.Caption = FrmFrente.Empresa("Cabecalho1")
                                     FrmRelPed.rpCabInformacao.Caption = FrmFrente.Empresa("Cabecalho2")
                                    
                                     If DBGrid.Columns(5) = "7" Then
                                        FrmRelPed.ReportField10.Caption = "ORÇAMENTO"
                                     End If
                                     If DBGrid.Columns(5) = "9" Then
                                        FrmRelPed.ReportField10.Caption = "DEVOLUÇÃO"
                                     End If

                                     
    '                                 If VarPedCompleto = True Then
                                         FrmRelPed.CabGrupo(1).Mostrar = True
    '                                 Else
    '                                     FormPed.CabGrupo(1).Mostrar = False
    '                                 End If
                                     
    '                                 If Not IsNull(Pedidos!Obs) Then
    '                                     FrmRelPed.Obs1.Caption = (Left(Pedidos("Obs"), 70))
    '                                     FrmRelPed.Obs2.Caption = (Mid(Pedidos("Obs"), 71, 70))
    '                                     FrmRelPed.Obs3.Caption = (Mid(Pedidos("Obs"), 141, 70))
    '                                     FrmRelPed.Obs4.Caption = (Mid(Pedidos("Obs"), 211, 70))
    '                                 End If
                                    Me.Hide
                                    FrmRelPed.Relatorio.Copias = VarCopias
                                    FrmRelPed.Config
                                    Me.Show vbModal
                                                                                                                                
                                Else
                            
                                '     DatPed.Recordset.MoveFirst
                                     
                                     Print #1, Left(FrmFrente.Empresa("Cabecalho1"), 40); Tab(65); "Pedido: "; psCompString(1, DatPed.Recordset("CodPed"), 6)
                                     Print #1, Left(FrmFrente.Empresa("Cabecalho2"), 80); Tab(60); Date & " - " & Time
                                     Print #1, ""
                                     Print #1, psCompString(1, DatPed.Recordset("Entidade"), 6) & "-" & Left(DatPed.Recordset("Nome"), 40)
                                     Print #1, DatPed.Recordset("Endereco") & " " & DatPed.Recordset("Nro")
                                     Print #1, DatPed.Recordset("Cidade") & "  " & DatPed.Recordset("Uf") & "  " & DatPed.Recordset("Bairro") & "    " & Format(DatPed.Recordset("Cep"), "@@.@@@-@@")
                                     Print #1, "Fone:"; Tab(7); Format(DatPed.Recordset("Fone"), "(@@)@@@@-@@@@");
                                     If DatPed.Recordset("Classificacao") = 2 Then
                                            Doc1 = Format(DatPed.Recordset("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                                            Doc2 = DatPed.Recordset("InscrEst")
                                            Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                                            Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
                                        Else
                                            Doc1 = Format(DatPed.Recordset("Rg"), "@@.@@@.@@@-@")
                                            Doc2 = Format(DatPed.Recordset("Cpf"), "@@@.@@@.@@@.@@")
                                            Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                                            Print #1, Tab(52); "CPF:"; Tab(62); Doc2
                                     End If
                                     Print #1, ""
                                     Print #1, "--------------------------------------------------------------------------------"
                                     Print #1, "CODIGO"; Tab(10); "QTD"; Tab(17); "UN"; Tab(25); "PRODUTO"; Tab(68); "UNIT."; Tab(76); "TOTAL"
                                     Print #1, "--------------------------------------------------------------------------------"
                                     Linhas = 11
'                                     FrmFrente.DatCompra.Recordset.MoveFirst
                                     DatItens.Recordset.MoveFirst
                                     Do While Not DatItens.Recordset.EOF
                                        Print #1, Tab(1); psCompString(1, Trim(DatItens.Recordset("Produto")), 6); Tab(10); Alinhar(DatItens.Recordset("Qtd"), 6, 0); Tab(17); DatItens.Recordset("Embalagem"); Tab(20); Left(DatItens.Recordset("Descrição_Produto"), 40); " "; Left(DatItens.Recordset("Complemento"), 20);
                                        Print #1, Tab(64); Alinhar(DatItens.Recordset("ValorUnit"), 8, 2); Tab(72); Alinhar(DatItens.Recordset("Valor"), 9, 2)
                                        Linhas = Linhas + 1
                                        varSubTotal = varSubTotal + DatPed.Recordset("Total")
'                                        If TxtPgto <> "7" Then
'                                            If TxtPgto = "9" Then
'                                                BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto"), False
'                                            Else
'                                                BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto"), True
'                                            End If
'                                        End If
                                        If Linhas >= 30 Then
                                                Print #1, ""
                                                Print #1, ""
                                                Print #1, ""
                                                Print #1, ""
                                                Print #1, ""
                                                
                                                Linhas = 2
                                        End If
                                        DatItens.Recordset.MoveNext
                                     Loop
                                
                                     VarTotal = varSubTotal
                                     If Linhas > 50 Then
                                        SaltoPagina Linhas
                                     End If
                     '                DatPed.Recordset.MovePrevious
                                     
                                     Print #1, ""
                                     Print #1, Tab(55); " SUB-TOTAL =>   "; Alinhar(DatPed.Recordset("Total"), 10, 2)
                                     Print #1, Tab(57); "DESCONTO =>   "; Alinhar(DatPed.Recordset("Desconto"), 10, 2)
                                     Print #1, Tab(60); "TOTAL =>   "; Alinhar(DatPed.Recordset("Total") - DatPed.Recordset("Desconto"), 10, 2)
                                     Linhas = Linhas + 2
                                                                      
                                     If Linhas > 50 Then
                                        SaltoPagina Linhas
                                     End If
                                 
                                     Print #1, Tab(1); TxtOperacao
                                     Linhas = Linhas + 1
                                                                           
                                     If VarOperacao = "6" Then
                                            Set VarDupli = FrmFrente.Banco.OpenRecordset("SELECT DTVENC, PEDIDO FROM DUP WHERE PEDIDO  = " & DatPed.Recordset("CodPed"))
                                            If VarDupli.RecordCount > 0 Then
                                                VarDtVenc = VarDupli!DtVenc
                                            Else
                                                VarDtVenc = ""
                                            End If
                                            If VarDtVenc <> "" Then
                                                Print #1, ""
                                                Print #1, Tab(1); "DATA VENCIMENTO - "; VarDtVenc
                                                Linhas = Linhas + 2
                                            End If
                                            If Linhas > 50 Then
                                               SaltoPagina Linhas
                                            End If
                                            Print #1, ""
                                            Print #1, ""
                                            Print #1, " Ass.Cli-----------------------------------------------"
                                            Linhas = Linhas + 3
                                      End If
                                      If Linhas <= 50 Then
                                        SaltoPagina Linhas
                                      End If
'                                      Fim (False)
                                 End If
                             Else
                                     Print #1, ""
                                     Print #1, "==============================================================================="
                                     Print #1, Gde & Centralizar(Left(FrmFrente.Empresa("Cabecalho1"), 40), 40) & Nor
                                     Print #1, Centralizar(Left(FrmFrente.Empresa("Cabecalho2"), 80), 80)
                                     Print #1, "==============================================================================="
                                     Print #1, "Data:"; Date; Tab(17); " - "; Time; Tab(66); "Pedido: "; psCompString(1, DatPed.Recordset("CodPed"), 6)
                                     Print #1, "-------------------------------------------------------------------------------"
                                     Print #1, psCompString(1, DatPed.Recordset("Entidade"), 6) & "-" & Left(DatPed.Recordset("Nome"), 40)
                                     Print #1, DatPed.Recordset("Endereco")
                                     Print #1, DatPed.Recordset("Cidade")
        '                             Print #1, "Fone:"; Tab(7); Format(FrmFrente.Clientes("Fone"), "(@@)@@@@-@@@@");
                                     If DatPed.Recordset("Classificacao") = 2 Then
                                            Doc1 = Format(DatPed.Recordset("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                                            Doc2 = DatPed.Recordset("InscrEst")
                                            Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                                            Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
                                        Else
                                            Doc1 = Format(DatPed.Recordset("Rg"), "@@.@@@.@@@-@")
                                            Doc2 = Format(DatPed.Recordset("Cpf"), "@@@.@@@.@@@.@@")
                                            Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                                            Print #1, Tab(52); "CPF:"; Tab(62); Doc2
                                     End If
                                     Print #1, "-------------------------------------------------------------------------------"
                                     Print #1, "CODIGO"; Tab(8); "PRODUTO"; Tab(44); "  "; Tab(53); "QTD"; Tab(61); "UNIT."; Tab(71); "TOTAL"
                                     Print #1, "-------------------------------------------------------------------------------"
                                     
                                  '   DatPed.Recordset.MoveFirst
                                     VarNroNota = DatPed.Recordset("NroNt")
                                     
                                     DatItens.Recordset.MoveFirst
                                     
                                     Do While Not DatItens.Recordset.EOF
                                        Print #1, Tab(1); psCompString(1, Trim(DatItens.Recordset("Produto")), 6); Tab(8); Left(DatItens.Recordset("Descrição_Produto"), 38); Tab(48); Alinhar(DatItens.Recordset("Qtd"), 8, 2); Tab(58); Alinhar(DatItens.Recordset("ValorUnit"), 9, 2); Tab(67); Alinhar(DatItens.Recordset("Qtd") * DatItens.Recordset("ValorUnit"), 9, 2)
                                        Linhas = Linhas + 1
        '                                varSubTotal = varSubTotal + DatPed.Recordset("Valor")
                                        DatItens.Recordset.MoveNext
                                     Loop
                                     
  '                                   DatItens.Recordset.MovePrevious
                                
                                     VarTotal = varSubTotal
                                     Print #1, ""
                                     Print #1, Tab(50); "SUB-TOTAL =>"; Tab(66); Alinhar(CStr(DatPed.Recordset("Total")), 10, 2)
                                     Print #1, Tab(50); "DESCONTO  =>"; Tab(66); Alinhar(CStr(DatPed.Recordset("Desconto")), 10, 2)
                                     Print #1, Tab(50); "TOTAL     =>"; Tab(66); Alinhar(CStr(DatPed.Recordset("Total") - DatPed.Recordset("Desconto")), 10, 2)
                                     Print #1, ""
                                     
                                     If VarOperacao = "1" Then
                                                Print #1, Tab(50); "RECEBIDO ==>"; Tab(66); Alinhar(CStr(DatPed.Recordset("Total") - DatPed.Recordset("Desconto")), 10, 2)
                                                Print #1, ""
                                     End If
                                     Print #1, Tab(1); TxtOperacao
                                      
                                      Print #1, ""
                                      Print #1, " Recebi os produtos acima - Ass.Cli ----------------------------------"
                                     
                                      Print #1, ""
                                      Print #1, ""
                                      Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 80)
                                      Print #1, ""
                                      Print #1, ""
                                      If VarOperacao = "6" Then
                                            Set VarDupli = FrmFrente.Banco.OpenRecordset("SELECT DTVENC, PEDIDO FROM DUP WHERE PEDIDO  = " & DatPed.Recordset("CodPed"))
                                            If VarDupli.RecordCount > 0 Then
                                                VarDtVenc = VarDupli!DtVenc
                                            Else
                                                VarDtVenc = ""
                                            End If
                                      
                                            Print #1, "==============================================================================="
                                            Print #1, Gde & VarEmpresa & Nor
                                            Print #1, VarEnd; Tab(40); VarCidade
                                            Print #1, "Insc.Est.: " & VarInscEst; Tab(30); "CNPJ: " & VarCGC
                                            Print #1, "==============================================================================="
                                            Print #1, "         FATURA           |       DUPLICATA         |         DATA             "
                                            Print #1, "   Valor     |   Numero   |   Valor     |  Numero   | Emissão     | Vencimento "
                                            Print #1, "-------------------------------------------------------------------------------"
'                                            Print #1, "             |            |"; Tab(30); Alinhar(CStr(MebTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, LblTot, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
                                            If VarNroNota = 0 Then
                                                Print #1, "             |            |"; Tab(30); Alinhar(CStr(DatPed.Recordset("Total") - DatPed.Recordset("Desconto")), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, DatPed.Recordset("Total") - DatPed.Recordset("Desconto"), 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
                                            Else
                                                Print #1, Tab(2); Alinhar(CStr(DatPed.Recordset("Total") - DatPed.Recordset("Desconto")), 10, 2); Tab(14); "|"; Tab(17); psCompString(1, VarNroNota, 6); Tab(27); "|"; Tab(30); Alinhar(CStr(DatPed.Recordset("Total") - DatPed.Recordset("Desconto")), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, VarNroNota, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
                                            End If
                                            Print #1, "-------------------------------------------------------------------------------"
                                            Print #1, " Nome do Sacado.: "; DatPed.Recordset("Nome")
                                            Print #1, " Endereco.......: "; Trim(DatPed.Recordset("Endereco")) & ", " & DatPed.Recordset("Nro")
                                            Print #1, " Municipio......: "; Left(DatPed.Recordset("Cidade"), 20) & " - " & Left(DatPed.Recordset("Uf"), 2) & "  " & Format(DatPed.Recordset("Cep"), "@@.@@@-@@")
                                            Print #1, " CNPJ/CPF.......: "; Doc1; Tab(45); "Insc.Est./RG.: "; Doc2
                                            Print #1, "-------------------------------------------------------------------------------"
                                            Print #1, " Valor por Extenso: "; Cond & UCase(PassaExtenso(DatPed.Recordset("Total") - DatPed.Recordset("Desconto"))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *" & Nor
                                            Print #1, "-------------------------------------------------------------------------------"
                                            Print #1, " Reconheco/emos a exatidao desta DUPLICATA de VENDA MERCANTIL, na importancia  "
                                            Print #1, " acima que pagarei/emos a " & Negi & VarEmpresa & NegF
                                            Print #1, " ou a sua ordem na praca e vencimento acima indicados."
                                            Print #1, " "
                                            Print #1, "           ____/____/_______            _______________________________"
                                            Print #1, "            Data do aceite                  Assinatura do Sacado       "
                                      End If
                                      Print #1, SaltoPag
                                
                       '               Fim (False)
                       '               Close #1
                            End If
                        End If
                    End If


 '                   Fim (False)
 '                   PortaNF = "TELA"
                    If VarOperacao = "6" And (PortaNF <> "TELA" And VarImpressoraNf <> "EPSON") Then
                        If MsgBox("Imprimir Comprovante ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
 '                           Iniciar (False)
                    
                            If Left(PortaNF, 3) <> "COM" Then
                            
                                Print #1, ""
                                Print #1, "======================================="
                                Print #1, " Pedido:"; Tab(10); psCompString(1, DatPed.Recordset("CodPed"), 6); Tab(18); "Data: "; Tab(24); Date
                                Print #1, ""
                                Print #1, Tab(1); psCompString(1, DatPed.Recordset("Entidade"), 6); Tab(10); "-"; Tab(12); Left(DatPed.Recordset("Nome"), 28)
                                Print #1, ""
                                Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(DatPed.Recordset("Total")), 10, 2)
                                Print #1, ""
                                Print #1, ""
                                Print #1, "   Ass.Cli---------------------------"
                                Print #1, "======================================="
                                Print #1, ""

                            End If
                            
 '                           Fim (False)
                        Else
                           If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Or VarImpressoraNf = "TERMICA" Then
 '                               Iniciar (False)
                                    Print #1, ""
                                    Print #1, "   Ass.Cli---------------------------"
                                    Print #1, "======================================="
                                    Print #1, ""
 '                               Fim (False)
                            End If
                        End If
                    End If
 '                   Iniciar (False)
                    If Left(PortaNF, 3) <> "COM" Then
                         If VarMeiaFolha = "Nao" Then
                    
                            Print #1, "" & Nor
                             For x = 1 To NroLinha
                                Print #1, ""
                            Next
                         End If
                   Else
                        For x = 1 To NroLinha
                            Comando = FormataTX("" + Chr(13) + Chr(10), 2, 0, 0, 0, 0)
                        Next
                    End If
                    If PortaNF <> "TELA" Then
                        Print #1, "" & Chr(27) & "i"         '& Chr(27) & "v50"
                    End If
                    
                    Fim (False)
                
            End Select
            
'            MensagemPromo = "           Agradecemos a Preferencia            "
            If IsNull(FrmFrente.Empresa("MensagemRodape")) Then
                MensagemPromo = "           Agradecemos a Preferencia            "
            Else
                MensagemPromo = psCompString(3, Centralizar(FrmFrente.Empresa("MensagemRodape"), 48), 48)
            End If
            If Impressora <> 99 Then
'                TxtImposto = "Trib Aprox R$ " & Format(VarTotImpostoNac, "#,##0.00") & " Federal e " & Format(VarTotImpostoEst, "##0.00") & "%) Fonte: IBPT"
'                TxtImposto = psCompString(3, TxtImposto, 48)
                TxtImposto = "Trib Aprox R$ " & Format(VarTotImpostoNac, "#,##0.00") & " Federal e R$ " & Format(VarImpostoEst, "##0.00") & " Estadual Fonte:IBPT"
                TxtImposto = psCompString(3, TxtImposto, 48)
            
            End If
            If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") <> "" Then
                If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") <> 3 Then
                    MensagemPromo = "ICMS RECOLHIDO CONF.LC 123/2006-SIMPLES NACIONAL"
                End If
            End If
                    
'            PortaNF = "TELA"
            If Impressora <= 8 Then
                If VarOperacao = "6" And MsgBox("Imprimir comprovante ?", vbYesNo, App.Title) = vbYes Then
                    Linha1 = "================================================"
                    Linha2 = " Data: " & Space(5) & Date & "Pedido: " & psCompString(1, VarPedido, 6) & Space(12)
                    Linha3 = psCompString(1, DatPed.Recordset("Entidade"), 6) & "-" & Left(DatPed.Recordset("Nome"), 40) + Space(48 - Len(psCompString(1, DatPed.Recordset("Entidade"), 6) & "-" & Left(DatPed.Recordset("Nome"), 40)))
                    Linha4 = " Valor ==> " & Alinhar(CStr(DatPed.Recordset("Total")), 10, 2) & Space(27)
                    Linha5 = Space(48)
                    Linha6 = "          Ass.Cli-------------------------------"
                    Linha7 = Space(48)
                Else
                    Linha4 = ""
                    If FrmFrente.CodCli = 1 Then
                        Linha1 = ""
                        Linha2 = ""
                        Linha3 = ""
                        Linha4 = ""
                        Linha5 = ""
                        Linha6 = ""
                        Linha7 = ""
                    Else
                        
                        Linha1 = RTrim(DatPed.Recordset("Nome")) + Space(48 - Len(RTrim(DatPed.Recordset("Nome"))))
                        
                        If DatPed.Recordset.Fields("Classificacao") = 1 Then

                            If IsNull(DatPed.Recordset("Cpf")) Then
                                  Linha2 = ""
                            Else
                                  Linha2 = RTrim(Format(DatPed.Recordset("Cpf"), "@@@.@@@.@@@-@@")) + Space(48 - Len(RTrim(Format(DatPed.Recordset("Cpf"), "@@@.@@@.@@@-@@"))))
                            End If

                            If IsNull(DatPed.Recordset("Rg")) Then
                                  Linha3 = ""
                            Else
                                  Linha3 = RTrim(DatPed.Recordset("Rg")) + Space(48 - Len(RTrim(DatPed.Recordset("Rg"))))
                            End If

                         Else
                            If IsNull(DatPed.Recordset("Cgc")) Then
                                  Linha2 = ""
                            Else
                                  Linha2 = RTrim(Format(DatPed.Recordset("Cgc"), "@@.@@@.@@@/@@@@-@@")) + Space(48 - Len(RTrim(Format(DatPed.Recordset("Cgc"), "@@.@@@.@@@/@@@@-@@"))))
                            End If
                            If IsNull(DatPed.Recordset("InscrEst")) Then
                                  Linha3 = ""
                            Else
                                  Linha3 = RTrim(DatPed.Recordset("InscrEst")) + Space(48 - Len(RTrim(DatPed.Recordset("InscrEst"))))
                            End If
                            
                        End If
                                                                       
                        If IsNull(DatPed.Recordset("Endereco")) Then
                              Linha4 = ""
                        Else
                              Linha4 = RTrim(DatPed.Recordset("Endereco")) & ", " & DatPed.Recordset("Nro") + Space(48 - Len(RTrim(DatPed.Recordset("Endereco") + ", " + DatPed.Recordset("Nro"))))
                        End If

                            
                        
                        If IsNull(DatPed.Recordset("Cidade")) Then
                              Linha5 = ""
                        Else
                              Linha5 = RTrim(DatPed.Recordset("Cidade")) & "-" & DatPed.Recordset("Uf") + Space(48 - Len(RTrim(DatPed.Recordset("Cidade") & "-" & DatPed.Recordset("Uf"))))
                        End If

                        Linha6 = ""
                        Linha7 = ""
                    
                    End If
                End If
            End If
            Select Case Impressora
                Case 1
                    ' Fecha o Cupom
                    If VarOperacao = "6" Then
'                            If MsgBox("Imprimir Comprovante ?", vbYesNo, App.Title) = vbYes Then
                            Retorno = FECHACUPOMFISCAL("S" & psCompString(1, Len(MensagemPromo + Linha1 + Linha2 + Linha3 + Linha4 + Linha5 + Linha6 + Linha7), 3), _
                                    MensagemPromo & Linha1 & Linha2 & Linha3 & Linha4 & Linha5 & Linha6 & Linha7)
'                            Else
'                                Retorno = FECHACUPOMFISCAL("S" & psCompString(1, Len(MensagemPromo + Space(48) + Linha1 + Linha2 + Linha3 + Linha4), 3), _
'                                        MensagemPromo & Space(48) & Linha1 & Linha2 & Linha3 & Linha4)
'                            End If
                    Else
                            Retorno = FECHACUPOMFISCAL("S" & psCompString(1, Len(MensagemPromo + Space(48) + Linha1 + Linha2 + Linha3 + Linha4 + Linha5 + Linha6), 3), _
                                    MensagemPromo & Space(48) & Linha1 & Linha2 & Linha3 & Linha4 & Linha5 & Linha6)
                    End If
                    MensErro (Retorno)
                    ' Abrir Gaveta
                    Retorno = AcionarGaveta() ' "0", "32", "38")
                    MensErro (Retorno)
                Case 5
                    
                    If VarOperacao = "6" Then
                        TxtMensagem = (TxtImposto & Left(Linha1, 48) & Left(Linha2, 48) & Left(Linha4, 48) & Left(Linha6, 48))
                    Else
                        TxtMensagem = (TxtImposto & Left(Linha1, 48) & Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha5, 48))
                    End If
                    Retorno = Bematech_FI_TerminaFechamentoCupom(TxtMensagem)
                    'FrmFrente.Mens.Caption = Retorno
                                        
                    ' Abrir Gaveta
                    Retorno = Bematech_FI_AcionaGaveta()
                    'FrmFrente.Mens.Caption = Retorno
                Case 6
                    If VarOperacao = "6" Then
                        TxtMensagem = (Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha4, 48) & Left(Linha6, 48))
                    Else
                        TxtMensagem = (Left(Linha1, 48) & Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha5, 48))
                    End If
                    Retorno = Daruma_FI_TerminaFechamentoCupom(TxtMensagem)
                    'FrmFrente.Mens.Caption = Retorno
                                        
                    ' Abrir Gaveta
                    Retorno = Daruma_FI_AcionaGaveta()
                    'FrmFrente.Mens.Caption = Retorno
                                        
                Case 7
                        If txtpgto >= 6 Then
                            If Comprove = True Then
                                TxtMensagem = (Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha4, 48) & Left(Linha6, 48))
                            Else
                                TxtMensagem = (Left(Linha1, 48) & Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha4, 48))
                            End If
                        Else
                            TxtMensagem = (Left(Linha1, 48) & Left(Linha2, 48) & Left(Linha3, 48) & Left(Linha4, 48))
                        End If
                        Retorno = Elgin.Elgin_TerminaFechamentoCupom(TxtMensagem)
                        TrataRetorno2 Retorno
                                                                                    
            End Select
            NroCupom = MebNfe
            If NroCupom = "" Then NroCupom = 0
          
            If Impressora < 9 And IsNumeric(NroCupom) Then
                sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.NroNt = " & NroCupom & " WHERE PED.CodPed = " & VarPedido
                FrmFrente.Banco.Execute sql
            End If
            
     ElseIf Index = 2 Then
            Set Empresa = FrmFrente.Banco.OpenRecordset("CFG")
            Empresa.MoveFirst
            Iniciar (False)
            varSubTotal = 0

     
             Print #1, Cond & Left(FrmFrente.Empresa("Cabecalho1"), 40); Tab(109); "Pedido: "; psCompString(1, DatPed.Recordset("CodPed"), 6)
             Print #1, Left(FrmFrente.Empresa("Cabecalho2"), 80); Tab(103); Date & " - " & Time
             Print #1, ""
             Print #1, Nor & Negi & psCompString(1, DatPed.Recordset("Entidade"), 6) & "-" & Left(DatPed.Recordset("Nome"), 40) & NegF & Cond
             Print #1, DatPed.Recordset("Endereco") & " " & DatPed.Recordset("Nro")
             Print #1, DatPed.Recordset("Cidade") & "  " & DatPed.Recordset("Uf") & "  " & DatPed.Recordset("Bairro") & "    " & Format(DatPed.Recordset("Cep"), "@@.@@@-@@")
             Print #1, "Fone:"; Tab(7); Format(DatPed.Recordset("Fone"), "(@@)@@@@-@@@@");
             If DatPed.Recordset("Classificacao") = 2 Then
                    Doc1 = Format(DatPed.Recordset("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                    Doc2 = DatPed.Recordset("InscrEst")
                    Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                    Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
                Else
                    Doc1 = Format(DatPed.Recordset("Rg"), "@@.@@@.@@@-@")
                    Doc2 = Format(DatPed.Recordset("Cpf"), "@@@.@@@.@@@.@@")
                    Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                    Print #1, Tab(52); "CPF:"; Tab(62); Doc2
             End If
             Print #1, ""
             Print #1, "---------------------------------------------------------------------------------------------------------------------------"
             Print #1, "CODIGO"; Tab(15); "QTD"; Tab(20); "UN"; Tab(25); "PRODUTO"; Tab(75); "VR.CUSTO"; Tab(88); "UNIT."; ; Tab(98); "CUSTO"; Tab(110); "TOTAL"; Tab(119); "( % )"
             Print #1, "---------------------------------------------------------------------------------------------------------------------------"
             Linhas = 11
'                                     FrmFrente.DatCompra.Recordset.MoveFirst
             DatItens.Recordset.MoveFirst
             Do While Not DatItens.Recordset.EOF
                Print #1, Tab(1); psCompString(1, Trim(DatItens.Recordset("Produto")), 6); Tab(10); Alinhar(DatItens.Recordset("Qtd"), 8, 2); Tab(20); DatItens.Recordset("Embalagem"); Tab(25); Left(DatItens.Recordset("Descrição_Produto"), 45); Tab(75); Alinhar(DatItens.Recordset("CustoDia"), 9, 2); Tab(84); Alinhar(DatItens.Recordset("ValorUnit"), 9, 2); Tab(95); Alinhar(Format(DatItens.Recordset("CustoDia") * DatItens.Recordset("Qtd"), "#,##0.00"), 9, 2); Tab(107); Alinhar(DatItens.Recordset("Valor"), 9, 2); Tab(118); Alinhar(Format(((DatItens.Recordset("Valor") / (DatItens.Recordset("CustoDia") * DatItens.Recordset("Qtd"))) - 1) * 100, "#,##0.00"), 6, 2)
                Linhas = Linhas + 1
                varSubTotal = varSubTotal + DatItens.Recordset("Valor")
                TotalCusto = TotalCusto + DatItens.Recordset("CustoDia") * DatItens.Recordset("Qtd")
                DatItens.Recordset.MoveNext
             Loop
        
             VarTotal = varSubTotal
             Print #1, "---------------------------------------------------------------------------------------------------------------------------"
'             Print #1, ""
             Print #1, Tab(95); Alinhar(Format(TotalCusto, "#,##0.00"), 9, 2); Tab(107); Alinhar(Format(varSubTotal, "#,##0.00"), 9, 2); Tab(118); Alinhar(Format(((varSubTotal / TotalCusto) - 1) * 100, "#,##0.00"), 6, 2)
'             'Print #1, Tab(74); "DESCONTO =>  "; Alinhar(DatPed.Recordset("Desconto"), 10, 2);
'             Print #1, Tab(103); "TOTAL =>   "; Alinhar(DatPed.Recordset("Total") - DatPed.Recordset("Desconto"), 10, 2)
             Print #1, SaltoPag
             Linhas = Linhas + 2
             Fim (False)
             
                                              
'             If Linhas > 30 Then
'                SaltoPagina Linhas
'             End If

             If VarImprEntrega = "C:\TELA.TXT" Then
                Unload Me
'                FrmRelTXT.Relatorio.Orientacao = oriPaisagem
                FrmRelTXT2.Config "C:\TELA.Txt"
             End If

            
        End If
        BtAtualiza_Click
        VarOperacao = ""
        FrmFrente.CodCli = 0
     
'     End If
    End If
Impressora = Varimpr

Exit Sub
erro2:
If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If

End Sub


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdCarne_Click()
    If DBGrid.Columns(9) = "PARCELADO" Then
        FrmParcelas.Tag = "P"
        FrmParcelas.CodDoc = DBGrid.Columns(0)
        FrmParcelas.CodEnt = MebEntidade
        FrmParcelas.MebVlrPed = Format(DBGrid.Columns(2), "#,##0.00")
        FrmParcelas.TxtNome = psCompString(1, MebEntidade, 6) & "-" & Entidade.Text
        FrmParcelas.OK.Enabled = False
        FrmParcelas.GridParc.Enabled = False
        Me.Hide
        FrmParcelas.Show vbModal
        Me.Show
    Else
        MsgBox "Natureza da Operação Inválida", vbInformation, App.Title
    End If

End Sub

Private Sub CmdSat_Click()

If txtsat <> "" Then
    If MsgBox("Confirma Cancelamento do Nfc-e Nro: " & txtsat, vbYesNo) = vbYes Then
        If LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini") <> "" Then
             VarCnpjEmite = LerINI("EMITENTE", "Cnpj", App.Path & "\nfe.ini")
        End If

        If dir(PortaAcbr & "\Sai.txt") <> "" Then
            Kill PortaAcbr & "\Sai.txt"
        End If

       AbreAcbr "Nfe.CANCELARNFE(" & txtsat & "," & "Cupom emitido com erro de digitacao" & "," & VarCnpjEmite & ")"
       VarResp = TrataAcbr2("Cancela", 1)
       If VarResp = 135 Or VarResp = 155 Then
                MsgBox "Cupom cancelado com sucesso !", vbInformation, App.Title
'                DatPed.Recordset.Edit
'                DatPed.Recordset("NroNt") = 0
'                DatPed.Recordset("Cfo") = ""
'                DatPed.Recordset.Update
                sql = "UPDATE PED SET  PED.NroNt = 0  , PED.Cfo = ' '  WHERE PED.CodPed = " & DBGrid.Columns(0)
                FrmFrente.Banco.Execute sql
       
       
       End If
  
'       AbreAcbr "SAT.CancelarCFe(" & """" & PortaAcbr & "\Arqs\SAT\Vendas\" & VarCnpjEmite & "\" & Year(Date) & psCompString(1, Month(Date), 2) & "\" & Trim(txtsat) & ".xml" & """" & ")"
'       TrataAcbr
        
'       If ResultSat = "7000" Then
'              sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.Sat = " & """" & CStr(Mid(Trim(txtsat), 60, 44)) & """" & " , PED.Cfo = 'SAT', PED.PrevEntrega = 'CANCELADO', Ped.NroNt = 0  WHERE PED.CodPed = " & DBGrid.Columns(0)
'              FrmFrente.Banco.Execute sql
'        End If
        
    End If
End If


End Sub

Private Sub DBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub


Private Sub DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DBGrid.Columns(0) <> "" Then
    DatItens.RecordSource = ("SELECT Prd.Icm, Prd.CfopPrd, Prd.SitTrib, Prd.Tributo, Prd.Aliquota, Prd.embalagem, itp.complemento, ITP.Pedido, ITP.Produto, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.CodItp, Itp.CustoDia, Classe FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE ITP.Pedido =   " & DBGrid.Columns(0) & " ORDER BY ITP.CodItp ")
    DatItens.Refresh
    
    Entidade = ""
    EnderecoEntidade = ""
    Nro = ""
    Bairro = ""
    CidadeEntidade = ""
    UfEntidade = ""
    CepEntidade = ""
    VarDocumento = ""
    VarCMun = ""
    
    If DatItens.Recordset.RecordCount <> 0 Then
         DatItens.Recordset.MoveFirst
    End If
    If Not IsNull(DatPed.Recordset.Fields("CondPgto")) Then
        Tipo = DatPed.Recordset.Fields("CondPgto")
    End If
    MebEntidade = DatPed.Recordset.Fields("CodEntidade")
    Entidade = DatPed.Recordset.Fields("Nome")
    If Not IsNull(DatPed.Recordset.Fields("Endereco")) Then
        EnderecoEntidade = DatPed.Recordset.Fields("Endereco")
    End If
    If Not IsNull(DatPed.Recordset.Fields("Nro")) Then
    
        Nro = DatPed.Recordset.Fields("Nro")
    End If
    If Not IsNull(DatPed.Recordset.Fields("Bairro")) Then
    
        Bairro = DatPed.Recordset.Fields("Bairro")
    End If
    If Not IsNull(DatPed.Recordset.Fields("Cidade")) Then
    
        CidadeEntidade = DatPed.Recordset.Fields("Cidade")
    End If
    If Not IsNull(DatPed.Recordset.Fields("Uf")) Then
    
        UfEntidade = DatPed.Recordset.Fields("Uf")
    End If
    If Not IsNull(DatPed.Recordset.Fields("Cep")) Then
        CepEntidade = DatPed.Recordset.Fields("Cep")
    End If
    
    If Not IsNull(DatPed.Recordset.Fields("Matricula")) Then
        VarCMun = DatPed.Recordset.Fields("Matricula")
    End If
    
    If DatPed.Recordset.Fields("Classificacao") = "1" Then
        VarDocumento = Format(DatPed.Recordset.Fields("Cpf"), "@@@.@@@.@@@-@@")
    Else
        VarDocumento = Format(DatPed.Recordset.Fields("CGC"), "@@.@@@.@@@/@@@@-@@")
    End If
    
    MebTotal = DatPed.Recordset.Fields("Total") '- DatPed.Recordset.Fields("Desconto")
    Txtped.Text = DBGrid.Columns(0)
    If Boleto.Enabled = True Then
        DatPcl.RecordSource = "SELECT dup.DtEmiss, dup.coddup, dup.nossonumero, dup.CodigoBarra, LinhaDigitalizavel, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.Valor, DUP.Cond FROM DUP WHERE DUP.Pedido = " & DBGrid.Columns(0) & " ORDER BY DUP.NroPcl"
        DatPcl.Refresh
        If DatPcl.Recordset.RecordCount > 0 Then
            DatPcl.Recordset.MoveFirst
        End If
    Else
    
        DatPcl.RecordSource = "SELECT DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.Valor, DUP.Cond FROM DUP WHERE DUP.Pedido = " & DBGrid.Columns(0) & " ORDER BY DUP.NroPcl"
        DatPcl.Refresh
    End If
    
    If VarCaixa Then
        Perc = 0
        MebDesconto = 0
        Total = 0
        MebAcresc = 0
'        If Not IsNull(DatPed.Recordset.Fields("Desconto")) Then
'            MebDesconto = DatPed.Recordset.Fields("Desconto")
'        End If
        Total = (MebTotal - MebDesconto) + MebAcresc
'        Perc.SetFocus
    End If
    If Not IsNull(DatPed.Recordset("SAt")) Then
        txtsat = DatPed.Recordset("Sat")
    Else
        txtsat = ""
    End If
End If
End Sub

Private Sub Form_Activate()
BtAtualiza_Click

End Sub

Private Sub Form_Load()

DatPed.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"

DatItens.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"

DatPcl.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"


DtData.Value = Date

If EditarFrente = True Then  'Or CfgDesc = 1 Then
'    BtEditar.Visible = True
Else
    BtEditar.Visible = False
End If

If Impressora = 99 Then
    BtReinpressao(1).Visible = False
End If

If Impressora = 10 Then
    BtReinpressao(1).Visible = False
End If

If VarBanco = "" Then
    Boleto.Enabled = False
    Receb.Enabled = False
    Frame2.Enabled = False
End If

Perc = 0
MebDesconto = 0
Total = 0
MebAcresc = 0



End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Function LimpaPrd() As String
Dim Produto As String
Dim PrdSujo As String
Dim x As Integer
Produto = ""
PrdSujo = Left(DatItens.Recordset("Descrição_Produto"), 38)
For x = 1 To Len(PrdSujo)
    If Mid$(PrdSujo, x, 1) >= Chr(32) And Mid$(PrdSujo, x, 1) <= Chr(122) Then
            Produto = Produto + Mid$(PrdSujo, x, 1)
    End If
Next
LimpaPrd = psCompString(3, Produto, 38)
End Function


Private Sub Boleto_Click()
Dim TESTE As String
Dim Registro As String
Dim VarEmitente As String
Dim VarEndEmitente As String
Dim VarNroemitente As String
Dim VarCidadeEmitente As String
Dim VarUfEmitente As String


    If DatPcl.Recordset.RecordCount > 0 Then
'        Dim sc As New LancecomBrasil
        If MsgBox("Confirma Emissão Boleto ?", vbYesNo, App.Title) = vbYes Then
            VarCnpj = LerINI("Config", "Cnpj", FrmFrente.Caminho & "\config.ini")
        
            VarEmitente = LerINI("Emitente", "RazaoSocial", FrmFrente.Caminho & "\nfe.ini")
            VarEndEmitente = LerINI("Emitente", "Logradouro", FrmFrente.Caminho & "\nfe.ini")
            VarNroemitente = LerINI("Emitente", "Nro", FrmFrente.Caminho & "\nfe.ini")
            VarCidadeEmitente = LerINI("Emitente", "Cidade", FrmFrente.Caminho & "\nfe.ini")
            VarUfEmitente = LerINI("Emitente", "Uf", FrmFrente.Caminho & "\nfe.ini")
            
            
            DatPcl.Recordset.MoveFirst
            Open FrmFrente.Caminho & "\env_Boleto.txt" For Output As #1
            
            While Not DatPcl.Recordset.EOF
            
                Registro = VarConvenio & "|"
                Registro = Registro & VarCnpj & "|"
                Registro = Registro & VarAgencia & "|"
                Registro = Registro & VarConta & "|"
                Registro = Registro & "18|"
                Registro = Registro & VarEmitente & "|"
                Registro = Registro & VarEndEmitente & "|"
                Registro = Registro & VarNroemitente & "|"
                Registro = Registro & VarCidadeEmitente & "|"
                Registro = Registro & VarUfEmitente & "|"
                Registro = Registro & VarDocumento & "|"
                Registro = Registro & Entidade & "|"
                Registro = Registro & EnderecoEntidade & "|"
                Registro = Registro & Bairro & "|"
                Registro = Registro & CidadeEntidade & "|"
                Registro = Registro & UfEntidade & "|"
                Registro = Registro & CepEntidade & "|"
                Registro = Registro & psCompString(1, DatPcl.Recordset.Fields("pedido"), 6) & "|"
                Registro = Registro & CStr(DatPcl.Recordset.Fields("DtEmiss")) & "|"
                Registro = Registro & CStr(DatPcl.Recordset.Fields("DtVenc")) & "|"
                Registro = Registro & CStr(Format(DatPcl.Recordset.Fields("Valor"), "###,##0.00")) & "|"
                Registro = Registro & CStr(DatPcl.Recordset.Fields("CodDup")) & "|"
                Registro = Registro & "Após o Vencimento, corrigir pelo IGPM/FGV" & "|"
                Registro = Registro & "Cobrar JUROS de 1% ao mês e MULTA de 10%" & "|"
                Registro = Registro & "Pagável em qualquer agência bancária até o vencimento" & "|"
                Print #1, Registro
            
            
            
    '            Call sc.GeraBoleto(psCompString(1, VarConvenio, 6), VarCnpj, Mid$(VarAgencia, 6, 1), Left(VarAgencia, 4), Left(VarConta, 5), _
    '                        Mid(VarConta, 7, 1), "18", "ENERGIA MATERIAIS PARA CONSTRUÇÃO EIRELI – ME", "RUA SÃO VICENTE DE PAULA 617 – CENTRO <br>IBIRÁ – SP  - FONE – (17) 3551-3900", _
    '                        Entidade, EnderecoEntidade, Bairro, CidadeEntidade, UfEntidade, CepEntidade, "", psCompString(1, DatPcl.Recordset.Fields("CodDup"), 6), _
    '                        CStr(DatPcl.Recordset.Fields("DtEmiss")), CStr(DatPcl.Recordset.Fields("DtVenc")), CStr(Format(DatPcl.Recordset.Fields("Valor"), "###,##0.00")), _
    '                        CStr(DatPcl.Recordset.Fields("Pedido")), "<B> Após o Vencimento, corrigir pelo IGPM/FGV <BR> Cobrar JUROS de 1% ao mês e MULTA de 10% <BR> Pagável em qualquer agência bancária até o vencimento <BR></b>", _
    '                        "aa ", "bb", "1", "SIDCOMP", CStr(DatPcl.Recordset.Fields("DtVenc")))
            
            
                DatPcl.Recordset.MoveNext
            Wend
            Close #1
            DoEvents
            
                        
            Shell FrmFrente.Caminho & "\ERP.IntegraTxt.exe " & FrmFrente.Caminho & "\env_boleto.txt 10 1"
            
        End If
    End If
End Sub

Private Sub GeraCodigos()
   'calculo do nosso numero
   NossoNumero = Calculo_NossoNumero(psCompString(1, DatPcl.Recordset.Fields("CodDup"), 6))
  'calculo do codigo de barras
   CodigoBarras = Monta_CodBarras(VarBanco, "9", CCur(DatPcl.Recordset.Fields("valor")), CDate(DatPcl.Recordset.Fields("DtVenc")), VarConta & NossoNumero & "21")
   
   TxtCodBarra.Text = CodigoBarras
   TxtLinhaDigitavel.Text = LinhaDigitavel
End Sub
Private Sub ImprimirBol()
    RptRel.DataFiles(0) = FrmFrente.Caminho & "\Dados.mdb"
    RptRel.WindowState = 2
    RptRel.ReportFileName = FrmFrente.Caminho & "\boleto.rpt"
    RptRel.SelectionFormula = "{Dup.Pedido} = " & DBGrid.Columns(0)
    RptRel.Formulas(1) = "agencia = '" & VarAgencia & "'"
    RptRel.Formulas(2) = "conta = '" & VarConta & "'"
    RptRel.WindowControlBox = True
    
    RptRel.Destination = 0
    RptRel.Action = 1

End Sub

Private Function Monta_CodBarras(Banco As String, Moeda As String, Valor As Single, vencimento As Date, Livre As String)

Dim codigo_sequencia As String
Dim database As Date
Dim fator As Integer
Dim intDac As Integer

'database para calculo do fator
database = CDate("7/10/1997")
fator = DateDiff("d", database, Format(vencimento, "dd/mm/yyyy"))
Valor = Int(Valor * 100)
Livre = Format(Livre, "0000000000000000000000000")

' sequencia sem o DV
codigo_sequencia = Banco & Moeda & fator & Format(Valor, "0000000000") & Livre

' calculo do DV do codigo de barras
intDac = calcula_DV_CodBarras(codigo_sequencia)

' monta a sequencia para o codigo de barras com o DV
Monta_CodBarras = (Left(codigo_sequencia, 4) & intDac & Right(codigo_sequencia, 39))

' monta sequencia da linha digitavel com os dvs
LinhaDigitavel = Linha_Digitavel(VarBanco & "9" & VarConta & NossoNumero, CStr(intDac), CCur(Valor))

End Function


Private Function calcula_DV_CodBarras(sequencia As String) As Integer

Dim intcontador, intnumero, intTotalNumero As Integer
Dim intMultiplicador, intResto, intresultado As Integer
Dim caracter As String

intMultiplicador = 2

For intcontador = 1 To 43
    caracter = Mid(Right(sequencia, intcontador), 1, 1)
    If intMultiplicador > 9 Then
       intMultiplicador = 2
       intnumero = 0
    End If
    intnumero = caracter * intMultiplicador
    intTotalNumero = intTotalNumero + intnumero
    intMultiplicador = intMultiplicador + 1
Next

intResto = intTotalNumero Mod 11

intresultado = 11 - intResto

If intresultado = 10 Or intresultado = 11 Then
    calcula_DV_CodBarras = 1
Else
    calcula_DV_CodBarras = intresultado
End If
    
End Function
Function Linha_Digitavel(sequencia As String, DV_CodBarras As String, Valor As Single) As String

Dim seq1 As String
Dim seq2 As String
Dim seq3 As String

Dim dv1, dv2, dv3 As Integer

'separa a sequencia e prepara o valor
seq1 = Left(sequencia, 9)
seq2 = Mid(sequencia, 10, 10)
seq3 = Right(sequencia, 10)
Valor = Int(Valor * 100)

' calcula os dvs
dv1 = Val(Calculo_DV10(seq1))
dv2 = Val(Calculo_DV10(seq2))
dv3 = Val(Calculo_DV10(seq3))

'formata a sequencia
seq1 = Left(seq1 & dv1, 5) & "." & Right(seq1 & dv1, 5)
seq2 = Left(seq2 & dv2, 5) & "." & Right(seq2 & dv2, 6)
seq3 = Left(seq3 & dv3, 5) & "." & Right(seq3 & dv3, 6)

Linha_Digitavel = seq1 & " " & seq2 & " " & seq3 & " " & DV_CodBarras & " " & Valor

End Function
Function Calculo_DV10(strNumero As String) As String
'declara As variáveis
Dim intcontador, intnumero, intTotalNumero, intMultiplicador, intResto As Integer

' se nao for um valor numerico sai da função
If Not IsNumeric(strNumero) Then
  Calculo_DV10 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 2

'pega cada caracter do numero a partir da direita
For intcontador = Len(strNumero) To 1 Step -1

'extrai o caracter e multiplica prlo multiplicador
intnumero = Val(Mid(strNumero, intcontador, 1)) * intMultiplicador

' se o resultado for maior que nove soma os algarismos do resultado
If intnumero > 9 Then
  intnumero = Val(Left(intnumero, 1)) + Val(Right(intnumero, 1))
End If

'soma o resultado para totalização
intTotalNumero = intTotalNumero + intnumero

'se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
intMultiplicador = IIf(intMultiplicador = 2, 1, 2)

Next

'calcula o resto da divisao do total por 10
intResto = intTotalNumero Mod 10

'calcula o valor do DV ( equivale a subtrair da dezena anterior)
intResto = 10 - intResto

'verifica as exceções ( 0 -> DV=0 )
Select Case intResto
  Case 0
     Calculo_DV10 = "0"
  Case Else
     Calculo_DV10 = str(intResto)
End Select

End Function
Function Calculo_DV11(strNumero As String) As String
'declara as variáveis
Dim intcontador, intnumero, intTotalNumero, intMultiplicador, intResto As Integer

' se nao for um valor numerico sai da função
If Not IsNumeric(strNumero) Then
  Calculo_DV11 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 9

'pega cada caracter do numero a partir da direita
For intcontador = Len(strNumero) To 1 Step -1

'extrai o caracter e multiplica prlo multiplicador
intnumero = Val(Mid(strNumero, intcontador, 1)) * intMultiplicador

'soma o resultado para totalização
intTotalNumero = intTotalNumero + intnumero

'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)

Next

'calcula o resto da divisao do total por 11
intResto = intTotalNumero Mod 11

'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
Select Case intResto
  Case 0
    Calculo_DV11 = "0"
  Case 10
    Calculo_DV11 = "X"
  Case Else
    Calculo_DV11 = str(intResto)
End Select

End Function
Function Calculo_NossoNumero(sequencia As String) As String
 'montamos o nosso numero com o numero do convenio ( 6 posicoes)
    Dim dv As Integer
 
    dv = Calculo_DV11(sequencia)
    Calculo_NossoNumero = Format(sequencia & dv, "00000000000000000")
    
End Function


Private Sub Receb_Click()
Dim VarDdesc As Currency
Dim VarSoma As Currency
Dim VarDesconto As Currency
VarDesc = 0

If MsgBox("Confirma recebimento ?", vbYesNo, App.Title) = vbYes Then

    If DatPed.Recordset.RecordCount > 0 Then
        If DatPed.Recordset.Fields("PrevEntrega") = "" Or IsNull(DatPed.Recordset.Fields("PrevEntrega")) Then
            DatPed.Recordset.Edit
            DatPed.Recordset.Fields("PrevEntrega") = "RECEBIDO"
            DatPed.Recordset.Update
            If CCur(MebTotal) < DatPed.Recordset.Fields("Total") Then
                VarDesc = DatPed.Recordset.Fields("Total") - MebTotal
                DatPed.Recordset.Edit
                DatPed.Recordset.Fields("Total") = DatPed.Recordset.Fields("Total") - VarDesc
                DatPed.Recordset.Fields("Desconto") = VarDesc
                DatPed.Recordset.Update
            End If
            If CCur(MebTotal) > DatPed.Recordset.Fields("Total") Then
               sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, D_C ) VALUES ( "
               sql = sql + """" & Maquina & """" & ", "
               sql = sql & """" & "ACRESCIMO DE RECEBIMENTO" & """" & " , #"
               sql = sql & Format(Date, "mm/dd/yyyy") & "# , #"
               sql = sql & Format(Time, "hh:mm:ss") & "# , "
               sql = sql & Numero2(CStr(CCur(MebTotal) - DatPed.Recordset.Fields("Total"))) & " ,"
    '           sql = sql & "" & " ,"
               sql = sql & """" & "C" & """" & ")"
               FrmFrente.Banco.Execute sql
            End If
        Else
            DatPed.Recordset.Edit
            DatPed.Recordset.Fields("PrevEntrega") = ""
            DatPed.Recordset.Update
            DatItens.Recordset.MoveFirst
            VarSoma = 0
            VarDesconto = 0
            While Not DatItens.Recordset.EOF
                VarSoma = VarSoma + DatItens.Recordset.Fields("Valor")
                VarDesconto = VarDesconto + DatItens.Recordset.Fields("Desconto")
                DatItens.Recordset.MoveNext
            Wend
            DatPed.Recordset.Edit
            DatPed.Recordset.Fields("Total") = VarSoma
            DatPed.Recordset.Fields("Desconto") = 0
            DatPed.Recordset.Update
            MebTotal = VarSoma
            
            MsgBox "Se houve algum acrescimo nesse pedido vc deve extorna-lo no caixa !", vbInformation
            
        End If
    End If
    
    Perc = 0
    MebDesconto = 0
    Total = 0
    MebAcresc = 0
'    If Not IsNull(DatPed.Recordset.Fields("Desconto")) Then
'        MebDesconto = DatPed.Recordset.Fields("Desconto")
'    End If
    Total = (MebTotal - MebDesconto) + MebAcresc
    
    
    
    If Total > (MebTotal - MebDesconto) + MebAcresc Then
        MsgBox "Troco ===>  " & Format(Total > ((MebTotal - MebDesconto) + MebAcresc), "###.##0,00"), vbInformation
    End If
End If

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub



Private Sub Total_LostFocus()
If Total = "" Then
Total = (CCur(MebTotal) - CCur(MebDesconto)) + CCur(MebAcresc)
End If
End Sub
