VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormEntrega 
   BackColor       =   &H00C0BB9C&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7605
   ClientLeft      =   2010
   ClientTop       =   2220
   ClientWidth     =   13455
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   52
      Top             =   7245
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13470
      TabIndex        =   50
      Top             =   0
      Width           =   13470
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de  Entrega"
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
         TabIndex        =   51
         Top             =   45
         Width           =   5370
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0BB9C&
      Caption         =   "Pagamento"
      Height          =   6180
      Left            =   10080
      TabIndex        =   37
      Top             =   855
      Width           =   3165
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0BB9C&
         Height          =   735
         Left            =   180
         TabIndex        =   39
         Top             =   180
         Width           =   2850
         Begin VB.Label LblValor 
            BackColor       =   &H00C0BB9C&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   135
            TabIndex        =   40
            Top             =   120
            Width           =   2640
         End
      End
      Begin VB.ComboBox CmbTipo 
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
         Height          =   2340
         ItemData        =   "FormEntrega.frx":0000
         Left            =   225
         List            =   "FormEntrega.frx":0016
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Top             =   2925
         Width           =   2805
      End
      Begin MSMask.MaskEdBox MebPago 
         Height          =   405
         Left            =   225
         TabIndex        =   41
         Top             =   1350
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483641
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTroco 
         Height          =   405
         Left            =   225
         TabIndex        =   43
         Top             =   2115
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483641
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CodDoc 
         Height          =   270
         Left            =   1710
         TabIndex        =   46
         Top             =   5805
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CodEnt 
         Height          =   270
         Left            =   315
         TabIndex        =   47
         Top             =   5805
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pedido"
         Height          =   195
         Left            =   1710
         TabIndex        =   49
         Top             =   5580
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Entidade"
         Height          =   195
         Left            =   315
         TabIndex        =   48
         Top             =   5580
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   225
         TabIndex        =   45
         Top             =   2610
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco"
         Height          =   195
         Left            =   225
         TabIndex        =   44
         Top             =   1845
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do Pagamento"
         Height          =   195
         Left            =   225
         TabIndex        =   42
         Top             =   1080
         Width           =   1440
      End
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
      Left            =   3870
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2025
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FormEntrega.frx":0073
      Height          =   2550
      Left            =   180
      OleObjectBlob   =   "FormEntrega.frx":008D
      TabIndex        =   27
      ToolTipText     =   "Duplo Click p/ Alterar Entrega"
      Top             =   855
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0BB9C&
      Caption         =   "Atualiza Tabela"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   7560
      TabIndex        =   24
      Top             =   4230
      Width           =   2085
      Begin VB.CommandButton BtAtualiza 
         Caption         =   "Atualiza Tabela"
         Enabled         =   0   'False
         Height          =   375
         Left            =   270
         TabIndex        =   14
         Top             =   900
         Width           =   1680
      End
      Begin VB.OptionButton OptCompra 
         BackColor       =   &H00C0BB9C&
         Caption         =   "Margem"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   540
         Width           =   1455
      End
      Begin VB.OptionButton OptCompra 
         BackColor       =   &H00C0BB9C&
         Caption         =   "Preço Venda"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0BB9C&
      Caption         =   "Transportadora"
      Height          =   3495
      Left            =   180
      TabIndex        =   16
      Top             =   3510
      Width           =   9735
      Begin VB.CheckBox ChkIpi 
         BackColor       =   &H00C0BB9C&
         Caption         =   "Calcular I.P.I."
         Height          =   375
         Left            =   7605
         TabIndex        =   36
         Top             =   2295
         Width           =   1410
      End
      Begin VB.CommandButton CmdCusto 
         Caption         =   "Interface Centro Custo"
         Height          =   330
         Left            =   7380
         TabIndex        =   35
         Top             =   225
         Width           =   2085
      End
      Begin VB.CommandButton BtOk 
         Caption         =   "&Confirma"
         Height          =   420
         Left            =   7605
         TabIndex        =   34
         Top             =   2835
         Width           =   1635
      End
      Begin VB.CheckBox ChkSub 
         BackColor       =   &H00C0BB9C&
         Caption         =   "&Nota c/ Sub. Tributária"
         Height          =   465
         Left            =   135
         TabIndex        =   5
         Top             =   1215
         Width           =   1275
      End
      Begin VB.TextBox TxtContato 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2505
         Width           =   3510
      End
      Begin VB.ComboBox TxtForma 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormEntrega.frx":1B3C
         Left            =   135
         List            =   "FormEntrega.frx":1B4C
         TabIndex        =   11
         Top             =   2505
         Width           =   3120
      End
      Begin VB.Data DatDados 
         Caption         =   "Transport"
         Connect         =   "Access"
         DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   3
         Left            =   2070
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodEntidade, Nome  From Ent Where Tipo = 4 or Tipo = 6 Order by Nome"
         Top             =   180
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.ComboBox CmbFrete 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormEntrega.frx":1B74
         Left            =   135
         List            =   "FormEntrega.frx":1B7E
         TabIndex        =   1
         Top             =   810
         Width           =   2310
      End
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   135
         TabIndex        =   13
         Top             =   3060
         Width           =   6855
      End
      Begin VB.TextBox txtEntrega 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1950
         Width           =   3510
      End
      Begin VB.ComboBox TxtCondPgto 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormEntrega.frx":1B8C
         Left            =   135
         List            =   "FormEntrega.frx":1B96
         TabIndex        =   9
         Top             =   1980
         Width           =   3105
      End
      Begin MSMask.MaskEdBox MebVlrFrete 
         Height          =   315
         Left            =   5580
         TabIndex        =   4
         Top             =   810
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebFrete 
         Height          =   315
         Left            =   4095
         TabIndex        =   3
         Top             =   810
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebValor 
         Height          =   315
         Left            =   2610
         TabIndex        =   2
         Top             =   810
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo DbcTransport 
         Bindings        =   "FormEntrega.frx":1BA8
         DataField       =   "Nome"
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   210
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MebSub 
         Height          =   315
         Left            =   1575
         TabIndex        =   6
         Top             =   1350
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebIpi 
         Height          =   315
         Left            =   3465
         TabIndex        =   7
         Top             =   1350
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebCredito 
         Bindings        =   "FormEntrega.frx":1BC2
         Height          =   315
         Left            =   5445
         TabIndex        =   8
         Top             =   1350
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Substituição"
         Height          =   195
         Left            =   1575
         TabIndex        =   33
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.P.I."
         Height          =   195
         Left            =   3465
         TabIndex        =   32
         Top             =   1170
         Width           =   330
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Presumido"
         Height          =   195
         Left            =   5445
         TabIndex        =   31
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   6570
         TabIndex        =   30
         Top             =   1395
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contato (Comprador/Vendedor)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   29
         Top             =   2310
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Forma Pagto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   2310
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo Frete"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor do Frete"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5580
         TabIndex        =   22
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Calcula Frete"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4095
         TabIndex        =   21
         Top             =   615
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2610
         TabIndex        =   20
         Top             =   615
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Observação"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   2865
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cond.Pagto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1710
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Prazo Entrega"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   1710
         Width           =   1005
      End
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormEntrega.frx":1BCD
      Height          =   2550
      Left            =   180
      OleObjectBlob   =   "FormEntrega.frx":1BE7
      TabIndex        =   15
      ToolTipText     =   "Duplo Click p/ Alterar Entrega"
      Top             =   855
      Width           =   9735
   End
End
Attribute VB_Name = "FormEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Loja As Recordset
Public Baixa As Boolean
Public VarIcms As Boolean
Public VarSomaIPi As Currency
Public VarSub As Boolean

Private Sub BtAtualiza_Click()
Dim VarVenda As Currency

DatDados(1).Recordset.MoveFirst
While Not DatDados(1).Recordset.Eof

    DatDados(1).Recordset.Edit
    DatDados(1).Recordset.Fields("Custo") = DatDados(1).Recordset.Fields("ValorUnit")
    If OptCompra(0).value = True Then
        VarVenda = DatDados(1).Recordset.Fields("ValorUnit") + (DatDados(1).Recordset.Fields("ValorUnit") * (DatDados(1).Recordset.Fields("Margem") / 100))
        DatDados(1).Recordset.Fields("Venda") = VarVenda
        If VarPzoProduto = True Then
            DatDados(1).Recordset.Fields("PcoPrazo") = VarVenda + (VarVenda * (DatDados(1).Recordset.Fields("MargemPrazo") / 100))
        End If
    Else
        DatDados(1).Recordset.Fields("Margem") = ((DatDados(1).Recordset.Fields("Venda") / (DatDados(1).Recordset.Fields("ValorUnit")) - 1) * 100)
    End If
    DatDados(1).Recordset.Fields("DtAlt") = Date
    DatDados(1).Recordset.Update
    DatDados(1).Recordset.MoveNext
Wend

DBGrid2.Visible = True

End Sub

Private Sub BtOK_Click()
Dim Criterio As String
Dim flag As Boolean
Dim AtuPco As Integer
Dim VarSomaEst As Double
Dim VarSaldoEst As Double
Dim VarPedido As Double
Dim VarImpSub As Currency
Dim VarBaseSub As Currency
Dim TotImpSub As Currency
Dim VarProd As Currency
Dim VarIcmsSub As Currency
Dim VarFreteProp As Currency
Dim VarBaseSub2 As Currency
Dim VarIcmsSub2 As Currency
Dim SQL As String

VarBaseSub = 0
VarImpSub = 0
TotImpSub = 0
VarProd = 0
VarFreteProp = 0
VarIcmsSub = 0

VarSomaEst = 0

'On Error GoTo erro

Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")

Loja.Requery

flag = False

DatDados(1).Recordset.MoveFirst

While Not DatDados(1).Recordset.Eof
    If DatDados(1).Recordset.Fields("ITP.Entregue") = True Then
        If DatDados(1).Recordset.Fields("Loja") = "" Then
            flag = True
        End If
    End If
    
    If ChkSub.value = 1 Then
        If VarSub = True Then
                Dim VarUfIva As Recordset
                Dim VarUfEstado As Recordset
'                SQL = "Select * from IVA Where Uf = " & """" & Trim(FrmPedidos.Datdados(0).Recordset.Fields("Uf")) & """" & " AND Classe = " & """" & Trim(Datdados(1).Recordset.Fields("Classe")) & """"
                Set VarUfEstado = Banco.OpenRecordset("Select * from ICM Where Uf = " & """" & Trim(FrmPedidos.DatDados(0).Recordset.Fields("Uf")) & """")
                VarUfEstado.Requery
                Set VarUfIva = Banco.OpenRecordset("Select * from IVA Where Uf = " & """" & Trim(FrmPedidos.DatDados(0).Recordset.Fields("Uf")) & """" & " AND Ncm = " & """" & Trim(DatDados(1).Recordset.Fields("Classe")) & """") ' & "AND Cest = " & """" & Trim(DatDados(1).Recordset.Fields("ClasseFiscal")) & """")
                VarUfIva.Requery
                If VarUfIva.RecordCount > 0 Then
                    
                  If DatDados(1).Recordset.Fields("ITP.Complemento") <> "BONIFICACAO-teste" Then
                       If MebVlrFrete <> 0 Then
                            VarFreteProp = (DatDados(1).Recordset("Valor") / CCur(FrmPedidos.TxtDados(1))) * MebVlrFrete
                            VarProd = DatDados(1).Recordset("Valor") + VarFreteProp
                       Else
                            VarProd = DatDados(1).Recordset("Valor")
                       End If
                                  
                       VarBaseSub = VarBaseSub + (((VarProd / 100) * VarUfIva("Iva")) + VarProd)
                       VarBaseSub2 = (((VarProd / 100) * VarUfIva("Iva")) + VarProd)
                       VarIcmsSub = ((VarProd / 100) * VarUfEstado!Icms)
                       
                       
                       VarIcmsSub2 = ((DatDados(1).Recordset("Valor") / 100) * VarUfIva!Icms)
                       If VarUfIva!Icms <> 20 Then
                            VarImpSub = (VarBaseSub2 * (VarUfIva!Icms / 100)) - VarIcmsSub
                       Else
                            VarImpSub = (VarBaseSub2 * (VarUfIva!Icms / 100))
                       End If
                       
                       If Not IsNumeric(MebCredito) Then
                           MebCredito = 0
                       End If
                       If MebCredito <> 0 Then
                            TotImpSub = TotImpSub + ((VarImpSub - ((DatDados(1).Recordset("Valor") / 100) * MebCredito)))
                       Else
                            TotImpSub = TotImpSub + ((VarImpSub)) ' - ((Datdados(1).Recordset("Valor") / 100) * Datdados(1).Recordset("Redutor"))))
                       End If
                  End If
                Else
                    MsgBox "Nao tem Icms cadastrado !", vbInformation, App.Title
                    Exit Sub
                End If
        Else
            If FrmPedidos.DatDados(0).Recordset.Fields("Uf") <> "SP" Then
                Dim VarUf As Recordset
                Set VarUf = Banco.OpenRecordset("Select * from ICM Where Uf = " & """" & Trim(FrmPedidos.DatDados(0).Recordset.Fields("Uf")) & """")
                VarUf.Requery
                If VarUf.RecordCount > 0 Then
                    
                    If MebSub = 0 And DatDados(1).Recordset("Iva") <> 0 Then
                       If MebVlrFrete <> 0 Then
                            VarFreteProp = (DatDados(1).Recordset("Valor") / CCur(FrmPedidos.TxtDados(1))) * MebVlrFrete
                            VarProd = DatDados(1).Recordset("Valor") + VarFreteProp
                       Else
                            VarProd = DatDados(1).Recordset("Valor")
                       End If
                                  
                       VarBaseSub = VarBaseSub + (((VarProd / 100) * DatDados(1).Recordset("Iva")) + VarProd)
                       VarBaseSub2 = (((VarProd / 100) * DatDados(1).Recordset("Iva")) + VarProd)
                       If VarUf!Entrada <> 0 Then
                            VarIcmsSub = ((VarProd / 100) * VarUf!Entrada)
                       Else
                            VarIcmsSub = ((VarProd / 100) * VarUf!Icms)
                       End If
                       
                       VarIcmsSub2 = ((DatDados(1).Recordset("Valor") / 100) * VarUf!Icms)
                                          
                       VarImpSub = (VarBaseSub2 * (VarUf!Icms / 100)) - VarIcmsSub
                       
                       If Not IsNumeric(MebCredito) Then
                           MebCredito = 0
                       End If
                       If MebCredito <> 0 Then
                            TotImpSub = TotImpSub + ((VarImpSub - ((DatDados(1).Recordset("Valor") / 100) * MebCredito)))
                       Else
                            TotImpSub = TotImpSub + ((VarImpSub)) ' - ((Datdados(1).Recordset("Valor") / 100) * Datdados(1).Recordset("Redutor"))))
                       End If
                    End If
                Else
                    MsgBox "Nao tem Icms cadastrado !", vbInformation, App.Title
                    Exit Sub
                End If
                    
            Else
                If MebSub = 0 And DatDados(1).Recordset("Iva") <> 0 Then
                   If MebVlrFrete <> 0 Then
                        VarFreteProp = (DatDados(1).Recordset("Valor") / CCur(FrmPedidos.TxtDados(1))) * MebVlrFrete
                        VarProd = DatDados(1).Recordset("Valor") + VarFreteProp
                   Else
                        
                        VarProd = DatDados(1).Recordset("Valor")
                   End If
                              
                   VarBaseSub = VarBaseSub + (((VarProd / 100) * DatDados(1).Recordset("Iva")) + VarProd)
                   VarBaseSub2 = (((VarProd / 100) * DatDados(1).Recordset("Iva")) + VarProd)
                   If DatDados(1).Recordset("Redutor") <> 0 Then
                        VarIcmsSub = ((VarProd / 100) * DatDados(1).Recordset("Redutor"))
                   Else
                        VarIcmsSub = ((VarProd / 100) * DatDados(1).Recordset("Icm"))
                   End If
                   
                   VarIcmsSub2 = ((DatDados(1).Recordset("Valor") / 100) * DatDados(1).Recordset("Icm"))
                   
                   
        '           VarImpSub = ((((VarProd / 100) * Datdados(1).Recordset("Iva")) + VarProd) * (Datdados(1).Recordset("Icm") / 100))
                   VarImpSub = (VarBaseSub2 * (DatDados(1).Recordset("Icm") / 100)) - VarIcmsSub
                   VarImpSub = Round(VarImpSub, 2)
                   
                   If Not IsNumeric(MebCredito) Then
                       MebCredito = 0
                   End If
                   If MebCredito <> 0 Then
                        TotImpSub = TotImpSub + ((VarImpSub - ((DatDados(1).Recordset("Valor") / 100) * MebCredito)))
                   Else
                        TotImpSub = TotImpSub + ((VarImpSub)) ' - ((Datdados(1).Recordset("Valor") / 100) * Datdados(1).Recordset("Redutor"))))
        '                TotImpSub = TotImpSub + ((VarImpSub - ((Datdados(1).Recordset("Valor") / 100) * Datdados(1).Recordset("Redutor"))))
                   
                   End If
                End If
            End If
        End If
    Else
       If MebVlrFrete <> 0 Then
             VarFreteProp = (DatDados(1).Recordset("Valor") / CCur(FrmPedidos.TxtDados(1))) * MebVlrFrete
             VarProd = DatDados(1).Recordset("Valor") + VarFreteProp
        Else
             VarProd = DatDados(1).Recordset("Valor")
        End If

    End If
        
    DatDados(1).Recordset.MoveNext
Wend

If flag Then
     MsgBox "Tem Movimentação de Estoque sem Loja Especificada", vbInformation, App.Title
     Exit Sub
End If

If DatDados(1).Recordset.RecordCount <> 0 Then
    DatDados(1).Recordset.MoveFirst
    
  
    Select Case DatDados(1).Recordset.Fields("Operacao")
           Case 0
                Baixa = False
            Case 1
                Baixa = True
            Case 2
                Baixa = True
            Case 3
                Baixa = True
            Case 6
                Baixa = True
            Case 8
                Baixa = True
            Case 9
                Baixa = False
            Case 10
                Baixa = False
            Case 11
                Baixa = True
            Case 14
                Baixa = True
            Case 15
                Baixa = False
            Case 16
                Baixa = False
                            
    End Select
    
    DatDados(1).Recordset.Edit
    DatDados(1).Recordset("Ped.Entregue") = 1
    DatDados(1).Recordset.Update
    
    BaixaEstoque DatDados(1).Recordset.Fields("CodPed"), Baixa, DatDados(1).Recordset.Fields("Operacao")
    
    DatDados(1).Recordset.Edit
    DatDados(1).Recordset("Ped.Entregue") = 0
    DatDados(1).Recordset.Update
    
    DatDados(1).Recordset.Edit
    If DbcTransport.BoundText <> "" Then
        If IsNumeric(DbcTransport.BoundText) Then
            If DbcTransport.BoundText <> 0 Then
                DatDados(1).Recordset("Transportadora") = DbcTransport.BoundText
            End If
        End If
    End If
        
    VarPedido = DatDados(1).Recordset.Fields("CodPed")
    
    DatDados(1).Recordset("TipoFrete") = CmbFrete.Text
    If MebVlrFrete <> 0 Then
        DatDados(1).Recordset("ValorFrete") = MebVlrFrete
    Else
        DatDados(1).Recordset("ValorFrete") = 0
    End If
    DatDados(1).Recordset("Ped.Obs") = TxtObs.Text
    DatDados(1).Recordset("CondPgto") = Mid(CmbTipo.Text, 4, 15)
    DatDados(1).Recordset("PrevEntrega") = txtEntrega.Text
    DatDados(1).Recordset("FormaPgto") = TxtForma.Text
    DatDados(1).Recordset("Contato") = TxtContato.Text
    
    DatDados(1).Recordset("PedCredito") = MebCredito
    DatDados(1).Recordset("PedTipoSub") = ChkSub
        
    If MebSub = 0 Then
'        If VarSub = True Then
'            Datdados(1).Recordset("PedSub") = Format(Truncar(VarImpSub, 2), "##,##0.00")   ' TotImpSub
'        Else
            DatDados(1).Recordset("PedSub") = Format(Truncar(TotImpSub, 2), "##,##0.00")     ' TotImpSub
'        End If
            
        DatDados(1).Recordset("PedBaseSub") = VarBaseSub
    Else
        DatDados(1).Recordset("PedSub") = MebSub
    End If
    
    If ChkIpi.value = 1 Then
        DatDados(1).Recordset("Ped.Ipi") = VarSomaIPi
        MebIpi = VarSomaIPi
    Else
        DatDados(1).Recordset("Ped.Ipi") = 0
        MebIpi = 0
    End If
    
    DatDados(1).Recordset("Ped.Ipi") = MebIpi
                
    DatDados(1).Recordset.Update
'    If FrmPedidos.MebDados(1) = Date Then
        SQL = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Maquina  ) VALUES ( "
        SQL = SQL & CodEnt & " , "
        SQL = SQL & "#" & Format(FrmPedidos.MebDados(1), "mm/dd/yyyy") & "# , "
        SQL = SQL & """" & Time & """" & " , "
        SQL = SQL & Numero(CStr(LblValor), 2) & " , "
        SQL = SQL & 0 & " , "
        SQL = SQL & CmbTipo.ListIndex + 11 & " ,"
        SQL = SQL & CodDoc & " ,"
        SQL = SQL & """" & LIMPACampo(Trim(Maquina)) & """" & ")"
        Banco.Execute SQL
'    End If
Else
    MsgBox "Erro na baixa/Entrada de de estoque", vbInformation, App.Title
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, "Erro Estoque modulo Entrega " & DatDados(1).Recordset.Fields("CodPed") & "; " & Date & "; " & Time & "; " & Maquina
    Close #1
End If

Unload Me

Exit Sub
Erro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
    MsgBox "Ouve um erro inesperado favor comunicar o responsavel pelo sistema, poderá haver divergência de estoque", vbInformation, App.Title
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, "Ouve um erro inesperado favor comunicar o responsavel pelo sistema, poderá haver divergência de estoque" & "; " & Date & "; " & Time & "; " & Maquina
    Close #1

End Sub


Private Sub BtSair_Click()

End Sub

Private Sub CmbTipo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 9 Then
    
    BtOK_Click
End If

End Sub

Private Sub CmbTipo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    BtOk.SetFocus
'End If
End Sub

Private Sub CmdCusto_Click()
Dim SQL As String

If MsgBox("Deseja lançar valores no centro de custo ?", vbYesNo, App.Title) = vbYes Then
    
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.Eof
                               
        SQL = "INSERT INTO CCU (Filial, Data, Tipo, Historico, Valor ) VALUES ( "
        SQL = SQL & 1 & " , #"
        SQL = SQL & Format(Date, "mm/dd/yyyy") & "#,"
        SQL = SQL & """" & "D" & """" & " ,"
        SQL = SQL & """" & DatDados(1).Recordset("Descrição_Produto") & """" & " ,"
        SQL = SQL & Numero2(DatDados(1).Recordset("VlrVdaDia")) & ")"
        Banco.Execute SQL
        
        DatDados(1).Recordset.MoveNext
        
    Wend
    Unload Me
    
    FormCentroCusto.Show
    
End If

End Sub

Private Sub DbcTransport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 9 Or KeyCode = 13 Then

    If DbcTransport.Text = "" Then
        DbcTransport.SetFocus
        Exit Sub
    End If
    If IsNumeric(DbcTransport.Text) Then
       Criterio = "CodEntidade = " & DbcTransport
       DatDados(3).Recordset.FindFirst Criterio
       If DatDados(3).Recordset.NoMatch Then
           MsgBox "Transportadora não encontrada !!!", 16
           DbcTransport.Text = ""
           DbcTransport.SetFocus
           Exit Sub
       Else
           DbcTransport.Text = DatDados(3).Recordset.Fields("Nome")
       End If
    Else
        If DbcTransport.BoundText = 0 Or DbcTransport.BoundText = "" Then
                If DbcTransport.Text <> "" Then
                    MsgBox "Transportadora não encontrada !!!", 16
                    DbcTransport.SetFocus
                    Exit Sub
                End If
         End If
    End If

End If
End Sub

Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
If DBGrid2.Columns(10) <> "" Then
    If IsNumeric(DBGrid2.Columns(10)) Then
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset.Fields("Margem") = ((DBGrid2.Columns(10) / DBGrid2.Columns(5)) - 1) * 100
        DatDados(1).Recordset.Update
    End If
End If

End Sub

Private Sub Form_Activate()
Dim Criterio As String
Dim SomaIpi As Currency
Dim SQL As String


MebVlrFrete = 0
SomaIpi = 0
MebSub = 0
MebIpi = 0
MebCredito = 0
VarSub = 0
'BtOk.SetFocus

If LerINI("Nota", "Substituicao", App.Path & "\config.ini") <> "" Then
    VarSub = LerINI("Nota", "substituicao", App.Path & "\config.ini")
Else
    VarSub = 0
End If


VarIcms = LerINI("Nota", "CalculaIcms", App.Path & "\config.ini")

FormEntrega.Caption = psCompString(1, FrmPedidos.LblCodigo, 6) & "-" & psCompString(3, FrmPedidos.DbcDados(0).Text, 40) & "     Pedido:" & psCompString(1, FrmPedidos.TxtDados(0), 6)

MebValor.Text = FrmPedidos.TxtDados(2)
'MebIpi.Text = FrmPedidos.TxtDados(4)



DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.* FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & FrmPedidos.TxtDados(0) & " Order by Itp.CodItp"
DatDados(1).Refresh

'MebSub = Datdados(1).Recordset("PedSub")
MebIpi.Text = DatDados(1).Recordset("Ped.Ipi")

'If VarIcms Then
'    ChkIpi.value = 1
'Else
'    ChkIpi.value = 0
'End If

'Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")
Set Grupo = Banco.OpenRecordset("SELECT * FROM GRU")

If DatDados(1).Recordset.RecordCount <> 0 Then
    
    DatDados(1).Recordset.MoveFirst
    
    If DatDados(1).Recordset.Fields("Operacao") = 0 Then
        BtAtualiza.Enabled = True
        Frame2.Enabled = True
    End If

    If Not IsNull(DatDados(1).Recordset.Fields("Ped.Obs")) Then
        TxtObs.Text = DatDados(1).Recordset.Fields("Ped.Obs")
    End If
    Criterio = "CodEntidade = " & FrmPedidos.Entidade!Transportadora
    DatDados(3).Recordset.FindFirst Criterio
    If Not DatDados(3).Recordset.NoMatch Then
        DbcTransport.Text = DatDados(3).Recordset.Fields("Nome")
    Else
        DbcTransport.Text = "(Nenhum)"
    End If
    If Not IsNull(DatDados(1).Recordset("ValorFrete")) Then
        MebVlrFrete.Text = DatDados(1).Recordset("ValorFrete")
    End If
    If Not IsNull(DatDados(1).Recordset("Ped.Obs")) Then
        TxtObs.Text = DatDados(1).Recordset("Ped.Obs")
    End If
    If Not IsNull(DatDados(1).Recordset("Contato")) Then
        TxtContato.Text = DatDados(1).Recordset("Contato")
    End If
    
    If Not IsNull(DatDados(1).Recordset("FormaPgto")) Then
        TxtForma.Text = DatDados(1).Recordset("FormaPgto")
    End If
           
    If Not IsNull(DatDados(1).Recordset("CondPgto")) Then
        TxtCondPgto.Text = DatDados(1).Recordset("CondPgto")
    Else
        If FrmPedidos.CmbDados(3).ListIndex = 2 Then
            TxtCondPgto.ListIndex = 1
        Else
            TxtCondPgto.ListIndex = 0
        End If
    End If
    If Not IsNull(DatDados(1).Recordset("PrevEntrega")) Then
        txtEntrega.Text = DatDados(1).Recordset("PrevEntrega")
    End If
    
    If Not IsNull(DatDados(1).Recordset("TipoFrete")) Then
        CmbFrete = DatDados(1).Recordset("TipoFrete")
    Else
        If LerINI("Nota", "TipoFrete", App.Path & "\config.ini") <> "" Then
            CmbFrete.ListIndex = LerINI("Nota", "TipoFrete", App.Path & "\config.ini")
        Else
            CmbFrete.ListIndex = 1
        End If
    End If
                
    If Not IsNull(DatDados(1).Recordset("PedCredito")) Then
        MebCredito = DatDados(1).Recordset("PedCredito")
    End If
                
    If Not IsNull(DatDados(1).Recordset("PedTipoSub")) Then
        If DatDados(1).Recordset("PedTipoSub") = True Then
            ChkSub.value = 1
        Else
            ChkSub.value = 0
        End If
    End If
                                
        While Not DatDados(1).Recordset.Eof
            If LerINI("LocalBco", "Local1", App.Path & "\config.ini") = "" Then
                Criterio = "CodGru = " & DatDados(1).Recordset.Fields("Grupo")
                Grupo.FindFirst Criterio
                If Grupo.Fields("BaixarEstoque") = True Then
'                    If IsNull(Datdados(1).Recordset("CondPgto")) Then
                        DatDados(1).Recordset.Edit
                        DatDados(1).Recordset.Fields("Itp.Entregue") = 1
                        DatDados(1).Recordset.Update
'                    End If
                Else
                    DatDados(1).Recordset.Edit
                    DatDados(1).Recordset.Fields("Itp.Entregue") = 0
                    DatDados(1).Recordset.Update
                End If
 '               If FrmPedidos.CmbDados(3).ListIndex = 0 Then
 '                   Datdados(1).Recordset.Edit
 '                   Datdados(1).Recordset.Fields("Itp.Entregue") = 1
 '                   Datdados(1).Recordset.Update
 '               End If
            Else
                DatDados(1).Recordset.Edit
                DatDados(1).Recordset.Fields("Itp.Entregue") = 0
                DatDados(1).Recordset.Update
            End If
            
            If VarIcms Then
                If IsNull(DatDados(1).Recordset("Prd.Ipi")) Then
                    MsgBox "Cadastro do Ipi com problema", vbInformation, App.Title
                Else
                    SomaIpi = SomaIpi + ((DatDados(1).Recordset("VlrVdaDia") / 100) * DatDados(1).Recordset("Prd.Ipi"))
                End If
            End If
            If Not IsNull(DatDados(1).Recordset("Prd.Ipi")) Then
                    VarSomaIPi = VarSomaIPi + ((DatDados(1).Recordset("VlrVdaDia") / 100) * DatDados(1).Recordset("Prd.Ipi"))
            End If
            DatDados(1).Recordset.MoveNext
        Wend
        If SomaIpi = 0 Then
            SomaIpi = MebIpi
        End If
        
        DatDados(1).Recordset.MoveFirst
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset("Ped.Ipi") = SomaIpi
        'DatDados(1).Recordset("Ped.PedIpi") = SomaIpi
        DatDados(1).Recordset.Update
'        MebIpi = SomaIpi
Else
    Unload Me
End If

If FrmPedidos.CmbDados(3).ListIndex = 1 Then
    Frame3.Enabled = True
    CmbTipo.ListIndex = 0
    MebPago.SetFocus
Else
    If FrmPedidos.CmbDados(3).ListIndex = 2 Then
        CmbTipo.ListIndex = 5
    End If
    Frame3.Enabled = False
    BtOK.SetFocus
End If

If VarSub = True Then
    ChkSub.value = 1
End If

LblValor.Caption = FrmPedidos.TxtDados(2).Text
CodDoc = FrmPedidos.TxtDados(0).Text
CodEnt = FrmPedidos.LblCodigo
MebPago = LblValor


End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

VarSomaIPi = 0


End Sub

Private Sub Grid_DblClick()
If DatDados(1).Recordset.RecordCount > 0 Then
    If LerINI("LocalBco", "Local1", App.Path & "\config.ini") = "" Then
        DatDados(1).Recordset.Edit
        If DatDados(1).Recordset.Fields("Itp.Entregue") = 0 Then
            DatDados(1).Recordset.Fields("Itp.Entregue") = 1
        Else
            DatDados(1).Recordset.Fields("Itp.Entregue") = 0
        End If
        DatDados(1).Recordset.Update
    End If
End If
' DatDados(1).Refresh
Grid.Refresh

End Sub

Private Sub MebFrete_LostFocus()
If MebFrete.Text <> "" Then
    MebVlrFrete = (MebValor * (MebFrete / 100))
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub MebPago_LostFocus()
If IsNumeric(MebValor) Then
    MebTroco = MebPago - CCur(LblValor)
    CmbTipo.SetFocus
End If
End Sub

Private Sub TxtCondPgto_Change()
If Len(Trim(TxtCondPgto)) > 20 Then
    MsgBox "Tamanho do texto não permitido !", vbInformation
    TxtCondPgto = Left(TxtCondPgto, 20)
    TxtCondPgto.SetFocus
End If
End Sub

Private Sub TxtForma_Change()
If Len(Trim(TxtForma)) > 20 Then
    MsgBox "Tamanho do texto não permitido !", vbInformation
    TxtForma = Left(TxtForma, 20)
    TxtCondPgto.SetFocus
End If

End Sub
