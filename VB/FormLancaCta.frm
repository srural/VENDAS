VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormLancaCta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11280
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -495
      ScaleHeight     =   345
      ScaleWidth      =   11820
      TabIndex        =   54
      Top             =   6930
      Width           =   11850
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11310
      TabIndex        =   52
      Top             =   0
      Width           =   11310
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançar Contas a Receber"
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
         TabIndex        =   53
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0FFFF&
      Height          =   825
      Left            =   4140
      TabIndex        =   46
      Top             =   3105
      Visible         =   0   'False
      Width           =   3120
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde . . . "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   810
         TabIndex        =   47
         Top             =   270
         Width           =   1755
      End
   End
   Begin VB.TextBox TxtPedido 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CheckBox ChkJuros 
      Caption         =   "&C/Juros"
      Height          =   285
      Left            =   8580
      TabIndex        =   2
      Top             =   855
      Width           =   915
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   5715
      TabIndex        =   41
      Top             =   855
      Width           =   1095
   End
   Begin VB.CheckBox ChkPago 
      Caption         =   "&Mostrar os Pagos"
      Height          =   285
      Left            =   9570
      TabIndex        =   3
      Top             =   855
      Width           =   1590
   End
   Begin VB.Frame Frame2 
      Caption         =   "Duplicata/Parcela"
      Height          =   1410
      Left            =   1425
      TabIndex        =   31
      Top             =   5130
      Width           =   6885
      Begin VB.CommandButton Btdupli 
         Caption         =   "&Carnê"
         Height          =   330
         Index           =   1
         Left            =   5850
         TabIndex        =   51
         Top             =   990
         Width           =   840
      End
      Begin VB.CommandButton btPromi 
         Caption         =   "&Promissória"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4710
         TabIndex        =   50
         Top             =   990
         Width           =   1080
      End
      Begin VB.CommandButton BtBoleto 
         Caption         =   "&Boleto"
         Height          =   330
         Left            =   2730
         TabIndex        =   43
         Top             =   990
         Width           =   900
      End
      Begin VB.CommandButton Btdupli 
         Caption         =   "&Mercantil"
         Height          =   330
         Index           =   0
         Left            =   3675
         TabIndex        =   35
         Top             =   990
         Width           =   990
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Parcela"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   34
         Top             =   990
         Width           =   1005
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Documento"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   33
         Top             =   652
         Width           =   1230
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pedido"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   32
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin MSMask.MaskEdBox MebDtSaida 
         DataField       =   "Custo"
         Height          =   315
         Left            =   1350
         TabIndex        =   36
         Tag             =   "S"
         Top             =   990
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebBanco 
         DataField       =   "Custo"
         Height          =   315
         Left            =   1350
         TabIndex        =   38
         Top             =   450
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebNota 
         DataField       =   "Custo"
         Height          =   315
         Left            =   4290
         TabIndex        =   39
         Top             =   435
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotalNota 
         DataField       =   "Custo"
         Height          =   315
         Left            =   5460
         TabIndex        =   48
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   15
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total"
         Height          =   195
         Index           =   9
         Left            =   5460
         TabIndex        =   49
         Top             =   225
         Width           =   765
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Nro Nota"
         Height          =   195
         Index           =   8
         Left            =   4290
         TabIndex        =   45
         Top             =   210
         Width           =   645
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Index           =   7
         Left            =   1350
         TabIndex        =   40
         Top             =   225
         Width           =   465
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Data &Saida"
         Height          =   195
         Index           =   6
         Left            =   1350
         TabIndex        =   37
         Top             =   765
         Width           =   795
      End
   End
   Begin VB.ComboBox CmbObs 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormLancaCta.frx":0000
      Left            =   8145
      List            =   "FormLancaCta.frx":0010
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox TxtPcl 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1590
      TabIndex        =   5
      Top             =   1440
      Width           =   870
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormLancaCta.frx":0044
      Left            =   6930
      List            =   "FormLancaCta.frx":0066
      TabIndex        =   1
      Top             =   855
      Width           =   1560
   End
   Begin VB.Data DatDados 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4725
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa"
      Height          =   315
      Left            =   8505
      TabIndex        =   27
      Top             =   6210
      Width           =   1020
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Item"
      Height          =   315
      Left            =   225
      TabIndex        =   12
      Top             =   5130
      Width           =   1050
   End
   Begin VB.Data DatDados 
      Caption         =   "Dup"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2070
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3615
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormLancaCta.frx":0095
      Height          =   3180
      Left            =   180
      OleObjectBlob   =   "FormLancaCta.frx":00AF
      TabIndex        =   17
      Top             =   1875
      Width           =   10935
   End
   Begin VB.CommandButton CmdLancar 
      Caption         =   "&Lançar"
      Height          =   315
      Left            =   8505
      TabIndex        =   11
      Top             =   5820
      Width           =   1020
   End
   Begin VB.TextBox TxtDoc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6735
      TabIndex        =   9
      Top             =   1440
      Width           =   1275
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormLancaCta.frx":202E
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Custo"
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Tag             =   "S"
      Top             =   1440
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebSubTot 
      DataField       =   "Custo"
      Height          =   225
      Left            =   9750
      TabIndex        =   18
      Tag             =   "S"
      Top             =   5295
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebJuros 
      DataField       =   "Custo"
      Height          =   255
      Left            =   9750
      TabIndex        =   20
      Tag             =   "S"
      Top             =   5790
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTotal 
      DataField       =   "Custo"
      Height          =   255
      Left            =   9750
      TabIndex        =   23
      Tag             =   "S"
      Top             =   6330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebCtaPagas 
      DataField       =   "Custo"
      Height          =   285
      Left            =   180
      TabIndex        =   25
      Tag             =   "S"
      Top             =   6195
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebEmiss 
      DataField       =   "Custo"
      Height          =   315
      Left            =   2595
      TabIndex        =   6
      Tag             =   "S"
      Top             =   1440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebVenc 
      DataField       =   "Custo"
      Height          =   315
      Left            =   3915
      TabIndex        =   7
      Tag             =   "S"
      Top             =   1440
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   90
      Top             =   5925
      Width           =   1305
   End
   Begin VB.Line Line2 
      X1              =   9660
      X2              =   11100
      Y1              =   6075
      Y2              =   6075
   End
   Begin VB.Line Line1 
      X1              =   9660
      X2              =   11100
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Shape Shape1 
      Height          =   1635
      Left            =   9660
      Top             =   5025
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Pedido"
      Height          =   195
      Left            =   180
      TabIndex        =   44
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   5715
      TabIndex        =   42
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nro Parcela"
      Height          =   195
      Left            =   1575
      TabIndex        =   30
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Emissão"
      Height          =   195
      Index           =   3
      Left            =   2565
      TabIndex        =   29
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   6900
      TabIndex        =   28
      Top             =   630
      Width           =   315
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Contas Pagas"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   26
      Top             =   5940
      Width           =   990
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   9720
      TabIndex        =   24
      Top             =   6120
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Historico"
      Height          =   195
      Left            =   8190
      TabIndex        =   22
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Juros"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   9750
      TabIndex        =   21
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   9720
      TabIndex        =   19
      Top             =   5070
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento"
      Height          =   195
      Left            =   6750
      TabIndex        =   16
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Vencimento"
      Height          =   195
      Index           =   15
      Left            =   3915
      TabIndex        =   15
      Top             =   1230
      Width           =   1230
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor"
      Height          =   195
      Index           =   5
      Left            =   5310
      TabIndex        =   14
      Top             =   1230
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Cliente"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   630
      Width           =   1170
   End
End
Attribute VB_Name = "FormLancaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub BtBoleto_Click()

With FormBoleto
    If Opt(2).value = True Then
        .TxtPedido = DBGrid1.Columns(0)
    Else
        .TxtPedido = DBGrid1.Columns(1)
    End If
    If DBGrid1.Columns(2) = "" Then
        MsgBox "Nenhuma parcela selecionada", vbInformation, App.Title
        CmdLimpa_Click
        Exit Sub
    End If
    .TxtPcl = DBGrid1.Columns(3)
    .TxtDoc = DBGrid1.Columns(14)
    .Opt(0) = Opt(0).value
    .Opt(1) = Opt(1).value
    .Opt(2) = Opt(2).value
'    .MebBanco = MebBanco
    .TxtEnt = txtCod
    .MebDtSaida = MebDtSaida
    .MebEmiss = DBGrid1.Columns(2)
    .MebVenc = DBGrid1.Columns(5)
    .MebValor = CCur(DBGrid1.Columns(7)) + CCur(DBGrid1.Columns(8))
    If Opt(0).value = True Then
        Set soma = Banco.OpenRecordset("SELECT DUP.Pedido,Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor  FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(1)))
    ElseIf Opt(1).value = True Then
            If Not IsNull(DBGrid1.Columns(13)) And DBGrid1.Columns(14) <> "" Then
                Set soma = Banco.OpenRecordset("SELECT Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Documento HAVING DUP.Documento Like '" & Trim(DBGrid1.Columns(14)) & "'")
            Else
                MsgBox "Campo documento não está preenchido", vbInformation, App.Title
            End If
        Else
            Set soma = Banco.OpenRecordset("SELECT DUP.Pedido,Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor  FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(0)))
    End If
    If soma.RecordCount > 0 Then
        .MebTotal = CCur(soma!SomaDeValor) + CCur(soma!somadejuros)
    Else
        .MebTotal = CCur(DBGrid1.Columns(7)) + CCur(DBGrid1.Columns(8))
    End If
End With

FormBoleto.Show vbModal

End Sub

Private Sub BtCarne_Click()

End Sub

Private Sub BtDupli_Click(Index As Integer)
Dim Completa As String
Dim TotExtenso As String
Dim soma As Recordset
Dim SQL As String

Completa = " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
'1=fisica
'2=juridica

'MDIPrincipal.ComRel.ShowPrinter
'MDIPrincipal.RptRel.ReportFileName = Caminho & "\Boleto.RPT"
If MebBanco <> "" Then
    FrmRelBoleto.rpBanco.Caption = MebBanco
    FrmRelBoleto.rpDtSaida.Caption = CStr(MebDtSaida)
End If

FrmRelBoleto.RpDtEmissao.Caption = CStr(MebDtSaida)

If Opt(0).value = True Then
        Set soma = Banco.OpenRecordset("SELECT DUP.Pedido,Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor  FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(1)))
        Set varPesquisa(1) = Banco.OpenRecordset("SELECT  DUP.Juros, DUP.Valor, ENT.Nro, ENT.Rg, ENT.CPF, ENT.Classificacao, ENT.CGC, ENT.InscrEst, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, Ent.Endereço_cob, Ent.Nro_cob, Ent.Cidade_cob, Ent.Uf_Cob, Ent.Cep_cob, DUP.NroPcl, DUP.DtVenc FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY ENT.CPF, ENT.Classificacao, ENT.CGC,  ENT.InscrEst, Dup.Juros, DUP.Valor, ENT.Nro, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Cep,  ENT.Rg, ENT.Fone, Ent.Endereço_cob, Ent.Nro_cob, Ent.Cidade_cob, Ent.Uf_Cob, Ent.Cep_cob, DUP.Pedido, DUP.Cond, DUP.NroPcl, DUP.DtVenc HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(1)) & " and Dup.Cond = False")
        If Index = 0 And varPesquisa(1).RecordCount > 0 Then
            If MebNota <> "" Then
                FrmRelBoleto.rpNota.Caption = MebNota
            End If
            If MebNota = "" Then
                FrmRelBoleto.rpOrdem.Caption = DBGrid1.Columns(1)
            Else
                FrmRelBoleto.rpOrdem.Caption = MebNota
            End If
            If MebTotalNota <> "" Then
                FrmRelBoleto.rpSoma.Caption = Format(MebTotalNota, "#,##0.00")
            Else
                FrmRelBoleto.rpSoma.Caption = Format(soma!SomaDeValor + soma!somadejuros, "#,##0.00")
            End If
            FrmRelBoleto.rpEmitenteCPF.Caption = IIf(varPesquisa(1)("Classificacao") = 1, IIf(IsNull(varPesquisa(1)("CPF")), "", Format(varPesquisa(1)("CPF"), "@@@.@@@.@@@-@@")), IIf(IsNull(varPesquisa(1)("CGC")), "", Format(varPesquisa(1)("CGC"), "@@.@@@.@@@/@@@@-@@")))
            FrmRelBoleto.VarEmpresa.Caption = NomeEmpresa
            FrmRelBoleto.VarEmpresa2.Caption = NomeEmpresa
            If varPesquisa(1)("Classificacao") = 2 Then
                If Not IsNull(varPesquisa(1)("InscrEst")) Then
                    FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("InscrEst")
                End If
            Else
                If Not IsNull(varPesquisa(1)("Rg")) Then
                    FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("Rg")
                End If
            End If
            
            If varPesquisa(1)("Endereço_cob") <> "" Then
                FrmRelBoleto.RptEnd.Caption = varPesquisa(1)("Endereço_cob")
                FrmRelBoleto.RptNro.Caption = varPesquisa(1)("Nro_cob")
                FrmRelBoleto.RptCidade.Caption = varPesquisa(1)("Cidade_cob")
                FrmRelBoleto.RptUf.Caption = varPesquisa(1)("Uf_cob")
                FrmRelBoleto.RptCep.Caption = varPesquisa(1)("Cep_cob")
            Else
                FrmRelBoleto.RptEnd.Caption = varPesquisa(1)("Endereco")
                FrmRelBoleto.RptNro.Caption = varPesquisa(1)("Nro")
                FrmRelBoleto.RptCidade.Caption = varPesquisa(1)("Cidade")
                FrmRelBoleto.RptUf.Caption = varPesquisa(1)("Uf")
                FrmRelBoleto.RptCep.Caption = varPesquisa(1)("Cep")
            End If
            
            If MebBanco = "" Then
                FrmRelBoleto.Label8.Visible = False
                FrmRelBoleto.Label9.Visible = False
                FrmRelBoleto.Label10.Visible = False
                FrmRelBoleto.Label11.Visible = False
                FrmRelBoleto.Label12.Visible = False
                FrmRelBoleto.Label13.Visible = False
                FrmRelBoleto.rpDtSaida.Visible = False
                FrmRelBoleto.rpBanco.Visible = False
            Else
                FrmRelBoleto.VarEmpresa2.Visible = False
                FrmRelBoleto.Line12.Visible = False
            End If
            FrmRelBoleto.SecPromissoria.Mostrar = False
            Set FrmRelBoleto.Relatorio.Recordset = varPesquisa(1)
            FrmRelBoleto.Config
'            Sql = "UPDATE DUP SET DUP.Tipo_Dup = 'DP' WHERE DUP.Pedido = " & CLng(DBGrid1.Columns(1)) & " AND DUP.Cond=False "
'            Banco.Execute Sql
        Else
            FrmRelCarne.rpTitulo.Caption = Cabecalho1
            FrmRelCarne.RpTitulo2.Caption = Cabecalho1
            FrmRelCarne.rpNome.Caption = txtCod & " - " & Left(DbcPesquisa, 40)
            FrmRelCarne.rpNome2.Caption = txtCod & " - " & Left(DbcPesquisa, 40)
'            FrmRelCarne.RpMens.Caption = mensagemRodape
'            FrmRelCarne.RpMens2.Caption = mensagemRodape
            FrmRelCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            FrmRelCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            Set FrmRelCarne.Relatorio.Recordset = varPesquisa(1)
            FrmRelCarne.Config
            
        End If
    ElseIf Opt(1).value = True Then
            If Not IsNull(DBGrid1.Columns(13)) And DBGrid1.Columns(14) <> "" Then
                Set soma = Banco.OpenRecordset("SELECT Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Documento, DUP.Entidade HAVING DUP.Documento Like '" & Trim(DBGrid1.Columns(14)) & "'  AND DUP.Entidade =" & txtCod)
                Set varPesquisa(1) = Banco.OpenRecordset("SELECT ENT.Classificacao, ENT.RG, ENT.CPF, ENT.CGC, ENT.Nro, ENT.InscrEst, ENT.Cep, ENT.Endereço_Cob, ENT.Nro_Cob, ENT.Cidade_Cob, ENT.Uf_Cob, ENT.Cep_Cob, DUP.Documento, DUP.Valor, DUP.Juros, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone, DUP.NroPcl, DUP.DtVenc, DUP.Entidade FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade GROUP BY ENT.Classificacao, ENT.RG, ENT.CPF, ENT.CGC, ENT.Nro, ENT.InscrEst, ENT.Cep, ENT.Endereço_Cob, ENT.Nro_Cob, ENT.Cidade_Cob, ENT.Uf_Cob, ENT.Cep_Cob, DUP.Documento, DUP.Valor, DUP.Juros, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone, DUP.NroPcl, DUP.DtVenc, DUP.Entidade, DUP.Pedido, DUP.Cond HAVING Dup.Documento like '" & Trim(DBGrid1.Columns(14)) & "' AND DUP.Entidade =" & txtCod)
                SQL = "SELECT Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Documento, DUP.Entidade HAVING DUP.Documento Like '" & Trim(DBGrid1.Columns(14)) & "'  AND DUP.Entidade =" & txtCod
                SQL = "SELECT ENT.Classificacao, ENT.RG, ENT.CPF, ENT.CGC, ENT.Nro, ENT.InscrEst, ENT.Cep, ENT.Endereço_Cob, ENT.Nro_Cob, ENT.Cidade_Cob, ENT.Uf_Cob, ENT.Cep_Cob, DUP.Documento, DUP.Valor, DUP.Juros, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone, DUP.NroPcl, DUP.DtVenc, DUP.Entidade FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade GROUP BY ENT.Classificacao, ENT.RG, ENT.CPF, ENT.CGC, ENT.Nro, ENT.InscrEst, ENT.Cep, ENT.Endereço_Cob, ENT.Nro_Cob, ENT.Cidade_Cob, ENT.Uf_Cob, ENT.Cep_Cob, DUP.Documento, DUP.Valor, DUP.Juros, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone, DUP.NroPcl, DUP.DtVenc, DUP.Entidade, DUP.Pedido, DUP.Cond HAVING Dup.Documento like '" & Trim(DBGrid1.Columns(14)) & "' AND DUP.Entidade =" & txtCod
                If Index = 0 Then
                    If MebNota <> "" Then
                        FrmRelBoleto.rpNota.Caption = MebNota
                    Else
                        FrmRelBoleto.rpNota.Caption = " " ' Trim(DBGrid1.Columns(14))
                    End If
                    FrmRelBoleto.rpOrdem.Caption = Trim(DBGrid1.Columns(14))
                    If MebTotalNota <> "" Then
                            FrmRelBoleto.rpSoma.Caption = Format(MebTotalNota, "#,##0.00")
                    Else
                            FrmRelBoleto.rpSoma.Caption = Format(soma!SomaDeValor + soma!somadejuros, "#,##0.00")
                    End If
                                    
                    FrmRelBoleto.rpEmitenteCPF.Caption = IIf(varPesquisa(1)("Classificacao") = 1, IIf(IsNull(varPesquisa(1)("CPF")), "", Format(varPesquisa(1)("CPF"), "@@@.@@@.@@@-@@")), IIf(IsNull(varPesquisa(1)("CGC")), "", Format(varPesquisa(1)("CGC"), "@@.@@@.@@@/@@@@-@@")))
                    FrmRelBoleto.VarEmpresa.Caption = NomeEmpresa
                    FrmRelBoleto.VarEmpresa2.Caption = NomeEmpresa
                    
                    If varPesquisa(1)("Classificacao") = 2 Then
                        If Not IsNull(varPesquisa(1)("InscrEst")) Then
                            FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("InscrEst")
                        End If
                    Else
                        If Not IsNull(varPesquisa(1)("Rg")) Then
                            FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("Rg")
                        End If
                    End If
                    
                    If varPesquisa(1)("Endereço_cob") <> "" Then
                        FrmRelBoleto.RptEnd.Caption = varPesquisa(1)("Endereço_cob")
                        FrmRelBoleto.RptNro.Caption = varPesquisa(1)("Nro_cob")
                        FrmRelBoleto.RptCidade.Caption = varPesquisa(1)("Cidade_cob")
                        FrmRelBoleto.RptUf.Caption = varPesquisa(1)("Uf_cob")
                        FrmRelBoleto.RptCep.Caption = varPesquisa(1)("Cep_cob")
                    Else
                        FrmRelBoleto.RptEnd.Caption = varPesquisa(1)("Endereco")
                        FrmRelBoleto.RptNro.Caption = varPesquisa(1)("Nro")
                        FrmRelBoleto.RptCidade.Caption = varPesquisa(1)("Cidade")
                        FrmRelBoleto.RptUf.Caption = varPesquisa(1)("Uf")
                        FrmRelBoleto.RptCep.Caption = varPesquisa(1)("Cep")
                    End If
                    
                    
                    If MebBanco = "" Then
                        FrmRelBoleto.Label8.Visible = False
                        FrmRelBoleto.Label9.Visible = False
                        FrmRelBoleto.Label10.Visible = False
                        FrmRelBoleto.Label11.Visible = False
                        FrmRelBoleto.Label12.Visible = False
                        FrmRelBoleto.Label13.Visible = False
                        FrmRelBoleto.rpDtSaida.Visible = False
                        FrmRelBoleto.rpBanco.Visible = False
                    Else
                        FrmRelBoleto.VarEmpresa2.Visible = False
                        FrmRelBoleto.Line12.Visible = False
                    End If
                    FrmRelBoleto.SecPromissoria.Mostrar = False
                    
                    Set FrmRelBoleto.Relatorio.Recordset = varPesquisa(1)
                    FrmRelBoleto.Config
                Else
                    FrmRelCarne.rpTitulo.Caption = Cabecalho1
                    FrmRelCarne.RpTitulo2.Caption = Cabecalho1
                    FrmRelCarne.rpNome.Caption = Left(DbcPesquisa, 40)
                    FrmRelCarne.rpNome2.Caption = Left(DbcPesquisa, 40)
        '            FrmRelCarne.RpMens.Caption = mensagemRodape
        '            FrmRelCarne.RpMens2.Caption = mensagemRodape
                    FrmRelCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                    FrmRelCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                    Set FrmRelCarne.Relatorio.Recordset = varPesquisa(1)
                    FrmRelCarne.Config
                    
                End If
            
            Else
                MsgBox "Campo documento não está preenchido", vbInformation, App.Title
            End If
            If Trim(DBGrid1.Columns(14)) <> "" Then
                SQL = "UPDATE DUP SET DUP.Tipo_Dup = 'DP' WHERE DUP.Documento = " & """" & Trim(DBGrid1.Columns(14)) & """" & " AND DUP.Cond=False"
                Banco.Execute SQL
            End If
        Else
            Set varPesquisa(1) = Banco.OpenRecordset("SELECT Ent.Classificacao, ENT.Rg, ENT.Nro, ENT.CPF, ENT.CGC, ENT.InscrEst, Ent.Cep, Dup.CodDup, DUP.Documento, DUP.Valor, DUP.Juros, ENT.Nome, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone,  Ent.Endereço_cob, Ent.Nro_cob, Ent.Cidade_cob, Ent.Uf_Cob, Ent.Cep_cob, DUP.NroPcl, DUP.DtVenc FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY Ent.Classificacao, ENT.CPF, ENT.CGC, ENT.InscrEst, Ent.Cep, Dup.CodDup, DUP.Documento, DUP.Valor, Dup.Juros, ENT.Nome,  ENT.Rg, ENT.Nro, ENT.Endereco, ENT.Cidade, ENT.Uf, ENT.Fone,  Ent.Endereço_cob, Ent.Nro_cob, Ent.Cidade_cob, Ent.Uf_Cob, Ent.Cep_cob, DUP.Pedido, DUP.Cond, DUP.NroPcl, DUP.DtVenc HAVING DUP.CodDup = " & Trim(DBGrid1.Columns(0)))
            If Index = 0 Then
                If MebNota <> "" Then
                    FrmRelBoleto.rpNota.Caption = MebNota
                End If
                If Trim(DBGrid1.Columns(1)) <> "" Then
                    FrmRelBoleto.rpOrdem.Caption = Trim(DBGrid1.Columns(1))
                Else
                    FrmRelBoleto.rpOrdem.Caption = Trim(DBGrid1.Columns(13))
                End If
                
                If MebTotalNota <> "" Then
                    FrmRelBoleto.rpSoma.Caption = Format(MebTotalNota, "#,##0.00")
                Else
                    FrmRelBoleto.rpSoma.Caption = Format(varPesquisa(1)!Valor + varPesquisa(1)!Juros, "#,##0.00")
                End If
                FrmRelBoleto.VarEmpresa.Caption = NomeEmpresa
                FrmRelBoleto.VarEmpresa2.Caption = NomeEmpresa
                           
                FrmRelBoleto.rpEmitenteCPF.Caption = IIf(varPesquisa(1)("Classificacao") = 1, IIf(IsNull(varPesquisa(1)("CPF")), "", Format(varPesquisa(1)("CPF"), "@@@.@@@.@@@-@@")), IIf(IsNull(varPesquisa(1)("CGC")), "", Format(varPesquisa(1)("CGC"), "@@.@@@.@@@/@@@@-@@")))
                
                If varPesquisa(1)("Classificacao") = 2 Then
                    If Not IsNull(varPesquisa(1)("InscrEst")) Then
                        FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("InscrEst")
                    End If
                Else
                    If Not IsNull(varPesquisa(1)("Rg")) Then
                        FrmRelBoleto.rpInscricao.Caption = varPesquisa(1)("Rg")
                    End If
                End If
                
                If varPesquisa(1)("Endereço_cob") <> "" Then
                    FrmRelBoleto.RptEnd = varPesquisa(1)("Endereço_cob")
                    FrmRelBoleto.RptNro = varPesquisa(1)("Nro_cob")
                    FrmRelBoleto.RptCidade = varPesquisa(1)("Cidade_cob")
                    FrmRelBoleto.RptUf = varPesquisa(1)("Uf_cob")
                    FrmRelBoleto.RptCep = varPesquisa(1)("Cep_cob")
                Else
                    FrmRelBoleto.RptEnd.Caption = varPesquisa(1)("Endereco")
                    FrmRelBoleto.RptNro.Caption = varPesquisa(1)("Nro")
                    FrmRelBoleto.RptCidade.Caption = varPesquisa(1)("Cidade")
                    FrmRelBoleto.RptUf.Caption = varPesquisa(1)("Uf")
                    FrmRelBoleto.RptCep.Caption = varPesquisa(1)("Cep")
                End If
                                
                If MebBanco = "" Then
                    FrmRelBoleto.Label8.Visible = False
                    FrmRelBoleto.Label9.Visible = False
                    FrmRelBoleto.Label10.Visible = False
                    FrmRelBoleto.Label11.Visible = False
                    FrmRelBoleto.Label12.Visible = False
                    FrmRelBoleto.Label13.Visible = False
                    FrmRelBoleto.rpDtSaida.Visible = False
                    FrmRelBoleto.rpBanco.Visible = False
                Else
                    FrmRelBoleto.VarEmpresa2.Visible = False
                    FrmRelBoleto.Line12.Visible = False
                End If
                FrmRelBoleto.SecPromissoria.Mostrar = False
                
                Set FrmRelBoleto.Relatorio.Recordset = varPesquisa(1)
                FrmRelBoleto.Config
                SQL = "UPDATE DUP SET DUP.Tipo_Dup = 'DP' WHERE  DUP.CodDup = " & Trim(DBGrid1.Columns(0)) & " And Dup.Cond = False"
                Banco.Execute SQL
            Else
                FrmRelCarne.rpTitulo.Caption = Cabecalho1
                FrmRelCarne.RpTitulo2.Caption = Cabecalho1
                FrmRelCarne.rpNome.Caption = Left(DbcPesquisa, 40)
                FrmRelCarne.rpNome2.Caption = Left(DbcPesquisa, 40)
    '            FrmRelCarne.RpMens.Caption = mensagemRodape
    '            FrmRelCarne.RpMens2.Caption = mensagemRodape
                FrmRelCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                FrmRelCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                Set FrmRelCarne.Relatorio.Recordset = varPesquisa(1)
                FrmRelCarne.Config
                
            End If

End If


End Sub


Private Sub btPromi_Click()

        FrmRelBoleto.SecPromissoria.Mostrar = True
        FrmRelBoleto.Det.Mostrar = False
        FrmRelBoleto.NroPedido.Caption = Format(Datdados(5).Recordset("CodPed"), "000000")
        FrmRelBoleto.ExtensoDia.Caption = Day(Parcelas("DtVenc")) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(Parcelas("DtVenc")))) & " DO ANO DE " & Year(Parcelas("DtVenc"))
        FrmRelBoleto.QuantiaExtenso.Caption = UCase(PassaExtenso(Datdados(5).Recordset("Total"))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
        FrmRelBoleto.ValorPromissoria.Caption = Format(Datdados(5).Recordset("Total"), "#,#0.00")
        FrmRelBoleto.DiaVenc.Caption = Day(Parcelas("DtVenc"))
        FrmRelBoleto.MesVenc.Caption = UCase(RetornaMes(Month(Parcelas("DtVenc"))))
        FrmRelBoleto.AnoVenc.Caption = Year(Parcelas("DtVenc"))
        FrmRelBoleto.Emitente.Caption = varPesquisa(1).Fields("Nome")
        FrmRelBoleto.Emitente2.Caption = varPesquisa(1).Fields("Nome")
        FrmRelBoleto.EmitenteCpf.Caption = varPesquisa(1).Fields("CPF")
        FrmRelBoleto.Pagavel.Caption = CidadeEmpresa
        FrmRelBoleto.EmitenteEndereco.Caption = varPesquisa(1).Fields("Endereco") & " - " & varPesquisa(1).Fields("Nro")
        FrmRelBoleto.NomeEmpresa.Caption = NomeInscr
        FrmRelBoleto.EmpresaCGCCPF.Caption = Cnpj
        FrmRelBoleto.RpDia.Caption = Day(Date)
        FrmRelBoleto.RpMes.Caption = UCase(RetornaMes(Month(Date)))
        FrmRelBoleto.RpAno.Caption = Year(Date)


End Sub



Private Sub ChkJuros_Click()
If DbcPesquisa.Text <> "" Then
    Atualiza
End If
End Sub

Private Sub ChkPago_Click()
If DbcPesquisa.Text <> "" Then
    Atualiza
End If
End Sub

Private Sub CmbTipo_GotFocus()
    If Dir(App.Path & "\Tipo.Txt") <> "" Then
        CmbTipo.ListIndex = 1
    Else
        CmbTipo.ListIndex = 0
    End If
End Sub

Private Sub CmdCancela_Click()
    Unload Me
End Sub

Private Sub CmdExcluir_Click()
Dim SQL As String
If Datdados(0).Recordset.RecordCount <> 0 Then
    If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
        Open Caminho & "\Delete.txt" For Append As #1
        Print #1, "Parcela; " & Str(DBGrid1.Columns(1)) & "/"; Str(DBGrid1.Columns(3)) & "; " & DbcPesquisa.Text & "; " & Date & "; " & Time & "; " & Maquina & "; " & Alinhar(DBGrid1.Columns(7), 10, 2)
        Close #1
        If DBGrid1.Columns(13) = "CRE" Then
            SQL = "DELETE REC.CodEnt, REC.Data, REC.Valor FROM REC WHERE (((REC.CodEnt)=" & txtCod & ") AND ((REC.Data)=#" & Format(DBGrid1.Columns(2), "mm/dd/yyyy") & "#) AND ((REC.Valor)=" & Numero4(DBGrid1.Columns(7)) & "))"
            Banco.Execute SQL
        End If
        Datdados(0).Recordset.Delete
        Datdados(0).Refresh
        
        SomaGrid
    End If
End If
End Sub

Private Sub CmdLancar_Click()
Dim SQL As String
Dim Pedido As Long
On Error GoTo TrataErro
If IsNull(MebValor) Or MebValor = "" Then
    MsgBox "Valor não pode ser nulo !", vbExclamation, App.Title
    MebValor.SetFocus
    Exit Sub
End If

If IsNull(MebEmiss) Or MebEmiss = "" Then
    MsgBox "Data Emissão não pode ser nulo !", vbExclamation, App.Title
    MebEmiss.SetFocus
    Exit Sub
End If

If IsNull(MebVenc) Or MebVenc = "" Then
    MsgBox "Vencimento não pode ser nulo !", vbExclamation, App.Title
    MebVenc.SetFocus
    Exit Sub
End If

If TxtPedido = "" Then
    TxtPedido = 0
End If

If TxtPcl = "" Then
    TxtPcl = 1
End If

SQL = "INSERT INTO DUP ( DtEmiss, NroPcl, Entidade, DtVenc, Valor, Cond, Obs, Tipo_Dup, Pedido, Documento  ) VALUES ( "
SQL = SQL & "#" & Format(MebEmiss, "mm/dd/yyyy") & "# , "
SQL = SQL & TxtPcl & ", "
SQL = SQL & Datdados(1).Recordset.Fields("CodEntidade") & " , #"
SQL = SQL & Format(MebVenc, "mm/dd/yyyy") & "# , "
SQL = SQL & Numero2(MebValor.Text) & " ,"
SQL = SQL & 0 & " ,"
SQL = SQL & """" & CmbObs & """" & " ,"
SQL = SQL & """" & CmbTipo & """" & " ,"
SQL = SQL & TxtPedido & " ,"
SQL = SQL & """" & TxtDoc & """" & ")"
Banco.Execute SQL

CmdLimpa_Click
Atualiza
Datdados(0).Recordset.MoveLast
'SomaGrid
DbcPesquisa.SetFocus

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdLimpa_Click()
'DbcPesquisa = ""
MebValor = ""
TxtDoc = ""
CmbObs = ""
TxtPcl = ""
CmbTipo.ListIndex = 1
MebEmiss = Date
MebVenc = Date + 30
TxtPedido = 0
TxtPcl = 1

Atualiza
'SomaGrid
DbcPesquisa.SetFocus

End Sub


Private Sub Command1_Click()

End Sub

Private Sub DbcPesquisa_LostFocus()
If DbcPesquisa.Text <> "" Then
    If Not IsNumeric(DbcPesquisa.Text) Then
        If IsNumeric(DbcPesquisa.BoundText) Then
            Pesquisa DbcPesquisa.BoundText
        End If
    Else
        Pesquisa DbcPesquisa.Text
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Pesquisa(Codigo As Integer)

Dim Criterio As String
    Frame.Visible = True
    DoEvents
   Criterio = Datdados(1).Recordset.Fields("CodEntidade").Name & " = "
   Criterio = Criterio & Codigo
   Datdados(1).Recordset.FindFirst Criterio
   If Datdados(1).Recordset.NoMatch Then
       MsgBox "Entidade não Localizado", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
       DbcPesquisa.Text = Datdados(1).Recordset.Fields("Nome")
       txtCod.Text = Datdados(1).Recordset.Fields("CodEntidade")
       Atualiza
       CmbTipo.SetFocus
   End If

End Sub



Private Sub Form_Activate()

Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

'Atualiza
If Admin = True Then
    CmdExcluir.Visible = True
Else
    CmdExcluir.Visible = False
End If

If LerINI("FINANCEIRO", "NaoExluirDup", App.Path & "\config.ini") <> "" Then
    If LerINI("FINANCEIRO", "NaoExluirDup", App.Path & "\config.ini") = "1" Then
        CmdExcluir.Visible = False
    End If
End If

CarregaTipo

End Sub

Private Sub Form_Load()
Dim Componente As Variant
'Set Registros = Banco.OpenRecordset("SELECT *  FROM Dup")


For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

Datdados(1).RecordSource = "SELECT CodEntidade , Nome, Prazo, Juros FROM Ent  WHERE Tipo = 1 or Tipo >= 5 Order By Nome"
Datdados(1).Refresh

MebDtSaida.Text = Date
MebEmiss = Date
MebVenc = Date + 30
TxtPedido = 0
TxtPcl = 1

End Sub


Private Sub SomaGrid()
Dim soma As Recordset
Dim Total As Currency
Dim flag As Boolean
Dim DtTmp As Date
Dim DtAtraso As Integer
Dim Correcao As Currency
Dim TotJuros As Currency
Dim TotCtaPagas As Currency
Dim TotSaldo As Currency
flag = True
TotJuros = 0
TotCtaPagas = 0
Set soma = Datdados(0).Recordset.Clone
If soma.RecordCount <> 0 Then
    soma.MoveFirst
    Do Until soma.Eof
        If GetInputState() Then DoEvents

        If soma.Fields("Cond") = 0 Then
            DtTmp = soma.Fields("DtVenc")
            If flag Then
                DtTmp = soma.Fields("DtVenc")
'                If DtTmp < Date - 30 Then
'                    LblAtraso.Caption = (Date - DtTmp) - IIf(IsNull(Datdados(1).Recordset.Fields("Prazo")), 0, Datdados(1).Recordset.Fields("Prazo"))
'                Else
'                    LblAtraso.Caption = "0"
'                End If
                flag = False
            End If
            If ChkJuros.value = 1 Then
                DtAtraso = (Date - DtTmp)
                Correcao = soma.Fields("Valor") * (((Datdados(1).Recordset.Fields("Juros") / 100) / 30) * DtAtraso)
                soma.Edit
                If Correcao > 0 Then
                    soma.Fields("Juros") = Correcao
                    TotJuros = TotJuros + Correcao
                Else
                    soma.Fields("Juros") = 0
                End If
                soma.Update
            Else
                soma.Edit
                soma!Juros = 0
                soma.Update
            End If
            Total = Total + soma.Fields("Valor")
        Else
            If soma.Fields("OBS") = "SALDO" Then
            
            Else
                TotCtaPagas = TotCtaPagas + soma.Fields("Valor")
            End If
        End If
        soma.MoveNext
    Loop
End If
MebSubTot = Total
MebJuros = TotJuros
MebTotal = CDec(MebSubTot) + TotJuros
MebCtaPagas = TotCtaPagas
Frame.Visible = False
End Sub


Private Sub Atualiza()
If ChkPago.value = 1 Then
    Datdados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " ORDER BY DtEmiss,CodDup"
Else
    Datdados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " AND COND = FALSE ORDER BY DtEmiss, CodDup"
End If
Datdados(0).Refresh
SomaGrid
If Datdados(0).Recordset.RecordCount <> 0 Then
    Datdados(0).Recordset.MoveLast
End If

End Sub



