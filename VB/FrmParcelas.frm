VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmParcelas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   9120
   Begin VB.CommandButton BtBoletoBco 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Boleto via Banco"
      Height          =   345
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7290
      Width           =   1890
   End
   Begin VB.CommandButton BtImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Promissoria c/ Parcelas"
      Height          =   345
      Index           =   7
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7290
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   9390
      TabIndex        =   44
      Top             =   7965
      Width           =   9420
   End
   Begin VB.CommandButton BtImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Canê com Capa"
      Height          =   345
      Index           =   5
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7290
      Width           =   1890
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000002&
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9090
      TabIndex        =   41
      Top             =   0
      Width           =   9150
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão de Contas a Receber"
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
         TabIndex        =   42
         Top             =   45
         Width           =   6045
      End
   End
   Begin VB.CheckBox ChkPgto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lançar já quitado"
      Height          =   330
      Left            =   6795
      TabIndex        =   40
      Top             =   2880
      Width           =   1590
   End
   Begin VB.CommandButton BtImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Canê &Promissoria"
      Height          =   345
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6435
      Width           =   1890
   End
   Begin VB.CommandButton BtImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Promissoria"
      Height          =   345
      Index           =   4
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6030
      Width           =   1890
   End
   Begin VB.CommandButton BtImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Canê com Promissoria"
      Height          =   345
      Index           =   3
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6885
      Width           =   1890
   End
   Begin VB.TextBox TxtNome 
      Height          =   330
      Left            =   1455
      TabIndex        =   36
      Top             =   4995
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.ComboBox CmbVenc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3150
      TabIndex        =   7
      Top             =   2340
      Width           =   3480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   5220
      TabIndex        =   29
      Top             =   720
      Width           =   3750
      Begin VB.Label LblValor 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   30
         Top             =   165
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impressão Parcelas"
      Height          =   1095
      Left            =   4545
      TabIndex        =   24
      Top             =   5985
      Width           =   4200
      Begin VB.CommandButton BtBoleto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Boleto"
         Height          =   720
         Left            =   225
         Picture         =   "FrmParcelas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton BtImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Formulário"
         Height          =   720
         Index           =   0
         Left            =   3195
         Picture         =   "FrmParcelas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton BtImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Bobina"
         Height          =   720
         Index           =   1
         Left            =   2205
         Picture         =   "FrmParcelas.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton BtImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Página"
         Height          =   720
         Index           =   2
         Left            =   1215
         Picture         =   "FrmParcelas.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox TxtDoc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2205
      TabIndex        =   9
      Top             =   2925
      Width           =   1740
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmParcelas.frx":10C0
      Left            =   2205
      List            =   "FrmParcelas.frx":10E2
      TabIndex        =   1
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Data DatDados 
      Caption         =   "Parcelas"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Index           =   0
      Left            =   2475
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4455
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSMask.MaskEdBox CodDoc 
      Height          =   270
      Left            =   180
      TabIndex        =   14
      Top             =   7155
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
   Begin VB.CommandButton OK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Gerar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2865
      Width           =   1305
   End
   Begin MSMask.MaskEdBox MebEntrada 
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   1710
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebParcelas 
      Height          =   315
      Left            =   2130
      TabIndex        =   3
      Top             =   1710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid GridParc 
      Bindings        =   "FrmParcelas.frx":1111
      Height          =   2535
      Left            =   225
      OleObjectBlob   =   "FrmParcelas.frx":112B
      TabIndex        =   12
      Top             =   3375
      Width           =   8535
   End
   Begin MSMask.MaskEdBox MebDia 
      Height          =   315
      Left            =   225
      TabIndex        =   5
      Top             =   2340
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebValorPcl 
      Height          =   315
      Left            =   225
      TabIndex        =   8
      Top             =   2925
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebVlrPed 
      Height          =   405
      Left            =   225
      TabIndex        =   0
      Top             =   1035
      Width           =   1875
      _ExtentX        =   3307
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
   Begin MSMask.MaskEdBox TxtMes 
      Height          =   315
      Left            =   1695
      TabIndex        =   6
      Top             =   2340
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebIndices 
      Height          =   315
      Left            =   3735
      TabIndex        =   4
      Top             =   1710
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CodEnt 
      Height          =   270
      Left            =   180
      TabIndex        =   34
      Top             =   6615
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Entidade"
      Height          =   195
      Left            =   180
      TabIndex        =   35
      Top             =   6390
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indices = 0 (zero) para não cobrar juros"
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5625
      TabIndex        =   33
      Top             =   1890
      Width           =   2760
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vencimentos"
      Height          =   195
      Left            =   3150
      TabIndex        =   31
      Top             =   2115
      Width           =   915
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Se deixar o dia e mês de pagamento sem preencher será assumido a data do dia"
      ForeColor       =   &H80000007&
      Height          =   720
      Left            =   6840
      TabIndex        =   28
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indices"
      Height          =   195
      Left            =   3735
      TabIndex        =   27
      Top             =   1485
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês Inicio"
      Height          =   195
      Left            =   1665
      TabIndex        =   23
      Top             =   2115
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento"
      Height          =   195
      Left            =   2205
      TabIndex        =   22
      Top             =   2700
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   2205
      TabIndex        =   21
      Top             =   765
      Width           =   315
   End
   Begin VB.Label LblOk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Left            =   225
      TabIndex        =   20
      Top             =   6075
      Width           =   2220
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Pedido"
      Height          =   195
      Left            =   225
      TabIndex        =   19
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
      Height          =   195
      Left            =   225
      TabIndex        =   18
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   225
      TabIndex        =   17
      Top             =   2700
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dia Pagamento"
      Height          =   195
      Left            =   225
      TabIndex        =   16
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Parcelas"
      Height          =   195
      Left            =   2160
      TabIndex        =   15
      Top             =   1485
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Pedido"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   6930
      Width           =   495
   End
End
Attribute VB_Name = "FrmParcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Parcelas As Recordset
'Public CodEnt As Integer
Dim Ind As Recordset
Dim VarIcms As Boolean
Dim VarPclAuto As Boolean
Dim VarDiaFixo As String
Dim VarEnt As Recordset
Public Senha As Boolean
Dim PassWord As String

Private Sub BtBoleto_Click()
With FormBoleto
    .TxtPedido = CodDoc
    .TxtPcl = GridParc.Columns(0)
    .TxtDoc = ""
    .OPt(0).value = True
    .MebDtSaida = Date
    .MebEmiss = GridParc.Columns(1)
    .MebVenc = GridParc.Columns(3)
    .MebValor = GridParc.Columns(4)
    .MebTotal = LblValor
End With

FormBoleto.Show vbModal
                     
End Sub

Private Sub CmbPromi_Click()
'ImprimiPromiMat
End Sub


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtBoletoBco_Click()
Dim VarEmitente As String
Dim VarEndEmitente As String
Dim VarNroemitente As String
Dim VarCidadeEmitente As String
Dim VarUfEmitente As String
Dim VarBairroEmitente As String
Dim VarCont As Integer
Dim VarCnpj  As String
VarCont = 1


'            FormBaixarDupli.DatDados(0).Recordset.MoveFirst
'            Do Until FormBaixarDupli.DatDados(0).Recordset.Eof
'                If FormBaixarDupli.DatDados(0).Recordset("PgtTemp") = -1 Then
'                    FormBaixarDupli.DatDados(0).Recordset.Edit
'                    FormBaixarDupli.DatDados(0).Recordset.Fields("Tipo_Dup") = "BO"
'                    FormBaixarDupli.DatDados(0).Recordset.Fields("DtVenc") = Format(MebVenc, "dd/mm/yyyy")
'                    FormBaixarDupli.DatDados(0).Recordset.Update
'
'                End If
'                FormBaixarDupli.DatDados(0).Recordset.MoveNext
'            Loop
            
        If MsgBox("Confirma Emissão Boleto ?", vbYesNo, App.Title) = vbYes Then
            VarCnpj = LerINI("Config", "Cnpj", App.Path & "\config.ini")
        
            VarEmitente = LerINI("Emitente", "RazaoSocial", App.Path & "\nfe.ini")
            VarEndEmitente = LerINI("Emitente", "Logradouro", App.Path & "\nfe.ini")
            VarNroemitente = LerINI("Emitente", "Nro", App.Path & "\nfe.ini")
            VarCidadeEmitente = LerINI("Emitente", "Cidade", App.Path & "\nfe.ini")
            VarUfEmitente = LerINI("Emitente", "Uf", App.Path & "\nfe.ini")
            VarBairroEmitente = LerINI("Emitente", "Bairro", App.Path & "\nfe.ini")
            
            
'            Open App.Path & "\Titulos.ini" For Output As #1
            If Dir(App.Path & "\Titulos.Ini") <> "" Then
               Kill App.Path & "\Titulos.ini"
            End If
            
            If Dir(App.Path & "\Cedente.Ini") <> "" Then
               Kill App.Path & "\Cedente.ini"
            End If
            
            
'            AbreAcbr "Boleto.LimparLista"
           
            EscreveINI "CEDENTE", "Nome", VarEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "CNPJCPF", VarCnpj, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Logradouro", VarEndEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Numero", VarNroemitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Bairro", VarBairroEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Cidade", VarCidadeEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Uf", VarUfEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "Cidade", VarCidadeEmitente, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "CEP", LerINI("Emitente", "CEP", App.Path & "\nfe.ini"), App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "RespEmis", 0, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "TipoPessoa", 1, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "CodigoCedente", LerINI("Boleto", "Convenio", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "LayOutBol", 3, App.Path & "\Cedente.ini"
            EscreveINI "CEDENTE", "CaracTitulo", 0, App.Path & "\Cedente.ini"
            
            EscreveINI "CONTA", "Conta", LerINI("Boleto", "Conta", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            EscreveINI "CONTA", "DigitoConta", LerINI("Boleto", "DigConta", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            EscreveINI "CONTA", "Agencia", LerINI("Boleto", "Agencia", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            EscreveINI "CONTA", "DigitoAgencia", LerINI("Boleto", "DigAgencia", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            
            EscreveINI "BANCO", "Numero", LerINI("Boleto", "Banco", App.Path & "\Config.ini"), App.Path & "\Cedente.ini"
            EscreveINI "BANCO", "CNAB", 1, App.Path & "\Cedente.ini"
            
            DoEvents
            AbreAcbr "Boleto.ConfigurarDados(" & App.Path & "\cedente.ini" & ")"
            
            
            DatDados(0).Recordset.MoveFirst

'            BOLETO.IncluirTitulos("c:\titulos.ini", "I")
            While Not DatDados(0).Recordset.Eof
            
                EscreveINI "TITULO" & VarCont, "NumeroDocumento", CodDoc, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "NossoNumero", DatDados(0).Recordset("CodDup"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "Carteira", "9", App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "ValorDocumento", DatDados(0).Recordset.Fields("Valor"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "Sacado.NomeSacado", FrmPedidos.Entidade("Nome"), App.Path & "\Titulos.ini"
                If FrmPedidos.Entidade("Classificacao") = 1 Then
                    EscreveINI "TITULO" & VarCont, "Sacado.CNPJCPF", FrmPedidos.Entidade("Cpf"), App.Path & "\Titulos.ini"
                Else
                    EscreveINI "TITULO" & VarCont, "Sacado.CNPJCPF", FrmPedidos.Entidade("Cgc"), App.Path & "\Titulos.ini"
                End If
                EscreveINI "TITULO" & VarCont, "Sacado.Logradouro", FrmPedidos.Entidade("Endereco"), App.Path & "\Titulos.ini"
                If Not IsNull(FrmPedidos.Entidade("Nro")) Then
                    EscreveINI "TITULO" & VarCont, "Sacado.Numero", FrmPedidos.Entidade("Nro"), App.Path & "\Titulos.ini"
                Else
                    EscreveINI "TITULO" & VarCont, "Sacado.Numero", "", App.Path & "\Titulos.ini"
                End If
                If Not IsNull(FrmPedidos.Entidade("Bairro")) Then
                    EscreveINI "TITULO" & VarCont, "Sacado.Bairro", FrmPedidos.Entidade("Bairro"), App.Path & "\Titulos.ini"
                Else
                    EscreveINI "TITULO" & VarCont, "Sacado.Bairro", "", App.Path & "\Titulos.ini"
                End If
                EscreveINI "TITULO" & VarCont, "Sacado.Cidade", FrmPedidos.Entidade("Cidade"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "Sacado.Uf", FrmPedidos.Entidade("Uf"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "Sacado.CEP", FrmPedidos.Entidade("Cep"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "ValorMoraJuros", 0, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "Vencimento", DatDados(0).Recordset("DtVenc"), App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "DataDocumento", Date, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "DataProcessamento", Date, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "DataDesconto", Date, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "DataMoraJuros", Date, App.Path & "\Titulos.ini"
                EscreveINI "TITULO" & VarCont, "DataProtesto", Date, App.Path & "\Titulos.ini"
                
                DatDados(0).Recordset.MoveNext
                VarCont = VarCont + 1
            Wend

            DoEvents
            AbreAcbr "Boleto.IncluirTitulos(" & App.Path & "\Titulos.ini" & ")"
            MsgBox "Boleto gerado !!!", vbInformation, App.Title
            
            AbreAcbr "Boleto.Imprimir"
            
            If MsgBox("Imprimir Boleto ?", vbYesNo, App.Title) = vbYes Then
                    
                ShellExecute 0&, vbNullString, "c:\SidComp\ACBrMonitorPLUS\boleto\boleto.pdf", _
                      vbNullString, vbNullString, vbNormalFocus
                    
            End If
            If Dir("c:\SidComp\ACBrMonitorPLUS\Sai.txt") <> "" Then
    
                Kill "c:\SidComp\ACBrMonitorPLUS\Sai.txt"
            End If
            
            Unload Me
        End If

End Sub

Private Sub BtIm_Click(Index As Integer)

End Sub

Private Sub DatDados_Reposition(Index As Integer)
Dim Total As Recordset
If CodDoc <> "" Then
    Set Total = Banco.OpenRecordset("SELECT DUP.Pedido, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CodDoc)
    If Total.RecordCount <> 0 Then
        If CCur(MebVlrPed) <= Round(Total("SomadeValor"), 2) Then
            BtImprimir(0).Enabled = True
            BtImprimir(1).Enabled = True
            BtImprimir(2).Enabled = True
            BtBoleto.Enabled = True
            LblOk.Caption = "Parcelas Corretas !"
        Else
            BtImprimir(0).Enabled = False
            BtImprimir(1).Enabled = False
            BtImprimir(2).Enabled = False
            BtBoleto.Enabled = False
            LblOk.Caption = "Parcelas Incorretas !"
        End If
        LblValor.Caption = Format(Total("SomadeValor"), "###,##0.00;(###,##0.00)")

'        LblValor.Caption = Format(CCur(MebVlrPed) - Round(Total("SomadeValor")), "#,##0.00;(#,##0.00)")
    
    End If

End If
End Sub

Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

VarIcms = LerINI("Nota", "CalculaIcms", App.Path & "\config.ini")

Set Ind = Banco.OpenRecordset("Select * From Ind")


If CodDoc <> "" Then
    DatDados(0).RecordSource = "SELECT * From DUP Where Pedido = " & CodDoc & " Order by NroPcl"
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount > 0 Then
        MebParcelas = DatDados(0).Recordset.RecordCount
    End If
    If DatDados(0).Recordset.RecordCount <= 0 Then
        If VarPclAuto Then
            DatDados(0).Recordset.AddNew
            DatDados(0).Recordset("Pedido") = CodDoc
            DatDados(0).Recordset("NroPcl") = 1
            DatDados(0).Recordset("Tipo_Dup") = CmbTipo.Text
            DatDados(0).Recordset("Documento") = TxtDoc.Text
            DatDados(0).Recordset("DtEmiss") = Date
            DatDados(0).Recordset("Entidade") = CodEnt
            DatDados(0).Recordset("DtVenc") = Date + 30
            DatDados(0).Recordset("Valor") = MebVlrPed
            DatDados(0).Recordset.Update
        End If
    Else
        SQL = " UPDATE DUP SET DUP.Entidade = " & CodEnt & " WHERE DUP.Pedido = " & CodDoc
        Banco.Execute SQL
    End If
End If


BtImprimir(1).Enabled = False
BtImprimir(0).Enabled = False
BtImprimir(2).Enabled = False

'CodDoc.Text = FrmPedidos.TxtDados(0)
If MebEntrada = "" Then
    MebEntrada = 0
End If
If MebParcelas = "" Then
    MebParcelas = 1
End If

MebDia = VarDiaFixo
TxtMes = Month(Date)
CmbTipo.ListIndex = 0
DatDados_Reposition 0

If VarIcms Then
    MebVlrPed = MebVlrPed '+ FrmPedidos.Tabela("Ipi")
    MebValorPcl = MebValorPcl '+ FrmPedidos.Tabela("Ipi")
End If

CarregaTipo

If DatDados(0).Recordset.RecordCount > 0 Then
    Dim VARPASS As Recordset
    Set VARPASS = Banco.OpenRecordset("SELECT * FROM PWD WHERE NOME = 'ALTPARCELA'")
    VARPASS.Requery
    If VARPASS.RecordCount = 1 Then
'        OK.Enabled = False
'        GridParc.Enabled = False
'        PassWord = Trim(VarPass!Senha)

    End If
End If
                   
End Sub

Private Sub CmbTipo_GotFocus()
    If Dir(App.Path & "\Tipo.Txt") <> "" Then
        CmbTipo.ListIndex = 1
    Else
        CmbTipo.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

'MebVlrPed.Text = Format(FrmPedidos.TxtDados(2), "#,##0.00")
Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

If LerINI("PEDIDO", "ParcelaAutomatica", App.Path & "\Config.ini") <> "" Then
    VarPclAuto = LerINI("PEDIDO", "ParcelaAutomatica", App.Path & "\Config.ini")

End If

If LerINI("PEDIDO", "PclDiaFixo", App.Path & "\Config.ini") <> "" Then
    VarDiaFixo = LerINI("PEDIDO", "PclDiaFixo", App.Path & "\Config.ini")
    If VarDiaFixo = 0 Then
        VarDiaFixo = ""
    End If
Else
    VarDiaFixo = 10
End If

Dim TipoVenc As String

    If Dir(App.Path & "\Vencimentos.Txt") <> "" Then
    On Error GoTo ErroVenc
    CmbVenc.Clear
    Open App.Path & "\Vencimentos.Txt" For Input As #1
        Do
            Input #1, TipoVenc
            CmbVenc.AddItem TipoVenc
        Loop Until TipoVenc = ""
    Close #1
End If

ErroVenc:
If Err.Number = 62 Then
    Close #1
End If

End Sub

Private Sub BtImprimir_Click(Index As Integer)
Dim Linha02Col1 As String
Dim Linha03Col1 As String
Dim Linha04Col1 As String
Dim Linha05Col1 As String
Dim Linha06Col1 As String
Dim Linha07Col1 As String
Dim Linha08Col1 As String

Dim Linha02Col2 As String
Dim Linha03Col2 As String
Dim Linha04Col2 As String
Dim Linha05Col2 As String
Dim Linha06Col2 As String
Dim Linha07Col2 As String
Dim Linha08Col2 As String

Dim flag As Boolean
Dim Configura As Recordset
Dim LinhaSimples As String
Dim LinhaDupla As String
Dim Corte As String
Dim PortaTemp As String
Dim VarFormCarne As Form

PortaTemp = Porta

If Index = 5 Then
    Set VarFormCarne = FrmRelCapa
Else
    Set VarFormCarne = FrmRelCarnePromi
End If

If Index = 7 Then
    Set VarFormCarne = FrmRelCarnePromi3
End If

If LerINI("Impressora", "ImprNota", App.Path & "\config.ini") <> "" Then
    Porta = LerINI("Impressora", "ImprNota", App.Path & "\config.ini")
End If

If LerINI("Impressora", "ImprCarne", App.Path & "\config.ini") <> "" Then
    Porta = LerINI("Impressora", "ImprCarne", App.Path & "\config.ini")
End If

LinhaSimples = "--------------------------------------"
LinhaDupla = "======================================"
Corte = "......................................"

DatDados(0).Refresh

If BtImprimir(Index).Enabled = False Then
    Exit Sub
End If

flag = False
If MsgBox("Confirma Impressão do Carnê ? ", vbYesNo, App.Title) = vbYes Then
    Set Configura = Banco.OpenRecordset("SELECT * FROM CFG")
    DatDados(0).Recordset.MoveFirst
    Select Case Index
        Case 0
                Iniciar True
                Do While Not DatDados(0).Recordset.Eof
                    
                        Linha02Col1 = DatDados(0).Recordset("DtVenc")
                        Linha03Col1 = DatDados(0).Recordset("Valor")
                        Linha04Col1 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col1 = Date
                        Linha06Col1 = Format(CodEnt, "0000")
                        Linha07Col1 = Configura.Fields("MensagemRodape")
                        Linha08Col1 = Configura.Fields("ObsParcelas")
                                         
                        Linha02Col2 = DatDados(0).Recordset("DtVenc")
                        Linha03Col2 = DatDados(0).Recordset("Valor")
                        Linha04Col2 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col2 = Date
                        Linha06Col2 = Format(CodEnt, "0000")
                        Linha07Col2 = Configura.Fields("MensagemRodape")
                        Linha08Col2 = Configura.Fields("ObsParcelas")
                                
                        ImprimeLinha 1, 0, 0, 0, False, 0
                        ImprimeLinha 0, 30, CmbTipo.Text, 5, False, 0
                        ImprimeLinha 0, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, CmbTipo.Text, 5, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        
                        ImprimeLinha 0, 0, TxtNome.Text, 38, False, 0
                        ImprimeLinha 0, 0, "", 4, False, 0
                        ImprimeLinha 0, 0, TxtNome.Text, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha02Col1, 19, False, 0
                        ImprimeLinha 0, 0, Linha03Col1, 17, False, 2
                        
                        ImprimeLinha 0, 6, Linha02Col2, 19, False, 0
                        ImprimeLinha 0, 0, Linha03Col2, 17, False, 2
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha04Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha04Col2, 38, False, 0
                
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha05Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha05Col2, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha06Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha06Col2, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha07Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha07Col2, 38, False, 0
                        ImprimeLinha 1, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha08Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha08Col2, 38, False, 0
                        
                        ImprimeLinha 4, 0, 0, 0, False, 0
                        
                        DatDados(0).Recordset.MoveNext
                Loop
                Fim True
                Unload Me
        Case 1
                Iniciar True
                Do While Not DatDados(0).Recordset.Eof
                    
                        Linha02Col1 = DatDados(0).Recordset("DtVenc")
                        Linha03Col1 = DatDados(0).Recordset("Valor")
                        Linha04Col1 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col1 = Date
                        Linha06Col1 = Format(CodEnt, "00000")
                        Linha07Col1 = Configura.Fields("MensagemRodape")
                        Linha08Col1 = Configura.Fields("ObsParcelas")
                                         
                                
                        ImprimeLinha 1, 0, 0, 0, False, 0
                        ImprimeLinha 0, 0, LinhaDupla, 38, False, 0
                        ImprimeLinha 1, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha06Col1 & "-" & TxtNome, 32, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Tipo:", 6, False, 0
                        ImprimeLinha 0, 1, CmbTipo.Text, 5, False, 0
                        ImprimeLinha 0, 5, "Pedido:", 7, False, 0
                        ImprimeLinha 0, 1, Linha04Col1, 8, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Data:", 5, False, 0
                        ImprimeLinha 0, 1, Date, 10, False, 0
                        ImprimeLinha 0, 1, "Valor:", 6, False, 0
                        ImprimeLinha 0, 1, Linha03Col1, 10, False, 2
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Venc.", 5, False, 0
                        ImprimeLinha 0, 1, Linha02Col1, 10, False, 0
                        ImprimeLinha 0, 1, "Juros:", 6, False, 0
                        ImprimeLinha 0, 1, "", 2, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Pago:", 5, False, 0
                        ImprimeLinha 0, 1, "", 10, False, 0
                        ImprimeLinha 0, 1, "Total:", 6, False, 0
                        ImprimeLinha 0, 1, "", 10, False, 0
                        
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, LinhaSimples, 38, False, 0
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, Linha07Col1, 38, False, 0
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, Linha08Col1, 38, False, 0
                        ImprimeLinha 2, 0, "", 35, False, 0
                                                
                        DatDados(0).Recordset.MoveNext
                Loop
                ImprimeLinha 2, 0, "", 35, False, 0
                ImprimeLinha 2, 0, "", 35, False, 0
                ImprimeLinha 2, 0, "", 35, False, 0
                Fim True
                Unload Me
        Case 2
            Set varPesquisa(1) = DatDados(0).Recordset.Clone
            varPesquisa(1).Requery
            varPesquisa(1).MoveFirst
            
            FrmRelCarne.rpTitulo.Caption = Configura!Cabecalho1
            FrmRelCarne.RpTitulo2.Caption = Configura!Cabecalho1
            FrmRelCarne.rpNome.Caption = Left(TxtNome, 40)
            FrmRelCarne.rpNome2.Caption = Left(TxtNome, 40)
            FrmRelCarne.RpMens.Caption = Configura!obsparcelas
            FrmRelCarne.RpMens2.Caption = Configura!obsparcelas
            FrmRelCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            FrmRelCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                        
            Set FrmRelCarne.Relatorio.Recordset = varPesquisa(1)
            Me.Hide
            FrmRelCarne.Config
            Me.Show vbModal
                                                             
        Case 3 To 5
            Dim varValor As Currency
            varValor = LblValor
'            If IsNumeric(MebEntrada) Then
'                varValor = LblValor - MebEntrada
'            End If
            Set varPesquisa(1) = DatDados(0).Recordset.Clone
            varPesquisa(1).Requery
            varPesquisa(1).MoveFirst
            Set VarEnt = Banco.OpenRecordset("Select CodEntidade, Nome, Endereco, cpf, Nro, Cidade, Uf FROM Ent")
            Criterio = "CodEntidade=" & CodEnt
            VarEnt.FindFirst Criterio
            
            DatDados(0).Recordset.MoveLast
            VarFormCarne.NroPedido.Caption = Format(CodDoc, "000000")
            VarFormCarne.NroPedido1.Caption = Format(CodDoc, "000000")
            If VarPrecoPedido = True Then
                VarFormCarne.QuantiaExtenso.Caption = UCase(PassaExtenso(CStr(varValor))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                VarFormCarne.ValorPromissoria.Caption = Format(varValor, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
            Else
                VarFormCarne.QuantiaExtenso.Caption = UCase(PassaExtenso(CStr(varValor - MebEntrada))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                VarFormCarne.ValorPromissoria.Caption = Format(varValor - MebEntrada, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
            End If
            If VarPromissoriaBca = False Then
                VarFormCarne.ExtensoDia.Caption = valorextenso(Day(DatDados(0).Recordset("DtVenc"))) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(DatDados(0).Recordset("DtVenc")))) & " DO ANO DE " & valorextenso(Year(DatDados(0).Recordset("DtVenc")))
                VarFormCarne.DiaVenc.Caption = Day(DatDados(0).Recordset("DtVenc"))
                VarFormCarne.MesVenc.Caption = UCase(RetornaMes(Month(DatDados(0).Recordset("DtVenc"))))
                VarFormCarne.AnoVenc.Caption = Year(DatDados(0).Recordset("DtVenc"))
                VarFormCarne.Pagavel.Caption = CidadeEmpresa
                VarFormCarne.NomeEmpresa.Caption = NomeInscr
                VarFormCarne.EmpresaCGCCPF.Caption = Cnpj
                VarFormCarne.RpDia.Caption = Day(Date)
                VarFormCarne.RpMes.Caption = UCase(RetornaMes(Month(Date)))
                VarFormCarne.RpAno.Caption = Year(Date)
            
            End If
            DatDados(0).Recordset.MoveFirst
            VarFormCarne.Emitente.Caption = TxtNome ' Left(VarEnt!Nome, 40)
            VarFormCarne.Emitente2.Caption = TxtNome  'Left(VarEnt!Nome, 40)
            If Not IsNull(VarEnt("Cpf")) Then
                  VarFormCarne.EmitenteCpf.Caption = VarEnt("CPF")
            End If
            VarFormCarne.EmitenteEndereco.Caption = VarEnt("Endereco") & " - " & VarEnt("Nro")
            VarFormCarne.EmitenteCidade.Caption = VarEnt("Cidade") & " - " & VarEnt("Uf")
            VarFormCarne.rpTitulo.Caption = Configura!Cabecalho1
            VarFormCarne.RpTitulo2.Caption = Configura!Cabecalho1
            VarFormCarne.rpNome.Caption = TxtNome ' Left(VarEnt!Nome, 40)
            VarFormCarne.rpNome2.Caption = TxtNome 'Left(VarEnt!Nome, 40)
            VarFormCarne.RpMens.Caption = Configura!obsparcelas
            VarFormCarne.RpMens2.Caption = Configura!obsparcelas
            VarFormCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            VarFormCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            If Index = 4 Then
                VarFormCarne.Det.Mostrar = False
            End If
                        
            Set VarFormCarne.Relatorio.Recordset = varPesquisa(1)
            Me.Hide
            If Index = 5 Then
                FrmRelCapa.Config
            Else
                FrmRelCarnePromi.Config
            End If
            Me.Show vbModal
        Case 6
'            Dim varValor As Currency
            varValor = LblValor
            If IsNumeric(MebEntrada) Then
                varValor = (LblValor - MebEntrada) / MebParcelas
            Else
                varValor = LblValor / MebParcelas
            End If
            Set varPesquisa(1) = DatDados(0).Recordset.Clone
            varPesquisa(1).Requery
            varPesquisa(1).MoveFirst
            Set VarEnt = Banco.OpenRecordset("Select CodEntidade, Nome, Endereco, cpf, Nro, Cidade, uf FROM Ent")
            Criterio = "CodEntidade=" & CodEnt
            VarEnt.FindFirst Criterio
            
'            FrmRelCarnePromi2.NroPedido.Caption = Format(CodDoc, "000000")
            FrmRelCarnePromi2.ExtensoDia.Caption = valorextenso(Day(DatDados(0).Recordset("DtVenc"))) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(DatDados(0).Recordset("DtVenc")))) & " DO ANO DE " & valorextenso(Year(DatDados(0).Recordset("DtVenc")))
   '         FrmRelCarnePromi2.valorextenso.Caption = UCase(PassaExtenso(CStr(varValor))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
'            FrmRelCarnePromi2.ValorPromissoria.Caption = Format(varValor, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
   '         FrmRelCarnePromi2.DiaVenc.Caption = Day(Datdados(0).Recordset("DtVenc"))
   '         FrmRelCarnePromi2.MesVenc.Caption = UCase(RetornaMes(Month(Datdados(0).Recordset("DtVenc"))))
   '         FrmRelCarnePromi2.AnoVenc.Caption = Year(Datdados(0).Recordset("DtVenc"))
            FrmRelCarnePromi2.Emitente.Caption = Left(VarEnt!Nome, 40)
            FrmRelCarnePromi2.Emitente2.Caption = psCompString(2, VarEnt!CodEntidade, 6) & "-" & Left(VarEnt!Nome, 40)
            If Not IsNull(VarEnt("Cpf")) Then
                FrmRelCarnePromi2.EmitenteCpf.Caption = VarEnt("CPF")
            End If
            FrmRelCarnePromi2.Pagavel.Caption = CidadeEmpresa
'            FrmRelCarnePromi2.Pagavel.Caption = CidadeEmpresa
            FrmRelCarnePromi2.EmitenteEndereco.Caption = VarEnt("Endereco") & " - " & VarEnt("Nro")
            FrmRelCarnePromi2.EmitenteCidade.Caption = VarEnt("Cidade") & " - " & VarEnt("Uf")
            FrmRelCarnePromi2.NomeEmpresa.Caption = NomeInscr
            FrmRelCarnePromi2.EmpresaCGCCPF.Caption = Cnpj
            FrmRelCarnePromi2.RpDia.Caption = Day(Date)
            FrmRelCarnePromi2.RpMes.Caption = UCase(RetornaMes(Month(Date)))
            FrmRelCarnePromi2.RpAno.Caption = Year(Date)
                    
            FrmRelCarnePromi2.rpTitulo.Caption = Configura!Cabecalho1
'            FrmRelCarnePromi2.RpTitulo2.Caption = Configura!Cabecalho1
            FrmRelCarnePromi2.rpNome.Caption = Left(TxtNome, 40)
'            FrmRelCarnePromi2.RpNome2.Caption = Left(TxtNome, 40)
            FrmRelCarnePromi2.RpMens.Caption = Configura!MensagemRodape
'            FrmRelCarnePromi2.RpMens2.Caption = Configura!MensagemRodape
            FrmRelCarnePromi2.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            FrmRelCarnePromi2.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            
            Set FrmRelCarnePromi2.Relatorio.Recordset = varPesquisa(1)
            Me.Hide
            FrmRelCarnePromi2.Config
            Me.Show vbModal
    Case 7
    
'            Dim varValor As Currency
            Dim X As Byte
            X = 0
            
            varValor = LblValor
            If IsNumeric(MebEntrada) Then
                varValor = LblValor - MebEntrada
            End If
            Do While Not DatDados(0).Recordset.Eof
                
               VarFormCarne.R1(X).Caption = DatDados(0).Recordset("NroPcl")
'               VarFormCarne.R2(X).Caption = DatDados(0).Recordset.RecordCount
               VarFormCarne.R3(X).Caption = DatDados(0).Recordset("DtVenc")
               VarFormCarne.R4(X).Caption = DatDados(0).Recordset("Valor")
               VarFormCarne.R2(X).Mostrar = False
               X = X + 1
               DatDados(0).Recordset.MoveNext
            Loop
                                    
            Set varPesquisa(1) = DatDados(0).Recordset.Clone
            varPesquisa(1).Requery
            varPesquisa(1).MoveFirst
            Set VarEnt = Banco.OpenRecordset("Select CodEntidade, Nome, Endereco, cpf, Nro, Cidade, Uf FROM Ent")
            Criterio = "CodEntidade=" & CodEnt
            VarEnt.FindFirst Criterio
            
            DatDados(0).Recordset.MoveLast
            VarFormCarne.NroPedido.Caption = "Pedido:" & CodDoc & " Nº Parcelas: " & varPesquisa(1).RecordCount ' & "/" & varPesquisa(1).RecordCount
            VarFormCarne.QuantiaExtenso.Caption = UCase(PassaExtenso(CStr(varValor))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
            VarFormCarne.ValorPromissoria.Caption = Format(varValor, "#,#0.00") ' Format(Datdados(5).Recordset("Total"), "#,#0.00")
            VarFormCarne.RpNomeCarne.Caption = VarEnt!Nome
            
            
            If VarPromissoriaBca = False Then
                VarFormCarne.ExtensoDia.Caption = valorextenso(Day(DatDados(0).Recordset("DtVenc"))) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(DatDados(0).Recordset("DtVenc")))) & " DO ANO DE " & valorextenso(Year(DatDados(0).Recordset("DtVenc")))
                VarFormCarne.DiaVenc.Caption = Day(DatDados(0).Recordset("DtVenc"))
                VarFormCarne.MesVenc.Caption = UCase(RetornaMes(Month(DatDados(0).Recordset("DtVenc"))))
                VarFormCarne.AnoVenc.Caption = Year(DatDados(0).Recordset("DtVenc"))
                VarFormCarne.Pagavel.Caption = CidadeEmpresa
                VarFormCarne.NomeEmpresa.Caption = NomeInscr
                VarFormCarne.EmpresaCGCCPF.Caption = Cnpj
                VarFormCarne.RpDia.Caption = Day(Date)
                VarFormCarne.RpMes.Caption = UCase(RetornaMes(Month(Date)))
                VarFormCarne.RpAno.Caption = Year(Date)
            
            End If
            
            
            DatDados(0).Recordset.MoveFirst
            VarFormCarne.Emitente.Caption = TxtNome ' Left(VarEnt!Nome, 40)
            VarFormCarne.Emitente2.Caption = TxtNome  'Left(VarEnt!Nome, 40)
            If Not IsNull(VarEnt("Cpf")) Then
                  VarFormCarne.EmitenteCpf.Caption = VarEnt("CPF")
            End If
            VarFormCarne.EmitenteEndereco.Caption = VarEnt("Endereco") & " - " & VarEnt("Nro")
            VarFormCarne.EmitenteCidade.Caption = VarEnt("Cidade") & " - " & VarEnt("Uf")
            VarFormCarne.rpTitulo.Caption = Configura!Cabecalho1
            VarFormCarne.RpTitulo2.Caption = Configura!Cabecalho1
            VarFormCarne.rpNome.Caption = TxtNome ' Left(VarEnt!Nome, 40)
            VarFormCarne.rpNome2.Caption = TxtNome 'Left(VarEnt!Nome, 40)
            VarFormCarne.RpMens.Caption = Configura!MensagemRodape
            VarFormCarne.RpMens2.Caption = Configura!MensagemRodape
            VarFormCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            VarFormCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            
            VarFormCarne.Det.Mostrar = False

                        
            Set VarFormCarne.Relatorio.Recordset = varPesquisa(1)
            Me.Hide
            FrmRelCarnePromi3.Config
            Me.Show vbModal
    
                                                                                                                          
    End Select

End If
Porta = PortaTemp
End Sub


Private Sub GridParc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 9 Then
    If DatDados(0).Recordset.RecordCount > 0 Then
        If DatDados(0).Recordset.Bof Then
            DatDados(0).Recordset.MoveNext
        Else
            DatDados(0).Recordset.MoveLast
        End If
    End If
End If
End Sub

Private Sub MebEntrada_LostFocus()
If MebEntrada.Text = "" Then
    MebEntrada = 0
End If
End Sub

Private Sub MebIndices_Change()
If MebIndices = "" Or IsNull(MebIndices) Then
    MebIndices = 0
End If
If MebParcelas <> "" And IsNumeric(MebEntrada) And IsNumeric(MebIndices) Then
    If MebEntrada.Text <> 0 And IsNumeric(MebEntrada) Then
        If MebIndices = 0 Then
            MebValorPcl = (MebVlrPed - MebEntrada) / MebParcelas
        Else
            MebValorPcl = (MebVlrPed - MebEntrada) * MebIndices ' Round(MebIndices, 4)
        End If
    Else
        If MebIndices = 0 Then
            MebValorPcl = MebVlrPed / MebParcelas
        Else
            MebValorPcl = MebVlrPed * MebIndices ' Round(MebIndices, 4)
        End If
    End If
End If
End Sub

Private Sub MebParcelas_Change()
Dim Criterio As String
If MebParcelas <> "" Then
    If Ind.RecordCount <> 0 Then
        Criterio = "NroIndice = " & MebParcelas
        Ind.FindFirst Criterio
        If Not Ind.NoMatch Then
            MebIndices = Ind("Indice")
        End If
    Else
        MebIndices = 0
    End If
End If
End Sub

Private Sub MebParcelas_LostFocus()
MebParcelas_Change
End Sub

Private Sub MebValorPcl_LostFocus()
If MebValorPcl.Text = "" Then
    If MebParcelas.Text <> "" Then
        MebValorPcl.Text = (CCur(MebVlrPed) - CCur(MebEntrada)) / MebParcelas
    Else
        MebParcelas = ""
        MebParcelas.SetFocus
    End If
End If
End Sub

Private Sub OK_Click()
Dim SQL As String
Dim Parc As Integer
Dim DtVenc As Date
Dim X As Integer
Dim Venc(50) As Integer
Dim Cont As Integer
Dim VarTemp As String
Cont = 1

If MebDia <> "" And Not IsNumeric(MebDia) Then
    If MebDia < 0 And MebDia > 28 Then
        MsgBox "Dia Inválido !", vbExclamation, App.Title
        MebDia.SetFocus
        Exit Sub
    End If
    MebDia.SetFocus
    Exit Sub
End If

If TxtMes <> "" And Not IsNumeric(TxtMes) Then
    If TxtMes < 0 And TxtMes > 13 Then
        MsgBox "Mes Inválido !", vbExclamation, App.Title
        TxtMes.SetFocus
        Exit Sub
    End If
    TxtMes.SetFocus
    Exit Sub
End If

If MebValorPcl = "" And Not IsNumeric(MebValorPcl) Then
    MebValorPcl.SetFocus
    Exit Sub
End If

If CmbVenc <> "" Then
    VarTemp = ""
    For X = 1 To Len(CmbVenc)
        If Mid$(CmbVenc, X, 1) <> "/" Then
            VarTemp = VarTemp & Mid$(CmbVenc, X, 1)
        Else
            Venc(Cont) = CInt(VarTemp)
            Cont = Cont + 1
            VarTemp = ""
        End If
    Next
    Venc(Cont) = CInt(VarTemp)
End If

SQL = "Delete * From Dup Where Pedido = " & CodDoc
Banco.Execute SQL
DatDados(0).Refresh
GridParc.Refresh
Set Parcelas = Banco.OpenRecordset("DUP")

If MebDia = "" Or TxtMes = "" Or MebDia = "0" Or TxtMes = "0" Then
    DtVenc = Date
Else
    If TxtMes = "1" And Month(Date) = 12 Then
        DtVenc = Format(MebDia & "/" & CInt(TxtMes) & "/" & Year(Date) + 1, "dd/mm/yyyy")
    Else
        
        DtVenc = Format(MebDia & "/" & CInt(TxtMes) & "/" & Year(Date), "dd/mm/yyyy")
    End If
End If

If MebEntrada.Text <> "0" And IsNumeric(MebEntrada) Then
    Parc = 1
    Parcelas.AddNew
    Parcelas("Pedido") = CodDoc
    Parcelas("NroPcl") = Parc
    Parcelas("Tipo_Dup") = CmbTipo.Text
    Parcelas("Documento") = TxtDoc.Text
    Parcelas("DtEmiss") = Date
    Parcelas("Entidade") = CodEnt
    Parcelas("DtVenc") = Date
    Parcelas("Valor") = CCur(MebEntrada)
    Parcelas("Documento") = "ENTRADA"
    Parcelas("Cond") = True
    Parcelas.Update
    
    If CodDoc <> 0 Then
        SQL = "Delete * From Rec Where CodDupli = " & CodDoc
        Banco.Execute SQL
    End If
    
    SQL = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Maquina  ) VALUES ( "
    SQL = SQL & CodEnt & " , "
    SQL = SQL & "#" & Format(Date, "mm/dd/yyyy") & "# , "
    SQL = SQL & """" & Time & """" & " , "
    SQL = SQL & Numero(CStr(MebEntrada), 2) & " , "
    SQL = SQL & 0 & " , "
    SQL = SQL & 1 & " ,"
    SQL = SQL & CodDoc & " ,"
    SQL = SQL & """" & LIMPACampo(Trim(Maquina)) & """" & ")"
    Banco.Execute SQL
      
Else
    Parc = 0
End If
For X = 1 To MebParcelas
    Parc = Parc + 1
    If CmbVenc <> "" Then
        DtVenc = Date + Venc(X)
    Else
        DtVenc = DtVenc + 30
    End If
    Parcelas.AddNew
    Parcelas("Pedido") = CodDoc
    Parcelas("NroPcl") = Parc
    Parcelas("Tipo_Dup") = CmbTipo.Text
    Parcelas("Documento") = TxtDoc.Text
    Parcelas("DtEmiss") = Date
    If IsNumeric(CodEnt) Then
        Parcelas("Entidade") = CodEnt
    End If
    If MebDia = "" Then
        Parcelas("DtVenc") = DtVenc
    Else
        If IsDate(Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy")) Then
            Parcelas("DtVenc") = Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy")
        Else
            Parcelas("DtVenc") = Null
        End If
    End If
    If ChkPgto.value = 1 Then
        Parcelas("Cond") = True
        Parcelas("DtPgto") = Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy")
        Parcelas("VlrPago") = Round(CCur(MebValorPcl), 2)
        
        SQL = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Juros, Maquina  ) VALUES ( "
        SQL = SQL & CodEnt & " , "
'        SQL = SQL & "#" & Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy") & "# , "
        SQL = SQL & "#" & Format(Parcelas("DtVenc"), "mm/dd/yyyy") & "# , "
        SQL = SQL & """" & Time & """" & " , "
        SQL = SQL & Numero2(CStr(MebValorPcl)) & " , "
        SQL = SQL & Numero2(CStr(0)) & " , "
        SQL = SQL & 1 & " ,"
        SQL = SQL & """" & TxtDoc & """" & " ,"
        SQL = SQL & Numero2(CStr(0)) & " ,"
        SQL = SQL & """" & LIMPACampo(Trim(Maquina)) & """" & ")"
        Banco.Execute SQL
        
        
    End If
    
    Parcelas("Valor") = Round(CCur(MebValorPcl), 2)
    Parcelas.Update
Next

DatDados(0).RecordSource = "SELECT DUP.* FROM DUP WHERE (((DUP.Pedido)= " & CodDoc & ")) ORDER BY DUP.NroPcl"
DatDados(0).Refresh


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyPress = 0
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

