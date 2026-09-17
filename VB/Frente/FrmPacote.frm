VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form FrmPacote 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8565
   ClientLeft      =   4095
   ClientTop       =   3075
   ClientWidth     =   12810
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtConsulta 
      Appearance      =   0  'Flat
      Caption         =   "Consulta"
      DisabledPicture =   "FrmPacote.frx":0000
      Height          =   435
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Pesquisar"
      Top             =   1035
      Width           =   1050
   End
   Begin VB.Frame FrameEnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cadastro"
      Height          =   2625
      Left            =   2745
      TabIndex        =   12
      Top             =   2655
      Visible         =   0   'False
      Width           =   7845
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         Height          =   285
         Left            =   7515
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   285
      End
      Begin VB.CommandButton CmdSalvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salvar"
         Height          =   375
         Left            =   6075
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2025
         Width           =   1320
      End
      Begin VB.TextBox Email 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Email"
         Height          =   315
         Left            =   945
         MaxLength       =   60
         TabIndex        =   17
         Tag             =   "S"
         Top             =   1440
         Width           =   6435
      End
      Begin VB.TextBox Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Nome"
         Height          =   315
         Left            =   945
         MaxLength       =   60
         TabIndex        =   13
         Tag             =   "S"
         Top             =   540
         Width           =   6435
      End
      Begin MSMask.MaskEdBox Fone 
         DataField       =   "Celular"
         Height          =   315
         Left            =   5625
         TabIndex        =   16
         Top             =   990
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         Mask            =   "(##)#.####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cpf 
         DataField       =   "CPF"
         DataSource      =   "DatDados(1)"
         Height          =   315
         Left            =   945
         TabIndex        =   14
         Tag             =   "S"
         Top             =   990
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Rg 
         DataField       =   "RG"
         DataSource      =   "DatDados(1)"
         Height          =   315
         Left            =   3330
         TabIndex        =   15
         Tag             =   "S"
         Top             =   990
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "##.###.###-C"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG:"
         Height          =   195
         Index           =   29
         Left            =   2880
         TabIndex        =   23
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Index           =   55
         Left            =   315
         TabIndex        =   22
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Index           =   37
         Left            =   435
         TabIndex        =   21
         Top             =   990
         Width           =   345
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         Height          =   195
         Index           =   8
         Left            =   5175
         TabIndex        =   20
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label LblDados 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   19
         Top             =   540
         Width           =   465
      End
   End
   Begin VB.CommandButton BtPesquisa 
      Appearance      =   0  'Flat
      Caption         =   "Cadastra"
      DisabledPicture =   "FrmPacote.frx":0532
      Height          =   435
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Pesquisar"
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Data DatPesquisaG 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7245
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   585
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton CmbFecha 
      BackColor       =   &H8000000A&
      Caption         =   "&Baixar o Produto/Serviço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7605
      Width           =   3150
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -315
      ScaleHeight     =   330
      ScaleWidth      =   17445
      TabIndex        =   8
      Top             =   8235
      Width           =   17475
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3825
      TabIndex        =   7
      Top             =   7740
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   630
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   17145
      TabIndex        =   5
      Top             =   0
      Width           =   17145
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "LANÇA COMPRA ANTECIPADA"
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
         Visible         =   0   'False
         Width           =   5685
      End
   End
   Begin VB.TextBox TxtCartao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6210
      TabIndex        =   1
      Top             =   1035
      Width           =   1155
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\AATESTE\AATesteBica\Posto\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   585
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSComctlLib.ListView ListVda 
      Height          =   5685
      Left            =   180
      TabIndex        =   3
      Top             =   1530
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   10028
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CodPrd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Complemento"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Produto"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qtd"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Unitário"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSDBCtls.DBCombo TxtPessoa 
      Bindings        =   "FrmPacote.frx":0A64
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Top             =   1035
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11325
      TabIndex        =   25
      Top             =   7290
      Width           =   570
   End
   Begin VB.Label MebTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   9450
      TabIndex        =   24
      Top             =   7605
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   765
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6210
      TabIndex        =   4
      Top             =   765
      Width           =   675
   End
   Begin VB.Label Mens 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   8175
      Width           =   270
   End
End
Attribute VB_Name = "FrmPacote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Banco As Database
Public Formulario As Form
Public Registros As DAO.Recordset
Public Clientes As DAO.Recordset
Public Vendas As DAO.Recordset
Public Bicos As DAO.Recordset
Public Bico As DAO.Recordset
Public Encerrante As DAO.Recordset
'Public Ofertas As Dao.RecordSet
Public TotTemp As DAO.Recordset
Public Linha As Integer
Public ContPro As Integer
Public Soma As Double
'Public Caminho As String
Public Flag2 As Boolean
Public Pessoa As String
Public NroCupom As String
Public Itens As Integer
Public Ponto As Integer
Public Config As DAO.Recordset
Public PistaAnt As String
Public PosicaoAba As Long
Public LstPesq As ListItem
Public PrdAnt As Long
Public Varvda As Recordset
Public VarProdutos As Recordset
Public VarFlagVer As Boolean
Public VarX As Integer

Public CodCard As Double

'Public Automa As Dao.RecordSet

Private Sub CmbCalc_Click()
    Shell ("calc")
End Sub

Private Sub CmbCancela_Click()

Dim sql As String

If MsgBox("Confirma cancelamento da venda ?", vbYesNo) = vbYes Then
    ListVda.ListItems.Clear
        sql = "UPDATE Conv SET Cancelado = -1, Totaliza = -1   WHERE Cliente = " & CodCard & " AND Flag = 0 " ' AND Maquina = " & """" & LimpaCampo(Left(Maquina, 15)) & """"
        FrmFrente.Banco.Execute sql
'        DatDados(1).Refresh
        LimpaTela
Else
'    TxtDados_KeyPress 2, 13
End If
End Sub



Private Sub BtConsulta_Click()
Dim VarConv As Recordset
'Dim soma As Currency
Soma = 0
VarFlagVer = True
VarX = 0
If TxtCartao <> "" And MebTotal = "" Then
    If IsNumeric(TxtCartao) Then
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Sum(CONV.Valor) AS SomaDeValor, CONV.Cliente, CONV.Flag, CONV.Totaliza From CONV GROUP BY CONV.Cliente, CONV.Flag, CONV.Totaliza HAVING cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 ")
        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Conv.Maquina,Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
        
        VarConv.Requery
        If VarConv.RecordCount > 0 Then
'                MebTotal = VarConv!somadevalor
            VarConv.MoveFirst
'            TxtDados(0).Enabled = False
'            TxtDados(2).Enabled = False
            ListVda.ListItems.Clear
                        
            While Not VarConv.EOF
                
                If VarConv.Fields("Valor") <> 0 Then
                    Set LstPesq = ListVda.ListItems.Add(, , VarConv!CodVda)

                
                    LstPesq.SubItems(1) = VarConv.Fields("Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(2) = VarConv.Fields("Maquina")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(3) = VarConv.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
                    LstPesq.SubItems(4) = Numero(Format(VarConv.Fields("Qtd"), " ##,##0.00"), 2)  'SeCampoNulo(Consulta, "Conta")
                    LstPesq.SubItems(5) = Numero(Format(VarConv.Fields("Valor") / VarConv.Fields("qtd"), "##,##0.00"), 2)  'SeCampoNulo(Consulta, "DataOper")
                    LstPesq.SubItems(6) = Format(VarConv.Fields("Valor"), "##,##0.00")  'SeCampoNulo(Consulta, "BomPara")
'                    LstPesq.SubItems(6) = VarConv.Fields("Cupom")   'SeCampoNulo(Consulta, "Agencia")
                    Soma = Soma + VarConv!Valor
                    MebTotal.Caption = Format(Soma, "     ##,##0.00")
                    VarX = VarX + 1
                End If
                VarConv.MoveNext
            Wend
            VarX = VarX + 1
        End If
    End If
End If
End Sub

Private Sub BtPesquisa_Click()
FrameEnt.Visible = True
Nome.SetFocus
End Sub

Private Sub CmbFecha_Click()
If MsgBox("Confirma Serviço ? ", vbYesNo, App.Path) = vbYes Then

    For x = 1 To ListVda.ListItems.Count
        If ListVda.ListItems(x).Checked Then
                        
            sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom,Flag, Qtd, Produto,  Valor ) Values ("
            sql = sql & "99" & psCompString(1, TxtCartao, 6) & " ,"
            sql = sql & """" & LimpaCampo(Left(Pessoa, 15)) & """" & " , #"
            sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
            sql = sql & """" & CStr(Time) & """" & ", "
            sql = sql & 0 & " , "
            sql = sql & -1 & ", "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(4) / 100, " ##,##0.00"), 2) & " , "
            sql = sql & ListVda.ListItems.Item(x).SubItems(1) & " , "
            sql = sql & Numero(Format(ListVda.ListItems.Item(x).SubItems(6), "##,##0.00"), 2) & ")"
            FrmFrente.Banco.Execute sql
            
            Criterio = "CodVda = " & ListVda.ListItems.Item(x)
            Varvda.FindFirst Criterio
            If Not Varvda.NoMatch Then
                Varvda.Edit
                Varvda!cancelado = -1
                Varvda.Update
            End If
            
        End If
    Next
    

    Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Conv.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & "99" & psCompString(1, TxtCartao, 6) & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
    
    VarConv.Requery
    If VarConv.RecordCount > 0 Then
        If VarConv!Cupom <> 0 Then
            FrmFrente.CliCartao = VarConv!Cupom
        End If
    End If
'    TxtDados(0).SetFocus
    FrmFrente.TxtDados(0) = "/" & Trim("99" & psCompString(1, TxtCartao, 6))
    FrmFrente.TxtTmp2 = Trim("99" & psCompString(1, TxtCartao, 6))
'    Frame5.Visible = False
    SendKeys ("{ENTER}")
    Unload Me
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub CmdLimpa_Click()
'LimpaTela
End Sub

Private Sub CmdSalvar_Click()
If Nome <> "" Then
    DatPesquisaG.Recordset.AddNew
    DatPesquisaG.Recordset.Fields("Nome") = Nome
    DatPesquisaG.Recordset.Fields("Cpf") = Cpf
    DatPesquisaG.Recordset.Fields("Rg") = Rg
    DatPesquisaG.Recordset.Fields("Fone") = Fone
    DatPesquisaG.Recordset.Fields("email") = Email
    DatPesquisaG.Recordset.Fields("Tipo") = 1
    DatPesquisaG.Recordset.Fields("Classificacao") = 1
    DatPesquisaG.Recordset.Fields("Ativo") = 1
    DatPesquisaG.Recordset.Update
    MsgBox "Cliente Gravado com Sucesso !", vbInformation
    TxtPessoa.Text = Nome
    FrameEnt.Visible = False
    FrmFrente.Clientes.Requery
    FrmFrente.Registros.Requery
    FrmFrente.DatCli.Recordset.Requery
    
    TxtPessoa.SetFocus
Else
    FrameEnt.Visible = False
    TxtPessoa.SetFocus
End If

End Sub

Private Sub DbcProd_GotFocus()
    DbcProd.BackColor = &HC0FFFF
    DatDados(0).Recordset.MoveFirst
End Sub



Private Sub DbcProd_LostFocus()
DbcProd.BackColor = &H80000005
End Sub

Private Sub DblPesquisaG_DblClick()
Dim sql As String
'FramePesquisaG.Visible = False
sql = "Nome = " & """" & DblPesquisaG.BoundText & """"
DatPesquisaG.Recordset.FindFirst sql
TxtPessoa = DatPesquisaG.Recordset.Fields("Nome")
txtcomplemento = DatPesquisaG.Recordset.Fields("CodEntidade")
'MebDados(1) = DatPesquisaG.Recordset.Fields("CEP")
'MebDados(16) = DatPesquisaG.Recordset.Fields("CodCidade")
'TxtDados(0).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmFrente.Timer2.Enabled = True
End Sub

Private Sub ListVda_DblClick()
Dim Criterio As String
Dim VarItem As String
If ListVda.ListItems.Count > 0 Then
    If MsgBox("Confirma exlusão deste item ? " & ListVda.SelectedItem.Text, vbYesNo, App.Title) = vbYes Then
        VarItem = ListVda.SelectedItem.Index
        Criterio = "CodVda = " & ListVda.SelectedItem.Text
        Varvda.FindFirst Criterio
        If Not Varvda.NoMatch Then
            Soma = MebTotal
            Soma = Soma - Varvda("Valor")
            MebTotal = Format(Soma, "     ##,##0.00")
            ListVda.ListItems.Remove CDec(VarItem)
            
            Varvda.Edit
            Varvda!cancelado = -1
            VarX = VarX - 1
'            VarVda.Delete
            Varvda.Update
        End If
    End If
End If
'If TxtDados(0).Enabled = True Then
'    TxtDados(0).SetFocus
'End If
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim PesquisaSql As String
Dim Retorno As Integer
Dim Status As String
Dim x As Integer
Dim P As Long
Dim Pista As String
Dim Ponto1 As String
Dim Ponto2 As String
Dim Ponto3 As String
Dim Ponto4 As String
Dim Comando As String
Dim VarCodigo As Double
PistaAnt = ""
Caminho = App.Path
VarFlagVer = False

'If App.PrevInstance Then
'    MsgBox "Existe uma aplicação do " & App.ProductName & " aberta nesta máquina, um ou mais programas abertos pode causar lentidão no sistema!", 16
'    End
'End If

'FrmFrente.Timer2.Enabled = False
'If Dir(App.Path & "\Config.ini") <> "" Then
    
'    Caminho = LerINI("Config", "LocalBco", App.Path & "\config.ini")
'    Impressora = LerINI("Config", "Ecf", App.Path & "\config.ini")
'    PortaF = LerINI("Config", "Porta", App.Path & "\config.ini")
'    PortaNF = LerINI("Config", "Impr", App.Path & "\config.ini")
'    Leitor = LerINI("Config", "Leitor", App.Path & "\config.ini")
'    VarTef = LerINI("Config", "Tef", App.Path & "\config.ini")
'    VarTipoVda = LerINI("Config", "TipoVda", App.Path & "\config.ini")
'    VarTermica = LerINI("Config", "Termica", App.Path & "\config.ini")
'    VarComprovante = LerINI("Config", "Comprovante", App.Path & "\config.ini")
'    VarPrecoLiberado = LerINI("Config", "PrecoLiberado", App.Path & "\config.ini")
'    VarTipoPreco = LerINI("Config", "TipoPreco", App.Path & "\config.ini")
'
'End If

If dir(App.Path & "Cond.Txt") <> "" Then
    Nor = Chr(18)
    Cond = Chr(15)
Else
    Nor = ""
    Cond = ""
End If
Gde = Chr(14)
Negi = Chr(27) + Chr(69)
NegF = Chr(27) + Chr(70)
SaltoPag = Chr(12)

'FrmFrente.Caminho = App.Path
Soma = 0
For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

'P = NameOfPC(Maquina)

'Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Dados.mdb")
'Set Config = Banco.OpenRecordset("Select Cfg.* FROM Cfg")

'Config.MoveFirst
'Porta = Config.Fields("Porta")

DatPesquisaG.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"


'If FramePesquisaG.Visible = False Then
'    If TxtPessoa.Text <> "" Then
'       PesquisaGeral FrmFrente.Caminho & "\dados.mdb", "SELECT CodEntidade, Nome From Ent  where Nome like " & """" & "*" & TxtPessoa.Text & "*" & """" & "  ORDER BY Nome", "Nome"
'    Else
       PesquisaGeral FrmFrente.Caminho & "\dados.mdb", "SELECT CodEntidade, Nome, Cpf, Rg, Ativo, Tipo, Classificacao, Fone , Email From Ent  ORDER BY Nome", "Nome"
'    End If
'Else
'    FramePesquisaG.Visible = False
'End If


Set Registros = FrmFrente.Banco.OpenRecordset("SELECT PRD.* , BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto  WHERE Ativo = true ORDER BY PRD.CodPrd ")

Set Vendas = FrmFrente.Banco.OpenRecordset("SELECT * FROM CONV")

Set Varvda = FrmFrente.Banco.OpenRecordset("Select * From Conv Order By CodVda")

Set VarProdutos = FrmFrente.Banco.OpenRecordset("Select CodPrd, Concorrencia, Local From Prd Order By CodPrd")

'DatDados(1).RecordSource = "SELECT vda.pedido, vda.bico, Vda.Unitario, Prd.Estoque, Vda.BaixarEst, vda.cliente, VDA.CodAba, VDA.CodVda, Prd.CodPrd, PRD.Descrição_Produto, VDA.Qtd, PRD.Preco, VDA.Valor, Vda.Maquina FROM PRD INNER JOIN VDA ON PRD.CodPrd = VDA.Produto WHERE VDA.Flag = False and Vda.Maquina = '" & LimpaCampo(Maquina) & "' ORDER BY VDA.CodVda"
'DatDados(1).Refresh

LblHora = Time
LblData = Date

DatDados(0).RecordSource = "SELECT CodPrd, Descrição_Produto FROM Prd  WHERE Ativo = true ORDER BY Descrição_produto"
DatDados(0).Refresh

'sql = "UPDATE Conv SET Conv.Cancelado = -1   WHERE Flag = 0  " 'AND Maquina = " & """" & Trim(LimpaCampo(Left(FrmFrente.Maquina, 15))) & """"
'FrmFrente.Banco.Execute sql


End Sub

Function LimpaTela()
'TxtDados(0) = ""

txtcomplemento = ""
MebTotal = ""
TxtCartao = ""
Soma = 0
ContPro = 0
Valor = 0
txtpgto = ""
'LblCard = ""
'CmbImprimi.Enabled = True
'CmbLista.Enabled = True
CmbFecha.Enabled = True


'TxtDados(2).SetFocus
'TxtDados(0).SetFocus

'TxtDados(0).Enabled = True
'TxtDados(0).SetFocus
'(2).Enabled = True
'TxtDados(2).SetFocus
'TxtCartao.Enabled = True
'TxtCartao.SetFocus

Linha = 0
Itens = 0
ListVda.ListItems.Clear

End Function


Private Sub Timer1_Timer()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then
    LimpaTela
End If

If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Transfere_Click()


End Sub

Private Sub TxtCartao_GotFocus()
TxtCartao.BackColor = &HC0FFFF
End Sub

Private Sub TxtCartao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub TxtCartao_KeyPress(KeyAscii As Integer)
'Dim VarConv As Recordset
'Dim Soma As Currency
'Soma = 0
If KeyAscii = 13 And TxtCartao <> "" Then
    If IsNumeric(TxtCartao) Then
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT Sum(CONV.Valor) AS SomaDeValor, CONV.Cliente, CONV.Flag, CONV.Totaliza From CONV GROUP BY CONV.Cliente, CONV.Flag, CONV.Totaliza HAVING cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 ")
'        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, PRD.Descrição_Produto FROM CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ")
        Set VarConv = FrmFrente.Banco.OpenRecordset("SELECT CONV.Cupom, CONV.Produto, CONV.CodVda,CONV.Valor, CONV.Qtd, CONV.Cliente, CONV.Flag, CONV.Totaliza, Conv.Maquina FROM CONV WHERE cliente = " & TxtCartao & " AND Totaliza = 0 AND Flag = -1 and Cancelado = 0 ORDER BY CODVDA ")
        
        VarConv.Requery
        If VarConv.RecordCount > 0 Then
            If IsNull(VarConv!Maquina) Then
                TxtPessoa = ""
            Else
                TxtPessoa = VarConv!Maquina
            End If
            If VarConv!Cupom <> 0 Then
                txtcomplemento = VarConv!Cupom
            End If
            
        End If
'                MebTotal = VarConv!somadevalor
'            VarConv.MoveFirst
'            While Not VarConv.EOF
            
'                Set LstPesq = ListVda.ListItems.Add(, , VarConv!CodVda)
                
'                LstPesq.SubItems(1) = VarConv.Fields("Produto")   'SeCampoNulo(Consulta, "Agencia")
'                LstPesq.SubItems(2) = VarConv.Fields("Descrição_Produto")   'SeCampoNulo(Consulta, "Agencia")
'                LstPesq.SubItems(3) = Numero(Format(VarConv.Fields("Qtd"), " ##,##0.000"), 3)  'SeCampoNulo(Consulta, "Conta")
'                LstPesq.SubItems(4) = Numero(Format(VarConv.Fields("Valor") / VarConv.Fields("qtd"), "##,##0.000"), 3)  'SeCampoNulo(Consulta, "DataOper")
'                LstPesq.SubItems(5) = Format(VarConv.Fields("Valor"), "##,##0.00")  'SeCampoNulo(Consulta, "BomPara")
'                Soma = Soma + VarConv!Valor
'                MebTotal.Text = Format(Soma, "     ##,##0.00")'''

'                VarConv.MoveNext
'            Wend
'        End If
'        LblCard = TxtCartao
        VarItem = 0
        CodCard = TxtCartao

        TxtCartao.Enabled = False
'        TxtDados(0).SetFocus
    Else
        MsgBox "Mesa não indentificada !", vbInformation
        TxtCartao = ""
        TxtCartao.SetFocus
    End If
End If
End Sub

Private Sub TxtCartao_LostFocus()

'If TxtCartao = "" Then
'        MsgBox "Mesa não indentificada !", vbInformation
'        TxtCartao = ""
'        TxtCartao.SetFocus
'Else
'    If TxtCartao = 0 Then
'            MsgBox "Mesa Inválida !", vbInformation
'            TxtCartao = ""
'            TxtCartao.SetFocus
'    End If

'End If

If IsNumeric(TxtCartao) Then TxtCartao_KeyPress 13
TxtCartao.BackColor = &H80000005
End Sub

Private Sub txtcomplemento_GotFocus()
txtcomplemento.BackColor = &HC0FFFF
End Sub

Private Sub txtcomplemento_LostFocus()
txtcomplemento.BackColor = &H80000005
End Sub






Function LimpaPrd() As String
Dim Produto As String
Dim PrdSujo As String
Dim x As Integer
Produto = ""
PrdSujo = Left(Registros.Fields("Descrição_Produto"), 38)
For x = 1 To Len(PrdSujo)
    If Mid$(PrdSujo, x, 1) >= Chr(32) And Mid$(PrdSujo, x, 1) <= Chr(122) Then
            Produto = Produto + Mid$(PrdSujo, x, 1)
    End If
Next
LimpaPrd = psCompString(3, Produto, 38)
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"

If KeyCode = 121 Then
    LimpaTela
End If

If KeyCode = 27 Then
    Unload Me
End If

End Sub


Private Sub ListaBicos_LostFocus()
ListaBicos.BackColor = &HFFFFFF ' Branco
MebValor.BackColor = &H80FFFF ' Amarelo
MebValor.SetFocus
End Sub

Private Sub ListaPreco_GotFocus()
ListaPreco.BackColor = &H80FFFF ' Amarelo
End Sub


Private Sub ListaPreco_LostFocus()
ListaPreco.BackColor = &HFFFFFF ' Branco
End Sub



Private Sub MebValor_GotFocus()
MebValor.SelStart = 0
MebValor.SelLength = 12
End Sub

Private Sub MebUnitario_GotFocus()
MebValor.SelStart = 0
MebValor.SelLength = 12
End Sub

Private Sub MebValor_LostFocus()
MebValor.BackColor = &HFFFFFF ' Branco
FrmPreco.Visible = False
End Sub


Private Sub TxtPessoa_GotFocus()
TxtPessoa.BackColor = &HC0FFFF
End Sub


Public Sub PesquisaGeral(Tabela As String, sql As String, Campo As String)
'    FramePesquisaG.Visible = True

    DatPesquisaG.DatabaseName = Tabela
    DatPesquisaG.RecordSource = sql
'    Formulário.DatPesquisaG.DataField = "Cidade"
'    Formulário.DatPesquisaG.ListField = Campo
    DatPesquisaG.Refresh
'    DblPesquisaG.Refresh
'    Formulário.DblPesquisaG.SetFocus

End Sub

Private Sub TxtPessoa_LostFocus()

If TxtPessoa = "" Then
        MsgBox "Cliente não indentificado !", vbInformation
        TxtPessoa = ""
'        TxtPessoa.SetFocus
'Else
'    If TxtCartao = 0 Then
'            MsgBox "Mesa Inválida !", vbInformation
'            TxtCartao = ""
'            TxtCartao.SetFocus
'    End If

End If

TxtPessoa.BackColor = &H80000005

If IsNumeric(TxtPessoa.BoundText) Then
'    txtcomplemento = TxtPessoa.BoundText
    TxtCartao = TxtPessoa.BoundText
    BtConsulta.SetFocus

End If
End Sub
