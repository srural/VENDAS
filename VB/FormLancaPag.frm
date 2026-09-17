VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormLancaPag 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DatDados 
      Caption         =   "Pag"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5715
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5535
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14385
      TabIndex        =   33
      Top             =   6840
      Width           =   14415
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14415
      TabIndex        =   31
      Top             =   0
      Width           =   14415
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançar Contas a Pagar"
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
         TabIndex        =   32
         Top             =   45
         Width           =   5505
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "FormLancaPag.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   11310
      TabIndex        =   29
      Top             =   -1755
      Width           =   11310
      Begin VB.Label Label5 
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
         TabIndex        =   30
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados"
      Height          =   2445
      Left            =   270
      TabIndex        =   17
      Top             =   1305
      Width           =   13650
      Begin VB.TextBox TxtParcelas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8640
         TabIndex        =   6
         Top             =   540
         Width           =   1545
      End
      Begin VB.CommandButton CmdCusto 
         Caption         =   "Interface Centro Custo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11430
         TabIndex        =   35
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CommandButton CmdLancar 
         Caption         =   "&Lançar"
         Height          =   375
         Left            =   6570
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CmdLimpa 
         Caption         =   "&Limpa"
         Height          =   375
         Left            =   8415
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtPedido 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7290
         TabIndex        =   5
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   225
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1845
         Width           =   5910
      End
      Begin VB.ComboBox CmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormLancaPag.frx":A4F6
         Left            =   225
         List            =   "FormLancaPag.frx":A506
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   2805
      End
      Begin VB.CheckBox ChkPago 
         Caption         =   "&Pago"
         Height          =   285
         Left            =   3915
         TabIndex        =   9
         Top             =   1215
         Width           =   735
      End
      Begin VB.Data DatDados 
         Caption         =   "Conta"
         Connect         =   "Access"
         DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   11475
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT * From Cta Order By Descric"
         Top             =   450
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSDBCtls.DBCombo DbcConta 
         Bindings        =   "FormLancaPag.frx":A53A
         DataSource      =   "DatDados(0)"
         Height          =   315
         Left            =   7785
         TabIndex        =   12
         Top             =   1215
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Descric"
         BoundColumn     =   "CodCta"
         Text            =   ""
      End
      Begin VB.TextBox TxtNota 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5985
         TabIndex        =   4
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox TxtDupli 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4590
         TabIndex        =   3
         Top             =   540
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DtVenc 
         Height          =   315
         Left            =   225
         TabIndex        =   7
         Top             =   1200
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90832897
         CurrentDate     =   36759
      End
      Begin MSMask.MaskEdBox MebValor 
         DataField       =   "Custo"
         Height          =   315
         Left            =   2070
         TabIndex        =   8
         Tag             =   "S"
         Top             =   1200
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker Dtpgto 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   1215
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90832897
         CurrentDate     =   36759
      End
      Begin MSComCtl2.DTPicker DTData 
         Height          =   315
         Left            =   3150
         TabIndex        =   2
         Top             =   540
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90832897
         CurrentDate     =   36759
      End
      Begin MSMask.MaskEdBox MebCta 
         Height          =   315
         Left            =   6390
         TabIndex        =   11
         Top             =   1215
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.###.###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Numero de Parcelas"
         Height          =   195
         Left            =   8640
         TabIndex        =   36
         Top             =   315
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Pedido/Nota"
         Height          =   195
         Left            =   7290
         TabIndex        =   28
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código Conta"
         Height          =   195
         Left            =   6390
         TabIndex        =   27
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Observação"
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   25
         Top             =   315
         Width           =   315
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Data &Lançamento"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   24
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Data Pagamento"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   23
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Descrição Conta"
         Height          =   195
         Left            =   7785
         TabIndex        =   22
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Valor"
         Height          =   195
         Index           =   5
         Left            =   2070
         TabIndex        =   21
         Top             =   990
         Width           =   360
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Data &Vencimento"
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   20
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Do&cumento"
         Height          =   195
         Left            =   5985
         TabIndex        =   19
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Duplicata"
         Height          =   195
         Left            =   4590
         TabIndex        =   18
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Forncedor"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6255
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 Or Tipo = 4 ORDER BY Ent.nome"
      Top             =   810
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormLancaPag.frx":A554
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   855
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "nome"
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormLancaPag.frx":A56E
      Height          =   2580
      Left            =   270
      OleObjectBlob   =   "FormLancaPag.frx":A588
      TabIndex        =   34
      Top             =   3915
      Width           =   13650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   630
      Width           =   810
   End
End
Attribute VB_Name = "FormLancaPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public Pcl As Byte


Private Sub ChkPago_Click()
If ChkPago.value = 1 Then
    Dtpgto.Enabled = True
Else
    Dtpgto.Enabled = False
End If
End Sub

Private Sub CmdCusto_Click()

Dim Sql As String

If MsgBox("Deseja lançar valores no centro de custo ?", vbYesNo, App.Title) = vbYes Then
                                  
        Sql = "INSERT INTO CCU (Filial, Data, Tipo, Historico, Valor ) VALUES ( "
        Sql = Sql & 1 & " , #"
        Sql = Sql & Format(DtData, "mm/dd/yyyy") & "#,"
        Sql = Sql & """" & "D" & """" & " ,"
        Sql = Sql & """" & TxtObs & """" & " ,"
        Sql = Sql & Numero2(MebValor) & ")"
        Banco.Execute Sql
    
    FormCentroCusto.Show
End If

End Sub

Private Sub CmdLancar_Click()
Dim Sql As String
Dim Tipo As String
Dim Caixa As Recordset
Dim X As Integer
Dim VarDias As Integer
'On Error GoTo TrataErro
If IsNull(MebValor) Or MebValor = "" Then
    MsgBox "Valor não pode ser nulo !", vbExclamation, App.Title
    MebValor.SetFocus
    Exit Sub
End If

VarDias = 0

If TxtPedido = "" Then
    TxtPedido = 0
End If

If ChkPago.value = 1 Then
    Sql = "INSERT INTO PAG (Fornecedor, Data, DtVenc, Valor, VlrPago, Numero , NroNota ,Obs, Conta, CodCta, Tipo, DtPgto, Pedido, Cond ) VALUES ( "
    Sql = Sql & DatDados(1).Recordset.Fields("CodEntidade") & " , #"
    Sql = Sql & Format(DtData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(DtVenc, "mm/dd/yyyy") & "# , "
    Sql = Sql & Numero2(MebValor.Text) & " ,"
    Sql = Sql & Numero2(MebValor.Text) & " ,"
    Sql = Sql & """" & (TxtDupli) & """" & " ,"
    Sql = Sql & """" & (TxtNota) & """" & " ,"
    Sql = Sql & """" & (TxtObs) & """" & " ,"
    Sql = Sql & """" & (DbcConta) & """" & " ,"
    Sql = Sql & """" & (MebCta) & """" & " ,"
    Sql = Sql & """" & CmbTipo.ListIndex & """" & " ,#"
    Sql = Sql & Format(Dtpgto, "mm/dd/yyyy") & "# , "
    Sql = Sql & TxtPedido & " ,"
    Sql = Sql & ChkPago & ")"
    Banco.Execute Sql
    
    MsgBox "Lançamento feito com sucesso!", vbInformation, App.Title
    
    If MsgBox("O dinheiro está sendo pago pelo caixa (gaveta) ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
    
        Set Caixa = Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
        Caixa.Requery
        
        If Caixa.RecordCount > 0 Then
            Caixa.MoveLast
            VarCaixa = Caixa!CodControle
        End If
    
        Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
        Sql = Sql + """" & Maquina & """" & ", "
        Sql = Sql & """" & Format(Trim(DbcConta) & " - " & Trim(TxtObs), ">") & """" & " , #"
        Sql = Sql & Format(Dtpgto, "mm/dd/yyyy") & "# , #"
        Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
        Sql = Sql & Numero2(MebValor.Text) & " ,"
        Sql = Sql & VarCaixa & " ,"
        Sql = Sql & """" & "D" & """" & ")"
        Banco.Execute Sql
            
    End If
    
    If InterFinanc = True Then
          FrmMov.MebDados(3).Text = MebValor
          FrmMov.MebDados(10).Text = Dtpgto
          FrmMov.TxtDados(0).Text = "N: " & Trim(TxtDupli) & " | " & Left(DbcPesquisa.Text, 20)
          FrmMov.Show
    End If
        
Else
    If TxtParcelas = "" Then TxtParcelas = 1
    
    For X = 1 To TxtParcelas
        Sql = "INSERT INTO PAG (Fornecedor, Data, DtVenc, Valor, Numero , NroNota ,Obs, Conta, CodCta, Tipo, Pedido, Cond ) VALUES ( "
        Sql = Sql & DatDados(1).Recordset.Fields("CodEntidade") & " , #"
        Sql = Sql & Format(DtData, "mm/dd/yyyy") & "# , #"
        Sql = Sql & Format(DtVenc, "mm/dd/yyyy") & "# , "
        Sql = Sql & Numero2(MebValor.Text / CDbl(TxtParcelas)) & " ,"
        If TxtParcelas = 1 Then
            Sql = Sql & """" & (TxtDupli) & """" & " ,"
        Else
            Sql = Sql & """" & (TxtDupli) & "/" & X & """" & " ,"
        End If
        Sql = Sql & """" & (TxtNota) & """" & " ,"
        Sql = Sql & """" & (TxtObs) & """" & " ,"
        Sql = Sql & """" & (DbcConta) & """" & " ,"
        Sql = Sql & """" & (MebCta) & """" & " ,"
        Sql = Sql & """" & CmbTipo.ListIndex & """" & " ,"
        Sql = Sql & TxtPedido & " ,"
        Sql = Sql & ChkPago & ")"
        Banco.Execute Sql
        DtVenc = DtVenc + 30
    Next
    DatDados(2).RecordSource = " Select * from Pag Where Fornecedor = " & DatDados(1).Recordset.Fields("CodEntidade") & " and Pedido =  " & TxtPedido & " Order By CodDup"
    DatDados(2).Refresh
End If


Pcl = Pcl + 1

If TxtPedido <> "" Then
    TxtDupli = TxtPedido & "/" & Pcl
End If

MebValor = ""
'TxtDupli = ""
TxtNota = ""
DbcConta = ""
TxtObs = ""
Dtpgto.Enabled = False
ChkPago.value = 0
DbcPesquisa.SetFocus

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdLimpa_Click()
MebValor = ""
TxtDupli = ""
TxtNota = ""
TxtObs = ""
DbcConta = ""
TxtParcelas = 1

DatDados(2).RecordSource = "Select * from Pag Where Fornecedor = 0 "
DatDados(2).Refresh

DbcPesquisa = ""
DbcPesquisa.SetFocus

End Sub


Private Sub Command1_Click()

End Sub

Private Sub DbcConta_Change()
TxtObs = DbcConta
End Sub

Private Sub DbcConta_LostFocus()
If IsNumeric(DbcConta.BoundText) Then
    MebCta = DbcConta.BoundText

End If
End Sub

Private Sub DbcPesquisa_LostFocus()
 On Error Resume Next
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

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String
   Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
   Criterio = Criterio & Codigo
   DatDados(1).Recordset.FindFirst Criterio
   If DatDados(1).Recordset.NoMatch Then
       MsgBox "Cliente não Localizado", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
       DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
       TxtParcelas = 1
'       TxtDupli.SetFocus
   End If

End Sub

Private Sub DTData_LostFocus()
DtVenc = DtData + 30
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_Activate()
'Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

DtData = Date
DtVenc = Date
Dtpgto = Date
Dtpgto.Enabled = False
CmbTipo.ListIndex = 0

End Sub

Private Sub Form_Load()
Dim Componente As Variant

'Set Registros = Banco.OpenRecordset("SELECT * FROM Pag")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

Pcl = 1

DatDados(2).RecordSource = "Select * from Pag Where Fornecedor = 0 "
DatDados(2).Refresh

If TxtPedido <> "" Then
    TxtDupli = TxtPedido & "/" & Pcl
End If

 On Error Resume Next
End Sub

Private Sub MebCta_Change()
Dim Criterio As String
If IsNumeric(MebCta) Then
    Criterio = "CodCta like " & """" & MebCta & "*"""
    DatDados(0).Recordset.FindFirst Criterio
    If Not DatDados(0).Recordset.NoMatch Then
        DbcConta.Text = DatDados(0).Recordset("Descric")
    End If
End If
End Sub
