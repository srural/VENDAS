VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormBaixarPag 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   14040
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -315
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   46
      Top             =   7920
      Width           =   14370
   End
   Begin VB.CommandButton cmdped 
      Caption         =   "Consulta"
      Height          =   330
      Left            =   10395
      TabIndex        =   44
      Top             =   810
      Width           =   915
   End
   Begin VB.TextBox TxtPed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8910
      TabIndex        =   42
      Top             =   810
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14055
      TabIndex        =   40
      Top             =   0
      Width           =   14055
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Baixar Contas a Pagar"
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
         TabIndex        =   41
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.CommandButton CmdBaixaMulti 
      Caption         =   "Baixa Multipla de duplicatas"
      Height          =   375
      Left            =   8685
      TabIndex        =   39
      Top             =   5850
      Width           =   2670
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   135
      TabIndex        =   27
      Top             =   6300
      Width           =   11265
      Begin VB.CommandButton CmdPgto 
         Caption         =   "&Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9960
         TabIndex        =   38
         Top             =   990
         Width           =   1185
      End
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         TabIndex        =   8
         Top             =   990
         Width           =   3195
      End
      Begin VB.TextBox TxtDupl 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox TxtNota 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1428
         TabIndex        =   3
         Top             =   390
         Width           =   1080
      End
      Begin MSMask.MaskEdBox MebPagto 
         Height          =   330
         Left            =   6390
         TabIndex        =   11
         Top             =   990
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox ChkPago 
         Caption         =   "&Pago"
         Height          =   285
         Left            =   5310
         TabIndex        =   10
         Top             =   990
         Width           =   870
      End
      Begin VB.ComboBox CmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormBaixarPag.frx":0000
         Left            =   7920
         List            =   "FormBaixarPag.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   390
         Width           =   3210
      End
      Begin MSMask.MaskEdBox MebPago 
         DataField       =   "Custo"
         Height          =   330
         Left            =   7920
         TabIndex        =   12
         Tag             =   "S"
         Top             =   990
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebValor 
         DataField       =   "Custo"
         Height          =   330
         Left            =   6102
         TabIndex        =   6
         Tag             =   "S"
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebCta 
         Height          =   315
         Left            =   3555
         TabIndex        =   9
         Top             =   990
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##.###.###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   330
         Left            =   2781
         TabIndex        =   4
         Top             =   390
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebVenc 
         Height          =   330
         Left            =   4464
         TabIndex        =   5
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código Conta"
         Height          =   195
         Left            =   3555
         TabIndex        =   37
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Observação"
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo"
         Height          =   195
         Index           =   7
         Left            =   7920
         TabIndex        =   35
         Top             =   180
         Width           =   315
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         Height          =   195
         Index           =   6
         Left            =   6120
         TabIndex        =   34
         Top             =   180
         Width           =   360
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Vencimento"
         Height          =   195
         Index           =   5
         Left            =   4455
         TabIndex        =   33
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data"
         Height          =   195
         Index           =   4
         Left            =   2790
         TabIndex        =   32
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota"
         Height          =   195
         Left            =   1425
         TabIndex        =   31
         Top             =   180
         Width           =   510
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         Height          =   195
         Index           =   3
         Left            =   7920
         TabIndex        =   30
         Top             =   765
         Width           =   780
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Pagamento"
         Height          =   195
         Index           =   15
         Left            =   6390
         TabIndex        =   29
         Top             =   765
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicata"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.CheckBox ChkPagos 
      Caption         =   "&Pagos"
      Height          =   285
      Left            =   7650
      TabIndex        =   1
      Top             =   855
      Width           =   870
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6030
      TabIndex        =   25
      Top             =   855
      Width           =   1275
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Duplicata"
      Height          =   330
      Left            =   12330
      TabIndex        =   15
      Top             =   810
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   11700
      TabIndex        =   21
      Top             =   5850
      Width           =   2175
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   300
         Left            =   135
         TabIndex        =   17
         Tag             =   "S"
         Top             =   1575
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPagos 
         DataField       =   "Custo"
         Height          =   300
         Left            =   135
         TabIndex        =   18
         Tag             =   "S"
         Top             =   945
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebAPagar 
         DataField       =   "Custo"
         Height          =   300
         Left            =   135
         TabIndex        =   19
         Tag             =   "S"
         Top             =   360
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         X1              =   0
         X2              =   4635
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   2160
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicatas à Pagar"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   24
         Top             =   135
         Width           =   1350
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicatas Pagas"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   22
         Top             =   1350
         Width           =   360
      End
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa"
      Height          =   405
      Left            =   1800
      TabIndex        =   13
      Top             =   5895
      Width           =   1320
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   180
      TabIndex        =   14
      Top             =   5895
      Width           =   1320
   End
   Begin VB.Data DatDados 
      Caption         =   "Pag"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6930
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3285
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormBaixarPag.frx":0044
      Height          =   4020
      Left            =   180
      OleObjectBlob   =   "FormBaixarPag.frx":005E
      TabIndex        =   16
      Top             =   1260
      Width           =   13695
   End
   Begin VB.Data DatDados 
      Caption         =   "Fornecedor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1620
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 Or Tipo = 4 ORDER BY Ent.nome"
      Top             =   855
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormBaixarPag.frx":1C81
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   5490
      _ExtentX        =   9684
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
   Begin VB.Label LblFor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   45
      Top             =   5400
      Width           =   5775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      Height          =   195
      Left            =   8910
      TabIndex        =   43
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   6030
      TabIndex        =   26
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   630
      Width           =   810
   End
End
Attribute VB_Name = "FormBaixarPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public VarCodDup As Double

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub ChkPago_Click()
If ChkPago.value = 1 Then
    MebPagto = Date
    MebPago = MebValor
Else
    MebPagto = "__/__/____"
    MebPago = 0
End If
End Sub

Private Sub ChkPagos_Click()
'If TxtCod <> "" Then
    LIMPACampo
    AtualizaGrid
'End If
End Sub

Private Sub AtualizaGrid()

If ChkPagos.value = False Then
    DatDados(0).RecordSource = "SELECT * FROM Pag WHERE Fornecedor = " & Val(TxtCod) & " AND Cond = FALSE ORDER BY DtVenc"
Else
    DatDados(0).RecordSource = "SELECT * FROM Pag WHERE Fornecedor = " & Val(TxtCod) & " ORDER BY DtVenc"
End If
DatDados(0).Refresh
If DatDados(0).Recordset.RecordCount > 0 Then
     DatDados(0).Recordset.MoveLast
     CmdBaixaMulti.Enabled = True
Else
     CmdBaixaMulti.Enabled = False
End If
SomaGrid

End Sub

Private Sub CmdBaixaMulti_Click()

'If DatDados(0).Recordset.RecordCount > 0 Then
    FormConsolidaPgt.Show
'Else
'    MsgBox "Nao tem duplicatas em aberto", vbInformation
'End If

End Sub

Private Sub CmdExcluir_Click()
Dim Criterio As String
If DatDados(0).Recordset.RecordCount <> 0 And VarCodDup <> 0 Then
    If MsgBox("Confirma Exclusão da Duplicata", vbYesNo, App.Title) = vbYes Then
        Open Caminho & "\Delete.txt" For Append As #1
        Print #1, "Duplicata; " & VarCodDup & "; " & DbcPesquisa.Text & "; " & Date & "; " & Time & "; " & Maquina & "; " & Alinhar(MebValor, 10, 2)
        Close #1
        Criterio = "CodDup  = " & VarCodDup
        DatDados(0).Recordset.FindFirst Criterio
        If Not DatDados(0).Recordset.NoMatch Then
            DatDados(0).Recordset.Delete
        End If
        DatDados(0).Refresh
        SomaGrid
        MebPago = ""
        TxtDupl = ""
        VarCodDup = 0
    End If
Else
    MsgBox "Duplicata não selecionada", vbInformation, App.Title
End If

End Sub

Private Sub CmdImprimir_Click()

MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELPAG.RPT"
If ChkPagos.value = 1 Then
    MDIPrincipal.RptRel.SelectionFormula = "{Pag.Fornecedor} = " & DbcPesquisa.BoundText
Else
    MDIPrincipal.RptRel.SelectionFormula = "{Pag.Cond} = False and {Pag.Fornecedor} = " & DbcPesquisa.BoundText
End If
MDIPrincipal.RptRel.RetrieveDataFiles
MDIPrincipal.RptRel.PrintReport

End Sub

Private Sub CmdLimpa_Click()
LIMPACampo
MebPago = 0
MebSubTot = 0
MebPagos = 0
MebAPagar = 0
DbcPesquisa = ""
TxtCod.Text = ""
DatDados(0).RecordSource = "SELECT * FROM Pag  WHERE Fornecedor = " & Val(DbcPesquisa.BoundText) & " ORDER BY DtVenc"
DatDados(0).Refresh
SomaGrid

DbcPesquisa.SetFocus

End Sub

Private Sub LIMPACampo()

TxtDupl = ""
TxtNota = ""
'DtData = Date
'DtVenc = Date
'DtPagto = Date
MebValor = ""
MebPago = ""
TxtObs = ""
MebCta = "__.___.___"
CmbTipo.ListIndex = 0

End Sub

Private Sub cmdped_Click()
If ChkPagos.value = False Then
    DatDados(0).RecordSource = "SELECT * FROM Pag WHERE Pedido = " & Val(TxtPed) & " AND Cond = FALSE ORDER BY DtVenc"
Else
    DatDados(0).RecordSource = "SELECT * FROM Pag WHERE Pedido = " & Val(TxtPed) & " ORDER BY DtVenc"
End If
DatDados(0).Refresh
If DatDados(0).Recordset.RecordCount > 0 Then
     DatDados(0).Recordset.MoveLast
     CmdBaixaMulti.Enabled = Trueejs
Else
     CmdBaixaMulti.Enabled = False
End If
SomaGrid

End Sub

Private Sub CmdPgto_Click()
Dim Criterio As String
Dim Sql As String
Dim caxia As Recordset
Dim VarCaixa As Double
If DatDados(0).Recordset.RecordCount <> 0 And MebValor <> "" Then
    If MsgBox("Confirma ?", vbYesNo, App.Title) = vbYes Then
        Criterio = "CodDup = " & DBGrid1.Columns(0)
        DatDados(0).Recordset.FindFirst Criterio
        If Not DatDados(0).Recordset.NoMatch Then
            DatDados(0).Recordset.Edit
            DatDados(0).Recordset.Fields("Cond") = ChkPago
            If IsDate(MebPagto) Then
                DatDados(0).Recordset.Fields("DtPgto") = MebPagto
            Else
                DatDados(0).Recordset.Fields("DtPgto") = Null
            End If
            DatDados(0).Recordset.Fields("VlrPago") = MebPago
            DatDados(0).Recordset.Fields("Numero") = TxtDupl
            DatDados(0).Recordset.Fields("NroNota") = TxtNota
            DatDados(0).Recordset.Fields("Data") = MebData
            DatDados(0).Recordset.Fields("DtVenc") = MebVenc
            DatDados(0).Recordset.Fields("Valor") = MebValor
            DatDados(0).Recordset.Fields("Tipo") = CmbTipo.ListIndex
            DatDados(0).Recordset.Fields("Obs") = TxtObs
            DatDados(0).Recordset.Fields("CodCta") = MebCta
            DatDados(0).Recordset.Update
        End If
        DatDados(0).Refresh
        SomaGrid
        If MsgBox("O dinheiro está sendo pago pelo caixa (gaveta) ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        
            Set Caixa = Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
            Caixa.Requery
            
            If Caixa.RecordCount > 0 Then
                Caixa.MoveLast
                VarCaixa = Caixa!CodControle
            End If
        
            Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
            Sql = Sql + """" & Maquina & """" & ", "
            Sql = Sql & """" & Format(TxtObs, ">") & """" & " , #"
            Sql = Sql & Format(MebPagto, "mm/dd/yyyy") & "# , #"
            Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
            Sql = Sql & Numero2(MebPago.Text) & " ,"
            Sql = Sql & VarCaixa & " ,"
            Sql = Sql & """" & "D" & """" & ")"
            Banco.Execute Sql
                
        End If
        
        If InterFinanc = True Then
                FrmMov.MebDados(3).Text = MebValor
                FrmMov.MebDados(10).Text = MebPagto
                FrmMov.TxtDados(0).Text = "N: " & TxtDupl & " | " & Left(DbcPesquisa.Text, 20)
                FrmMov.Show
        End If
        MebPago = ""
        TxtDupl = ""
        LIMPACampo
    End If
Else
    MsgBox "Clique duas vezes na duplicata que deseja pagar", vbInformation, App.Title
End If

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

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String
Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
Criterio = Criterio & Codigo
DatDados(1).Recordset.FindFirst Criterio
If DatDados(1).Recordset.NoMatch Then
    MsgBox "Fornecedor não Localizado", vbExclamation, App.Title
    DbcPesquisa.Text = ""
    DbcPesquisa.SetFocus
Else
    DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
    TxtCod.Text = Codigo
    AtualizaGrid
End If

End Sub

Private Sub DBGrid1_DBLClick()
LIMPACampo
If DatDados(0).Recordset.RecordCount <> 0 Then
    VarCodDup = DBGrid1.Columns(0)
    TxtDupl.Text = DBGrid1.Columns(1)
    TxtNota.Text = DBGrid1.Columns(2)
    MebData = DBGrid1.Columns(4)
    MebVenc = DBGrid1.Columns(5)
    MebValor = DBGrid1.Columns(6)
    CmbTipo.ListIndex = DBGrid1.Columns(12)
    TxtObs = DBGrid1.Columns(10)
    If DBGrid1.Columns(11) <> "" Then
        MebCta = DBGrid1.Columns(11)
    End If
    If DBGrid1.Columns(7) = "Sim" Then
        ChkPago = 1
    Else
        ChkPago = 0
    End If
    If ChkPago.value = 1 Then
        MebPago.Text = DBGrid1.Columns(9)
        MebPagto = DBGrid1.Columns(8)
    Else
    '    DtPagto = Date
        MebPago = 0
    End If
Else
    MsgBox "Duplicata não selecionada", vbInformation, App.Title
End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DatDados(0).Recordset.RecordCount > 0 Then
'        TxtDados(0).Text = DatDados(5).Recordset("CodPed")
        DatDados(1).Recordset.FindFirst "CodEntidade = " & DatDados(0).Recordset("Fornecedor")
        LblFor = DatDados(1).Recordset("Nome")
'        DBGrid1.SetFocus
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyPress = 0
    End If

End Sub

Private Sub Form_Load()
Dim Componente As Variant

Set Registros = Banco.OpenRecordset("SELECT * FROM Pag")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

'Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

DatDados(1).RecordSource = "SELECT Ent.CodEntidade,Ent.nome  FROM Ent Where Tipo = 2 Or Tipo = 5 Or Tipo = 4 ORDER BY Ent.nome"
DatDados(1).Refresh

DatDados(0).RecordSource = "SELECT * FROM Pag WHERE Fornecedor = " & Val(TxtCod) & " AND Cond = FALSE ORDER BY DtVenc"



'DtData = Date
MebPago = 0
VarCodDup = 0

End Sub

Private Sub MebPago_GotFocus()
MebPago = ""
End Sub

Private Sub SomaGrid()
Dim SomaGrid As Recordset
Dim Pago  As Currency
Dim Apagar As Currency
TotSub = 0
Pago = 0
Apagar = 0

Set SomaGrid = DatDados(0).Recordset.Clone
Do Until SomaGrid.Eof
    If GetInputState() Then DoEvents

    If SomaGrid.Fields("Cond") = -1 Then
        Pago = Pago + SomaGrid.Fields("VlrPago")
    Else
        Apagar = Apagar + SomaGrid.Fields("Valor")
    End If
   SomaGrid.MoveNext
Loop
MebPagos.Text = Pago
MebAPagar.Text = Apagar
MebSubTot = Pago + Apagar

End Sub

