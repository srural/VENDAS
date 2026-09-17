VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormBalancete 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13125
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   20
      Top             =   7920
      Width           =   14370
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
      TabIndex        =   18
      Top             =   0
      Width           =   14055
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Balancete"
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
         TabIndex        =   19
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gráfico"
      Height          =   5625
      Left            =   10530
      TabIndex        =   11
      Top             =   2070
      Width           =   2310
      Begin VB.Shape ShVista 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   4740
         Left            =   150
         Top             =   390
         Width           =   540
      End
      Begin VB.Shape SHRecebe 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   4740
         Left            =   1590
         Top             =   390
         Width           =   540
      End
      Begin VB.Shape ShPrazo 
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   4740
         Left            =   870
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vista"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   5250
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Prazo"
         Height          =   195
         Left            =   945
         TabIndex        =   13
         Top             =   5250
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rec."
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   5250
         Width           =   345
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6285
      Left            =   180
      TabIndex        =   8
      Top             =   1395
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   11086
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão"
      Height          =   1095
      Left            =   10530
      TabIndex        =   7
      Top             =   765
      Width           =   2325
      Begin VB.CommandButton BtLaser 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   165
         TabIndex        =   6
         Top             =   645
         Width           =   1995
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Visualiza"
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1995
      End
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   85852161
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   85852161
      CurrentDate     =   36641
   End
   Begin MSMask.MaskEdBox MebDebito 
      Height          =   315
      Left            =   7290
      TabIndex        =   4
      Top             =   975
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebCredito 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   975
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebAcrescimo 
      Height          =   315
      Left            =   5430
      TabIndex        =   3
      Top             =   975
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Faturamento Extra"
      Height          =   195
      Left            =   5430
      TabIndex        =   17
      Top             =   750
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Receita Extra"
      Height          =   195
      Left            =   3645
      TabIndex        =   16
      Top             =   750
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despesa &Extra"
      Height          =   195
      Left            =   7290
      TabIndex        =   15
      Top             =   750
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   705
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   1890
      TabIndex        =   9
      Top             =   750
      Width           =   720
   End
End
Attribute VB_Name = "FormBalancete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VdaVista As Recordset
Dim VdaPrazo As Recordset
Dim Receb As Recordset
Dim Pagto As Recordset
Dim RecebFut As Recordset
Dim PagtoFut As Recordset
Dim Rec As Recordset
Dim PgtoCta As Recordset
Dim RegCta As Recordset
Dim RegVda As Recordset
Dim RegGru As Recordset

Public lstPesq As ListItem

Private Sub BtImprimir_Click()
Dim Sql As String
Dim VarVdaVista As Currency
Dim VarVdaPrazo As Currency
Dim VarDescVista As Currency
Dim VarDescPrazo As Currency
Dim VarReceb As Currency
Dim VarDescReceb As Currency
Dim VarJurosReceb As Currency
Dim VarRecebFut As Currency
Dim VarDescRecebFut As Currency
Dim VarPagto As Currency
Dim VarJurosPagto As Currency
Dim VarPagtoFut As Currency
Dim VarRecebAtras As Currency

Sql = "SELECT Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Sum(DUP.Desconto) AS SomaDeDesconto, Sum(DUP.Juros) AS SomaDeJuros FROM DUP WHERE (((DUP.DtVenc)>=#  " & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=# " & Format(DtData2, "mm/dd/yyyy") & " #)) GROUP BY DUP.Cond HAVING DUP.Cond =FALSE"
VarVdaVista = 0
VarVdaPrazo = 0
VarDescVista = 0
VarReceb = 0
VarDescReceb = 0
VarJurosReceb = 0
VarRecebFut = 0
VarDescRecebFut = 0
VarPagto = 0
VarJurosPagto = 0
VarPagtoFut = 0
VarDescPrazo = 0
VarRecebAtras = 0

ShVista.Height = 1
ShPrazo.Height = 1
SHRecebe.Height = 1

DoEvents

Lista.ListItems.Clear

Set VdaVista = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao = 1 )  AND PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# ")
VdaVista.Requery
If VdaVista.RecordCount > 0 Then
    If Not IsNull(VdaVista!SomaDeTotal) Then
        VarVdaVista = VdaVista!SomaDeTotal
        VarDescVista = VdaVista!SomaDeDesconto
    End If
End If

Set VdaPrazo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao = 2 )  AND PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# ")
'Sql = "SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao = 2 )  AND PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# "

VdaPrazo.Requery
If VdaPrazo.RecordCount > 0 Then
    If Not IsNull(VdaPrazo!SomaDeTotal) Then
        VarVdaPrazo = VdaPrazo!SomaDeTotal
        VarDescPrazo = VdaPrazo!SomaDeDesconto
    End If
End If

'Set Receb = Banco.OpenRecordset("SELECT Sum(DUP.VlrPago) AS SomaDeVlrPago, DUP.Cond, Sum(DUP.Valor) AS SomaDeValor, Sum(DUP.Desconto) AS SomaDeDesconto FROM DUP WHERE (((DUP.DtPgto)>=#  " & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=# " & Format(DtData2, "mm/dd/yyyy") & " #)) GROUP BY DUP.Cond HAVING DUP.Cond =True")
Set Receb = Banco.OpenRecordset("SELECT Sum(REC.Desconto) AS SomaDeDesconto, Sum(Rec.Valor) as SomaDeValor   FROM REC WHERE REC.Data >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND REC.Data <= #" & Format(DtData2, "mm/dd/yyyy") & "#  AND REC.Tipo < 10")
Receb.Requery
If Receb.RecordCount > 0 Then
    If Not IsNull(Receb!SomaDeValor) Then
        VarReceb = Receb!SomaDeValor ' - Receb!SomaDeDesconto
        VarDescReceb = Receb!SomaDeDesconto ' (Receb!SomaDeValor - Receb!SomaDeVlrPago) +
    End If
End If

Set Pagto = Banco.OpenRecordset("SELECT Sum(PAG.VlrPago) AS SomaDeVlrPago, Sum(PAG.Valor) AS SomaDeValor, PAG.Cond FROM PAG WHERE (((PAG.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PAG.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PAG.Cond HAVING (((PAG.Cond)=True))")
Pagto.Requery
If Pagto.RecordCount > 0 Then
    VarPagto = Pagto!SomaDeVlrPago
    If Pagto!SomaDeVlrPago > Pagto!SomaDeValor Then
        VarJurosPagto = Pagto!SomaDeVlrPago - Pagto!SomaDeValor
    End If
End If

'DtData1 = DtData1 - 5000
'DtData2 = DtData2 + 5000

Set RecebFut = Banco.OpenRecordset("SELECT Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Sum(DUP.Desconto) AS SomaDeDesconto, Sum(DUP.Juros) AS SomaDeJuros FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade WHERE (((DUP.DtVenc)>=#  " & Format(DtData1 - 5000, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=# " & Format(DtData2, "mm/dd/yyyy") & " #)) GROUP BY DUP.Cond HAVING DUP.Cond =FALSE")
RecebFut.Requery

Sql = "SELECT Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Sum(DUP.Desconto) AS SomaDeDesconto, Sum(DUP.Juros) AS SomaDeJuros FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade WHERE (((DUP.DtVenc)>=#  " & Format(DtData1 - 5000, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=# " & Format(DtData2, "mm/dd/yyyy") & " #)) GROUP BY DUP.Cond HAVING DUP.Cond =FALSE"
If RecebFut.RecordCount > 0 Then
    VarRecebAtras = RecebFut!SomaDeValor
End If


Set RecebFut = Banco.OpenRecordset("SELECT Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Sum(DUP.Desconto) AS SomaDeDesconto, Sum(DUP.Juros) AS SomaDeJuros FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade WHERE (((DUP.DtVenc)>=#  " & Format(DtData2, "mm/dd/yyyy") & "# And (DUP.DtVenc)<=# " & Format(DtData2 + 5000, "mm/dd/yyyy") & " #)) GROUP BY DUP.Cond HAVING DUP.Cond =FALSE")
RecebFut.Requery
If RecebFut.RecordCount > 0 Then
    VarRecebFut = RecebFut!SomaDeValor
End If


Set PagtoFut = Banco.OpenRecordset("SELECT Sum(PAG.VlrPago) AS SomaDeVlrPago, Sum(PAG.Valor) AS SomaDeValor, PAG.Cond FROM PAG WHERE (((PAG.DtVenc)>=#" & Format(DtData1 - 5000, "mm/dd/yyyy") & "# And (PAG.DtVenc)<=#" & Format(DtData2 + 5000, "mm/dd/yyyy") & "#)) GROUP BY PAG.Cond HAVING (((PAG.Cond)=False))")
Pagto.Requery
If PagtoFut.RecordCount > 0 Then
    VarPagtoFut = PagtoFut!SomaDeValor
End If



'Set Rec = Banco.OpenRecordset("SELECT Sum(REC.Desconto) AS SomaDeDesconto  FROM REC WHERE REC.Data >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND REC.Data <= #" & Format(DtData2, "mm/dd/yyyy") & "# ")
'Rec.Requery
'If Rec.RecordCount > 0 Then
'    VarDescReceb = Rec!SomaDeDesconto
'End If

ShVista.Height = 4950 * VarVdaVista / (VarVdaVista + VarVdaPrazo + VarReceb)
ShPrazo.Height = 4950 * VarVdaPrazo / (VarVdaVista + VarVdaPrazo + VarReceb)
SHRecebe.Height = 4950 * VarReceb / (VarVdaVista + VarVdaPrazo + VarReceb)

ShVista.Top = 4950 - (4950 * VarVdaVista / (VarVdaVista + VarVdaPrazo + VarReceb))
ShPrazo.Top = 4950 - (4950 * VarVdaPrazo / (VarVdaVista + VarVdaPrazo + VarReceb))
SHRecebe.Top = 4950 - (4950 * VarReceb / (VarVdaVista + VarVdaPrazo + VarReceb))


Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "VENDA A VISTA"
lstPesq.SubItems(2) = "(+)"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarVdaVista), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "RECEBIMENTO"
lstPesq.SubItems(2) = "(+)"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarReceb), 12, 2), "##,###,##0.00")

If MebCredito > 0 Then
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "RECEITA EXTRA"
    lstPesq.SubItems(2) = "(+)"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(MebCredito), 12, 2), "##,###,##0.00")
End If

Set lstPesq = Lista.ListItems.Add(, , 1)
'lstPesq.SubItems(1) = Space(20)
'lstPesq.SubItems(2) = Space(5)
lstPesq.SubItems(3) = "==========="

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "SUB TOTAL DE RECEITAS"
lstPesq.SubItems(2) = "(=)"
lstPesq.SubItems(3) = Format(Alinhar(CStr((VarVdaVista + VarReceb + CCur(MebCredito))), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "PAGAMENTOS"
lstPesq.SubItems(2) = "(-)"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarPagto), 12, 2), "##,###,##0.00")

If MebDebito > 0 Then
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "DESPESA EXTRA"
    lstPesq.SubItems(2) = "(-)"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(MebDebito), 12, 2), "##,###,##0.00")
End If

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(3) = "==========="

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "TOTAL (Vista + Recebimento + Receita Extra - Pagamento - Despesa Extra)"
lstPesq.SubItems(2) = "(=)"
lstPesq.SubItems(3) = Format(Alinhar(CStr((VarVdaVista + VarReceb + CCur(MebCredito)) - (VarPagto + CCur(MebDebito))), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "VENDA A PRAZO"
lstPesq.SubItems(2) = " "
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarVdaPrazo), 12, 2), "##,###,##0.00")

If MebAcrescimo > 0 Then
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "FATURAMENTO EXTRA"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = Format(Alinhar(CStr(MebAcrescimo), 12, 2), "##,###,##0.00")
End If

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "FATURAMENTO (Venda Vista + Venda a Prazo + Faturamento Extra)"
lstPesq.SubItems(2) = " "
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarVdaPrazo + VarVdaVista + CCur(MebAcrescimo)), 12, 2), "##,###,##0.00")


Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "CONTAS A RECEBER VENCIDAS"
lstPesq.SubItems(2) = " "
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarRecebAtras), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "CONTAS A RECEBER À VENCER"
lstPesq.SubItems(2) = " "
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarRecebFut), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(3) = "==========="

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "TOTAL A RECEBER"
lstPesq.SubItems(2) = "(=)"
lstPesq.SubItems(3) = Format(Alinhar(CStr((VarRecebAtras) + CCur(VarRecebFut)), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "CONTAS A PAGAR"
lstPesq.SubItems(2) = " "
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarPagtoFut), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "DESCONTO DE VENDA VISTA"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarDescVista), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "DESCONTO DE VENDA PRAZO"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarDescPrazo), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "DESCONTO DE RECEBIMENTO"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarDescReceb), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "TOTAL DOS DESCONTOS"
lstPesq.SubItems(3) = Format(Alinhar(CStr(CCur(VarDescReceb) + CCur(VarDescVista) + CCur(VarDescPrazo)), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)


Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "JUROS RECEBIDOS"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarJurosReceb), 12, 2), "##,###,##0.00")

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "JUROS PAGOS"
lstPesq.SubItems(3) = Format(Alinhar(CStr(VarJurosPagto), 12, 2), "##,###,##0.00")

Set RegCta = Banco.OpenRecordset("Select * From Cta where right(TRIM(codcta),3) = '000' Order by CodCta")
'RegCta.Requery
If RegCta.RecordCount <= 0 Then
    Exit Sub
End If

RegCta.MoveFirst
Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)
Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

While Not RegCta.Eof
            
'        While Not PgtoCta.Eof
'            Sql = "SELECT PAG.Cond, Sum(PAG.Valor) AS SomaDeValor, Sum(PAG.VlrPago) AS SomaDeVlrPago, Left([CodCta],6) AS Expr1 FROM Pag  WHERE PAG.DtPgto>=#" & Format(DtData1, "mm/dd/yyyy") & "# And PAG.DtPgto<=#" & Format(DtData2, "mm/dd/yyyy") & "# GROUP BY PAG.Cond, Left([CodCta],6)  HAVING PAG.Cond=True AND Left([CodCta],6)= " & Left(RegCta!CodCta, 6)
            Set PgtoCta = Banco.OpenRecordset(" SELECT PAG.Cond, Sum(PAG.Valor) AS SomaDeValor, Sum(PAG.VlrPago) AS SomaDeVlrPago, Left([CodCta],6) AS Expr1 FROM Pag  WHERE PAG.DtPgto>=#" & Format(DtData1, "mm/dd/yyyy") & "# And PAG.DtPgto<=#" & Format(DtData2, "mm/dd/yyyy") & "# GROUP BY PAG.Cond, Left([CodCta],6)  HAVING PAG.Cond=True AND Left([CodCta],6)= " & """" & Left(RegCta!CodCta, 6) & """")
            PgtoCta.Requery
            If PgtoCta.RecordCount > 0 Then
                PgtoCta.MoveFirst
                Set lstPesq = Lista.ListItems.Add(, , 1)
'                If Right(Trim(PgtoCta.Fields("CodCta")), 3) = "000" Then
                    lstPesq.SubItems(1) = Space(5) & RegCta!CodCta & " - " & RegCta!Descric
'                Else
'                    LstPesq.SubItems(1) = Space(15) & PgtoCta.Fields("CodCta") & " - " & PgtoCta!Descric
'                End If
                lstPesq.SubItems(2) = Format(Alinhar(CStr((PgtoCta!SomaDeVlrPago / VarPagto) * 100), 6, 2), "##,###,##0.00") & " %"
                lstPesq.SubItems(3) = Format(Alinhar(CStr(PgtoCta!SomaDeVlrPago), 12, 2), "##,###,##0.00")
            End If
'            PgtoCta.MoveNext
'        Wend
   RegCta.MoveNext
Wend

Set RegGru = Banco.OpenRecordset("Select * From Gru Order By CodGru")

RegGru.MoveFirst

If RegGru.RecordCount <= 0 Then
    Exit Sub
End If

Set RegCta = Banco.OpenRecordset("SELECT GRU.CodGru, GRU.Descrição_Grupo, Sum(PED.Desconto) AS SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Avg(ITP.ValorUnit) AS MédiaDeValor FROM (PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE (((PED.DtSaida)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DtSaida)<=# " & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3) AND ((DUP.Cond)=True)) GROUP BY GRU.CodGru, GRU.Descrição_Grupo ORDER BY GRU.CodGru, GRU.Descrição_Grupo")

'Set RegVda = Banco.OpenRecordset("SELECT GRU.CodGru, GRU.Descrição_Grupo, Sum(PED.Desconto) AS SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Avg(ITP.ValorUnit) AS MédiaDeValor FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.DtSaida)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And #( " & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1)) GROUP BY GRU.CodGru, GRU.Descrição_Grupo ORDER BY GRU.CodGru, GRU.Descrição_Grupo")

Set RegVda = Banco.OpenRecordset("SELECT GRU.CodGru, GRU.Descrição_Grupo, Sum(PED.Desconto) AS SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Avg(ITP.ValorUnit) AS MédiaDeValor FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.DtSaida)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DtSaida)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1)) GROUP BY GRU.CodGru, GRU.Descrição_Grupo ORDER BY GRU.CodGru, GRU.Descrição_Grupo")

'Sql = "SELECT GRU.CodGru, GRU.Descrição_Grupo, Sum(PED.Desconto) AS SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Avg(ITP.ValorUnit) AS MédiaDeValor FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.DtSaida)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DtSaida)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1)) GROUP BY GRU.CodGru, GRU.Descrição_Grupo ORDER BY GRU.CodGru, GRU.Descrição_Grupo"

VarPagto = 0

RegGru.MoveFirst

Dim Var1 As Currency
Dim Var2 As Currency
Dim Var3 As Currency
Dim Var4 As Currency
Dim Var5 As Currency
Dim Var6 As Currency

Var1 = 0
Var2 = 0
Var3 = 0
Var4 = 0
Var5 = 0
Var6 = 0

'While Not RegGru.Eof
    
'    Criterio = "Gru.CodGru = " & RegGru!codgru
'    RegCta.FindFirst Criterio
'    If Not RegCta.NoMatch Then
'        Var1 = (RegCta!SomaDeVlrVdaDia + RegCta!somaDesconto)
'    Else
'        Var1 = 0
'    End If
'    Criterio = "Gru.CodGru = " & RegGru!codgru
'    RegVda.FindFirst Criterio
'    If Not RegVda.NoMatch Then
'        Var2 = (RegVda!SomaDeVlrVdaDia + RegVda!somaDesconto)
'    Else
'        Var2 = 0
'    End If
'    VarPagto = VarPagto + (Var1 + Var2)
'
'    Var1 = 0
'    Var2 = 0
'
'    RegGru.MoveNext
'Wend

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(5) & "VENDA POR GRUPO CONSIDERANDO SÓ OS PAGOS"
Set lstPesq = Lista.ListItems.Add(, , 1)
lstPesq.SubItems(1) = Space(50)

RegGru.MoveFirst

While Not RegGru.Eof
            
    Criterio = "Gru.CodGru = " & RegGru!codgru
    RegCta.FindFirst Criterio
    If Not RegCta.NoMatch Then
        Var1 = (RegCta!SomaDeVlrVdaDia) ' + RegCta!SomaDesconto)
        Var4 = (RegCta!Custo)
    Else
        Var1 = 0
        Var4 = 0
    End If
    Criterio = "Gru.CodGru = " & RegGru!codgru
    RegVda.FindFirst Criterio
    If Not RegVda.NoMatch Then
        Var2 = (RegVda!SomaDeVlrVdaDia) ' + RegVda!SomaDesconto)
        Var5 = (RegVda!Custo)
    Else
        Var2 = 0
        Var5 = 0
    End If
    
    Var3 = (Var1 + Var2)
    Var6 = (Var4 + Var5)
            
    If Var3 <> 0 Then
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & RegGru!codgru & " - " & RegGru!descrição_Grupo
        If Var6 <> 0 Then
            lstPesq.SubItems(2) = Format(Alinhar(CStr((((Var3) / Var6) - 1) * 100), 12, 2), "##,###,##0.00") & " %"
        Else
            lstPesq.SubItems(2) = 0
        End If
        lstPesq.SubItems(3) = Format(Alinhar(CStr(Var3), 12, 2), "##,###,##0.00")
    End If
    
    Var1 = 0
    Var2 = 0
    Var3 = 0
    Var4 = 0
    Var5 = 0
    Var6 = 0
        
    RegGru.MoveNext

Wend

End Sub

Private Sub BtLaser_Click()
Dim rptGrid As FrmRelPedFut

If Lista.ListItems.Count > 0 Then
    Set rptGrid = New FrmRelPedFut
    Set rptGrid.Grid = Lista
    rptGrid.Relatorio.Escala = scmCentimetros
    rptGrid.Relatorio.MargemEsquerda = 1
    rptGrid.Relatorio.MargemDireita = 1
    rptGrid.TituloRelatorio.Caption = "Balancete     -  Periodo: " & DtData1 & " a " & DtData2
    rptGrid.Cab.Mostrar = False
'    rptGrid.Cabec1.Caption = LblNome
'    rptGrid.Cabec2.Caption = LblCidade
'    rptGrid.Cabec3.Caption = Format(Arq!CGC, "@@.@@@.@@@/@@@@-@@") & "      Inscr.Est: " & Arq!InscrEst
'    rptGrid.Cabec4.Caption = LblPedido & "        Data: " & LblData
'    rptGrid.Total.Caption = lblRestante
'    rptGrid.TotPed.Caption = LblValor
    rptGrid.SecTotal.Mostrar = False
    rptGrid.Executar
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Me.Left = 100 '(MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim X As Integer

MebValor = 0

Lista.ListItems.Clear

DtData1 = Date - 30
DtData2 = Date

MebCredito = 0
MebDebito = 0
MebAcrescimo = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub AtualizaLista()
Dim X As Integer
Dim Sql As String
Dim Criterio As String
Dim VarSaldo As Currency
Lista.ListItems.Clear
X = 0
Saldo = 0
Sql = "SELECT ITP.Produto, Sum(ITP.Qtd) AS SomaDeQtd FROM ITP WHERE ITP.Pedido = " & LstRel.List(X)
If LstRel.ListCount > 0 Then
    For X = 1 To LstRel.ListCount - 1
        Sql = Sql & " Or ITP.Pedido = " & LstRel.List(X)
    Next
    Sql = Sql & "   GROUP BY ITP.Produto"
    Set Fut = Banco.OpenRecordset(Sql)
End If

lblRestante = 0

Set Ped = Banco.OpenRecordset("SELECT ITP.Pedido, PRD.CodPrd, PRD.Descrição_Produto, Itp.ValorUnit, ITP.Qtd, Itp.CodItp FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido = " & TxtPedido & " ORDER BY Itp.CodItp ")
If Ped.RecordCount > 0 Then
    Ped.MoveFirst
    While Not Ped.Eof
        If LstRel.ListCount > 0 Then
            Criterio = "Produto = " & Ped!CodPrd
            Fut.FindFirst Criterio
            If Not Fut.NoMatch Then
                Saldo = Fut!SomaDeQtd
            Else
                Saldo = 0
            End If
        End If
'        Lista.AddItem psCompString(1, Left(Ped!codprd, 8), 8) & " " & psCompString(3, Left(Ped!Descrição_Produto, 30), 30) & " " & Alinhar(Ped!ValorUnit, 9, 3) & " " & Alinhar(Ped!Qtd, 9, 2) & " " & Alinhar(Ped!Qtd - Saldo, 9, 2)
        Set lstPesq = Lista.ListItems.Add(, , Ped("CodPrd"))

        lstPesq.SubItems(1) = psCompString(3, Left(Ped!Descrição_Produto, 40), 40)
        If Right(Format(Alinhar(Ped!Qtd, 9, 2), "##,###.00"), 2) = "00" Then
            lstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 9, 2), "##,###,##0")
        Else
            lstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 9, 2), "###,##0.0#")
        End If
        If Right(Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "##,###.00"), 2) = "00" Then
            lstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "##,###,##0")
        Else
            lstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "###,##0.0#")
        End If
        lstPesq.SubItems(4) = Alinhar(Ped!ValorUnit, 9, 3)
        lstPesq.SubItems(5) = Alinhar(Ped!ValorUnit * (Ped!Qtd - Saldo), 9, 2)
        lblRestante = lblRestante + (Ped!ValorUnit * (Ped!Qtd - Saldo))
        Ped.MoveNext
    Wend
End If

lblRestante = Format(lblRestante, "###,##0.00")

End Sub




