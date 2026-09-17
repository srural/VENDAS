VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormComissao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8400
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11940
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -720
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   20
      Top             =   8055
      Width           =   14370
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comissão"
      Height          =   1095
      Left            =   6885
      TabIndex        =   15
      Top             =   765
      Width           =   2625
      Begin MSMask.MaskEdBox MebVista 
         Height          =   330
         Left            =   945
         TabIndex        =   16
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPrazo 
         Height          =   330
         Left            =   945
         TabIndex        =   17
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Vista"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11940
      TabIndex        =   13
      Top             =   0
      Width           =   11940
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões"
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
         Left            =   225
         TabIndex        =   14
         Top             =   90
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado por"
      Height          =   1050
      Left            =   2115
      TabIndex        =   8
      Top             =   810
      Width           =   4575
      Begin VB.CheckBox ChkResumo 
         Caption         =   "Resumo"
         Height          =   240
         Left            =   2745
         TabIndex        =   22
         Top             =   630
         Width           =   1545
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Parceiro"
         Height          =   375
         Index           =   4
         Left            =   2745
         TabIndex        =   21
         Top             =   240
         Width           =   1425
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Vendedor"
         Height          =   375
         Index           =   3
         Left            =   1350
         TabIndex        =   12
         Top             =   540
         Width           =   1080
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Supervisor"
         Height          =   375
         Index           =   2
         Left            =   1350
         TabIndex        =   11
         Top             =   240
         Width           =   1080
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Produto"
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   540
         Width           =   885
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Grupo"
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   810
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5625
      Left            =   180
      TabIndex        =   5
      Top             =   2085
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9922
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   10936
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão"
      Height          =   1140
      Left            =   9765
      TabIndex        =   4
      Top             =   720
      Width           =   1830
      Begin VB.CommandButton BtLaser 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   225
         TabIndex        =   3
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Visualiza"
         Height          =   315
         Left            =   210
         TabIndex        =   2
         Top             =   270
         Width           =   1365
      End
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   870
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Format          =   33030145
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   1500
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Format          =   33030145
      CurrentDate     =   36641
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   660
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1290
      Width           =   720
   End
End
Attribute VB_Name = "FormComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VdaVenda As Recordset
Dim VdaPrazo As Recordset
Public lstPesq As ListItem

Private Sub BtImprimir_Click()
Dim SQL As String

Dim GruAnt As String
Dim VarTotGru As Double
Dim VarMedia As Double
Dim Cont As Double
Dim VarTotal As Currency

Dim VarQtd As Currency
Dim varValor As Currency
Dim VarComissao As Currency
Dim VarTotQtd As Currency
Dim VarTotValor As Currency
Dim VarTotComissao As Currency

VarQtd = 0
varValor = 0
VarComissao = 0
VarTotQtd = 0
VarTotValor = 0
VarTotComissao = 0



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



' TOTAL DAS VENDAS

DoEvents

Lista.ListItems.Clear

If OptTipo(0).value = True Then

    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo")
    SQL = "SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo"
    
    VdaVenda.Requery
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DAS VENDAS"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "GRUPO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!descrição_Grupo
        While Not VdaVenda.Eof
            If VdaVenda!descrição_Grupo = GruAnt Then
    
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!descrição_Grupo
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
            
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS A VISTA
    
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo")
    
    VdaVenda.Requery
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DE VENDAS A VISTA"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "GRUPO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!descrição_Grupo
        While Not VdaVenda.Eof
            If VdaVenda!descrição_Grupo = GruAnt Then
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!descrição_Grupo
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS RECEBIDAS
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO RECEBIDAS"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "GRUPO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((DUP.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, DUP.Cond HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True)) ORDER BY GRU.Descrição_Grupo")
    
    VdaVenda.Requery
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!descrição_Grupo
        While Not VdaVenda.Eof
            If VdaVenda!descrição_Grupo = GruAnt Then
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!descrição_Grupo
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")

ElseIf OptTipo(1).value = True Then

    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, PRD.Descrição_Produto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.CodPrd, PRD.Descrição_Produto HAVING (((First(PED.Total))<>0)) ORDER BY PRD.Descrição_Produto")
    
    SQL = "SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, PRD.Descrição_Produto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.CodPrd, PRD.Descrição_Produto HAVING (((First(PED.Total))<>0)) ORDER BY PRD.Descrição_Produto"
    
    VdaVenda.Requery
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DAS VENDAS"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "PRODUTO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!Descrição_Produto
        While Not VdaVenda.Eof
            If VdaVenda!Descrição_Produto = GruAnt Then
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!Descrição_Produto
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
            
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS A VISTA
    
'    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo")
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, PRD.Descrição_Produto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.CodPrd, PRD.Descrição_Produto HAVING (((First(PED.Total))<>0)) ORDER BY PRD.Descrição_Produto")
    
    
    VdaVenda.Requery
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DE VENDAS A VISTA"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "PRODUTO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!Descrição_Produto
        While Not VdaVenda.Eof
            If VdaVenda!Descrição_Produto = GruAnt Then
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!Descrição_Produto
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS RECEBIDAS
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO RECEBIDAS"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "PRODUTO"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
    lstPesq.SubItems(4) = "DESCONTO"
    
    
    'Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto, sum(DUP.VlrPago) AS PrimeiroDePago  FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.CodPed, PRD.CodPrd, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True) AND ((DUP.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) ORDER BY PRD.Descrição_Produto")
    
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto, Dup.Tipo_Dup FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.CodPed, PRD.CodPrd, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto, Dup.Tipo_Dup HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True) AND (Dup.Tipo_Dup <> 'SD')  AND ((DUP.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) ORDER BY PRD.Descrição_Produto")


    SQL = "SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto, Dup.Documento FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.CodPed, PRD.CodPrd, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto, Dup.Documento HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True) AND (Dup.Documento = SPACE(0))  AND ((DUP.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) ORDER BY PRD.Descrição_Produto"

    VdaVenda.Requery
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        GruAnt = VdaVenda!Descrição_Produto
        While Not VdaVenda.Eof
            If VdaVenda!Descrição_Produto = GruAnt Then
                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
                VarMedia = VarMedia + ((VdaVenda!PrimeiroDeDesconto / VdaVenda!Expr1) * 100)
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
                    Cont = Cont + 1
    '            End If
                VdaVenda.MoveNext
            Else
                If Cont <> 0 Then
                    VarMedia = VarMedia / Cont
                End If
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & GruAnt
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
                VarTotal = VarTotal + VarTotGru
                Cont = 0
                VarMedia = 0
                VarTotGru = 0
                GruAnt = VdaVenda!Descrição_Produto
            End If
        Wend
        If Cont <> 0 Then
            VarMedia = VarMedia / Cont
        End If
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & GruAnt
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        lstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(3) = "=========="
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50) & "TOTAL"
    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")

ElseIf OptTipo(2).value = True Then

'    Set VdaVenda = Banco.OpenRecordset("SELECT ENT.NomeVendedor, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.DataEmiss>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao =2 Or PED.Operacao=3) GROUP BY ENT.NomeVendedor")

    Set VdaVenda = Banco.OpenRecordset("SELECT ENT.NomeVendedor, Sum(ITP.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto, PRD.Grupo, GRU.Descrição_Grupo FROM GRU INNER JOIN (PRD INNER JOIN ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo WHERE (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY ENT.NomeVendedor, PRD.Grupo, GRU.Descrição_Grupo")

   ' Sql = " SELECT ENT.NomeVendedor, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto, PRD.Grupo, GRU.Descrição_Grupo FROM GRU INNER JOIN (PRD INNER JOIN ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo WHERE (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ENT.NomeVendedor, PRD.Grupo, GRU.Descrição_Grupo"
'    Sql = "SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo"
    
    VdaVenda.Requery
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = "TOTAL DAS VENDAS"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(5) & "SUPERVISOR"
    lstPesq.SubItems(2) = " "
    lstPesq.SubItems(3) = "VENDA"
'    LstPesq.SubItems(4) = "DESCONTO"
    
    Set lstPesq = Lista.ListItems.Add(, , 1)
    lstPesq.SubItems(1) = Space(50)
    
    VarTotGru = 0
    Cont = 0
    VarMedia = 0
    VarTotal = 0
    
    If VdaVenda.RecordCount > 0 Then
        VdaVenda.MoveFirst
        While Not VdaVenda.Eof
    
'                VarTotGru = VarTotGru + VdaVenda!SomaDeValor
'                varDesconto = varDesconto + VdaVenda!SomaDeDesconto
    '            If VdaVenda!PrimeiroDeDesconto <> 0 Then
    '                Cont = Cont + 1
    '            End If
    '            VdaVenda.MoveNext
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = Space(5) & VdaVenda!NomeVendedor & "   -   " & VdaVenda!descrição_Grupo
                lstPesq.SubItems(2) = " "
                lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDetotal - VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
'                LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
'                VarTotal = VarTotal + VarTotGru
'                Cont = 0
'                VarMedia = 0
'                VarTotGru = 0
'                GruAnt = VdaVenda!descrição_Grupo
'            End If
             VdaVenda.MoveNext
        Wend
'        If Cont <> 0 Then
'            VarMedia = VarMedia / Cont
'        End If
'        Set LstPesq = Lista.ListItems.Add(, , 1)
'        LstPesq.SubItems(1) = Space(5) & GruAnt
'        LstPesq.SubItems(2) = " "
'        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
'        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
'        VarTotal = VarTotal + VarTotGru
            
    End If
ElseIf OptTipo(3).value = True Then
    Dim VarVendAnt As String
    Dim VarVendAnt2 As String
    Dim VarVendAtual As String
    Dim VarTotVenda  As Currency
    If MebVista = "" Then

    '    Set VdaVenda = Banco.OpenRecordset("SELECT ENT.NomeVendedor, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.DataEmiss>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao =2 Or PED.Operacao=3) GROUP BY ENT.NomeVendedor")
    
        Set VdaVenda = Banco.OpenRecordset("SELECT ENT.Nome, Sum(ITP.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto, PRD.Grupo, GRU.Descrição_Grupo, DUP.Cond FROM ((PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON PED.Vendedor = ENT.CodEntidade) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE (((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY ENT.Nome, PRD.Grupo, GRU.Descrição_Grupo, DUP.Cond  HAVING DUP.Cond=True")
    
        
        VdaVenda.Requery
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DAS VENDAS RECEBIDAS NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & "VENDEDOR"
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = "VENDA"
    '    LstPesq.SubItems(4) = "DESCONTO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotVenda = 0
        VarTotal = 0
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
        VarVendAnt = psCompString(3, Left(VdaVenda!Nome, 20), 30)
        VarVendAtual = psCompString(3, Left(VdaVenda!Nome, 20), 30)
            While Not VdaVenda.Eof
                
                While Not VdaVenda.Eof And VarVendAnt = VarVendAtual
                    
                    VarTotVenda = VarTotVenda + (VdaVenda!SomaDetotal - VdaVenda!SomaDeDesconto)
                    If ChkResumo.value = 0 Then
                        Set lstPesq = Lista.ListItems.Add(, , 1)
                        lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaVenda!Nome, 20), 30) & "   -   " & VdaVenda!descrição_Grupo
                        lstPesq.SubItems(2) = " "
                        lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDetotal - VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
                   End If
                    
                 VdaVenda.MoveNext
                 If Not VdaVenda.Eof Then
                    VarVendAtual = psCompString(3, Left(VdaVenda!Nome, 20), 30)
                 End If
               Wend
                If ChkResumo.value = 1 Then
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = Space(3) & VarVendAnt
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotVenda), 12, 2), "##,###,##0.00")
                End If

               
               If Not VdaVenda.Eof Then
'                    VdaVenda.MovePrevious
                    VarVendAnt = psCompString(3, Left(VdaVenda!Nome, 20), 30)
'                    VarVendAnt2 = psCompString(3, Left(VdaVenda!Nome, 20), 30)
'                    VdaVenda.MoveNext
                    
'                    VarVendAtual = psCompString(3, Left(VdaVenda!Nome, 20), 30)
               End If
               If ChkResumo.value = O Then
                    lstPesq.SubItems(1) = Space(5) & "SUBTOTAL VENDEDOR " & VarVendAnt2
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotVenda), 12, 2), "##,###,##0.00")
               End If
               VarTotVenda = 0
            Wend
    '        If Cont <> 0 Then
    '            VarMedia = VarMedia / Cont
    '        End If
    '        Set LstPesq = Lista.ListItems.Add(, , 1)
    '        LstPesq.SubItems(1) = Space(5) & GruAnt
    '        LstPesq.SubItems(2) = " "
    '        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
    '        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
    '        VarTotal = VarTotal + VarTotGru
                
        End If
    ElseIf OptTipo(3).value = True Then
    
        Set VdaVenda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome ")
        
        SQL = "SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome "
        VdaVenda.Requery
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DAS VENDAS A VISTA  NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & "VENDEDOR"
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = "VENDA"
        lstPesq.SubItems(4) = "COMISSAO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotal = 0
        
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
            While Not VdaVenda.Eof
        
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaVenda!Nome, 40), 40)
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDeVlrVdaDia), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(4) = Format(Alinhar(CStr((VdaVenda!SomaDeVlrVdaDia) * (MebVista / 100)), 12, 2), "##,###,##0.00")
                    
'                    LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
                 VdaVenda.MoveNext
            Wend
        End If
        
    End If
    
    If MebPrazo <> "" Then
    
        Set VdaVenda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 2 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome ")
        
        SQL = "SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome "
        VdaVenda.Requery
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO  NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
    '    Set LstPesq = Lista.ListItems.Add(, , 1)
    '    LstPesq.SubItems(1) = Space(5) & "VENDEDOR"
    '    LstPesq.SubItems(2) = " "
    '    LstPesq.SubItems(3) = "VENDA"
    '    LstPesq.SubItems(4) = "COMISSAO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotal = 0
        
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
            While Not VdaVenda.Eof
        
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaVenda!Nome, 40), 40)
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDeVlrVdaDia), 12, 2), "##,###,##0.00")
'                    If MebPrazo <> "" Then
                        lstPesq.SubItems(4) = Format(Alinhar(CStr((VdaVenda!SomaDeVlrVdaDia) * (MebPrazo / 100)), 12, 2), "##,###,##0.00")
'                    End If
                    
    '                    LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
                 VdaVenda.MoveNext
            Wend
        End If
        
    Else
    
        Set VdaVenda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome ")
        
        SQL = "SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome "
        VdaVenda.Requery
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DAS VENDAS A VISTA  NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(5) & "NOME"
        lstPesq.SubItems(2) = " "
        lstPesq.SubItems(3) = "VALOR"
        lstPesq.SubItems(4) = "COMISSAO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotal = 0
        If MebVista = "" Then
            MebVista = 0
        End If
        
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
            While Not VdaVenda.Eof
                
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaVenda!Nome, 40), 40)
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDeVlrVdaDia), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(4) = Format(Alinhar(CStr((VdaVenda!SomaDeVlrVdaDia) * (MebVista / 100)), 12, 2), "##,###,##0.00")
                    
'                    LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
                 VdaVenda.MoveNext
            Wend
        End If
        
    End If
    
    If MebPrazo <> "" Then
    
        Set VdaVenda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 2 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome ")
        
        SQL = "SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido  Where DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# and (Operacao = 1 )  GROUP BY ENT.CodEntidade, ENT.Nome  Order by Nome "
        VdaVenda.Requery
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO  NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
    '    Set LstPesq = Lista.ListItems.Add(, , 1)
    '    LstPesq.SubItems(1) = Space(5) & "VENDEDOR"
    '    LstPesq.SubItems(2) = " "
    '    LstPesq.SubItems(3) = "VENDA"
    '    LstPesq.SubItems(4) = "COMISSAO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotal = 0
        
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
            While Not VdaVenda.Eof
        
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaVenda!Nome, 40), 40)
                    lstPesq.SubItems(2) = " "
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaVenda!SomaDeVlrVdaDia), 12, 2), "##,###,##0.00")
'                    If MebPrazo <> "" Then
                        lstPesq.SubItems(4) = Format(Alinhar(CStr((VdaVenda!SomaDeVlrVdaDia) * (MebPrazo / 100)), 12, 2), "##,###,##0.00")
'                    End If
                    
    '                    LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")
                 VdaVenda.MoveNext
            Wend
        End If
    
    End If
   MebVista = ""
Else


'        Set VdaVenda = Banco.OpenRecordset("SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, First(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade FROM (PED INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON ITP.Vendedor = ENT.CodEntidade WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade HAVING DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ENT.Nome")
        
        Dim VarCartao  As Currency
        Dim VarTaxa  As Currency
        Dim VarTotCartao As Currency
'        Dim VarTotValor  As Currency
'        Dim VarTotQtd As Currency
'        Dim VarTotComissao As Currency
        
        VarTotCartao = 0
        VarTotValor = 0
        VarTotQtd = 0
        VarTotComissao = 0
        
        If LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini") <> "" Then
            VarTaxa = LerINI("Parceiro", "TaxaCartao", App.Path & "\config.ini")
        Else
        '    TxtEvento = ""
        End If

        VarCartao = 0
        Set VdaVenda = Banco.OpenRecordset("SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, First(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade, PED.CondPgto FROM PED INNER JOIN ((COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON COMIS.CodPrd = PRD.CodPrd) INNER JOIN ENT ON (ENT.CodEntidade = COMIS.CodFun) AND (ITP.Vendedor = ENT.CodEntidade)) ON PED.CodPed = ITP.Pedido WHERE DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao=1 Or PED.Operacao=2 Or PED.Operacao=4) GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade, PED.CondPgto ORDER BY ENT.Nome")
                
        
'        SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, First(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade FROM PED INNER JOIN ((COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON COMIS.CodPrd = PRD.CodPrd) INNER JOIN ENT ON (ENT.CodEntidade = COMIS.CodFun) AND (ITP.Vendedor = ENT.CodEntidade)) ON PED.CodPed = ITP.Pedido WHERE DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# AND PED.Operacao=1 Or PED.Operacao=2 GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade ORDER BY ENT.Nome"
        SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, First(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade FROM PED INNER JOIN ((COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON COMIS.CodPrd = PRD.CodPrd) INNER JOIN ENT ON (ENT.CodEntidade = COMIS.CodFun) AND (ITP.Vendedor = ENT.CodEntidade)) ON PED.CodPed = ITP.Pedido WHERE DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao=1 Or PED.Operacao=2 Or PED.Operacao=4) GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade ORDER BY ENT.Nome"
        VdaVenda.Requery
        
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "TOTAL DE SERVIÇOS NO PERIODO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = "SERVIÇO"
        lstPesq.SubItems(2) = "QUANTIDADE"
        lstPesq.SubItems(3) = "VALOR"
        lstPesq.SubItems(4) = "COMISSAO"
        
        Set lstPesq = Lista.ListItems.Add(, , 1)
        lstPesq.SubItems(1) = Space(50)
        
        VarTotGru = 0
        Cont = 0
        VarMedia = 0
        VarTotal = 0
        Dim VarParceiroAnt As Integer
        
        If VdaVenda.RecordCount > 0 Then
            VdaVenda.MoveFirst
            While Not VdaVenda.Eof
                    VarParceiroAnt = VdaVenda!CodEntidade
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    lstPesq.SubItems(1) = VdaVenda!Nome
'                    Set VdaPrazo = Banco.OpenRecordset("SELECT PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, First(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade FROM (PED INNER JOIN (COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON (ITP.Vendedor = COMIS.CodFun) AND (COMIS.CodPrd = PRD.CodPrd)) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON ITP.Vendedor = ENT.CodEntidade WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.DataEmiss, PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade HAVING DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#  and CodEntidade = " & VdaVenda!CodEntidade & " ORDER BY ENT.Nome")
                    
                    Set VdaPrazo = Banco.OpenRecordset("SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, last(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade,PED.CondPgto FROM PED INNER JOIN ((COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON COMIS.CodPrd = PRD.CodPrd) INNER JOIN ENT ON (ENT.CodEntidade = COMIS.CodFun) AND (ITP.Vendedor = ENT.CodEntidade)) ON PED.CodPed = ITP.Pedido WHERE DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao=1 Or PED.Operacao=2 Or PED.Operacao=4) GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade,PED.CondPgto HAVING CodEntidade = " & VdaVenda!CodEntidade & " ORDER BY ENT.Nome")
                    VdaPrazo.Requery
                    SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Last(COMIS.Comissao) AS PrimeiroDeComissao, ENT.Nome, ENT.CodEntidade FROM PED INNER JOIN ((COMIS INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON COMIS.CodPrd = PRD.CodPrd) INNER JOIN ENT ON (ENT.CodEntidade = COMIS.CodFun) AND (ITP.Vendedor = ENT.CodEntidade)) ON PED.CodPed = ITP.Pedido WHERE DataEmiss >= # " & Format(DtData1, "mm/dd/yyyy") & "# AND DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "# AND (PED.Operacao=1 Or PED.Operacao=2 Or PED.Operacao=4) GROUP BY PRD.CodPrd, PRD.Descrição_Produto, ENT.Nome, ENT.CodEntidade HAVING CodEntidade = " & VdaVenda!CodEntidade & " ORDER BY ENT.Nome"
                    
                    While Not VdaPrazo.Eof
                            If ChkResumo = False Then
                                Set lstPesq = Lista.ListItems.Add(, , 1)
                                lstPesq.SubItems(1) = Space(3) & psCompString(3, Left(VdaPrazo!Descrição_Produto, 40), 40)
                                lstPesq.SubItems(2) = Format(Alinhar(CStr(VdaPrazo!SomaDeQtd), 12, 2), "##,###,##0.00")
                                lstPesq.SubItems(3) = Format(Alinhar(CStr(VdaPrazo!SomaDeValor), 12, 2), "##,###,##0.00")
                                lstPesq.SubItems(4) = Format(Alinhar(CStr((VdaPrazo!SomaDeValor) * (VdaPrazo!PrimeiroDeComissao / 100)), 12, 2), "##,###,##0.00")
                            End If
                            VarQtd = VarQtd + VdaPrazo!SomaDeQtd
                            VarTotQtd = VarTotQtd + VdaPrazo!SomaDeQtd
                            varValor = varValor + VdaPrazo!SomaDeValor
                            VarTotValor = VarTotValor + VdaPrazo!SomaDeValor
                            VarComissao = VarComissao + VdaPrazo!SomaDeValor * (VdaPrazo!PrimeiroDeComissao / 100)
                            VarTotComissao = VarTotComissao + VdaPrazo!SomaDeValor * (VdaPrazo!PrimeiroDeComissao / 100)
                            If VdaPrazo!CondPgto = "CARTAO DEBITO" Or VdaPrazo!CondPgto = "CARTAO CREDITO" Then
                                VarCartao = VarCartao + VdaPrazo!SomaDeValor
                                VarTotCartao = VarTotCartao + VdaPrazo!SomaDeValor
                            End If
                            

'                         VarParceiroAnt = VdaVenda!CodEntidade
                         VdaPrazo.MoveNext
                         VdaVenda.MoveNext
                    Wend
                    If ChkResumo = False Then
                        Set lstPesq = Lista.ListItems.Add(, , 1)
                        lstPesq.SubItems(1) = Space(50)
                        lstPesq.SubItems(2) = "--------"
                        lstPesq.SubItems(3) = "-----------"
                        lstPesq.SubItems(4) = "--------"
                        Set lstPesq = Lista.ListItems.Add(, , 1)
                        lstPesq.SubItems(1) = Space(50)
                    End If

                    lstPesq.SubItems(2) = Format(Alinhar(CStr(VarQtd), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(3) = Format(Alinhar(CStr(varValor), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(4) = Format(Alinhar(CStr(VarComissao), 12, 2), "##,###,##0.00")
                    
                    Set lstPesq = Lista.ListItems.Add(, , 1)
                    If ChkResumo = 1 Then
                        lstPesq.SubItems(1) = Space(50)
                    Else
                        lstPesq.SubItems(1) = Space(40) & "Cartão...."
                    End If
                    lstPesq.SubItems(2) = Format(Alinhar(CStr(VarTaxa), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(3) = Format(Alinhar(CStr((VarCartao * (VarTaxa / 100))), 12, 2), "##,###,##0.00")
                    lstPesq.SubItems(4) = Format(Alinhar(CStr(VarComissao - (VarCartao * (VarTaxa / 100))), 12, 2), "##,###,##0.00")
                    If ChkResumo = 1 Then
                        Set lstPesq = Lista.ListItems.Add(, , 1)
                        lstPesq.SubItems(1) = "-------------------------------------------------"
                        lstPesq.SubItems(2) = "--------"
                        lstPesq.SubItems(3) = "-----------"
                        lstPesq.SubItems(4) = "--------"
                    End If
                    
                    VarCartao = 0
                    VarQtd = 0
                    varValor = 0
                    VarComissao = 0
                    
                    
'                    LstPesq.SubItems(4) = Format(Alinhar(CStr(VdaVenda!SomaDeDesconto), 12, 2), "##,###,##0.00")

            Wend
            If ChkResumo = False Then
                Set lstPesq = Lista.ListItems.Add(, , 1)
                lstPesq.SubItems(1) = "-------------------------------------------------"
                lstPesq.SubItems(2) = "--------"
                lstPesq.SubItems(3) = "-----------"
                lstPesq.SubItems(4) = "--------"
            End If
            Set lstPesq = Lista.ListItems.Add(, , 1)
            lstPesq.SubItems(1) = Space(50)

            lstPesq.SubItems(2) = Format(Alinhar(CStr(VarTotQtd), 12, 2), "##,###,##0.00")
            lstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotValor), 12, 2), "##,###,##0.00")
            lstPesq.SubItems(4) = Format(Alinhar(CStr(VarTotComissao), 12, 2), "##,###,##0.00")
            
            Set lstPesq = Lista.ListItems.Add(, , 1)
            lstPesq.SubItems(1) = Space(40) & "Cartão...."
            lstPesq.SubItems(2) = Format(Alinhar(CStr(VarTaxa), 12, 2), "##,###,##0.00")
            lstPesq.SubItems(3) = Format(Alinhar(CStr((VarTotCartao * (VarTaxa / 100))), 12, 2), "##,###,##0.00")
            lstPesq.SubItems(4) = Format(Alinhar(CStr(VarTotComissao - (VarTotCartao * (VarTaxa / 100))), 12, 2), "##,###,##0.00")
            
        End If
End If


End Sub

Private Sub BtLaser_Click()
Dim rptGrid As FrmRelPedFut

If Lista.ListItems.Count > 0 Then
    Set rptGrid = New FrmRelPedFut
    Set rptGrid.Grid = Lista
    rptGrid.Relatorio.Escala = scmCentimetros
    rptGrid.Relatorio.MargemEsquerda = 1
    rptGrid.Relatorio.MargemDireita = 1
    rptGrid.TituloRelatorio.Caption = "Periodo: " & DtData1 & " a " & DtData2
    rptGrid.Cab.Mostrar = False
'    rptGrid.Cabec1.Caption = LblNome
'    rptGrid.Cabec2.Caption = LblCidade
'    rptGrid.Cabec3.Caption = Format(Arq!CGC, "@@.@@@.@@@/@@@@-@@") & "      Inscr.Est: " & Arq!InscrEst
'    rptGrid.Cabec4.Caption = LblPedido & "        Data: " & LblData
'    rptGrid.Total.Caption = lblRestante
'    rptGrid.TotPed.Caption = LblValor
    rptGrid.SecTotal.Mostrar = False
    rptGrid.rpCabSubTitulo.Caption = NomeEmpresa
    
    rptGrid.Executar
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Me.Left = 200 '(MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 ' (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim x As Integer

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
Dim x As Integer
Dim SQL As String
Dim Criterio As String
Dim VarSaldo As Currency
Lista.ListItems.Clear
x = 0
Saldo = 0
SQL = "SELECT ITP.Produto, Sum(ITP.Qtd) AS SomaDeQtd FROM ITP WHERE ITP.Pedido = " & LstRel.List(x)
If LstRel.ListCount > 0 Then
    For x = 1 To LstRel.ListCount - 1
        SQL = SQL & " Or ITP.Pedido = " & LstRel.List(x)
    Next
    SQL = SQL & "   GROUP BY ITP.Produto"
    Set Fut = Banco.OpenRecordset(SQL)
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




