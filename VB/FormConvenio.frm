VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormConvenio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7605
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11100
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   13
      Top             =   7290
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11130
      TabIndex        =   11
      Top             =   0
      Width           =   11130
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Convênio"
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
         TabIndex        =   12
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado por"
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   660
      Width           =   1695
      Begin VB.OptionButton OptTipo 
         Caption         =   "Produto"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Grupo"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4860
      Left            =   180
      TabIndex        =   5
      Top             =   1950
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   8573
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
         Name            =   "MS Sans Serif"
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
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão"
      Height          =   1095
      Left            =   9150
      TabIndex        =   4
      Top             =   705
      Width           =   1695
      Begin VB.CommandButton BtLaser 
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   165
         TabIndex        =   3
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Visualiza"
         Height          =   315
         Left            =   165
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   825
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   1410
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   615
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "FormConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VdaVenda As Recordset
Dim VdaPrazo As Recordset
Public LstPesq As ListItem

Private Sub BtImprimir_Click()
Dim SQL As String

Dim GruAnt As String
Dim VarTotGru As Double
Dim VarMedia As Double
Dim Cont As Double
Dim VarTotal As Currency

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
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DAS VENDAS"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "GRUPO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
            
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS A VISTA
    
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo")
    
    VdaVenda.Requery
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DE VENDAS A VISTA"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "GRUPO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS RECEBIDAS
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO RECEBIDAS"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "GRUPO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    
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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
Else

    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, PRD.Descrição_Produto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.CodPrd, PRD.Descrição_Produto HAVING (((First(PED.Total))<>0)) ORDER BY PRD.Descrição_Produto")
    
'    Sql = "SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo"
    
    VdaVenda.Requery
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DAS VENDAS"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "PRODUTO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
            
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS A VISTA
    
'    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1 FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.Grupo, GRU.Descrição_Grupo HAVING (((First(PED.Total))<>0)) ORDER BY GRU.Descrição_Grupo")
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, PRD.Descrição_Produto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 ) AND ((PED.DataEmiss)>=# " & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) GROUP BY PED.CodPed, PRD.CodPrd, PRD.Descrição_Produto HAVING (((First(PED.Total))<>0)) ORDER BY PRD.Descrição_Produto")
    
    
    VdaVenda.Requery
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DE VENDAS A VISTA"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "PRODUTO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")
    
    ' TOTAL DAS VENDAS RECEBIDAS
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = "TOTAL DAS VENDAS A PRAZO RECEBIDAS"
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(50)
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(5) & "PRODUTO"
    LstPesq.SubItems(2) = " "
    LstPesq.SubItems(3) = "VENDA"
    LstPesq.SubItems(4) = "DESCONTO"
    
    
    Set VdaVenda = Banco.OpenRecordset("SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.CodPed, PRD.CodPrd, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True) AND ((DUP.DtPgto)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (DUP.DtPgto)<=#" & Format(DtData2, "mm/dd/yyyy") & "#)) ORDER BY PRD.Descrição_Produto")

    SQL = "SELECT PED.CodPed, PRD.CodPrd, Sum(ITP.Valor) AS SomaDeValor, First(PED.Total) AS PrimeiroDeTotal, First(PED.Desconto) AS PrimeiroDeDesconto, First(([ped].[total]+[ped].[desconto])) AS Expr1, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto FROM (PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2)) GROUP BY PED.CodPed, PRD.CodPrd, DUP.Cond, PRD.Descrição_Produto, DUP.DtPgto HAVING (((First(PED.Total))<>0) AND ((DUP.Cond)=True) AND ((DUP.DtPgto)>#1/1/2009# And (DUP.DtPgto)<#1/30/2009#)) ORDER BY PRD.Descrição_Produto"

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
                Set LstPesq = Lista.ListItems.Add(, , 1)
                LstPesq.SubItems(1) = Space(5) & GruAnt
                LstPesq.SubItems(2) = " "
                LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
                LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
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
        Set LstPesq = Lista.ListItems.Add(, , 1)
        LstPesq.SubItems(1) = Space(5) & GruAnt
        LstPesq.SubItems(2) = " "
        LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotGru), 12, 2), "##,###,##0.00")
        LstPesq.SubItems(4) = Format(Alinhar(CStr(VarMedia), 12, 2), "##,###,##0.00") & "%"
        VarTotal = VarTotal + VarTotGru
        
    End If
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(3) = "=========="
    
    Set LstPesq = Lista.ListItems.Add(, , 1)
    LstPesq.SubItems(1) = Space(120) & "TOTAL"
    LstPesq.SubItems(3) = Format(Alinhar(CStr(VarTotal), 12, 2), "##,###,##0.00")



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
    rptGrid.Executar
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200  '(MDIPrincipal.Height - Me.Height) / 6

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
        Set LstPesq = Lista.ListItems.Add(, , Ped("CodPrd"))

        LstPesq.SubItems(1) = psCompString(3, Left(Ped!Descrição_Produto, 40), 40)
        If Right(Format(Alinhar(Ped!Qtd, 9, 2), "##,###.00"), 2) = "00" Then
            LstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 9, 2), "##,###,##0")
        Else
            LstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 9, 2), "###,##0.0#")
        End If
        If Right(Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "##,###.00"), 2) = "00" Then
            LstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "##,###,##0")
        Else
            LstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 9, 2), "###,##0.0#")
        End If
        LstPesq.SubItems(4) = Alinhar(Ped!ValorUnit, 9, 3)
        LstPesq.SubItems(5) = Alinhar(Ped!ValorUnit * (Ped!Qtd - Saldo), 9, 2)
        lblRestante = lblRestante + (Ped!ValorUnit * (Ped!Qtd - Saldo))
        Ped.MoveNext
    Wend
End If

lblRestante = Format(lblRestante, "###,##0.00")

End Sub




