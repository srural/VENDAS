VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormEntregaFutura 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7980
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12195
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   1890
      TabIndex        =   26
      Top             =   6750
      Width           =   6135
      Begin VB.CommandButton BtGerar 
         Caption         =   "Gerar a Pedido"
         Height          =   330
         Left            =   4680
         TabIndex        =   32
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton BtCalcula 
         Caption         =   "Calcula"
         Height          =   330
         Left            =   3330
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DtData1 
         Height          =   330
         Left            =   135
         TabIndex        =   27
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49938433
         CurrentDate     =   36641
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   330
         Left            =   1665
         TabIndex        =   28
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49938433
         CurrentDate     =   36641
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   1665
         TabIndex        =   29
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   25
      Top             =   7605
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
      ScaleWidth      =   12210
      TabIndex        =   23
      Top             =   0
      Width           =   12210
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Entregas Futuras"
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
         TabIndex        =   24
         Top             =   45
         Width           =   5730
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Produtos à entregar"
      Height          =   315
      Left            =   90
      TabIndex        =   22
      Top             =   7005
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão"
      Height          =   780
      Left            =   8460
      TabIndex        =   18
      Top             =   6750
      Width           =   3510
      Begin VB.CommandButton BtLaser 
         Caption         =   " Laser/Jato"
         Height          =   315
         Left            =   1815
         TabIndex        =   20
         Top             =   315
         Width           =   1515
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "Matricial"
         Height          =   315
         Left            =   135
         TabIndex        =   19
         Top             =   315
         Width           =   1515
      End
   End
   Begin VB.ListBox LstRel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   135
      TabIndex        =   15
      Top             =   1935
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   1845
      TabIndex        =   4
      Top             =   720
      Width           =   10095
      Begin VB.Label LblEnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   7875
         TabIndex        =   33
         Top             =   180
         Width           =   45
      End
      Begin VB.Label lblRestante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5640
         TabIndex        =   17
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Restante:"
         Height          =   195
         Left            =   4440
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   435
         TabIndex        =   14
         Top             =   900
         Width           =   390
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         Height          =   195
         Left            =   2175
         TabIndex        =   13
         Top             =   900
         Width           =   810
      End
      Begin VB.Label LblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   12
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   11
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   180
         Width           =   465
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3075
         TabIndex        =   7
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label LblCidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   6
         Top             =   420
         Width           =   45
      End
      Begin VB.Label LblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   5
         Top             =   180
         Width           =   45
      End
   End
   Begin VB.TextBox TxtPedRel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   135
      TabIndex        =   1
      Tag             =   "N"
      Top             =   1575
      Width           =   1515
   End
   Begin VB.TextBox TxtPedido 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Tag             =   "N"
      Top             =   855
      Width           =   1470
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4635
      Left            =   1845
      TabIndex        =   21
      Top             =   2055
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8176
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CodPrd"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qtd"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Unitário"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Pedido Orginal"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Relacionados"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   1350
      Width           =   975
   End
End
Attribute VB_Name = "FormEntregaFutura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arq As Recordset
Dim Ped As Recordset
Dim Lst As Recordset
Dim Fut As Recordset
Dim Cli As Recordset
Dim LstTab As Recordset
Dim VarItens As Recordset
Dim pedidos As Recordset

Public VarOper As Byte
Public lstPesq As ListItem

Private Sub BtCalcula_Click()

Set LstTab = Banco.OpenRecordset("SELECT PED.Operacao, PED.DataEmiss, PED.CodPed FROM PED WHERE (((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=# " & Format(DtData2, "mm/dd/yyyy") & "#)) ORDER BY PED.CodPed")
LstTab.Requery

If LstTab.RecordCount > 0 Then
    LstTab.MoveFirst
    While Not LstTab.Eof
        TxtPedRel.SetFocus
        TxtPedRel = LstTab!CodPed
        TxtPedRel_KeyPress 13
'        SendKeys "{Enter}"
        LstTab.MoveNext
'        TxtPedRel = ""
    Wend
End If

End Sub

Private Sub BtGerar_Click()
Dim VarCodPed As String
            
If MsgBox("Deseja gerar o pedido ?", vbYesNo) = vbYes Then

'    (Operacao, Entidade, DataEmiss, Total, Desconto, Vendedor, Entregue, Transportadora, Comissao, IcmSubst, Ipi, Icms, ValorFrete,  SubTotal, TotalPrd ) "
    
'    SELECT " & "4" & " ," & Str(LblEnt) & "," & Date & ", " & Numero2(Str(lblRestante)) & ", 0, 0, False AS Expr1, 0, 0, 0, 0, 0, 0, " & Numero2(Str(lblRestante)) & ", " & Numero2(Str(lblRestante)) & " FROM PED "
 '0   Banco.Execute SQL
    
    
        Set pedidos = Banco.OpenRecordset("Select * From Ped Order by CodPed")
        pedidos.Requery
        
        pedidos.AddNew
        pedidos!Operacao = "4"
        pedidos!Entidade = LblEnt
        pedidos!DataEmiss = Date
        pedidos!Total = Numero2(lblRestante)
        pedidos!Entregue = False
        pedidos!SubTotal = Numero2(lblRestante)
        pedidos!TotalPrd = Numero2(lblRestante)
        pedidos.Update
        
        pedidos.Requery
        
        
        If pedidos.RecordCount > 0 Then
            pedidos.MoveLast
            VarCodPed = pedidos!CodPed
        End If
    
    Set VarItens = Banco.OpenRecordset("SELECT * FROM Itp")
    VarItens.Requery

            
    For X = 1 To Lista.ListItems.Count
'        While Not VarItp.Eof
            VarItens.AddNew
            VarItens!Pedido = VarCodPed
            VarItens!Produto = Lista.ListItems.Item(X).Text
            VarItens!Qtd = Lista.ListItems.Item(X).SubItems(3)
            VarItens!ValorUnit = Lista.ListItems.Item(X).SubItems(4)
            VarItens!Valor = Lista.ListItems.Item(X).SubItems(5)
            VarItens!Desconto = 0
            VarItens!CustoDia = 0
            VarItens!Loja = "LOJA"
            VarItens!Entregue = False
            VarItens!AliqIcms = 0
            VarItens!AliqIPI = 0
            VarItens!ValorIPI = 0
'            VarItens!Cst = " "
            VarItens!Comissao = 0
            VarItens!VlrVdaDia = Lista.ListItems.Item(X).SubItems(5)
            VarItens!VendaDia = Lista.ListItems.Item(X).SubItems(4)
            VarItens.Update
'            VarItp.MoveNext
 '       Wend
    Next
    MsgBox "Pedido Criado com Sucesso !", vbInformation
End If
End Sub

Private Sub BtImprimir_Click()
Dim Linha As Integer
    
Linha = 8
If Lista.ListItems.Count > 0 Then
    If MsgBox("Confirma Impressão ?", vbYesNo, App.Title) = vbYes Then
        Iniciar (False)
        
'        Print #1, "==============================================================================="
'        Print #1, Gde & Centralizar(Left(Empresa("Cabecalho1"), 40), 40) & Nor
'        Print #1, Centralizar(Left(Empresa("Cabecalho2"), 80), 80)
        Print #1, "==============================================================================="
        Print #1, "Data:"; lbldadta; Tab(17); "   Saldo do Pedido:  "; psCompString(1, TxtPedido, 6)
        Print #1, "-------------------------------------------------------------------------------"
        Print #1, LblNome
        Print #1, LblCidade
        Print #1, "-------------------------------------------------------------------------------"
        Print #1, "CODIGO"; Tab(8); "PRODUTO"; Tab(49); "  "; Tab(56); "UNIT"; Tab(66); "QTD"; Tab(75); "SALDO"
        Print #1, "-------------------------------------------------------------------------------"
    
        For X = 1 To Lista.ListItems.Count
            If Linha >= 55 Then
                Print #1, SaltoPag
                Linha = 1
            End If
             Print #1, Lista.ListItems.Item(X).Text; Tab(8);
             Print #1, Lista.ListItems.Item(X).SubItems(1); Tab(51);
             Print #1, Lista.ListItems.Item(X).SubItems(2); Tab(61);
             Print #1, Lista.ListItems.Item(X).SubItems(3); Tab(71);
             Print #1, Lista.ListItems.Item(X).SubItems(4)
            
            Linha = Linha + 1
            
            
        Next
        Print #1, "==============================================================================="
        
        Print #1, SaltoPag
    
        Fim (False)
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
    rptGrid.TituloRelatorio.Caption = "Entregas Futuras"
    rptGrid.Cabec1.Caption = LblNome
    rptGrid.Cabec2.Caption = LblCidade
    rptGrid.Cabec3.Caption = Format(Cli!CGC, "@@.@@@.@@@/@@@@-@@") & "      Inscr.Est: " & Cli!InscrEst
    rptGrid.Cabec4.Caption = LblPedido & "        Data: " & LblData
    rptGrid.Total.Caption = lblRestante
    rptGrid.TotPed.Caption = LblValor
    rptGrid.Executar
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Lista.ListItems.Clear

Set Ped = Banco.OpenRecordset("SELECT  PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, Sum(ITP_1.Qtd) AS SomaDeQtd1, Sum([ITP].[Qtd])-Sum([ITP_1].[Qtd]) FROM (FUT INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON FUT.Pedido = ITP.Pedido) INNER JOIN ITP AS ITP_1 ON FUT.PedFuturo = ITP_1.Pedido GROUP BY PRD.Descrição_Produto, ITP.Produto HAVING (((Sum([ITP].[Qtd])-Sum([ITP_1].[Qtd]))<>0)) ORDER BY PRD.Descrição_Produto")
Ped.Requery

'Sql = "SELECT  PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, Sum(ITP_1.Qtd) AS SomaDeQtd1, Sum([ITP].[Qtd])-Sum([ITP_1].[Qtd]) FROM (FUT INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON FUT.Pedido = ITP.Pedido) INNER JOIN ITP AS ITP_1 ON FUT.PedFuturo = ITP_1.Pedido GROUP BY PRD.Descrição_Produto, ITP.Produto HAVING (((Sum([ITP].[Qtd])-Sum([ITP_1].[Qtd]))<>0)) ORDER BY PRD.Descrição_Produto"

If Ped.RecordCount > 0 Then
    Ped.MoveFirst
    While Not Ped.Eof
        Set lstPesq = Lista.ListItems.Add(, , Ped("Produto"))

        lstPesq.SubItems(1) = psCompString(3, Left(Ped!Descrição_Produto, 40), 40)
        If Right(Format(Alinhar(Ped!SomaDeQtd, 10, 2), "##,###.00"), 2) = "00" Then
            lstPesq.SubItems(2) = Format(Alinhar(Ped!SomaDeQtd, 10, 2), "##,###,##0")
        Else
            lstPesq.SubItems(2) = Format(Alinhar(Ped!SomaDeQtd, 10, 2), "###,##0.0#")
        End If
'        If Right(Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "##,###.00"), 2) = "00" Then
'            LstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "##,###,##0")
'        Else
'            LstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "###,##0.0#")
'        End If
'        LstPesq.SubItems(4) = Alinhar(Ped!ValorUnit, 10, 3)
'        LstPesq.SubItems(5) = Alinhar(Ped!ValorUnit * (Ped!Qtd - Saldo), 12, 2)
'        lblRestante = lblRestante + (Ped!ValorUnit * (Ped!Qtd - Saldo))
        Ped.MoveNext
    Wend
End If



End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()

Me.Left = 200 'MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 '(MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim X As Integer

MebValor = 0

Lista.ListItems.Clear
LstRel.Clear

'Set Arq = Banco.OpenRecordset("SELECT PED.*, PED.CodPed, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.UF, ENT.Cidade, Ent.Cgc, Ent.InscrEst, Ent.Fone FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade ORDER BY PED.CodPed")

Set Cli = Banco.OpenRecordset("SELECT ENT.CodEntidade,  ENT.Nome, ENT.Endereco, ENT.Nro, ENT.UF, ENT.Cidade, Ent.Cgc, Ent.InscrEst, Ent.Fone FROM ENT  ORDER BY CodEntidade")

Set Arq = Banco.OpenRecordset("SELECT PED.*  FROM PED  ORDER BY PED.CodPed")

DtData1 = Date
DtData2 = Date

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub LstRel_DBLClick()
Dim Criterio As String
If MsgBox("Deseja excluir ?", vbYesNo, App.Title) = vbYes Then
    Criterio = "Pedido = " & TxtPedido & " and PedFuturo = " & LstRel.Text
    Lst.FindFirst Criterio
    If Not Lst.NoMatch Then
        Lst.Delete
        LstRel.RemoveItem (LstRel.ListIndex)
        AtualizaLista
    End If
End If
End Sub


Private Sub TxtPedRel_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If TxtPedRel = TxtPedido Then
    MsgBox "Pedido não pode ser igual", vbInformation, App.Title
    TxtPedRel = ""
    TxtPedRel.SetFocus
    Exit Sub
End If
If TxtPedRel.Text <> "" Then
    If KeyAscii = 13 Then
        Criterio = "Ped.CodPed = " & TxtPedRel
        Arq.FindFirst Criterio
        If Not Arq.NoMatch Then
            LstRel.AddItem psCompString(1, TxtPedRel, 6)
            Lst.AddNew
            Lst("Pedido") = TxtPedido
            Lst("PedFuturo") = TxtPedRel
            Lst("Data") = Date
            Lst.Update
            AtualizaLista
            TxtPedRel.Text = ""
            TxtPedRel.SetFocus
        Else
            MsgBox "Pedido não encontrado !", vbInformation, App.Path
        End If
    End If
End If

End Sub

Private Sub TxtPedido_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 And IsNumeric(TxtPedido) Then
    Criterio = "Ped.CodPed = " & TxtPedido
    Arq.FindFirst Criterio
    If Not Arq.NoMatch Then
        Criterio = "CodEntidade = " & Arq!Entidade
        Cli.FindFirst Criterio
        LblNome.Caption = Cli!Nome
        LblCidade.Caption = Trim(Cli!Endereco) & ", " & Cli!Nro & "       " & Trim(Cli!Cidade) & "-" & Cli.Fields("Uf") & "     Fone " & Format(Cli!Fone, "(@@)@@@@-@@@@")
        LblValor.Caption = Format(Arq!Total, "##,##0.00")
        LblPedido.Caption = Arq.Fields("CodPed")
        LblData.Caption = Arq!DataEmiss
        LblEnt = Cli!CodEntidade
        Set Lst = Banco.OpenRecordset("SELECT * FROM FUT WHERE Pedido = " & TxtPedido & " ORDER BY Pedido, PedFuturo")
'        Lista.Clear
        LstRel.Clear
        If Lst.RecordCount > 0 Then
            Lst.MoveFirst
            While Not Lst.Eof
                LstRel.AddItem psCompString(1, Lst("PedFuturo"), 6)
                Lst.MoveNext
            Wend
        End If
        AtualizaLista
        TxtPedRel.SetFocus
    Else
        MsgBox "Pedido não encontrado !", vbInformation, App.Path
        Lista.ListItems.Clear
        LstRel.Clear
    End If
Else
    Lista.ListItems.Clear
    LstRel.Clear
End If
End Sub


Private Sub AtualizaLista()
Dim X As Integer
Dim SQL As String
Dim Criterio As String
Dim VarSaldo As Currency
Lista.ListItems.Clear
X = 0
Saldo = 0
SQL = "SELECT ITP.Produto, Sum(ITP.Qtd) AS SomaDeQtd FROM ITP WHERE ITP.Pedido = " & LstRel.List(X)
If LstRel.ListCount > 0 Then
    For X = 1 To LstRel.ListCount - 1
        SQL = SQL & " Or ITP.Pedido = " & LstRel.List(X)
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
        If Right(Format(Alinhar(Ped!Qtd, 10, 2), "##,###.00"), 2) = "00" Then
            lstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 10, 2), "##,###,##0")
        Else
            lstPesq.SubItems(2) = Format(Alinhar(Ped!Qtd, 10, 2), "###,##0.0#")
        End If
        If Right(Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "##,###.00"), 2) = "00" Then
            lstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "##,###,##0")
        Else
            lstPesq.SubItems(3) = Format(Alinhar(Ped!Qtd - Saldo, 10, 2), "###,##0.0#")
        End If
        lstPesq.SubItems(4) = Alinhar(Ped!ValorUnit, 10, 3)
        lstPesq.SubItems(5) = Alinhar(Ped!ValorUnit * (Ped!Qtd - Saldo), 12, 2)
        lblRestante = lblRestante + (Ped!ValorUnit * (Ped!Qtd - Saldo))
        Ped.MoveNext
    Wend
End If

lblRestante = Format(lblRestante, "###,##0.00")

End Sub

