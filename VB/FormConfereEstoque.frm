VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormConfereEstoque 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8205
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13650
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   15
      Top             =   7830
      Width           =   14370
   End
   Begin MSMask.MaskEdBox MebTotal 
      Height          =   420
      Left            =   11250
      TabIndex        =   14
      Top             =   1260
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13650
      TabIndex        =   11
      Top             =   0
      Width           =   13650
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Conferência de Estoque"
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
   Begin VB.CommandButton CmdCorrigir 
      Caption         =   "&Corrigir Estoque"
      Enabled         =   0   'False
      Height          =   540
      Left            =   9045
      TabIndex        =   10
      Top             =   1035
      Width           =   1110
   End
   Begin VB.CheckBox ChkTipo 
      Caption         =   "Sem Historico"
      Height          =   345
      Left            =   5850
      TabIndex        =   9
      Top             =   1755
      Width           =   1605
   End
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2250
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT HISTEST.DTEstoque FROM HISTEST GROUP BY HISTEST.DTEstoque ORDER BY HISTEST.DTEstoque desc"
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBCtls.DBList DBList 
      Bindings        =   "FormConfereEstoque.frx":0000
      DataSource      =   "Data"
      Height          =   1035
      Left            =   2220
      TabIndex        =   7
      Top             =   1020
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   1826
      _Version        =   393216
      Appearance      =   0
      ListField       =   "dtestoque"
   End
   Begin VB.CommandButton CmbExibir 
      Caption         =   "&Verificar"
      Height          =   585
      Left            =   5895
      TabIndex        =   6
      Top             =   990
      Width           =   1110
   End
   Begin VB.CommandButton BtLaser 
      Caption         =   "&Imprimir"
      Height          =   540
      Left            =   7785
      TabIndex        =   5
      Top             =   1035
      Width           =   1110
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5145
      Left            =   150
      TabIndex        =   0
      Top             =   2220
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9075
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CodPrd"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Antigo"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "+ Vda"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "- Vda"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "+ Est"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "- Est"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Atual"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Loja"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   1035
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   87031809
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   1695
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   87031809
      CurrentDate     =   36641
   End
   Begin VB.Label LblContador 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   450
      TabIndex        =   16
      Top             =   7470
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total do Estoque"
      Height          =   195
      Left            =   11250
      TabIndex        =   13
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico do Estoque"
      Height          =   195
      Left            =   2220
      TabIndex        =   8
      Top             =   810
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   810
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1470
      Width           =   720
   End
End
Attribute VB_Name = "FormConfereEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Est As Recordset
Dim Ped As Recordset
Dim Lst As Recordset
Dim Prd As Recordset
Dim Hist As Recordset
Dim Loja As Recordset
Dim ContPed As Recordset
Dim ContImp As Boolean

Public VarOper As Byte
Public lstPesq As ListItem

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbExibir_Click()
Dim Criterio As String
Dim VarCodigo As Double
Dim varNome As String
Dim varEstAnt As Currency
Dim VarSomaVda As Currency
Dim VarTiraVda As Currency
Dim VarSomaEst As Currency
Dim VarTiraEst As Currency
Dim VarSaldo As Currency
Dim VarAtual As Currency
Dim SQL As Recordset
Dim PedIni As Double
Dim PedFim As Double
Dim VarSql As String
Dim VarSaldoQtd As Double
Dim ContPrd As Double
Dim Contador As Double
Dim VarLoja As Currency
MebTotal = 0
ContPrd = 0
Contador = 0

Dim sql2 As String

'On Error Resume Next

CmbExibir.Enabled = False
BtLaser.Enabled = False
'sql2 = "SELECT codped, DataEmiss  FROM PED WHERE DataEmiss >=# " & Format(DtData1, "mm/dd/yyyy") & " # And DataEmiss <=#" & Format(DtData2, "mm/dd/yyyy") & "#  ORDER BY CodPed"
'Set ContPed = Banco.OpenRecordset("SELECT CodPed, DataEmiss  FROM PED WHERE DataEmiss >=#" & Format(DtData1, "mm/dd/yyyy") & "# And DataEmiss <=#" & Format(DtData2, "mm/dd/yyyy") & "#  ORDER BY CodPed")
'ContPed.Requery
'If ContPed.RecordCount > 0 Then
'    ContPed.MoveFirst
'    PedIni = ContPed!CodPed
'    ContPed.MoveLast
'    PedIni = ContPed!CodPed
'End If

'Set Ped = Banco.OpenRecordset("SELECT PED.DataEmiss, ITP.Produto, ITP.Qtd, PED.Operacao FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE PED.codped >= " & PedIni & "  And PED.CodPed <=" & PedFim & "  ORDER BY ITP.Produto")
Set Ped = Banco.OpenRecordset("SELECT PED.DataEmiss, ITP.Produto, ITP.Qtd, PED.Operacao FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE PED.DataEmiss >=# " & Format(DtData1, "mm/dd/yyyy") & " # And PED.DataEmiss <=#" & Format(DtData2, "mm/dd/yyyy") & "#  ORDER BY ITP.Produto")

'Sql = "SELECT PED.DataEmiss, ITP.Produto, ITP.Qtd, PED.Operacao FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE PED.DataEmiss >=# " & Format(DtData1, "mm/dd/yyyy") & " # And PED.DataEmiss <=#" & Format(DtData2, "mm/dd/yyyy") & "#  ORDER BY ITP.Produto"
Ped.Requery
If Ped.RecordCount > 0 Then
    Ped.MoveFirst
End If

Set Loja = Banco.OpenRecordset("SELECT * FROM Loj ORDER BY  Produto ")
Loja.Requery
Loja.MoveFirst


Set Prd = Banco.OpenRecordset("SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem,Prd.Custo,  PRD.Estoque, PRD.Grupo FROM Prd WHERE Ativo = true ORDER BY  PRD.Descrição_Produto ")
Prd.Requery
Prd.MoveFirst
Set Est = Banco.OpenRecordset("SELECT EST.Tipo, EST.Produto, Sum(EST.Qtd) AS SomaDeQtd FROM EST WHERE  EST.Data >= #" & Format(DtData1, "mm/dd/yyyy") & "# And EST.Data <= #" & Format(DtData2, "mm/dd/yyyy") & "#  GROUP BY EST.Tipo, EST.Produto Having Est.Tipo <> 3 ORDER BY EST.Produto")
Est.Requery
If Est.RecordCount > 0 Then
    Est.MoveFirst
End If
If ChkTipo.value = 0 Then
'    VarSql = "SELECT DTEstoque, CodPrd, QtdEst FROM HISTEST GROUP BY HISTEST.DTEstoque, HISTEST.CodPrd, HISTEST.QtdEst HAVING (((HISTEST.DTEstoque)=# " & Format(DBList, "mm/dd/yyyy") & " #)) ORDER BY HISTEST.CodPrd"
    
    Set Hist = BcoEst.OpenRecordset("SELECT DTEstoque, CodPrd, QtdEst FROM HISTEST GROUP BY HISTEST.DTEstoque, HISTEST.CodPrd, HISTEST.QtdEst HAVING (((HISTEST.DTEstoque)=# " & Format(DBList, "mm/dd/yyyy") & " #)) ORDER BY HISTEST.CodPrd")
    
    Hist.Requery
    Hist.MoveFirst
End If
Prd.MoveFirst

Lista.ListItems.Clear

While Not Prd.Eof
    varEstAnt = 0
    VarSomaVda = 0
    VarTiraVda = 0
    VarSomaEst = 0
    VarTiraEst = 0
    VarSaldo = 0
    VarAtual = 0
    DoEvents
    ContPrd = ContPrd + 1
    
    VarCodigo = Prd!CodPrd
    varNome = Prd!Descrição_Produto
    Criterio = "Produto =  " & VarCodigo
    Ped.FindFirst Criterio

'    If Ped.RecordCount > 0 Then
'        Ped.MoveFirst
            While Not Ped.Eof
                 If Ped!Produto = VarCodigo Then
        
                    Select Case Ped!Operacao
                           Case 0
                                 VarSomaVda = VarSomaVda + Ped!Qtd
                           Case 1
                                 VarTiraVda = VarTiraVda + Ped!Qtd
                           Case 2
                                 VarTiraVda = VarTiraVda + Ped!Qtd
                           Case 3
                           Case 4
                           Case 5
                           Case 6
                                 VarTiraVda = VarTiraVda + Ped!Qtd
                           Case 7
                           Case 8
                                 VarTiraVda = VarTiraVda + Ped!Qtd
                           Case 9
                                 VarSomaVda = VarSomaVda + Ped!Qtd
                           Case 10
                                 VarSomaVda = VarSomaVda + Ped!Qtd
                           Case 11
                                 VarTiraVda = VarTiraVda + Ped!Qtd
                           Case 12
                           Case 13
                           Case 14
                           Case 15
                                 VarSomaVda = VarSomaVda + Ped!Qtd
                                 
                    End Select
                Else
                    Ped.MoveLast
                End If
                
                Ped.MoveNext
            Wend
'End If
'            If VarCodigo = 953 Then
'                VarCodigo = VarCodigo
'            End If

            Criterio = "Produto =  " & VarCodigo & " AND Tipo = 1 "
            Est.FindFirst Criterio
            If Not Est.NoMatch Then
                VarSomaEst = VarSomaEst + Est.Fields("SomaDeQtd")
            End If
            
            Criterio = "Produto =  " & VarCodigo & " AND Tipo = 2 "
            Est.FindFirst Criterio
            If Not Est.NoMatch Then
                VarTiraEst = VarTiraEst + Est!SomaDeQtd
            End If
            
            VarAtual = Prd!estoque
            
            If ChkTipo.value = 0 Then
                Criterio = "CodPrd =  " & VarCodigo & " "
                Hist.FindFirst Criterio
                If Not Hist.NoMatch Then
                    varEstAnt = Hist!QtdEst
                End If
            Else
                varEstAnt = 0
            End If
'            VarSaldo = (varEstAnt + VarSomaVda + VarSomaEst) - (VarTiraVda + VarTiraEst)
            
            Criterio = "Produto =  " & VarCodigo & " "
            Loja.FindFirst Criterio
            If Not Loja.NoMatch Then
                VarLoja = Loja!estoque
            End If
                                    
            VarSaldoQtd = (CDec(varEstAnt) + CDec(VarSomaVda) + CDec(VarSomaEst)) - (CDec(VarTiraVda) + CDec(VarTiraEst))
            
            If Prd!estoque - VarSaldoQtd <> 0 Then
                    Contador = Contador + 1
'                If (CDec(VarEstAnt) + CDec(VarSomaVda) + CDec(VarSomaEst)) - (CDec(VarTiraVda) + CDec(VarTiraEst)) <> 0 And (CDec(VarEstAnt) + CDec(VarSomaVda) + CDec(VarSomaEst)) - (CDec(VarTiraVda) + CDec(VarTiraEst)) <> VarAtual Then
                    Set lstPesq = Lista.ListItems.Add(, , VarCodigo)
                
                    lstPesq.SubItems(1) = psCompString(3, Left(varNome, 40), 40)
                    lstPesq.SubItems(2) = CDec(varEstAnt)
                    lstPesq.SubItems(3) = CDec(VarSomaVda)
                    lstPesq.SubItems(4) = CDec(VarTiraVda)
                    lstPesq.SubItems(5) = CDec(VarSomaEst)
                    lstPesq.SubItems(6) = CDec(VarTiraEst)
                    lstPesq.SubItems(7) = (CDec(varEstAnt) + CDec(VarSomaVda) + CDec(VarSomaEst)) - (CDec(VarTiraVda) + CDec(VarTiraEst))
'                    LstPesq.SubItems(8) = CDec(VarAtual)
                    lstPesq.SubItems(8) = Prd!estoque  ' Format(Prd!Custo, "###,##0.00")
                    lstPesq.SubItems(9) = VarLoja   ' Prd!Estoque - VarSaldoQtd ' Format((Prd!Custo * VarSaldoQtd), "###,##0.00")
                    MebTotal = MebTotal + (Prd!Custo * VarSaldoQtd)
                    
                    If ContImp = True Then
'                            VarSql = " INSERT INTO HISTEST ( DTEstoque, CodPrd, QtdEst ) SELECT date() AS Expr1, " & VarCodigo & " AS Expr2, " & Numero(Str(VarSaldoQtd), 2) & " AS Expr3"
'                            Banco.Execute VarSql
                    End If
'                End If
            End If
    Prd.MoveNext
Wend
LblContador = Contador & "/" & ContPrd
CmbExibir.Enabled = True
BtLaser.Enabled = True
ContImp = False


End Sub


Private Sub BtLaser_Click()
Dim rptGrid As FrmRelConfEst

If Lista.ListItems.Count > 0 Then
    Set rptGrid = New FrmRelConfEst
    Set rptGrid.Grid = Lista
    rptGrid.Relatorio.Escala = scmCentimetros
    rptGrid.Relatorio.MargemEsquerda = 1
    rptGrid.Relatorio.MargemDireita = 1
    rptGrid.TituloRelatorio.Caption = "Conferência de Estoque"
    rptGrid.Cabec1.Caption = "Periodo:" & DtData1 & "  até  " & DtData2
    rptGrid.Executar
End If

End Sub


Private Sub CmdCorrigir_Click()
ContImp = True
CmbExibir_Click
End Sub

Private Sub Form_Activate()

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 '(MDIPrincipal.Height - Me.Height) / 6

DtData1 = Date
DtData2 = Date

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim x As Integer

Data.DatabaseName = Caminho & "\Estoque.mdb"

Lista.ListItems.Clear
ContImp = False
'Set Arq = Banco.OpenRecordset("SELECT PED.*, PED.CodPed, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.UF, ENT.Cidade, Ent.Cgc, Ent.InscrEst, Ent.Fone FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade ORDER BY PED.CodPed")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
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
        LblNome.Caption = Arq!Nome
        LblCidade.Caption = Trim(Arq!Endereco) & ", " & Arq!Nro & "       " & Trim(Arq!Cidade) & "-" & Arq.Fields("Ent.Uf") & "     Fone " & Format(Arq!Fone, "(@@)@@@@-@@@@")
        LblValor.Caption = Format(Arq!Total, "##,##0.00")
        LblPedido.Caption = Arq.Fields("Ped.CodPed")
        LblData.Caption = Arq!DataEmiss
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




