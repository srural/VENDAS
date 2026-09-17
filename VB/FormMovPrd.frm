VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMovPrd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4920
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6615
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   19
      Top             =   4590
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   17
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Movimentação de Estoque"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   6315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loja"
      Height          =   1545
      Left            =   2205
      TabIndex        =   14
      Top             =   795
      Width           =   1860
      Begin VB.TextBox TxtOrigem 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   3
         Top             =   405
         Width           =   1545
      End
      Begin VB.TextBox TxtDestino 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   990
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Destino"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Destino"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   765
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimentação"
      Height          =   1410
      Left            =   450
      TabIndex        =   11
      Top             =   2520
      Width           =   2445
      Begin VB.OptionButton OptArq 
         Caption         =   "Movimentação Média"
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   20
         Top             =   630
         Width           =   2100
      End
      Begin VB.OptionButton OptArq 
         Caption         =   "Estoque Normal"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton OptArq 
         Caption         =   "Produtos da composição"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   990
         Width           =   2100
      End
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   0
      Left            =   4365
      TabIndex        =   5
      Top             =   1065
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52625409
      CurrentDate     =   37181
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Frame Frm1 
      Caption         =   "&Tipo"
      Height          =   1545
      Left            =   450
      TabIndex        =   8
      Top             =   795
      Width           =   1605
      Begin VB.OptionButton Opt1 
         Caption         =   "Tranferência"
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   2
         Top             =   1080
         Width           =   1275
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Entrada"
         Height          =   375
         Index           =   2
         Left            =   165
         TabIndex        =   0
         Top             =   225
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Saída"
         Height          =   375
         Index           =   3
         Left            =   165
         TabIndex        =   1
         Top             =   630
         Width           =   915
      End
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   1
      Left            =   4335
      TabIndex        =   6
      Top             =   1785
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52625409
      CurrentDate     =   37181
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
      Height          =   195
      Left            =   4335
      TabIndex        =   10
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   4335
      TabIndex        =   9
      Top             =   840
      Width           =   795
   End
End
Attribute VB_Name = "FormMovPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub BtImprimir_Click()
Dim X As Variant
Dim Oper As String
Dim SQL As String
Dim VarOrigem As String
Dim VarDestino As String
                       
If OptArq(0).value = True Then
            If Opt1(2).value = True Then
                    Oper = "Ped.Operacao = 0"
               ElseIf Opt1(3).value = True Then
                    Oper = "Ped.Operacao  = 1 Or Ped.Operacao = 2 Or Ped.Operacao = 3"
               Else
                    MsgBox "Operação não suportada", vbInformation
                    Exit Sub
            End If

            SQL = "SELECT PED.Operacao, Sum(ITP.Qtd) AS SomaDeQtd, First(COMP.Qtd) AS PrimeiroDeQtd, PRD.Descrição_Produto, Comp.PrdComp, Loja FROM PED INNER JOIN ((ITP INNER JOIN COMP ON ITP.Produto = Comp.Produto) INNER JOIN PRD ON Comp.PrdComp = PRD.CodPrd) ON PED.CodPed = ITP.Pedido "
            SQL = SQL + " WHERE PED.DataEmiss >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And PED.DataEmiss <= #" & Format(DtData(1), "mm/dd/yyyy") & "# GROUP BY PED.Operacao, PRD.Descrição_Produto, Comp.PrdComp, Loja HAVING " & Oper & " ORDER BY  PRD.Descrição_Produto "

' ************************
'    If Opt1(2).Value = True Then
'            Oper = "EST.Tipo = 1"
'       ElseIf Opt1(3).Value = True Then
'            Oper = "EST.Tipo = 2"
'       Else
'            Oper = "EST.Tipo = 3"
'    End If
'
'    SQL = " SELECT EST.Produto, Sum(EST.Qtd) AS SomaDeQtd, EST.Loja, EST.Data, EST.Hora, PRD.CodPrd, PRD.Descrição_Produto, [est].[qtd]*[comp].[qtd] AS Expr1 FROM (COMP INNER JOIN EST ON COMP.Produto = EST.Produto) INNER JOIN PRD ON COMP.PrdComp = PRD.CodPrd WHERE (((EST.Tipo)=2)) GROUP BY EST.Produto, EST.Loja, EST.LojaDestino, EST.Data, EST.Hora, PRD.CodPrd, PRD.Descrição_Produto, EST.Codigo, [est].[qtd]*[comp].[qtd] HAVING (((EST.Data)>=#2/6/2006# And (EST.Data)<=#2/6/2006#)) ORDER BY EST.Codigo"
'
    
'    SQL = "SELECT PED.Operacao, Sum(ITP.Qtd) AS SomaDeQtd1, Sum(COMP.Qtd) AS SomaDeQtd, PRD.Descrição_Produto, Comp.PrdComp, Loja FROM PED INNER JOIN ((ITP INNER JOIN COMP ON ITP.Produto = Comp.Produto) INNER JOIN PRD ON Comp.PrdComp = PRD.CodPrd) ON PED.CodPed = ITP.Pedido "
'    SQL = SQL + " WHERE PED.DataEmiss >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And PED.DataEmiss <= #" & Format(DtData(1), "mm/dd/yyyy") & "# GROUP BY PED.Operacao, PRD.Descrição_Produto, Comp.PrdComp, Loja HAVING " & Oper & " ORDER BY  PRD.Descrição_Produto "
        
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelMovPrd.Relatorio.Recordset = varPesquisa(1)
    FrmRelMovPrd.rpCabTitulo.Caption = "Relatório de Movimentação de Produtos"
    FrmRelMovPrd.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelMovPrd.Config

ElseIf OptArq(1).value = True Then
    If Opt1(2).value = True Then
            Oper = "EST.Tipo = 1"
            
       ElseIf Opt1(3).value = True Then
            Oper = "EST.Tipo = 2"
            
       Else
            Oper = "EST.Tipo = 3"
            
    End If
      
    If TxtOrigem = "" Then
        VarOrigem = ""
    Else
        VarOrigem = " AND Est.Loja = '" & TxtOrigem & "'"
    End If
    
    If TxtDestino = "" Then
        Vardstino = ""
    Else
        VarDestino = " AND Est.LojaDestino = '" & TxtDestino & "'"
    End If
        
    
    SQL = " SELECT EST.Tipo, EST.Produto,Embalagem, EST.Qtd, EST.Loja, Prd.Classe, EST.LojaDestino, EST.Data, EST.Hora, Prd.CodPrd, PRD.Descrição_Produto, Prd.Custo, EST.Codigo FROM EST INNER JOIN PRD ON EST.Produto = PRD.CodPrd "
    SQL = SQL + " WHERE " & Oper & " AND EST.Data >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And EST.Data <= #" & Format(DtData(1), "mm/dd/yyyy") & "# " & VarOrigem & VarDestino & " ORDER BY PRD.Descrição_Produto"
        
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelMov2Prd.Relatorio.Recordset = varPesquisa(1)
    FrmRelMov2Prd.rpCabTitulo.Caption = "Relatório de Movimentação Manual de Produtos"
    FrmRelMov2Prd.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelMov2Prd.Config
Else
    Dim VarPeriodo As Currency

    If Opt1(2).value = True Then
            Oper = "EST.Tipo = 1"
            
       ElseIf Opt1(3).value = True Then
            Oper = "EST.Tipo = 2"
            
       Else
            Oper = "EST.Tipo = 3"
            
    End If
      
    If TxtOrigem = "" Then
        VarOrigem = ""
    Else
        VarOrigem = " AND Est.Loja = '" & TxtOrigem & "'"
    End If
    
    If TxtDestino = "" Then
        Vardstino = ""
    Else
        VarDestino = " AND Est.LojaDestino = '" & TxtDestino & "'"
    End If
        
    
'    SQL = " SELECT EST.Tipo, EST.Produto,Embalagem, EST.Qtd, EST.Loja, Prd.Classe, EST.LojaDestino, EST.Data, EST.Hora, Prd.CodPrd, PRD.Descrição_Produto, Prd.Custo, EST.Codigo FROM EST INNER JOIN PRD ON EST.Produto = PRD.CodPrd "
'    SQL = SQL + " WHERE " & Oper & " AND EST.Data >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And EST.Data <= #" & Format(DtData(1), "mm/dd/yyyy") & "# " & VarOrigem & VarDestino & " ORDER BY PRD.Descrição_Produto"
        
     SQL = "SELECT EST.Tipo, EST.Produto,  Last(Prd.Custo) as UltimoCusto, PRD.Embalagem, Prd.Validade, Sum(EST.Qtd) AS SomaDeQtd, PRD.Classe, PRD.CodPrd, PRD.Descrição_Produto, Last(PRD.Estoque) AS UltimoDeEstoque FROM EST INNER JOIN PRD ON EST.Produto = PRD.CodPrd "
     SQL = SQL + "WHERE (((EST.Data)>=#" & Format(DtData(0), "mm/dd/yyyy") & "# And (EST.Data)<=#" & Format(DtData(1), "mm/dd/yyyy") & "#)) GROUP BY EST.Tipo, EST.Produto, PRD.Embalagem, Prd.Validade, Prd.Custo, PRD.Classe, PRD.CodPrd, PRD.Descrição_Produto HAVING " & Oper & " ORDER BY PRD.Descrição_Produto"
    If CDbl((CDbl(DtData(1)) - CDbl(DtData(0)))) > 32 Then
        Periodo = (CDbl((CDbl(DtData(1)) - CDbl(DtData(0)))) / 30)
    Else
        Periodo = 1
    End If
        
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelMov3Prd.Relatorio.Recordset = varPesquisa(1)
    FrmRelMov3Prd.rpCabTitulo.Caption = "Relatório de Movimentação Manual de Produtos"
    FrmRelMov3Prd.rpCabPeriodo.Caption = DtData(0) & " a " & DtData(1)
    FrmRelMov3Prd.VarMes = Periodo
    
    FrmRelMov3Prd.Config

End If
        

End Sub


Private Sub Form_Load()

Set Formulário = Me

DtData(0).value = Date
DtData(1).value = Date

End Sub

Private Sub Opt1_Click(Index As Integer)

If Opt1(2).value = True Or Opt1(3).value = True Then
    TxtOrigem.Enabled = True
Else
    TxtOrigem.Enabled = True
    TxtDestino.Enabled = True
End If




End Sub
