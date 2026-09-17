VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMovEnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2505
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5985
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   10
      Top             =   2205
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
      TabIndex        =   8
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Produtos por Entidade"
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
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   5865
      End
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   0
      Left            =   795
      TabIndex        =   4
      Top             =   885
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   37181
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Frame Frm1 
      Caption         =   "&Tipo"
      Height          =   585
      Left            =   2685
      TabIndex        =   0
      Top             =   765
      Width           =   2535
      Begin VB.OptionButton Opt1 
         Caption         =   "&Clientes"
         Height          =   375
         Index           =   2
         Left            =   165
         TabIndex        =   2
         Top             =   165
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Fornecedores"
         Height          =   375
         Index           =   3
         Left            =   1110
         TabIndex        =   1
         Top             =   165
         Width           =   1365
      End
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   1
      Left            =   795
      TabIndex        =   5
      Top             =   1515
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   37181
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   795
      TabIndex        =   7
      Top             =   1290
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   795
      TabIndex        =   6
      Top             =   660
      Width           =   795
   End
End
Attribute VB_Name = "FormMovEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub BtImprimir_Click()
Dim x As Variant
Dim Oper As String
Dim SQL As String
                       
'If OptArq(0).Value = True Then
'    If Opt1(2).Value = True Then
'            Oper = "Ped.Operacao = 0"
'       ElseIf Opt1(3).Value = True Then
'            Oper = "Ped.Operacao  = 1 Or Ped.Operacao = 2 Or Ped.Operacao = 3"
'       Else
'            MsgBox "Operação não suportada", vbInformation
'            Exit Sub
'    End If
'
'    Sql = "SELECT PED.Operacao, Sum(ITP.Qtd) AS SomaDeQtd1, Sum(COMP.Qtd) AS SomaDeQtd, PRD.Descrição_Produto, Comp.PrdComp, Loja FROM PED INNER JOIN ((ITP INNER JOIN COMP ON ITP.Produto = Comp.Produto) INNER JOIN PRD ON Comp.PrdComp = PRD.CodPrd) ON PED.CodPed = ITP.Pedido "
'    Sql = Sql + " WHERE PED.DataEmiss >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And PED.DataEmiss <= #" & Format(DtData(1), "mm/dd/yyyy") & "# GROUP BY PED.Operacao, PRD.Descrição_Produto, Comp.PrdComp, Loja HAVING " & Oper & " ORDER BY  PRD.Descrição_Produto "
'
'    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
'    Set FrmRelMovPrd.Relatorio.Recordset = varPesquisa(1)
'    FrmRelMovPrd.rpCabTitulo.Caption = "Relatório de Movimentação de Produtos"
'    FrmRelMovPrd.rpCabSubTitulo.Caption = NomeEmpresa
'    FrmRelMovPrd.Config

'Else
'    If Opt1(2).Value = True Then
'            Oper = "EST.Tipo = 1"
'       ElseIf Opt1(3).Value = True Then
'            Oper = "EST.Tipo = 2"
'       Else
'            Oper = "EST.Tipo = 3"
'    End If
'
'    Sql = " SELECT EST.Tipo, EST.Produto, EST.Qtd, EST.Loja, EST.LojaDestino, EST.Data, EST.Hora, Prd.CodPrd, PRD.Descrição_Produto, EST.Codigo FROM EST INNER JOIN PRD ON EST.Produto = PRD.CodPrd "
'    Sql = Sql + " WHERE " & Oper & " AND EST.Data >= #" & Format(DtData(0), "mm/dd/yyyy") & "# And EST.Data <= #" & Format(DtData(1), "mm/dd/yyyy") & "# ORDER BY EST.Codigo"
 '
 ''   Set varPesquisa(1) = Banco.OpenRecordset(Sql)
'    Set FrmRelMov2Prd.Relatorio.Recordset = varPesquisa(1)
'    FrmRelMov2Prd.rpCabTitulo.Caption = "Relatório de Movimentação Manual de Produtos"
'    FrmRelMov2Prd.rpCabSubTitulo.Caption = NomeEmpresa
'    FrmRelMov2Prd.Config

'End If
        

End Sub


Private Sub Form_Load()

Set Formulário = Me

DtData(0).value = Date
DtData(1).value = Date

End Sub
