VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmParcelas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNome 
      Height          =   330
      Left            =   1485
      TabIndex        =   35
      Top             =   3600
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -90
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   32
      Top             =   7605
      Width           =   11355
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7935
      TabIndex        =   30
      Top             =   0
      Width           =   7935
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
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
         TabIndex        =   31
         Top             =   45
         Width           =   5325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5265
      TabIndex        =   28
      Top             =   585
      Width           =   2625
      Begin VB.Label LblValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   90
         TabIndex        =   29
         Top             =   135
         Width           =   2370
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão do Carnê"
      Height          =   1095
      Left            =   4455
      TabIndex        =   22
      Top             =   6390
      Width           =   3075
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Formulário"
         Enabled         =   0   'False
         Height          =   720
         Index           =   0
         Left            =   2070
         Picture         =   "FrmParcelas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Bobina"
         Height          =   720
         Index           =   1
         Left            =   1125
         Picture         =   "FrmParcelas.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton BtImprimir 
         Caption         =   "&Página"
         Enabled         =   0   'False
         Height          =   720
         Index           =   2
         Left            =   180
         Picture         =   "FrmParcelas.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox TxtDoc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4140
      TabIndex        =   8
      Top             =   2385
      Width           =   1095
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmParcelas.frx":0C7E
      Left            =   2475
      List            =   "FrmParcelas.frx":0C97
      TabIndex        =   1
      Top             =   1170
      Width           =   1680
   End
   Begin VB.Data DatDados 
      Caption         =   "Parcelas"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Index           =   0
      Left            =   2565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * From Dup Order by Pedido"
      Top             =   4500
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSMask.MaskEdBox CodDoc 
      Height          =   270
      Left            =   360
      TabIndex        =   12
      Top             =   7155
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton OK 
      Height          =   405
      Left            =   5445
      Picture         =   "FrmParcelas.frx":0CB7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2295
      Width           =   450
   End
   Begin MSMask.MaskEdBox MebEntrada 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1755
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebParcelas 
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   1755
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid GridParc 
      Bindings        =   "FrmParcelas.frx":10F9
      Height          =   3375
      Left            =   360
      OleObjectBlob   =   "FrmParcelas.frx":1113
      TabIndex        =   10
      Top             =   2835
      Width           =   7230
   End
   Begin MSMask.MaskEdBox MebDia 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   2385
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebValorPcl 
      Height          =   315
      Left            =   2715
      TabIndex        =   7
      Top             =   2385
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebVlrPed 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtMes 
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Top             =   2385
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebIndices 
      Height          =   315
      Left            =   3195
      TabIndex        =   4
      Top             =   1755
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CodEnt 
      Height          =   270
      Left            =   360
      TabIndex        =   33
      Top             =   6615
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483644
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Entidade"
      Height          =   195
      Left            =   360
      TabIndex        =   34
      Top             =   6390
      Width           =   630
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Se deixar o dia e mês de pagamento sem preencher é assumido a data do dia"
      ForeColor       =   &H80000007&
      Height          =   675
      Left            =   4920
      TabIndex        =   27
      Top             =   1575
      Width           =   2595
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indices"
      Height          =   195
      Left            =   3195
      TabIndex        =   26
      Top             =   1530
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mês Inicio"
      Height          =   195
      Left            =   1665
      TabIndex        =   21
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento"
      Height          =   195
      Left            =   4140
      TabIndex        =   20
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   2475
      TabIndex        =   19
      Top             =   945
      Width           =   315
   End
   Begin VB.Label LblOk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1710
      TabIndex        =   18
      Top             =   6435
      Width           =   2535
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Pedido"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   810
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor das Parcelas"
      Height          =   195
      Left            =   2700
      TabIndex        =   15
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dia Pagamento"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Parcelas"
      Height          =   195
      Left            =   1935
      TabIndex        =   13
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   6930
      Width           =   825
   End
End
Attribute VB_Name = "FrmParcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Parcelas As Recordset
Dim Ind As Recordset

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub DatDados_Reposition(Index As Integer)
Dim Total As Recordset
If CodDoc <> "" Then
    Set Total = FrmFrente.Banco.OpenRecordset("SELECT DUP.Pedido, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CodDoc)
    If Total.RecordCount <> 0 Then
'        If CCur(MebVlrPed) = Round(Total("SomadeValor"), 2) Then
            BtImprimir(0).Enabled = False
            BtImprimir(1).Enabled = True
            BtImprimir(2).Enabled = True
'            LblOk.Caption = "Parcelas Corretas !"
'        Else
'            BtImprimir(0).Enabled = False
'            BtImprimir(1).Enabled = False
'            BtImprimir(2).Enabled = False
'            LblOk.Caption = "Parcelas Incorretas !"
'        End If
        LblValor.Caption = Format(Total("SomadeValor"), "###,##0.00;(###,##0.00)")
'        LblValor.Caption = Format(CCur(MebVlrPed) - Round(Total("SomadeValor")), "#,##0.00;(#,##0.00)")
    
    End If
End If
End Sub

Private Sub Form_Activate()

BtImprimir(1).Enabled = False
BtImprimir(0).Enabled = False
BtImprimir(2).Enabled = False

'CodDoc.Text = FrmTotal.LblTot
'MebDia = 10
MebEntrada = 0
TxtMes = Month(Date)
CmbTipo.ListIndex = 0
MebIndices = 0
DatDados_Reposition 0

    DatDados(0).RecordSource = "SELECT * From DUP Where Pedido = " & CodDoc & " Order by NroPcl"
    DatDados(0).Refresh

End Sub

Private Sub Form_Load()
For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\dados.mdb"
Next
'If FrmParcelas.Tag = "" Then
'    MebVlrPed.Text = Format(FrmTotal.MebTotal, "#,##0.00")
'Else
'    DatDados(0).RecordSource = "SELECT * From DUP Where Pedido = " & FrmTotal.LblTot & " Order by NroPcl"
'    DatDados(0).Refresh
'    MebVlrPed.Text = Format(FrmTotal.MebTotal, "#,##0.00")
'End If


Set Ind = FrmFrente.Banco.OpenRecordset("Select * From Ind")

End Sub

Private Sub BtImprimir_Click(Index As Integer)
Dim Linha02Col1 As String
Dim Linha03Col1 As String
Dim Linha04Col1 As String
Dim Linha05Col1 As String
Dim Linha06Col1 As String
Dim Linha07Col1 As String
Dim Linha08Col1 As String

Dim Linha02Col2 As String
Dim Linha03Col2 As String
Dim Linha04Col2 As String
Dim Linha05Col2 As String
Dim Linha06Col2 As String
Dim Linha07Col2 As String
Dim Linha08Col2 As String

Dim Flag As Boolean
Dim Configura As Recordset
Dim LinhaSimples As String
Dim LinhaDupla As String
Dim Corte As String

LinhaSimples = "--------------------------------------"
LinhaDupla = "======================================"
Corte = "......................................"

Flag = False
DatDados(0).Refresh
If MsgBox("Confirma Impressão do Carnê ? ", vbYesNo, App.Title) = vbYes Then
    Set Configura = FrmFrente.Banco.OpenRecordset("SELECT * FROM CFG")
    DatDados(0).Recordset.MoveFirst
    Select Case Index
        Case 0
                Iniciar (True)
                Do While Not DatDados(0).Recordset.EOF
                    
                        Linha02Col1 = DatDados(0).Recordset("DtVenc")
                        Linha03Col1 = DatDados(0).Recordset("Valor")
                        Linha04Col1 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col1 = Date
                        Linha06Col1 = Format(FrmFrente.LblCodigo, "0000")
                        Linha07Col1 = Configura.Fields("MensagemRodape")
                        Linha08Col1 = Configura.Fields("ObsParcelas")
                                         
                        Linha02Col2 = DatDados(0).Recordset("DtVenc")
                        Linha03Col2 = DatDados(0).Recordset("Valor")
                        Linha04Col2 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col2 = Date
                        Linha06Col2 = Format(FrmFrente.LblCodigo, "0000")
                        Linha07Col2 = Configura.Fields("MensagemRodape")
                        Linha08Col2 = Configura.Fields("ObsParcelas")
                                
                        ImprimeLinha 1, 0, 0, 0, False, 0
                        ImprimeLinha 0, 30, CmbTipo.Text, 5, False, 0
                        ImprimeLinha 0, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, CmbTipo.Text, 5, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        
                        ImprimeLinha 0, 0, FrmFrente.LblNome, 38, False, 0
                        ImprimeLinha 0, 0, "", 4, False, 0
                        ImprimeLinha 0, 0, FrmFrente.LblNome, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha02Col1, 19, False, 0
                        ImprimeLinha 0, 0, Linha03Col1, 17, False, 2
                        
                        ImprimeLinha 0, 6, Linha02Col2, 19, False, 0
                        ImprimeLinha 0, 0, Linha03Col2, 17, False, 2
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha04Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha04Col2, 38, False, 0
                
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha05Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha05Col2, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha06Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha06Col2, 38, False, 0
                        
                        ImprimeLinha 2, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha07Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha07Col2, 38, False, 0
                        ImprimeLinha 0, 0, Linha08Col1, 38, False, 0
                        ImprimeLinha 0, 4, Linha08Col2, 38, False, 0
                        
                        ImprimeLinha 4, 0, 0, 0, False, 0
                        
                        DatDados(0).Recordset.MoveNext
                Loop
                Fim (True)
                Unload Me
        Case 1
                Iniciar (True)
                Do While Not DatDados(0).Recordset.EOF
                    
                        Linha02Col1 = DatDados(0).Recordset("DtVenc")
                        Linha03Col1 = DatDados(0).Recordset("Valor")
                        Linha04Col1 = CDec(CodDoc) & "/" & DatDados(0).Recordset("NroPcl")
                        Linha05Col1 = Date
                        Linha06Col1 = Format(FrmFrente.LblCodigo, "00000")
                        Linha07Col1 = Configura.Fields("MensagemRodape")
                        Linha08Col1 = Configura.Fields("ObsParcelas")
                                         
                                
                        ImprimeLinha 1, 0, 0, 0, False, 0
                        ImprimeLinha 0, 0, LinhaDupla, 38, False, 0
                        ImprimeLinha 1, 0, "", 0, False, 0
                        ImprimeLinha 0, 0, Linha06Col1 & "-" & FrmFrente.LblNome, 32, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Tipo:", 6, False, 0
                        ImprimeLinha 0, 1, CmbTipo.Text, 5, False, 0
                        ImprimeLinha 0, 5, "Pedido:", 7, False, 0
                        ImprimeLinha 0, 1, Linha04Col1, 8, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Data:", 5, False, 0
                        ImprimeLinha 0, 1, Date, 10, False, 0
                        ImprimeLinha 0, 1, "Valor:", 6, False, 0
                        ImprimeLinha 0, 1, Linha03Col1, 10, False, 2
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Venc.", 5, False, 0
                        ImprimeLinha 0, 1, Linha02Col1, 10, False, 0
                        ImprimeLinha 0, 1, "Juros:", 6, False, 0
                        ImprimeLinha 0, 1, "", 2, False, 0
                        
                        ImprimeLinha 2, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, "Pago:", 5, False, 0
                        ImprimeLinha 0, 1, "", 10, False, 0
                        ImprimeLinha 0, 1, "Total:", 6, False, 0
                        ImprimeLinha 0, 1, "", 10, False, 0
                        
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, LinhaSimples, 38, False, 0
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, Linha07Col1, 38, False, 0
                        ImprimeLinha 1, 0, "", 35, False, 0
                        ImprimeLinha 0, 0, Linha08Col1, 38, False, 0
                        ImprimeLinha 2, 0, "", 35, False, 0
                                                
                        DatDados(0).Recordset.MoveNext
                Loop
                ImprimeLinha 2, 0, "", 35, False, 0
                ImprimeLinha 2, 0, "", 35, False, 0
                ImprimeLinha 2, 0, "", 35, False, 0
                Fim (True)
                Unload Me
        Case 2
        
            Set varPesquisa(1) = DatDados(0).Recordset.Clone
            varPesquisa(1).Requery
            varPesquisa(1).MoveFirst
            
            FrmRelCarne.rpTitulo.Caption = Configura!Cabecalho1
            FrmRelCarne.RpTitulo2.Caption = Configura!Cabecalho1
            FrmRelCarne.RpNome.Caption = Left(TxtNome, 40)
            FrmRelCarne.RpNome2.Caption = Left(TxtNome, 40)
            FrmRelCarne.RpMens.Caption = Configura!MensagemRodape
            FrmRelCarne.RpMens2.Caption = Configura!MensagemRodape
            FrmRelCarne.RpTotPcl.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
            FrmRelCarne.RpTotPcl2.Caption = varPesquisa(1).RecordCount ' Trim(MebParcelas)
                        
            Set FrmRelCarne.Relatorio.Recordset = varPesquisa(1)
            Me.Hide
            FrmRelCarne.Config
            Me.Show vbModal
   
    End Select

End If
End Sub

Private Sub MebEntrada_LostFocus()
If MebEntrada.Text = "" Then
    MebEntrada = 0
End If
End Sub

Private Sub MebIndices_Change()
If MebParcelas <> "" And IsNumeric(MebEntrada) And IsNumeric(MebIndices) Then
    If MebEntrada.Text <> 0 And IsNumeric(MebEntrada) Then
        If MebIndices = 0 Then
            MebValorPcl = (MebVlrPed - MebEntrada) / MebParcelas
        Else
            MebValorPcl = (MebVlrPed - MebEntrada) * Round(MebIndices, 4)
        End If
    Else
        If MebIndices = 0 Then
            MebValorPcl = MebVlrPed / MebParcelas
        Else
            MebValorPcl = MebVlrPed * Round(MebIndices, 4)
        End If
    End If
End If
End Sub

Private Sub MebParcelas_Change()
Dim Criterio As String
If MebParcelas <> "" Then
    If Ind.RecordCount <> 0 Then
        Criterio = "NroIndice = " & MebParcelas
        Ind.FindFirst Criterio
        If Not Ind.NoMatch Then
            MebIndices = Ind("Indice")
        End If
    Else
        MebIndices = 0
    End If
End If
End Sub

Private Sub MebParcelas_LostFocus()
MebParcelas_Change
End Sub

Private Sub MebValorPcl_LostFocus()
If MebValorPcl.Text = "" Then
    If MebParcelas.Text <> "" Then
        MebValorPcl.Text = (CCur(MebVlrPed) - CCur(MebEntrada)) / MebParcelas
    Else
        MebParcelas = ""
        MebParcelas.SetFocus
    End If
End If
End Sub

Private Sub OK_Click()
Dim sql As String
Dim Parc As Integer
Dim DtVenc As Date

If MebDia <> "" And Not IsNumeric(MebDia) Then
    If MebDia < 0 And MebDia > 32 Then
        MsgBox "Dia Inválido !", vbExclamation, App.Title
        MebDia.SetFocus
        Exit Sub
    End If
    MebDia.SetFocus
    Exit Sub
End If

If TxtMes <> "" And Not IsNumeric(TxtMes) Then
    If TxtMes < 0 And TxtMes > 13 Then
        MsgBox "Mes Inválido !", vbExclamation, App.Title
        TxtMes.SetFocus
        Exit Sub
    End If
    TxtMes.SetFocus
    Exit Sub
End If

If MebValorPcl = "" And Not IsNumeric(MebValorPcl) Then
    MebValorPcl.SetFocus
    Exit Sub
End If

sql = "Delete * From Dup Where Pedido = " & CodDoc
FrmFrente.Banco.Execute sql
DatDados(0).Refresh
GridParc.Refresh
Set Parcelas = FrmFrente.Banco.OpenRecordset("DUP")

If MebDia = "" Then
    DtVenc = Date
Else
    If TxtMes = "1" And Month(Date) = 12 Then
        DtVenc = Format(MebDia & "/" & CInt(TxtMes) & "/" & Year(Date) + 1, "dd/mm/yyyy")
    Else
        DtVenc = Format(MebDia & "/" & CInt(TxtMes) & "/" & Year(Date), "dd/mm/yyyy")
    End If
End If

If MebEntrada.Text <> "0" And IsNumeric(MebEntrada) Then
'    Parc = 1
'    Parcelas.AddNew
'    Parcelas("Pedido") = CodDoc
'    Parcelas("NroPcl") = Parc
'    Parcelas("Tipo_Dup") = CmbTipo.Text
'    Parcelas("Documento") = TxtDoc.Text
'    Parcelas("DtEmiss") = Date
'    Parcelas("Entidade") = FrmFrente.LblCodigo
'    Parcelas("DtVenc") = Date
'    Parcelas("Valor") = CCur(MebEntrada)
'    Parcelas.Update
    
    Parc = 1
    Parcelas.AddNew
    Parcelas("Pedido") = CodDoc
    Parcelas("NroPcl") = Parc
    Parcelas("Tipo_Dup") = CmbTipo.Text
    Parcelas("Documento") = TxtDoc.Text
    Parcelas("DtEmiss") = Date
    Parcelas("Entidade") = FrmFrente.LblCodigo  'CodEnt
    Parcelas("DtVenc") = Date
    Parcelas("Valor") = CCur(MebEntrada)
    Parcelas("Documento") = "ENTRADA"
    Parcelas("Cond") = True
    Parcelas.Update
    
    If CodDoc <> 0 Then
        sql = "Delete * From Rec Where CodDupli = " & CodDoc
        FrmFrente.Banco.Execute sql
    End If
    
    sql = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Maquina  ) VALUES ( "
    sql = sql & FrmFrente.LblCodigo & " , "
    sql = sql & "#" & Format(Date, "mm/dd/yyyy") & "# , "
    sql = sql & """" & Time & """" & " , "
    sql = sql & Numero(CStr(MebEntrada), 2) & " , "
    sql = sql & 0 & " , "
    sql = sql & 1 & " ,"
    sql = sql & CodDoc & " ,"
    sql = sql & """" & "FRENTE" & """" & ")"
    FrmFrente.Banco.Execute sql
        
Else
    Parc = 0
End If
For x = 1 To MebParcelas
    Parc = Parc + 1
    DtVenc = DtVenc + 30
    Parcelas.AddNew
    Parcelas("Pedido") = CodDoc
    Parcelas("NroPcl") = Parc
    Parcelas("Tipo_Dup") = CmbTipo.Text
    Parcelas("Documento") = TxtDoc.Text
    Parcelas("DtEmiss") = Date
    If IsNumeric(FrmFrente.LblCodigo) Then
        Parcelas("Entidade") = FrmFrente.LblCodigo
    End If
    If MebDia = "" Then
        Parcelas("DtVenc") = DtVenc
    Else
        If IsDate(Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy")) Then
            Parcelas("DtVenc") = Format(MebDia & "/" & Month(DtVenc) & "/" & Year(DtVenc), "dd/mm/yyyy")
        Else
            Parcelas("DtVenc") = Null
        End If
    End If
    Parcelas("Valor") = Round(CCur(MebValorPcl), 2)
    Parcelas.Update
Next

DatDados(0).RecordSource = "SELECT DUP.* FROM DUP WHERE (((DUP.Pedido)= " & CodDoc & ")) ORDER BY DUP.NroPcl"
DatDados(0).Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyPress = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub

