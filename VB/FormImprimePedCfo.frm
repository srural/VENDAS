VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormImprimeCfo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6855
   Begin VB.CheckBox ChkSat 
      Caption         =   "Cupom/Sat"
      Height          =   330
      Left            =   4995
      TabIndex        =   10
      Top             =   1170
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   9
      Top             =   2790
      Width           =   11535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6855
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório Vendas por C.F.Op."
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
         TabIndex        =   8
         Top             =   45
         Width           =   6045
      End
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From CFO Order by Descricao"
      Top             =   1170
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   4515
      TabIndex        =   4
      Top             =   1845
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   315
      Left            =   630
      TabIndex        =   0
      Top             =   1890
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   87031809
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   315
      Left            =   2715
      TabIndex        =   2
      Top             =   1890
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   87031809
      CurrentDate     =   36641
   End
   Begin MSDBCtls.DBCombo DbcDados 
      Bindings        =   "FormImprimePedCfo.frx":0000
      Height          =   315
      Index           =   0
      Left            =   630
      TabIndex        =   5
      Top             =   1170
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483643
      ListField       =   "Descricao"
      BoundColumn     =   "Codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFO"
      Height          =   195
      Left            =   630
      TabIndex        =   6
      Top             =   945
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   2715
      TabIndex        =   3
      Top             =   1650
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   630
      TabIndex        =   1
      Top             =   1650
      Width           =   795
   End
End
Attribute VB_Name = "FormImprimeCfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()

    SQL = "SELECT Ped.Sat, PED.CodPed, PED.NroNt, PED.DtFat, ENT.Nome, PED.Total, CFO.Codigo, CFO.Descricao, Nota.TotOutras,  Sum(NOTA.ValorIcms) AS SomaDeValorIcms FROM (CFO INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON CFO.Codigo = PED.Cfo) INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed  GROUP BY Ped.Sat, PED.CodPed, PED.NroNt, PED.DtFat, ENT.Nome, PED.Total, CFO.Codigo, CFO.Descricao, Nota.TotOutras HAVING PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And PED.DtFat<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# "
    If ChkSat.value = 0 Then
        If IsNumeric(DbcDados(0).BoundText) Then
            SQL = SQL & " AND   PED.NroNt <> 0 AND CFO.Codigo='" & DbcDados(0).BoundText & "'"
        Else
            SQL = SQL & " AND   PED.NroNt <> 0  "
        End If
        SQL = SQL & " order by CFO.Codigo,  NroNt"
   
    Else
        SQL = "SELECT PED.Cfo, PED.Sat, PED.CodPed, PED.NroNt, PED.DtFat, ENT.Nome, PED.Total FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.Cfo, PED.Sat, PED.CodPed, PED.NroNt, PED.DtFat, ENT.Nome, PED.Total HAVING PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And PED.DtFat<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# "

        SQL = SQL & " AND   PED.NroNt <> 0  and (Ped.cfo = 'SAT' OR Ped.Cfo = 'Nfc-e')"
        
        SQL = SQL & " order by  NroNt"
        FrmRelPedCfo.ReportField5.Caption = "Sat"
    End If
      
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPedCfo.Relatorio.Recordset = varPesquisa(1)
    

    FrmRelPedCfo.rpCabTitulo.Caption = "Relatório de Pedidos por CFO"
    FrmRelPedCfo.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPedCfo.SubTitulo.Caption = "Período: " & DtData1 & " à " & DtData2
    
    
    FrmRelPedCfo.Config
    
End Sub

Private Sub Form_Load()

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

CentralizarForm Me, 1

DtData1 = Date
DtData2 = Date

End Sub

