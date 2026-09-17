VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormImprimeProtocolos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Protocolos"
   ClientHeight    =   1260
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5160
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3630
      TabIndex        =   4
      Top             =   450
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36641
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   1860
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "FormImprimeProtocolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim Sql As String
Dim SqlCampos As String
Dim SqlWhere As String
Dim SqlOrdem As String

    SqlCampos = "SELECT PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat, ENT.CPF, PED.CodPed, PED.Total FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat, ENT.CPF, PED.CodPed, PED.Total "
    'SqlCampos = "SELECT PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade "
    SqlWhere = "Having PED.Operacao=0 AND PED.NroNt Is Not Null and PED.NroNt<>0 and PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DtFat)<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "#"
    SqlOrdem = "Order by Nome"
    
    Sql = SqlCampos & SqlWhere & SqlOrdem
    
    FrmRelPedProtocolo.rpCabTitulo.Caption = "Relatório de Protocolos"
    FrmRelPedProtocolo.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPedProtocolo.SubTitulo.Caption = "Período: " & DtData1 & " à " & DtData2
    
    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
    Set FrmRelPedProtocolo.Relatorio.Recordset = varPesquisa(1)
    FrmRelPedProtocolo.Config
    
End Sub

Private Sub Form_Load()

CentralizarForm Me, 1

DtData1 = Date
DtData2 = Date

End Sub

