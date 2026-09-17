VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormRelProd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6360
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -315
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   8
      Top             =   2430
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
      TabIndex        =   6
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios para Produção"
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
         Left            =   135
         TabIndex        =   7
         Top             =   45
         Width           =   5685
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   3105
      TabIndex        =   5
      Top             =   1305
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   3105
      TabIndex        =   4
      Top             =   1755
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   1380
      TabIndex        =   0
      Top             =   1005
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   53280769
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   1380
      TabIndex        =   2
      Top             =   1725
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   53280769
      CurrentDate     =   36641
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   1380
      TabIndex        =   3
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   1380
      TabIndex        =   1
      Top             =   765
      Width           =   795
   End
End
Attribute VB_Name = "FormRelProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String

Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"
Data2 = "Date(" & Year(DtData2) & "," & Month(DtData2) & "," & Day(DtData2) & ")"
MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELPRODUCAO.RPT"
MDIPrincipal.RptRel.SelectionFormula = ""
MDIPrincipal.RptRel.Formulas(1) = "Formulas = '" & Periodo & "'"
'MDIPrincipal.RptRel.Formulas(2) = "Empresa = '" & NomeEmpresa & "'"

MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2

MDIPrincipal.RptRel.RetrieveDataFiles
MDIPrincipal.RptRel.PrintReport

End Sub

Private Sub Form_Load()
Me.Left = 1500
Me.Top = 2000
DtData1 = Date
DtData2 = Date
End Sub

