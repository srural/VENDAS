VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormRelCxa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2970
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6300
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   8
      Top             =   2610
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
         Caption         =   "Relatórios de Movimentação do Caixa"
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
         TabIndex        =   7
         Top             =   45
         Width           =   6090
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   375
      Left            =   3285
      TabIndex        =   5
      Top             =   1395
      Width           =   1590
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3285
      TabIndex        =   4
      Top             =   1890
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   375
      Left            =   1245
      TabIndex        =   0
      Top             =   1095
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   375
      Left            =   1245
      TabIndex        =   2
      Top             =   1860
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1245
      TabIndex        =   3
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1245
      TabIndex        =   1
      Top             =   855
      Width           =   990
   End
End
Attribute VB_Name = "FormRelCxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConfImpr_Click()
 MDIPrincipal.RptRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String

Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"
Data2 = "Date(" & Year(DtData2) & "," & Month(DtData2) & "," & Day(DtData2) & ")"

MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELCXA.RPT"
'If DbcCaixa.Text = "" Then
MDIPrincipal.RptRel.Formulas(2) = "Empresa = '" & NomeEmpresa & "'"
MDIPrincipal.RptRel.SelectionFormula = "{CXA.Data} >= " & Data1 & " AND {CXA.Data} <= " & Data2
'Else
'    MDIPrincipal.RptRel.SelectionFormula = "{CXA.Data} >= " & Data1 & " AND {CXA.Data} <= " & Data2 & " AND {Cxa.Maquina} =  '" & Trim(DbcCaixa) & "'"
'End If
Imprimir
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

Me.Left = 2000
Me.Top = 2500
DtData1 = Date
DtData2 = Date

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub
