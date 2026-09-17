VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMovEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Movimentação de Estoque"
   ClientHeight    =   2145
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5100
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   37181
   End
   Begin VB.CommandButton BtImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   630
      Left            =   3960
      Picture         =   "FormMovEst.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   765
      Width           =   765
   End
   Begin VB.Frame Frm1 
      Caption         =   "&Tipo"
      Height          =   1635
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   1695
      Begin VB.OptionButton Opt1 
         Caption         =   "Tranferência"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   1125
         Width           =   1275
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "&Entrada"
         Height          =   375
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Saída"
         Height          =   375
         Index           =   3
         Left            =   195
         TabIndex        =   1
         Top             =   720
         Width           =   915
      End
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   1305
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   37181
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   270
      Width           =   795
   End
End
Attribute VB_Name = "FormmoVeST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Activate()


Left = 350
Top = 150


End Sub

Private Sub BtImprimir_Click()
Dim x As Variant
Dim Op1 As Integer
Dim Op2 As Integer
Dim Formula As String

MDIPrincipal.ComRel.ShowPrinter
                        
For Each x In Opt1
    If x.Value Then
        Op1 = x.Index
    End If
Next

MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelEstLoja.rpt")
MDIPrincipal.RptRel.WindowTitle = "Relatório de Estoque po Loja"

End Sub

Private Sub Imprimir_Click()

End Sub


Private Sub Form_Load()

Set Formulário = Me

DtData(0).Value = Date
DtData(1).Value = Date

End Sub
