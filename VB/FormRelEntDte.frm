VERSION 5.00
Begin VB.Form FormRelEntDte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios de Entidades"
   ClientHeight    =   2850
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4515
   Begin VB.ComboBox CmbAniver 
      Height          =   315
      ItemData        =   "FormRelEntDte.frx":0000
      Left            =   180
      List            =   "FormRelEntDte.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2250
      Width           =   1905
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   2250
      Width           =   1800
   End
   Begin VB.ListBox Lista 
      Height          =   1425
      ItemData        =   "FormRelEntDte.frx":0091
      Left            =   180
      List            =   "FormRelEntDte.frx":00AA
      TabIndex        =   6
      Top             =   405
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   1635
      Left            =   2475
      TabIndex        =   0
      Top             =   270
      Width           =   1815
      Begin VB.OptionButton Opt 
         Caption         =   "Aniversariantes"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   1215
         Width           =   1500
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Inativos"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   870
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ativos"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   525
         Width           =   825
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Aniversariante (Mês)"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Label LblDados 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Index           =   56
      Left            =   225
      TabIndex        =   4
      Top             =   135
      Width           =   315
   End
End
Attribute VB_Name = "FormRelEntDte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbCancel_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula  As String

MDIPrincipal.ComRel.ShowPrinter

Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"

'If Opt(3).Value = True Then
'    MDIPrincipal.RptRel.ReportFileName = Caminho & "\relAni.RPT"
'Else
'    MDIPrincipal.RptRel.ReportFileName = Caminho & "\relENT.RPT"
'End If

Select Case Lista.ListIndex
    Case 0
        Formula = "Tipo = 1"
    Case 1
        Formula = "Tipo = 2"
    Case 2
        Formula = "Tipo = 3"
    Case 3
        Formula = "Tipo = 4"
    Case 4
        Formula = "Tipo = 5"
    Case 5
        Formula = "Tipo = 6"
    Case 6
        Formula = "Tipo = 7"
End Select

If Opt(3).Value = True Then
    Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1
Else
    If Opt(1).Value = True Then
            Formula = Formula + " and Ativo = True "
        ElseIf Opt(2).Value = True Then
            Formula = Formula + " and Ativo = False "
    End If
End If

'MDIPrincipal.RptRel.SelectionFormula = Formula
'MDIPrincipal.RptRel.Formulas(0) = "Formulas = '" & Periodo & "' "
'Imprimir

If DteBanco.rsEntidades.State = 1 Then DteBanco.rsEntidades.Close
DteBanco.rsEntidades.Source = "Select * From Ent Where " & Formula
DteBanco.rsEntidades.Open

RelEnt.Show

End Sub

Private Sub Form_Load()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

Lista.ListIndex = 0
CmbAniver.Enabled = False
End Sub

Private Sub Opt_Click(Index As Integer)

If Index = 3 Then
    If Opt(3).Value = True Then
        CmbAniver.Enabled = True
    End If
End If

End Sub
