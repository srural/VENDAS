VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormRelPag 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4470
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6285
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   19
      Top             =   4140
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
         Caption         =   "Relatórios de Contas a Pagar"
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
         TabIndex        =   18
         Top             =   45
         Width           =   5145
      End
   End
   Begin VB.CheckBox Chkdoc 
      Caption         =   "Documento"
      Height          =   330
      Left            =   3105
      TabIndex        =   16
      Top             =   2700
      Width           =   1680
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   3285
      TabIndex        =   14
      Top             =   3465
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   270
      TabIndex        =   11
      Top             =   2430
      Width           =   2445
      Begin VB.OptionButton Opt 
         Caption         =   "Por &Data"
         Height          =   375
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   945
         Width           =   1860
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Por &Tipo"
         Height          =   375
         Index           =   5
         Left            =   270
         TabIndex        =   13
         Top             =   180
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Pelo Plano de Contas"
         Height          =   375
         Index           =   4
         Left            =   270
         TabIndex        =   12
         Top             =   562
         Width           =   1860
      End
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      ItemData        =   "FormRelPag.frx":0000
      Left            =   3105
      List            =   "FormRelPag.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2250
      Width           =   2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duplicatas"
      Height          =   1410
      Left            =   270
      TabIndex        =   5
      Top             =   765
      Width           =   2445
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   2
         Left            =   270
         TabIndex        =   8
         Top             =   945
         Width           =   1050
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pagas"
         Height          =   375
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   608
         Width           =   1050
      End
      Begin VB.OptionButton Opt 
         Caption         =   "À Pagar"
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   4635
      TabIndex        =   4
      Top             =   3465
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   375
      Left            =   3135
      TabIndex        =   0
      Top             =   855
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   53280769
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   375
      Left            =   3135
      TabIndex        =   2
      Top             =   1530
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   53280769
      CurrentDate     =   36641
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo"
      Height          =   195
      Left            =   3105
      TabIndex        =   10
      Top             =   2025
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   3135
      TabIndex        =   3
      Top             =   1305
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   3135
      TabIndex        =   1
      Top             =   630
      Width           =   795
   End
End
Attribute VB_Name = "FormRelPag"
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
Dim Tipo As String
Dim SQL As String
Dim Formula As String
Dim Ordenado As String

If Opt(3).value = True Then
    SQL = "SELECT PAG.Data, PAG.DtVenc, PAG.DtPgto, PAG.Fornecedor, ENT.Nome, PAG.Valor, PAG.VlrPago, PAG.NroNota, PAG.Cond, PAG.Numero, PAG.Obs, ENT.Cidade, PAG.Tipo FROM ENT INNER JOIN PAG ON ENT.CodEntidade = PAG.Fornecedor "
    If Opt(0).value = True Then
            Formula = " WHERE  PAG.DtVenc >=# " & Format(DtData1, "mm/dd/yyyy") & " # And PAG.DtVenc <=# " & Format(DtData2, "mm/dd/yyyy") & " # AND PAG.Cond = False "
            If CmbTipo.ListIndex < 4 Then
                Formula = Formula & " AND Pag.Tipo = " & CmbTipo.ListIndex
            End If
            
            If Chkdoc.value = 1 Then
                 Formula = Formula & " AND Pag.NroNota <>  " & """" & "" & """"
            End If
            
            Ordenado = " ORDER BY PAG.DtVenc"
            FrmRelPagData.VarData.Campo = "DtVenc"
        
        ElseIf Opt(1).value = True Then
           Formula = " WHERE PAG.DtPgto >=# " & Format(DtData1, "mm/dd/yyyy") & " # And PAG.DtPgto <=# " & Format(DtData2, "mm/dd/yyyy") & " # AND PAG.Cond = True "
            If CmbTipo.ListIndex < 4 Then
                Formula = Formula & " AND Pag.Tipo = " & CmbTipo.ListIndex
            End If
           
            If Chkdoc.value = 1 Then
                 Formula = Formula & " AND Pag.NroNota <>  " & """" & "" & """"
            End If
           
           Ordenado = " ORDER BY PAG.DtPgto"
           FrmRelPagData.VarData.Campo = "DtPgto"
           
        Else
           Formula = " WHERE PAG.DtData >=# " & Format(DtData1, "mm/dd/yyyy") & " # And PAG.DtData <=# " & Format(DtData2, "mm/dd/yyyy") & " # "
            If CmbTipo.ListIndex < 4 Then
                Formula = Formula & " AND Pag.Tipo = " & CmbTipo.ListIndex
            End If
           
            If Chkdoc.value = 1 Then
                 Formula = Formula & " AND Pag.NroNota <>  " & """" & "" & """"
            End If
           
           Ordenado = " ORDER BY PAG.Data"
           FrmRelPagData.VarData.Campo = "Data"
           
    End If

    SQL = SQL & Formula & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPagData.Relatorio.Recordset = varPesquisa(1)
    FrmRelPagData.rpCabTitulo.Caption = "Relatório de Contas a Pagar / Pagas"
    FrmRelPagData.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPagData.rpCabPeriodo.Caption = Periodo
    FrmRelPagData.Config


Else
    Periodo = Format(DtData1, "dd/mm/yyyy") & " a " & Format(DtData2, "dd/mm/yyyy")
    
    Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"
    Data2 = "Date(" & Year(DtData2) & "," & Month(DtData2) & "," & Day(DtData2) & ")"
    
    If Opt(5).value = True Then
        MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELPAG.RPT"
    ElseIf Opt(4).value = True Then
        If MsgBox("Relatorio Completo ?", vbYesNo, App.Title) = vbYes Then
            MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELPAGCTA.RPT"
        Else
            MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELPAGCTA2.RPT"
        End If
    End If
    MDIPrincipal.RptRel.SelectionFormula = ""
    MDIPrincipal.RptRel.Formulas(1) = "Periodo = '" & Periodo & "'"
    MDIPrincipal.RptRel.Formulas(2) = "Empresa = '" & NomeEmpresa & "'"
    
    
    If CmbTipo.ListIndex = 4 Then
        Tipo = "{Pag.Tipo} <= 10"
    Else
        Tipo = "{Pag.Tipo} =  " & CmbTipo.ListIndex
    End If
    
    If Opt(0).value = True Then
        MDIPrincipal.RptRel.SelectionFormula = "{Pag.DtVenc} >= " & Data1 & " AND {Pag.DtVenc} <= " & Data2 & "AND {Pag.Cond} = False AND " & Tipo
    ElseIf Opt(1).value = True Then
            MDIPrincipal.RptRel.SelectionFormula = "{Pag.DtPgto} >= " & Data1 & " AND {Pag.DtPgto} <= " & Data2 & " AND {Pag.Cond} = true AND" & Tipo
        ElseIf Opt(2).value = True Then
            MDIPrincipal.RptRel.SelectionFormula = "{Pag.DtVenc} >= " & Data1 & " AND {Pag.DtVenc} <= " & Data2 & " AND " & Tipo
    End If
    
    MDIPrincipal.RptRel.RetrieveDataFiles
    MDIPrincipal.RptRel.PrintReport

End If

End Sub

Private Sub Form_Load()
Me.Left = 1500
Me.Top = 2000
DtData1 = Date
DtData2 = Date

CmbTipo.ListIndex = 4
End Sub

