VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormGerPrd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6915
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7395
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -630
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   18
      Top             =   6570
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
      ScaleWidth      =   7395
      TabIndex        =   16
      Top             =   0
      Width           =   7395
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatística de Produtos"
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
         TabIndex        =   17
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagamento em Cheque"
      Height          =   1470
      Left            =   1170
      TabIndex        =   11
      Top             =   4815
      Width           =   5475
      Begin VB.Data DatDados 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   2
         Left            =   4095
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT REC.Terceiro FROM REC GROUP BY REC.Terceiro ORDER BY REC.Terceiro"
         Top             =   270
         Visible         =   0   'False
         Width           =   1230
      End
      Begin MSDBCtls.DBCombo DBCTerceiro 
         Bindings        =   "FormGerPrd.frx":0000
         Height          =   420
         Left            =   135
         TabIndex        =   12
         Top             =   900
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MebChPre 
         Height          =   375
         Left            =   1575
         TabIndex        =   13
         Top             =   225
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Datado para"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Classificado"
      Height          =   2715
      Left            =   495
      TabIndex        =   5
      Top             =   1755
      Width           =   3750
      Begin VB.OptionButton Opt 
         Caption         =   "Produtos por Giro"
         Height          =   375
         Index           =   4
         Left            =   585
         TabIndex        =   10
         Top             =   2160
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Produtos por Data de Compra"
         Height          =   375
         Index           =   0
         Left            =   585
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   2805
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Produtos por Data de Venda"
         Height          =   375
         Index           =   1
         Left            =   585
         TabIndex        =   8
         Top             =   765
         Width           =   2535
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Produtos mais Comprados"
         Height          =   375
         Index           =   2
         Left            =   585
         TabIndex        =   7
         Top             =   1230
         Width           =   2670
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Produtos mais Vendidos"
         Height          =   375
         Index           =   3
         Left            =   585
         TabIndex        =   6
         Top             =   1695
         Width           =   2940
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   4770
      TabIndex        =   4
      Top             =   2070
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   1275
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   2550
      TabIndex        =   2
      Top             =   1275
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   2550
      TabIndex        =   3
      Top             =   1035
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   495
      TabIndex        =   1
      Top             =   1035
      Width           =   795
   End
End
Attribute VB_Name = "FormGerPrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Selecao As Recordset
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String

'MDIPrincipal.ComRel.ShowPrinter

Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"
Data2 = "Date(" & Year(DtData2) & "," & Month(DtData2) & "," & Day(DtData2) & ")"

Periodo = "Período de " & DtData1 & " a " & DtData2

If Opt(0).value = True Then
        MDIPrincipal.RptRel.ReportFileName = Caminho & "\relconf.RPT"
    ElseIf Opt(1).value = True Then
            MDIPrincipal.RptRel.ReportFileName = Caminho & "\relVda.RPT"
        ElseIf Opt(2).value = True Then
                MDIPrincipal.RptRel.ReportFileName = Caminho & "\relVdaCli.RPT"
            ElseIf Opt(3).value = True Then
                    If ListaSel.ListIndex = 0 Then
                        MDIPrincipal.RptRel.ReportFileName = Caminho & "\TOTGRU.RPT"
                    Else
                        MDIPrincipal.RptRel.ReportFileName = Caminho & "\relVdaGru.RPT"
                    End If
                Else
                    MDIPrincipal.RptRel.ReportFileName = Caminho & "\comiss.RPT"
End If

If ListaOper.ListIndex = 0 Then
    If ListaSel.ListIndex = 0 Or ListaSel.Enabled = False Then
        MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3] "
    Else
        If Opt(2).value = True Then
              MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3] and {Ped.Entidade} = " & CLng(Left(ListaSel.Text, 5))
           ElseIf Opt(3).value = True Then
                  If ListaSel.ListIndex = 0 Then
                      MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao}  = " & ListaOper.ListIndex - 1
                  ElseIf ListaSel.ListIndex = 1 Then
                          MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3]"
                      Else
                          MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3] and {Prd.Grupo} = " & CLng(Left(ListaSel.Text, 4))
                      End If
              ElseIf Opt(4).value = True Then
                    MDIPrincipal.RptRel.ReportFileName = Caminho & "\RelVdaFun.RPT"
                    MDIPrincipal.RptRel.Formulas(1) = "Vendedor = '" & Mid$(ListaSel.Text, 9, 40) & "' "
                    MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3] and {Ped.Vendedor} = " & CLng(Left(ListaSel.Text, 5))
                  Else
                     MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3] "
        End If
    End If
Else
    If ListaSel.ListIndex = 0 Or ListaSel.Enabled = False Then
        MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} = " & ListaOper.ListIndex - 1
    Else
        If Opt(2).value = True Then
              MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} = " & ListaOper.ListIndex - 1 & " and {Ped.Entidade} = " & CLng(Left(ListaSel.Text, 5))
           ElseIf Opt(3).value = True Then
                  If ListaSel.ListIndex = 0 Then
                      MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao}  = " & ListaOper.ListIndex - 1
                  ElseIf ListaSel.ListIndex = 1 Then
                          MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} in [1,2,3]"
                      Else
                          MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao}  = " & ListaOper.ListIndex - 1 & " and {Prd.Grupo} = " & CLng(Left(ListaSel.Text, 4))
                      End If
              ElseIf Opt(4).value = True Then
                    MDIPrincipal.RptRel.ReportFileName = Caminho & "\RelVdaFun.RPT"
                    MDIPrincipal.RptRel.Formulas(1) = "Vendedor = '" & Mid$(ListaSel.Text, 9, 40) & "' "
                    MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} = " & ListaOper.ListIndex - 1 & " and {Ped.Vendedor} = " & CLng(Left(ListaSel.Text, 5))
                  Else
                     MDIPrincipal.RptRel.SelectionFormula = "{Ped.DataEmiss} >= " & Data1 & " AND {Ped.DataEmiss} <= " & Data2 & " and {Ped.Operacao} = " & ListaOper.ListIndex - 1
        End If
    
    End If

End If

MDIPrincipal.RptRel.Formulas(0) = "Formulas = '" & Periodo & "' "

Imprimir

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DtData1 = Date
DtData2 = Date

ListaOper.Enabled = True
ListaSel.Enabled = False
ListaOper.ListIndex = 0
End Sub

Private Sub Opt_Click(Index As Integer)
Select Case Index
    Case 0
        ListaSel.Enabled = False
    Case 1
        ListaSel.Enabled = False
    Case 2
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Todos)")
        Set Selecao = Banco.OpenRecordset("select * from ent where tipo = 1 or tipo >= 5 order by nome")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodEntidade"), 5) & " - " & Selecao.Fields("Nome"))
                Selecao.MoveNext
            Wend
        End If
        ListaSel.ListIndex = 0
    Case 3
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Totalizado)")
        ListaSel.AddItem ("(Todos)")
        Set Selecao = Banco.OpenRecordset("select * from gru  order by Descrição_Grupo")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodGru"), 4) & " - " & Selecao("Descrição_Grupo"))
                Selecao.MoveNext
            Wend
         End If
         ListaSel.ListIndex = 0
    Case 4
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Todos)")
        Set Selecao = Banco.OpenRecordset("select * from ent where tipo = 3 or tipo = 7 order by nome")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodEntidade"), 5) & " - " & Selecao.Fields("Nome"))
                Selecao.MoveNext
            Wend
        End If
        ListaSel.ListIndex = 0
End Select

End Sub
