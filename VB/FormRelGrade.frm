VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormRelGrade 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6150
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   9
      Top             =   3195
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
      TabIndex        =   7
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios de Vendas por Grade"
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
         TabIndex        =   8
         Top             =   45
         Width           =   5415
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   1485
      TabIndex        =   5
      Top             =   1845
      Width           =   1320
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   3240
      TabIndex        =   4
      Top             =   1845
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   1425
      TabIndex        =   0
      Top             =   1095
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   53149697
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   3180
      TabIndex        =   2
      Top             =   1095
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   53149697
      CurrentDate     =   36641
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   2655
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
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
      Left            =   3180
      TabIndex        =   3
      Top             =   855
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
      Left            =   1425
      TabIndex        =   1
      Top             =   855
      Width           =   990
   End
End
Attribute VB_Name = "FormRelGrade"
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
Dim SQL As String

' MDIPrincipal.Commom.ShowPrinter
Data1 = Format(DtData1, "mm/dd/yyyy")
Data2 = Format(DtData2, "mm/dd/yyyy")



    Barra.Visible = True
    Barra.Min = 0

    Barra.value = 0
    
    Set Ind = Banco.OpenRecordset("Select * from  Ind")
    Ind.Requery
    
    Set Prod = Banco.OpenRecordset("Select * From Prd")
    Set Rel = Banco.OpenRecordset("Select * From RelPrd")
    SQL = "Delete * from RelPrd"
    Banco.Execute SQL
    
    Prod.MoveLast
    Barra.Max = Prod.RecordCount
    Prod.MoveFirst
    While Not Prod.Eof
            Rel.AddNew
            Rel("produto") = Prod("CodPrd")
            Ind.MoveFirst
            x = 1
            While Not Ind.Eof
                If x <= 5 Then
                        Rel("Prazo" & x) = Prod("Venda") * Ind("Indice")
                End If
                Ind.MoveNext
                x = x + 1
            Wend
            Rel.Update
            Barra.value = Prod.AbsolutePosition
            Prod.MoveNext
    Wend
    Command1.Enabled = True
    Barra.Visible = False









If DbcEnt = "" Then
    If LstTipo.ListIndex = 0 Then
        SQL = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo < 10 ORDER BY REC.CodRec "
    Else
        SQL = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo = " & LstTipo.ListIndex & " ORDER BY REC.CodRec "
    End If
Else
    If LstTipo.ListIndex = 0 Then
        SQL = " SELECT REC.Juros , REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo < 10  AND Ent.CodEntidade = " & DbcEnt.BoundText & " ORDER BY REC.CodRec "
    Else
        SQL = " SELECT REC.Juros , REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo = " & LstTipo.ListIndex & "  AND Ent.CodEntidade = " & DbcEnt.BoundText & " ORDER BY REC.CodRec "
    End If
End If

Set varPesquisa(1) = Banco.OpenRecordset(SQL)
Set FrmRelRec.Relatorio.Recordset = varPesquisa(1)
FrmRelRec.rpCabTitulo.Caption = "Relatório de Recebimento"
FrmRelRec.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelRec.Config

'MDIPrincipal.RptRel.ReportFileName = Caminho & "\RELREC.RPT"
'If DbcEnt.Text = "" Then
'    MDIPrincipal.RptRel.SelectionFormula = "{Rec.Data} >= " & Data1 & " AND {Rec.Data} <= " & Data2
'Else
'    MDIPrincipal.RptRel.SelectionFormula = "{Rec.Data} >= " & Data1 & " AND {Rec.Data} <= " & Data2 & " AND {Rec.CodEnt} =  " & DbcEnt.BoundText
'End If
'Imprimir
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = 2000
Me.Top = 2500
DtData1 = Date
DtData2 = Date
LstTipo.ListIndex = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


