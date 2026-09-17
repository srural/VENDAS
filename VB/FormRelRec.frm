VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelRec 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4065
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5970
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   315
      TabIndex        =   14
      Top             =   630
      Width           =   5145
      Begin VB.OptionButton OptVda 
         Caption         =   "Vendas"
         Height          =   330
         Left            =   3060
         TabIndex        =   16
         Top             =   225
         Width           =   1005
      End
      Begin VB.OptionButton OptRec 
         Caption         =   "Recebimento"
         Height          =   285
         Left            =   765
         TabIndex        =   15
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.CheckBox chkjuros 
      Caption         =   "Controle Juros"
      Height          =   240
      Left            =   3870
      TabIndex        =   13
      Top             =   1620
      Width           =   1590
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   12
      Top             =   3735
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
      TabIndex        =   10
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios de Recebimento"
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
         TabIndex        =   11
         Top             =   45
         Width           =   5775
      End
   End
   Begin VB.ComboBox LstTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormRelRec.frx":0000
      Left            =   300
      List            =   "FormRelRec.frx":0019
      TabIndex        =   8
      Top             =   3105
      Width           =   1950
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   2565
      TabIndex        =   7
      Top             =   3105
      Width           =   1320
   End
   Begin VB.Data DatDados 
      Caption         =   "Entidade"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM ENT ORDER BY Nome"
      Top             =   2295
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   4095
      TabIndex        =   4
      Top             =   3105
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   300
      TabIndex        =   0
      Top             =   1590
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   85852161
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   2055
      TabIndex        =   2
      Top             =   1590
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   85852161
      CurrentDate     =   36641
   End
   Begin MSDBCtls.DBCombo DbcEnt 
      Bindings        =   "FormRelRec.frx":0066
      Height          =   360
      Left            =   300
      TabIndex        =   5
      Top             =   2250
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Caption         =   "Tipo"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   2835
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Entidade"
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
      Left            =   300
      TabIndex        =   6
      Top             =   1575
      Width           =   810
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
      Left            =   2055
      TabIndex        =   3
      Top             =   1350
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
      Left            =   300
      TabIndex        =   1
      Top             =   1350
      Width           =   990
   End
End
Attribute VB_Name = "FormRelRec"
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
Dim Sql As String
' MDIPrincipal.Commom.ShowPrinter
Data1 = Format(DtData1, "mm/dd/yyyy")
Data2 = Format(DtData2, "mm/dd/yyyy")

If OptRec.value = True Then
    If chkjuros.value = 1 Then
    '        If LstTipo.ListIndex = 0 Then
                Sql = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo , DUP.DtVenc FROM (REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade) INNER JOIN DUP ON REC.CodDupli = DUP.CodDup Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo < 10 ORDER BY REC.CodRec "
    '        Else
    '            SQL = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo = " & LstTipo.ListIndex & " ORDER BY REC.CodRec "
    '        End If
    Else
    
        If DbcEnt = "" Then
            If LstTipo.ListIndex = 0 Then
                Sql = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo < 10 ORDER BY REC.CodRec "
            Else
                Sql = " SELECT REC.Juros, REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo = " & LstTipo.ListIndex & " ORDER BY REC.CodRec "
            End If
        Else
            If LstTipo.ListIndex = 0 Then
                Sql = " SELECT REC.Juros , REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo < 10  AND Ent.CodEntidade = " & DbcEnt.BoundText & " ORDER BY REC.CodRec "
            Else
                Sql = " SELECT REC.Juros , REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo = " & LstTipo.ListIndex & "  AND Ent.CodEntidade = " & DbcEnt.BoundText & " ORDER BY REC.CodRec "
            End If
        End If
    End If
    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
    Set FrmRelRec.Relatorio.Recordset = varPesquisa(1)
    FrmRelRec.rpCabTitulo.Caption = "Relatório de Recebimento"
    FrmRelRec.rpCabSubTitulo.Caption = NomeEmpresa
    If chkjuros.value = 1 Then
        FrmRelRec.RelTEmp.Tag = 1
    Else
        FrmRelRec.RelTEmp.Tag = 0
    End If
    FrmRelRec.Config
Else
    Sql = " SELECT REC.Juros , REC.CodEnt, ENT.Nome, REC.Data, REC.Hora, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM REC INNER JOIN ENT ON REC.CodEnt = ENT.CodEntidade Where Rec.Data >= #" & Data1 & "# AND rec.Data <= #" & Data2 & "# and Rec.Tipo > 10 ORDER BY REC.CodRec "
    
    Set varPesquisa(1) = Banco.OpenRecordset(Sql)
    Set FrmRelRec.Relatorio.Recordset = varPesquisa(1)
    FrmRelRec.rpCabTitulo.Caption = "Relatório de Recebimento/Vendas"
    FrmRelRec.rpCabSubTitulo.Caption = NomeEmpresa
    If chkjuros.value = 1 Then
        FrmRelRec.RelTEmp.Tag = 1
    Else
        FrmRelRec.RelTEmp.Tag = 0
    End If
    FrmRelRec.Config


End If
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


Private Sub OptRec_Click()
chkjuros.Enabled = True
DbcEnt.Enabled = True
LstTipo.Enabled = True

End Sub

Private Sub OptVda_Click()
chkjuros.Enabled = False
DbcEnt.Enabled = False
LstTipo.Enabled = False
End Sub
