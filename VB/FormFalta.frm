VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormFalta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10755
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -675
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   11
      Top             =   7425
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
      ScaleWidth      =   10770
      TabIndex        =   9
      Top             =   0
      Width           =   10770
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Relação de Faltas / Troca"
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
         TabIndex        =   10
         Top             =   45
         Width           =   5010
      End
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormFalta.frx":0000
      Height          =   5460
      Left            =   4455
      OleObjectBlob   =   "FormFalta.frx":001A
      TabIndex        =   8
      Top             =   1575
      Width           =   6000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Faltas / Troca"
      Height          =   600
      Left            =   315
      TabIndex        =   4
      Top             =   630
      Width           =   2265
      Begin VB.OptionButton OptFalta 
         Caption         =   "&Trocas"
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   7
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton OptFalta 
         Caption         =   "&Faltas"
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Com Cotação"
      Height          =   330
      Left            =   6930
      TabIndex        =   3
      Top             =   765
      Value           =   1  'Checked
      Width           =   1410
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8685
      TabIndex        =   2
      Top             =   720
      Width           =   1545
   End
   Begin VB.Data DatDados 
      Caption         =   "Falta"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4095
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormFalta.frx":0D6C
      Top             =   810
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1125
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT prd.CodPrd,prd.descrição_Produto, Prd.Falta  FROM prd ORDER BY prd.Descrição_Produto"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSDBCtls.DBCombo DbcProdutos 
      Bindings        =   "FormFalta.frx":0DF4
      DataField       =   "descrição_Produto"
      DataSource      =   "DatDados(1)"
      Height          =   5445
      Left            =   315
      TabIndex        =   6
      Top             =   1575
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   9578
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   "descrição_Produto"
      BoundColumn     =   "CodPrd"
      Text            =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos em Falta"
      Height          =   195
      Left            =   4455
      TabIndex        =   1
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos da Tabela"
      Height          =   195
      Left            =   315
      TabIndex        =   0
      Top             =   1305
      Width           =   1395
   End
End
Attribute VB_Name = "FormFalta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
If Check1.value = 1 Then
    MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelFalta.rpt")
Else
    MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelPrd.rpt")
End If
MDIPrincipal.RptRel.WindowTitle = "Relatório de Produtos em Falta / Trocas"
If OptFalta(0).value = True Then
    MDIPrincipal.RptRel.SelectionFormula = "{Prd.falta} = true"
Else
    MDIPrincipal.RptRel.SelectionFormula = "{Prd.Troca} = true"
End If
MDIPrincipal.RptRel.WindowState = crptMaximized
Imprimir

End Sub


Private Sub DbGrid_dblClick()
Dim Criterio As String
Criterio = "CodPrd = " & DBGrid.Columns(0)
DatDados(0).Recordset.FindFirst Criterio
If Not DatDados(0).Recordset.NoMatch Then
    DatDados(0).Recordset.Edit
    If OptFalta(0).value = True Then
        DatDados(0).Recordset.Fields("Falta") = 0
    Else
        DatDados(0).Recordset.Fields("Troca") = 0
    End If
    DatDados(0).Recordset.Update
    DatDados(0).Refresh
End If

End Sub

Private Sub DbcProdutos_dblClick(Area As Integer)
Dim Criterio As String
Criterio = "CodPrd = " & DbcProdutos.BoundText
DatDados(1).Recordset.FindFirst Criterio
If Not DatDados(1).Recordset.NoMatch Then
    DatDados(1).Recordset.Edit
    If OptFalta(0).value = True Then
        DatDados(1).Recordset.Fields("Falta") = -1
    Else
        DatDados(1).Recordset.Fields("Troca") = -1
    End If
    DatDados(1).Recordset.Update
    DatDados(0).Refresh
End If
End Sub

Private Sub Form_Load()
Dim Componente As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
FormFalta.Top = 200
FormFalta.Left = 200

DatDados(1).RecordSource = "SELECT prd.CodPrd,prd.descrição_Produto, Prd.Falta, Prd.Troca  FROM prd ORDER BY prd.Descrição_Produto"
DatDados(1).Refresh
OptFalta_Click 0

End Sub

Private Sub OptFalta_Click(Index As Integer)
If Index = 0 Then
    DatDados(0).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Falta, prd.Troca FROM PRD WHERE (((PRD.Falta)=True)) ORDER BY PRD.Descrição_Produto"
Else
    DatDados(0).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Falta, prd.Troca FROM PRD WHERE (((PRD.Troca)=True)) ORDER BY PRD.Descrição_Produto"
End If
DatDados(0).Refresh
End Sub

