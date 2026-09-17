VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmManut 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "FrmManut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8415
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   8670
      TabIndex        =   17
      Top             =   5940
      Width           =   8700
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11085
      TabIndex        =   15
      Top             =   0
      Width           =   11085
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão de Controle de Entrega"
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
         TabIndex        =   16
         Top             =   45
         Width           =   6540
      End
   End
   Begin VB.TextBox TxtMontador 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2565
      MaxLength       =   30
      TabIndex        =   6
      Top             =   4635
      Width           =   3840
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   360
      MaxLength       =   200
      TabIndex        =   7
      Top             =   5220
      Width           =   6045
   End
   Begin VB.CheckBox Chk 
      Caption         =   "&Check List"
      Height          =   330
      Left            =   6795
      TabIndex        =   9
      Top             =   4050
      Width           =   1320
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      Height          =   2055
      ItemData        =   "FrmManut.frx":000C
      Left            =   360
      List            =   "FrmManut.frx":000E
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1620
      Width           =   7740
   End
   Begin VB.CommandButton BtConfirma 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   6795
      TabIndex        =   8
      Top             =   5220
      Width           =   1200
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   4050
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   7
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entidade"
      Height          =   765
      Left            =   330
      TabIndex        =   0
      Top             =   765
      Width           =   7770
      Begin VB.Data DatDados 
         Caption         =   "Entidade"
         Connect         =   "Access"
         DatabaseName    =   "\\Frente\c\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   2025
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From ENT Order by Nome"
         Top             =   270
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FrmManut.frx":0010
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
   End
   Begin MSMask.MaskEdBox Revenda 
      Height          =   315
      Left            =   1755
      TabIndex        =   4
      Top             =   4050
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   4635
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   40
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Observação (roteiro de acesso)"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4995
      Width           =   2220
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Montador"
      Height          =   195
      Left            =   2565
      TabIndex        =   13
      Top             =   4410
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nro Série"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   4410
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Revenda"
      Height          =   195
      Left            =   1755
      TabIndex        =   11
      Top             =   3825
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nota Fiscal"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   3825
      Width           =   795
   End
End
Attribute VB_Name = "FrmManut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtConfirma_Click()
On Error Resume Next

'MDIPrincipal.ComRel.ShowPrinter



With MDIPrincipal.RptRel
    .ReportFileName = Caminho & "\rptmanut.rpt"
    .WindowState = crptMaximized
    .WindowTitle = "PLANILHA DE MONTAGEM"
    For x = 1 To Lista.ListCount
        If Lista.Selected(x - 1) = True Then
            Select Case x
                Case 1
                    .Formulas(0) = "Bens1 = '" & Mid$(Lista.List(x - 1), 9, 40) & "'"
                Case 2
                    .Formulas(1) = "Bens2 = '" & Mid$(Lista.List(x - 1), 9, 40) & "'"
                Case 3
                    .Formulas(2) = "Bens3 = '" & Mid$(Lista.List(x - 1), 9, 40) & "'"
                Case 4
                    .Formulas(3) = "Bens4 = '" & Mid$(Lista.List(x - 1), 9, 40) & "'"
            End Select
        End If
    Next
    
    .Formulas(4) = "Valor = '" & Valor.Text & "'"
    .Formulas(5) = "Obs = '" & TxtObs.Text & "'"
    .Formulas(6) = "Revenda = '" & Revenda & "'"
    .Formulas(7) = "Montador = '" & TxtMontador & "'"
    .SelectionFormula = "{Ent.CodEntidade} = " & DbcDados(0).BoundText & ""
    
'    .Formulas(0) = "Entidade = '" & DbcDados(0).Text & "'"
'    .Formulas(1) = "Nacionalidade = '" & DatDados(0).Recordset("Pais") & "'"
'    .Formulas(2) = "EstadoCivil = '" & DatDados(0).Recordset("EstadoCivil") & "'"
'    .Formulas(3) = "Profissao = '" & DatDados(0).Recordset("Profissao") & "'"
'    .Formulas(4) = "CPF = '" & DatDados(0).Recordset("CPF") & "'"
'    .Formulas(5) = "RG = '" & DatDados(0).Recordset("RG") & "'"
'    .Formulas(6) = "Endereco = '" & DatDados(0).Recordset("Endereco") & "'"
'    .Formulas(7) = "Nro = '" & DatDados(0).Recordset("Nro") & "'"
'    .Formulas(8) = "Municipio = '" & DatDados(0).Recordset("Cidade") & "'"
'    .Formulas(9) = "Estado = '" & DatDados(0).Recordset("Uf") & "'"
'    .Formulas(10) = "Valor = '" & Format(Valor.Text, "#,#0.00") & "'"
'    .Formulas(11) = "FormaDePagamento = '" & FormaPagamento.Text & "'"
    
End With

Imprimir
End Sub

Private Sub DbcDados_Click(Index As Integer, Area As Integer)
Lista.Clear
Lista.Enabled = True
If IsNumeric(DbcDados(0).BoundText) Then
    Set pedidos = Banco.OpenRecordset("SELECT PED.Entidade, PRD.Descrição_Produto, First(PED.CodPed) AS PrimeiroCodPed FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where Entidade = " & DbcDados(0).BoundText & " GROUP BY PED.Entidade, PRD.Descrição_Produto")
    If pedidos.RecordCount <> 0 Then
        Do While Not pedidos.Eof
            Lista.AddItem psCompString(1, pedidos("PrimeiroCodPed"), 6) & " - " & pedidos("Descrição_Produto")
            pedidos.MoveNext
        Loop
    Else
        Lista.AddItem "Nenhum Produto Encontrado"
        Lista.Enabled = False
    End If
    
DatDados(0).Recordset.FindFirst "CodEntidade = " & DbcDados(0).BoundText

End If
End Sub

Private Sub Form_Activate()
'CmbEnd.ListIndex = 0
Chk.value = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub
