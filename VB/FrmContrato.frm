VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmContrato 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "FrmContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8475
   Begin VB.TextBox FormaPagamento 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   270
      TabIndex        =   11
      Top             =   6030
      Width           =   7935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -720
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   10
      Top             =   6705
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
      ScaleWidth      =   14655
      TabIndex        =   8
      Top             =   0
      Width           =   14655
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão de Contrato"
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
         TabIndex        =   9
         Top             =   45
         Width           =   7800
      End
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      Height          =   3405
      ItemData        =   "FrmContrato.frx":000C
      Left            =   225
      List            =   "FrmContrato.frx":000E
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1650
      Width           =   8055
   End
   Begin VB.CommandButton BtConfirma 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5310
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   270
      TabIndex        =   5
      Top             =   5445
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   195
      TabIndex        =   0
      Top             =   720
      Width           =   8085
      Begin VB.Data DatDados 
         Caption         =   "Entidade"
         Connect         =   "Access"
         DatabaseName    =   "\\Frente\c\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   2790
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From ENT Order by Nome"
         Top             =   360
         Visible         =   0   'False
         Width           =   2100
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FrmContrato.frx":0010
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entidade"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   150
         Width           =   630
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pagamento"
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   5805
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   5175
      Width           =   360
   End
End
Attribute VB_Name = "FrmContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pedidos As Recordset

Private Sub BtConfirma_Click()
On Error Resume Next
Dim Var1 As String
Dim Var2 As String
Dim Var3 As String
Dim Var4 As String

Var1 = ""
Var2 = ""
Var3 = ""
Var4 = ""

'MDIPrincipal.ComRel.ShowPrinter

For X = 0 To pedidos.RecordCount
    If Len(Var1 & ", " & Lista.List(X)) <= 96 Then
        If Lista.Selected(X) = True Then Var1 = Var1 & Lista.List(X) & " ,"
        
        ElseIf Len(Var2 & ", " & Lista.List(X)) <= 96 Then
            If Lista.Selected(X) = True Then Var2 = Var2 & Lista.List(X) & " ,"
            
                ElseIf Len(Var3 & ", " & Lista.List(X)) <= 96 Then
                    If Lista.Selected(X) = True Then Var3 = Var3 & Lista.List(X) & " ,"
                    
                        ElseIf Len(Var4 & ", " & Lista.List(X)) <= 96 Then
                            If Lista.Selected(X) = True Then Var4 = Var4 & Lista.List(X) & " ,"
    End If
Next

With MDIPrincipal.RptRel
    .ReportFileName = Caminho & "\rptcontrato.rpt"
    .WindowState = crptMaximized
    .WindowTitle = "RELATÓRIO DE CONTRATO"
    .Formulas(0) = "Entidade = '" & DbcDados(0).Text & "'"
    .Formulas(1) = "Nacionalidade = '" & DatDados(0).Recordset("Pais") & "'"
    .Formulas(2) = "EstadoCivil = '" & DatDados(0).Recordset("EstadoCivil") & "'"
    .Formulas(3) = "Profissao = '" & DatDados(0).Recordset("Profissao") & "'"
    .Formulas(4) = "CPF = '" & DatDados(0).Recordset("CPF") & "'"
    .Formulas(5) = "RG = '" & DatDados(0).Recordset("RG") & "'"
    .Formulas(6) = "Endereco = '" & DatDados(0).Recordset("Endereco") & "'"
    .Formulas(7) = "Nro = '" & DatDados(0).Recordset("Nro") & "'"
    .Formulas(8) = "Municipio = '" & DatDados(0).Recordset("Cidade") & "'"
    .Formulas(9) = "Estado = '" & DatDados(0).Recordset("Uf") & "'"
    .Formulas(10) = "Valor = '" & Format(Valor.Text, "#,#0.00") & "'"
    .Formulas(11) = "FormaDePagamento = '" & Left(FormaPagamento.Text, 70) & "'"
    .Formulas(12) = "Bens1 = '" & Var1 & "'"
    .Formulas(13) = "Bens2 = '" & Var2 & "'"
    .Formulas(14) = "Bens3 = '" & Var3 & "'"
    .Formulas(15) = "Bens4 = '" & Var4 & "'"
    .Formulas(16) = "FormaDePagamento2 = '" & Mid$(FormaPagamento.Text, 71, 100) & "'"
    
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

Private Sub Form_Load()
Dim Componentes As Variant

Set Formulário = Me

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

End Sub
