VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormEditar 
   BorderStyle     =   0  'None
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbarra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3015
      TabIndex        =   14
      Top             =   1575
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   90
      ScaleHeight     =   555
      ScaleWidth      =   8700
      TabIndex        =   12
      Top             =   90
      Width           =   8700
      Begin VB.Label LblProduto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         TabIndex        =   13
         Top             =   45
         Width           =   7665
      End
   End
   Begin VB.TextBox txtLblItens 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton BtSalva 
      Appearance      =   0  'Flat
      Caption         =   "&Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1260
      Width           =   780
   End
   Begin MSMask.MaskEdBox MebUnitario 
      Height          =   375
      Left            =   2895
      TabIndex        =   2
      Top             =   1215
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTot 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1215
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDesc 
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Top             =   1215
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebQtd 
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Top             =   1215
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebFina 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1215
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   90
      ScaleHeight     =   330
      ScaleWidth      =   8670
      TabIndex        =   15
      Top             =   2070
      Width           =   8700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   5
      Height          =   2400
      Left            =   45
      Top             =   45
      Width           =   8790
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unitário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   2880
      TabIndex        =   10
      Top             =   945
      Width           =   690
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   4410
      TabIndex        =   9
      Top             =   945
      Width           =   465
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desc(%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   1575
      TabIndex        =   8
      Top             =   945
      Width           =   780
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   315
      TabIndex        =   7
      Top             =   945
      Width           =   315
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Financeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6165
      TabIndex        =   6
      Top             =   945
      Width           =   945
   End
End
Attribute VB_Name = "FormEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vendas As Recordset

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Criterio As String
Dim VarItem As Double
Set Vendas = FrmFrente.RecVendas.Clone
VarItem = FrmFrente.ListVda.SelectedItem.Index

Criterio = "CodItp = " & FrmFrente.ListVda.SelectedItem.SubItems(7)
Vendas.FindFirst Criterio

End Sub

Private Sub MebQtd_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10
End Sub

Private Sub MebQtd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MebDesc.SetFocus
End If
End Sub

Private Sub MebUnitario_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10
End Sub

Private Sub MebUnitario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MebFina.SetFocus
End If

End Sub

Private Sub MebUnitario_LostFocus()
If MebUnitario <> "" And MebQtd <> "" Then
    MebTot = CCur(MebQtd) * CCur(MebUnitario)
    MebFina = MebTot
Else
    MebUnitario.SetFocus
End If
End Sub

Private Sub MebDesc_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10
End Sub

Private Sub MebDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MebUnitario.SetFocus
End If
End Sub

Private Sub MebDesc_LostFocus()
If MebDesc <> "" Then
    MebUnitario = CCur(Vendas("ValorUnit")) - ((CCur(Vendas("ValorUnit")) * (CCur(MebDesc)) / 100))
Else
    MebDesc.SetFocus
End If
End Sub

Private Sub MebFina_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10

End Sub

Private Sub MebFina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtSalva.SetFocus
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub BtSalva_Click()
        
Vendas.Edit
Vendas("Qtd") = MebQtd
Vendas("ValorUnit") = MebUnitario
Vendas("Valor") = MebTot
Vendas("Desconto") = MebDesc
Vendas("VlrVdaDia") = MebFina
Vendas.Update

With FrmFrente
     .ListVda.ListItems.Remove CDec(.VarItem)

    Set .LstPesq = .ListVda.ListItems.Add(, , VarItem)

    .LstPesq.SubItems(1) = Format(MebQtd, " ##,##0.000")  'SeCampoNulo(Consulta, "Agencia")
    .LstPesq.SubItems(2) = Vendas("Produto")
    .LstPesq.SubItems(3) = TxtBarra
    .LstPesq.SubItems(4) = LblProduto   'SeCampoNulo(Consulta, "Agencia")
    .LstPesq.SubItems(5) = Numero(Format(MebUnitario, " ##,##0.00"), 2)  'SeCampoNulo(Consulta, "Conta")
    .LstPesq.SubItems(6) = Numero(Format(MebFina, "##,##0.00"), 2)  'SeCampoNulo(Consulta, "DataOper")
    .LstPesq.SubItems(7) = txtLblItens
    
    'FrmFrente.DatCompra.Refresh
    'FrmFrente.DatCompra.Recordset.MoveLast
    
    .soma = CCur(FrmFrente.soma) + CCur(MebFina)
    .MebTotal = Format(FrmFrente.soma, "##,##0.00")
End With

Unload Me
FrmFrente.TxtDados(0).SetFocus
End Sub

