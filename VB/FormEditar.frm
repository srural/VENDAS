VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormEditar 
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRef 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3960
      TabIndex        =   7
      Top             =   1890
      Width           =   2940
   End
   Begin VB.TextBox txtComplemento 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   1890
      Width           =   3570
   End
   Begin VB.TextBox txtLblItens 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7335
      TabIndex        =   14
      Top             =   1395
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton BtSalva 
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1890
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MebUnitario 
      Height          =   330
      Left            =   2745
      TabIndex        =   2
      Top             =   1410
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
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
      Height          =   330
      Left            =   4320
      TabIndex        =   3
      Top             =   1410
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
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
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   1410
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
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
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   1410
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
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
      Height          =   330
      Left            =   5895
      TabIndex        =   4
      Top             =   1410
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   8925
      TabIndex        =   15
      Top             =   45
      Width           =   8925
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Altera Itens Pedido"
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
         Width           =   3795
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   5
      Height          =   2850
      Left            =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label LblProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   180
      TabIndex        =   13
      Top             =   690
      Width           =   8430
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unitário"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2745
      TabIndex        =   12
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   1185
      Width           =   360
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desc(%)"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   1395
      TabIndex        =   10
      Top             =   1185
      Width           =   630
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1185
      Width           =   255
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Financeiro"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   5895
      TabIndex        =   8
      Top             =   1185
      Width           =   735
   End
End
Attribute VB_Name = "FormEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vendas As Recordset

Private Sub BtSair_Click()

End Sub

Private Sub Form_Activate()
Dim Criterio As String
Dim VarItem As Double
Set Vendas = FrmPedidos.DatDados(2).Recordset.Clone
'VarItem = FrmFrente.ListVda.SelectedItem.Index

Criterio = "CodItp = " & txtLblItens
Vendas.FindFirst Criterio
MebDesc = 0
If txtComplemento <> "" Then
    txtComplemento.Enabled = True
Else
    txtComplemento.Enabled = False
End If

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

Private Sub MebQtd_LostFocus()
MebUnitario_LostFocus
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
'Vendas("Itp.Entregue") = 0
Vendas("Complemento") = txtComplemento
Vendas("Referencia") = TxtRef
Vendas.Update

'With FrmFrente
'     .ListVda.ListItems.Remove CDec(.VarItem)

'    Set .LstPesq = .ListVda.ListItems.Add(, , VarItem)
    
'    .LstPesq.SubItems(1) = Format(MebQtd, " ##,##0.000")  'SeCampoNulo(Consulta, "Agencia")
'    .LstPesq.SubItems(2) = LblProduto   'SeCampoNulo(Consulta, "Agencia")
'    .LstPesq.SubItems(3) = Numero(Format(MebUnitario, " ##,##0.000"), 3)  'SeCampoNulo(Consulta, "Conta")
'    .LstPesq.SubItems(4) = Numero(Format(MebFina, "##,##0.000"), 3)  'SeCampoNulo(Consulta, "DataOper")
'    .LstPesq.SubItems(5) = txtLblItens
    
    
'    .soma = CCur(FrmFrente.soma) + CCur(MebFina)
'    .MebTotal = FrmFrente.soma
'End With

Unload Me
FrmPedidos.MebDados(4).SetFocus
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MebQtd.SetFocus
End If
End Sub
