VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormVista 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   -450
      ScaleHeight     =   360
      ScaleWidth      =   11355
      TabIndex        =   21
      Top             =   8190
      Width           =   11355
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   405
      TabIndex        =   16
      Top             =   6660
      Width           =   5820
      Begin VB.Label LblDoc 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   945
         Width           =   5415
      End
      Begin VB.Label LblCidade 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   705
         Width           =   5415
      End
      Begin VB.Label LblEndereco 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   465
         Width           =   5415
      End
      Begin VB.Label LblNome 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   17
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6630
      TabIndex        =   14
      Top             =   0
      Width           =   6630
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Operação"
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
         TabIndex        =   15
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   405
      ScaleHeight     =   3405
      ScaleWidth      =   5820
      TabIndex        =   8
      Top             =   615
      Width           =   5820
      Begin VB.CommandButton BtOk 
         Height          =   420
         Left            =   4440
         Picture         =   "FormVista.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Gravar"
         Top             =   2880
         Width           =   435
      End
      Begin VB.ListBox ListaOper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1470
         ItemData        =   "FormVista.frx":0532
         Left            =   1770
         List            =   "FormVista.frx":0542
         TabIndex        =   1
         Top             =   15
         Width           =   2265
      End
      Begin VB.Data DatVend 
         Caption         =   "Vendedor"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2655
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodEntidade,Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
         Top             =   2205
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSMask.MaskEdBox MebDataVenc 
         Height          =   405
         Left            =   2460
         TabIndex        =   2
         Top             =   1635
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo DbcVendedor 
         Bindings        =   "FormVista.frx":056C
         Height          =   360
         Left            =   1605
         TabIndex        =   3
         Top             =   2310
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   16711680
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
      Begin MSMask.MaskEdBox MaskFone 
         Height          =   405
         Left            =   1635
         TabIndex        =   23
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
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
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   600
         TabIndex        =   24
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natureza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   375
         TabIndex        =   13
         Top             =   -30
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   330
         TabIndex        =   12
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   285
         TabIndex        =   11
         Top             =   2280
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   225
         TabIndex        =   10
         Top             =   1665
         Width           =   2100
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Está natureza define o valor dos produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   4335
         TabIndex        =   9
         Top             =   315
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   405
      ScaleHeight     =   2580
      ScaleWidth      =   5820
      TabIndex        =   0
      Top             =   4050
      Width           =   5820
      Begin VB.TextBox TxtPesAut 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   345
         Width           =   5415
      End
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         Height          =   1050
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1470
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoas Autrorizada a comprar no nome do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   4320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs do Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1245
         Width           =   1260
      End
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   450
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   855
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      BorderWidth     =   5
      Height          =   7800
      Left            =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "FormVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Contador As Byte

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtOk_Click()
FrmFrente.Clientes.Edit
FrmFrente.Clientes!Fone = MaskFone.Text
FrmFrente.Clientes.Update

End Sub

Private Sub DbcVendedor_GotFocus()
DbcVendedor.Text = FrmFrente.DbcVendedor.Text

'If FormSenha = True Then
'    DbcVendedor.Enabled = False
'End If
End Sub

Private Sub DbcVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Criterio As String


If KeyCode = 13 Then

    If IsNumeric(DbcVendedor) Then
        Criterio = "CodEntidade = " & DbcVendedor
        DatVend.Recordset.FindFirst Criterio
        If DatVend.Recordset.NoMatch Then
            MsgBox "Vendedor não cadastrado !", vbInformation, App.Title
            DbcVendedor.Text = ""
            DbcVendedor.SetFocus
        Else
            DbcVendedor.Text = DatVend.Recordset("Nome")
            FrmFrente.DbcVendedor.Text = DatVend.Recordset("Nome")
            Unload Me
        End If
    Else
        If IsNumeric(DbcVendedor.BoundText) Then
            Criterio = "CodEntidade = " & DbcVendedor.BoundText
            DatVend.Recordset.FindFirst Criterio
            If DatVend.Recordset.NoMatch Then
                MsgBox "Vendedor não cadastrado !", vbInformation, App.Title
                DbcVendedor.Text = ""
                DbcVendedor.SetFocus
            Else
                DbcVendedor.Text = DatVend.Recordset("Nome")
                FrmFrente.DbcVendedor.Text = DatVend.Recordset("Nome")
                Unload Me
            End If
        End If
    End If

End If

End Sub

Private Sub Form_Activate()

'VarDtVenc = Date


MebDataVenc = CDate(Format(VarDtVenc, "dd/mm/yyyy"))

DatVend.DatabaseName = FrmFrente.Caminho & "\Dados.Mdb"
DatVend.RecordSource = "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
DatVend.Refresh

With FrmFrente.Clientes

    LblEndereco.Caption = Trim(.Fields("Endereco")) & " ," & .Fields("Nro")
    LblCidade.Caption = Trim(.Fields("Cidade")) & "-" & .Fields("Uf")
    LblNome.Caption = psCompString(1, .Fields("CodEntidade"), 6) & " - " & .Fields("Nome")
    If .Fields("Classificacao") = 2 Then
        If .Fields("CGC") <> "" Then
            LblDoc.Caption = Format(.Fields("CGC"), "@@.@@@.@@@/@@@@-@@") & "   |   " & .Fields("InscrEst")
        Else
            LblDoc.Caption = Format(.Fields("Cpf"), "@@@.@@@.@@@-@@") & "   |   " & .Fields("InscrEst")
        End If
    ElseIf .Fields("Classificacao") = 1 Then
        LblDoc.Caption = Format(.Fields("Cpf"), "@@@.@@@.@@@-@@") & "   |   " & .Fields("InscrEst")
    End If
    If Not IsNull(.Fields("Fone")) Then
         MaskFone = .Fields("Fone")
    End If
End With

'ListaOper.SetFocus

If FrmFrente.TxtDados(0) = 1 Then
    ListaOper.ListIndex = 2
Else
    ListaOper.ListIndex = 0
End If
'Text1.SetFocus

End Sub

Private Sub Form_Load()
If CfgDesc = "Nao" Then
    MebDataVenc.Enabled = False
Else
    MebDataVenc.Enabled = True
End If

If VarTelaPreco = True Then
    ListaOper.Enabled = True
Else
    ListaOper.Enabled = False
    If FrmFrente.CodCli = 1 Then
        Vista = True
    Else
        Vista = False
    End If
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
VarDtVenc = MebDataVenc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
           Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub

Private Sub ListaOper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If ListaOper.ListIndex < 1 Then
        Vista = False
    Else
        Vista = True
    End If
    NatOper = ListaOper.ListIndex
'    If MebDataVenc.Enabled Then
'        MebDataVenc.SetFocus
'        contador = 2
    Else
'        contador = 2
'        DbcVendedor.Text = FrmFrente.DbcVendedor.Text
'        DbcVendedor.SetFocus
'    End If

End If

End Sub

Private Sub ListaOper_KeyPress(KeyAscii As Integer)


If KeyAscii = 27 Then
    Unload Me
End If

End Sub
