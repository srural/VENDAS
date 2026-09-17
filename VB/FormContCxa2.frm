VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormContCxa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   13260
   Begin VB.Data DatDados 
      Caption         =   "entidade"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   3
      Left            =   6975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DatDados 
      Caption         =   "pedido"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   5805
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5850
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DatDados 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3285
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1260
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Ok 
      Caption         =   "OK"
      Height          =   285
      Left            =   630
      TabIndex        =   4
      Top             =   1530
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "FormContCxa2.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   13245
      TabIndex        =   2
      Top             =   0
      Width           =   13245
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle do Caixa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   135
         TabIndex        =   3
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.CommandButton BtExclui 
      Caption         =   "&Exclui"
      Height          =   330
      Left            =   11700
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data DatDados 
      Caption         =   "Caixa"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5130
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2970
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormContCxa2.frx":28D4
      Height          =   2355
      Left            =   180
      OleObjectBlob   =   "FormContCxa2.frx":28EE
      TabIndex        =   0
      Top             =   1935
      Width           =   11715
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   1
      Left            =   615
      TabIndex        =   5
      Top             =   675
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   48824321
      CurrentDate     =   36992
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Index           =   2
      Left            =   615
      TabIndex        =   6
      Top             =   1140
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   48824321
      CurrentDate     =   36992
   End
   Begin MSDBGrid.DBGrid DbgFecha 
      Bindings        =   "FormContCxa2.frx":3E51
      Height          =   1125
      Left            =   2160
      OleObjectBlob   =   "FormContCxa2.frx":3E6B
      TabIndex        =   7
      Top             =   675
      Width           =   3585
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FormContCxa2.frx":485B
      Height          =   2490
      Left            =   225
      OleObjectBlob   =   "FormContCxa2.frx":4875
      TabIndex        =   10
      Top             =   4545
      Width           =   12885
   End
   Begin MSMask.MaskEdBox MebPed 
      Height          =   195
      Left            =   6615
      TabIndex        =   11
      Top             =   855
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebRec 
      Height          =   195
      Left            =   6615
      TabIndex        =   12
      Top             =   1350
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebPed2 
      Height          =   195
      Left            =   7470
      TabIndex        =   13
      Top             =   855
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebRec2 
      Height          =   195
      Left            =   7470
      TabIndex        =   14
      Top             =   1350
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lbloperacao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9225
      TabIndex        =   18
      Top             =   7200
      Width           =   3885
   End
   Begin VB.Label lblnome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   270
      TabIndex        =   17
      Top             =   7155
      Width           =   5460
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Receb."
      Height          =   195
      Left            =   6030
      TabIndex        =   16
      Top             =   1350
      Width           =   525
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Pedido"
      Height          =   195
      Left            =   6030
      TabIndex        =   15
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Inicio"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   705
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Fim"
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   1140
      Width           =   240
   End
End
Attribute VB_Name = "FormContCxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtExclui_Click()
If MsgBox("Confirma Exlusão  ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(2).Recordset.RecordCount <> 0 Then
        DatDados(2).Recordset.Delete
        DatDados(2).Refresh
    End If
End If
End Sub


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub DbgFecha_Click()
Dim VarCaixa As Recordset
Dim Criterio As String


Set VarCaixa = Banco.OpenRecordset("Select * FROM  ContCxa WHERE  Maquina = " & """" & DbgFecha.Columns(1) & """" & " Order By CodControle")
VarCaixa.Requery

If VarCaixa.RecordCount > 0 Then
    Criterio = "CodControle = " & DbgFecha.Columns(0)
    VarCaixa.FindFirst Criterio
    
    If Not VarCaixa.NoMatch Then
        
        MebPed = VarCaixa!Pedido
        MebRec = VarCaixa!Rec
        VarCaixa.MoveNext
        If Not VarCaixa.Eof Then
           MebPed2 = VarCaixa!Pedido
           MebRec2 = VarCaixa!Rec
        Else
           MebPed2 = MebPed
           MebRec2 = MebRec
        End If
        
    End If

        DatDados(2).RecordSource = "SELECT * FROM CXA  WHERE NroCaixa = " & DbgFecha.Columns(0) & "  ORDER BY DATA DESC, HORA DESC"
        DatDados(2).Refresh
        Sql = "SELECT ITP.Pedido, PED.DataEmiss, ITP.Qtd, PRD.Descrição_Produto, ITP.ValorUnit, ITP.VlrVdaDia, PED.Desconto, PED.Contato FROM PRD INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE ITP.Pedido > " & MebPed & " And ITP.Pedido <= " & MebPed2 & " AND PED.Contato = " & DbgFecha.Columns(1) & " ORDER BY ITP.Pedido"
        
        DatDados(1).RecordSource = "SELECT Ped.Operacao, Ped.Entidade, ITP.Pedido, PED.DataEmiss, ITP.Qtd, PRD.Descrição_Produto, ITP.ValorUnit, ITP.VlrVdaDia, PED.Desconto, PED.Contato FROM PRD INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE ITP.Pedido > " & MebPed & " And ITP.Pedido <= " & MebPed2 & " AND PED.Contato = " & """" & DbgFecha.Columns(1) & """" & " ORDER BY ITP.Pedido"
        DatDados(1).Refresh

End If
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DatDados(1).Recordset.RecordCount > 0 Then
    DatDados(3).Recordset.FindFirst "CodEntidade = " & DatDados(1).Recordset("Entidade")
    LblNome = DatDados(3).Recordset("Nome")
    If DatDados(1).Recordset("Operacao") = 1 Then
        lbloperacao = "VISTA"
    ElseIf DatDados(1).Recordset("Operacao") = 2 Then
            lbloperacao = "PRAZO"
         ElseIf DatDados(1).Recordset("Operacao") = 4 Then
                lbloperacao = "ORCAMENTO"
         Else
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = 300 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300  ' (MDIPrincipal.Height - Me.Height) / 6

DatDados(0).RecordSource = "SELECT * FROM CONTCXA  ORDER BY CODCONTROLE DESC"
DatDados(0).Refresh


DatDados(2).RecordSource = "SELECT * FROM CXA  ORDER BY DATA DESC, HORA DESC"
DatDados(2).Refresh

DatDados(1).RecordSource = "SELECT Ped.Operacao, Ped.Entidade, ITP.Pedido, PED.DataEmiss, ITP.Qtd, PRD.Descrição_Produto, ITP.ValorUnit, ITP.VlrVdaDia, PED.Desconto, PED.Contato FROM PRD INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE (((ITP.Pedido)=1 And (ITP.Pedido)=4) AND ((PED.Contato)='aaa')) ORDER BY ITP.Pedido"
DatDados(1).Refresh

DatDados(3).RecordSource = "Select CodEntidade, Nome From Ent Order by CodEntidade"
DatDados(3).Refresh

End Sub

