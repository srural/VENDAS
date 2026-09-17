VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormEntregas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Pendentes"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11550
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   180
      ScaleHeight     =   735
      ScaleWidth      =   7080
      TabIndex        =   4
      Top             =   5220
      Width           =   7080
      Begin VB.TextBox TxtQtd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5670
         TabIndex        =   7
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox TxtProduto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         TabIndex        =   5
         Top             =   315
         Width           =   5460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
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
         Left            =   5670
         TabIndex        =   8
         Top             =   45
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
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
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdConfirma 
      Caption         =   "&Confirma"
      Height          =   375
      Left            =   9315
      TabIndex        =   2
      Top             =   5535
      Width           =   1320
   End
   Begin VB.TextBox TxtLoja 
      Height          =   330
      Left            =   7515
      TabIndex        =   1
      Top             =   5535
      Width           =   1410
   End
   Begin VB.Data DatDados 
      Caption         =   "Entregas"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2430
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormEntregas.frx":0000
      Top             =   2070
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormEntregas.frx":013F
      Height          =   4875
      Left            =   135
      OleObjectBlob   =   "FormEntregas.frx":0159
      TabIndex        =   0
      Top             =   180
      Width           =   11265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loja"
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
      Left            =   7470
      TabIndex        =   3
      Top             =   5220
      Width           =   465
   End
End
Attribute VB_Name = "FormEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Soma As Recordset

Private Sub CmdConfirma_Click()
Dim Criterio As String
Dim TotEst As Currency
If MsgBox("Confirma entrega ? ", vbYesNo, App.Title) = vbYes Then
    Set Soma = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto = " & DatDados(0).Recordset.Fields("CodPrd") & " ORDER BY DescricaoLoja")
    
    If TxtQtd = "" Or Not IsNumeric(TxtQtd) Then
        Exit Sub
    End If
    
    If TxtLoja = "" Or Not IsNumeric(TxtLoja) Then
        TxtLoja = ""
        TxtLoja.SetFocus
        Exit Sub
    End If
    
    Criterio = "DescricaoLoja = " & """" & TxtLoja & """"
    Soma.FindFirst Criterio
    If Soma.NoMatch Then
        MsgBox "Loja não cadastrada !", vbExclamation, App.Title
    Else
        Soma.Edit
        Soma.Fields("Estoque") = CCur(Soma.Fields("Estoque")) - CCur(TxtQtd)
        Soma.Update
    End If
    Soma.Requery
    Soma.MoveFirst
    TotEst = 0
    While Not Soma.EOF
        TotEst = TotEst + Soma.Fields("Estoque")
        Soma.MoveNext
    Wend
    DatDados(0).Recordset.Edit
    DatDados(0).Recordset.Fields("Estoque") = TotEst
    DatDados(0).Recordset.Fields("Entregue") = 1
    DatDados(0).Recordset.Update
    TxtProduto = ""
    TxtQtd = ""
    TxtLoja = ""
    DatDados(0).Refresh
End If
End Sub

Private Sub DbGrid_dblClick()
TxtProduto.Text = DatDados(0).Recordset.Fields("Descrição_Produto")
TxtQtd.Text = Format(DatDados(0).Recordset.Fields("Qtd"), "#,##0.00")
TxtLoja.SetFocus
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Formulario As Form

Set Formulario = Me

Formulario.Left = 200
Formulario.Top = 200


For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "SELECT PED.CodPed, PED.DataEmiss, PED.Entregue, ITP.Qtd, PRD.Descrição_Produto, ENT.Nome, ENT.Endereco, ENT.Fone, PRD.CodPrd, PRD.Estoque, PED.Operacao FROM (PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.Entregue)=False) AND ((PED.Operacao)=1 OR (PED.Operacao)=2)) ORDER BY PED.DataEmiss"


DatDados(0).Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyPress = 0
    End If
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

