VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormEstoque 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9045
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   28
      Top             =   6345
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9060
      TabIndex        =   26
      Top             =   0
      Width           =   9060
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Movimentação de Estoque"
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
         TabIndex        =   27
         Top             =   45
         Width           =   5730
      End
   End
   Begin VB.CommandButton CmdCorrige 
      Caption         =   "Corrige Estoque"
      Height          =   420
      Left            =   7335
      TabIndex        =   25
      Top             =   1980
      Width           =   1545
   End
   Begin VB.ComboBox CmbPrima 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FormEstoque.frx":0000
      Left            =   3870
      List            =   "FormEstoque.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3015
      Width           =   1725
   End
   Begin VB.Data DatDados 
      Caption         =   "Depto"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2130
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Dep Order By Descricao"
      Top             =   3450
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox TxtNroSerie 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1230
      MaxLength       =   11
      TabIndex        =   2
      Top             =   3015
      Width           =   2415
   End
   Begin VB.TextBox TxtLocal 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1230
      TabIndex        =   21
      Top             =   2610
      Width           =   2415
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1260
      TabIndex        =   20
      Top             =   1710
      Width           =   1875
   End
   Begin VB.PictureBox PctBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   3510
      ScaleHeight     =   555
      ScaleWidth      =   3150
      TabIndex        =   19
      Top             =   1665
      Width           =   3180
   End
   Begin VB.TextBox TxtData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5685
      TabIndex        =   18
      Top             =   720
      Width           =   3120
   End
   Begin VB.Data DatDados 
      Caption         =   "Loja"
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   5
      Left            =   1590
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Loj ORDER BY DescricaoLoja"
      Top             =   5415
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Data DatDados 
      Caption         =   "CLoja"
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   7
      Left            =   1560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM CLoj ORDER BY Descricao"
      Top             =   4845
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox LojaDest 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   5430
      Width           =   2460
   End
   Begin VB.ComboBox DbcLoja 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4845
      Width           =   2460
   End
   Begin VB.CommandButton CmdConfirma 
      Caption         =   "&Confirma"
      Height          =   465
      Left            =   4185
      TabIndex        =   8
      Top             =   5310
      Width           =   1545
   End
   Begin MSMask.MaskEdBox MebQtd 
      Height          =   330
      Left            =   1230
      TabIndex        =   5
      Top             =   4320
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame9 
      Height          =   3210
      Left            =   6075
      TabIndex        =   11
      Top             =   2520
      Width           =   2835
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "FormEstoque.frx":0020
         Height          =   2310
         Left            =   45
         OleObjectBlob   =   "FormEstoque.frx":003A
         TabIndex        =   12
         Top             =   135
         Width           =   2700
      End
      Begin MSMask.MaskEdBox MebTotal 
         Height          =   330
         Left            =   1515
         TabIndex        =   13
         Top             =   2655
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total de Estoque"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2655
         Width           =   1215
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2550
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From PRD Order By Descrição_Produto"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormEstoque.frx":0A25
      Left            =   1230
      List            =   "FormEstoque.frx":0A32
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   2085
   End
   Begin MSDBCtls.DBCombo DbcDados 
      Bindings        =   "FormEstoque.frx":0A55
      Height          =   315
      Index           =   1
      Left            =   1230
      TabIndex        =   0
      Top             =   1275
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Descrição_Produto"
      BoundColumn     =   "CodPrd"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo DbcDados 
      Bindings        =   "FormEstoque.frx":0A6F
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   4
      Top             =   3480
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      ListField       =   "Descricao"
      BoundColumn     =   "CodDep"
      Text            =   ""
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Depto"
      Height          =   195
      Left            =   615
      TabIndex        =   24
      Top             =   3510
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NroSerie"
      Height          =   195
      Left            =   435
      TabIndex        =   23
      Top             =   3015
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      Height          =   195
      Left            =   660
      TabIndex        =   22
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      Height          =   195
      Left            =   495
      TabIndex        =   17
      Top             =   5475
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Origem"
      Height          =   195
      Left            =   540
      TabIndex        =   16
      Top             =   4875
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   4350
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   195
      Left            =   495
      TabIndex        =   10
      Top             =   1290
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   735
      TabIndex        =   9
      Top             =   2160
      Width           =   315
   End
End
Attribute VB_Name = "FormEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Formulario As Form
Public soma As Recordset
Public Barra As Recordset

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbTipo_LostFocus()
If CmbTipo.ListIndex = 2 Then
    LojaDest.Enabled = True
Else
    LojaDest.Enabled = False
End If
End Sub

Private Sub CmdConfirma_Click()
Dim Criterio As String
Dim TotEst As Currency
Dim Controle As Recordset

'On Error GoTo ErroEst

If Not IsNumeric(DbcDados(1).BoundText) Then

    MsgBox "Produto não encontrado", vbExclamation, App.Title
    DbcDados(1).Text = ""
    DbcDados(1).SetFocus
End If

If MebQtd = "" Or Not IsNumeric(MebQtd) Then
    MsgBox "Quantidade não pode ser zero", vbExclamation, App.Title
    MebQtd = ""
    MebQtd.SetFocus
    Exit Sub
End If


Criterio = "CodPrd = " & DbcDados(1).BoundText
DatDados(1).Recordset.FindFirst Criterio


If MsgBox("Confirma operação ?", vbYesNo, App.Title) = vbYes Then
        
    Criterio = "DescricaoLoja = " & """" & DbcLoja & """"
    soma.FindFirst Criterio
    If soma.NoMatch Then
        If CmbTipo.ListIndex = 0 Then
            soma.AddNew
            soma.Fields("Produto") = DbcDados(1).BoundText
            soma.Fields("DescricaoLoja") = DbcLoja
            soma.Fields("Estoque") = MebQtd
            soma.Update
            DatDados(1).Recordset.Edit
'            DatDados(1).Recordset.Fields("DtCompra") = Date
            If IsNull(DatDados(1).Recordset.Fields("Estoque")) Then
                DatDados(1).Recordset.Fields("Estoque") = CCur(MebQtd)
            Else
                DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") + CCur(MebQtd)
            End If
            DatDados(1).Recordset.Update
        Else
            MsgBox "Loja sem estoque cadastrado", vbInformation, App.Title
        End If
    Else
        If CmbTipo.ListIndex = 0 Then
            soma.Edit
            soma.Fields("Estoque") = CCur(soma.Fields("Estoque")) + CCur(MebQtd)
            soma.Update
            DatDados(1).Recordset.Edit
'            DatDados(1).Recordset.Fields("DtCompra") = Date
            DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") + CCur(MebQtd)
            DatDados(1).Recordset.Update
        Else
            soma.Edit
            soma.Fields("Estoque") = CCur(soma.Fields("Estoque")) - CCur(MebQtd)
            soma.Update
            If CmbTipo.ListIndex = 1 Then
                DatDados(1).Recordset.Edit
                DatDados(1).Recordset.Fields("DtVenda") = Date
                DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") - CCur(MebQtd)
                DatDados(1).Recordset.Update
                If DatDados(1).Recordset.Fields("BaixarComp") = True Then
                      BaixarComposicao True, 1, TxtCod, CCur(MebQtd), DbcLoja.Text
                End If
            End If
            If CmbTipo.ListIndex = 2 Then
                Criterio = "DescricaoLoja = " & """" & LojaDest & """"
                soma.FindFirst Criterio
                If Not soma.NoMatch Then
                    soma.Edit
                    soma.Fields("Estoque") = CCur(soma.Fields("Estoque")) + CCur(MebQtd)
                    soma.Update
                Else
                    soma.AddNew
                    soma.Fields("Produto") = DbcDados(1).BoundText
                    soma.Fields("DescricaoLoja") = LojaDest
                    soma.Fields("Estoque") = MebQtd
                    soma.Update
                End If
            End If
        End If
        soma.Requery
        soma.MoveFirst
        TotEst = 0
        While Not soma.Eof
            TotEst = TotEst + soma.Fields("Estoque")
            soma.MoveNext
        Wend
        Criterio = "CodPrd = " & DbcDados(1).BoundText
        DatDados(1).Recordset.FindFirst Criterio
        If DatDados(1).Recordset.Fields("Estoque") <> TotEst Then
            If MsgBox("Produto com divergência de Estoque, quer corrigir ?", vbYesNo, App.Title) = vbYes Then
                DatDados(1).Recordset.Edit
                DatDados(1).Recordset.Fields("Estoque") = TotEst
                DatDados(1).Recordset.Update
                MsgBox "Estoque corrigido !", vbInformation
            End If
        End If
        Set Controle = Banco.OpenRecordset("SELECT * FROM Est")
        
        Controle.AddNew
        Controle.Fields("Tipo") = CmbTipo.ListIndex + 1
        Controle.Fields("Produto") = DbcDados(1).BoundText
        Controle.Fields("Qtd") = MebQtd
        Controle.Fields("Loja") = DbcLoja
        Controle.Fields("Data") = Date
        Controle.Fields("Hora") = Time
        Controle.Fields("NroSerie") = TxtNroSerie.Text
        Controle.Fields("MatPrima") = CmbPrima.Text
        If IsNumeric(DbcDados(0).BoundText) Then
            Controle.Fields("Depto") = DbcDados(0).BoundText
        End If
        
        If CmbTipo.ListIndex = 2 Then
            Controle.Fields("LojaDestino") = LojaDest
        End If
        Controle.Update
    End If
End If
DbcDados(1).Text = ""
MebQtd = ""
TxtLocal = ""
TxtCod = ""
TxtNroSerie = 0
DbcDados(0).Text = ""
CmbPrima.ListIndex = 0
PctBarra.Visible = False
MebTotal.Text = 0
DatDados(5).RecordSource = "SELECT * FROM Loj WHERE Produto = 0 ORDER BY DescricaoLoja"
DatDados(5).Refresh
TxtData.Text = Date & " - " & Time
CmbTipo.SetFocus

Exit Sub
    
ErroEst:
        
        MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdCorrige_Click()
FormCorrigeEstoque.Show
End Sub

Private Sub DbcDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim CodigoBarra As String
Dim Criterio As String
Select Case Index
    Case 0
        If KeyCode = 9 Or KeyCode = 13 Then
            If IsNumeric(DbcDados(0).Text) Then
                Criterio = "CodDep = " & DbcDados(0).Text
                DatDados(0).Recordset.FindFirst Criterio
                If DatDados(0).Recordset.NoMatch Then
                    MsgBox "Departamento não encontrado !", vbInformation, App.Title
                    DbcDados(0).Text = ""
                    DbcDados(0).SetFocus
                Else
                    DbcDados(0).Text = DatDados(0).Recordset.Fields("Descricao")
                End If
            End If
        End If
    Case 1
        If KeyCode = 9 Or KeyCode = 13 Then
            CodigoBarra = ""
            If IsNumeric(DbcDados(1).Text) Then
               If Len(DbcDados(1).Text) < VarTamBarra Then
                    Criterio = "CodPrd = " & DbcDados(1)
                    DatDados(1).Recordset.FindFirst Criterio
'                    If DatDados(1).Recordset.NoMatch Then
'                        MsgBox "Produto não encontrado !!!", 16
'                        DbcDados(1).Text = ""
'                        DbcDados(1).SetFocus
'                        Exit Sub
'                    Else
                        DbcDados(1).Text = DatDados(1).Recordset.Fields("Descrição_Produto")
                        DbcDados(1).BoundText = DatDados(1).Recordset("CodPrd")
'                    End If
               Else
                    Criterio = "BarBarra = " & """" & Trim(Left(DbcDados(1).Text, 13)) & """"
                    Barra.FindFirst Criterio
                    If Barra.NoMatch Then
                        MsgBox "Produto não encontrado !!!", 16
                        DbcDados(1).Text = ""
                        DbcDados(1).SetFocus
                        Exit Sub
                    Else
                        CodigoBarra = Barra!BarBarra
                        Criterio = "CodPrd = " & Barra!Produto
                        DatDados(1).Recordset.FindFirst Criterio
                        DbcDados(1).Text = DatDados(1).Recordset.Fields("Descrição_Produto")
                        DbcDados(1).BoundText = DatDados(1).Recordset("CodPrd")
                        TxtCod = DbcDados(1).BoundText
                        TxtLocal = DatDados(1).Recordset("Classe")
                    End If
               End If
            Else
                If DbcDados(1).BoundText = "" Then
'                    If DbcDados(1).Text <> "" Then
'                        MsgBox "Produto não encontrado !!!", 16
'                        DbcDados(1).SetFocus
'                        Exit Sub
'                    End If
                Else
                    If IsNumeric(DbcDados(1).BoundText) Then
                        Criterio = "CodPrd = " & DbcDados(1).BoundText
                        DatDados(1).Recordset.FindFirst Criterio
                    End If
                End If
            End If
            
            If Not IsNumeric(DbcDados(1).BoundText) Then
                   MsgBox "Produto não encontrado !!!", 16
                   DbcDados(1).Text = ""
                   DbcDados(1).SetFocus
                   Exit Sub
            End If
            
            If IsNumeric(DbcDados(1).BoundText) Then
                TxtCod = DbcDados(1).BoundText
            Else
                TxtCod = ""
            End If
            If Not IsNull(DatDados(1).Recordset("Classe")) Then
                TxtLocal = DatDados(1).Recordset("Classe")
            Else
                TxtLocal = ""
            End If
            
            MebTotal = 0
            Criterio = "Produto = " & DbcDados(1).BoundText
            Barra.FindFirst Criterio
            If Not Barra.NoMatch Then
                CodigoBarra = Barra!BarBarra
            End If
            
            If CodigoBarra <> "" Then
                PctBarra.Visible = True
                PctBarra.ScaleMode = 3
                PctBarra.Height = PctBarra.Height * (1 * 40 / PctBarra.ScaleHeight)
                PctBarra.FontSize = 7
                Call DrawBarcode(CodigoBarra, PctBarra)
            Else
                PctBarra.Visible = False
            End If
            
            If DbcDados(1).Text <> "" And Not IsNumeric(DbcDados(1)) Then
                If IsNumeric(DbcDados(1).BoundText) Then
                    DatDados(5).RecordSource = "SELECT * FROM Loj WHERE Produto = " & DbcDados(1).BoundText & " ORDER BY DescricaoLoja"
                    DatDados(5).Refresh
                    DBGrid3.Refresh
                    Set soma = DatDados(5).Recordset.Clone
                    If soma.RecordCount <> 0 Then
                        soma.MoveFirst
                        While Not soma.Eof
                            If Not IsNull(soma.Fields("Estoque")) Then
                                MebTotal = CCur(MebTotal) + CCur(soma.Fields("Estoque"))
                            End If
                            soma.MoveNext
                        Wend
                    End If
                End If
            Else
                DatDados(5).RecordSource = "SELECT * FROM Loj WHERE Produto = 0 ORDER BY DescricaoLoja"
                DatDados(5).Refresh
            End If
            
        End If
End Select
End Sub

Private Sub DbcDados_LostFocus(Index As Integer)
DbcDados_KeyDown 1, 13, 0
End Sub

Private Sub Form_Activate()
'DbcLoja.Text = ""
TxtNroSerie.Text = 0
DbcLoja.Enabled = True
DatDados(7).RecordSource = "SELECT CLOJ.* FROM CLOJ ORDER BY CLOJ.Descricao"
DatDados(7).Refresh
If DatDados(7).Recordset.RecordCount <> 0 Then
    DatDados(7).Recordset.MoveFirst
    DbcLoja.Clear
    While Not DatDados(7).Recordset.Eof
        DbcLoja.AddItem DatDados(7).Recordset.Fields("Descricao")
        LojaDest.AddItem DatDados(7).Recordset.Fields("Descricao")
        DatDados(7).Recordset.MoveNext
    Wend
    DbcLoja.ListIndex = 0
    LojaDest.ListIndex = 0
End If
TxtData.Text = Date & " - " & Time
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

Me.Left = 400 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 500  '(MDIPrincipal.Height - Me.Height) / 6

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(5).RecordSource = "SELECT * FROM Loj WHERE Produto = 0 ORDER BY DescricaoLoja"

Set Barra = Banco.OpenRecordset("Select * From Bar Order By BarBarra")

CmbTipo.ListIndex = 0
CmbPrima.ListIndex = 0

            
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


Private Sub MebQtd_KeyPress(KeyAscii As Integer)
If CmbTipo.ListIndex = 1 And MebQtd <> "" Then
    If KeyAscii = 13 Then
        If MebQtd > CCur(MebTotal) Then
            If MsgBox("Não tem estoque suficiente, deseja continuar ?", vbYesNo, App.Title) = vbNo Then
                MebQtd = ""
                MebQtd.SetFocus
                Exit Sub
            End If
        End If
    End If
End If
End Sub
