VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormIntegra 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7455
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   4
      Top             =   3150
      Width           =   11535
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   330
      Left            =   585
      TabIndex        =   1
      Top             =   2070
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7485
      TabIndex        =   2
      Top             =   0
      Width           =   7485
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Integração site e-comerce"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   3
         Top             =   45
         Width           =   5010
      End
   End
   Begin VB.CommandButton CmbOk 
      Caption         =   "&Inicia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1095
      Width           =   1500
   End
End
Attribute VB_Name = "FormIntegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnn As ADODB.Connection
Public Rs As ADODB.Recordset
Public VarIntegra As ADODB.Recordset
Public VarEmpresa As String

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbOk_Click()
Dim Ped As Double
Dim Itp As Double
Dim PedAnt As Double
Dim Criterio As String
Dim VarPrd As Recordset
Dim VarGrade As Recordset
Dim SQL As String

'On Error GoTo TrataErro

CmbOk.Caption = "Aguarde ..."
CmbOk.Enabled = False

Set VarPrd = Banco.OpenRecordset("SELECT PRD.CodPrd, FAM.Descrição_Familia, GRU.Descrição_Grupo, PRD.Marca, PRD.Descrição_Produto, PRD.Complemento, PRD.Custo, PRD.Venda, PRD.PesoBruto, PRD.Estoque, PRD.Obs, PRD.Embalagem, PRD.Concorrencia FROM FAM INNER JOIN (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) ON FAM.CodFam = GRU.Familia WHERE (((PRD.Concorrencia)=True)) ORDER BY PRD.CodPrd")
VarPrd.Requery

If VarPrd.RecordCount > 0 Then
    VarPrd.MoveLast
    DoEvents
    Bar1.Min = 0
    Bar1.Max = VarPrd.RecordCount
    VarPrd.MoveFirst
    SQL = "Delete from produtosSincronizacao WHERE codempresa = " & VarEmpresa
    cnn.Execute SQL
    While Not VarPrd.Eof
'        Criterio = "CodProduto = " & VarPrd!CodPrd
'        Rs.Find Criterio, , , 0
'        If Rs.Eof = False Then
'            Rs.EditMode
'            Rs!codEmpresa = VarEmpresa
'            Rs!codProduto = VarPrd!CodPrd
'            Rs!descricaoDepartamento = VarPrd!Descrição_Familia
'            Rs!descricaoCategoria = VarPrd!descrição_Grupo
'            Rs!descricaoMarca = VarPrd!Marca
'            Rs!descricaoProduto = VarPrd!Descrição_Produto
'            Rs!resumoProduto = VarPrd!Complemento
'            Rs!custoProduto = VarPrd!Custo
'            Rs!valorProduto = VarPrd!Venda
'            Rs!pesoProduto = VarPrd!PesoBruto
'            Rs!estoqueProduto = VarPrd!estoque
'            Rs!caracteristicasProduto = Left(VarPrd!Obs, 300)
'            Rs!descricaoUnidade = VarPrd!Embalagem
'            Rs.Update
'        Else
            Rs.AddNew
            Rs!codEmpresa = VarEmpresa
            Rs!codProduto = VarPrd!CodPrd
            Rs!descricaoDepartamento = VarPrd!Descrição_Familia
            Rs!descricaoCategoria = VarPrd!descrição_Grupo
            Rs!descricaoMarca = VarPrd!Marca
            Rs!descricaoProduto = VarPrd!Descrição_Produto
            Rs!resumoProduto = VarPrd!Complemento
            Rs!custoProduto = VarPrd!Custo
            Rs!valorProduto = VarPrd!Venda
            Rs!pesoProduto = VarPrd!PesoBruto
            Rs!estoqueProduto = VarPrd!estoque
            Rs!caracteristicasProduto = VarPrd!Obs
            Rs!descricaoUnidade = VarPrd!Embalagem
            Rs.Update
'        End If
        
        
        Bar1.value = VarPrd.AbsolutePosition
        VarPrd.MoveNext
    Wend
    
    Bar1.value = 0
    
    Set VarGrade = Banco.OpenRecordset("SELECT LOJ.Produto, LOJ.DescricaoLoja, LOJ.Estoque, LOJ.CodBarLoja, PRD.Concorrencia FROM PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto WHERE (((PRD.Concorrencia)=True)) ORDER BY LOJ.Produto")
    VarGrade.Requery
    
    
    VarGrade.MoveLast
    DoEvents
    Bar1.Min = 0
    Bar1.Max = VarGrade.RecordCount
    VarGrade.MoveFirst
    SQL = "Delete from produtosgrades WHERE codempresa = " & VarEmpresa
    cnn.Execute SQL
    
    While Not VarGrade.Eof
'        Criterio = "CodProduto = " & VarGrade!Produto
'        VarIntegra.Find Criterio, , , 0
'        If VarIntegra.Eof = False Then
'            VarIntegra!codEmpresa = VarEmpresa
'            VarIntegra!codProduto = VarGrade!Produto
'            VarIntegra!descricao = VarGrade!Descricaoloja
            
'            Rs!descricaoMarca = VarPrd!Marca
'            Rs!descricaoProduto = VarPrd!Descrição_Produto
'            VarIntegra.Update
'        Else
            VarIntegra.AddNew
            VarIntegra!codEmpresa = VarEmpresa
            VarIntegra!codProduto = VarGrade!Produto
            VarIntegra!descricao = VarGrade!Descricaoloja
            VarIntegra.Update
'        End If
        
        
        Bar1.value = VarGrade.AbsolutePosition
        VarGrade.MoveNext
    Wend
    
    Bar1.value = 0
    
    MsgBox "Operação concluída com sucesso !", vbInformation, App.Title
Else
    MsgBox "Arquivo sem movimento !", vbInformation, App.Title
End If
CmbOk.Caption = "Inicia"
CmbOk.Enabled = True

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub


Private Sub Form_Activate()
Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6
End Sub

Private Sub Form_Load()


'DtData.Value = Date


Set cnn = New ADODB.Connection


cnn.ConnectionString = "DSN=automasis;uid=sinc;pwd=sinc123;"
cnn.Open


If LerINI("Config", "Site", App.Path & "\config.ini") <> "" Then
    VarEmpresa = LerINI("Config", "Site", App.Path & "\config.ini")
Else
    MsgBox "Empresa não configura no config.ini !!", vbInformation
    Exit Sub
End If


'Set AdoAutomasis = New Recordset

Set Rs = New ADODB.Recordset

'Set Rs = cnn.Execute("select * from produtossincronizacao where codempresa=" & VarEmpresa)  ' , adOpenDynamic, adLockOptimistic, adCmdText)      '  adOpenStatic, adLockOptimistic

Rs.Open "select * from produtossincronizacao where codempresa=" & VarEmpresa, cnn, adOpenDynamic, adLockOptimistic

Set VarIntegra = New ADODB.Recordset

VarIntegra.Open "select * from produtosgrades where codempresa=" & VarEmpresa, cnn, adOpenDynamic, adLockOptimistic

' Senha.Open "Select * From Pwd Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText


'    StringConexao = "PROVIDER=MSDASQL;dsn=automasis;uid=;pwd=;"
        
        
'        "Provider=SQLNCLI; " & _
'                  "Initial Catalog=bdautomasis; " & _
'                  "Data Source=mssql.automasis.com.br; " & _
'                  "integrated security=SSPI; persist security info=True;"'

        
'        LocalCursor = adUseClient
'        If BcoAutomasis.State = 1 Then Banco.Close
'        BcoAutomasis.CursorLocation = adUseClient
'        BcoAutomasis.ConnectionString = StringConexao
'        BcoAutomasis.CommandTimeout = 0
'        BcoAutomasis.Open

End Sub
