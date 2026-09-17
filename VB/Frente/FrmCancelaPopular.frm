VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCancelaPopular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancela Popular"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtCancelaTotal 
      Caption         =   "Cancela Venda Total"
      Height          =   375
      Left            =   9090
      TabIndex        =   18
      Top             =   5355
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -135
      ScaleHeight     =   330
      ScaleWidth      =   11505
      TabIndex        =   17
      Top             =   7650
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
      ScaleWidth      =   11400
      TabIndex        =   15
      Top             =   0
      Width           =   11400
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancela Popular"
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
         Width           =   5325
      End
   End
   Begin VB.TextBox TxtAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   4365
      TabIndex        =   14
      Text            =   "     Aguarde . . ."
      Top             =   2610
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton BtCancela 
      Caption         =   "Cancela Venda"
      Height          =   375
      Left            =   7290
      TabIndex        =   13
      Top             =   5355
      Width           =   1500
   End
   Begin VB.TextBox TxtQtd 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6030
      TabIndex        =   11
      Top             =   5400
      Width           =   1050
   End
   Begin VB.TextBox TxtBarra 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3465
      TabIndex        =   9
      Top             =   5400
      Width           =   2265
   End
   Begin VB.TextBox TxtAutorizacao 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   270
      TabIndex        =   7
      Top             =   5400
      Width           =   3030
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   5895
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   5
      Top             =   6180
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   270
      TabIndex        =   4
      Top             =   6360
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   270
      TabIndex        =   3
      Top             =   6600
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   270
      TabIndex        =   2
      Top             =   6870
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   270
      TabIndex        =   1
      Top             =   7110
      Width           =   10740
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FrmCancelaPopular.frx":0000
      Height          =   4245
      Left            =   270
      OleObjectBlob   =   "FrmCancelaPopular.frx":0015
      TabIndex        =   0
      Top             =   720
      Width           =   10860
   End
   Begin VB.Data DatPop 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\DADOS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FrmCancelaPopular.frx":1747
      Top             =   4950
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6030
      TabIndex        =   12
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Código Barra"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3465
      TabIndex        =   10
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pré-Autorização"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   5220
      Width           =   1275
   End
End
Attribute VB_Name = "FrmCancelaPopular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VarUsuario As String
Dim VarSenha As String
Dim VarCnpj As String

Private Sub BtCancela_Click()
Dim VarChave As String
If MsgBox("Confirma Cancelamento da Autorização: " & TxtAutorizacao, vbYesNo, App.Title) = vbYes Then
    TxtAguarde.Visible = True
    If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
        Kill "C:\FPOPULAR\FP_REQ.txt"
    End If
    If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
        Kill "C:\FPOPULAR\FP_ENV.txt"
    End If
    
    If LerINI("Sistema", "Chave", App.Path & "\Popular.ini") <> "" Then
        VarChave = LerINI("sistema", "Chave", App.Path & "\Popular.ini")
    End If
    
                    
    Open "c:\FPopular\FP_ENV.TXT" For Output As #2
    Print #2, "01" & Trim(VarUsuario) & Trim(DBGrid.Columns(8)) & Trim(VarChave)
    Print #2, "02" & Trim(VarSenha) & Trim(DBGrid.Columns(9))
    Print #2, "06" & psCompString(1, LIMPACampo2(VarCnpj), 14) & Trim(psCompString(1, TxtAutorizacao, 19))
    Print #2, "07" & psCompString(1, Trim(TxtBarra), 13) & psCompString(1, LimpaNumero(TxtQtd * 100), 5)
    Close #2
    TrataPopular2
    TxtAguarde.Visible = False
End If
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub BtCancelaTotal_Click()
If MsgBox("Confirma Cancelamento do Pedido: " & DBGrid.Columns(0), vbYesNo, App.Title) = vbYes Then
    Dim SqlPopular As Recordset
    Set SqlPopular = FrmFrente.Banco.OpenRecordset("SELECT Vendedor, Senha, POPULAR.Pedido, POPULAR.Cpf, POPULAR.Crm, POPULAR.DtPrescricao, POPULAR.CodBarra, POPULAR.QtdSolicitada, POPULAR.QtdAprovada, POPULAR.PreAutoriza FROM POPULAR WHERE Popular.Pedido = " & DBGrid.Columns(0) & "  ORDER BY POPULAR.Pedido")
    SqlPopular.Requery
    SqlPopular.MoveLast
    SqlPopular.MoveFirst
    If SqlPopular.RecordCount > 1 Then
    
        TxtAguarde.Visible = True
        If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
            Kill "C:\FPOPULAR\FP_REQ.txt"
        End If
        If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
            Kill "C:\FPOPULAR\FP_ENV.txt"
        End If
        
        If LerINI("Sistema", "Chave", App.Path & "\Popular.ini") <> "" Then
            VarChave = LerINI("sistema", "Chave", App.Path & "\Popular.ini")
        End If
        
                        
        Open "c:\FPopular\FP_ENV.TXT" For Output As #2
        Print #2, "01" & Trim(VarUsuario) & Trim(DBGrid.Columns(8)) & VarChave
        Print #2, "02" & Trim(VarSenha) & Trim(DBGrid.Columns(9))
        Print #2, "06" & psCompString(1, LIMPACampo2(VarCnpj), 14) & Trim(psCompString(1, TxtAutorizacao, 19))
        SqlPopular.MoveFirst
        While Not SqlPopular.EOF
            Print #2, "07" & psCompString(1, Trim(SqlPopular!CodBarra), 13) & psCompString(1, LimpaNumero(SqlPopular!QtdAprovada * 100), 5)
            SqlPopular.MoveNext
        Wend
        Close #2
    End If
    TrataPopular2
    TxtAguarde.Visible = False
End If
End Sub

Private Sub DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
TxtAutorizacao.Text = DBGrid.Columns(7)
TxtBarra.Text = DBGrid.Columns(4)
TxtQtd.Text = DBGrid.Columns(6)
End Sub

Private Sub Form_Load()
DatPop.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"

If LerINI("Sistema", "cnpj", App.Path & "\Popular.ini") <> "" Then
    VarCnpj = LerINI("sistema", "cnpj", App.Path & "\Popular.ini")
Else
    MsgBox "Falta C.N.P.J. do estabelecimento !", vbInformation, App.Title
End If

If LerINI("Sistema", "Login", App.Path & "\Popular.ini") <> "" Then
    VarUsuario = LerINI("sistema", "Login", App.Path & "\Popular.ini")
End If

If LerINI("Sistema", "Senha", App.Path & "\Popular.ini") <> "" Then
    VarSenha = LerINI("sistema", "Senha", App.Path & "\Popular.ini")
End If


End Sub

Function TrataPopular2() As String
'Dim VarPopular(1 To 6) As String
Dim VarTemp As String
Dim sql As String
Dim x As Double
Dim Var
Dim VarQtd As Currency
Dim ValorGov As Currency
Dim ValorCli As Currency
Dim CodBarPop As String
Dim Criterio As String

'On Error Resume Next
DoEvents
x = 0

For x = 1 To 6
    VarPopular(x) = ""
Next

While dir("C:\FPOPULAR\FP_REQ.txt") = ""
 '   DoEvents
Wend

MsgBox "Retorno recebido !", vbInformation

x = 0
LblAutoriza = ""

'For X = 1 To 700000
'   DoEvents
'Next

Open "C:\FPOPULAR\FP_REQ.txt" For Input As #5

    Do
        Input #5, VarTemp
        Select Case Left(VarTemp, 2)
            Case "05"
                VarPopular(1) = VarTemp
            Case "06"
                VarPopular(2) = VarTemp
        End Select
            
    Loop Until EOF(5)
    If Mid$(VarPopular(1), 3, 3) = "00E" Then
            sql = "INSERT INTO ESTORNO ( PreAutorizacao, AutorizaCancela, Descricao, CodigoBarra, DescrEstorno ) VALUES ( "
            sql = sql & """" & TxtAutorizacao & """" & " , "
            sql = sql & """" & Mid$(VarPopular(1), 206, 19) & """" & ", "
            sql = sql & """" & Mid$(VarPopular(1), 6, 200) & """" & ", "
            sql = sql & """" & Trim(TxtBarra) & """" & ", "
            sql = sql & """" & Mid$(VarPopular(1), 16, 200) & """" & ")"
            FrmFrente.Banco.Execute sql
            MsgBox "Medicamento cancelado com sucesso !", vbInformation
    End If
Close #5


If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_REQ.txt"
End If
If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_ENV.txt"
End If

End Function


