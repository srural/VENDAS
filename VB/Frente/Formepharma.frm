VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Formepharma 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   28
      Top             =   7380
      Width           =   11355
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Formepharma.frx":0000
      Left            =   225
      List            =   "Formepharma.frx":000A
      TabIndex        =   27
      Text            =   "CRM"
      Top             =   1620
      Width           =   1050
   End
   Begin VB.CommandButton BtExporta 
      Caption         =   "Exporta Produto"
      Height          =   375
      Left            =   9630
      TabIndex        =   26
      Top             =   2475
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11265
      TabIndex        =   24
      Top             =   0
      Width           =   11265
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "e-Pharma"
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
         TabIndex        =   25
         Top             =   90
         Width           =   5325
      End
   End
   Begin VB.Data DatUsu 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "I:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   8955
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from UsuPop"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   270
      TabIndex        =   22
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autorização"
      Height          =   645
      Left            =   3465
      TabIndex        =   19
      Top             =   2295
      Width           =   3435
      Begin VB.Label LblAutoriza 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   20
         Top             =   225
         Width           =   3030
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
      Height          =   600
      Left            =   4320
      TabIndex        =   18
      Text            =   "     Aguarde . . ."
      Top             =   4230
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   255
      TabIndex        =   17
      Top             =   6930
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   255
      TabIndex        =   16
      Top             =   6675
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   255
      TabIndex        =   15
      Top             =   6465
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   255
      TabIndex        =   14
      Top             =   6240
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   255
      TabIndex        =   13
      Top             =   6000
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   255
      TabIndex        =   12
      Top             =   5760
      Width           =   10740
   End
   Begin MSMask.MaskEdBox MebData 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   5085
      TabIndex        =   3
      Top             =   1620
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtUf 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Tag             =   "N"
      Text            =   "SP"
      Top             =   1620
      Width           =   1275
   End
   Begin VB.TextBox TxtCrm 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "N"
      Top             =   1620
      Width           =   2040
   End
   Begin VB.TextBox TxtCartao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   225
      TabIndex        =   0
      Tag             =   "N"
      Top             =   990
      Width           =   2535
   End
   Begin VB.TextBox TxtPedido 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   9450
      TabIndex        =   8
      Tag             =   "N"
      Top             =   675
      Width           =   1680
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Pre-Autoização"
      Height          =   510
      Left            =   7110
      TabIndex        =   4
      Top             =   2430
      Width           =   1500
   End
   Begin VB.Data DatDados 
      Caption         =   "Popular"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Index           =   1
      Left            =   675
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"Formepharma.frx":0018
      Top             =   4860
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "Formepharma.frx":010E
      Height          =   2475
      Left            =   225
      OleObjectBlob   =   "Formepharma.frx":0128
      TabIndex        =   5
      Top             =   3105
      Width           =   10770
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   195
      Left            =   270
      TabIndex        =   23
      Top             =   2115
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   225
      TabIndex        =   21
      Top             =   1440
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   7740
      Picture         =   "Formepharma.frx":11E0
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   3450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Prescrição"
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
      Height          =   240
      Left            =   5085
      TabIndex        =   11
      Top             =   1305
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U.F. Prescritor"
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
      Height          =   240
      Left            =   3600
      TabIndex        =   10
      Top             =   1350
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRM do Prescritor"
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
      Height          =   240
      Left            =   1440
      TabIndex        =   9
      Top             =   1395
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
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
      Height          =   240
      Left            =   8685
      TabIndex        =   7
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero do Cartão"
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
      Height          =   240
      Left            =   225
      TabIndex        =   6
      Top             =   720
      Width           =   1650
   End
End
Attribute VB_Name = "Formepharma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Arq As Recordset
Dim VarUsuario As String
Dim VarSenha As String



Private Sub BtExporta_Click()

Dim VarPrd As Recordset
Dim sql As String
Dim VarLinha As String

If MsgBox("Confirma exportação de produtos ? ", vbYesNo, App.Title) = vbYes Then

        sql = "SELECT Prd.CodPrd, Prd.Descrição_Produto, Prd.Complemento, Prd.Venda, BAR.BarBarra FROM Prd INNER JOIN BAR ON Prd.CodPrd = BAR.Produto ORDER BY BAR.BarBarra"
        Set VarPrd = FrmFrente.Banco.OpenRecordset(sql)
        VarPrd.Requery
        If VarPrd.RecordCount > 0 Then
           VarPrd.MoveFirst
           Open "c:\e-pharma\Tabela.txt" For Output As #1
           While Not VarPrd.EOF
               
               If Len(Trim(VarPrd!BarBarra)) >= 8 Then
                    VarLinha = psCompString(3, Left(VarPrd!BarBarra, 13), 13)
                    VarLinha = VarLinha + psCompString(3, Left(VarPrd!Descrição_Produto, 40), 40)
                    VarLinha = VarLinha + psCompString(2, LimpaNro(VarPrd!Venda), 7)
                    VarLinha = VarLinha + psCompString(2, LimpaNro(VarPrd!Venda), 7)
                    VarLinha = VarLinha + psCompString(1, 0, 8)
                    Print #1, VarLinha & Chr(13)
               End If
               VarPrd.MoveNext
           Wend
           MsgBox "Arquivo exportado com sucesso!", vbInformation, App.Title
           Close #1
      End If
End If

End Sub

Private Sub BtOk_Click()
Dim VarQtd As String
Dim VarPreco As String
Dim VarPrescrita As String

TxtAguarde.Visible = True

If DatDados(1).Recordset.RecordCount > 0 Then
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.EOF
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset("Cartao") = TxtCartao
        DatDados(1).Recordset("Crm") = Left(TxtCrm, 6)
        DatDados(1).Recordset("UfPrescritor") = TxtUf
        DatDados(1).Recordset("DtPrescricao") = MebData
        DatDados(1).Recordset.Update
        DatDados(1).Recordset.MoveNext
    Wend
    
    If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
        Kill "C:\FPOPULAR\FP_REQ.txt"
    End If
    If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
        Kill "C:\FPOPULAR\FP_ENV.txt"
    End If
                    
    Open "c:\FPopular\FP_ENV.TXT" For Output As #2
    Print #2, "01" & Trim(VarUsuario) & Trim(CmbFunc)
    Print #2, "02" & Trim(VarSenha) & Trim(TxtSenha)
'    Print #2, "03" & psCompString(1, TxtPedido, 6) & psCompString(1, LIMPACampo2(TxtCnpj), 14) & Trim(psCompString(1, Left(LIMPACampo2(TxtCpf), 11), 11)) & psCompString(1, Left(TxtCrm, 6), 6) & Left(TxtUf, 2) & Format(MebData, "dd/mm/yyyy")
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.EOF
        VarQtd = psCompString(1, LimpaNumero(DatDados(1).Recordset("QtdSolicitada") * 100), 5)
        VarPreco = psCompString(1, LimpaNumero(DatDados(1).Recordset("PcoVenda") * 100), 15)
        VarPrescrita = psCompString(1, LimpaNumero(DatDados(1).Recordset("QtdPrescrita") * 100), 5)
        
        Print #2, "04" & psCompString(1, DatDados(1).Recordset("CodBarra"), 13) & VarQtd & VarPreco & VarPrescrita
        DatDados(1).Recordset.MoveNext
    Wend
    Close #2
        
    TrataPopular
                
End If
TxtAguarde.Visible = False

'Unload Me
End Sub



Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbFunc_Change()
Dim Criterio As String
If DatUsu.Recordset.RecordCount > 0 Then
    Criterio = "Cpf = " & """" & CmbFunc.Text & """"
    DatUsu.Recordset.FindFirst Criterio

    If DatUsu.Recordset.Fields("Senha") <> "" Then
        TxtSenha = DatUsu.Recordset("Senha")
    End If
End If
End Sub



'Open PortaAcbr & "\Sai.txt" For Input As #5

'Do
'    Input #5, VarTemp
'    VarAcbr = VarTemp
'    If Not EOF(5) Then
'        Input #5, VarTemp
'        If Left(VarTemp, 2) <> "OK" Or Left(VarTemp, 2) <> "ER" Then
'            VarAcbr = VarAcbr & "," & VarTemp
'        End If
'    End If
'Loop Until EOF(5)
'Close #5

'If Left(VarAcbr, 2) = "OK" Then
'    TrataAcbr = VarAcbr
'Else
'    If VarAcbr <> "," Then
'        MsgBox VarAcbr, vbInformation, App.Title
'    End If
'End If

'Kill PortaAcbr & "\Sai.txt"

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim sql As String
Dim VarItem As Recordset

MebData = Date
TxtUf = "SP"
'TxtCpf = "00106424122"
'TxtCrm = "25142"

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\dados.mdb"
Next
DatUsu.DatabaseName = FrmFrente.Caminho & "\dados.mdb"

If LerINI("Sistema", "cnpj", App.Path & "\Popular.ini") <> "" Then
    TxtCnpj = LerINI("sistema", "cnpj", App.Path & "\Popular.ini")
Else
    MsgBox "Falta C.N.P.J. do estabelecimento !", vbInformation, App.Title
End If

If LerINI("Sistema", "Login", App.Path & "\Popular.ini") <> "" Then
    VarUsuario = LerINI("sistema", "Login", App.Path & "\Popular.ini")
End If

If LerINI("Sistema", "Senha", App.Path & "\Popular.ini") <> "" Then
    VarSenha = LerINI("sistema", "Senha", App.Path & "\Popular.ini")
End If

'If Not IsNull(FrmFrente.Clientes("Cpf")) Then
'    TxtCpf.Text = LIMPACampo2(FrmFrente.Clientes("Cpf"))
'End If

If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_REQ.txt"
End If
If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_ENV.txt"
End If

TxtPedido.Text = FrmFrente.LblPedido
If TxtPedido = "" Then
    TxtPedido = 0
End If

DatDados(1).RecordSource = "SELECT Pedido, PRD.Descrição_Produto,epharma.Cartao, epharma.Crm, epharma.UfPrescritor,epharma.DtPrescricao,  epharma.CodBarra, epharma.QtdSolicitada, epharma.PcoVenda, epharma.QtdPrescrita, epharma.Codigo, epharma.PreAutoriza FROM epharma INNER JOIN PRD ON epharma.CodPrd = PRD.CodPrd WHERE (((epharma.Pedido)= " & TxtPedido & ")) ORDER BY epharma.Codigo "
DatDados(1).Refresh



If DatDados(1).Recordset.RecordCount <= 0 Then
    Set VarItem = FrmFrente.Banco.OpenRecordset("Select * From Itp INNER JOIN BAR ON ITP.Produto = BAR.Produto Where Pedido = " & TxtPedido & " AND BAR.BarBarra<>CStr([itp].[produto]) Order By CodItp")
    VarItem.Requery
    If VarItem.RecordCount > 0 Then
        VarItem.MoveFirst
        While Not VarItem.EOF
            sql = "INSERT INTO epharma ( Pedido, Cartao, Crm, UfPrescritor, DtPrescricao, CodPrd, CodBarra, QtdSolicitada, PcoVenda, QtdPrescrita ) VALUES ( "
            sql = sql & TxtPedido & " , "
            sql = sql & """" & Trim(TxtCartao) & """" & ", "
            sql = sql & """" & Trim(TxtCrm) & """" & ", "
            sql = sql & """" & Trim(TxtUf) & """" & ",# "
            sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
            sql = sql & VarItem("Itp.Produto") & " , "
            sql = sql & VarItem!BarBarra & " , "
            sql = sql & 0 & " , "
            sql = sql & Numero(VarItem!ValorUnit, 2) & " , "
            sql = sql & 0 & ")"
            FrmFrente.Banco.Execute sql
            
            VarItem.MoveNext
        Wend
        DatDados(1).Refresh
    End If
End If

'If dir(App.Path & "\VendPopular.Txt") <> "" Then
'    On Error GoTo ErroFunc
'    CmbFunc.Clear
'    Open App.Path & "\VendPopular.Txt" For Input As #1
'        Do
'            Input #1, TipoVenc
'            CmbVenc.AddItem TipoVenc
'        Loop Until TipoVenc = ""
'    Close #1
'End If
'ErroFunc:
'If Err.Number = 62 Then
'    Close #1
'End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub



Function TrataPopular() As String
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
            Case "01"
                VarPopular(1) = VarTemp
                If Mid$(VarTemp, 3, 3) = "00S" Then
                    LblAutoriza = Mid$(VarTemp, 306, 19)
                    DatDados(1).Recordset.MoveFirst
                    While Not DatDados(1).Recordset.EOF
                        DatDados(1).Recordset.Edit
                        DatDados(1).Recordset("PreAutoriza") = LblAutoriza
                        DatDados(1).Recordset.Update
                        DatDados(1).Recordset.MoveNext
                    Wend
                End If
            Case "02"
                VarPopular(2) = VarTemp
                If Mid$(VarTemp, 51, 3) = "00S" Then
                    CodBarPop = Mid$(VarTemp, 3, 13)
                    VarQtd = (Mid$(VarTemp, 16, 5) / 100)
                    ValorGov = (Mid$(VarTemp, 21, 15) / 100)
                    ValorCli = (Mid$(VarTemp, 36, 15) / 100)
                    sql = "UPDATE POPULAR SET POPULAR.ParcCli =  " & Numero(Format(ValorCli, "#.##0,00"), 2)
                    sql = sql + " , POPULAR.ParcGov = " & Numero(Format(ValorGov, "#.##0,00"), 2) & ", POPULAR.QtdAprovada = " & Numero(Format(VarQtd, "#.##0,00"), 2)
                    sql = sql + ",  POPULAR.Vendedor = " & """" & Trim(CmbFunc) & """" & ", POPULAR.Senha = " & """" & Trim(TxtSenha) & """"
                    sql = sql + " WHERE POPULAR.CodBarra = " & """" & CodBarPop & """" & " AND POPULAR.Pedido =  " & TxtPedido
                    FrmFrente.Banco.Execute sql

'                    Criterio = "CodBarra = " & CodBarPop
'                    DatDados(1).Recordset.FindFirst Criterio
'                    If Not DatDados(1).Recordset.NoMatch Then
'                        DatDados(1).Recordset.Edit
'                        DatDados(1).Recordset("QtdAprovada") = VarQtd
'                        DatDados(1).Recordset("ParcGov") = ValorGov
'                        DatDados(1).Recordset("ParcCli") = ValorCli
'                        DatDados(1).Recordset.Update
'                    End If
                End If
            Case "03"
                VarPopular(3) = VarTemp
            Case "04"
                VarPopular(4) = VarTemp
            Case "05"
                VarPopular(5) = VarTemp
            Case "06"
                VarPopular(6) = VarTemp
        End Select
    Loop Until EOF(5)

Close #5


If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_REQ.txt"
End If
If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_ENV.txt"
End If

End Function

