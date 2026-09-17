VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormVidaLink 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   7920
      TabIndex        =   26
      Top             =   1935
      Width           =   2310
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   25
      Top             =   7515
      Width           =   11355
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
      TabIndex        =   23
      Top             =   0
      Width           =   11265
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "VidaLink"
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
         TabIndex        =   24
         Top             =   90
         Width           =   5325
      End
   End
   Begin MSDBCtls.DBCombo CmbFunc 
      Bindings        =   "FormVidaLink.frx":0000
      DataSource      =   "DatUsu"
      Height          =   315
      Left            =   3285
      TabIndex        =   22
      Top             =   1035
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Cpf"
      Text            =   ""
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autorização"
      Height          =   645
      Left            =   3600
      TabIndex        =   19
      Top             =   2430
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
      Height          =   645
      Left            =   4230
      TabIndex        =   18
      Text            =   "     Aguarde . . ."
      Top             =   4095
      Visible         =   0   'False
      Width           =   2940
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
      Height          =   375
      Left            =   4095
      TabIndex        =   2
      Top             =   1665
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtUf 
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
      Left            =   3285
      TabIndex        =   1
      Tag             =   "N"
      Text            =   "SP"
      Top             =   1665
      Width           =   645
   End
   Begin VB.TextBox TxtCpf 
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
      Left            =   315
      TabIndex        =   0
      Tag             =   "N"
      Top             =   1710
      Width           =   2535
   End
   Begin VB.TextBox TxtCnpj 
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
      Left            =   315
      TabIndex        =   9
      Tag             =   "N"
      Top             =   1035
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
      Left            =   9225
      TabIndex        =   8
      Tag             =   "N"
      Top             =   855
      Width           =   1680
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Pre-Autoização"
      Height          =   510
      Left            =   495
      TabIndex        =   3
      Top             =   2385
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
      RecordSource    =   $"FormVidaLink.frx":0015
      Top             =   4860
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormVidaLink.frx":010B
      Height          =   2475
      Left            =   225
      OleObjectBlob   =   "FormVidaLink.frx":0125
      TabIndex        =   4
      Top             =   3105
      Width           =   10770
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionario"
      Height          =   195
      Left            =   3285
      TabIndex        =   21
      Top             =   810
      Width           =   825
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
      Left            =   4095
      TabIndex        =   11
      Top             =   1395
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U.F."
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
      Left            =   3285
      TabIndex        =   10
      Top             =   1395
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J. do Estabelecimento"
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
      Left            =   315
      TabIndex        =   7
      Top             =   810
      Width           =   2550
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
      Left            =   9225
      TabIndex        =   6
      Top             =   630
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.F. do Participante"
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
      Left            =   315
      TabIndex        =   5
      Top             =   1440
      Width           =   1920
   End
End
Attribute VB_Name = "FormVidaLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Arq As Recordset
Dim VarUsuario As String
Dim VarSenha As String
Public VarEnvio As String
Public VarResposta As String


Private Sub BtOk_Click()
Dim VarQtd As String
Dim VarPreco As String
Dim VarPrescrita As String

TxtAguarde.Visible = True

If DatDados(1).Recordset.RecordCount > 0 Then
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.EOF
        DatDados(1).Recordset.Edit
        DatDados(1).Recordset("CPF") = TxtCpf
        If TxtCrm <> "" Then
            DatDados(1).Recordset("Crm") = Left(TxtCrm, 6)
        Else
            DatDados(1).Recordset("Crm") = Left(TxtRms, 10)
        End If
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
    Print #2, "03" & psCompString(1, TxtPedido, 6) & psCompString(1, LIMPACampo2(TxtCnpj), 14) & Trim(psCompString(1, Left(LIMPACampo2(TxtCpf), 11), 11)) & psCompString(2, Left(TxtCrm, 6), 6) & Left(TxtUf, 2) & Format(MebData, "dd/mm/yyyy") & psCompString(2, Left(TxtRms, 10), 10)
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

If LerINI("VidaLink", "Envio", App.Path & "\Config.ini") <> "" Then
    VarEnvio = LerINI("VidaLink", "Envio", App.Path & "\Popular.ini")
Else
    MsgBox "Falta configurar caminho VidaLink", vbInformation, App.Title
End If

If LerINI("VidaLink", "Resposta", App.Path & "\Config.ini") <> "" Then
    VarResposta = LerINI("VidaLink", "Resposta", App.Path & "\Popular.ini")
Else
    MsgBox "Falta configurar caminho VidaLink", vbInformation, App.Title
End If


Open VarEnvio & "\TesteDeAtividade.txt" For Output As #2
Print #2, "000111"
Close #2



TxtPedido.Text = FrmFrente.LblPedido

'DatDados(1).RecordSource = "SELECT Pedido, PRD.Descrição_Produto,POPULAR.Cpf, POPULAR.Crm, POPULAR.UfPrescritor,POPULAR.DtPrescricao,  POPULAR.CodBarra, POPULAR.QtdSolicitada, POPULAR.PcoVenda, POPULAR.QtdPrescrita, POPULAR.Codigo, Popular.PreAutoriza FROM POPULAR INNER JOIN PRD ON POPULAR.CodPrd = PRD.CodPrd WHERE (((POPULAR.Pedido)= " & TxtPedido & ")) ORDER BY POPULAR.Codigo "
'DatDados(1).Refresh

DatUsu.RecordSource = "Select * From UsuPop "
DatUsu.Refresh


If DatDados(1).Recordset.RecordCount <= 0 Then
    Set VarItem = FrmFrente.Banco.OpenRecordset("Select * From Itp INNER JOIN BAR ON ITP.Produto = BAR.Produto Where Pedido = " & TxtPedido & " AND BAR.BarBarra <> CStr([itp].[produto]) Order By CodItp")
    SALQ = "Select * From Itp INNER JOIN BAR ON ITP.Produto = BAR.Produto Where Pedido = " & TxtPedido & " AND BAR.BarBarra<>CStr([itp].[produto]) Order By CodItp"
    VarItem.Requery
    If VarItem.RecordCount > 0 Then
        VarItem.MoveFirst
        While Not VarItem.EOF
            sql = "INSERT INTO POPULAR ( Pedido, Cpf, Crm, UfPrescritor, DtPrescricao, CodPrd, CodBarra, QtdSolicitada, PcoVenda, QtdPrescrita ) VALUES ( "
            sql = sql & TxtPedido & " , "
            sql = sql & """" & Trim(TxtCpf) & """" & ", "
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
Else

    Open App.Path & "\Popular.txt" For Append As #9
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.EOF
        Print #9, "Pedido; " & DatDados(1).Recordset.Fields("Pedido") & "; " & DatDados(1).Recordset.Fields("Cpf") & "; " & DatDados(1).Recordset.Fields("CodBarra") & "; " & DatDados(1).Recordset.Fields("PreAutoriza")
        DatDados(1).Recordset.MoveNext
    Wend
    Close #9

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


Private Sub TxtCpf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtCpf <> "" Then
        If IsNumeric(LIMPACampo2(TxtCpf)) Then
            If Len(Trim(LIMPACampo2(TxtCpf))) = 11 Then
                TxtCpf = Format(Trim(LIMPACampo2(TxtCpf)), "@@@.@@@.@@@-@@")
                If Not CalculaDacCPF(TxtCpf) Then
                    MsgBox "CPF Inválido", vbInformation
                    TxtCpf.SetFocus
                Else
 '                   TxtPgto.SetFocus
                End If
            ElseIf Len(Trim(LIMPACampo2(TxtCpf))) = 14 Then
                TxtCpf = Format(Trim(LIMPACampo2(TxtCpf)), "@@.@@@.@@@/@@@@-@@")
                If Not VerificaCGC(TxtCpf) Then
                    MsgBox "CNPJ Inválido.", vbInformation
                    TxtCpf.SetFocus
                Else
'                    TxtPgto.SetFocus
                End If
            Else
                MsgBox "Documento invalido", vbInformation
                TxtCpf.SetFocus
            End If
        End If
    Else
'        TxtPgto.SetFocus
    End If
End If

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



Function TrataVidaLink() As String
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



