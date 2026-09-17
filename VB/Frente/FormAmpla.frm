VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormAmpla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoriza Amplacard"
   ClientHeight    =   6600
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Autorização"
      Height          =   645
      Left            =   2835
      TabIndex        =   21
      Top             =   1575
      Width           =   3210
      Begin VB.Label LblAutoriza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   22
         Top             =   225
         Width           =   2805
      End
   End
   Begin VB.TextBox TxtAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4410
      TabIndex        =   20
      Text            =   "     Aguarde . . ."
      Top             =   3060
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   6120
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   5865
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   5655
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   5430
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   5190
      Width           =   10740
   End
   Begin VB.TextBox VarPopular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   4950
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
      Left            =   6435
      TabIndex        =   3
      Top             =   1080
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtUf 
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
      Left            =   4995
      TabIndex        =   2
      Tag             =   "N"
      Text            =   "SP"
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox TxtCrm 
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
      Left            =   2835
      TabIndex        =   1
      Tag             =   "N"
      Top             =   1080
      Width           =   2040
   End
   Begin VB.TextBox TxtCpf 
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
      Left            =   180
      TabIndex        =   0
      Tag             =   "N"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox TxtCnpj 
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
      Left            =   180
      TabIndex        =   10
      Tag             =   "N"
      Top             =   405
      Width           =   2535
   End
   Begin VB.TextBox TxtPedido 
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
      Left            =   9090
      TabIndex        =   9
      Tag             =   "N"
      Top             =   225
      Width           =   1680
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Pre-Autoização"
      Height          =   510
      Left            =   585
      TabIndex        =   4
      Top             =   1665
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
      Left            =   4275
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormAmpla.frx":0000
      Top             =   4140
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormAmpla.frx":00F6
      Height          =   2475
      Left            =   90
      OleObjectBlob   =   "FormAmpla.frx":0110
      TabIndex        =   5
      Top             =   2340
      Width           =   10770
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   8190
      Picture         =   "FormAmpla.frx":11C8
      Stretch         =   -1  'True
      Top             =   765
      Width           =   2640
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
      Left            =   6435
      TabIndex        =   13
      Top             =   810
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
      Left            =   4995
      TabIndex        =   12
      Top             =   810
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
      Left            =   2835
      TabIndex        =   11
      Top             =   810
      Width           =   1635
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
      Left            =   180
      TabIndex        =   8
      Top             =   180
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
      Left            =   9090
      TabIndex        =   7
      Top             =   0
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
      Left            =   180
      TabIndex        =   6
      Top             =   810
      Width           =   1920
   End
End
Attribute VB_Name = "FormAmpla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arq As Recordset
Dim VarUsuario As String
Dim VarSenha As String



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
        DatDados(1).Recordset("Crm") = TxtCrm
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
                    
    Open "c:\FPopular\FP_ENV.TXT" For Output As #1
    Print #1, "01" & Trim(VarUsuario)
    Print #1, "02" & Trim(VarSenha)
    Print #1, "03" & psCompString(1, TxtPedido, 6) & psCompString(1, LIMPACampo2(TxtCnpj), 14) & Trim(psCompString(1, LIMPACampo2(TxtCpf), 11)) & psCompString(1, TxtCrm, 6) & Left(TxtUf, 2) & Format(MebData, "dd/mm/yyyy")
    DatDados(1).Recordset.MoveFirst
    While Not DatDados(1).Recordset.EOF
        VarQtd = psCompString(1, LimpaNumero(DatDados(1).Recordset("QtdSolicitada") * 100), 5)
        VarPreco = psCompString(1, LimpaNumero(DatDados(1).Recordset("PcoVenda") * 100), 15)
        VarPrescrita = psCompString(1, LimpaNumero(DatDados(1).Recordset("QtdPrescrita") * 100), 5)
        
        Print #1, "04" & psCompString(1, DatDados(1).Recordset("CodBarra"), 13) & VarQtd & VarPreco & VarPrescrita
        DatDados(1).Recordset.MoveNext
    Wend
    Close #1
        
    TrataPopular
                
End If
TxtAguarde.Visible = False

'Unload Me
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim sql As String
Dim VarItem As Recordset

MebData = Date
TxtUf = "SP"
'TxtCpf = "00106424122"
'TxtCrm = "25142"

Dim VarLoja  As String
Dim VarCxa As String

If LerINI("Sistema", "loja", App.Path & "\Amplacard.ini") <> "" Then
    VarLoja = LerINI("sistema", "Loja", App.Path & "\Amplacard.ini")
End If

If LerINI("Sistema", "Caixa", App.Path & "\Amplacard.ini") <> "" Then
    VarCxa = LerINI("sistema", "Caixa", App.Path & "\Amplacard.ini")
End If

Open "C:\Winguest\Ped\" & VarLoja & VarCxa & ".STS" For Output As #1
Print #1, "9002"
Close #1

If dir("C:\Winguest\Ret\" & VarLoja & VarCxa & ".STS") <> "" Then
    Kill "C:\FPOPULAR\FP_REQ.txt"
End If

If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
    Kill "C:\FPOPULAR\FP_ENV.txt"
End If

TxtPedido.Text = FrmFrente.LblPedido

DatDados(1).RecordSource = "SELECT Pedido, PRD.Descrição_Produto,POPULAR.Cpf, POPULAR.Crm, POPULAR.UfPrescritor,POPULAR.DtPrescricao,  POPULAR.CodBarra, POPULAR.QtdSolicitada, POPULAR.PcoVenda, POPULAR.QtdPrescrita, POPULAR.Codigo, Popular.PreAutoriza FROM POPULAR INNER JOIN PRD ON POPULAR.CodPrd = PRD.CodPrd WHERE (((POPULAR.Pedido)= " & TxtPedido & ")) ORDER BY POPULAR.Codigo "
DatDados(1).Refresh

If DatDados(1).Recordset.RecordCount <= 0 Then
    Set VarItem = FrmFrente.Banco.OpenRecordset("Select * From Itp INNER JOIN BAR ON ITP.Produto = BAR.Produto Where Pedido = " & TxtPedido & " AND BAR.BarBarra<>CStr([itp].[produto]) Order By CodItp")
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
End If

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

x = 0
LblAutoriza = ""

'For X = 1 To 700000
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

