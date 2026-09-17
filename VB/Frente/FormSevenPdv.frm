VERSION 5.00
Begin VB.Form FormSevenPdv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNroCartao 
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
      Left            =   3870
      TabIndex        =   3
      Tag             =   "N"
      Top             =   3285
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.ComboBox TxtOperadora 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FormSevenPdv.frx":0000
      Left            =   225
      List            =   "FormSevenPdv.frx":000A
      TabIndex        =   4
      Top             =   2475
      Width           =   2805
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
      Left            =   4815
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormSevenPdv.frx":0020
      Top             =   2700
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Pre-Autoização"
      Height          =   510
      Left            =   3735
      TabIndex        =   12
      Top             =   1845
      Width           =   1635
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
      Left            =   225
      TabIndex        =   2
      Tag             =   "N"
      Top             =   1710
      Width           =   2850
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   11
      Top             =   4455
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
      ScaleWidth      =   6855
      TabIndex        =   9
      Top             =   0
      Width           =   6855
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Seven Pdv"
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
         TabIndex        =   10
         Top             =   90
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
      Height          =   645
      Left            =   2340
      TabIndex        =   8
      Text            =   "     Aguarde . . ."
      Top             =   3690
      Visible         =   0   'False
      Width           =   2940
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
      Left            =   225
      TabIndex        =   0
      Tag             =   "N"
      Top             =   1035
      Width           =   2850
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
      Left            =   3690
      TabIndex        =   1
      Tag             =   "N"
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Cartão"
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
      Left            =   3870
      TabIndex        =   14
      Top             =   3015
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
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
      TabIndex        =   13
      Top             =   2205
      Width           =   990
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
      Left            =   225
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
      Left            =   3690
      TabIndex        =   6
      Top             =   810
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Atendimento (NSU)"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "FormSevenPdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Arq As Recordset
Dim VarUsuario As String
Dim VarSenha As String
Public VarEnvio As String
Public VarResposta As String
Public VarCnpjSeven As String
   
   
'          001-000 = 153342           Pedido
Public VarOperadora As String
Public VarNsu As String
Public VarStatus As String
Public VarCodBar As String
Public VarQtd As Currency
Public VarValor As Currency
Public VarLiquido As Currency
Public VarDesconto As Currency
Public VarErro As String


'          040-000 = LOGIXPHARMA      Operadora
'          012-000 = 003734627501     Nsu
'          009-000 = 00              status
'          902-000 = 7896422504843   CodBar
'          905-000 = 001             Qtd
'          910-000 = 0001000         Preço Bruto
'          911-000 = 0000450         Preço Liquido
'          909-000 = 05500           Desconto




'00-000 = 0110
'011-000 = 800500
'          001-000 = 153342           Pedido
'          040-000 = LOGIXPHARMA      Operadora
'          012-000 = 003734627501     Nsu
'940-000 = 003
'942-000 = 99999994002455
'941-000 = 001
'015-000 = 1123153148
'          009-000 = 00              status
'900-000 = 0
'901-000 = 02
'          902-000 = 7896422504843   CodBar
'906-000 = U
'          905-000 = 001             Qtd
'          910-000 = 0001000         Preço Bruto
'          911-000 = 0000450         Preço Liquido
'912-000 = 0000450
'          909-000 = 05500           Desconto
'919-000 = 00
'902-001 = 7896422507134
'906-001 = U
'905-001 = 001
'910-001 = 0000270
'911-001 = 0000121
'912-001 = 0000121
'909-001 = 05518
'919-001 = 00




Private Sub BtOk_Click()
Dim RetSeven As String
TxtAguarde.Visible = True

VarOperadora = ""
VarNsu = ""
VarStatus = ""
VarCodBar = ""
VarQtd = 0
VarValor = 0
VarLiquido = 0
VarDesconto = 0
VarErro = ""

'TxtPedido = psCompString(1, Month(Date), 2) & psCompString(1, Day(Date), 2) & Second(Time)

'TxtPedido = 153342

Open "C:\TRNCENTR\REQ\ENVIO.001" For Output As #1
Print #1, "000-000 = 0100"
Print #1, "011-000 = 800500"
Print #1, "001-000 = " & psCompString(1, FrmFrente.LblPedido, 6)
Print #1, "040-000 = " & TxtOperadora
Print #1, "012-000 = " & TxtCpf   ' NSU
Print #1, "940-000 = 003"
Print #1, "942-000 = " & VarCnpjSeven   ' Cnpj
Print #1, "941-000 = 001"  ' 153342
Print #1, "023-000 = " & Format(Time(), "hhmmss")   ' & psCompString(1, FrmFrente.LblPedido, 6)  ' 153342"   '& TxtPedido ' Numero seguencial
Print #1, "022-000 = " & psCompString(1, Month(Date), 2) & psCompString(1, Day(Date), 2) ' Data  MMDD
Print #1, "004-000 = REA"
Print #1, "900-000 ="

Close #1

DatDados(1).RecordSource = "SELECT SevenPdv.*, SevenPdv.Código FROM SevenPdv WHERE (((SevenPdv.Código)=0))"
DatDados(1).Refresh

If dir("C:\TRNCENTR\RESP\ENVIO.001") <> "" Then
    Kill "C:\TRNCENTR\RESP\ENVIO.001"
End If

DoEvents
x = 0


While dir("C:\TRNCENTR\RESP\ENVIO.001") = ""
 '   DoEvents
Wend

For x = 1 To 700000

Next

x = 0
If dir("C:\TRNCENTR\RESP\ENVIO.001") <> "" Then
'    On Error Resume Next
    Open "C:\TRNCENTR\RESP\ENVIO.001" For Input As #1
        Do
            Line Input #1, RetSeven
            Select Case Left(RetSeven, 7)
                Case "040-000"
                    VarOperadora = Mid(RetSeven, 11, 50)
                Case "012-000"
                    VarNsu = Mid(RetSeven, 11, 50)
                Case "009-000"
                    VarStatus = Mid(RetSeven, 11, 50)
                Case "902-" & psCompString(1, x, 3)
                    VarCodBar = Mid(RetSeven, 11, 50)
                Case "030-000"
                    VarErro = Mid(RetSeven, 11, 50)
                Case "905-" & psCompString(1, x, 3)
                    VarQtd = Mid(RetSeven, 11, 50)
                Case "910-" & psCompString(1, x, 3)
                    VarValor = Mid(RetSeven, 11, 50) / 100
                Case "911-" & psCompString(1, x, 3)
                    VarLiquido = Mid(RetSeven, 11, 50) / 100
                Case "909-" & psCompString(1, x, 3)
                    VarDesconto = Mid(RetSeven, 11, 50) / 100
            End Select
                    
            If Left(RetSeven, 3) = "919" Then
                    DatDados(1).Recordset.AddNew
                    DatDados(1).Recordset.Fields("Nsu") = VarNsu
                    DatDados(1).Recordset.Fields("Pedido") = TxtPedido
                    DatDados(1).Recordset.Fields("Operadora") = TxtOperadora
                    DatDados(1).Recordset.Fields("Data") = Date
                    DatDados(1).Recordset.Fields("CodBar") = VarCodBar
                    DatDados(1).Recordset.Fields("Qtd") = VarQtd
                    DatDados(1).Recordset.Fields("Valor") = VarValor
                    DatDados(1).Recordset.Fields("Liquido") = VarLiquido
                    DatDados(1).Recordset.Fields("Desconto") = VarDesconto
                    DatDados(1).Recordset.Update
                    x = x + 1
            End If
            
        Loop Until EOF(1)
    Close #1
End If

TxtAguarde.Visible = False
If VarStatus = "00" Then
    MsgBox "Produto recuperado com Sucesso !", vbInformation
    
    FrmFrente.LblCartao = "BPMS"
    FrmFrente.LblBPms = TxtOperadora
    FrmFrente.LblNsu = TxtCpf
    FrmFrente.LblNroCartao = TxtNroCartao
    Unload Me
    FrmFrente.TxtDados(0).SetFocus
    FrmFrente.TxtDados(0) = "*"
    SendKeys ("{ENTER}")

Else
    MsgBox "Erro na recuperação da venda !  " & VarErro, vbInformation
    Unload Me
End If

'          040-000 = LOGIXPHARMA      Operadora
'          012-000 = 003734627501     Nsu
'          009-000 = 00              status
'          902-000 = 7896422504843   CodBar
'          905-000 = 001             Qtd
'          910-000 = 0001000         Preço Bruto
'          911-000 = 0000450         Preço Liquido
'          909-000 = 05500           Desconto


'000-000 = 0100
'011-000 = 800500
'001-000 = 153342
'040-000 = LOGIXPHARMA
'012-000 = 3734627501
'940-000 = 003
'942-000 = 99999994002455
'941-000 = 001
'023-000 = 153342
'022-000 = 0817
'004-000 = REA
'900-000 =


End Sub


Private Sub BtSair_Click()
Unload Me
End Sub



Private Sub Form_Activate()
If TxtPedido = "" Then
    MsgBox "Venda não foi aberta !", vbInformation
    Unload Me
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


If LerINI("SevenPdv", "Cnpj", App.Path & "\config.ini") <> "" Then
    VarCnpjSeven = LerINI("SevenPdv", "Cnpj", App.Path & "\config.ini")
    TxtCnpj = VarCnpjSeven
Else
    MsgBox "Cnpj Farmacia não cadadastrado !! ", vbInformation
End If


TxtPedido.Text = FrmFrente.LblPedido

TxtOperadora.ListIndex = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub






Function TrataVidaLink() As String
Dim VarPopular(1 To 6) As String
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



