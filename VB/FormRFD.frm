VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormRFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação RFD"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCopia 
      Caption         =   "&Importar"
      Height          =   420
      Left            =   5535
      TabIndex        =   3
      Top             =   765
      Width           =   1500
   End
   Begin VB.CommandButton Sair 
      Caption         =   "&Sair"
      Height          =   420
      Left            =   5535
      TabIndex        =   2
      Top             =   1425
      Width           =   1500
   End
   Begin VB.TextBox TxtCaminho 
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Top             =   795
      Width           =   4875
   End
   Begin VB.CommandButton CmdLoca 
      Caption         =   "&Localizar"
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Top             =   1440
      Width           =   1680
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2835
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe o drive destino"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   585
      Width           =   1605
   End
End
Attribute VB_Name = "FormRFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VarPed As Recordset
Public VarItp As Recordset

Private Sub CmdCopia_Click()
Dim VarTemp As String
Dim VarData As Date
Dim VarDtTemp As String
Open TxtCaminho For Input As #6
Do
    Line Input #6, VarTemp
    If Left(VarTemp, 3) = "E14" And Mid(VarTemp, 123, 1) = "N" Then
        VarPed.AddNew
        VarPed!CodPed = Mid(VarTemp, 47, 6)
        VarPed!Operacao = 1
        VarPed!Entidade = 1
        VarDtTemp = Mid(VarTemp, 65, 2) & "/" & Mid(VarTemp, 63, 2) & "/" & Mid(VarTemp, 59, 4)
        VarPed!DataEmiss = CDate(VarDtTemp)
        VarPed!Total = Mid(VarTemp, 109, 14) / 100
        VarPed!NroNt = Mid(VarTemp, 53, 6)
        VarPed!DtFat = CDate(VarDtTemp)
        VarPed!DtSaida = CDate(VarDtTemp)
        VarPed!Desconto = Mid(VarTemp, 81, 13) / 100
        VarPed.Update
    End If
Loop Until Eof(6)
Close #6

Open TxtCaminho For Input As #7

Do
    Line Input #7, VarTemp
    If Left(VarTemp, 3) = "E15" And Mid(VarTemp, 231, 1) = "N" Then
        VarItp.AddNew
        VarItp!Pedido = Mid(VarTemp, 53, 6)
        VarItp!Produto = 0
        VarItp!Complemento = Mid(VarTemp, 76, 100)
'        VarDtTemp = Mid(Vartemp, 65, 2) & "/" & Mid(Vartemp, 63, 2) & "/" & Mid(Vartemp, 59, 4)
'        VarItp!DataEmiss = CDate(VarDtTemp)
        VarItp!Qtd = Mid(VarTemp, 176, 7) / 1000
        VarItp!ValorUnit = Mid(VarTemp, 186, 8) / 100
        VarItp!Valor = Mid(VarTemp, 210, 14) / 100
        VarItp!VlrVdaDia = Mid(VarTemp, 210, 14) / 100
'        VarItp!NroNt = Mid(Vartemp, 47, 6)
'        VarItp!DtFat = CDate(VarDtTemp)
'        VarItp!DtSaida = CDate(VarDtTemp)
'        VarItp!Desconto = Mid(Vartemp, 81, 13) / 100
        VarItp.Update
    End If
Loop Until Eof(7)
Close #7

MsgBox "Arquivo importado com sucesso !!!"


Kill TxtCaminho

'CmdLoca_Click

End Sub

Private Sub CmdLoca_Click()
CommonDialog.filename = ""
CommonDialog.ShowSave
TxtCaminho.Text = CommonDialog.filename
'CmdCopia_Click
End Sub

Private Sub Form_Load()

Set VarPed = Banco.OpenRecordset("Select * From Ped")
VarPed.Requery

Set VarItp = Banco.OpenRecordset("Select * From Itp")
VarItp.Requery


End Sub

Private Sub Sair_Click()
Unload Me
End Sub
