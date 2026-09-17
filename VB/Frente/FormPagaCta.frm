VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormPagaCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Contas"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox MebData 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1170
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa"
      Height          =   375
      Left            =   7965
      TabIndex        =   25
      Top             =   360
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1545
      Left            =   4680
      ScaleHeight     =   1545
      ScaleWidth      =   4650
      TabIndex        =   18
      Top             =   4815
      Width           =   4650
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   405
         Left            =   2790
         TabIndex        =   22
         Tag             =   "S"
         Top             =   90
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebJuros 
         DataField       =   "Custo"
         Height          =   405
         Left            =   2790
         TabIndex        =   23
         Tag             =   "S"
         Top             =   630
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotal 
         DataField       =   "Custo"
         Height          =   405
         Left            =   2790
         TabIndex        =   24
         Tag             =   "S"
         Top             =   1125
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "SubTotal Saldo Devedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   21
         Top             =   135
         Width           =   2250
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Juros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   20
         Top             =   675
         Width           =   495
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Total Saldo Devedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   765
         TabIndex        =   19
         Top             =   1215
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   270
      TabIndex        =   15
      Top             =   810
      Width           =   3165
      Begin VB.OptionButton OptPgt 
         Caption         =   "&Com Juros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   360
         Width           =   1410
      End
      Begin VB.OptionButton OptPgt 
         Caption         =   "&Sem Juros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdPgto 
      Caption         =   "&Efetua Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7470
      TabIndex        =   5
      Top             =   1170
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   315
      TabIndex        =   8
      Top             =   5040
      Width           =   3390
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dias em Atraso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   945
         Width           =   1425
      End
      Begin VB.Label LblAtraso 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1710
         TabIndex        =   13
         Top             =   945
         Width           =   135
      End
      Begin VB.Label LblCredito 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1305
         TabIndex        =   12
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   630
         Width           =   690
      End
      Begin VB.Label LblPrazo 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1260
         TabIndex        =   10
         Top             =   315
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prazo: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Cta"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   1665
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *  FROM Cta ORDER BY Data"
      Top             =   3375
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormPagaCta.frx":0000
      Height          =   2850
      Left            =   270
      OleObjectBlob   =   "FormPagaCta.frx":001A
      TabIndex        =   6
      Top             =   1935
      Width           =   9105
   End
   Begin VB.Data DatDados 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2475
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT cli.Codcli,cli.nome_cliente , Prazo, Crédito, Correção, Endereço, Nro FROM cli ORDER BY cli.nome_cliente"
      Top             =   360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormPagaCta.frx":13E9
      DataField       =   "Nome_Cliente"
      Height          =   420
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   741
      _Version        =   393216
      ListField       =   "Nome_Cliente"
      BoundColumn     =   "CodCli"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MebPago 
      DataField       =   "Custo"
      Height          =   405
      Left            =   5535
      TabIndex        =   4
      Tag             =   "S"
      Top             =   1170
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   714
      _Version        =   393216
      BackColor       =   16777152
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Data Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   3600
      TabIndex        =   17
      Top             =   900
      Width           =   1530
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5535
      TabIndex        =   16
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "FormPagaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()

Open "LPT1:" For Output As #1
Print #1, Chr(15)
Print #1, "====================================================="
Print #1, Date; "          EXTRATO DE CONTA        "; Time
Print #1, "====================================================="
Print #1, DbcPesquisa.BoundText; " - "; DbcPesquisa.Text
Print #1, Trim(DatDados(1).Recordset.Fields("Endereço")) & ", " & DatDados(1).Recordset.Fields("Nro")
Print #1, "-----------------------------------------------------"
Print #1, "Data"; Tab(13); "Valor"; Tab(22); "Documento"; Tab(32); "Obs"
Print #1, "-----------------------------------------------------"
DatDados(0).Recordset.MoveFirst
While Not DatDados(0).Recordset.EOF
    Print #1, Format(DatDados(0).Recordset.Fields("Data"), "dd/mm/yyyy"); Tab(13); _
              Format(DatDados(0).Recordset.Fields("Valor"), "#,##0.00;-#,##0.00"); Tab(22); _
              DatDados(0).Recordset.Fields("Cupom"); Tab(32); _
              DatDados(0).Recordset.Fields("Obs")
    DatDados(0).Recordset.MoveNext
Wend
Print #1, "-----------------------------------------------------"
    Print #1, "      Conta......:  "; Format(MebSubTot, "#,##0.00;-#,##0.00")
If OptPgt(1).Value = True Then
    Print #1, "      Juros......:  "; Format(MebJuros, "#,##0.00;-#,##0.00")
    Print #1, "      Total .....:  "; Format(MebTotal, "#,##0.00;-#,##0.00")
End If
Print #1, "     "
Print #1, "     "
Print #1, "     "
Print #1, "     "
Print #1, Chr(18);
Close #1


End Sub

Private Sub CmdLimpa_Click()
MebPago = 0

DbcPesquisa = ""
MebSubTot = 0
MebJuros = 0
MebTotal = 0
DatDados(0).RecordSource = "SELECT * FROM Cta  WHERE Cliente = " & Val(DbcPesquisa.BoundText) & " ORDER BY Data"
DatDados(0).Refresh
SomaGrid
DbcPesquisa.SetFocus

End Sub

Private Sub CmdPgto_Click()
Dim ValorPago As Currency
Dim DtSaldo As Date
Dim valor As Currency
Dim Resp
Dim Flag As Boolean
If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma pagamento ?", vbYesNo, App.Title) = vbYes Then
        DatDados(0).Recordset.MoveFirst
        While Not DatDados(0).Recordset.EOF
                 DatDados(0).Recordset.Edit
                 DatDados(0).Recordset.Fields("Cond") = -1
                 DatDados(0).Recordset.Fields("DtPgt") = MebData
                 DatDados(0).Recordset.Update
                 valor = valor - DatDados(0).Recordset("Valor")
                 If valor <= 0 And Flag = True Then
                      DtSaldo = DatDados(0).Recordset.Fields("Data")
                      Flag = False
                 End If
                 DatDados(0).Recordset.MoveNext
        Wend
        If OptPgt(0).Value = True Then
           ValorPago = MebPago - MebSubTot
        Else
           ValorPago = MebPago - MebTotal
        End If
        If ValorPago < 0 Then
            ValorPago = ValorPago - (ValorPago * 2)
            Resp = MsgBox("Ficou " & Format(ValorPago, "#,##0.00;-#,##0.00") & "reais ,clique  sim para gerar comprovante de saldo," & vbCr & " não para assumir como desconto", vbYesNoCancel)
            If Resp = vbYes Then
                DatDados(0).Recordset.AddNew
                DatDados(0).Recordset.Fields("Cliente") = Val(DbcPesquisa.BoundText)
                DatDados(0).Recordset.Fields("Data") = DtSaldo
                DatDados(0).Recordset.Fields("valor") = ValorPago
                DatDados(0).Recordset.Fields("Cupom") = "0"
                DatDados(0).Recordset.Fields("Cond") = 0
                DatDados(0).Recordset.Fields("Obs") = "SALDO"
                DatDados(0).Recordset.Update
                If Impressora = 99 Then
                    Iniciar (False)
                    Print #1, "================================="
                    Print #1, " COMPROVANTE DE SALDO DE CONTA  "
                    Print #1, "================================="
                    Print #1, "Codigo..: "; DbcPesquisa.BoundColumn
                    Print #1, Cond
                    Print #1, DbcPesquisa.Text
                    Print #1, Trim(DatDados(1).Recordset.Fields("Endereço")) & ", " & DatDados(1).Recordset.Fields("Nro")
                    Print #1, Nor
                    Print #1, "Data....: "; Format(MebData, "dd/mm/yyyy")
                    Print #1, "  "
                    Print #1, "Valor...: "; Format(ValorPago, "#,##0.00;-#,##0.00")
                    Print #1, " "
                    Print #1, " "
                    Print #1, " "
                    Print #1, " ---------------------------"
                    Print #1, "        Assinatura             "
                    Print #1, " "
                    Print #1, " "
                    Print #1, " "
                    Print #1, " "
                    Print #1, Nor;
                    Fim (False)
                End If
                Sql = "INSERT INTO REC ( Cliente, Data, Valor, Desconto ) VALUES ( "
                Sql = Sql & Val(DbcPesquisa.BoundText) & " , #"
                Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , "
                Sql = Sql & Numero(MebPago.Text, 2) & " ,"
                Sql = Sql & 0 & ")"
                FrmFrente.Banco.Execute Sql
            ElseIf Resp = vbNo Then
                Sql = "INSERT INTO REC ( Cliente, Data, Valor, Desconto ) VALUES ( "
                Sql = Sql & Val(DbcPesquisa.BoundText) & " , #"
                Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , "
                Sql = Sql & Numero(MebPago.Text, 2) & " ,"
                Sql = Sql & Numero(CStr(ValorPago), 2) & ")"
                FrmFrente.Banco.Execute Sql
            End If
        Else
            Sql = "INSERT INTO REC ( Cliente, Data, Valor, Desconto ) VALUES ( "
            Sql = Sql & Val(DbcPesquisa.BoundText) & " , #"
            Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , "
            Sql = Sql & Numero(MebPago.Text, 2) & " ,"
            Sql = Sql & 0 & ")"
            FrmFrente.Banco.Execute Sql
        End If
        
        CmdLimpa_Click
        DbcPesquisa.SetFocus
    End If
Else
    MebPago = ""
    MebPago.SetFocus
End If

End Sub

Private Sub DbcPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub DbcPesquisa_LostFocus()
If DbcPesquisa.Text <> "" Then
    If Not IsNumeric(DbcPesquisa.Text) Then
        If IsNumeric(DbcPesquisa.BoundText) Then
            Pesquisa DbcPesquisa.BoundText
        End If
    Else
        Pesquisa DbcPesquisa.Text
    End If
End If
End Sub

Private Sub DBGrid1_DblClick()
Dim Sql As String
Dim valor As String
If MsgBox("Confirma o pagamento da Conta --> " & DBGrid1.Columns(0) & " no valor de R$ " & DBGrid1.Columns(3), vbYesNo, App.Title) = vbYes Then
    valor = DBGrid1.Columns(3)
    DatDados(0).Recordset.Edit
    DatDados(0).Recordset.Fields("Cond") = True
    DatDados(0).Recordset.Fields("DtPgt") = Date
    DatDados(0).Recordset.Update
    DatDados(0).Refresh
    Sql = "INSERT INTO REC ( Cliente, Data, Valor, Desconto ) VALUES ( "
    Sql = Sql & Val(DbcPesquisa.BoundText) & " , #"
    Sql = Sql & Format(Date, "mm/dd/yyyy") & "# , "
    Sql = Sql & Numero(valor, 2) & " ,"
    Sql = Sql & 0 & ")"
    Banco.Execute Sql
    SomaGrid
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'Formulário.Left = 700
'Formulário.Top = 200

DatDados(0).RecordSource = "SELECT * FROM Cta  WHERE Cliente = " & Val(DbcPesquisa.BoundText) & " AND Cond = 0 ORDER BY Data, CodCta"
DatDados(0).Refresh
MebData.Text = Date
MebPago = 0
End Sub

Private Sub Form_Load()
Dim Componente As Variant
Set Registros = FrmFrente.Banco.OpenRecordset("SELECT *  FROM Cta")

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Produtos.mdb"
Next
Flag = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Flag = 0
End Sub

Private Sub MebPago_GotFocus()
MebPago = ""
End Sub


Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String

If IsNumeric(DbcPesquisa.BoundText) Then
   Criterio = DatDados(1).Recordset.Fields("Codcli").Name & " = "
   Criterio = Criterio & Codigo
   DatDados(1).Recordset.FindFirst Criterio
   If DatDados(1).Recordset.NoMatch Then
       MsgBox "Cliente não Localizado", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
       DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome_Cliente")
       DatDados(0).RecordSource = "SELECT * FROM Cta  WHERE Cliente = " & Val(Codigo) & "AND Cond = 0 ORDER BY Data, CodCta"
       DatDados(0).Refresh
       If DatDados(0).Recordset.RecordCount <> 0 Then
            DatDados(0).Recordset.MoveLast
       End If
       LblPrazo.Caption = IIf(IsNull(DatDados(1).Recordset.Fields("Prazo")), "", DatDados(1).Recordset.Fields("Prazo")) & " dias"
       LblCredito.Caption = IIf(IsNull(DatDados(1).Recordset.Fields("Crédito")), "", Format(DatDados(1).Recordset.Fields("Crédito"), "##,##0.00"))
       SomaGrid
'       MebPago.SetFocus
   End If
End If


End Sub

Private Sub SomaGrid()
Dim Soma As Recordset
Dim Total As Currency
Dim Flag As Boolean
Dim DtTmp As Date
Dim DtAtraso As Integer
Dim Correcao As Currency
Dim TotJuros As Currency
Dim TotCtaPagas As Currency
Dim TotSaldo As Currency
Flag = True
TotJuros = 0
TotCtaPagas = 0
Set Soma = DatDados(0).Recordset.Clone
If Soma.RecordCount <> 0 Then
    Soma.MoveFirst
    Do Until Soma.EOF
        If Soma.Fields("Cond") = 0 Then
            DtTmp = Soma.Fields("Data")
            If Flag Then
                DtTmp = Soma.Fields("Data")
                If DtTmp < Date - DatDados(1).Recordset.Fields("Prazo") Then
                    LblAtraso.Caption = (Date - DtTmp) - IIf(IsNull(DatDados(1).Recordset.Fields("Prazo")), 0, DatDados(1).Recordset.Fields("Prazo"))
                Else
                    LblAtraso.Caption = "0"
                End If
                Flag = False
            End If
            DtAtraso = (Date - DtTmp) - IIf(IsNull(DatDados(1).Recordset.Fields("Prazo")), 0, DatDados(1).Recordset.Fields("Prazo"))
            Soma.Edit
            Correcao = Soma.Fields("Valor") * (((DatDados(1).Recordset.Fields("Correção") / 100) / 30) * DtAtraso)
            If Soma.Fields("Obs") = "SALDO" Then
                TotSaldo = TotSaldo + Soma.Fields("Valor")
            End If
            If Correcao > 0 Then
                Soma.Fields("Juros") = Correcao
                TotJuros = TotJuros + Correcao
            Else
                Soma.Fields("Juros") = 0
            End If
            Soma.Update
            Total = Total + Soma.Fields("Valor")
        Else
            If Trim(Soma.Fields("Obs")) <> "SALDO" Then
                TotCtaPagas = TotCtaPagas + Soma.Fields("Valor")
            End If
        End If
        Soma.MoveNext
    Loop
End If
MebSubTot = Total
MebJuros = TotJuros
MebTotal = CDec(MebSubTot) + TotJuros
'MebCtaPagas = TotCtaPagas - TotSaldo
End Sub


Private Sub MebPago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    CmdLimpa_Click
End If
End Sub

' Dim ValorPago As Currency
' If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
'    If MsgBox("Confirma pagamento ?", vbYesNo, App.Title) = vbYes Then
'        ValorPago = MebPago
'        DatDados(0).Recordset.MoveFirst
'        While Not DatDados(0).Recordset.EOF
'             If ValorPago > 0 Then
'                 If OptPgt(0).Value = True Then
'                    ValorPago = ValorPago - DatDados(0).Recordset.Fields("Valor")
'                 Else
'                    ValorPago = ValorPago - (DatDados(0).Recordset.Fields("Valor") + DatDados(0).Recordset.Fields("Juros"))
'                 End If
'                 DatDados(0).Recordset.Edit
'                 DatDados(0).Recordset.Fields("Cond") = -1
'                 DatDados(0).Recordset.Fields("DtPgt") = DtData
'                 DatDados(0).Recordset.Update
'             End If
'             DatDados(0).Recordset.MoveNext
'        Wend
'        If ValorPago < 0 Then
'            DatDados(0).Recordset.AddNew
'            DatDados(0).Recordset.Fields("Cliente") = Val(PesqCod)
'            DatDados(0).Recordset.Fields("Data") = DtData
'            DatDados(0).Recordset.Fields("valor") = ValorPago - (ValorPago * 2)
'            DatDados(0).Recordset.Fields("Cupom") = "0"
'            DatDados(0).Recordset.Fields("Cond") = 0
'            DatDados(0).Recordset.Fields("Obs") = "SALDO"
'            DatDados(0).Recordset.Update
'        End If
'        CmdLimpa_Click
'    End If
'End If










Private Sub OptPgt_Click(Index As Integer)

End Sub

Private Sub OptPgt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    CmdLimpa_Click
End If
End Sub
