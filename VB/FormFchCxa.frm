VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormFchCxa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10245
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   21
      Top             =   7560
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
      ScaleWidth      =   10275
      TabIndex        =   17
      Top             =   0
      Width           =   10275
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechamento de Caixa por Período"
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
         TabIndex        =   18
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6360
      Left            =   180
      TabIndex        =   8
      Top             =   675
      Width           =   2340
      Begin VB.Data DatDados 
         Caption         =   "Caixa"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   225
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT MAQUINA  FROM CXA GROUP BY MAQUINA ORDER BY MAQUINA"
         Top             =   4995
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox ChkHist 
         Caption         =   "Agrupar"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Laser/Jato"
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   5940
         Width           =   1470
      End
      Begin VB.CommandButton BtCalcula 
         Caption         =   "&Calcula"
         Height          =   330
         Left            =   180
         TabIndex        =   1
         Top             =   3795
         Width           =   1515
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Fecha Caixa"
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   4290
         Width           =   1515
      End
      Begin VB.CommandButton CmbImprimir 
         Caption         =   "&Matricial"
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   5535
         Width           =   1470
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebValor 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebProximo 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   2685
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDif 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   3300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebData2 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   945
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo DbcCaixa 
         Bindings        =   "FormFchCxa.frx":0000
         Height          =   360
         Left            =   180
         TabIndex        =   19
         Top             =   4950
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "MAQUINA"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   4725
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   2070
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Proximo Caixa"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Retirada"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1815
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Diferença"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   3075
         Width           =   690
      End
   End
   Begin VB.ListBox LstCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6330
      Left            =   2745
      TabIndex        =   7
      Top             =   765
      Width           =   7200
   End
End
Attribute VB_Name = "FormFchCxa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Caixa As Recordset
Public Venda As Recordset
Public Cxa As Recordset
Public Dup As Recordset
Public Rec As Recordset
Public VdaPzo As Recordset
Public VdaVista As Recordset
Public CodVda As Recordset
Public TotRec As Currency
Public VarNroCxa As Integer
Public TotVdaVista As Currency



Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub BtCalcula_Click()
Dim Temp As String
Dim VarTotRec As Recordset
Dim VarVdaVista As Recordset

Dim TotDescVda As Currency
Dim TotVda As Currency
Dim TotDeb As Currency
Dim TotCred As Currency
Dim TotDup As Currency

Dim TotVdaPrazo As Currency
Dim SubRec As Currency
Dim SubDeb As Currency
Dim TotCaixa As Currency
Dim Criterio As String
Dim TotDesc As Currency
Dim TotOut As Currency
Dim TotBco As Currency
Dim totout2 As Currency

TotVda = 0
TotDeb = 0
TotCred = 0
TotRec = 0
TotDup = 0
TotVdaPrazo = 0
TotDesc = 0
TotVdaVista = 0
TotOut = 0
TotBco = 0
totout2 = 0

If MebValor = "" Then MebValor = 0
If MebProximo = "" Then MebProximo = 0
If MebDif = "" Then MebDif = 0

'Set Venda = Banco.OpenRecordset("SELECT First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE PED.DataEmiss = #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.Operacao = 1 Or PED.Operacao=2 Or PED.Operacao =3 GROUP BY ITP.Produto")

If DbcCaixa = "" Then
    Set Venda = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao = 1 or PED.Operacao = 2 or PED.Operacao = 3)  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "#   AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# ")
Else
    Set Venda = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao = 1 or PED.Operacao = 2 or PED.Operacao = 3)  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "#   AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Contato = " & """" & Left(Trim(DbcCaixa), 10) & """")
End If
    

Venda.Requery

If ChkHist.value = 1 Then

    If DbcCaixa = "" Then
        Set Cxa = Banco.OpenRecordset("SELECT cxa.Maquina, Cxa.Historico, Sum(Cxa.Valor) AS SomaDeValor, Cxa.D_C FROM Cxa WHERE Data >= #" & Format(MebData, "mm/dd/yyyy") & "# and Data <= #" & Format(MebData2, "mm/dd/yyyy") & "# GROUP BY cxa.Maquina, Cxa.Historico, Cxa.D_C ORDER BY Cxa.Historico")
    Else
        Set Cxa = Banco.OpenRecordset("SELECT cxa.Maquina, Cxa.Historico, Sum(Cxa.Valor) AS SomaDeValor, Cxa.D_C FROM Cxa WHERE Data >= #" & Format(MebData, "mm/dd/yyyy") & "# and Data <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Maquina = " & """" & Left(Trim(DbcCaixa), 10) & """" & " GROUP BY cxa.Maquina, Cxa.Historico, Cxa.D_C ORDER BY Cxa.Historico")
    End If
Else
    If DbcCaixa = "" Then
        Set Cxa = Banco.OpenRecordset("Select * From Cxa Where Data >= #" & Format(MebData, "mm/dd/yyyy") & "# and Data <= #" & Format(MebData2, "mm/dd/yyyy") & "#  ORDER BY CodCxa")
    Else
        Set Cxa = Banco.OpenRecordset("Select * From Cxa Where Data >= #" & Format(MebData, "mm/dd/yyyy") & "# and Data <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Maquina = " & """" & Left(Trim(DbcCaixa), 10) & """" & " ORDER BY CodCxa")
    End If
End If
SQL = "Select * From Cxa Where Data >= #" & Format(MebData, "mm/dd/yyyy") & "# and Data <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Maquina = " & """" & Left(Trim(DbcCaixa), 10) & """" & " ORDER BY CodCxa"
Cxa.Requery
'Set VdaPzo = Banco.OpenRecordset("SELECT PED.Operacao, First(PED.CodPed) AS PrimeiroDeCodPed, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED GROUP BY PED.Operacao HAVING (((PED.Operacao)=2) AND ((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & "))")

If DbcCaixa = "" Then
    Set VdaPzo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =2 OR PED.Operacao =3)  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "#  ")
Else
    Set VdaPzo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =2 OR PED.Operacao =3)  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Contato = " & """" & Left(Trim(DbcCaixa), 10) & """")
End If
VdaPzo.Requery

If DbcCaixa = "" Then
    Set Rec = Banco.OpenRecordset("SELECT Rec.Maquina, REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec, Rec.Tipo FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE REC.Tipo < 10 AND REC.Data >= #" & Format(MebData, "mm/dd/yyyy") & "# AND REC.Data <= #" & Format(MebData2, "mm/dd/yyyy") & "#  Order By CodRec")
Else
    Set Rec = Banco.OpenRecordset("SELECT Rec.Maquina, REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec, Rec.Tipo FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE REC.Tipo < 10 AND REC.Data >= #" & Format(MebData, "mm/dd/yyyy") & "# AND REC.Data <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Maquina = " & """" & Left(Trim(DbcCaixa), 10) & """" & " Order By CodRec")
End If
Rec.Requery

If DbcCaixa = "" Then
    Set VarVdaVista = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =1 )  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# ")
'    SQL = "SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =1 )  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# "
Else
    Set VarVdaVista = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =1 )  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Contato = " & """" & Left(Trim(DbcCaixa), 10) & """")
'    Sql = " SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (PED.Operacao =1 )  AND PED.DataEmiss >= #" & Format(MebData, "mm/dd/yyyy") & "# AND PED.DataEmiss <= #" & Format(MebData2, "mm/dd/yyyy") & "# AND Contato = " & """" & Trim(DbcCaixa) & """"
End If
VdaPzo.Requery

If Venda.RecordCount > 0 Then
    If Not IsNull(Venda!SomaDetotal) Then
        TotVda = Venda!SomaDetotal '- Venda!SomaDeDesconto
        TotDescVda = Venda!SomaDeDesconto
    End If
End If

If VdaPzo.RecordCount > 0 Then
    If Not IsNull(VdaPzo!SomaDetotal) Then
        TotVdaPrazo = VdaPzo!SomaDetotal ' - VdaPzo!SomaDeDesconto
    End If
End If

If VarVdaVista.RecordCount > 0 Then
    If Not IsNull(VarVdaVista!SomaDetotal) Then
        TotVdaVista = VarVdaVista!SomaDetotal ' - VarVdaVista!SomaDeDesconto
    End If
End If

    Open App.Path & "\Caixa.txt" For Output As #1
    Print #1, ""
    Print #1, Centralizar("FECHAMENTO CAIXA", 42)
    Print #1, "=========================================="
    Print #1, "Data: " & Date & "   Hora: " & Time
    Print #1, ""
    If VarMostraPgtoCaixa Then
        If DbcCaixa = "" Then
            Set VarTotRec = Banco.OpenRecordset("SELECT   REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC  WHERE  REC.Data >=# " & Format(MebData, "mm/dd/yyyy") & "# and REC.Data <=# " & Format(MebData2, "mm/dd/yyyy") & "#  GROUP BY REC.Tipo  ORDER BY REC.Tipo")
        Else
            Set VarTotRec = Banco.OpenRecordset("SELECT Rec.Maquina, REC.Data, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC GROUP BY Rec.Maquina, REC.Data, REC.Tipo HAVING  REC.Data =# " & Format(MebData, "mm/dd/yyyy") & "# AND Maquina = " & """" & Left(Trim(DbcCaixa), 10) & """" & " ORDER BY REC.Tipo")
        End If
        SQL = "SELECT   REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC  WHERE  REC.Data >=# " & Format(MebData, "mm/dd/yyyy") & "# and REC.Data <=# " & Format(MebData2, "mm/dd/yyyy") & "#  GROUP BY REC.Tipo  ORDER BY REC.Tipo"
        
        VarTotRec.Requery
        
        If VarTotRec.RecordCount > 0 Then
            VarTotRec.MoveFirst
            Print #1, "------------------------------------------"
            While Not VarTotRec.Eof
                
                Select Case VarTotRec("Tipo")
                    Case 1
                        Print #1, "Recebimento Dinheiro ....... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                    Case 2
                        Print #1, "Recebimento Cheque ......... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        totout2 = totout2 + VarTotRec("SomaDeValor")
                    Case 3
                        Print #1, "Recebimento Pix/Transf...... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        totout2 = totout2 + VarTotRec("SomaDeValor")
                    Case 4
                        Print #1, "Recebimento Cartão Débito... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        totout2 = totout2 + VarTotRec("SomaDeValor")
                    Case 5
                        Print #1, "Recebimento Cartão Crédito.. "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        totout2 = totout2 + VarTotRec("SomaDeValor")
                    Case 6
                        Print #1, "Recebimento Via Banco ...... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        totout2 = totout2 + VarTotRec("SomaDeValor")
                End Select
                
                
                Select Case VarTotRec("Tipo")
                    Case 11
                        Print #1, "Venda Dinheiro ............. "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
'                        If DbcCaixa = "" Then
                            TotVdaVista = VarTotRec("SomaDeValor")
'                        End If

                    Case 12
                        Print #1, "Venda Cheque ............... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 13
                        Print #1, "Venda Pix/Transf. .......... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 14
                        Print #1, "Venda Cartão Débito......... "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 15
                        Print #1, "Venda Cartão Crédito........ "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 16
                        Print #1, "Venda a Prazo .............. "; Alinhar(VarTotRec("SomaDeValor"), 10, 2)
                    Case Else
                        If DbcCaixa = "" Then
                            TotVdaVista = 0
                        End If
                        
                End Select

        
        
        '        If TotVda("Pgto") < 6 Then
        '            Entradas = Entradas + TotVda("SomaDeTotal")
        '        End If
                VarTotRec.MoveNext
            Wend
            Print #1, "------------------------------------------"
            
        End If
    End If

    If Cxa.RecordCount > 0 Then
        Print #1, "Data"; Tab(12); "Hora"; Tab(18); "Historico"; Tab(34); "Valor"; Tab(40); "D/C"
        
        With Cxa
            .MoveFirst
            While Not .Eof
                If ChkHist.value = 1 Then
                    Print #1, Tab(0); Left(.Fields("Historico"), 24); Tab(29); Alinhar(.Fields("SomaDeValor"), 10, 2); Tab(41); Left(.Fields("D_C"), 2)
                    If Left(.Fields("D_C"), 2) = "C" Then
                        TotCred = CCur(TotCred) + .Fields("SomaDeValor")
                    ElseIf Left(.Fields("D_C"), 2) = "D" Then
                        TotDeb = CCur(TotDeb) + .Fields("SomaDeValor")
                    End If
                
                Else
                    Print #1, Tab(0); .Fields("Data"); Tab(12); Left(.Fields("Hora"), 5); Tab(18); Left(.Fields("Historico"), 15); Tab(29); Alinhar(.Fields("Valor"), 10, 2); Tab(41); Left(.Fields("D_C"), 2)
                    If Left(.Fields("D_C"), 2) = "C" Then
                        TotCred = CCur(TotCred) + .Fields("Valor")
                    ElseIf Left(.Fields("D_C"), 2) = "D" Then
                        TotDeb = CCur(TotDeb) + .Fields("Valor")
                    End If
                End If
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Não tem movimentação manual de caixa"
    End If

    ' Recebimento
    Print #1, ""
    Print #1, "------------------------------------------"
    If Rec.RecordCount > 0 Then
 '       Print #1, "Hora"; Tab(8); "Nome"; Tab(33); "Valor"; Tab(41); "Desconto"
        
        With Rec
            .MoveFirst
            While Not .Eof
'                Print #1, Tab(0); Left(.Fields("Hora"), 5); Tab(8); Left(.Fields("Nome"), 20); Tab(28); Alinhar(.Fields("Valor"), 10, 2); Tab(38); Alinhar(.Fields("Desconto"), 10, 2)
                TotRec = CCur(TotRec) + (.Fields("Valor"))
                TotDesc = CCur(TotDesc) + (.Fields("Desconto"))
                .MoveNext
            Wend
        End With
    Else
 '       Print #1, "Nenhum recebimento de parcelas"
    End If


    Print #1, "Venda ...................(+)  "; Alinhar(CStr(TotVdaVista + TotOut), 11, 2)
    Print #1, "Recebimento Conta........(+)  "; Alinhar(CStr(TotRec), 11, 2)
    Print #1, "Entradas Diversas........(+)  "; Alinhar(CStr(TotCred), 11, 2)
    
    Print #1, "                             ------------- "
    SubRec = TotVdaVista + TotCred + TotRec + TotOut
    Print #1, "Sub Total ...............(=)  "; Alinhar(CStr(SubRec), 11, 2)
    
    Print #1, "Saidas Diversas..........(-)  "; Alinhar(CStr(TotDeb), 11, 2)
    Print #1, "Outras Formas Recebimento(-)  "; Alinhar(CStr(TotOut + totout2), 11, 2)
    
    Print #1, "                             ------------- "
    
    SubDeb = TotDeb ' + TotVdaPrazo
    TotCaixa = SubRec - (SubDeb + TotOut + totout2)
    Print #1, "Total do Caixa...........(=)  "; Alinhar(CStr(TotCaixa), 11, 2)
    Print #1, "                             ------------- "
    Print #1, ""
'    If CmdOk.Enabled = True Then
        MebDif = (TotCaixa - MebValor) - MebProximo
        Print #1, "Retirada do Fechamento...(-)  "; Alinhar(MebValor, 11, 2)
        Print #1, "Diferença..............(+/-)  "; Alinhar(MebDif, 11, 2)
        Print #1, "Proximo Caixa...............  "; Alinhar(MebProximo, 11, 2)
'    Else
'        MebDif = 0
'    End If
'
    Print #1, ""
    Print #1, "     INFORMACOES ADICIONAIS "
    Print #1, ""
    
    Print #1, "Venda Total..............( )   "; Alinhar(CStr(TotVda), 10, 2)
    Print #1, "Venda a Prazo............( )   "; Alinhar(CStr(TotVdaPrazo), 10, 2)

    Print #1, "Desconto de Venda vista..( )   "; Alinhar(CStr(TotDescVda), 10, 2)
    Print #1, "Desconto de Recebimento..( )   "; Alinhar(CStr(TotDesc), 10, 2)
    
    
    
    Print #1, "" & Nor
    Print #1, "<>"

    Close #1
    
    On Error GoTo Erro
    LstCaixa.Clear
    Open App.Path & "\Caixa.txt" For Input As #1
    Do
        Line Input #1, Temp
        LstCaixa.AddItem Temp
'        Lista.Selected(Lista.ListCount - 1) = True
    Loop Until Temp = "<>"
    Close #1

Erro:
If Err.Number = 62 Then
    Close #1
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbImprimir_Click()
    
    If MsgBox("Confirma impressão?", 36) = vbYes Then
    
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
        
        Fs2.CopyFile App.Path & "\caixa.txt", Porta
        
'        Fs2.Close
        
    End If

End Sub

Private Sub CmdOk_Click()
Dim SQL As String
Dim PgtAnt As Byte
Dim Total As Currency
Dim Texto As String
Dim TotVdaVista As Currency

' V= VISTA   R=RECEBIMENTO   S=DIFERENCIA/SALDO  F=FECHAMENTO

BtCalcula_Click

If MebValor = "" Then
    MebValor.SetFocus
    Exit Sub
End If

If MebData <> MebData2 Then

    MsgBox "O Caixa só pode ser fechado com as datas iguais !!!", vbInformation, App.Title
    Exit Sub
    
End If

If MsgBox("Confirma Fechamento do Caixa ?", vbYesNo, App.Title) = vbYes Then
        If DbcCaixa = "" Then
            Set VdaVista = Banco.OpenRecordset("SELECT PED.Operacao, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE PED.DataEmiss = #" & Format(MebData, "mm/dd/yyyy") & "#  GROUP BY PED.Operacao HAVING (((PED.Operacao)=1))")
        Else
            Set VdaVista = Banco.OpenRecordset("SELECT PED.Operacao, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE PED.DataEmiss = #" & Format(MebData, "mm/dd/yyyy") & "#  GROUP BY PED.Operacao, PED.Contato HAVING (((PED.Operacao)=1)) AND Contato = " & """" & Left(Trim(DbcCaixa), 10) & """")
        End If
        VdaVista.Requery
        
        If VdaVista.RecordCount <> 0 Then
            
            If DbcCaixa <> "" Then

                TotVdaVista = VdaVista!SomaDetotal '- VdaVista!SomaDeDesconto
            End If
            Texto = "VENDA VISTA"
            
            SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
            SQL = SQL + """" & DbcCaixa & """" & ", "
            SQL = SQL & """" & Texto & """" & " , #"
            SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
            SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
            SQL = SQL & Numero2(CStr(TotVdaVista)) & " ,"
            SQL = SQL & VarNroCxa & " ,"
            SQL = SQL & """" & "V" & """" & ")"
            Banco.Execute SQL
            Total = 0
       End If
'    Next
    Texto = "RECEBIMENTO"
    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & DbcCaixa & """" & ", "
    SQL = SQL & """" & Texto & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(CStr(TotRec)) & " ,"
    SQL = SQL & VarNroCxa & " ,"
    SQL = SQL & """" & "R" & """" & ")"
    Banco.Execute SQL
    Total = 0
    
    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & DbcCaixa & """" & ", "
    SQL = SQL & """" & "FECHAMENTO" & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebValor) & " ,"
    SQL = SQL & VarNroCxa & " ,"
    SQL = SQL & """" & "F" & """" & ")"
    Banco.Execute SQL

    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & DbcCaixa & """" & ", "
    SQL = SQL & """" & "DIFERENCA" & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebDif) & " ,"
    SQL = SQL & VarNroCxa & " ,"
    SQL = SQL & """" & "S" & """" & ")"
    Banco.Execute SQL

    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & DbcCaixa & """" & ", "
    SQL = SQL & """" & "ABRE" & """" & " , #"
    SQL = SQL & Format(CDate(MebData) + 1, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebProximo) & " ,"
    SQL = SQL & VarNroCxa + 1 & " ,"
    SQL = SQL & """" & "C" & """" & ")"
    Banco.Execute SQL


    MebData.Text = Date
    MebValor = 0
    MebDif = 0
    MebProximo = 0
    MebData = CDate(MebData) + 1
    
    BtCalcula_Click
        
    MebValor.Text = 0
    MebData.SetFocus
End If

End Sub

Private Sub Command1_Click()

    FrmRelTXT.Config App.Path & "\Caixa.Txt"

End Sub

Private Sub Form_Activate()
'If DatDados(1).Recordset.RecordCount > 0 Then
'    DbcCaixa.Text = DatDados(1).Recordset.Fields("Maquina")
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

MebData.Text = Date
MebData2.Text = Date

MebValor = 0
MebDif = 0
MebProximo = 0

Set Caixa = Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
Caixa.Requery
If Caixa.RecordCount > 0 Then
    Caixa.MoveLast
    VarNroCxa = Caixa!CodControle
End If

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

DatDados(1).DatabaseName = App.Path & "\dados.mdb"
 

End Sub

Private Sub MebDif_LostFocus()
BtCalcula_Click
End Sub


Private Sub MebProximo_LostFocus()
BtCalcula_Click
End Sub


Private Sub MebValor_LostFocus()
BtCalcula_Click
End Sub
