VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormFchCxanovo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -900
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   29
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
      ScaleWidth      =   9915
      TabIndex        =   24
      Top             =   0
      Width           =   9915
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechamento de Caixa"
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
         Top             =   45
         Width           =   5325
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   6660
      TabIndex        =   7
      Top             =   900
      Width           =   3030
      Begin VB.CommandButton Command1 
         Caption         =   "&Laser/Jato"
         Height          =   375
         Left            =   1530
         TabIndex        =   30
         Top             =   5355
         Width           =   1335
      End
      Begin VB.CheckBox ChkDiscr 
         Caption         =   "Discriminar Pgto"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox CHKDET 
         Caption         =   "Completo"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton CmbImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   5355
         Width           =   1335
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Fecha Caixa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   810
         TabIndex        =   5
         Top             =   4860
         Width           =   1335
      End
      Begin VB.CommandButton BtCalcula 
         Caption         =   "&Calcula"
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Top             =   1080
         Width           =   915
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   315
         Left            =   315
         TabIndex        =   0
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
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
         Left            =   315
         TabIndex        =   3
         Top             =   3195
         Width           =   1635
         _ExtentX        =   2884
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
      Begin MSMask.MaskEdBox MebNroCxa 
         Height          =   315
         Left            =   315
         TabIndex        =   1
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebProximo 
         Height          =   315
         Left            =   315
         TabIndex        =   4
         Top             =   3780
         Width           =   1635
         _ExtentX        =   2884
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
      Begin MSMask.MaskEdBox MebPed 
         Height          =   195
         Left            =   870
         TabIndex        =   8
         Top             =   2160
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPcl 
         Height          =   195
         Left            =   870
         TabIndex        =   9
         Top             =   2400
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebRec 
         Height          =   195
         Left            =   870
         TabIndex        =   10
         Top             =   2655
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPed2 
         Height          =   195
         Left            =   1725
         TabIndex        =   11
         Top             =   2160
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPcl2 
         Height          =   195
         Left            =   1725
         TabIndex        =   12
         Top             =   2400
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebRec2 
         Height          =   195
         Left            =   1725
         TabIndex        =   13
         Top             =   2655
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDif 
         Height          =   315
         Left            =   315
         TabIndex        =   21
         Top             =   4365
         Width           =   1635
         _ExtentX        =   2884
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
      Begin VB.Label LblMaquina 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   5940
         Width           =   2535
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Diferença"
         Height          =   195
         Left            =   315
         TabIndex        =   22
         Top             =   4140
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         Height          =   195
         Left            =   315
         TabIndex        =   20
         Top             =   225
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Retirada"
         Height          =   195
         Left            =   315
         TabIndex        =   19
         Top             =   2970
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Nro Caixa"
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   855
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Proximo Caixa"
         Height          =   195
         Left            =   315
         TabIndex        =   17
         Top             =   3555
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   285
         TabIndex        =   15
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Receb."
         Height          =   195
         Left            =   285
         TabIndex        =   14
         Top             =   2655
         Width           =   525
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
      Left            =   270
      TabIndex        =   6
      Top             =   945
      Width           =   6225
   End
End
Attribute VB_Name = "FormFchCxanovo"
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
Public ArqDesc As Recordset

Public TotRec As Currency


Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub BtCalcula_Click()
Dim Temp As String
Dim TotVda As Currency
Dim TotDeb As Currency
Dim TotCred As Currency
Dim TotDup As Currency
Dim TotVdaPrazo As Currency
Dim SubRec As Currency
Dim SubDeb As Currency
Dim TotCaixa As Currency
Dim sql As String
Dim Criterio As String
Dim VarTotRec As Recordset
Dim TotOut As Currency
Dim TotDescVda As Currency
Dim MovManu  As Recordset


TotVda = 0
TotDeb = 0
TotCred = 0
TotRec = 0
TotDup = 0
TotVdaPrazo = 0
TotOut = 0

Set CodVda = FrmFrente.Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
CodVda.Requery

Set Dup = FrmFrente.Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
Dup.Requery

Set Rec = FrmFrente.Banco.OpenRecordset("Select CodRec From Rec  Order by CodRec")
Rec.Requery

Set MovManu = FrmFrente.Banco.OpenRecordset("SELECT EST.Data, EST.Codigo, PRD.Descrição_Produto, EST.Tipo, EST.Data, EST.Hora, EST.Qtd FROM EST INNER JOIN PRD ON EST.Produto = PRD.CodPrd WHERE EST.Data = #" & Format(MebData, "mm/dd/yyyy") & "# ORDER BY EST.Codigo")
Rec.Requery


Criterio = "CodControle = " & MebNroCxa + 1
Caixa.FindFirst Criterio

If Not Caixa.NoMatch Then
    CmdOk.Enabled = False
'    MebNroCxa = Caixa!CodControle
'    Caixa.MoveLast
    MebPed2 = Caixa!Pedido
    MebPcl2 = Caixa!Dup
    MebRec2 = Caixa!Rec

Else
    CmdOk.Enabled = True
    If CodVda.RecordCount > 0 Then
       CodVda.MoveLast
       MebPed2 = CodVda!CodPed
    Else
        MebPed2 = MebPed
    End If
    
    If Dup.RecordCount > 0 Then
       Dup.MoveLast
       MebPcl2 = Dup!CodDup
    Else
       MebPcl2 = MebPcl
    End If
    
    If Rec.RecordCount > 0 Then
       Rec.MoveLast
       MebRec2 = Rec!CodRec
    Else
       MebRec2 = MebRec
    End If

End If



Criterio = "CodControle = " & MebNroCxa
Caixa.FindFirst Criterio

MebNroCxa = Caixa!CodControle
MebPed = Caixa!Pedido
MebPcl = Caixa!Dup
MebRec = Caixa!Rec

'LblMaquina = "DESKTOP-VN"

sql = "SELECT  First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto  FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3))AND Ped.Contato = " & """" & Trim(LblMaquina) & """" & " GROUP BY ITP.Produto"

Set Venda = FrmFrente.Banco.OpenRecordset(sql)

'sql = "SELECT Ped.Contato, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) AND Ped.contato = " & """" & LblMaquina & """"
Venda.Requery
Set ArqDesc = FrmFrente.Banco.OpenRecordset("SELECT  Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) AND Ped.contato = " & """" & LblMaquina & """")
ArqDesc.Requery

sql = "SELECT  Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) AND Ped.contato = " & """" & LblMaquina & """"

'Set Venda = Banco.OpenRecordset("SELECT  First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor FROM (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd GROUP BY ITP.Produto HAVING (((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & " ) AND ((First(PED.Operacao))=1 Or (First(PED.Operacao))=2 Or (First(PED.Operacao))=3))")

Set Cxa = FrmFrente.Banco.OpenRecordset("Select * From Cxa Where NroCaixa = " & MebNroCxa & " AND Maquina = " & """" & LblMaquina & """")
Cxa.Requery
'Set VdaPzo = Banco.OpenRecordset("SELECT PED.Operacao, First(PED.CodPed) AS PrimeiroDeCodPed, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED GROUP BY PED.Operacao HAVING (((PED.Operacao)=2) AND ((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & "))")

Set VdaPzo = FrmFrente.Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.Operacao)=2) AND ((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ")) AND Contato = " & """" & LblMaquina & """")
VdaPzo.Requery

Set Rec = FrmFrente.Banco.OpenRecordset("SELECT REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE REC.Tipo < 10 AND (((REC.CodRec)> " & MebRec & " And (REC.CodRec) <= " & MebRec2 & ")) AND Maquina = " & """" & LblMaquina & """" & " Order By CodRec")
Rec.Requery

If Venda.RecordCount > 0 Then
    If Not IsNull(Venda!SomaDetotal) Then
'        TotVda = Venda!SomaDetotal '- Venda!SomaDeDesconto
        TotDescVda = Venda!SomaDeDesconto
    End If
End If


If VdaPzo.RecordCount > 0 Then
    If Not IsNull(VdaPzo!SomaDetotal) Then
        TotVdaPrazo = VdaPzo!SomaDetotal '- VdaPzo!SomaDeDesconto
    End If
End If

    Open App.Path & "\Caixa.txt" For Output As #1
    Print #1, "" + Cond
    Print #1, Centralizar("FECHAMENTO CAIXA", 42)
    Print #1, "=========================================="
    If CmdOk.Enabled = False Then
        Print #1, "Data: " & Cxa!Data & "   Hora: " & Cxa!Hora
        Print #1, ""
        Print #1, Tab(1); "Ref.Cxa: "; psCompString(1, MebNroCxa, 6) & " - " & Cxa!Data
    Else
        Print #1, "Data: " & Date & "   Hora: " & Time
        Print #1, ""
        Print #1, Tab(1); "Ref.Cxa: "; psCompString(1, MebNroCxa, 6) & " - " & MebData
    End If
    ' Produtos
    If CHKDET.Value = 1 Then
        Print #1, ""
        Print #1, "------------------------------------------"
    End If
    If Venda.RecordCount > 0 Then
        If CHKDET.Value = 1 Then
            Print #1, "Produto"; Tab(19); "Quantidade"; Tab(37); "Valor"
        End If
    
        With Venda
            .MoveFirst
            While Not .EOF
                If CHKDET.Value = 1 Then
                    Print #1, Tab(0); Left(.Fields("PrimeiroDeDescrição_Produto"), 17); Tab(21); Alinhar(.Fields("SomaDeQtd"), 9, 2); Tab(34); Alinhar(.Fields("SomaDeValor"), 9, 2)
                End If
                TotVda = CCur(TotVda) + (.Fields("SomaDeTotal")) ' - .Fields("SomaDeDesconto"))
                .MoveNext
            Wend
        End With
        TotVda = TotVda - ArqDesc.Fields("SomaDeDesconto")
        
    Else
        Print #1, "Não tem produto vendido"
    End If
    
    ' Caixa
    Print #1, ""
    Print #1, "------------------------------------------"
    If Cxa.RecordCount > 0 Then
        Print #1, "Data"; Tab(12); "Hora"; Tab(18); "Historico"; Tab(31); "Valor"; Tab(40); "D/C"
        
        With Cxa
            .MoveFirst
            While Not .EOF
                Print #1, Tab(0); .Fields("Data"); Tab(12); Left(.Fields("Hora"), 5); Tab(18); Left(.Fields("Historico"), 12); Tab(31); Alinhar(.Fields("Valor"), 9, 2); Tab(41); Left(.Fields("D_C"), 2)
                If Left(.Fields("D_C"), 2) = "C" Then
                    TotCred = CCur(TotCred) + .Fields("Valor")
                ElseIf Left(.Fields("D_C"), 2) = "D" Then
                    TotDeb = CCur(TotDeb) + .Fields("Valor")
                End If
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Não tem movimentação de caixa"
    End If

    ' Recebimento
    If CHKDET.Value = 1 Then
        Print #1, ""
        Print #1, "------------------------------------------"
    End If
    If Rec.RecordCount > 0 Then
        If CHKDET.Value = 1 Then
            
            Print #1, "Hora"; Tab(8); "Nome"; Tab(24); "Valor"; Tab(35); "Desconto"
        End If
        
        With Rec
            .MoveFirst
            While Not .EOF
                If CHKDET.Value = 1 Then
                    Print #1, Tab(0); Left(.Fields("Hora"), 5); Tab(8); Left(.Fields("Nome"), 15); Tab(21); Alinhar(.Fields("Valor"), 9, 2); Tab(32); Alinhar(.Fields("Desconto"), 9, 2)
                End If
                TotRec = CCur(TotRec) + (.Fields("Valor"))
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Nenhum recebimento de parcelas"
    End If
    
    If MovManu.RecordCount > 0 Then
        Print #1, "------------------------------------------"
        Print #1, "Data"; Tab(12); "Hora"; Tab(18); "Produto"; Tab(31); "Qtd"; Tab(40); "E/S"

        With MovManu
            .MoveFirst
            While Not .EOF
                Print #1, Tab(0); .Fields("Data"); Tab(12); Left(.Fields("Hora"), 5); Tab(18); Left(.Fields("Descrição_Produto"), 12); Tab(31); Alinhar(.Fields("Qtd"), 8, 2); Tab(41); IIf(.Fields("Tipo") = 1, "E", "S")
                .MoveNext
            Wend
        End With
    Else
'        Print #1, "Não tem movimentação de caixa"
    End If
    
    
'        sql = "SELECT Rec.Maquina, REC.Data, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC GROUP BY Rec.Maquina, REC.Data, REC.Tipo HAVING  Maquina = " & """" & LblMaquina & """" & " ORDER BY REC.Tipo"
    If ChkDiscr.Value Then
'        Set VarTotRec = FrmFrente.Banco.OpenRecordset("SELECT  Rec.Maquina,  First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE  REC.Tipo > 10 AND (((REC.CodRec)> " & MebRec & " And (REC.CodRec) <= " & MebRec2 & ")) GROUP BY  Rec.Maquina,  REC.Tipo HAVING  Maquina = " & """" & LblMaquina & """" & "  ORDER BY REC.Tipo")
            Set VarTotRec = FrmFrente.Banco.OpenRecordset("SELECT  Rec.Maquina,  First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE   (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) GROUP BY  Rec.Maquina,  REC.Tipo HAVING  Maquina = " & """" & LblMaquina & """" & "  ORDER BY REC.Tipo")
        
        VarTotRec.Requery
        
        If VarTotRec.RecordCount > 0 Then
            VarTotRec.MoveFirst
            Print #1, "------------------------------------------"
            While Not VarTotRec.EOF
                
                Select Case VarTotRec("Tipo")
                    Case 1
                        Print #1, "Recebimento Dinheiro ..... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                    Case 2
                        Print #1, "Recebimento Cheque ....... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 3
                        Print #1, "Recebimento Pix/Transf ... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 4
                        Print #1, "Recebimento Cartão Débito... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 5
                        Print #1, "Recebimento Cartão Crédito.. "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 6
                        Print #1, "Recebimento Venda a Prazo  "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                End Select
                
                
                Select Case VarTotRec("Tipo")
                    Case 11
                        Print #1, "Venda Dinheiro ........... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                    Case 12
                        Print #1, "Venda Cheque ............. "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 13
                        Print #1, "Venda Pix/Transf ......... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 14
                        Print #1, "Venda Cartão Débito....... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 15
                        Print #1, "Venda Cartão Crédito...... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 16
                        Print #1, "Venda a Prazo ............ "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                End Select

        
        
        '        If TotVda("Pgto") < 6 Then
        '            Entradas = Entradas + TotVda("SomaDeTotal")
        '        End If
                VarTotRec.MoveNext
            Wend
            Print #1, "------------------------------------------"
            
        End If

    End If


'    ' Movimentação Financeira
    Print #1, ""
    Print #1, "------------------------------------------"
'    Print #1, "Caixa Anterior.........(+)  "; Alinhar(MebValor, 12, 2)
    Print #1, "Venda de Produto.......(+)   "; Alinhar(CStr(TotVda), 12, 2)
    Print #1, "Entradas Diversas......(+)   "; Alinhar(CStr(TotCred), 12, 2)
    Print #1, "Recebimento............(+)   "; Alinhar(CStr(TotRec), 12, 2)
    Print #1, "                             ------------- "
    SubRec = TotVda + TotCred + TotRec
    Print #1, "Receita Bruta..........(=)   "; Alinhar(CStr(SubRec), 12, 2)
    Print #1, "Saidas Diversas........(-)   "; Alinhar(CStr(TotOut), 12, 2)
    Print #1, "Venda a Prazo..........(-)   "; Alinhar(CStr(TotVdaPrazo), 12, 2)
    Print #1, "                             ------------- "
    SubDeb = TotDeb + TotVdaPrazo
    TotCaixa = SubRec - (SubDeb + TotOut)
    Print #1, "Total do Caixa.........(=)   "; Alinhar(CStr(TotCaixa), 12, 2)
    Print #1, ""
    If CmdOk.Enabled = True Then
        MebDif = (TotCaixa - MebValor) - MebProximo
        Print #1, "Retirada do Fechamento.(-)   "; Alinhar(MebValor, 12, 2)
        Print #1, "Diferença............(+/-)   "; Alinhar(MebDif, 12, 2)
        Print #1, "Proximo Caixa.............   "; Alinhar(MebProximo, 12, 2)
        Print #1, ""
        Print #1, "Desconto de Venda vista( )   "; Alinhar(CStr(TotDescVda), 12, 2)
        
    Else
        MebDif = 0
    End If
'
    Print #1, "" & Nor
    Print #1, "<>"

    Close #1
    
    On Error GoTo erro
    LstCaixa.Clear
    Open App.Path & "\Caixa.txt" For Input As #1
    Do
        Line Input #1, Temp
        LstCaixa.AddItem Temp
'        Lista.Selected(Lista.ListCount - 1) = True
    Loop Until Temp = "<>"
    Close #1

erro:
If Err.Number = 62 Then
    Close #1
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbImprimir_Click()
    
    If MsgBox("Confirma impressão?", 36) = vbYes Then
'        Dim Fs2 As New FileSystemObject, Arq
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
'        Set Arq = Fs2.CreateTextFile(Porta, True)
        
        Fs2.CopyFile App.Path & "\caixa.txt", PortaNF
        
'        Fs2.Close
    End If

End Sub

Private Sub CmdOk_Click()
Dim sql As String
Dim PgtAnt As Byte
Dim Total As Currency
Dim texto As String
Dim TotVdaVista As Currency

' V= VISTA   R=RECEBIMENTO   S=DIFERENCIA/SALDO  F=FECHAMENTO

BtCalcula_Click

If MebValor = "" Then
    MebValor.SetFocus
    Exit Sub
End If

If MsgBox("Confirma Fechamento do Caixa ?", vbYesNo, App.Title) = vbYes Then
    
        Set VdaVista = FrmFrente.Banco.OpenRecordset("SELECT PED.Operacao, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)>  " & MebPed & " And (PED.CodPed)<= " & MebPed2 & "3)) GROUP BY PED.Operacao HAVING (((PED.Operacao)=1))")

        VdaVista.Requery
        
        If VdaVista.RecordCount <> 0 Then
            
            TotVdaVista = VdaVista!SomaDetotal ' - VdaVista!SomaDeDesconto
            texto = "VENDA VISTA"
            
            sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
            sql = sql + """" & LblMaquina & """" & ", "
            sql = sql & """" & texto & """" & " , #"
            sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
            sql = sql & Format(Time, "hh:mm:ss") & "# , "
            sql = sql & Numero(CStr(TotVdaVista), 2) & " ,"
            sql = sql & MebNroCxa & " ,"
            sql = sql & """" & "V" & """" & ")"
            FrmFrente.Banco.Execute sql
            Total = 0
       End If
'    Next
    texto = "RECEBIMENTO"
    sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    sql = sql + """" & LblMaquina & """" & ", "
    sql = sql & """" & texto & """" & " , #"
    sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    sql = sql & Format(Time, "hh:mm:ss") & "# , "
    sql = sql & Numero(CStr(TotRec), 2) & " ,"
    sql = sql & MebNroCxa & " ,"
    sql = sql & """" & "R" & """" & ")"
    FrmFrente.Banco.Execute sql
    Total = 0
    
    sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    sql = sql + """" & LblMaquina & """" & ", "
    sql = sql & """" & "FECHAMENTO" & """" & " , #"
    sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    sql = sql & Format(Time, "hh:mm:ss") & "# , "
    sql = sql & Numero(MebValor, 2) & " ,"
    sql = sql & MebNroCxa & " ,"
    sql = sql & """" & "F" & """" & ")"
    FrmFrente.Banco.Execute sql

    sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    sql = sql + """" & LblMaquina & """" & ", "
    sql = sql & """" & "DIFERENCA" & """" & " , #"
    sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    sql = sql & Format(Time, "hh:mm:ss") & "# , "
    sql = sql & Numero(MebDif, 2) & " ,"
    sql = sql & MebNroCxa & " ,"
    sql = sql & """" & "S" & """" & ")"
    FrmFrente.Banco.Execute sql

    sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    sql = sql + """" & LblMaquina & """" & ", "
    sql = sql & """" & "ABRE" & """" & " , #"
    sql = sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    sql = sql & Format(Time, "hh:mm:ss") & "# , "
    sql = sql & Numero(MebProximo, 2) & " ,"
    sql = sql & MebNroCxa + 1 & " ,"
    sql = sql & """" & "C" & """" & ")"
    FrmFrente.Banco.Execute sql


    sql = "INSERT INTO CONTCXA ( Pedido , Rec, Dup, Maquina ) VALUES ( "
    sql = sql & MebPed2 & " ,"
    sql = sql & MebRec2 & " ,"
    sql = sql & MebPcl2 & " ,"
    sql = sql & """" & LblMaquina & """" & " )"
    FrmFrente.Banco.Execute sql
    
    MebData.Text = Date
    MebValor = 0
    MebDif = 0
    MebProximo = 0
    Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa Where Maquina = " & """" & LblMaquina & """" & " Order by CodControle")
    Caixa.Requery
    Set CodVda = FrmFrente.Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
    CodVda.Requery
    
    Set Dup = FrmFrente.Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
    Dup.Requery
    
    Set Rec = FrmFrente.Banco.OpenRecordset("Select CodRec From Rec Order by CodRec")
    Rec.Requery
    
    If Caixa.RecordCount > 0 Then
        Caixa.MoveLast
        MebNroCxa = Caixa!CodControle
        MebPed = Caixa!Pedido
        MebPcl = Caixa!Dup
        MebRec = Caixa!Rec
        
    End If

    BtCalcula_Click
        
    MebValor.Text = 0
    MebData.SetFocus
End If

End Sub

Private Sub DbcHist_LostFocus()
DbcHist.Text = Format(DbcHist, ">")
End Sub


Private Sub Command1_Click()

   
    FrmRelTXT.Config App.Path & "\Caixa.Txt"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant


'For Each Componentes In DatDados
'    Componentes.DatabaseName = Caminho & "\Dados.mdb"
'Next

MebData.Text = Date
MebValor = 0
MebDif = 0
MebProximo = 0
LblMaquina = Trim(LimpaCampo(FrmFrente.Maquina))

 Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa Where Maquina = " & """" & LblMaquina & """" & "  Order by CodControle")
Caixa.Requery
Set CodVda = FrmFrente.Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
CodVda.Requery

Set Dup = FrmFrente.Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
Dup.Requery

Set Rec = FrmFrente.Banco.OpenRecordset("Select CodRec From Rec Order by CodRec")
Rec.Requery

If Caixa.RecordCount > 0 Then
    Caixa.MoveLast
    MebNroCxa = Caixa!CodControle
    MebPed = Caixa!Pedido
    MebPcl = Caixa!Dup
    MebRec = Caixa!Rec
Else
    Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa  Order by CodControle")
    Caixa.Requery
    If Caixa.RecordCount > 0 Then
        Caixa.MoveLast
        MebNroCxa = Caixa!CodControle
        MebPed = Caixa!Pedido
        MebPcl = Caixa!Dup
        MebRec = Caixa!Rec
    End If
    
End If


'Me.Width = 11700
'Me.Height = 7500

'e.Left = (MDIPrincipal.Width - Me.Width) / 2
'Me.Top = (MDIPrincipal.Height - Me.Height) / 6

'Me.WindowState = 2



End Sub


Private Sub MebNroCxa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
