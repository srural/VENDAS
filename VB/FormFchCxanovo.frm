VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormFchCxanovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9555
   Begin VB.Frame Frame1 
      Height          =   6225
      Left            =   6660
      TabIndex        =   7
      Top             =   180
      Width           =   2760
      Begin VB.CommandButton CmbImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   810
         TabIndex        =   23
         Top             =   5265
         Width           =   1335
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Fecha Caixa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   810
         TabIndex        =   5
         Top             =   4590
         Width           =   1335
      End
      Begin VB.CommandButton BtCalcula 
         Caption         =   "&Calcula"
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   990
         Width           =   915
      End
      Begin MSMask.MaskEdBox MebData 
         Height          =   315
         Left            =   315
         TabIndex        =   0
         Top             =   405
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
         Top             =   2610
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
         Top             =   990
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
      Begin MSMask.MaskEdBox MebPed 
         Height          =   195
         Left            =   870
         TabIndex        =   8
         Top             =   1530
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
         Top             =   1770
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
         Top             =   2025
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
         Top             =   1530
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
         Top             =   1770
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
         Top             =   2025
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
         Top             =   3780
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Diferença"
         Height          =   195
         Left            =   315
         TabIndex        =   22
         Top             =   3555
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         Height          =   195
         Left            =   315
         TabIndex        =   20
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Retirada"
         Height          =   195
         Left            =   315
         TabIndex        =   19
         Top             =   2385
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Nro Caixa"
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Proximo Caixa"
         Height          =   195
         Left            =   315
         TabIndex        =   17
         Top             =   2970
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   285
         TabIndex        =   15
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Receb."
         Height          =   195
         Left            =   285
         TabIndex        =   14
         Top             =   2025
         Width           =   525
      End
   End
   Begin VB.ListBox LstCaixa 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   180
      TabIndex        =   6
      Top             =   270
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
Dim Criterio As String
TotVda = 0
TotDeb = 0
TotCred = 0
TotRec = 0
TotDup = 0
TotVdaPrazo = 0

Set CodVda = Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
CodVda.Requery

Set Dup = Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
Dup.Requery

Set Rec = Banco.OpenRecordset("Select CodRec From Rec Order by CodRec")
Rec.Requery


Criterio = "CodControle = " & MebNroCxa + 1
Caixa.FindFirst Criterio

If Not Caixa.NoMatch Then
    CmdOk.Enabled = False
'    MebNroCxa = Caixa!CodControle
'    Caixa.MoveLast
    MebPed2 = Caixa!pedido
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
MebPed = Caixa!pedido
MebPcl = Caixa!Dup
MebRec = Caixa!Rec

Set Venda = Banco.OpenRecordset("SELECT First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ITP.Produto")

'Set Venda = Banco.OpenRecordset("SELECT  First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor FROM (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd GROUP BY ITP.Produto HAVING (((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & " ) AND ((First(PED.Operacao))=1 Or (First(PED.Operacao))=2 Or (First(PED.Operacao))=3))")
Venda.Requery
Set Cxa = Banco.OpenRecordset("Select * From Cxa Where NroCaixa = " & MebNroCxa)
Cxa.Requery
'Set VdaPzo = Banco.OpenRecordset("SELECT PED.Operacao, First(PED.CodPed) AS PrimeiroDeCodPed, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED GROUP BY PED.Operacao HAVING (((PED.Operacao)=2) AND ((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & "))")

Set VdaPzo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.Operacao)=2) AND ((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & "))")
VdaPzo.Requery

Set Rec = Banco.OpenRecordset("SELECT REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) Order By CodRec")
Rec.Requery

If VdaPzo.RecordCount > 0 Then
    If Not IsNull(VdaPzo!SomaDetotal) Then
        TotVdaPrazo = VdaPzo!SomaDetotal - VdaPzo!SomaDeDesconto
    End If
End If

    Open App.Path & "\Caixa.txt" For Output As #1
    Print #1, ""
    Print #1, Centralizar("FECHAMENTO CAIXA", 48)
    Print #1, "================================================"
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
    Print #1, ""
    Print #1, "------------------------------------------------"
    If Venda.RecordCount > 0 Then
        Print #1, "Produto"; Tab(22); "Quantidade"; Tab(43); "Valor"
    
        With Venda
            .MoveFirst
            While Not .EOF
                Print #1, Tab(0); Left(.Fields("PrimeiroDeDescrição_Produto"), 20); Tab(22); Alinhar(.Fields("SomaDeQtd"), 10, 2); Tab(36); Alinhar(.Fields("SomaDeValor"), 12, 2)
                TotVda = CCur(TotVda) + (.Fields("SomaDeTotal") - .Fields("SomaDeDesconto"))
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Não tem produto vendido"
    End If
    
    ' Caixa
    Print #1, ""
    Print #1, "------------------------------------------------"
    If Cxa.RecordCount > 0 Then
        Print #1, "Data"; Tab(12); "Hora"; Tab(18); "Historico"; Tab(40); "Valor"; Tab(46); "D/C"
        
        With Cxa
            .MoveFirst
            While Not .EOF
                Print #1, Tab(0); .Fields("Data"); Tab(12); Left(.Fields("Hora"), 5); Tab(18); Left(.Fields("Historico"), 15); Tab(35); Alinhar(.Fields("Valor"), 10, 2); Tab(47); Left(.Fields("D_C"), 2)
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
    Print #1, ""
    Print #1, "------------------------------------------------"
    If Rec.RecordCount > 0 Then
        Print #1, "Hora"; Tab(8); "Nome"; Tab(33); "Valor"; Tab(41); "Desconto"
        
        With Rec
            .MoveFirst
            While Not .EOF
                Print #1, Tab(0); Left(.Fields("Hora"), 5); Tab(8); Left(.Fields("Nome"), 20); Tab(28); Alinhar(.Fields("Valor"), 10, 2); Tab(38); Alinhar(.Fields("Desconto"), 10, 2)
                TotRec = CCur(TotRec) + (.Fields("Valor"))
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Nenhum recebimento de parcelas"
    End If


'    ' Movimentação Financeira
    Print #1, ""
    Print #1, "------------------------------------------------"
'    Print #1, "Caixa Anterior...............(+)  "; Alinhar(MebValor, 12, 2)
    Print #1, "Venda de Produto.............(+)   "; Alinhar(CStr(TotVda), 12, 2)
    Print #1, "Entradas Diversas............(+)   "; Alinhar(CStr(TotCred), 12, 2)
    Print #1, "Recebimento..................(+)   "; Alinhar(CStr(TotRec), 12, 2)
    Print #1, "                                   ------------- "
    SubRec = TotVda + TotCred + TotRec
    Print #1, "Receita Bruta................(=)   "; Alinhar(CStr(SubRec), 12, 2)
    Print #1, "Saidas Diversas..............(-)   "; Alinhar(CStr(TotDeb), 12, 2)
    Print #1, "Venda a Prazo................(-)   "; Alinhar(CStr(TotVdaPrazo), 12, 2)
    Print #1, "                                   ------------- "
    SubDeb = TotDeb + TotVdaPrazo
    TotCaixa = SubRec - SubDeb
    Print #1, "Total do Caixa...............(=)   "; Alinhar(CStr(TotCaixa), 12, 2)
    Print #1, ""
    If CmdOk.Enabled = True Then
        MebDif = (TotCaixa - MebValor) - MebProximo
        Print #1, "Retirada do Fechamento.......(-)   "; Alinhar(MebValor, 12, 2)
        Print #1, "Diferença..................(+/-)   "; Alinhar(MebDif, 12, 2)
        Print #1, "Proximo Caixa...................   "; Alinhar(MebProximo, 12, 2)
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

Private Sub CmbImprimir_Click()
    
    If MsgBox("Confirma impressão?", 36) = vbYes Then
'        Dim Fs2 As New FileSystemObject, Arq
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
'        Set Arq = Fs2.CreateTextFile(Porta, True)
        
        Fs2.CopyFile App.Path & "\caixa.txt", Porta
    End If


End Sub

Private Sub CmdOk_Click()
Dim Sql As String
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

If MsgBox("Confirma Fechamento do Caixa ?", vbYesNo, App.Title) = vbYes Then
    
        Set VdaVista = Banco.OpenRecordset("SELECT PED.Operacao, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)>  " & MebPed & " And (PED.CodPed)<= " & MebPed2 & "3)) GROUP BY PED.Operacao HAVING (((PED.Operacao)=1))")

        VdaVista.Requery
        
        If VdaVista.RecordCount <> 0 Then
            
            TotVdaVista = VdaVista!SomaDetotal - VdaVista!SomaDeDesconto
            Texto = "VENDA VISTA"
            
            Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
            Sql = Sql + """" & Maquina & """" & ", "
            Sql = Sql & """" & Texto & """" & " , #"
            Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
            Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
            Sql = Sql & Numero(CStr(TotVdaVista)) & " ,"
            Sql = Sql & MebNroCxa & " ,"
            Sql = Sql & """" & "V" & """" & ")"
            Banco.Execute Sql
            Total = 0
       End If
'    Next
    Texto = "RECEBIMENTO"
    Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    Sql = Sql + """" & Maquina & """" & ", "
    Sql = Sql & """" & Texto & """" & " , #"
    Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
    Sql = Sql & Numero(CStr(TotRec)) & " ,"
    Sql = Sql & MebNroCxa & " ,"
    Sql = Sql & """" & "R" & """" & ")"
    Banco.Execute Sql
    Total = 0
    
    Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    Sql = Sql + """" & Maquina & """" & ", "
    Sql = Sql & """" & "FECHAMENTO" & """" & " , #"
    Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
    Sql = Sql & Numero(MebValor) & " ,"
    Sql = Sql & MebNroCxa & " ,"
    Sql = Sql & """" & "F" & """" & ")"
    Banco.Execute Sql

    Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    Sql = Sql + """" & Maquina & """" & ", "
    Sql = Sql & """" & "DIFERENCA" & """" & " , #"
    Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
    Sql = Sql & Numero(MebDif) & " ,"
    Sql = Sql & MebNroCxa & " ,"
    Sql = Sql & """" & "S" & """" & ")"
    Banco.Execute Sql

    Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    Sql = Sql + """" & Maquina & """" & ", "
    Sql = Sql & """" & "ABRE" & """" & " , #"
    Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
    Sql = Sql & Numero(MebProximo) & " ,"
    Sql = Sql & MebNroCxa + 1 & " ,"
    Sql = Sql & """" & "C" & """" & ")"
    Banco.Execute Sql


    Sql = "INSERT INTO CONTCXA ( Pedido , Rec, Dup ) VALUES ( "
    Sql = Sql & MebPed2 & " ,"
    Sql = Sql & MebRec2 & " ,"
    Sql = Sql & MebPcl2 & " )"
    Banco.Execute Sql
    
    MebData.Text = Date
    MebValor = 0
    MebDif = 0
    MebProximo = 0
    Set Caixa = Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
    Caixa.Requery
    Set CodVda = Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
    CodVda.Requery
    
    Set Dup = Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
    Dup.Requery
    
    Set Rec = Banco.OpenRecordset("Select CodRec From Rec Order by CodRec")
    Rec.Requery
    
    If Caixa.RecordCount > 0 Then
        Caixa.MoveLast
        MebNroCxa = Caixa!CodControle
        MebPed = Caixa!pedido
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
Set Caixa = Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
Caixa.Requery
Set CodVda = Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
CodVda.Requery

Set Dup = Banco.OpenRecordset("Select CodDup From Dup Order by CodDup")
Dup.Requery

Set Rec = Banco.OpenRecordset("Select CodRec From Rec Order by CodRec")
Rec.Requery

If Caixa.RecordCount > 0 Then
    Caixa.MoveLast
    MebNroCxa = Caixa!CodControle
    MebPed = Caixa!pedido
    MebPcl = Caixa!Dup
    MebRec = Caixa!Rec
    
End If


'Me.Width = 11700
'Me.Height = 7500

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

'Me.WindowState = 2



End Sub


