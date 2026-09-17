VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormFchCxaTurno 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9585
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -720
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   26
      Top             =   7740
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
      ScaleWidth      =   9600
      TabIndex        =   24
      Top             =   0
      Width           =   9600
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechamento de Caixa por Turno"
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
         Width           =   5820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   6615
      TabIndex        =   7
      Top             =   675
      Width           =   2760
      Begin VB.CheckBox ChkPrd 
         Caption         =   "Discriminar Produtos"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   2835
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.Data DatDados 
         Caption         =   "Caixa"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   405
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT MAQUINA  FROM CXA GROUP BY MAQUINA ORDER BY MAQUINA"
         Top             =   6345
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox ChkDiscr 
         Caption         =   "Discriminar Pgto"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CommandButton CmbImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   675
         TabIndex        =   23
         Top             =   5535
         Width           =   1335
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Fecha Caixa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   675
         TabIndex        =   5
         Top             =   5040
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
         Left            =   135
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
         Left            =   225
         TabIndex        =   3
         Top             =   3420
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
         Left            =   135
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
         Left            =   225
         TabIndex        =   4
         Top             =   4005
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
         Left            =   225
         TabIndex        =   21
         Top             =   4590
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
      Begin MSDBCtls.DBCombo DbcCaixa 
         Bindings        =   "FormFchCxaTurno.frx":0000
         Height          =   360
         Left            =   225
         TabIndex        =   28
         Top             =   6300
         Width           =   2400
         _ExtentX        =   4233
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   6075
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   135
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Diferença"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   4365
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Retirada"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   3195
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Nro Caixa"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Proximo Caixa"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   3780
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
      Height          =   6780
      Left            =   180
      TabIndex        =   6
      Top             =   765
      Width           =   6225
   End
End
Attribute VB_Name = "FormFchCxaTurno"
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
Public VarTotRec As Recordset
Public TotVdaVista As Currency
Public TotRec As Currency


Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub BtCalcula_Click()
Dim Temp As String
Dim SQL As String
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
TotVdaVista = 0

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

'Sql = "SELECT First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ITP.Produto"
If DbcCaixa = "" Then
    SQL = "SELECT First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ITP.Produto"
Else
    SQL = "SELECT  First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor, Sum(Itp.VlrVdaDia) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto  FROM PED INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3))AND Ped.Contato = " & """" & Trim(DbcCaixa) & """" & " GROUP BY ITP.Produto"
End If

Set Venda = Banco.OpenRecordset(SQL)
Venda.Requery

If DbcCaixa = "" Then
    Set ArqDesc = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3))")
Else
    Set ArqDesc = Banco.OpenRecordset("SELECT  Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ") AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) AND Ped.contato = " & """" & Trim(DbcCaixa) & """")
End If
ArqDesc.Requery

'Set Venda = Banco.OpenRecordset("SELECT  First(PED.DataEmiss) AS PrimeiroDeDataEmiss, Sum(ITP.Qtd) AS SomaDeQtd, ITP.Produto, First(PRD.Descrição_Produto) AS PrimeiroDeDescrição_Produto, Sum(ITP.Valor) AS SomaDeValor FROM (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd GROUP BY ITP.Produto HAVING (((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & " ) AND ((First(PED.Operacao))=1 Or (First(PED.Operacao))=2 Or (First(PED.Operacao))=3))")

If DbcCaixa = "" Then
    Set Cxa = Banco.OpenRecordset("Select * From Cxa Where NroCaixa = " & MebNroCxa)
Else
    Set Cxa = Banco.OpenRecordset("Select * From Cxa Where NroCaixa = " & MebNroCxa & " AND Maquina = " & """" & Trim(DbcCaixa) & """")
End If
Cxa.Requery
'Set VdaPzo = Banco.OpenRecordset("SELECT PED.Operacao, First(PED.CodPed) AS PrimeiroDeCodPed, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED GROUP BY PED.Operacao HAVING (((PED.Operacao)=2) AND ((First(PED.CodPed))> " & MebPed & " And (First(PED.CodPed))<= " & MebPed2 & "))")
If DbcCaixa = "" Then
    Set VdaPzo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE ((PED.Operacao=2 OR Ped.Operacao =3 ) AND ((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & "))")
Else
    Set VdaPzo = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.Operacao)=2) AND ((PED.CodPed)> " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ")) AND Contato = " & """" & Trim(DbcCaixa) & """")
End If
VdaPzo.Requery

If DbcCaixa = "" Then
    Set Rec = Banco.OpenRecordset("SELECT REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE REC.Tipo < 10 AND  (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) Order By CodRec")
Else
    Set Rec = Banco.OpenRecordset("SELECT REC.Data, REC.Hora, ENT.Nome, REC.Valor, REC.Desconto, REC.CodRec, REC.Tipo FROM ENT INNER JOIN REC ON ENT.CodEntidade = REC.CodEnt WHERE REC.Tipo < 10 AND (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) AND Maquina = " & """" & Trim(DbcCaixa) & """" & " Order By CodRec")
End If
Rec.Requery

If VdaPzo.RecordCount > 0 Then
    If Not IsNull(VdaPzo!SomaDetotal) Then
        TotVdaPrazo = VdaPzo!SomaDetotal '- VdaPzo!SomaDeDesconto
    End If
End If

    Open App.Path & "\Caixa.txt" For Output As #1
    Print #1, ""
    Print #1, Centralizar("FECHAMENTO CAIXA", 48)
    Print #1, "================================================"
    If Cxa.RecordCount > 0 Then
        If CmdOk.Enabled = False Then
            Print #1, "Data: " & Cxa!Data & "   Hora: " & Cxa!Hora
            Print #1, ""
            Print #1, Tab(1); "Ref.Cxa: "; psCompString(1, MebNroCxa, 6) & " - " & Cxa!Data
        Else
            Print #1, "Data: " & Date & "   Hora: " & Time
            Print #1, ""
            Print #1, Tab(1); "Ref.Cxa: "; psCompString(1, MebNroCxa, 6) & " - " & MebData
        End If
    End If
    
    ' Produtos
    Print #1, ""
    If ChkPrd.value = 1 Then
        Print #1, "------------------------------------------------"
        If Venda.RecordCount > 0 Then
            Print #1, "Produto"; Tab(22); "Quantidade"; Tab(43); "Valor"
        
            With Venda
                .MoveFirst
                While Not .Eof
                    Print #1, Tab(0); Left(.Fields("PrimeiroDeDescrição_Produto"), 20); Tab(22); Alinhar(.Fields("SomaDeQtd"), 10, 2); Tab(36); Alinhar(.Fields("SomaDeValor"), 12, 2)
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
    Else
        If Venda.RecordCount > 0 Then
            With Venda
                .MoveFirst
                While Not .Eof
'                    Print #1, Tab(0); Left(.Fields("PrimeiroDeDescrição_Produto"), 20); Tab(22); Alinhar(.Fields("SomaDeQtd"), 10, 2); Tab(36); Alinhar(.Fields("SomaDeValor"), 12, 2)
                    TotVda = CCur(TotVda) + (.Fields("SomaDeTotal")) ' - .Fields("SomaDeDesconto"))
                    .MoveNext
                Wend
            End With
            TotVda = TotVda - ArqDesc.Fields("SomaDeDesconto")
        End If
    End If
    
    Print #1, "------------------------------------------------"
    If Cxa.RecordCount > 0 Then
        Print #1, "Data"; Tab(12); "Hora"; Tab(18); "Historico"; Tab(40); "Valor"; Tab(46); "D/C"
        
        With Cxa
            .MoveFirst
            While Not .Eof
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
            While Not .Eof
                Print #1, Tab(0); Left(.Fields("Hora"), 5); Tab(8); Left(.Fields("Nome"), 20); Tab(28); Alinhar(.Fields("Valor"), 10, 2); Tab(38); Alinhar(.Fields("Desconto"), 10, 2)
                TotRec = CCur(TotRec) + (.Fields("Valor"))
                .MoveNext
            Wend
        End With
    Else
        Print #1, "Nenhum recebimento de parcelas"
    End If


    If ChkDiscr.value Then
'        Set VarTotRec = Banco.OpenRecordset("SELECT  First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE  REC.Tipo > 10 AND (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) GROUP BY   REC.Tipo  ORDER BY REC.Tipo")
        If DbcCaixa = "" Then
            Set VarTotRec = Banco.OpenRecordset("SELECT   First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE   (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) GROUP BY  REC.Tipo  ORDER BY REC.Tipo")
 '           SQL = "SELECT  Rec.Maquina,  First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE   (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) GROUP BY  Rec.Maquina,  REC.Tipo  ORDER BY REC.Tipo"
        Else
            Set VarTotRec = Banco.OpenRecordset("SELECT  Rec.Maquina,  First(REC.Data) AS PrimeiroDeData, REC.Tipo, Sum(REC.Valor) AS SomaDeValor FROM REC WHERE   (((REC.CodRec)> " & MebRec & " And (REC.CodRec)<= " & MebRec2 & ")) GROUP BY  Rec.Maquina,  REC.Tipo HAVING  Maquina = " & """" & DbcCaixa & """" & "  ORDER BY REC.Tipo")
        End If
        VarTotRec.Requery
        
        If VarTotRec.RecordCount > 0 Then
            VarTotRec.MoveFirst
            Print #1, "------------------------------------------"
            While Not VarTotRec.Eof
                
                Select Case VarTotRec("Tipo")
                    Case 1
                        Print #1, "Recebimento Dinheiro ....... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                    Case 2
                        Print #1, "Recebimento Cheque ......... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 3
                        Print #1, "Recebimento Pix/Transf...... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 4
                        Print #1, "Recebimento Cartão Débito... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 5
                        Print #1, "Recebimento Cartão Crédito.. "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 6
                        Print #1, "Recebimento Venda a Prazo .. "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                End Select
                
                
                Select Case VarTotRec("Tipo")
                    Case 11
                        Print #1, "Venda Dinheiro ............. "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotVdaVista = VarTotRec("SomaDeValor")
                    Case 12
                        Print #1, "Venda Cheque ............... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 13
                        Print #1, "Venda Pix/Transf. .......... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 14
                        Print #1, "Venda Cartão Débito......... "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 15
                        Print #1, "Venda Cartão Crédito........ "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
                        TotOut = TotOut + VarTotRec("SomaDeValor")
                    Case 16
                        Print #1, "Venda Venda a Prazo ........ "; Alinhar(VarTotRec("SomaDeValor"), 12, 2)
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
    Print #1, "------------------------------------------------"
'    Print #1, "Caixa Anterior...............(+)  "; Alinhar(MebValor, 12, 2)
    Print #1, "Venda Produtos ..............(+)   "; Alinhar(CStr(TotVda), 12, 2)
    Print #1, "Entradas Diversas............(+)   "; Alinhar(CStr(TotCred), 12, 2)
    Print #1, "Recebimento Conta ...........(+)   "; Alinhar(CStr(TotRec), 12, 2)
    Print #1, "                                   ------------- "
    SubRec = TotVda + TotCred + TotRec
    Print #1, "Receita Bruta................(=)   "; Alinhar(CStr(SubRec), 12, 2)
    Print #1, "Saidas Diversas..............(-)   "; Alinhar(CStr(TotDeb), 12, 2)
    Print #1, "Outras Formas Recebimento....(-)   "; Alinhar(CStr(TotOut), 12, 2)
    
    Print #1, "                                   ------------- "
    SubDeb = TotDeb + TotVdaPrazo
    TotCaixa = SubRec - (SubDeb + TotOut)
    Print #1, "Total do Caixa...............(=)   "; Alinhar(CStr(TotCaixa), 12, 2)
    Print #1, ""
    
'    Print #1, "Venda a Prazo................(-)   "; Alinhar(CStr(TotVdaPrazo), 12, 2)
    
    If CmdOk.Enabled = True Then
        MebDif = (TotCaixa - MebValor) - MebProximo
        Print #1, "Retirada do Fechamento.......(-)   "; Alinhar(MebValor, 12, 2)
        Print #1, "Diferença..................(+/-)   "; Alinhar(MebDif, 12, 2)
        Print #1, "Proximo Caixa...................   "; Alinhar(MebProximo, 12, 2)
    Else
        MebDif = 0
    End If
    
    Print #1, ""
    Print #1, "     INFORMACOES ADICIONAIS "
    Print #1, ""
    
'    Print #1, "Venda Total....................( )   "; Alinhar(CStr(TotVda), 10, 2)
    Print #1, "Venda a Prazo..................( )   "; Alinhar(CStr(TotVdaPrazo), 10, 2)
    
    
'    Print #1, "Desconto de Venda vista..( )   "; Alinhar(CStr(TotDescVda), 10, 2)
'    Print #1, "Desconto de Recebimento..( )   "; Alinhar(CStr(TotDesc), 10, 2)
    
'
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
'        Dim Fs2 As New FileSystemObject, Arq
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
'        Set Arq = Fs2.CreateTextFile(Porta, True)
        
        Fs2.CopyFile App.Path & "\caixa.txt", Porta
    End If

End Sub

Private Sub CmdOk_Click()
Dim SQL As String
Dim PgtAnt As Byte
Dim Total As Currency
Dim Texto As String


' V= VISTA   R=RECEBIMENTO   S=DIFERENCIA/SALDO  F=FECHAMENTO

BtCalcula_Click

If MebValor = "" Then
    MebValor.SetFocus
    Exit Sub
End If

If MsgBox("Confirma Fechamento do Caixa ?", vbYesNo, App.Title) = vbYes Then
    
        Set VdaVista = Banco.OpenRecordset("SELECT PED.Operacao, Sum(PED.Total) AS SomaDeTotal, Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE (((PED.CodPed)>  " & MebPed & " And (PED.CodPed)<= " & MebPed2 & ")) GROUP BY PED.Operacao HAVING (((PED.Operacao)=1))")

        VdaVista.Requery
        
        If VdaVista.RecordCount <> 0 Then
            
'            TotVdaVista = VdaVista!SomaDetotal ' - VdaVista!SomaDeDesconto
            Texto = "VENDA VISTA"
            
            SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
            SQL = SQL + """" & Maquina & """" & ", "
            SQL = SQL & """" & Texto & """" & " , #"
            SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
            SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
            SQL = SQL & Numero2(CStr(TotVdaVista)) & " ,"
            SQL = SQL & MebNroCxa & " ,"
            SQL = SQL & """" & "V" & """" & ")"
            Banco.Execute SQL
            Total = 0
       End If
'    Next
    Texto = "RECEBIMENTO"
    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & Maquina & """" & ", "
    SQL = SQL & """" & Texto & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(CStr(TotRec)) & " ,"
    SQL = SQL & MebNroCxa & " ,"
    SQL = SQL & """" & "R" & """" & ")"
    Banco.Execute SQL
    Total = 0
    
    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & Maquina & """" & ", "
    SQL = SQL & """" & "FECHAMENTO" & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebValor) & " ,"
    SQL = SQL & MebNroCxa & " ,"
    SQL = SQL & """" & "F" & """" & ")"
    Banco.Execute SQL

    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & Maquina & """" & ", "
    SQL = SQL & """" & "DIFERENCA" & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebDif) & " ,"
    SQL = SQL & MebNroCxa & " ,"
    SQL = SQL & """" & "S" & """" & ")"
    Banco.Execute SQL

    SQL = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
    SQL = SQL + """" & Maquina & """" & ", "
    SQL = SQL & """" & "ABRE" & """" & " , #"
    SQL = SQL & Format(MebData, "mm/dd/yyyy") & "# , #"
    SQL = SQL & Format(Time, "hh:mm:ss") & "# , "
    SQL = SQL & Numero2(MebProximo) & " ,"
    SQL = SQL & MebNroCxa + 1 & " ,"
    SQL = SQL & """" & "C" & """" & ")"
    Banco.Execute SQL


    SQL = "INSERT INTO CONTCXA ( Pedido , Rec, Dup ) VALUES ( "
    SQL = SQL & MebPed2 & " ,"
    SQL = SQL & MebRec2 & " ,"
    SQL = SQL & MebPcl2 & " )"
    Banco.Execute SQL
    
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
    MebPed = Caixa!Pedido
    MebPcl = Caixa!Dup
    MebRec = Caixa!Rec
    
End If


'Me.Width = 11700
'Me.Height = 7500

Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

'Me.WindowState = 2

DatDados(1).DatabaseName = App.Path & "\dados.mdb"

End Sub


