VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormConsolida 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6870
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OK 
      Caption         =   ">>"
      Height          =   330
      Left            =   10935
      TabIndex        =   17
      Top             =   900
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   12
      Top             =   6525
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11940
      TabIndex        =   10
      Top             =   0
      Width           =   11940
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Consolida Pedidos"
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
         TabIndex        =   11
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1440
      Width           =   11580
   End
   Begin VB.TextBox TxtCodigo 
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
      Left            =   135
      TabIndex        =   4
      Tag             =   "N"
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox TxtNome 
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
      Left            =   1305
      TabIndex        =   3
      Tag             =   "N"
      Top             =   900
      Width           =   5685
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Confirma "
      Height          =   330
      Left            =   1710
      TabIndex        =   0
      Top             =   5940
      Width           =   1815
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Venda"
      Height          =   375
      Left            =   9945
      TabIndex        =   6
      Top             =   5895
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebData 
      Height          =   330
      Left            =   180
      TabIndex        =   8
      Top             =   5940
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14737632
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   7200
      TabIndex        =   13
      Top             =   930
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Format          =   88473601
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   8955
      TabIndex        =   14
      Top             =   915
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Format          =   88473601
      CurrentDate     =   36641
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   7200
      TabIndex        =   16
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   8955
      TabIndex        =   15
      Top             =   705
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Operação"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   5670
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   9900
      TabIndex        =   5
      Top             =   5625
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   135
      TabIndex        =   2
      Top             =   675
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   1350
      TabIndex        =   1
      Top             =   675
      Width           =   555
   End
End
Attribute VB_Name = "FormConsolida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Arq As Recordset
Public VarOper As Byte

Private Sub BtOK_Click()
Dim flag As Boolean
Dim y As Integer
Dim VarData As Date
Dim Formula As String
Dim VarCodPed As Double
Dim pedidos As Recordset
Dim VarItp  As Recordset
Dim VarItens  As Recordset
Dim Criterio As String
Dim VarObs As String

flag = True
VarObs = ""

Formula = ""
For y = 0 To Lista.ListCount - 1
    If Lista.Selected(y) = True Then
        Criterio = "CodPed = " & CLng(Left(Lista.List(y), 6))
        Arq.FindFirst Criterio
            If Not Arq.NoMatch Then
                If Arq!NroNt <> "" And Arq!NroNt <> 0 Then
                    VarObs = VarObs & " / " & Arq!NroNt
                    SQL = " INSERT INTO NFE ( CodPed, NroChave, CodNota ) SELECT " & Arq!CodPed & " AS Expr1, " & """" & Arq!sat & """" & " AS Expr2, 0 AS Expr3 "
                    Banco.Execute SQL
                    
                ElseIf Arq!sat <> "" Then
                    SQL = " INSERT INTO NFE ( CodPed, NroChave, CodNota ) SELECT " & Arq!CodPed & " AS Expr1, " & Arq!sat & " AS Expr2, 0 AS Expr3 "
                    Banco.Execute SQL
                End If
            End If
        
        If flag Then
            Formula = "Itp.Pedido  = " & CLng(Left(Lista.List(y), 6))
            flag = False
            VarCodPed = CLng(Left(Lista.List(y), 6))
                    
        Else
            Formula = Formula & " or  Itp.Pedido = " & CLng(Left(Lista.List(y), 6))
'            MebData = CDate(Mid(Lista.List(y), 10, 10))
        End If
    End If
Next

If Formula <> "" Then
    If MsgBox("Deseja consolidar os pedidos marcados", vbYesNo, App.Title) = vbYes Then
    
    
        SQL = " INSERT INTO PED  (Operacao, Entidade, DataEmiss, Total, Desconto, Vendedor, Entregue, Transportadora, Comissao, IcmSubst, Ipi, Icms, ValorFrete, NroPedido, SubTotal, TotalPrd, FormaPgto, Obs ) SELECT " & "4" & " ," & TxtCodigo & ", PED.DataEmiss, " & Numero2(MebValor.Text) & ", PED.Desconto, PED.Vendedor, False AS Expr1, PED.Transportadora, PED.Comissao, PED.IcmSubst, PED.Ipi, PED.Icms, PED.ValorFrete, PED.NroPedido, PED.SubTotal, PED.TotalPrd, PED.FormaPgto, " & """" & VarObs & """" & " FROM PED WHERE PED.CodPed = " & VarCodPed
        Banco.Execute SQL
        
        Set pedidos = Banco.OpenRecordset("Select CodPed From Ped Order by CodPed")
        pedidos.Requery
        If pedidos.RecordCount > 0 Then
            pedidos.MoveLast
            VarCodPed = pedidos!CodPed
            SQL = " UPDATE NFE SET NFE.CodNota = " & VarCodPed & " WHERE CodNota = 0 "
            Banco.Execute SQL
        End If
        
        
'        Set Pedidos = Banco.OpenRecordset("Select * From Itp Where Pedido = " & GridPed.Columns(2))
'        Pedidos.Requery
                        
        If MsgBox("Agrupar os produtos ??", vbYesNo, App.Title) = vbYes Then

            Set VarItp = Banco.OpenRecordset("SELECT ITP.Produto, First(ITP.Complemento) AS PrimeiroDeComplemento, Sum(ITP.Qtd) AS SomaDeQtd, Avg(ITP.ValorUnit) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.Desconto) AS SomaDeDesconto, Avg(ITP.CustoDia) AS MédiaDeCustoDia, First(ITP.Loja) AS PrimeiroDeLoja, First(False) AS Expr1, Avg(IIf(IsNull(ITP.AliqIcms),0,ITP.AliqIcms)) AS Expr2, Avg(IIf(IsNull(ITP.AliqIPI),0,ITP.AliqIPI)) AS Expr3, Avg(IIf(IsNull(ITP.ValorIPI),0,ITP.ValorIPI)) AS Expr4, Avg(IIf(IsNull(Itp.Cst),0,Itp.Cst)) AS Expr5, Avg(IIf(IsNull(ITP.Comissao),0,ITP.Comissao)) AS Expr6, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Avg(ITP.VendaDia) AS MédiaDeVendaDia FROM ITP WHERE  ((" & Formula & ")) GROUP BY  ITP.Produto")
            Set VarItens = Banco.OpenRecordset(" Select * From Itp ")
            
            If VarItp.RecordCount > 0 Then
                VarItp.MoveFirst
                While Not VarItp.Eof
                    VarItens.AddNew
                    VarItens!Pedido = VarCodPed
                    VarItens!Produto = VarItp!Produto
                    VarItens!Qtd = VarItp!SomaDeQtd
                    VarItens!ValorUnit = VarItp!MédiaDeValorUnit
                    VarItens!Valor = VarItp!SomaDeValor
                    VarItens!Desconto = VarItp!SomaDeDesconto
                    VarItens!CustoDia = VarItp!MédiaDeCustoDia
                    VarItens!Loja = VarItp!PrimeiroDeLoja
                    VarItens!Entregue = True
                    VarItens!AliqIcms = VarItp!Expr2
                    VarItens!AliqIPI = VarItp!Expr3
                    VarItens!ValorIPI = VarItp!Expr4
                    VarItens!Cst = VarItp!Expr5
                    VarItens!Comissao = Expr6
                    VarItens!VlrVdaDia = VarItp!SomaDeVlrVdaDia
                    VarItens!VendaDia = VarItp!MédiaDeVendaDia
                    VarItens.Update
                    VarItp.MoveNext
                Wend
            End If
        Else
            SQL = " INSERT INTO ITP (Pedido, Produto, Complemento, Qtd, ValorUnit, Valor, Desconto, CustoDia, Loja, Entregue, AliqIcms, AliqIPI, ValorIPI, CST, Comissao , VlrVdaDia, VendaDia ) SELECT " & VarCodPed & ",  ITP.Produto, ITP.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.CustoDia, ITP.Loja, False AS Expr1, iif(isnull(ITP.AliqIcms),0,ITP.AliqIcms), iif(isnull(ITP.AliqIPI),0,ITP.AliqIPI), IIF(ISNULL(ITP.ValorIPI),0,ITP.ValorIPI), iif(ISNULL(Itp.Cst),0,Itp.Cst), IIF(ISNULL(ITP.Comissao),0,ITP.Comissao), Itp.VlrVdaDia, Itp.VendaDia FROM ITP WHERE " & Formula
            Banco.Execute SQL
        End If
                
        MsgBox "Pedido copiado com sucesso !"
        FrmPedidos.DatDados(5).Refresh
                        
    
    End If
Else
    MsgBox "Não tem nenhum pedido marcado para consolidar", vbInformation, App.Title
End If

Unload Me

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim X As Integer

TxtNome.Text = FrmPedidos.DbcDados(3).Text
TxtCodigo.Text = FrmPedidos.DbcDados(3).BoundText

MebData = Date
MebValor = 0

Lista.Clear
Set Arq = FrmPedidos.DatDados(5).Recordset.Clone
If Arq.RecordCount > 0 Then
    Arq.MoveLast
    While Not Arq.Bof
       If Lista.ListCount < 30000 Then
            If Arq!Operacao = 1 Or Arq!Operacao = 2 Or Arq!Operacao = 5 Then
                Lista.AddItem psCompString(1, Arq!CodPed, 6) & "   " & Arq!DataEmiss & "   " & Alinhar(Arq!Total, 10, 2) & "   " & psCompString(1, Arq!NroNt, 6) & "   " & Arq!sat
                Lista.Selected(Lista.ListCount - 1) = False
            End If
        End If
        Arq.MovePrevious
    Wend
     
End If

DtData1 = Date
DtData2 = Date

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Lista_Click()

'If Lista.Selected(Lista.ListIndex) = True Then
'    MebValor = CCur(MebValor) + Right(Lista.Text, 12)
'Else
'    MebValor = CCur(MebValor) - Right(Lista.Text, 12)
'End If
'MebValor = Format(MebValor, "#,##0.00")

End Sub

Private Sub Lista_ItemCheck(Item As Integer)

MebValor = 0
For y = 0 To Lista.ListCount - 1
    If Lista.Selected(y) = True Then
        MebValor = CCur(MebValor) + Mid(Lista.List(y), 23, 12)
    End If
Next
MebValor = Format(MebValor, "#,##0.00")


End Sub

Private Sub OK_Click()
Dim Componentes As Variant
Dim Criterio As String
Dim X As Integer

TxtNome.Text = FrmPedidos.DbcDados(3).Text
TxtCodigo.Text = FrmPedidos.DbcDados(3).BoundText

MebData = Date
MebValor = 0

Lista.Clear
    
Set Arq = Banco.OpenRecordset("SELECT Ped.PrevEntrega, Ped.Sat, Ped.Cfo,Ped.Entregue, Ped.CondPgto, Ped.Entidade, Ped.NroPedido, Ped.Operacao, iif(Ped.Operacao=0,'Entrada',iif(Ped.Operacao=1,'Venda Vista',iif(Ped.Operacao=2,'Venda Prazo',iif(Ped.Operacao=13,'Ped.Entrega',iif(Ped.Operacao=4,'Orçamento','Outros'))))) as TipoOper, PED.CodPed, PED.DataEmiss, PED.DtFat, Ped.DtSaida, Ped.DtSaida, Ped.NroNt, PED.Total, PED.Desconto FROM PED WHERE PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# AND PED.DataEmiss <=#" & Format(DtData2, "mm/dd/yyyy") & "# AND Entidade = " & TxtCodigo.Text & " ORDER BY PED.CodPed")
Arq.Requery
If Arq.RecordCount > 0 Then
    Arq.MoveLast
    While Not Arq.Bof
       If Lista.ListCount < 30000 Then
            If Arq!Operacao = 1 Or Arq!Operacao = 2 Or Arq!Operacao = 5 Then
                Lista.AddItem psCompString(1, Arq!CodPed, 6) & "   " & Arq!DataEmiss & "   " & Alinhar(Arq!Total, 10, 2) & "   " & psCompString(1, Arq!NroNt, 6) & "   " & Arq!sat
                Lista.Selected(Lista.ListCount - 1) = True
            End If
        End If
        Arq.MovePrevious
    Wend
     
End If

End Sub
