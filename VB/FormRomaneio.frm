VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormRomaneio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10860
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   32
      Top             =   6660
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   30
      Top             =   0
      Width           =   13155
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Romaneio de Entrega"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   90
         TabIndex        =   31
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.TextBox TxtNro 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5625
      TabIndex        =   28
      Top             =   5715
      Width           =   1500
   End
   Begin VB.CommandButton Btok 
      Caption         =   "Grava Romaneio"
      Height          =   330
      Index           =   4
      Left            =   5625
      TabIndex        =   27
      Top             =   6165
      Width           =   1545
   End
   Begin VB.CommandButton Btok 
      Caption         =   "Produtos Produção"
      Height          =   330
      Index           =   2
      Left            =   3825
      TabIndex        =   26
      Top             =   6165
      Width           =   1545
   End
   Begin VB.CheckBox ChkFina 
      Caption         =   "Com valor financeiro"
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      Top             =   5655
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton Btok 
      Caption         =   "Pedidos Produção"
      Height          =   330
      Index           =   1
      Left            =   2025
      TabIndex        =   24
      Top             =   6165
      Width           =   1545
   End
   Begin VB.CommandButton BtConfirma 
      Caption         =   "&Confirma"
      Height          =   330
      Left            =   225
      TabIndex        =   11
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   2070
      TabIndex        =   10
      Top             =   855
      Width           =   8430
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   435
         TabIndex        =   23
         Top             =   1035
         Width           =   390
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   4335
         TabIndex        =   22
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         Height          =   195
         Left            =   4350
         TabIndex        =   21
         Top             =   750
         Width           =   390
      End
      Begin VB.Label LblNota 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4950
         TabIndex        =   20
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label LblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   19
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label LblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   18
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   285
         TabIndex        =   17
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   180
         Width           =   465
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4860
         TabIndex        =   14
         Top             =   1035
         Width           =   1305
      End
      Begin VB.Label LblCidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   13
         Top             =   465
         Width           =   45
      End
      Begin VB.Label LblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   900
         TabIndex        =   12
         Top             =   180
         Width           =   45
      End
   End
   Begin VB.TextBox TxtNota 
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
      TabIndex        =   1
      Tag             =   "N"
      Top             =   1530
      Width           =   1365
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
      Height          =   2580
      Left            =   225
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   2520
      Width           =   10275
   End
   Begin VB.TextBox TxtPedido 
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
      TabIndex        =   0
      Tag             =   "N"
      Top             =   900
      Width           =   1365
   End
   Begin VB.CommandButton Btok 
      Caption         =   "&Imprimir Romaneio"
      Height          =   330
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   6165
      Width           =   1545
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Venda"
      Height          =   375
      Left            =   8820
      TabIndex        =   6
      Top             =   6120
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
      Left            =   225
      TabIndex        =   8
      Top             =   5670
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14737632
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
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Romeio"
      Height          =   195
      Left            =   5625
      TabIndex        =   29
      Top             =   5490
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data da Entrega"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   5400
      Width           =   1170
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
      Left            =   8775
      TabIndex        =   5
      Top             =   5850
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Pedido"
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
      TabIndex        =   4
      Top             =   675
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nota"
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
      TabIndex        =   3
      Top             =   1305
      Width           =   435
   End
End
Attribute VB_Name = "FormRomaneio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arq As Recordset
Public VarOper As Byte

Private Sub BtConfirma_Click()

If LblPedido <> "" Then

    If TxtPedido.Text <> "" Then
        Lista.AddItem psCompString(1, LblPedido, 6) & "   " & LblData & "   " & psCompString(1, LblNota, 6) & "   " & Alinhar(LblValor, 10, 2)
        Lista.Selected(Lista.ListCount - 1) = True
        TxtPedido.Text = ""
        TxtPedido.SetFocus
    End If
    
    If TxtNota.Text <> "" Then
        Lista.AddItem psCompString(1, LblPedido, 6) & "   " & LblData & "   " & psCompString(1, LblNota, 6) & "   " & Alinhar(LblValor, 10, 2)
        Lista.Selected(Lista.ListCount - 1) = True
        TxtNota.Text = ""
        TxtNota.SetFocus
    End If
    
End If
End Sub

Private Sub BtOK_Click(Index As Integer)
Dim flag As Boolean
Dim y As Integer
Dim VarData As Date
Dim Formula As String
Dim SQL As String
Dim CamposGroup As String
Dim VarPrd As Recordset
Dim Carga As Recordset

flag = True

Formula = ""
For y = 0 To Lista.ListCount - 1
    If Lista.Selected(y) = True Then
        If flag Then
            Formula = "Ped.CodPed  = " & CLng(Left(Lista.List(y), 6))
            flag = False
        Else
            Formula = Formula & " or  Ped.CodPed = " & CLng(Left(Lista.List(y), 6))
            MebData = CDate(Mid(Lista.List(y), 10, 10))
        End If
    End If
Next

If Formula <> "" Then

'    If MsgBox("Deseja consolidar os pedido marcados", vbYesNo, App.Title) = vbYes Then
    If Index = 0 Then
        SqlDiv(1) = "SELECT Ped.CodPed, ENT.Nome, ENT.Cidade, Ped.NroNt, Ped.Total,Ped.Dataemiss FROM Ped INNER JOIN ENT ON Ped.Entidade = ENT.CodEntidade GROUP BY Ped.CodPed, ENT.Nome, Ent.Cidade, Ped.NroNt, Ped.Total, Ped.DataEmiss Having " & Formula
        SqlDiv(2) = "SELECT PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia FROM (Ped INNER JOIN ENT ON Ped.Entidade = ENT.CodEntidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON Ped.CodPed = ITP.Pedido Where " & Formula & " GROUP BY PRD.Descrição_Produto"
        If ChkFina.value = 0 Then
            FrmRelRomaneio.rpfDet(2).ForeColor = &HFFFFFF
            FrmRelRomaneio.rpfDet(5).ForeColor = &HFFFFFF
        End If
        FrmRelRomaneio.rpfDet(23).Caption = MebValor
        FrmRelRomaneio.rpCabTitulo.Caption = "Romaneio de Embarque"
        FrmRelRomaneio.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelRomaneio.Config
    ElseIf Index = 1 Then
            SQL = " SELECT ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, Ped.Desconto, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, Itp.Complemento, Itp.Referencia, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, Ped.Desconto, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, Itp.Complemento, Itp.Referencia, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ENT.Nome, PED.Operacao, ITP.CodItp "
            SQL = SQL & " Where " & Formula & CamposGroup & " ORDER BY Ped.NroPedido "
'            FrmRelConf.ReportField7.Mostrar = False
'            FrmRelConf.ReportField8.Mostrar = False
'            FrmRelConf.ReportField9.Mostrar = False
'            FrmRelConf.ReportField10.Mostrar = False
            
            Set varPesquisa(1) = Banco.OpenRecordset(SQL)
            Set FrmRelConf.Relatorio.Recordset = varPesquisa(1)
            If ChkFina = 0 Then
                FrmRelConf.ReportField22.Mostrar = False
                FrmRelConf.ReportField23.Mostrar = False
                FrmRelConf.ReportField32.Mostrar = False
                FrmRelConf.ReportField33.Mostrar = False
                FrmRelConf.ReportField26.Mostrar = False
                FrmRelConf.ReportField28.Mostrar = False
            End If
            FrmRelConf.rpCabTitulo.Caption = "Pedidos Específicos"
            FrmRelConf.rpCabSubTitulo.Caption = NomeEmpresa
            FrmRelConf.Config
        ElseIf Index = 2 Then
        
        
            SQL = "  SELECT ITP.Produto, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Avg(ITP.ValorUnit) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto , Sum([ITP].[QTD]*[ITP].[CUSTODIA]) AS Somadecusto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto "
            SQL = SQL & " Where " & Formula & CamposGroup & " ORDER BY PRD.DESCRIÇÃO_PRODUTO "
        
            Set varPesquisa(1) = Banco.OpenRecordset(SQL)
            Set FrmRelVda.Relatorio.Recordset = varPesquisa(1)
            If ChkFina.value = 0 Then
                FrmRelVda.ReportField11.Mostrar = False
                FrmRelVda.ReportField14.Mostrar = False
                FrmRelVda.ReportField8.Mostrar = False
                FrmRelVda.ReportField21.Mostrar = False
                FrmRelVda.ReportField7.Mostrar = False
                FrmRelVda.ReportField22.Mostrar = False
                FrmRelVda.ReportField20.Mostrar = False
                FrmRelVda.ReportField19.Mostrar = False
                FrmRelVda.ReportField16.Mostrar = False
                FrmRelVda.ReportField25.Mostrar = False
                FrmRelVda.ReportField17.Mostrar = False
                FrmRelVda.ReportField26.Mostrar = False
            
            End If
            FrmRelVda.rpCabTitulo.Caption = "Produtos Produção"
            FrmRelVda.rpCabSubTitulo.Caption = NomeEmpresa
            FrmRelVda.Config
            
        Else
            Set Carga = Banco.OpenRecordset("Select * From Carga")
        
            SQL = "  SELECT ITP.Produto, Itp.Pedido, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, Avg(ITP.ValorUnit) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto , Sum([ITP].[QTD]*[ITP].[CUSTODIA]) AS Somadecusto FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY ITP.Produto, Itp.Pedido, PRD.Descrição_Produto "
            SQL = SQL & " Where " & Formula & CamposGroup & " ORDER BY PRD.DESCRIÇÃO_PRODUTO "
            Set VarPrd = Banco.OpenRecordset(SQL)
            VarPrd.Requery
            If VarPrd.RecordCount > 0 Then
            
                VarPrd.MoveFirst
                While Not VarPrd.Eof
                    Carga.AddNew
                    Carga!NroCarga = TxtNro
                    Carga!Data = MebData
                    Carga!Pedido = VarPrd.Fields("Pedido")
                    Carga!Produto = VarPrd!Produto
                    Carga!Qtd = VarPrd!SomaDeQtd
                    Carga.Update
                    VarPrd.MoveNext
                Wend
                EscreveINI "CARGA", "NroCarga", TxtNro, App.Path & "\config.ini"
                MsgBox "Carga Gravada com sucesso !", vbInformation
                        
            End If
                                            
    End If
Else
    MsgBox "Não tem nenhum pedido marcado para consolidar", vbInformation, App.Title
End If

End Sub

Private Sub Form_Activate()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim x As Integer

MebValor = 0

Lista.Clear

Set Arq = Banco.OpenRecordset("SELECT PED.*, PED.CodPed, ENT.Nome, ENT.Cidade FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade ORDER BY PED.CodPed")

'If Arq.RecordCount > 0 Then
'    Arq.MoveFirst
'    While Not Arq.EOF
'        If Arq!NroNt = 0 And Arq!Operacao = VarOper Then
'            Lista.AddItem psCompString(1, Arq!CodPed, 6) & "   " & Arq!DataEmiss & "   " & Alinhar(Arq!Total, 10, 2)
'            Lista.Selected(Lista.ListCount - 1) = False
'        End If
'        Arq.MoveNext
'    Wend
'End If

If LerINI("CARGA", "NroCarga", App.Path & "\config.ini") <> "" Then
    TxtNro = LerINI("Carga", "NroCarga", App.Path & "\config.ini") + 1
End If

MebData = Date


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Lista_Click()
If Lista.Selected(Lista.ListIndex) = True Then
    MebValor = CCur(MebValor) + Right(Lista.Text, 12)
Else
    MebValor = CCur(MebValor) - Right(Lista.Text, 12)
End If
MebValor = Format(MebValor, "#,##0.00")
End Sub

Private Sub TxtNota_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
    Criterio = "Ped.NroNt = " & TxtNota
    Arq.FindFirst Criterio
    If Not Arq.NoMatch Then
        LblNome.Caption = Arq!Nome
        LblCidade.Caption = Arq!Cidade
        LblValor.Caption = Format(Arq!Total, "##,#00.00")
        LblPedido.Caption = Arq.Fields("Ped.CodPed")
        LblNota.Caption = Arq!NroNt
        LblData.Caption = Arq!DataEmiss
        BtConfirma.SetFocus
    Else

        MsgBox "Nota não encontrada !", vbInformation, App.Path
    End If
End If

End Sub

Private Sub TxtPedido_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
    If TxtPedido <> "" Then
        Criterio = "Ped.CodPed = " & TxtPedido
        Arq.FindFirst Criterio
        If Not Arq.NoMatch Then
            LblNome.Caption = Arq!Nome
            LblCidade.Caption = Arq!Cidade
            LblValor.Caption = Format(Arq!Total, "##,#00.00")
            LblPedido.Caption = Arq.Fields("Ped.CodPed")
            LblNota.Caption = Arq!NroNt
            LblData.Caption = Arq!DataEmiss
    
            BtConfirma.SetFocus
        Else
            MsgBox "Pedido não encontrado !", vbInformation, App.Path
        End If
    End If
End If
End Sub
