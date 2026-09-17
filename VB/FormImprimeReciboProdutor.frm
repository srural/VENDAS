VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormImprimeReciboProdutor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7620
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   6855
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -540
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   25
      Top             =   7290
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
      ScaleWidth      =   6855
      TabIndex        =   23
      Top             =   0
      Width           =   6855
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Recibos e Protocolos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   24
         Top             =   45
         Width           =   6045
      End
   End
   Begin VB.CheckBox ChkNota 
      Caption         =   "&Nota"
      Height          =   240
      Left            =   3330
      TabIndex        =   22
      Top             =   6390
      Width           =   915
   End
   Begin VB.CheckBox ChkFunRural 
      Caption         =   "Descontar 2,3% do FunRural"
      Height          =   240
      Left            =   705
      TabIndex        =   21
      Top             =   6420
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.Frame FrameRecibos 
      Caption         =   "Condição de Seleção dos Recibos"
      Enabled         =   0   'False
      Height          =   2205
      Left            =   705
      TabIndex        =   17
      Top             =   1710
      Width           =   5205
      Begin VB.Data DatDados 
         Connect         =   "Access"
         DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   2700
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1620
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FormImprimeReciboProdutor.frx":0000
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   465
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox ListaEnt 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   885
         Width           =   4920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa Nome"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame FrameRecFormiga 
      Caption         =   "Dados para Emissão do Recibo do Comprador"
      Enabled         =   0   'False
      Height          =   2235
      Left            =   705
      TabIndex        =   15
      Top             =   4050
      Width           =   5205
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   945
         Left            =   930
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1110
         Width           =   4125
      End
      Begin MSMask.MaskEdBox MebValor 
         DataField       =   "Custo"
         Height          =   315
         Left            =   930
         TabIndex        =   7
         Tag             =   "S"
         Top             =   735
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebCpf 
         DataField       =   "Custo"
         Height          =   315
         Left            =   930
         TabIndex        =   6
         Tag             =   "S"
         Top             =   360
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&PF"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   300
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs.:"
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   1470
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condições do Relatório"
      Height          =   855
      Left            =   705
      TabIndex        =   14
      Top             =   720
      Width           =   5205
      Begin VB.OptionButton opTipoRel 
         Caption         =   "Declaração de Fornecedor"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   3
         Top             =   510
         Width           =   2625
      End
      Begin VB.OptionButton opTipoRel 
         Caption         =   "Recibos Compradores"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   1
         Top             =   510
         Width           =   2625
      End
      Begin VB.OptionButton opTipoRel 
         Caption         =   "Protocolos"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton opTipoRel 
         Caption         =   "Recibos Fornecedor"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   1785
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4665
      TabIndex        =   11
      Top             =   6690
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   315
      Left            =   975
      TabIndex        =   9
      Top             =   6750
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   315
      Left            =   2955
      TabIndex        =   10
      Top             =   6750
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Até"
      Height          =   195
      Left            =   2595
      TabIndex        =   13
      Top             =   6810
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "De:"
      Height          =   195
      Left            =   675
      TabIndex        =   12
      Top             =   6810
      Width           =   255
   End
End
Attribute VB_Name = "FormImprimeReciboProdutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub chReciboFormiga_Click()
    If chReciboFormiga.value = 1 Then
        FrameRecFormiga.Enabled = True
    Else
        FrameRecFormiga.Enabled = False
    End If
    ListaEnt.Clear
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim SQL As String
Dim sql2 As String
Dim SqlCampos As String
Dim SqlWhere As String
Dim SqlOrdem As String
Dim Formula As String
Dim Flag As Boolean
Dim Tabela As String
Dim Campo As String

SQL = ""
sql2 = ""
SqlCampos = ""
SqlWhere = ""
SqlOrdem = ""
Formula = ""
Tabela = ""
Campos = ""

    
    Dim Componente As Variant
    For Each Componentes In DatDados
        Componentes.DatabaseName = Caminho & "\Dados.mdb"
    Next
    
    If opTipoRel(1).value = True Or opTipoRel(2).value = True Or opTipoRel(3).value = True Then
    
         Flag = True
         If ListaEnt.ListCount > 0 Then
             If opTipoRel(1).value = True Then
                Campo = "Ent.CodEntidade"
                Tabela = "Ent"
             ElseIf opTipoRel(2).value = True Then
                Campo = "Ped.Vendedor"
                Tabela = "Ent"
                 FrmRelReciboProdutor.cabFormiga.Mostrar = True
                 FrmRelReciboProdutor.titCpf.Caption = mebCpf.Text
                 FrmRelReciboProdutor.titNome.Caption = Mid(ListaEnt.List(0), 10, Len(ListaEnt.List(0)))
                 FrmRelReciboProdutor.titNome2.Caption = Mid(ListaEnt.List(0), 10, Len(ListaEnt.List(0)))
                 FrmRelReciboProdutor.titValor.Caption = "R$ " & Format(MebValor, "#,#0.00")
                 FrmRelReciboProdutor.rpObs.Caption = TxtObs.Text
                 FrmRelReciboProdutor.titExtenso.Caption = "((" & Format(PassaExtenso(MebValor), ">") & "))" & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
                 
             ElseIf opTipoRel(3).value = True Then
                Campo = "Ent.CodEntidade"
                Tabela = "Vendedores"
            End If

            
             For y = 0 To ListaEnt.ListCount - 1
                 If ListaEnt.Selected(y) = True Then
                     If Flag Then
                         Formula = "And (" & Campo & " = " & CLng(Left(ListaEnt.List(y), 6))
                         Flag = False
                     Else
                         Formula = Formula & " or " & Campo & " = " & CLng(Left(ListaEnt.List(y), 6)) & " "
                     End If
                 End If
             Next
                 If ListaEnt.ListCount > 0 Then Formula = Formula & ") "
                 
                 'Sql = SqlCampos & SqlWhere & Formula & SqlOrdem
         ElseIf opTipoRel(2).value = True And ListaEnt.ListCount <= 0 Then
                MsgBox "É necessário selecionar um vendedor!", vbInformation
                DbcDados(0).SetFocus
                Exit Sub
         Else
                Tabela = "Ent"
                 'Sql = SqlCampos & SqlWhere & SqlOrdem
         End If
         
         FrmRelReciboProdutor.FunRural = ChkFunRural.value
        
         SqlCampos = "SELECT PED.Operacao, Ped.Vendedor, ENT.CodEntidade, ENT.Nome, ENT.CPF, Vendedores.CodEntidade, Vendedores.Nome, Vendedores.CPF,Ped.DataEmiss, PED.NroNt, PED.DtFat, Ped.CodPed, PED.Total FROM ENT AS Vendedores INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON Vendedores.CodEntidade = PED.Vendedor GROUP BY PED.Operacao, Ped.Vendedor, ENT.CodEntidade, ENT.Nome, ENT.CPF, Vendedores.CodEntidade, Vendedores.Nome, Vendedores.CPF, PED.NroNt, Ped.DataEmiss, PED.DtFat, PED.CodPed, Ped.Total "
         If ChkNota.value = 1 Then
             SqlWhere = " Having PED.Operacao=0 AND PED.NroNt Is Not Null and PED.NroNt<>0 and PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DtFat)<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# "
         Else
             SqlWhere = " Having PED.Operacao=0 AND  PED.DataEmiss >=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DataEmiss )<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# "
         End If
         SqlOrdem = " Order by " & Tabela & ".Nome "
         
         
         SQL = SqlCampos & SqlWhere & Formula & SqlOrdem
         
         If Tabela = "Vendedores" Then
             sql2 = "SELECT ENT.CodEntidade, Ped.Vendedor, Sum(Ped.Total) AS SomaDeTotal FROM ENT INNER JOIN Ped ON ENT.CodEntidade = Ped.Entidade Where PED.NroNt Is Not Null and PED.NroNt<>0 and PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DtFat)<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# GROUP BY ENT.CodEntidade, PED.Vendedor, PED.Operacao HAVING (((PED.Operacao)=0)) " & Formula
         Else
             If ChkNota.value = 1 Then
                sql2 = "SELECT PED.Entidade, Sum(PED.Total) AS SomaDeTotal from Ped Where PED.NroNt Is Not Null and PED.NroNt<>0 and PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DtFat)<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# GROUP BY PED.Entidade, PED.Operacao HAVING (((PED.Operacao)=0))"
             Else
                sql2 = "SELECT PED.Entidade, Sum(PED.Total) AS SomaDeTotal from Ped Where  PED.DataEmiss >=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DataEmiss )<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "# GROUP BY PED.Entidade, PED.Operacao HAVING (((PED.Operacao)=0))"
             End If
             
         End If
        
         Set varPesquisa(1) = Banco.OpenRecordset(SQL)
         Set varPesquisa(2) = Banco.OpenRecordset(sql2)
         Set FrmRelReciboProdutor.Relatorio.Recordset = varPesquisa(1)
         FrmRelReciboProdutor.Tabela = Tabela
         If opTipoRel(3).value = True Then FrmRelReciboProdutor.rpTitulo.Caption = "DECLARAÇÃO"
         FrmRelReciboProdutor.SqlCondicao = SqlWhere & SqlOrdem
         FrmRelReciboProdutor.FunRural = ChkFunRural.value
         FrmRelReciboProdutor.Config
    ElseIf opTipoRel(0).value = True Then
         SqlCampos = "SELECT PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat, ENT.CPF, PED.CodPed, PED.Total FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat, ENT.CPF, PED.CodPed, PED.Total "
         'SqlCampos = "SELECT PED.Operacao, ENT.Nome, PED.NroNt, PED.DtFat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade "
         SqlWhere = "Having PED.Operacao=0 AND PED.NroNt Is Not Null and PED.NroNt<>0 and PED.DtFat>=#" & Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1) & "# And (PED.DtFat)<=#" & Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2) & "#"
         SqlOrdem = "Order by Nome"
         
         SQL = SqlCampos & SqlWhere & SqlOrdem
        
         FrmRelPedProtocolo.rpCabTitulo.Caption = "Relatório de Protocolos"
         FrmRelPedProtocolo.rpCabSubTitulo.Caption = NomeEmpresa
         FrmRelPedProtocolo.SubTitulo.Caption = "Período: " & DtData1 & " à " & DtData2
        
         Set varPesquisa(1) = Banco.OpenRecordset(SQL)
         Set FrmRelPedProtocolo.Relatorio.Recordset = varPesquisa(1)
         FrmRelPedProtocolo.Config
         
    End If
    
End Sub

Private Sub DbcDados_GotFocus(Index As Integer)
Me.KeyPreview = False
End Sub

Private Sub DbcDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 9 Or KeyCode = 13 Then
    If Index = 0 Then
    
        If opTipoRel(2).value = True Then
            If ListaEnt.ListCount > 0 Then
                ListaEnt.Clear
            End If
        End If
    
        If DbcDados(0).Text = "" Then
            DbcDados(0).SetFocus
            Exit Sub
        End If
        Dim Criterio As String
        If IsNumeric(DbcDados(0).Text) Then
           Criterio = "CodEntidade = " & DbcDados(0)
           DatDados(0).Recordset.FindFirst Criterio
           If DatDados(0).Recordset.NoMatch Then
               MsgBox "Entidade não encontrada !!!", 16
               DbcDados(0).Text = ""
               DbcDados(0).SetFocus
           Else
               DbcDados(0).Text = DatDados(0).Recordset.Fields("Nome")
               ListaEnt.AddItem psCompString(1, DatDados(0).Recordset.Fields("CodEntidade"), 6) & " - " & DatDados(0).Recordset.Fields("Nome")
               ListaEnt.Selected(ListaEnt.ListCount - 1) = True
           End If
        Else
            If DbcDados(0).Text <> "" Then
                If Not IsNumeric(DbcDados(0).BoundText) Then
                    MsgBox "Entidade não encontrada !!!", 16
                    DbcDados(0).SetFocus
                    Exit Sub
                Else
                    Criterio = "CodEntidade = " & DbcDados(0).BoundText
                    DatDados(0).Recordset.FindFirst Criterio
                    If DatDados(0).Recordset.NoMatch Then
                        MsgBox "Entidade não encontrada !!!", 16
                        DbcDados(0).Text = ""
                        DbcDados(0).SetFocus
                    Else
                        DbcDados(0).Text = DatDados(0).Recordset.Fields("Nome")
                        ListaEnt.AddItem psCompString(1, DatDados(0).Recordset.Fields("CodEntidade"), 6) & " - " & DatDados(0).Recordset.Fields("Nome")
                        ListaEnt.Selected(ListaEnt.ListCount - 1) = True
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub DbcDados_LostFocus(Index As Integer)
Me.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()

CentralizarForm Me, 1

DtData1 = Date
DtData2 = Date

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next


End Sub

Private Sub opTipoRel_Click(Index As Integer)
If Index = 1 Then
    FrameRecibos.Enabled = False
'    chReciboFormiga.Enabled = False
    DatDados(0).RecordSource = "Select * From Ent Where Tipo = 2 or Tipo = 5 Order by Nome"
    DatDados(0).Refresh
ElseIf Index = 2 Then
    FrameRecibos.Enabled = True
    FrameRecFormiga.Enabled = True
    DatDados(0).RecordSource = "Select * From Ent Where Tipo = 3 Order by Nome"
    DatDados(0).Refresh
ElseIf Index = 3 Then
    FrameRecibos.Enabled = True
    FrameRecFormiga.Enabled = False
'    chReciboFormiga.Enabled = False
    DatDados(0).RecordSource = "Select * From Ent Where Tipo = 2 or Tipo = 5 Order by Nome"
    DatDados(0).Refresh
    
ElseIf Index = 0 Then
    FrameRecibos.Enabled = False
    FrameRecFormiga.Enabled = False
    DatDados(0).RecordSource = "Select * From Ent Where Tipo = -1"
    DatDados(0).Refresh
End If

    ListaEnt.Clear

End Sub
