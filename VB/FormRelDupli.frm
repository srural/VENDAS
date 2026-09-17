VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormRelDupli 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8220
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8265
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -585
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   32
      Top             =   7875
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios de Contas a Receber"
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
         Width           =   7575
      End
   End
   Begin VB.Data DatVend 
      Caption         =   "Vendedores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7"
      Top             =   7290
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordem"
      Height          =   1200
      Left            =   270
      TabIndex        =   23
      Top             =   3420
      Width           =   1635
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Emissão"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   510
         Width           =   1140
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pedido"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   180
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Vencimento"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   795
         Width           =   1275
      End
   End
   Begin VB.CheckBox ChkInativos 
      Caption         =   "&Inativos"
      Height          =   255
      Left            =   300
      TabIndex        =   22
      Top             =   6945
      Width           =   1275
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      ItemData        =   "FormRelDupli.frx":0000
      Left            =   2085
      List            =   "FormRelDupli.frx":0025
      TabIndex        =   21
      Top             =   6930
      Width           =   2190
   End
   Begin MSComctlLib.ProgressBar BarBarra 
      Height          =   210
      Left            =   2070
      TabIndex        =   20
      Top             =   7380
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox ChkCorrecao 
      Caption         =   "&Com Correção"
      Height          =   315
      Left            =   315
      TabIndex        =   2
      Top             =   2790
      Width           =   1515
   End
   Begin VB.CheckBox ChkData 
      Caption         =   "&Data Emissão"
      Height          =   315
      Left            =   315
      TabIndex        =   3
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   1200
      Left            =   285
      TabIndex        =   19
      Top             =   4680
      Width           =   1635
      Begin VB.OptionButton OptTipo 
         Caption         =   "Ordem Data"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   795
         Width           =   1275
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Completo"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumido"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   510
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   345
      Left            =   4425
      TabIndex        =   18
      Top             =   6915
      Width           =   1590
   End
   Begin VB.Frame Frame2 
      Caption         =   "Entidades (Clientes)"
      Height          =   885
      Left            =   255
      TabIndex        =   15
      Top             =   690
      Width           =   1635
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   165
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Selecionar"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   450
         Width           =   1050
      End
   End
   Begin VB.Data DatDados 
      Caption         =   "Ent"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3870
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Ent Where Tipo = 1 or Tipo = 5 Or Tipo = 6 Or Tipo = 7 Order by Nome"
      Top             =   2385
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.ListBox ListaEnt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   2070
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   1125
      Width           =   5790
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      Height          =   915
      Left            =   285
      TabIndex        =   12
      Top             =   5940
      Width           =   1635
      Begin VB.OptionButton Opt 
         Caption         =   "Recebidas"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   495
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "À Receber"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   195
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Visualiza"
      Height          =   345
      Left            =   6270
      TabIndex        =   9
      Top             =   6915
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   285
      TabIndex        =   0
      Top             =   1815
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Format          =   100007937
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   285
      TabIndex        =   1
      Top             =   2385
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Format          =   100007937
      CurrentDate     =   36641
   End
   Begin MSDBCtls.DBCombo DbcDados 
      Bindings        =   "FormRelDupli.frx":005D
      Height          =   315
      Index           =   0
      Left            =   2085
      TabIndex        =   13
      Top             =   720
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo DbcVend 
      Bindings        =   "FormRelDupli.frx":0077
      Height          =   315
      Left            =   2070
      TabIndex        =   27
      Top             =   6345
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   2070
      TabIndex        =   29
      Top             =   6705
      Width           =   510
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      Height          =   285
      Left            =   2070
      TabIndex        =   28
      Top             =   6120
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   285
      TabIndex        =   11
      Top             =   2175
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   285
      TabIndex        =   10
      Top             =   1605
      Width           =   795
   End
End
Attribute VB_Name = "FormRelDupli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter

End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula As String
Dim Condicao As String
Dim CamposSql As String
Dim Ordenado As String
Dim SQL As String
Dim CampoData As String
Dim Ativo As String
Dim VarOrdem As String

Periodo = Format(DtData1, "dd/mm/yyyy") & " a " & Format(DtData2, "dd/mm/yyyy")
Data1 = Month(DtData1) & "/" & Day(DtData1) & "/" & Year(DtData1)
Data2 = Month(DtData2) & "/" & Day(DtData2) & "/" & Year(DtData2)

               
If ChkInativos.value = 1 Then
    Ativo = "False"
Else
    Ativo = "True"
End If
               
If OPt(0).value = True Then
    Condicao = "False"
    CampoData = "DtVenc"
'    If OptTipo(1).Value Then
'            CampoData = "PrimeiroDeDtVenc"
'    End If
ElseIf OPt(1).value Then
   'MDIPrincipal.RptRel.ReportFileName = Caminho & "\relrecpcl.RPT"
   Condicao = "True"
   CampoData = "DtPgto"
End If

If ChkData.value = 1 Then
   CampoData = "DtEmiss"
End If

If OptOrdem(1).value = True Then
    VarOrdem = " ORDER BY ENT.Nome, DUP.Pedido, DUP.NroPcl "
ElseIf OptOrdem(2).value = True Then
        VarOrdem = " ORDER BY ENT.Nome, Dup.DtEmiss"
    Else
        VarOrdem = " ORDER BY ENT.Nome, DUP.DtVenc "
End If
                                             
If OptTipo(0).value = True Or OptTipo(1).value = True Or OptTipo(2).value = True Then

'    MDIPrincipal.RptRel.ReportFileName = Caminho & "\relpcl.RPT"
        
    If OPt(2).value = True Then
        flag = True
        If ListaEnt.Enabled = True Then
            For y = 0 To ListaEnt.ListCount - 1
                If ListaEnt.Selected(y) = True Then
                    If flag Then
                        Formula = " Having (Ent.CodEntidade = " & CLng(Left(ListaEnt.List(y), 6))
                        flag = False
                    Else
                        Formula = Formula & " or Ent.CodEntidade = " & CLng(Left(ListaEnt.List(y), 6))
                    End If
                End If
            Next
            Formula = Formula + ") and (" & CampoData & " >= #" & Data1 & "# AND " & CampoData & " <= #" & Data2 & "#  and DUP.Cond = " & Condicao & ")"
        End If
    Else
        Formula = " Having " & CampoData & " >= #" & Data1 & "# AND " & CampoData & " <= #" & Data2 & "# and DUP.Cond = " & Condicao & " AND Ent.Ativo = " & Ativo
    End If
Else
    'MDIPrincipal.RptRel.ReportFileName = Caminho & "\relpclres.RPT"
    Formula = " Having " & CampoData & " >= #" & Data1 & "# AND " & CampoData & " <= #" & Data2 & "# and Cond = " & Condicao & " AND Ent.Ativo = " & Ativo
End If

If DbcVend.Text <> "" Then

    Formula = Formula & " and Vendedores.CodEntidade =  " & DbcVend.BoundText

End If

'CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Fone, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, Vend.CodEntidade, Vend.Nome, DUP.DtPgto, DUP.Cond FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN (ENT AS Vend INNER JOIN PED ON Vend.CodEntidade = PED.Vendedor) ON DUP.Pedido = PED.CodPed GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Fone, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, Vend.CodEntidade, Vend.Nome, DUP.DtPgto, DUP.Cond "
'If OptTipo(0).Value = True Or OptTipo(2).Value = True Then

'    CamposSql = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, Ent.Cidade, ENT.Fone, Ent.Prazo, Ent.Juros, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, DUP.DtPgto, DUP.Cond, Dup.Juros FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, Ent.Cidade, ENT.Fone,  Ent.Prazo, Ent.Juros, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, DUP.DtPgto, DUP.Cond , Dup.Juros"
    CamposSql = "SELECT ENT.Ativo, ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ENT.Prazo, ENT.Juros, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, DUP.DtPgto, DUP.Cond, DUP.Juros, Vendedores.ENT.Nome, Vendedores.CodEntidade  FROM ((ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) LEFT JOIN PED ON DUP.Pedido = PED.CodPed) LEFT JOIN Vendedores ON PED.Vendedor = Vendedores.CodEntidade GROUP BY ENT.Ativo, ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ENT.Prazo, ENT.Juros, DUP.Documento, DUP.Pedido, DUP.NroPcl, DUP.DtEmiss, DUP.DtVenc, DUP.Valor, DUP.VlrPago, DUP.Obs, DUP.Tipo_Dup, DUP.DtPgto, DUP.Cond, DUP.Juros, Vendedores.ENT.Nome, Vendedores.CodEntidade "

'Else
'    CamposSql = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Fone, First(DUP.DtEmiss) AS PrimeiroDeDtEmiss, First(DUP.DtVenc) AS PrimeiroDeDtVenc, Sum(DUP.Valor) AS SomaDeValor, DUP.Cond FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Fone, DUP.Cond"
'End If


'Set varPesquisa(1) = Banco.OpenRecordset(Sql)

If ChkCorrecao.value = 1 Then
    CorrigePcl CamposSql & Formula
End If

If CmbTipo.Text <> "(Todos)" Then
    If CmbTipo.Text = "PC" Then
        Formula = Formula & " AND (Dup.Tipo_Dup =  " & """" & "SD" & """" & " OR Dup.Tipo_Dup = " & """" & CmbTipo & """" & ")"
    Else
        Formula = Formula & " and Dup.Tipo_Dup = " & """" & CmbTipo & """"
    End If
End If

If OptTipo(0).value = True Then
    Ordenado = VarOrdem
    SQL = CamposSql & Formula & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPcl.Relatorio.Recordset = varPesquisa(1)
    FrmRelPcl.rpCabTitulo.Caption = "Relatório de Contas a Receber / Recebidas"
    FrmRelPcl.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPcl.rpCabPeriodo.Caption = Periodo
    FrmRelPcl.Config
End If

If OptTipo(1).value = True Then
    Ordenado = VarOrdem
    SQL = CamposSql & Formula & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPclRes.Relatorio.Recordset = varPesquisa(1)
    If Condicao = True Then
        FrmRelPclRes.rpCabTitulo.Tag = "1"
    End If
    FrmRelPclRes.rpCabTitulo.Caption = "Relatório de Contas a Receber"
    FrmRelPclRes.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPclRes.rpCabPeriodo.Caption = Periodo
    FrmRelPclRes.Config
End If

If OptTipo(2).value = True Then
    Ordenado = " ORDER BY DUP.DtVenc "
    SQL = CamposSql & Formula & Ordenado
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelPclData.Relatorio.Recordset = varPesquisa(1)
    FrmRelPclData.rpCabTitulo.Caption = "Relatório de Contas a Receber / Recebidas"
    FrmRelPclData.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelPclData.rpCabPeriodo.Caption = Periodo
    FrmRelPclData.Config
End If

End Sub

Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

CarregaTipo
CmbTipo.ListIndex = 0

End Sub

Private Sub Form_Load()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

DtData1 = Date
DtData2 = Date

Dim Componente As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
DatVend.DatabaseName = Caminho & "\Dados.Mdb"

CmbTipo.ListIndex = 0

End Sub

Private Sub DbcDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 9 Or KeyCode = 13 Then
    If Index = 0 Then
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
                If DbcDados(0).BoundText = 0 Then
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

Private Sub Opt_Click(Index As Integer)
If Index = 3 Then
    DbcDados(0).Enabled = False
    ListaEnt.Enabled = False
ElseIf Index = 2 Then
    DbcDados(0).Enabled = True
    ListaEnt.Enabled = True
End If
End Sub

Private Sub CorrigePcl(SQL As String)
Dim SomaPcl As Recordset
Dim RecDup As Recordset
Dim VarJuros As Currency
Dim Criterio As String
Dim Cont As Double
Set RecDup = Banco.OpenRecordset("Select * From Dup Order By Pedido, NroPcl")

BarBarra.Min = 0
BarBarra.Max = 100

Set SomaPcl = Banco.OpenRecordset(SQL & " ORDER BY   DUP.Pedido ")
If SomaPcl.RecordCount > 0 Then
    SomaPcl.MoveLast
    SomaPcl.MoveFirst
    Do Until SomaPcl.Eof
        If GetInputState() Then DoEvents

        Cont = Cont + 1
        If SomaPcl("Cond") = False Then
            If SomaPcl("DtVenc") < Date Then
                Prazo = Date - SomaPcl("DtVenc")
                If IsNumeric(SomaPcl("Prazo")) Then
                    VarJuros = SomaPcl("Valor") * (((SomaPcl("Ent.Juros") / 30) / 100) * SomaPcl("Prazo"))
                    Criterio = "Entidade = " & SomaPcl("Ent.CodEntidade") & " AND Pedido = " & SomaPcl("Pedido") & " AND NroPcl = " & SomaPcl("NroPcl") & " AND Documento = " & """" & SomaPcl("Documento") & """"
                    RecDup.FindFirst Criterio
                    If Not RecDup.NoMatch Then
                        RecDup.Edit
                        RecDup("Juros") = VarJuros
                        RecDup.Update
                    End If
                End If
            End If
        End If
        SomaPcl.MoveNext
        DoEvents
        BarBarra.value = (Cont / SomaPcl.RecordCount) * 100
    Loop
End If

End Sub


