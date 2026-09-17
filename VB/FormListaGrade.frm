VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormListaGrade 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   9630
   ClientTop       =   4440
   ClientWidth     =   4500
   Icon            =   "FormListaGrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4500
   Begin VB.TextBox VarComiss 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   4860
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Data DatDados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1125
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DatDados 
      Caption         =   "Lista"
      Connect         =   "Access"
      DatabaseName    =   "app.path & ""\Produtos.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1395
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodPrd, Descrição_Produto From Prd Order By Descrição_Produto"
      Top             =   3015
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Confirma"
      Default         =   -1  'True
      Height          =   375
      Left            =   1575
      TabIndex        =   0
      Top             =   4770
      Width           =   1275
   End
   Begin VB.Data DatDados 
      Caption         =   "Hist"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, CodPrd, Qtd FROM  HISTEST ORDER BY Nome"
      Top             =   3780
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormListaGrade.frx":000C
      Height          =   4380
      Left            =   180
      OleObjectBlob   =   "FormListaGrade.frx":0026
      TabIndex        =   1
      Top             =   180
      Width           =   4065
   End
End
Attribute VB_Name = "FormListaGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset 'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas


Private Sub CmdOk_Click()
Dim Criterio As String
Dim Itens As Recordset
If DatDados(2).Recordset.RecordCount > 0 Then
    DatDados(2).Recordset.MoveFirst
    Set Itens = Banco.OpenRecordset("ITP")
    While Not DatDados(2).Recordset.Eof
        If DatDados(2).Recordset("Qtd") <> 0 Then
            Itens.AddNew
            Itens("Pedido") = FrmPedidos.TxtDados(0)
            Itens("Produto") = DatDados(2).Recordset("CodPrd")
            Itens("Desconto") = 0
            Itens("Qtd") = DatDados(2).Recordset("Qtd")
            Criterio = "CodPrd = " & DatDados(2).Recordset("CodPrd")
            DatDados(0).Recordset.FindFirst Criterio
            Itens("ValorUnit") = FrmPedidos.MebVlr
            Itens("Valor") = FrmPedidos.MebVlr * DatDados(2).Recordset("Qtd")
            Itens("VendaDia") = FrmPedidos.MebVlr
            Itens("VlrVdaDia") = FrmPedidos.MebVlr * DatDados(2).Recordset("Qtd")
            Itens("CustoDia") = DatDados(0).Recordset("Custo")
            Itens("Comissao") = (FrmPedidos.MebVlr * DatDados(2).Recordset("Qtd")) * (VarComiss / 100)
            Itens("Loja") = DatDados(2).Recordset("Nome")
            Itens("Complemento") = DatDados(2).Recordset("Nome")
            Itens("Referencia") = ""
            Itens.Update
        End If
        DatDados(2).Recordset.MoveNext
        
    Wend
    FrmPedidos.DatDados(2).RecordSource = "SELECT Tributo,Icm,Descrição_Produto,CondPgto, Itp.Referencia,Itp.complemento, Ped.Entregue,DataEmiss,Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
    FrmPedidos.DatDados(2).Refresh
    If FrmPedidos.DatDados(2).Recordset.RecordCount > 0 Then
        FrmPedidos.DatDados(2).Recordset.MoveLast
        FrmPedidos.BtCancela.Enabled = True
        FrmPedidos.BtOK.Enabled = True
        Unload Me
    End If
End If
End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'AtivarMenu (True)
'Formulário.Left = 30
'Formulário.Top = 0

'GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)
' testa se o bamco está em modo de consulta ou não
'DbcPesquisa.SetFocus

    DatDados(2).RecordSource = "SELECT * From HistEst WHERE HISTEST.Ordem = " & FrmPedidos.TxtDados(0) & " Order by Nome"
    DatDados(2).Refresh
    If DatDados(2).Recordset.RecordCount > 0 Then
        DatDados(2).Recordset.MoveLast
        DatDados(2).Recordset.MoveFirst
        
     End If
     DBGrid1.SetFocus
End Sub


Private Sub Form_Load()
Dim Componentes As Variant
Dim Sql As String

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatDados(2).RecordSource = "SELECT * From HistEst "
DatDados(2).Refresh

If FrmPedidos.DbcDados(1).BoundText <> "" Then
    DatDados(1).RecordSource = "Select  * from loj  where  Produto = " & FrmPedidos.DbcDados(1).BoundText & " Order by DescricaoLoja"
    DatDados(1).Refresh
End If

DatDados(0).RecordSource = "Select  * From Prd Order By CodPrd "
DatDados(0).Refresh


Sql = "DELETE HISTEST.*, HISTEST.Ordem FROM HISTEST WHERE HISTEST.Ordem = " & FrmPedidos.TxtDados(0)
Banco.Execute Sql

If FrmPedidos.DbcDados(1).BoundText <> "" Then

    If DatDados(1).Recordset.RecordCount > 0 Then
        DatDados(1).Recordset.MoveFirst
        While Not DatDados(1).Recordset.Eof
            DatDados(2).Recordset.AddNew
            DatDados(2).Recordset.Fields("CodPrd") = DatDados(1).Recordset.Fields("Produto")
            DatDados(2).Recordset.Fields("Nome") = DatDados(1).Recordset.Fields("DescricaoLoja")
            DatDados(2).Recordset.Fields("Qtd") = 0
            DatDados(2).Recordset.Fields("Ordem") = FrmPedidos.TxtDados(0)
            DatDados(2).Recordset.Update
            DatDados(1).Recordset.MoveNext
        Wend
    '    SQL = "SELECT * From HistEst WHERE HISTEST.Ordem = " & FrmPedidos.TxtDados(0) & " Order by HistEst.DescricaoLoja"
    '    DatDados(2).Refresh
    '    DBGrid1.Refresh
    
    End If
End If
End Sub



