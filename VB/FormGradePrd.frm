VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormGradePrd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   4065
   Icon            =   "FormGradePrd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4065
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   4440
      TabIndex        =   3
      Top             =   5220
      Width           =   4470
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Confirma"
      Height          =   375
      Left            =   2385
      TabIndex        =   2
      Top             =   4590
      Width           =   1275
   End
   Begin VB.Data DatDados 
      Caption         =   "Loj"
      Connect         =   "Access"
      DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT LOJ.Produto, LOJ.DescricaoLoja, LOJ.Estoque, LOJ.CodBarLoja FROM LOJ"
      Top             =   2205
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4245
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
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
         Left            =   225
         TabIndex        =   1
         Top             =   45
         Width           =   3930
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormGradePrd.frx":000C
      Height          =   3705
      Left            =   180
      OleObjectBlob   =   "FormGradePrd.frx":0026
      TabIndex        =   4
      Top             =   810
      Width           =   3480
   End
End
Attribute VB_Name = "FormGradePrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Itens As Recordset 'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas



Private Sub CmdOk_Click()


    Set Itens = Banco.OpenRecordset("ITP")

    If DatDados(0).Recordset.RecordCount > 0 Then
       DatDados(0).Recordset.MoveFirst
       While Not DatDados(0).Recordset.Eof
        
            Itens.AddNew
            Itens("Pedido") = FrmPedidos.TxtDados(0)
            Itens("Produto") = FrmPedidos.DbcDados(1).BoundText
            If IsNumeric(FrmPedidos.MebDados(2).Text) Then Itens("Desconto") = Round(FrmPedidos.MebDados(2).Text, CasasDec)
            If IsNumeric(FrmPedidos.MebDados(4).Text) Then Itens("Qtd") = FrmPedidos.MebDados(4).Text
            Itens("ValorUnit") = Round(FrmPedidos.MebVlr, 4)
            
            
            Itens("Valor") = Round(FrmPedidos.MebDados(5), 4)
'            Itens("CustoDia") = Round((CCur(FrmPedidos.MebDados(5).Text)), 4) ' - CCur(MebDados(5).Text) * Round((Extra("Desconto") / 100), 2)), 4)
                
            Itens("CustoDia") = Round(FrmPedidos.VarCusto, 4)
                        
            Itens("VendaDia") = Round(FrmPedidos.MebDados(5), 4)
            Itens("VlrVdaDia") = Round(CCur(FrmPedidos.MebDados(6).Text), 4) * CCur(FrmPedidos.MebDados(4).Text)
            Itens("Valor") = Round(CCur(FrmPedidos.MebDados(6).Text), 4) * CCur(FrmPedidos.MebDados(4).Text)
'            Itens("CustoDia") = Round(CCur(FrmPedidos.MebDados(6).Text), 4) * CCur(FrmPedidos.MebDados(4).Text)
                        
                If Not IsNull(FrmPedidos.MebComissao) Then
                    If IsNumeric(FrmPedidos.MebComissao) Then
                        Itens("Comissao") = Format((Round(CCur(FrmPedidos.MebDados(6).Text), 2) * CCur(FrmPedidos.MebDados(4).Text)) * (FrmPedidos.MebComissao / 100), "##,##0.00")
                    End If
                End If
                        

            Itens("Loja") = DatDados(0).Recordset("CodBarLoja")
            Itens("Complemento") = ""
            Itens("Referencia") = FrmPedidos.TxtRefItens
            Itens.Update
            
            DatDados(0).Recordset.MoveNext
       Wend
       
        FrmPedidos.MebDados(4).Text = ""
        FrmPedidos.MebDados(5).Text = ""
        FrmPedidos.MebDados(6).Text = ""
'        DbcDados(1).Text = ""
        FrmPedidos.MebVlr = ""
        FrmPedidos.MebDados(8).Text = ""
        FrmPedidos.TxtRefItens = ""
        FrmPedidos.CmbAcab = ""
        
        
'        Set Extra = Banco.OpenRecordset("SELECT Itp.Pedido, Sum(Itp.Vlrvdadia) AS SomaDeValor, Sum(Itp.Valor) AS SomaDeValorPrd  FROM PED INNER JOIN Itp ON Ped.CodPed = Itp.Pedido Where Pedido = " & VarCodPed & " GROUP BY Itp.Pedido")
'        TxtDados(2).Text = Format(Extra("SomaDeValor"), "#,##0.00")
'        TxtDados(1).Text = Format(Extra("SomaDeValor"), "#,##0.00")
'        FrmPedidos Totaliza
'        VarTotalPrd = Extra("SomaDeValorPrd")
'        FrmPedidos.MebDados(4).SetFocus
       
    End If
'       Exit Sub
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
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatDados(0).RecordSource = "SELECT  Produto, DescricaoLoja, Estoque, CodBarLoja FROM LOJ WHERE (((Produto)=" & FormProdutos.TxtDados(0) & " )) ORDER BY DescricaoLoja"
DatDados(0).Refresh


End Sub



