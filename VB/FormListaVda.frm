VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormListaVda 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   10395
   Icon            =   "FormListaVda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10395
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   7
      Top             =   7830
      Width           =   11535
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
      Left            =   5220
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodPrd, Descrição_Produto From Prd Order By Descrição_Produto"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Confirma Lista"
      Height          =   375
      Left            =   8820
      TabIndex        =   6
      Top             =   945
      Width           =   1275
   End
   Begin VB.Data DatDados 
      Caption         =   "Hist"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome FROM  HISTEST GROUP BY Nome ORDER BY Nome"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton CmdLer 
      Caption         =   "&Carrega Lista"
      Height          =   375
      Left            =   6435
      TabIndex        =   1
      Top             =   945
      Width           =   1230
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
      TabIndex        =   3
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Produtos "
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
         TabIndex        =   4
         Top             =   45
         Width           =   6630
      End
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
      Left            =   3465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4725
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormListaVda.frx":000C
      Height          =   6075
      Left            =   180
      OleObjectBlob   =   "FormListaVda.frx":0026
      TabIndex        =   2
      Top             =   1395
      Width           =   9945
   End
   Begin MSDBCtls.DBCombo TxtNome 
      Bindings        =   "FormListaVda.frx":10BD
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   945
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Lista"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1065
   End
End
Attribute VB_Name = "FormListaVda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas






Private Sub CmdLer_Click()
Dim SQL As String
If TxtNome.Text <> "" Then
    DatDados(0).RecordSource = "SELECT  Qtd, HISTEST.Nome, HISTEST.CodPrd, PRD.Descrição_Produto, HISTEST.Valor, HISTEST.Ordem FROM HISTEST INNER JOIN PRD ON HISTEST.CodPrd = PRD.CodPrd where nome = " & """" & TxtNome & """" & " Order by HISTEST.Ordem"
    DatDados(0).Refresh
    SQL = "UPDATE HISTEST SET HISTEST.Qtd = 0 WHERE HISTEST.Nome = " & """" & TxtNome & """"
    DatDados(0).Refresh
    Banco.Execute SQL
Else
    MsgBox "Digite o nome da lista", vbInformation
    TxtNome.SetFocus
End If

End Sub

Private Sub CmdOk_Click()
Dim Criterio As String
Dim Itens As Recordset
If DatDados(0).Recordset.RecordCount > 0 Then
    DatDados(0).Recordset.MoveFirst
    Set Itens = Banco.OpenRecordset("ITP")
    While Not DatDados(0).Recordset.Eof
        If DatDados(0).Recordset("Qtd") <> 0 Then
            Itens.AddNew
            Itens("Pedido") = FrmPedidos.TxtDados(0)
            Itens("Produto") = DatDados(0).Recordset("CodPrd")
            Itens("Desconto") = 0
            Itens("Qtd") = DatDados(0).Recordset("Qtd")
            
            Itens("ValorUnit") = DatDados(0).Recordset("Valor")
            Itens("Valor") = DatDados(0).Recordset("Valor") * DatDados(0).Recordset("Qtd")
            Itens("VendaDia") = DatDados(0).Recordset("Valor")
            Itens("VlrVdaDia") = DatDados(0).Recordset("Valor") * DatDados(0).Recordset("Qtd")
            Criterio = "CodPrd = " & DatDados(0).Recordset("CodPrd")
            DatDados(1).Recordset.FindFirst Criterio
            If Not DatDados(1).Recordset.NoMatch Then
                Itens("CustoDia") = DatDados(1).Recordset("Custo")
            Else
                Itens("CustoDia") = DatDados(0).Recordset("Valor")
            End If
            Itens("Comissao") = 0
            Itens("Loja") = "LOJA"
            Itens("Complemento") = ""
            Itens("Referencia") = ""
            Itens.Update
        End If
        DatDados(0).Recordset.MoveNext
        
    Wend
    FrmPedidos.DatDados(2).RecordSource = "SELECT Tributo,Icm,Descrição_Produto,CondPgto, Itp.Referencia,Itp.complemento, Ped.Entregue,DataEmiss,Hora,Operacao,Transportadora,PedSub,DtSaida,PrevEntrega,NroNt,Ped.Obs,ValorFrete,CodPed,Produto,Qtd,VlrVdaDia,CodItp,Itp.Desconto,Loja,Itp.Entregue,Pedido,Embalagem,ValorUnit,Valor,CodPrd,  [PRD]![Descrição_Produto] & SPACE(3) & [ITP]![Complemento] AS DescProd FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido Where CodPed =  " & VarCodPed & " Order by Itp.CodItp"
    FrmPedidos.DatDados(2).Refresh
    If FrmPedidos.DatDados(2).Recordset.RecordCount > 0 Then
        FrmPedidos.DatDados(2).Recordset.MoveLast
        FrmPedidos.BtCancela.Enabled = True
        FrmPedidos.BtOk.Enabled = True
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
End Sub


Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatDados(0).RecordSource = "SELECT Qtd, HISTEST.Nome, HISTEST.CodPrd, PRD.Descrição_Produto, HISTEST.Valor FROM HISTEST INNER JOIN PRD ON HISTEST.CodPrd = PRD.CodPrd where nome = " & """" & TxtNome & """"
DatDados(0).Refresh

DatDados(1).RecordSource = "Select  CodPrd, Descrição_Produto, Custo From Prd Order By CodPrd "
DatDados(1).Refresh

End Sub


Private Sub TxtNome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TxtNome <> "" Then
        CmdLer_Click
    End If
End If
End Sub

