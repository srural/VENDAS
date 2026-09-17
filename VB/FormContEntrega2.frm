VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormContEntrega2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Entregas"
   ClientHeight    =   5115
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11145
   Begin VB.CommandButton BtOk 
      Caption         =   "&Visusaliza"
      Height          =   330
      Left            =   8100
      TabIndex        =   2
      Top             =   4635
      Width           =   1320
   End
   Begin VB.CommandButton CmbConfirma 
      Caption         =   "&Confirma"
      Height          =   330
      Left            =   9630
      TabIndex        =   1
      Top             =   4635
      Width           =   1320
   End
   Begin VB.Data DatDados 
      Caption         =   "Itp"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4185
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2835
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormContEntrega2.frx":0000
      Height          =   4350
      Left            =   180
      OleObjectBlob   =   "FormContEntrega2.frx":001A
      TabIndex        =   0
      ToolTipText     =   "Duplo Click p/ Alterar Entrega"
      Top             =   135
      Width           =   10770
   End
End
Attribute VB_Name = "FormContEntrega2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Loja As Recordset
Public Baixa As Boolean

Private Sub BtOK2_Click()

End Sub


Private Sub BtOK_Click()
Dim Criterio As String
Dim Grupo As Recordset

'DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)=#" & Format(DtData, "mm/dd/yyyy") & "#) AND ((ITP.Entregue)=False)) ORDER BY PED.CodPed"

Datdados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue, PED.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE ITP.Entregue = False  AND PED.Operacao >=1 AND PED.Operacao <=2  ORDER BY PED.CodPed"

Datdados(1).Refresh

'Set Grupo = Banco.OpenRecordset("SELECT * FROM GRU")
'Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")

If Datdados(1).Recordset.RecordCount <> 0 Then
    Datdados(1).Recordset.MoveLast
Else
    MsgBox "Não tem entregas pendentes", vbInformation, App.Title
End If

End Sub



Private Sub CmbConfirma_Click()
Dim Criterio As String
Dim Flag As Boolean
Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")

Flag = False
Datdados(1).Recordset.MoveFirst

While Not Datdados(1).Recordset.EOF
    If Datdados(1).Recordset.Fields("ITP.Entregue") = True Then
        If Datdados(1).Recordset.Fields("Loja") = "" Then
            Flag = True
        End If
    End If
    Datdados(1).Recordset.MoveNext
Wend

If Flag Then
     MsgBox "Tem Movimentação de Estoque sem Loja Especificada", vbInformation, App.Title
     Exit Sub
End If

Datdados(1).Recordset.MoveFirst

If Datdados(1).Recordset.RecordCount <> 0 Then
    While Not Datdados(1).Recordset.EOF
        If Datdados(1).Recordset.Fields("Itp.Entregue") = True Then
            Datdados(1).Recordset.Edit
'            If Baixa Then
                Datdados(1).Recordset.Fields("Estoque") = Datdados(1).Recordset.Fields("Estoque") - Datdados(1).Recordset.Fields("Qtd")
'            Else
'                DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") + DatDados(1).Recordset.Fields("Qtd")
'            End If
            Datdados(1).Recordset.Update
            Criterio = "Produto = " & Datdados(1).Recordset.Fields("CodPrd") & " and DescricaoLoja = " & """" & Datdados(1).Recordset.Fields("Loja") & """"
            Loja.FindFirst Criterio
            If Not Loja.NoMatch Then
                Loja.Edit
'                If Baixa Then
                    Loja.Fields("Estoque") = Loja.Fields("Estoque") - Datdados(1).Recordset.Fields("Qtd")
'                Else
'                    Loja.Fields("Estoque") = Loja.Fields("Estoque") + DatDados(1).Recordset.Fields("Qtd")
'                End If
                Loja.Update
            Else
                Loja.AddNew
                Loja.Fields("Produto") = Datdados(1).Recordset.Fields("CodPrd")
                Loja.Fields("DescricaoLoja") = Datdados(1).Recordset.Fields("Loja")
                Loja.Fields("Estoque") = Datdados(1).Recordset.Fields("Qtd")
                Loja.Update
                Datdados(1).Recordset.Edit
                Datdados(1).Recordset.Fields("DtCompra") = Date
                Datdados(1).Recordset.Update
            End If
            If Datdados(1).Recordset.Fields("Ped.Operacao") = 1 Or Datdados(1).Recordset.Fields("Ped.Operacao") = 2 Then
                Datdados(1).Recordset.Edit
                Datdados(1).Recordset.Fields("DtVenda") = Date
                Datdados(1).Recordset.Update
            End If
        End If
        Datdados(1).Recordset.MoveNext
    Wend
End If
Datdados(1).Refresh
End Sub

Private Sub Form_Activate()
Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

End Sub

Private Sub Grid_DblClick()
Datdados(1).Recordset.Edit
If Datdados(1).Recordset.Fields("Itp.Entregue") = 0 Then
    Datdados(1).Recordset.Fields("Itp.Entregue") = 1
Else
    Datdados(1).Recordset.Fields("Itp.Entregue") = 0
End If
Datdados(1).Recordset.Update
' DatDados(1).Refresh
Grid.Refresh

End Sub
