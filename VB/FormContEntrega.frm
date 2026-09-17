VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormContEntrega 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11985
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   10
      Top             =   7380
      Width           =   14370
   End
   Begin VB.CommandButton BtMarcar 
      Caption         =   "Marcar/Desmarcar"
      Height          =   375
      Left            =   9765
      TabIndex        =   9
      Top             =   6750
      Width           =   2040
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11985
      TabIndex        =   7
      Top             =   0
      Width           =   11985
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Entregas"
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
         TabIndex        =   8
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.CheckBox ChkData 
      Caption         =   "Filtrar pedidos sem data de entrega"
      Height          =   315
      Left            =   1785
      TabIndex        =   6
      Top             =   810
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   5820
      TabIndex        =   3
      Top             =   660
      Width           =   2295
      Begin VB.OptionButton Opt 
         Caption         =   "Entrada"
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   5
         Top             =   150
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Saída"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CommandButton CmbConfirma 
      Caption         =   "&Confirma"
      Height          =   330
      Left            =   10275
      TabIndex        =   2
      Top             =   810
      Width           =   1320
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Visusaliza"
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   765
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
      Top             =   3420
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "FormContEntrega.frx":0000
      Height          =   5235
      Left            =   180
      OleObjectBlob   =   "FormContEntrega.frx":001A
      TabIndex        =   1
      ToolTipText     =   "Duplo Click p/ Alterar Entrega"
      Top             =   1305
      Width           =   11625
   End
End
Attribute VB_Name = "FormContEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Loja As Recordset
Public Baixa As Boolean

Private Sub BtOK2_Click()

End Sub


Private Sub BtMarcar_Click()
With DatDados(1).Recordset
    
    If .RecordCount > 0 Then
        .MoveFirst
        While Not .Eof
             Grid_DblClick
            .MoveNext
        Wend
    End If

End With
End Sub

Private Sub BtOK_Click()
Dim Criterio As String
Dim Grupo As Recordset
If Opt(0).value = True Then
    If ChkData.value = 1 Then
        DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue, PED.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE Ped.DtSaida is null AND PED.Operacao >=1 AND PED.Operacao <=3  ORDER BY PED.CodPed"
    Else
        DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue, PED.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE ITP.Entregue = False  AND PED.Operacao >=1 AND PED.Operacao <=3  ORDER BY PED.CodPed"
    End If
Else
    If ChkData.value = 1 Then
        DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue, PED.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE Ped.DtSaida is null AND PED.Operacao =0   ORDER BY PED.CodPed"
    Else
        DatDados(1).RecordSource = "SELECT PED.*, ITP.*, PRD.*, ENT.Nome, PED.DataEmiss, ITP.Entregue, PED.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE ITP.Entregue = False  AND PED.Operacao =0  ORDER BY PED.CodPed"
    End If
End If

DatDados(1).Refresh

'Set Grupo = Banco.OpenRecordset("SELECT * FROM GRU")
'Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")

If DatDados(1).Recordset.RecordCount <> 0 Then
    DatDados(1).Recordset.MoveLast
Else
    MsgBox "Não tem entregas pendentes", vbInformation, App.Title
End If

End Sub

Private Sub CmbCancela_Click()
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbConfirma_Click()
Dim Criterio As String
Dim Flag As Boolean
If MsgBox("Confirma entrega dos produtos selecionados ?", vbYesNo, App.Title) = vbYes Then
    Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")
    
    Flag = False
    DatDados(1).Recordset.MoveFirst
    
    While Not DatDados(1).Recordset.Eof
        If DatDados(1).Recordset.Fields("ITP.Entregue") = True Then
            If DatDados(1).Recordset.Fields("Loja") = "" Then
                Flag = True
            End If
        End If
        DatDados(1).Recordset.MoveNext
    Wend
    
    If Flag Then
         MsgBox "Tem Movimentação de Estoque sem Loja Especificada", vbInformation, App.Title
         Exit Sub
    End If
    
    DatDados(1).Recordset.MoveFirst
    
    If DatDados(1).Recordset.RecordCount <> 0 Then
        While Not DatDados(1).Recordset.Eof
            If DatDados(1).Recordset.Fields("Itp.Entregue") = True Then
                DatDados(1).Recordset.Edit
                If Opt(0).value = True Then
                    DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") - DatDados(1).Recordset.Fields("Qtd")
                Else
                    DatDados(1).Recordset.Fields("Estoque") = DatDados(1).Recordset.Fields("Estoque") + DatDados(1).Recordset.Fields("Qtd")
                End If
                
                DatDados(1).Recordset.Update
                Criterio = "Produto = " & DatDados(1).Recordset.Fields("CodPrd") & " and DescricaoLoja = " & """" & DatDados(1).Recordset.Fields("Loja") & """"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    If Opt(0).value = True Then
                        Loja.Fields("Estoque") = Loja.Fields("Estoque") - DatDados(1).Recordset.Fields("Qtd")
                    Else
                        Loja.Fields("Estoque") = Loja.Fields("Estoque") + DatDados(1).Recordset.Fields("Qtd")
                    End If
                    Loja.Update
                Else
                    Loja.AddNew
                    Loja.Fields("Produto") = DatDados(1).Recordset.Fields("CodPrd")
                    Loja.Fields("DescricaoLoja") = DatDados(1).Recordset.Fields("Loja")
                    Loja.Fields("Estoque") = DatDados(1).Recordset.Fields("Qtd")
                    Loja.Update
                    DatDados(1).Recordset.Edit
                    DatDados(1).Recordset.Fields("DtCompra") = Date
                    DatDados(1).Recordset.Update
                End If
                If DatDados(1).Recordset.Fields("Ped.Operacao") = 1 Or DatDados(1).Recordset.Fields("Ped.Operacao") = 2 Then
                    DatDados(1).Recordset.Edit
                    DatDados(1).Recordset.Fields("DtVenda") = Date
                    DatDados(1).Recordset.Update
                End If
                If DatDados(1).Recordset.Fields("BaixarComp") = True Then
                    BaixarComposicao True, DatDados(1).Recordset.Fields("Ped.Operacao"), DatDados(1).Recordset.Fields("CodPrd"), DatDados(1).Recordset.Fields("Qtd"), DatDados(1).Recordset.Fields("Loja")
                End If
                
            End If
            DatDados(1).Recordset.MoveNext
        Wend
    End If
    DatDados(1).Refresh
End If
End Sub



Private Sub Form_Load()
Dim Componentes As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next
Me.Left = 200
Me.Top = 200
End Sub

Private Sub Grid_DblClick()
DatDados(1).Recordset.Edit
If DatDados(1).Recordset.Fields("Itp.Entregue") = 0 Then
    DatDados(1).Recordset.Fields("Itp.Entregue") = 1
Else
    DatDados(1).Recordset.Fields("Itp.Entregue") = 0
End If
DatDados(1).Recordset.Update
' DatDados(1).Refresh
Grid.Refresh

End Sub

Private Sub Opt_Click(Index As Integer)
BtOK_Click
End Sub
