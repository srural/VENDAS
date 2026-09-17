VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormDevolucao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11235
   Begin VB.CommandButton BtItens 
      Caption         =   "Debater Itens"
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   6795
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   5
      Top             =   7335
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
      ScaleWidth      =   11265
      TabIndex        =   3
      Top             =   0
      Width           =   11265
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Devolução"
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
         TabIndex        =   4
         Top             =   45
         Width           =   2805
      End
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "Debater Valor"
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   6795
      Width           =   1545
   End
   Begin VB.Data DatPed 
      Caption         =   "Pedido"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas_1_8\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5490
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1935
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data DatDup 
      Caption         =   "Dup"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas_1_8\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4770
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormDevolucao.frx":0000
      Top             =   4455
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSDBGrid.DBGrid GridItens 
      Bindings        =   "FormDevolucao.frx":009E
      Height          =   2745
      Left            =   225
      OleObjectBlob   =   "FormDevolucao.frx":00B3
      TabIndex        =   0
      Top             =   3915
      Width           =   10725
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormDevolucao.frx":147C
      Height          =   2880
      Left            =   225
      OleObjectBlob   =   "FormDevolucao.frx":1491
      TabIndex        =   1
      Top             =   810
      Width           =   10725
   End
End
Attribute VB_Name = "FormDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VarTotal As Currency

Private Sub BtItens_Click()
Dim DatPed As Recordset
Dim DatDev As Recordset
Dim Criterio As String

If MsgBox("Confirma devolução ?", vbYesNo, App.Title) = vbYes Then
    
    Set DatDev = Banco.OpenRecordset("SELECT * FROM ITP WHERE Pedido = " & DBGrid.Columns(0))
    DatDev.Requery
    
    Set DatPed = Banco.OpenRecordset("SELECT * FROM ITP WHERE Pedido = " & FrmPedidos.TxtDados(0))
    
    
    DatPed.Requery
    If DatPed.RecordCount > 0 Then
        DatPed.MoveFirst
        While Not DatPed.Eof
            Criterio = " Produto =  " & DatPed!Produto
            DatDev.FindFirst Criterio
            If Not DatDev.NoMatch Then
                If DatDev!Qtd > 1 Then
                    DatDev.Edit
                    DatDev!Qtd = DatDev!Qtd - DatPed!Qtd
                    DatDev.Update
                Else
    '                DatDev.Edit
                    DatDev.Delete
                End If
  '
            End If
            DatPed.MoveNext
        Wend
    End If
    MsgBox "Devolução concluida ! ", vbInformation
    Unload Me
End If
End Sub

Private Sub BtOK_Click()
Dim VarPcl As Currency

    If DatDup.Recordset.RecordCount <> 0 Then
        VarPcl = (FrmPedidos.TxtDados(2) / DatDup.Recordset.RecordCount)
        DatDup.Recordset.MoveFirst
        If Left(DatDup.Recordset("Documento"), 2) <> "DV" Then
            If MsgBox("Confirma devolução ?", vbYesNo, App.Title) = vbYes Then
                DatDup.Recordset.MoveFirst
                While Not DatDup.Recordset.Eof
                    DatDup.Recordset.Edit
                    DatDup.Recordset("Valor") = DatDup.Recordset("Valor") - VarPcl
                    DatDup.Recordset("Documento") = "DV." & DBGrid.Columns(0)
                    DatDup.Recordset.Update
                    DatDup.Recordset.MoveNext
                Wend
            End If
        Else
            MsgBox "Pedido já devolvido !", vbInformation, App.Title
        End If
    Else
        If MsgBox("Não tem parcelas pra descontar ! Deseja lancar um crédito na conta do cliente ?", vbYesNo, App.Title) = vbYes Then
            SQL = "INSERT INTO DUP ( Entidade, DtEmiss, DtVenc , Valor,  Pedido, Tipo_Dup, NroPcl, Cond ) VALUES ( "
            SQL = SQL & FrmPedidos.LblCodigo & " , #"
            SQL = SQL & Format(Date, "mm/dd/yyyy") & "# , #"
            SQL = SQL & Format(Date, "mm/dd/yyyy") & "# , "
            SQL = SQL & Numero(FrmPedidos.TxtDados(2) * (-1), 2) & " ,"
            SQL = SQL & Val(DBGrid.Columns(0)) & " ,"
            SQL = SQL & """" & "DV" & """" & " ,"
            SQL = SQL & 1 & " ,"
            SQL = SQL & 0 & ")"
            FrmFrente.Banco.Execute SQL
        Else
                        
        End If
    End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DBGrid.Columns(0) <> "" Then
    DatDup.RecordSource = ("SELECT CodDup, DUP.Pedido, DUP.DtEmiss, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Cond, DUP.Documento FROM DUP WHERE DUP.Pedido =  " & DBGrid.Columns(0) & " AND DUP.Cond = False ORDER BY CodDup ")
    DatDup.Refresh
    If DatDup.Recordset.RecordCount <> 0 Then
        DatDup.Recordset.MoveFirst
        VarTotal = 0
        While Not DatDup.Recordset.Eof
            VarTotal = VarTotal + DatDup.Recordset("Valor")
            DatDup.Recordset.MoveNext
        Wend
    End If
      
End If

End Sub

Private Sub Form_Load()

DatPed.DatabaseName = Caminho & "\Dados.mdb"
DatDup.DatabaseName = Caminho & "\Dados.mdb"

DatPed.RecordSource = ("SELECT PED.CodPed, PED.DataEmiss, ENT.Nome, PED.Total, PED.Operacao, PED.Entidade FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.Entidade = " & FrmPedidos.LblCodigo & " ORDER BY CodPed")
DatPed.Refresh

End Sub
