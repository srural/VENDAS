VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormControleEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Entrega"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12945
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   12
      Top             =   7965
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
      ScaleWidth      =   12975
      TabIndex        =   10
      Top             =   0
      Width           =   12975
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
         TabIndex        =   11
         Top             =   45
         Width           =   4695
      End
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormControleEntrega.frx":0000
      Height          =   5775
      Left            =   270
      OleObjectBlob   =   "FormControleEntrega.frx":0014
      TabIndex        =   9
      Top             =   1305
      Width           =   12390
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "S:\Desenvolvimento\Vendas_2_0\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   765
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox TxtConfirma 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6570
      TabIndex        =   7
      Top             =   855
      Width           =   2490
   End
   Begin VB.TextBox TxtTotal 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10845
      TabIndex        =   6
      Top             =   7290
      Width           =   1815
   End
   Begin VB.TextBox txtBarra 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3105
      TabIndex        =   3
      Top             =   855
      Width           =   3120
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "Ok"
      Height          =   330
      Left            =   1980
      TabIndex        =   2
      Top             =   900
      Width           =   915
   End
   Begin VB.TextBox TxtCarga 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   270
      TabIndex        =   0
      Top             =   855
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmacao"
      Height          =   195
      Left            =   6570
      TabIndex        =   8
      Top             =   630
      Width           =   885
   End
   Begin VB.OLE OLE1 
      Height          =   60
      Left            =   8325
      TabIndex        =   5
      Top             =   2655
      Width           =   30
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Barra"
      Height          =   195
      Left            =   3105
      TabIndex        =   4
      Top             =   630
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Carga"
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "FormControleEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LstPesq As ListItem
Public VarCarga As Recordset
Public VarPrd As Recordset
Public VarTotal2 As Currency


Private Sub BtOK_Click()
Dim x As Integer
Dim SQL As String
Dim Criterio As String

Dim VarTotal As Currency

VarTotal = 0

'Lista1.ListItems.Clear

'Sql = "SELECT Carga.*, PRD.Descrição_Produto, Carga.Codigo, Carga.NroCarga FROM Carga INNER JOIN PRD ON Carga.Produto = PRD.CodPrd WHERE Carga.NroCarga  = " & TxtCarga & " ORDER BY Carga.Codigo"

SQL = "SELECT Carga.Produto, PRD.Descrição_Produto, Prd.Validade, Carga.Qtd, Carga.Aferido, Carga.Saldo FROM Carga INNER JOIN PRD ON Carga.Produto = PRD.CodPrd WHERE  NroCarga = " & TxtCarga


'Set VarCarga = Banco.OpenRecordset(Sql)

Data1.RecordSource = SQL

Data1.Refresh

'VarCarga.Requery

'If VarCarga.RecordCount > 0 Then
'    VarCarga.MoveFirst
'    While Not VarCarga.Eof
'        Set LstPesq = Lista1.ListItems.Add(, , VarCarga("Produto"))
'
'        LstPesq.SubItems(1) = psCompString(3, Left(VarCarga!Descrição_Produto, 40), 40)
'        LstPesq.SubItems(2) = Format(Alinhar(VarCarga!Qtd, 10, 2), "###,##0")
'        VarTotal = VarTotal + VarCarga!Qtd
'        VarCarga.MoveNext
'    Wend
'End If

'TxtTotal = Format(VarTotal, "###,##0.00")

'Lista2.ListItems.Clear




Set VarPrd = Banco.OpenRecordset("SELECT BAR.*, PRD.Descrição_Produto, Prd.Validade, BAR.BarBarra FROM BAR INNER JOIN PRD ON BAR.Produto = PRD.CodPrd ORDER BY BAR.BarBarra")

VarTotal = 0
txtBarra = ""
txtBarra.SetFocus
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Data1.DatabaseName = Caminho & "\Dados.Mdb"

End Sub

Private Sub txtBarra_KeyPress(KeyAscii As Integer)
Dim SQL As String
Dim VarItem As Currency
Dim Criterio As String

If KeyAscii = 13 Then
        
    Criterio = "Bar.BarBarra = " & """" & txtBarra & """"
    
    VarPrd.FindFirst Criterio
    
    If Not VarPrd.NoMatch Then
    
        Criterio = "Produto = " & VarPrd!Produto
        
        Data1.Recordset.FindFirst Criterio
        
        If Not Data1.Recordset.NoMatch Then
            
    '        Set LstPesq = Lista2.ListItems.Add(, , VarPrd("Produto"))
        
    '        LstPesq.SubItems(1) = psCompString(3, Left(VarPrd!Descrição_Produto, 40), 40)
    '        LstPesq.SubItems(2) = psCompString(3, Left(VarPrd.Fields("Bar.BarBarra"), 10), 10)
            Data1.Recordset.Edit
            If Data1.Recordset.Fields("Validade") = 0.5 Then
                   VarTotal = VarTotal + 0.5
                   Data1.Recordset.Fields("Aferido") = Data1.Recordset.Fields("Aferido") + 0.5
            ElseIf Data1.Recordset.Fields("Validade") = 2 Then
                   Data1.Recordset.Fields("Aferido") = Data1.Recordset.Fields("Aferido") + 2
                   VarTotal = VarTotal + 2
               Else
                   Data1.Recordset.Fields("Aferido") = Data1.Recordset.Fields("Aferido") + 1
                   VarTotal = VarTotal + 1
            End If
            Data1.Recordset.Fields("Saldo") = Data1.Recordset.Fields("Qtd") - Data1.Recordset.Fields("Aferido")
            Data1.Recordset.Update
            
            
    '        TxtTotal = Format(VarTotal, "###,##0.00")
        Else
            MsgBox "Item não encontrado !!", vbInformation
        End If
    Else
        MsgBox "Item não encontrado !!", vbInformation
        
    End If
    
    txtBarra = ""
    txtBarra.SetFocus

End If

End Sub
