VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormConsultaPed 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11970
   Begin VB.Data DatDados 
      Caption         =   "Hist"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   8865
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome FROM  HISTEST GROUP BY Nome ORDER BY Nome"
      Top             =   2835
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -315
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   10
      Top             =   7650
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
      ScaleWidth      =   11985
      TabIndex        =   8
      Top             =   0
      Width           =   11985
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta Pedido"
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
         TabIndex        =   9
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.Data DupAtual 
      Caption         =   "DupAtual"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from DUP"
      Top             =   3450
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data DupRemoto 
      Caption         =   "DupRemoto"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5415
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from DUP"
      Top             =   3420
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Data IteAtual 
      Caption         =   "IteAtual"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5535
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from Itp"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data IteRemoto 
      Caption         =   "IteRemoto"
      Connect         =   "Access"
      DatabaseName    =   "E:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5505
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from Itp"
      Top             =   4890
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data PedRemoto 
      Caption         =   "PedRemoto"
      Connect         =   "Access"
      DatabaseName    =   "E:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2985
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from Ped"
      Top             =   4890
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data PedAtual 
      Caption         =   "PedAtual"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3015
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from Ped"
      Top             =   4230
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Frame Frame6 
      Height          =   540
      Left            =   8190
      TabIndex        =   4
      Top             =   6885
      Width           =   3555
      Begin VB.OptionButton Opt 
         Caption         =   "Local 1"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   180
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Local 2"
         Height          =   225
         Index           =   1
         Left            =   1230
         TabIndex        =   6
         Top             =   180
         Width           =   945
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Local 3"
         Height          =   225
         Index           =   2
         Left            =   2385
         TabIndex        =   5
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.CommandButton Importa 
      Caption         =   "&Importa"
      Height          =   405
      Left            =   3615
      TabIndex        =   3
      Top             =   6990
      Width           =   1515
   End
   Begin VB.CommandButton BtAtualiza 
      Caption         =   "&Atualiza Lista"
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   6990
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Left            =   195
      TabIndex        =   1
      Top             =   6990
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   49217537
      CurrentDate     =   37501
   End
   Begin VB.Data DatPed 
      Caption         =   "Pedido"
      Connect         =   "Access"
      DatabaseName    =   "E:\(Rede)\Desenvolvimento\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormConsultaPed.frx":0000
      Top             =   2340
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormConsultaPed.frx":00CE
      Height          =   5970
      Left            =   180
      OleObjectBlob   =   "FormConsultaPed.frx":00E3
      TabIndex        =   0
      Top             =   765
      Width           =   11580
   End
   Begin MSDBCtls.DBCombo TxtNome 
      Bindings        =   "FormConsultaPed.frx":1181
      Height          =   360
      Left            =   5400
      TabIndex        =   11
      Top             =   7020
      Width           =   2670
      _ExtentX        =   4710
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
      Left            =   5400
      TabIndex        =   12
      Top             =   6795
      Width           =   1065
   End
End
Attribute VB_Name = "FormConsultaPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VarLocal As String
Public VarLocal1 As String
Public VarLocal2 As String
Public VarLocal3 As String
Public Senha As Boolean
Public VarDatTabela As Recordset

Private Sub BtAtualiza_Click()
              
DatPed.RecordSource = "SELECT PED.CodPed, PED.DataEmiss, PED.Total, PED.Desconto, ENT.Nome, PED.Operacao, Ped.NroNt FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade WHERE PED.DataEmiss =# " & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PED.CodPed DESC"

DatPed.Refresh
If DatPed.Recordset.RecordCount <> 0 Then
     DatPed.Recordset.MoveFirst
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
BtAtualiza_Click

End Sub

Private Sub Form_Load()
Set Formulário = Me

Senha = True

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next


If Not IsNull(VarSenhaBco) Then
    If VarSenhaBco <> "" Then
        Senha = False
        FormSenha2.Show vbModal
    End If
End If
If Senha = False Then
    Unload Me
End If

DtData.value = Date

If LerINI("LocalBco", "Local1", App.Path & "\config.ini") <> "" Then
    VarLocal1 = LerINI("LocalBco", "Local1", App.Path & "\config.ini")
    Opt(0).Enabled = True
Else
    VarLocal1 = ""
    Opt(0).Enabled = False
End If

If LerINI("LocalBco", "Local2", App.Path & "\config.ini") <> "" Then
    VarLocal2 = LerINI("LocalBco", "Local2", App.Path & "\config.ini")
    Opt(1).Enabled = True
Else
    VarLocal2 = ""
    Opt(1).Enabled = False
End If

If LerINI("LocalBco", "Local3", App.Path & "\config.ini") <> "" Then
    VarLocal3 = LerINI("LocalBco", "Local3", App.Path & "\config.ini")
    Opt(2).Enabled = True
Else
    VarLocal3 = ""
    Opt(2).Enabled = False
End If

If Opt(0).value = True Then
        VarLocal = VarLocal1
    ElseIf Opt(1).value = True Then
        VarLocal = varloca2
    Else
        Opt(2).value = True
        VarLocal = VarLocal3
End If
    
DatPed.DatabaseName = VarLocal & "\Dados.mdb"
PedAtual.DatabaseName = Caminho & "\Dados.mdb"
PedRemoto.DatabaseName = VarLocal & "\Dados.mdb"
IteAtual.DatabaseName = Caminho & "\Dados.mdb"
IteRemoto.DatabaseName = VarLocal & "\Dados.mdb"
DupAtual.DatabaseName = Caminho & "\Dados.mdb"
DupRemoto.DatabaseName = VarLocal & "\Dados.mdb"



End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Importa_Click()
Dim X As Integer
Dim VarCodPed As String
Dim Criterio As String
Dim TotPed As Currency

TotPed = 0

If TxtNome <> "" Then
    Set VarDatTabela = Banco.OpenRecordset("SELECT  *  FROM  HISTEST WHERE Nome = " & """" & Trim(TxtNome) & """" & " ORDER BY CodPrd")
    SQL = "SELECT  *  FROM  HISTEST WHERE Nome = " & """" & Trim(TxtNome) & """" & " ORDER BY CodPrd"
    VarDatTabela.Requery
End If

If DatPed.Recordset.RecordCount > 0 Then
    If MsgBox("Confirma a Importação do Pedido Nro " & DBGrid.Columns(0), vbYesNo, App.Title) = vbYes Then
            PedRemoto.RecordSource = "SELECT PED.* FROM Ped WHERE PED.CodPed =  " & DBGrid.Columns(0)
            PedRemoto.Refresh
            IteRemoto.RecordSource = "SELECT ITP.* FROM Itp WHERE Itp.Pedido =  " & DBGrid.Columns(0)
            IteRemoto.Refresh
            DupRemoto.RecordSource = "SELECT DUP.* FROM Dup WHERE Pedido =  " & DBGrid.Columns(0)
            DupRemoto.Refresh
                        
            PedAtual.RecordSource = "SELECT *  FROM PED "
            IteAtual.RecordSource = "SELECT *  FROM ITP "
            DupAtual.RecordSource = "SELECT *  FROM DUP "
            
            If PedRemoto.Recordset.RecordCount > 0 Then
                PedRemoto.Recordset.MoveFirst
                While Not PedRemoto.Recordset.Eof
                    PedAtual.Recordset.AddNew
                    For X = 1 To 42
                        PedAtual.Recordset.Fields(X) = PedRemoto.Recordset.Fields(X)
                    Next
                    PedAtual.Recordset.Update
                    PedAtual.Recordset.Bookmark = PedAtual.Recordset.LastModified

                    VarCodPed = PedAtual.Recordset("CodPed")
                    PedRemoto.Recordset.MoveNext
                Wend
                
            End If
            
            If IteRemoto.Recordset.RecordCount > 0 Then
                IteRemoto.Recordset.MoveFirst
                While Not IteRemoto.Recordset.Eof
                    IteAtual.Recordset.AddNew
                    For X = 1 To 17
                        If IteAtual.Recordset.Fields(X).Name = "Entregue" Then
                            IteAtual.Recordset.Fields(X) = False
                        Else
                            IteAtual.Recordset.Fields(X) = IteRemoto.Recordset.Fields(X)
                        End If
                    Next
                    If TxtNome <> "" Then
                        Criterio = "CodPrd = " & IteRemoto.Recordset.Fields("Produto")
                        VarDatTabela.FindFirst Criterio
                        If Not VarDatTabela.NoMatch Then
                            IteAtual.Recordset.Fields("ValorUnit") = VarDatTabela.Fields("Valor")
                            IteAtual.Recordset.Fields("Valor") = IteRemoto.Recordset.Fields("Qtd") * VarDatTabela.Fields("Valor")
                            IteAtual.Recordset.Fields("VlrVdaDia") = IteRemoto.Recordset.Fields("Qtd") * VarDatTabela.Fields("Valor")
                            TotPed = TotPed + (IteRemoto.Recordset.Fields("Qtd") * VarDatTabela.Fields("Valor"))
                        End If
                    End If
                    
                    IteAtual.Recordset.Fields("Pedido") = VarCodPed
                    IteAtual.Recordset.Update
                    IteRemoto.Recordset.MoveNext
                Wend
                
            End If
            If TxtNome <> "" Then
                PedAtual.Recordset.Edit
                PedAtual.Recordset.Fields("Total") = TotPed
                PedAtual.Recordset.Update
            End If
            
            If DupRemoto.Recordset.RecordCount > 0 Then
                DupRemoto.Recordset.MoveFirst
                While Not DupRemoto.Recordset.Eof
                    DupAtual.Recordset.AddNew
                    For X = 1 To 17
                        DupAtual.Recordset.Fields(X) = DupRemoto.Recordset.Fields(X)
                    Next
                    DupAtual.Recordset.Fields("Pedido") = VarCodPed
                    DupAtual.Recordset.Update
                    DupRemoto.Recordset.MoveNext
                Wend
                
            End If
            
                       
            MsgBox "O pedido passou a ter o numero " & VarCodPed, vbInformation, App.Title
                        
    End If

End If

End Sub
