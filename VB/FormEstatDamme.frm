VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormEstatDamme 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9180
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormEstatDamme.frx":0000
      Height          =   2625
      Left            =   225
      OleObjectBlob   =   "FormEstatDamme.frx":0016
      TabIndex        =   18
      Top             =   2925
      Width           =   8475
   End
   Begin VB.Data DatVend 
      Caption         =   "Vendedores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   945
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox TotCli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   225
      TabIndex        =   13
      Top             =   2205
      Width           =   1860
   End
   Begin VB.TextBox CliMD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4605
      TabIndex        =   12
      Top             =   2205
      Width           =   1860
   End
   Begin VB.TextBox cliped 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2415
      TabIndex        =   11
      Top             =   2205
      Width           =   1860
   End
   Begin VB.TextBox TicketMedio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6795
      TabIndex        =   10
      Top             =   2205
      Width           =   1860
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   9210
      TabIndex        =   5
      Top             =   5850
      Width           =   9240
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12930
      TabIndex        =   3
      Top             =   0
      Width           =   12930
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatistica Clientes"
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
         Width           =   6630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   5580
      Begin VB.CommandButton BtRel 
         Caption         =   "&Imprimir"
         Height          =   330
         Left            =   5220
         TabIndex        =   2
         Top             =   405
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton BtVisualiza 
         Caption         =   "&Visualiza"
         Height          =   330
         Left            =   3600
         TabIndex        =   0
         Top             =   405
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DtData1 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   86310913
         CurrentDate     =   36641
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   330
         Left            =   1845
         TabIndex        =   7
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   86310913
         CurrentDate     =   36641
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   1845
         TabIndex        =   9
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   17
      Top             =   1890
      Width           =   1995
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes Novos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4590
      TabIndex        =   16
      Top             =   1935
      Width           =   1770
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2385
      TabIndex        =   15
      Top             =   1935
      Width           =   2310
   End
   Begin VB.Label Ticket 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Médio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6750
      TabIndex        =   14
      Top             =   1935
      Width           =   2175
   End
End
Attribute VB_Name = "FormEstatDamme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Venda As Recordset

Private Sub CmdCancela_Click()
Unload Me
End Sub



Private Sub BtRel_Click()
Dim Sql As String
Dim Rs As Recordset

'Set Rs = Banco.OpenRecordset("SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data")
'Rs.Requery

'sql = "SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data"

'Set varPesquisa(1) = Banco.OpenRecordset("SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data")

If Rs.RecordCount > 0 Then


'                FrmRelVda.Sec2(0).Height = "6,35"
'                FrmRelVda.CampoQuebra.Mostrar = True
'                FrmRelVda.CampoQuebra.Campo = "Data"
'                FrmRelVda.rpCabTitulo.Caption = "Relatório de Vendas"
'                FrmRelVda.rpCabSubTitulo.Caption = Titulo
FrmRelGrafico.Data1.Caption = DtData1
FrmRelGrafico.Data2.Caption = DtData2
FrmRelGrafico.CabCli.Caption = "Total de Clientes :" & LblTotCli
FrmRelGrafico.CabTicket.Caption = "Ticket Medio: " & Format(LblTicket, "#,##0.00")
'Set FrmRelGrafico.Rsql = sql
FrmRelGrafico.Config


End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbImprimir_Click()
    
    If MsgBox("Confirma impressão?", 36) = vbYes Then
    
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
        
        Fs2.CopyFile App.Path & "\caixa.txt", Porta
        
'        Fs2.Close
        
    End If

End Sub

Private Sub BtVisualiza_Click()
Dim VarClientes As Recordset
Dim VarPedCli As Recordset
Dim VarServicos  As Recordset
Dim VarTicket As Recordset
Dim VarEstat As Recordset
Dim Criterio As String

Dim TotCliMes As String
Dim TotCliDia As String
Dim TotAtendMes As String
Dim TotAtendDia As String
Dim TotServicosMes As String
Dim TotServicosDia As String
Dim TotticketMes As String
Dim TotTicketDia As String
'DatVend.RecordSource = "SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data"
'DatVend.Refresh

Set VarClientes = Banco.OpenRecordset("SELECT ENT.Tipo, Count(ENT.Nome) AS ContarDeNome FROM ENT GROUP BY ENT.Tipo HAVING (((ENT.Tipo)=1))")
VarClientes.Requery
If VarClientes.RecordCount > 0 Then
    TotCli = VarClientes!ContarDeNome
Else
    TotCli = 0
End If


Set VarClientes = Banco.OpenRecordset("SELECT ENT.Tipo, Count(ENT.Nome) AS ContarDeNome FROM ENT WHERE (ENT.DtCadastro >= #" & Format(DtData1, "mm/dd/yyyy") & "# and ENT.DtCadastro <= #" & Format(DtData2, "mm/dd/yyyy") & "#) GROUP BY ENT.Tipo  HAVING (ENT.Tipo= 1)")
'Sql = "SELECT ENT.Tipo, Count(ENT.Nome) AS ContarDeNome FROM ENT WHERE (ENT.DtCadastro >= #" & Format(DtData1, "mm/dd/yyyy") & "# and ENT.DtCadastro <= #" & Format(DtData2, "mm/dd/yyyy") & "#) GROUP BY ENT.Tipo  HAVING (ENT.Tipo= 1)"

VarClientes.Requery
If VarClientes.RecordCount > 0 Then
    TotCliMes = VarClientes!ContarDeNome
Else
    TotCliMes = 0
End If

CliMD = Trim(TotCliMes)

Set VarPedCli = Banco.OpenRecordset("SELECT Count(PED.Entidade) AS ContarDeEntidade FROM PED WHERE (PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# and PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#) AND ( PED.Operacao = 1 OR  PED.Operacao = 2)")
VarPedCli.Requery
If VarPedCli.RecordCount > 0 Then
    TotAtendMes = VarPedCli!ContarDeEntidade
Else
    TotAtendMes = 0
End If

cliped = Trim(TotAtendMes)


 
Set VarTicket = Banco.OpenRecordset("SELECT Sum(PED.Total) AS SomaDeTotal, Avg(PED.Total) AS MediaDeTotal FROM PED WHERE (PED.DataEmiss >= #" & Format(DtData1, "mm/dd/yyyy") & "# and PED.DataEmiss <= #" & Format(DtData2, "mm/dd/yyyy") & "#) AND ( PED.Operacao = 1 OR  PED.Operacao = 2)")
VarTicket.Requery
If VarTicket.RecordCount > 0 And Not IsNull(VarTicket!MediaDeTotal) Then
    TotticketMes = VarTicket!MediaDeTotal
Else
    TotticketMes = 0
End If

TicketMedio = Format(Trim(TotticketMes), "#,##0.00")

DatVend.RecordSource = "SELECT Vendedores.ENT.Nome, Count(ENT.Nome) AS ContarDeNome FROM ENT INNER JOIN Vendedores ON ENT.Vendedor = Vendedores.CodEntidade WHERE ENT.DtCadastro >= #" & Format(DtData1, "mm/dd/yyyy") & "# and ENT.DtCadastro <= #" & Format(DtData2, "mm/dd/yyyy") & "# GROUP BY Vendedores.ENT.Nome, ENT.Tipo HAVING ENT.Tipo=1"
DatVend.Refresh

'MsgBox DatVend.Recordset.Fields(0)

End Sub

Private Sub Command1_Click()

    FrmRelTXT.Config App.Path & "\EstatVen.Txt"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant




'MebData.Text = Date
'MebData2.Text = Date


Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

'DatVend.DatabaseName = Caminho & "\Dados.Mdb"

DtData1 = Date
DtData2 = Date

'DatVend.RecordSource = "SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1 - 1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2 - 1, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data"
'DatVend.Refresh

'If DatVend.Recordset.RecordCount > 0 Then
'    LblTotCli = DatVend.Recordset("TotCli")
'    LblTicket = Format(DatVend.Recordset("TicketM"), "#,##0.00")
'End If


DatVend.DatabaseName = Caminho & "\Dados.Mdb"


End Sub



