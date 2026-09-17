VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormEstatDamme 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10710
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormEstatClientes.frx":0000
      Height          =   4740
      Left            =   225
      OleObjectBlob   =   "FormEstatClientes.frx":0016
      TabIndex        =   10
      Top             =   1845
      Width           =   8340
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -360
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   5
      Top             =   6840
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
      Width           =   8415
      Begin VB.Data DatVend 
         Connect         =   "Access"
         DatabaseName    =   "\\Servidor-pc\Sidival\Desenvolvimento\Vendas_2_0\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6705
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"FormEstatClientes.frx":13E1
         Top             =   135
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton BtRel 
         Caption         =   "&Imprimir"
         Height          =   330
         Left            =   6930
         TabIndex        =   2
         Top             =   405
         Width           =   1290
      End
      Begin VB.CommandButton BtVisualiza 
         Caption         =   "&Visualiza"
         Height          =   330
         Left            =   5400
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
         Format          =   91226113
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
         Format          =   91226113
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
   Begin VB.Label LblTicket 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8910
      TabIndex        =   14
      Top             =   1935
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   8775
      TabIndex        =   13
      Top             =   1665
      Width           =   1725
   End
   Begin VB.Label LblTotCli 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8910
      TabIndex        =   12
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8775
      TabIndex        =   11
      Top             =   855
      Width           =   1725
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

Set Rs = Banco.OpenRecordset("SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data")
Rs.Requery

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

DatVend.RecordSource = "SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data"
DatVend.Refresh


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

DatVend.DatabaseName = Caminho & "\Dados.Mdb"

DtData1 = Date
DtData2 = Date

DatVend.RecordSource = "SELECT ESTAT.Data, ESTAT.TotCli, ESTAT.MAtend, ESTAT.DAtend, ESTAT.MServicos, ESTAT.DServicos, ESTAT.TicketM, ESTAT.TicketD FROM ESTAT WHERE ESTAT.Data>=#" & Format(DtData1 - 1, "mm/dd/yyyy") & "# And ESTAT.Data<=#" & Format(DtData2 - 1, "mm/dd/yyyy") & "# ORDER BY ESTAT.Data"
DatVend.Refresh

If DatVend.Recordset.RecordCount > 0 Then
    LblTotCli = DatVend.Recordset("TotCli")
    LblTicket = Format(DatVend.Recordset("TicketM"), "#,##0.00")
End If

End Sub



