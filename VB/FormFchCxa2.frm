VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormFchCxa2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9165
   Begin VB.CommandButton BtCredito 
      Cancel          =   -1  'True
      Caption         =   "Exclui &Crédito"
      Height          =   330
      Left            =   225
      TabIndex        =   15
      Top             =   2655
      Width           =   1335
   End
   Begin VB.CommandButton BtDebito 
      Caption         =   "Exclui &Débito"
      Height          =   330
      Left            =   225
      TabIndex        =   14
      Top             =   4950
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MebCredito 
      Height          =   375
      Left            =   4230
      TabIndex        =   7
      Top             =   5220
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Data DatDados 
      Caption         =   "Debito"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * from cxa "
      Top             =   4095
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DatDados 
      Caption         =   "Credito"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3195
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * from cxa "
      Top             =   1980
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CommandButton CmdVisual 
      Caption         =   "&Visualiza"
      Height          =   375
      Left            =   7335
      TabIndex        =   3
      Top             =   315
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Fecha Caixa"
      Height          =   375
      Left            =   5310
      TabIndex        =   2
      Top             =   315
      Width           =   1560
   End
   Begin MSMask.MaskEdBox MebData 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebValor 
      Height          =   375
      Left            =   2250
      TabIndex        =   1
      Top             =   315
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDebito 
      Height          =   375
      Left            =   5805
      TabIndex        =   8
      Top             =   5220
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebSaldo 
      Height          =   375
      Left            =   7380
      TabIndex        =   9
      Top             =   5220
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormFchCxa2.frx":0000
      Height          =   1770
      Left            =   225
      OleObjectBlob   =   "FormFchCxa2.frx":001A
      TabIndex        =   6
      Top             =   3150
      Width           =   8655
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FormFchCxa2.frx":13DD
      Height          =   1770
      Left            =   225
      OleObjectBlob   =   "FormFchCxa2.frx":13F7
      TabIndex        =   13
      Top             =   855
      Width           =   8655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7380
      TabIndex        =   12
      Top             =   4995
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Débito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5805
      TabIndex        =   11
      Top             =   4995
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4230
      TabIndex        =   10
      Top             =   4995
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Retirada no Fechamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   5
      Top             =   45
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   45
      Width           =   435
   End
End
Attribute VB_Name = "FormFchCxa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub BtCredito_Click()
If MsgBox("Confirma Exclusão do Crédito ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(0).Recordset.RecordCount <> 0 Then
        DatDados(0).Recordset.Delete
        DatDados(0).Refresh
        CmdVisual_Click
    End If
End If
End Sub

Private Sub BtDebito_Click()
If MsgBox("Confirma Exlusão do Débito ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(2).Recordset.RecordCount <> 0 Then
        DatDados(2).Recordset.Delete
        DatDados(2).Refresh
        CmdVisual_Click
    End If
End If
End Sub

Private Sub CmdOk_Click()
Dim Venda As Recordset
Dim Caixa As Recordset
Dim Rec As Recordset
Dim Sql As String
Dim PgtAnt As Byte
Dim Total As Currency
Dim Texto As String

If MebValor = "" Then
    MebValor.SetFocus
    Exit Sub
End If

If MsgBox("Confirma Fechamento do Caixa ?", vbYesNo, App.Title) = vbYes Then
    Set Caixa = Banco.OpenRecordset("Select * From Cxa")
    
'    For PgtAnt = 1 To 5
        Set Venda = Banco.OpenRecordset("SELECT PED.CodPed, PED.Operacao, PED.DataEmiss, PED.Total FROM PED WHERE PED.Operacao = 1  AND PED.DataEmiss =#" & Format(MebData, "MM/DD/YYYY") & "#")
        Venda.Requery
        If Venda.RecordCount <> 0 Then
            Venda.MoveFirst
            While Not Venda.EOF
                 Total = Total + CCur(Venda("Total"))
                 Venda.MoveNext
            Wend
'            If PgtAnt = 1 Then
                Texto = "VENDA VISTA"
'            End If
'            If PgtAnt = 2 Then
'                Texto = "CHEQUE"
'            End If
'            If PgtAnt = 3 Then
'                Texto = "CHEQUE-PRÉ"
'            End If
'            If PgtAnt = 4 Then
'                Texto = "CARTÃO"
'            End If
'            If PgtAnt = 5 Then
'                Texto = "TICKET"
'            End If
            Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, D_C ) VALUES ( "
            Sql = Sql + """" & Maquina & """" & ", "
            Sql = Sql & """" & Texto & """" & " , #"
            Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
            Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
            Sql = Sql & Numero(CStr(Total)) & " ,"
            Sql = Sql & """" & "C" & """" & ")"
            Banco.Execute Sql
            Total = 0
       End If
'    Next
    Set Rec = Banco.OpenRecordset("SELECT REC.*, REC.Data, REC.Valor FROM REC WHERE REC.Data=#" & Format(MebData, "MM/DD/YYYY") & "#")
    Rec.Requery
    If Rec.RecordCount <> 0 Then
         Rec.MoveFirst
         While Not Rec.EOF
              Total = Total + CCur(Rec("Rec.Valor"))
              Rec.MoveNext
         Wend
         Texto = "RECEBIMENTO"
         Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, D_C ) VALUES ( "
         Sql = Sql + """" & Maquina & """" & ", "
         Sql = Sql & """" & Texto & """" & " , #"
         Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
         Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
         Sql = Sql & Numero(CStr(Total)) & " ,"
         Sql = Sql & """" & "C" & """" & ")"
         Banco.Execute Sql
         Total = 0
    End If
    
    Sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, D_C ) VALUES ( "
    Sql = Sql + """" & Maquina & """" & ", "
    Sql = Sql & """" & "FECHAMENTO" & """" & " , #"
    Sql = Sql & Format(MebData, "mm/dd/yyyy") & "# , #"
    Sql = Sql & Format(Time, "hh:mm:ss") & "# , "
    Sql = Sql & Numero(MebValor) & " ,"
    Sql = Sql & """" & "D" & """" & ")"
    Banco.Execute Sql
    CmdVisual_Click
    MebValor.Text = ""
    DatDados(0).Refresh
    MebData.SetFocus
End If

End Sub

Private Sub DbcHist_LostFocus()
DbcHist.Text = Format(DbcHist, ">")
End Sub

Private Sub CmdVisual_Click()
MebCredito = 0
MebDebito = 0
DatDados(0).RecordSource = "SELECT MAQUINA, HISTORICO, HORA, VALOR, D_C FROM CXA WHERE DATA = #" & Format(MebData, "MM/DD/YYYY") & "#  AND D_C = 'C' ORDER BY HORA"
DatDados(0).Refresh
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveFirst
    While Not DatDados(0).Recordset.EOF
        MebCredito = CCur(MebCredito) + CCur(DatDados(0).Recordset.Fields("Valor"))
        DatDados(0).Recordset.MoveNext
    Wend
End If
DatDados(2).RecordSource = "SELECT MAQUINA, HISTORICO, HORA, VALOR, D_C FROM CXA WHERE DATA = #" & Format(MebData, "MM/DD/YYYY") & "#  AND D_C = 'D' ORDER BY HORA"
DatDados(2).Refresh

If DatDados(2).Recordset.RecordCount <> 0 Then
    DatDados(2).Recordset.MoveFirst
    While Not DatDados(2).Recordset.EOF
        MebDebito = CCur(MebDebito) + CCur(DatDados(2).Recordset.Fields("Valor"))
        DatDados(2).Recordset.MoveNext
    Wend
End If
MebSaldo = CCur(MebCredito) - CCur(MebDebito)

End Sub

Private Sub Form_Activate()
MebData.Text = Date
DatDados(0).RecordSource = "SELECT MAQUINA, HISTORICO, HORA, VALOR,  D_C FROM CXA WHERE DATA = #" & Format(MebData, "MM/DD/YYYY") & "# AND MAQUINA = 'NENHUMA' AND D_C = 'C' ORDER BY HORA"
DatDados(0).Refresh
DatDados(2).RecordSource = "SELECT MAQUINA, HISTORICO, HORA, VALOR,  D_C FROM CXA WHERE DATA = #" & Format(MebData, "MM/DD/YYYY") & "# AND MAQUINA = 'NENHUMA' AND D_C = 'D' ORDER BY HORA"
DatDados(2).Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

End Sub

