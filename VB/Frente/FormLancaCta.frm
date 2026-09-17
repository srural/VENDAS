VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormLancaCta 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11100
      TabIndex        =   38
      Top             =   6975
      Width           =   11130
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11085
      TabIndex        =   36
      Top             =   0
      Width           =   11085
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançar Contas"
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
         TabIndex        =   37
         Top             =   45
         Width           =   5325
      End
   End
   Begin VB.ComboBox CmbHis 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormLancaCta.frx":0000
      Left            =   7650
      List            =   "FormLancaCta.frx":0010
      TabIndex        =   7
      Top             =   1530
      Width           =   3210
   End
   Begin VB.TextBox TxtPcl 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   285
      TabIndex        =   2
      Top             =   1530
      Width           =   915
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormLancaCta.frx":0044
      Left            =   5130
      List            =   "FormLancaCta.frx":005D
      TabIndex        =   1
      Top             =   945
      Width           =   1140
   End
   Begin VB.Data DatDados 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4860
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Ent.CodEntidade, Ent.Nome From Ent ORDER BY Ent.Nome"
      Top             =   3825
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton CmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   6120
      TabIndex        =   32
      Top             =   6390
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa"
      Height          =   390
      Left            =   6120
      TabIndex        =   31
      Top             =   5895
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
      Height          =   1320
      Left            =   2700
      TabIndex        =   19
      Top             =   5400
      Width           =   2670
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dias em Atraso:"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label LblAtraso 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1620
         TabIndex        =   25
         Top             =   900
         Width           =   135
      End
      Begin VB.Label LblCredito 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1305
         TabIndex        =   23
         Top             =   630
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Crédito:"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   630
         Width           =   540
      End
      Begin VB.Label LblPrazo 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1305
         TabIndex        =   21
         Top             =   315
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prazo: "
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Item"
      Enabled         =   0   'False
      Height          =   315
      Left            =   285
      TabIndex        =   9
      Top             =   5445
      Width           =   1185
   End
   Begin VB.Data DatDados 
      Caption         =   "Dup"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2295
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Dup"
      Top             =   3780
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormLancaCta.frx":007D
      Height          =   3180
      Left            =   285
      OleObjectBlob   =   "FormLancaCta.frx":0097
      TabIndex        =   14
      Top             =   2010
      Width           =   10530
   End
   Begin VB.CommandButton CmdLancar 
      Caption         =   "&Lançar"
      Height          =   390
      Left            =   6120
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox TxtDoc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6030
      TabIndex        =   6
      Top             =   1530
      Width           =   1455
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormLancaCta.frx":1E6A
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Top             =   945
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Custo"
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Tag             =   "S"
      Top             =   1530
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebSubTot 
      DataField       =   "Custo"
      Height          =   405
      Left            =   9045
      TabIndex        =   15
      Tag             =   "S"
      Top             =   5355
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebJuros 
      DataField       =   "Custo"
      Height          =   405
      Left            =   9045
      TabIndex        =   17
      Tag             =   "S"
      Top             =   5895
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTotal 
      DataField       =   "Custo"
      Height          =   405
      Left            =   9045
      TabIndex        =   27
      Tag             =   "S"
      Top             =   6390
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483633
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebCtaPagas 
      DataField       =   "Custo"
      Height          =   405
      Left            =   285
      TabIndex        =   29
      Tag             =   "S"
      Top             =   6300
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebEmiss 
      DataField       =   "Custo"
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Tag             =   "S"
      Top             =   1530
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebVenc 
      DataField       =   "Custo"
      Height          =   315
      Left            =   2925
      TabIndex        =   4
      Tag             =   "S"
      Top             =   1530
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nro Parcela"
      Height          =   195
      Left            =   285
      TabIndex        =   35
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Emissão"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   34
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   5130
      TabIndex        =   33
      Top             =   735
      Width           =   315
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Contas Pagas"
      Height          =   195
      Index           =   4
      Left            =   285
      TabIndex        =   30
      Top             =   6060
      Width           =   1395
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Index           =   2
      Left            =   8505
      TabIndex        =   28
      Top             =   6390
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Historico"
      Height          =   195
      Left            =   7695
      TabIndex        =   24
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Juros"
      Height          =   195
      Index           =   1
      Left            =   8505
      TabIndex        =   18
      Top             =   5895
      Width           =   375
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal"
      Height          =   195
      Index           =   0
      Left            =   8235
      TabIndex        =   16
      Top             =   5385
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento"
      Height          =   195
      Left            =   6030
      TabIndex        =   13
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Vencimento"
      Height          =   195
      Index           =   15
      Left            =   2925
      TabIndex        =   12
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor"
      Height          =   195
      Index           =   5
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Cliente"
      Height          =   195
      Left            =   285
      TabIndex        =   10
      Top             =   735
      Width           =   1170
   End
End
Attribute VB_Name = "FormLancaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbTipo_GotFocus()
CmbTipo.ListIndex = 0
End Sub

Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub CmdExcluir_Click()
If DatDados(0).Recordset.RecordCount <> 0 Then
    If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
        DatDados(0).Recordset.Delete
        DatDados(0).Refresh
'        SomaGrid
    End If
End If
End Sub

Private Sub CmdLancar_Click()
Dim sql As String
Dim Pedido As Long
On Error GoTo TrataErro
If IsNull(MebValor) Or MebValor = "" Then
    MsgBox "Valor não pode ser nulo !", vbExclamation, App.Title
    MebValor.SetFocus
    Exit Sub
End If

If IsNull(MebEmiss) Or MebEmiss = "" Then
    MsgBox "Data Emissão não pode ser nulo !", vbExclamation, App.Title
    MebEmiss.SetFocus
    Exit Sub
End If

If IsNull(MebVenc) Or MebVenc = "" Then
    MsgBox "Vencimento não pode ser nulo !", vbExclamation, App.Title
    MebVenc.SetFocus
    Exit Sub
End If

sql = "INSERT INTO DUP ( DtEmiss, NroPcl, Entidade, DtVenc, Valor, Cond, Obs, Tipo_Dup, Documento  ) VALUES ( "
sql = sql & "#" & Format(MebEmiss, "mm/dd/yyyy") & "# , "
sql = sql & TxtPcl & ", "
sql = sql & DatDados(1).Recordset.Fields("CodEntidade") & " , #"
sql = sql & Format(MebVenc, "mm/dd/yyyy") & "# , "
sql = sql & Numero(MebValor.Text, 2) & " ,"
sql = sql & 0 & " ,"
sql = sql & """" & CMBOBS & """" & " ,"
sql = sql & """" & CmbTipo & """" & " ,"
sql = sql & """" & TxtDoc & """" & ")"
FrmFrente.Banco.Execute sql

CmdLimpa_Click
Atualiza
DatDados(0).Recordset.MoveLast
'SomaGrid
DbcPesquisa.SetFocus

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdLimpa_Click()
'DbcPesquisa = ""
MebValor = ""
TxtDoc = ""
CMBOBS = ""
TxtPcl = ""
CmbTipo.ListIndex = 1
MebEmiss = "__/__/____"
MebVenc = "__/__/____"

Atualiza
'SomaGrid
DbcPesquisa.SetFocus

End Sub


Private Sub DbcPesquisa_LostFocus()
If DbcPesquisa.Text <> "" Then
    If Not IsNumeric(DbcPesquisa.Text) Then
        If IsNumeric(DbcPesquisa.BoundText) Then
            Pesquisa DbcPesquisa.BoundText
        End If
    Else
        Pesquisa DbcPesquisa.Text
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Pesquisa(Codigo As Integer)

Dim Criterio As String
   Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
   Criterio = Criterio & Codigo
   DatDados(1).Recordset.FindFirst Criterio
   If DatDados(1).Recordset.NoMatch Then
       MsgBox "Entidade não Localizado", vbExclamation, App.Title
       DbcPesquisa.Text = ""
       DbcPesquisa.SetFocus
   Else
       DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
       TxtPcl = 1
       MebEmiss = Date
       MebVenc = Date + 30
       Atualiza
       CmbTipo.SetFocus
   End If

End Sub



Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

'Me.Left = 700
'Me.Top = 200


Atualiza

If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveLast
End If

End Sub

Private Sub Form_Load()
Dim Componente As Variant
'Set Registros = Banco.OpenRecordset("SELECT *  FROM Dup")


For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

End Sub


Private Sub SomaGrid()
Dim soma As Recordset
Dim Total As Currency
Dim Flag As Boolean
Dim DtTmp As Date
Dim DtAtraso As Integer
Dim Correcao As Currency
Dim TotJuros As Currency
Dim TotCtaPagas As Currency
Dim TotSaldo As Currency
Flag = True
TotJuros = 0
TotCtaPagas = 0
Set soma = DatDados(0).Recordset.Clone
If soma.RecordCount <> 0 Then
    soma.MoveFirst
    Do Until soma.EOF
        If soma.Fields("Cond") = 0 Then
            DtTmp = soma.Fields("Data")
            If Flag Then
                DtTmp = soma.Fields("Data")
                If DtTmp < Date - 30 Then
                    LblAtraso.Caption = (Date - DtTmp) - IIf(IsNull(DatDados(1).Recordset.Fields("Prazo")), 0, DatDados(1).Recordset.Fields("Prazo"))
                Else
                    LblAtraso.Caption = "0"
                End If
                Flag = False
            End If
            DtAtraso = (Date - DtTmp) - IIf(IsNull(DatDados(1).Recordset.Fields("Prazo")), 0, DatDados(1).Recordset.Fields("Prazo"))
            soma.Edit
            Correcao = soma.Fields("Valor") * (((DatDados(1).Recordset.Fields("Correção") / 100) / 30) * DtAtraso)
            If Correcao > 0 Then
                soma.Fields("Juros") = Correcao
                TotJuros = TotJuros + Correcao
            Else
                soma.Fields("Juros") = 0
            End If
            soma.Update
            Total = Total + soma.Fields("Valor")
        Else
            If soma.Fields("OBS") <> "SALDO" Then
                TotCtaPagas = TotCtaPagas + soma.Fields("Valor")
            End If
        End If
        soma.MoveNext
    Loop
End If
MebSubTot = Total
MebJuros = TotJuros
MebTotal = CDec(MebSubTot) + TotJuros
MebCtaPagas = TotCtaPagas
End Sub


Private Sub Atualiza()
DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " AND COND = FALSE ORDER BY DtEmiss, NroPcl"
DatDados(0).Refresh
End Sub

