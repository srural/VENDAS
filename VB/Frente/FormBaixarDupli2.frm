VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormBaixarDupli2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixar Parcelas"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCod 
      Enabled         =   0   'False
      Height          =   330
      Left            =   6165
      TabIndex        =   35
      Top             =   495
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   690
      Left            =   3555
      TabIndex        =   31
      Top             =   4995
      Width           =   2175
      Begin MSMask.MaskEdBox MebTotal 
         Height          =   510
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483640
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton CmbImprimir 
      Caption         =   "&Imprimir"
      Height          =   630
      Left            =   4770
      Picture         =   "FormBaixarDupli2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Imprimir Extrato"
      Top             =   4095
      Width           =   765
   End
   Begin VB.ComboBox LstTipo 
      Height          =   315
      ItemData        =   "FormBaixarDupli2.frx":066A
      Left            =   1665
      List            =   "FormBaixarDupli2.frx":067D
      TabIndex        =   6
      Top             =   6210
      Width           =   1680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagamento em Cheque"
      Height          =   1470
      Left            =   5850
      TabIndex        =   26
      Top             =   5085
      Width           =   5475
      Begin VB.Data DatDados 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   2
         Left            =   4095
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT REC.Terceiro FROM REC GROUP BY REC.Terceiro ORDER BY REC.Terceiro"
         Top             =   270
         Visible         =   0   'False
         Width           =   1230
      End
      Begin MSDBCtls.DBCombo DBCTerceiro 
         Bindings        =   "FormBaixarDupli2.frx":06BA
         Height          =   420
         Left            =   135
         TabIndex        =   9
         Top             =   900
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   741
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MebChPre 
         Height          =   375
         Left            =   1575
         TabIndex        =   8
         Top             =   225
         Width           =   2040
         _ExtentX        =   3598
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Datado para"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   28
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   27
         Top             =   645
         Width           =   1815
      End
   End
   Begin VB.CheckBox ChkPagos 
      Caption         =   "&Mostrar os Pagos"
      Height          =   330
      Left            =   7875
      TabIndex        =   24
      Top             =   495
      Width           =   1590
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Parcela"
      Height          =   285
      Left            =   9990
      TabIndex        =   11
      Top             =   495
      Width           =   1365
   End
   Begin VB.TextBox TxtDupl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1665
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3915
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   915
      Left            =   5850
      TabIndex        =   19
      Top             =   4005
      Width           =   5505
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   405
         Left            =   180
         TabIndex        =   13
         Tag             =   "S"
         Top             =   405
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
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
      Begin MSMask.MaskEdBox MebPagos 
         DataField       =   "Custo"
         Height          =   405
         Left            =   1980
         TabIndex        =   14
         Tag             =   "S"
         Top             =   405
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
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
      Begin MSMask.MaskEdBox MebAPagar 
         DataField       =   "Custo"
         Height          =   405
         Left            =   3780
         TabIndex        =   15
         Tag             =   "S"
         Top             =   405
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   -2147483624
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
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Parcelas á pagar"
         Height          =   195
         Index           =   2
         Left            =   3780
         TabIndex        =   22
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Parcelas pagas"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   21
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Total"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   20
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa"
      Height          =   285
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   90
      Width           =   1365
   End
   Begin VB.CommandButton CmdPgto 
      Caption         =   "Pagar Parcela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3780
      TabIndex        =   7
      Top             =   5940
      Width           =   1770
   End
   Begin VB.Data DatDados 
      Caption         =   "Dupli"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT  *  FROM Dup  ORDER BY DtVenc"
      Top             =   1935
      Visible         =   0   'False
      Width           =   1545
   End
   Begin MSDBGrid.DBGrid dbgrid1 
      Bindings        =   "FormBaixarDupli2.frx":06D4
      Height          =   2850
      Left            =   135
      OleObjectBlob   =   "FormBaixarDupli2.frx":06EE
      TabIndex        =   12
      Top             =   945
      Width           =   11265
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1350
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From ENT Order by Nome"
      Top             =   495
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSDBCtls.DBCombo DbcPesquisa 
      Bindings        =   "FormBaixarDupli2.frx":2315
      DataField       =   "Nome"
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MebPago 
      DataField       =   "Custo"
      Height          =   360
      Left            =   1665
      TabIndex        =   5
      Tag             =   "S"
      Top             =   5715
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSComCtl2.DTPicker DtData 
      Height          =   360
      Left            =   1665
      TabIndex        =   2
      Top             =   4380
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   36759
   End
   Begin MSMask.MaskEdBox MebDesconto 
      DataField       =   "Custo"
      Height          =   360
      Left            =   1665
      TabIndex        =   3
      Tag             =   "S"
      Top             =   4815
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebJuros 
      DataField       =   "Custo"
      Height          =   360
      Left            =   1665
      TabIndex        =   4
      Tag             =   "S"
      Top             =   5295
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   8421631
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;- #,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   6165
      TabIndex        =   36
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Juros 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "Juros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   34
      Top             =   5310
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3600
      TabIndex        =   33
      Top             =   4815
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Pgto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   6255
      Width           =   840
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "Desconto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   25
      Top             =   4905
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Parcela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Data Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   4425
      Width           =   1425
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   5775
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entidade"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   270
      Width           =   630
   End
End
Attribute VB_Name = "FormBaixarDupli2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public Recebe As Recordset

Private Sub ChkPagos_Click()
If DbcPesquisa.Text <> "" Then
    Pesquisa DbcPesquisa.BoundText
End If
End Sub

Private Sub CmdCancela_Click()
Unload Me
End Sub



Private Sub CmbImprimir_Click()
MDIPrincipal.RptRel.ReportFileName = Caminho & "\relparcelas.RPT"
MDIPrincipal.RptRel.Formulas(0) = "Titulo = 'RELATÓRIO DE PARCELAS  " & "  -  " & NomeEmpresa & "'"
If ChkPagos.Value = 0 Then
    MDIPrincipal.RptRel.SelectionFormula = "{DUP.Entidade} = " & DbcPesquisa.BoundText & " AND {Dup.Cond} = false "
Else
    MDIPrincipal.RptRel.SelectionFormula = "{DUP.Entidade} = " & DbcPesquisa.BoundText
End If
Imprimir

End Sub

Private Sub CmdExcluir_Click()
Dim Criterio As String
If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma Exclusão da Duplicata ?", vbYesNo, App.Title) = vbYes Then
        Criterio = "Numero  = " & """" & TxtDupl & """"
        DatDados(0).Recordset.FindFirst Criterio
        If Not DatDados(0).Recordset.NoMatch Then
            DatDados(0).Recordset.Delete
        End If
        DatDados(0).Refresh
        SomaGrid
        MebPago = 0
        TxtDupl = ""
    End If
End If

End Sub

Private Sub CmdLimpa_Click()
MebPago = 0

DbcPesquisa = ""
TxtCod = ""
MebSubTot = 0
MebPagos = 0
MebDesconto = 0
MebJuros = 0
MebAPagar = 0
MebTotal = 0
DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " ORDER BY DtVenc"
DatDados(0).Refresh
SomaGrid
DbcPesquisa.SetFocus

End Sub

Private Sub CmdPgto_Click()
Dim Criterio As String
Dim Dupli As Long
If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma pagamento ?", vbYesNo, App.Title) = vbYes Then
        Dupli = DatDados(0).Recordset.Fields("CodDup")
        If CCur(MebTotal) = CCur(MebPago) Then
            DatDados(0).Recordset.Edit
            DatDados(0).Recordset.Fields("Cond") = -1
            DatDados(0).Recordset.Fields("DtPgto") = DtData
            DatDados(0).Recordset.Fields("VlrPago") = MebPago
            DatDados(0).Recordset.Update
            DatDados(0).Refresh
         ElseIf CCur(MebPago) > CCur(MebTotal) Then
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("DtPgto") = DtData
                DatDados(0).Recordset.Fields("Cond") = -1
                DatDados(0).Recordset.Fields("VlrPago") = MebPago
                DatDados(0).Recordset.Fields("Saldo") = 0
                DatDados(0).Recordset.Update
                If MebJuros = 0 Or MebJuros = "" Then
                    MsgBox " Troco de R$ " & MebPago - MebTotal, vbInformation, App.Title
                End If
            Else
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("DtPgto") = DtData
                DatDados(0).Recordset.Fields("VlrPago") = MebPago
                DatDados(0).Recordset.Fields("Saldo") = CCur(MebTotal) - CCur(MebPago)
                DatDados(0).Recordset.Update
         End If
         Set Recebe = FrmFrente.Banco.OpenRecordset("SELECT REC.* FROM REC")
         Recebe.AddNew
         Recebe.Fields("CodEnt") = DbcPesquisa.BoundText
         Recebe.Fields("Data") = Date
         Recebe.Fields("Hora") = Time
         Recebe.Fields("Valor") = CCur(MebPago)
         Recebe.Fields("Desconto") = CCur(MebDesconto)
         Recebe.Fields("Tipo") = LstTipo.ListIndex + 1
         If MebChPre <> "__/__/____" Then
              Recebe.Fields("ChPre") = MebChPre
              Recebe.Fields("Terceiro") = DBCTerceiro.Text
         End If
         Recebe.Fields("Coddupli") = Dupli
         Recebe.Fields("Maquina") = Trim(Maquina)
         Recebe.Update
         MebTotal = 0
         MebPago = 0
         MebDesconto = 0
         MebJuros = 0
         CmdPgto.Enabled = False
         TxtDupl = ""
         DatDados(0).Refresh
         SomaGrid
    End If
End If

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

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String
Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
Criterio = Criterio & Codigo
DatDados(1).Recordset.FindFirst Criterio
If DatDados(1).Recordset.NoMatch Then
    MsgBox "Entidade não Localizada", vbExclamation, App.Title
    DbcPesquisa.Text = ""
    DbcPesquisa.SetFocus
Else
    DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
    TxtCod.Text = DatDados(1).Recordset("CodEntidade")
    If ChkPagos.Value = 0 Then
        DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & " AND Cond = FALSE ORDER BY DtVenc"
    Else
        DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & " ORDER BY DtVenc"
    End If
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount <> 0 Then
         DatDados(0).Recordset.MoveLast
    End If
    SomaGrid
End If

End Sub

Private Sub DBGrid1_DBLClick()
TxtDupl.Text = ""
MebPago.Text = 0
If DBGrid1.Columns(7) = "Sim" Then
    MsgBox "Parcela já Paga", vbInformation, App.Title
Else
    TxtDupl.Text = DatDados(0).Recordset.Fields("NroPcl")
    If IsNull(DatDados(0).Recordset.Fields("Saldo")) Or DatDados(0).Recordset.Fields("Saldo") = 0 Then
        MebTotal.Text = DatDados(0).Recordset.Fields("Valor")
    Else
        MebTotal.Text = DatDados(0).Recordset.Fields("Saldo")
    End If
    MebPago.Text = MebTotal.Text
    CmdPgto.Enabled = True
    MebDesconto.SetFocus
End If
End Sub

Private Sub DtData_LostFocus()
MebDesconto.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyPress = 0
'    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
Set Formulário = Me

DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " AND Cond = false AND Saldo = 0 ORDER BY DtVenc"
DatDados(0).Refresh
DtData = Date
MebPago = 0
MebDesconto = 0
MebJuros = 0
MebTotal = 0
LstTipo.ListIndex = 0
CmdPgto.Enabled = False
End Sub

Private Sub Form_Load()
Dim Componente As Variant

'Set Registros = Banco.OpenRecordset("SELECT * FROM Dup")

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next

End Sub

Private Sub LstTipo_LostFocus()
If LstTipo.ListIndex = 1 Or LstTipo.ListIndex = 2 Then
    MebChPre.SetFocus
End If
End Sub

Private Sub MebDesconto_lostfocus()
If MebTotal <> "" And MebDesconto <> "" And IsNumeric(MebDesconto) Then
    MebTotal = CCur(MebTotal) - CCur(MebDesconto)
End If
MebPago = Format(CCur(MebTotal), "#,##0.00;(#,##0.00)")
MebJuros.SetFocus
End Sub

Private Sub SomaGrid()
Dim SomaGrid As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Dim Saldo As Currency
Pago = 0
Apagar = 0
Saldo = 0

'Set SomaGrid = DatDados(0).Recordset.Clone
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveFirst
    Do Until DatDados(0).Recordset.EOF
        If DatDados(0).Recordset.Fields("Cond") = True Then
            Pago = Pago + DatDados(0).Recordset.Fields("VlrPago")
        Else
            Apagar = Apagar + DatDados(0).Recordset.Fields("Valor")
        End If
        If Not IsNull(DatDados(0).Recordset.Fields("Saldo")) Then
            Saldo = Saldo + DatDados(0).Recordset.Fields("Saldo")
        End If
        DatDados(0).Recordset.MoveNext
    Loop
    MebPagos.Text = Pago + Saldo
    MebAPagar.Text = Apagar - Saldo
    MebSubTot = Pago + Apagar

End If

End Sub

Private Sub MebJuros_LostFocus()
If MebTotal <> "" And MebJuros <> "" And IsNumeric(MebJuros) Then
    MebTotal = CCur(MebTotal) + CCur(MebJuros)
End If
MebPago = CCur(MebTotal)

End Sub

'Private Sub MebPago_Change()
'If MebPago = "" Or MebPago = 0 Then
'    CmdPgto.Enabled = False
'Else
'    CmdPgto.Enabled = True
'End If

'End Sub
