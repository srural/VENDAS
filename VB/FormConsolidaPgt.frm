VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormConsolidaPgt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6825
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   12
      Top             =   6480
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
      TabIndex        =   10
      Top             =   0
      Width           =   11985
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento duplicatas"
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
   Begin VB.ListBox Lista 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1440
      Width           =   11625
   End
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   135
      TabIndex        =   4
      Tag             =   "N"
      Top             =   945
      Width           =   1095
   End
   Begin VB.TextBox TxtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Left            =   1305
      TabIndex        =   3
      Tag             =   "N"
      Top             =   945
      Width           =   5325
   End
   Begin VB.CommandButton BtOk 
      Caption         =   "&Confirma "
      Height          =   330
      Left            =   10440
      TabIndex        =   0
      Top             =   945
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MebValor 
      DataField       =   "Venda"
      Height          =   375
      Left            =   10035
      TabIndex        =   6
      Top             =   5805
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebData 
      Height          =   330
      Left            =   7110
      TabIndex        =   8
      Top             =   945
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Operação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   7110
      TabIndex        =   9
      Top             =   675
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   9405
      TabIndex        =   5
      Top             =   5850
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   1350
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "FormConsolidaPgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arq As Recordset
Public VarOper As Byte

Private Sub BtOK_Click()
Dim Flag As Boolean
Dim y As Integer
Dim VarData As Date
Dim Formula As String
Dim VarCodPed As Double
Dim pedidos As Recordset

Flag = True

Formula = ""
For y = 0 To Lista.ListCount - 1
    If Lista.Selected(y) = True Then
        If Flag Then
            Formula = "Pag.CodDup = " & CLng(Left(Lista.List(y), 6))
            Flag = False
            VarCodPed = CLng(Left(Lista.List(y), 6))
        Else
            Formula = Formula & " or  PAG.CodDup = " & CLng(Left(Lista.List(y), 6))
'            MebData = CDate(Mid(Lista.List(y), 10, 10))
        End If
    End If
Next

If Formula <> "" Then
    If MsgBox("Deseja quitar as duplicatas marcadas", vbYesNo, App.Title) = vbYes Then
    
    
        SQL = " UPDATE PAG SET PAG.DtPgto = #" & Format(MebData, "mm/dd/yyyy") & "#, PAG.VlrPago = [pag].[valor], PAG.Cond = True WHERE " & Formula
        Banco.Execute SQL
        
    End If
Else
    MsgBox "Não tem nenhum pedido marcado para consolidar", vbInformation, App.Title
End If

Unload Me

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
Dim Criterio As String
Dim x As Integer

TxtNome.Text = FormBaixarPag.DbcPesquisa.Text
TxtCodigo.Text = FormBaixarPag.txtCod.Text

MebData = Date
MebValor = 0

Lista.Clear
Set Arq = FormBaixarPag.DatDados(0).Recordset.Clone
If Arq.RecordCount > 0 Then
    Arq.MoveFirst
    While Not Arq.Eof

            Lista.AddItem psCompString(1, Arq!CodDup, 6) & " | " & Left(psCompString(1, Arq!Numero, 10), 10) & " | " & psCompString(1, Arq!NroNota, 6) & " | " & Arq!Data & " | " & Arq!DtVenc & "|" & Alinhar(Arq!Valor, 10, 2) & " | " & Left(Arq!Obs, 15)
            Lista.Selected(Lista.ListCount - 1) = False
        Arq.MoveNext
    Wend
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub Lista_ItemCheck(Item As Integer)

MebValor = 0
For y = 0 To Lista.ListCount - 1
    If Lista.Selected(y) = True Then
        MebValor = CCur(MebValor) + Mid(Lista.List(y), 57, 10)
    End If
Next
MebValor = Format(MebValor, "#,##0.00")


End Sub
