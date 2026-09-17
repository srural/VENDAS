VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormOferta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7290
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   9
      Top             =   3735
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   7
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Carta de Promoção"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
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
         Width           =   5820
      End
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4770
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Ent Where Tipo = 1 or Tipo = 5 Or Tipo = 6 Or Tipo = 7 Order by Nome"
      Top             =   855
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox ListaEnt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   2205
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1260
      Width           =   4830
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir"
      Top             =   2475
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1095
      Left            =   270
      TabIndex        =   0
      Top             =   765
      Width           =   1680
      Begin VB.OptionButton Opt 
         Caption         =   "Diversos"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   549
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin MSDBCtls.DBCombo DbcDados 
      Bindings        =   "FormOferta.frx":0000
      Height          =   315
      Index           =   0
      Left            =   2205
      TabIndex        =   5
      Top             =   900
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa Nome"
      Height          =   195
      Left            =   2205
      TabIndex        =   6
      Top             =   675
      Width           =   1110
   End
End
Attribute VB_Name = "FormOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabela As Recordset
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Periodo As String
Dim Formula  As String
Dim y As Integer
Dim Flag As Boolean
Dim SQL As String

'Dim ObjWord As Word.Document
On Error GoTo MergeErr

Formula = ""
'Set ObjWord = GetObject(Caminho & "\Carta.doc", "Word.Document")

MDIPrincipal.RptRel.ReportFileName = Caminho & "\Cobranca.RPT"

If Opt(0).value = True Then
        Formula = ""
Else
    Flag = True
    If ListaEnt.Enabled = True Then
        For y = 0 To ListaEnt.ListCount - 1
            If ListaEnt.Selected(y) = True Then
                If Flag Then
                    Formula = " {Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
                    Flag = False
                Else
                    Formula = Formula & " or  {Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
                End If
            End If
        Next
    End If
End If

MDIPrincipal.RptRel.ReportFileName = Caminho & "\Carta.RPT"
MDIPrincipal.RptRel.SelectionFormula = Formula
Imprimir

'    If Opt(0).Value = True Then
'        SQL = "Select Nome, Endereco, Nro, Cidade, Uf from Ent"
'    Else
'        SQL = "Select Nome, Endereco, Nro, Cidade, Uf from Ent"
'        Flag = True
'        If ListaEnt.Enabled = True Then
'            For y = 0 To ListaEnt.ListCount - 1
'                If ListaEnt.Selected(y) = True Then
'                    If Flag Then
'                        Formula = " Where CodEntidade = " & CLng(Left(ListaEnt.List(y), 6))
'                        Flag = False
'                    Else
'                        Formula = Formula & " or  CodEntidade = " & CLng(Left(ListaEnt.List(y), 6))
'                    End If
'                End If
'            Next
'        End If
'    End If

      'torna o word visível
    'ObjWord.Application.Visible = True
    
    
    ' ObjWord.MailMerge.OpenDataSource _
        Name:=Caminho & "\Dados.mdb", _
        LinkToSource:=True, _
        Connection:="TABLE Ent", _
        SQLStatement:=SQL & Formula
       
       'Executa a mala direta
    
    'ObjWord.MailMerge.Execute
    'Set ObjWord = Nothing



Exit Sub

MergeErr:
         ' campo em branco
            MsgBox Err.Number & vbCr & Err.Description

End Sub

Private Sub DbcDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 9 Or KeyCode = 13 Then
    If Index = 0 Then
        If DbcDados(0).Text = "" Then
            DbcDados(0).SetFocus
            Exit Sub
        End If
        Dim Criterio As String
        If IsNumeric(DbcDados(0).Text) Then
           Criterio = "CodEntidade = " & DbcDados(0)
           DatDados(0).Recordset.FindFirst Criterio
           If DatDados(0).Recordset.NoMatch Then
               MsgBox "Entidade não encontrada !!!", 16
               DbcDados(0).Text = ""
               DbcDados(0).SetFocus
           Else
               DbcDados(0).Text = DatDados(0).Recordset.Fields("Nome")
               ListaEnt.AddItem psCompString(1, DatDados(0).Recordset.Fields("CodEntidade"), 6) & " - " & DatDados(0).Recordset.Fields("Nome")
               ListaEnt.Selected(ListaEnt.ListCount - 1) = True
           End If
        Else
            If DbcDados(0).Text <> "" Then
                If DbcDados(0).BoundText = 0 Then
                    MsgBox "Entidade não encontrada !!!", 16
                    DbcDados(0).SetFocus
                    Exit Sub
                Else
                    Criterio = "CodEntidade = " & DbcDados(0).BoundText
                    DatDados(0).Recordset.FindFirst Criterio
                    If DatDados(0).Recordset.NoMatch Then
                        MsgBox "Entidade não encontrada !!!", 16
                        DbcDados(0).Text = ""
                        DbcDados(0).SetFocus
                    Else
                        DbcDados(0).Text = DatDados(0).Recordset.Fields("Nome")
                        ListaEnt.AddItem psCompString(1, DatDados(0).Recordset.Fields("CodEntidade"), 6) & " - " & DatDados(0).Recordset.Fields("Nome")
                        ListaEnt.Selected(ListaEnt.ListCount - 1) = True
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6
Dim Componente As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next


Set Tabela = Banco.OpenRecordset("Select Ent.CodEntidade, Ent.Nome FROM Ent Order By Ent.Nome")

Tabela.MoveFirst
ListaEnt.Clear
ListaEnt.Enabled = False
DbcDados(0).Enabled = False

End Sub

Private Sub Opt_Click(Index As Integer)

If Index = 1 Then
    DbcDados(0).Enabled = True
    ListaEnt.Enabled = True
    DbcDados(0).SetFocus
Else
    DbcDados(0).Enabled = False
    ListaEnt.Enabled = False
End If

End Sub
