VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormCobranca 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9495
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -360
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   11
      Top             =   4860
      Width           =   14370
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa"
      Height          =   3840
      Left            =   2565
      TabIndex        =   8
      Top             =   720
      Width           =   6405
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
         Height          =   3000
         Left            =   135
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   630
         Width           =   6090
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FormCobranca.frx":0000
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9510
      TabIndex        =   6
      Top             =   0
      Width           =   9510
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Carta de Cobrança"
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
         TabIndex        =   7
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.TextBox TxtDias 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Ent Where Tipo = 1 or Tipo = 5 Or Tipo = 6 Or Tipo = 7 Order by Nome"
      Top             =   3330
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   3870
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1545
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   2040
      Begin VB.OptionButton Opt 
         Caption         =   "Diversos"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   870
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dias"
      Height          =   240
      Left            =   270
      TabIndex        =   5
      Top             =   2655
      Width           =   510
   End
End
Attribute VB_Name = "FormCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabela As Recordset
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula  As String
Dim y As Integer
Dim Flag As Boolean

Data1 = Date - CDec(TxtDias)
Data1 = "Date(" & Year(Data1) & "," & Month(Data1) & "," & Day(Data1) & ")"
'MDIPrincipal.ComRel.ShowPrinter
MDIPrincipal.RptRel.ReportFileName = Caminho & "\Cobranca.RPT"
If Opt(0).value = True Then
    MDIPrincipal.RptRel.ReportFileName = Caminho & "\Cobranca.RPT"
    MDIPrincipal.RptRel.SelectionFormula = "{Dup.DtVenc} <= " & Data1 & " and {Dup.Cond} = False"
Else
    Flag = True
    If ListaEnt.Enabled = True Then
        For y = 0 To ListaEnt.ListCount - 1
            If ListaEnt.Selected(y) = True Then
                If Flag Then
                    Formula = "( {Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
                    Flag = False
                Else
                    Formula = Formula & " or  {Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
                End If
            End If
        Next
        Formula = Formula & " ) and {Dup.DtVenc} <= " & Data1 & " and {Dup.Cond} = False"
        MDIPrincipal.RptRel.SelectionFormula = Formula
    End If
End If
Imprimir
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

'Lista.ListIndex = 0
'CmbAniver.Enabled = False

Dim Componente As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next


Set Tabela = Banco.OpenRecordset("Select Ent.CodEntidade, Ent.Nome FROM Ent Order By Ent.Nome")

Tabela.MoveFirst
ListaEnt.Clear

TxtDias = 10

'While Not tabela.EOF
'    ListaEnt.AddItem psCompString(1, tabela("CodEntidade"), 6) & " - " & tabela("Nome")
'    tabela.MoveNext
'Wend

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
