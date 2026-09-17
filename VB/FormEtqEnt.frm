VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormEtqEnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6075
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FormEtqEnt.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   10440
   Begin VB.CommandButton BtLista2 
      Caption         =   "Monta Lista"
      Height          =   315
      Left            =   8595
      TabIndex        =   32
      Top             =   5040
      Width           =   1545
   End
   Begin VB.Data DatVend 
      Caption         =   "Vendedores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7"
      Top             =   5310
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   29
      Top             =   5715
      Width           =   11535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Desmarcar"
      Height          =   330
      Left            =   8595
      TabIndex        =   28
      Top             =   4545
      Width           =   1545
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10455
      TabIndex        =   26
      Top             =   0
      Width           =   10455
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Etiquetas de Entidades"
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
         TabIndex        =   27
         Top             =   45
         Width           =   5010
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Cartão"
      Height          =   375
      Index           =   2
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprimir"
      Top             =   4275
      Width           =   1890
   End
   Begin VB.Frame FrmEnd 
      Caption         =   "Endereços"
      Height          =   585
      Left            =   195
      TabIndex        =   21
      Top             =   4815
      Width           =   4155
      Begin VB.OptionButton OptEnd 
         Caption         =   "Entrega"
         Height          =   285
         Index           =   2
         Left            =   2910
         TabIndex        =   24
         Top             =   210
         Width           =   1005
      End
      Begin VB.OptionButton OptEnd 
         Caption         =   "Cobrança"
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   23
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton OptEnd 
         Caption         =   "Cadastro"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   210
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.CommandButton BtLista 
      Caption         =   "Monta Lista"
      Height          =   315
      Left            =   210
      TabIndex        =   20
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pessoa"
      Height          =   1320
      Left            =   2475
      TabIndex        =   16
      Top             =   3195
      Width           =   1815
      Begin VB.OptionButton OptTipo 
         Caption         =   "Jurídica"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   945
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Física"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   607
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Todos"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Etiqueta &Pequena"
      Height          =   375
      Index           =   1
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir"
      Top             =   3435
      Width           =   1890
   End
   Begin VB.Data DatDados 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Ent Where Tipo = 1 or Tipo = 5 Or Tipo = 6 Or Tipo = 7 Order by Nome"
      Top             =   2880
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
      Height          =   3270
      Left            =   4500
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   1170
      Width           =   5640
   End
   Begin VB.ComboBox CmbAniver 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormEtqEnt.frx":0342
      Left            =   210
      List            =   "FormEtqEnt.frx":036A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2670
      Width           =   1905
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Etiqueta &Grande"
      Height          =   375
      Index           =   0
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   3840
      Width           =   1890
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "FormEtqEnt.frx":03D3
      Left            =   210
      List            =   "FormEtqEnt.frx":03EC
      TabIndex        =   6
      Top             =   855
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   2490
      Left            =   2475
      TabIndex        =   0
      Top             =   630
      Width           =   1815
      Begin VB.OptionButton Opt 
         Caption         =   "Carnê"
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   11
         Top             =   2025
         Width           =   825
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Diversos"
         Height          =   375
         Index           =   4
         Left            =   180
         TabIndex        =   10
         Top             =   1656
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Aniversariantes"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   1287
         Width           =   1500
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Inativos"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   918
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ativos"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   549
         Width           =   825
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
      Bindings        =   "FormEtqEnt.frx":0471
      Height          =   315
      Index           =   0
      Left            =   4500
      TabIndex        =   13
      Top             =   810
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo DbcVend 
      Bindings        =   "FormEtqEnt.frx":048B
      Height          =   315
      Left            =   4500
      TabIndex        =   30
      Top             =   5040
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      Height          =   285
      Left            =   4500
      TabIndex        =   31
      Top             =   4815
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa Nome"
      Height          =   195
      Left            =   4500
      TabIndex        =   14
      Top             =   585
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aniversariante (Mês)"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   2445
      Width           =   1440
   End
   Begin VB.Label LblDados 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Index           =   56
      Left            =   210
      TabIndex        =   4
      Top             =   615
      Width           =   315
   End
End
Attribute VB_Name = "FormEtqEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabela As Recordset
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub BtLista_Click()
Dim VarAniver As Recordset

If CmbAniver <> "" Then
    Set VarAniver = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, ENT.DtNasc FROM ENT WHERE month(ENT.DtNasc) = " & CmbAniver.ListIndex + 1 & " ORDER BY ENT.Nome")
    VarAniver.Requery
    If VarAniver.RecordCount > 0 Then
        ListaEnt.Clear
        While Not VarAniver.Eof
            ListaEnt.AddItem psCompString(1, VarAniver!CodEntidade, 6) & " - " & VarAniver!Nome
            ListaEnt.Selected(ListaEnt.ListCount - 1) = True
            VarAniver.MoveNext
        Wend
        Opt(4).value = True
        Opt_Click 4
    End If
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtLista2_Click()
Dim VarVend As Recordset

If DbcVend <> "" Then
    Set VarVend = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, ENT.Vendedor FROM ENT WHERE Ent.Vendedor = " & DbcVend.BoundText & " ORDER BY ENT.Nome")
    VarVend.Requery
    If VarVend.RecordCount > 0 Then
        ListaEnt.Clear
        While Not VarVend.Eof
            ListaEnt.AddItem psCompString(1, VarVend!CodEntidade, 6) & " - " & VarVend!Nome
            ListaEnt.Selected(ListaEnt.ListCount - 1) = True
            VarVend.MoveNext
        Wend
        Opt(4).value = True
        Opt_Click 4
    End If
End If

End Sub

Private Sub CmdImprimir_Click(Index As Integer)
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula  As String
Dim y As Integer
Dim flag As Boolean

'MDIPrincipal.ComRel.ShowPrinter

Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"


Select Case Lista.ListIndex
    Case 0
        Formula = "{Ent.Tipo} = 1"
    Case 1
        Formula = "{Ent.Tipo} = 2"
    Case 2
        Formula = "{Ent.Tipo} = 3"
    Case 3
        Formula = "{Ent.Tipo} = 4"
    Case 4
        Formula = "{Ent.Tipo} = 5"
    Case 5
        Formula = "{Ent.Tipo} = 6"
    Case 6
        Formula = "{Ent.Tipo} = 7"
End Select

If Opt(3).value = True Then
    MDIPrincipal.RptRel.ReportFileName = Caminho & "\RelAni.RPT"
Else
    If Index = 0 Then
        If OptEnd(0).value = True Then
            MDIPrincipal.RptRel.ReportFileName = Caminho & "\MalaEnt.RPT"
        ElseIf OptEnd(1).value = True Then
                MDIPrincipal.RptRel.ReportFileName = Caminho & "\MalaEntCob.RPT"
                Formula = Formula + " and {Ent.Endereço_Cob} <> ' ' "
            Else
                MDIPrincipal.RptRel.ReportFileName = Caminho & "\MalaEntEntrega.RPT"
                Formula = Formula + " and {Ent.Endereço_Entrega} <> ' ' "
                
        End If
    ElseIf Index = 1 Then
        If OptEnd(0).value = True Then
            MDIPrincipal.RptRel.ReportFileName = Caminho & "\EtqEntPeq.RPT"
        ElseIf OptEnd(1).value = True Then
                MDIPrincipal.RptRel.ReportFileName = Caminho & "\EtqEntPeqCob.RPT"
                Formula = Formula + " and {Ent.Endereço_Cob} <> ' ' "
            Else
                MDIPrincipal.RptRel.ReportFileName = Caminho & "\EtqEntPeqEntrega.RPT"
                Formula = Formula + " and {Ent.Endereço_Entrega} <> ' ' "
        End If
    Else
        MDIPrincipal.RptRel.ReportFileName = Caminho & "\Cartao.RPT"
        
    End If
End If


If OptTipo(1).value = True Then
        Formula = Formula + " and {Ent.Classificacao} = 1 "
    ElseIf OptTipo(2).value = True Then
            Formula = Formula + " and {Ent.Classificacao} = 2 "
End If


If Opt(3).value = True Then
    Formula = Formula + " and Month({Ent.DtNasc}) = " & CmbAniver.ListIndex + 1
Else
    If Opt(1).value = True Then
            Formula = Formula + " and {Ent.Ativo} = true "
        ElseIf Opt(2).value = True Then
            Formula = Formula + " and {Ent.Ativo} = false "
    End If
End If

flag = True
If ListaEnt.Enabled = True Then
    For y = 0 To ListaEnt.ListCount - 1
        If ListaEnt.Selected(y) = True Then
            If flag Then
                Formula = "{Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
                flag = False
            Else
                Formula = Formula & " or  {Ent.CodEntidade} = " & CLng(Left(ListaEnt.List(y), 6))
            End If
        End If
    Next
End If

MDIPrincipal.RptRel.SelectionFormula = Formula

'MDIPrincipal.RptRel.Formulas(0) = "Formulas = '" & Periodo & "' "
Imprimir
End Sub


Private Sub Command1_Click()
Dim y As Integer
    For y = 0 To ListaEnt.ListCount - 1
         ListaEnt.Selected(y) = False
    Next
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

Lista.ListIndex = 0
CmbAniver.Enabled = False

Dim Componente As Variant
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

DatVend.DatabaseName = Caminho & "\Dados.Mdb"


Set Tabela = Banco.OpenRecordset("Select Ent.CodEntidade, Ent.Nome FROM Ent Order By Ent.Nome")

Tabela.MoveFirst
ListaEnt.Clear

'While Not tabela.EOF
'    ListaEnt.AddItem psCompString(1, tabela("CodEntidade"), 6) & " - " & tabela("Nome")
'    tabela.MoveNext
'Wend

ListaEnt.Enabled = False
DbcDados(0).Enabled = False

End Sub

Private Sub Opt_Click(Index As Integer)

If Index = 3 Then
    If Opt(3).value = True Then
        CmbAniver.Enabled = True
    End If
End If

If Index >= 4 Then
    DbcDados(0).Enabled = True
    ListaEnt.Enabled = True
    DbcDados(0).SetFocus
Else
    DbcDados(0).Enabled = False
    ListaEnt.Enabled = False
End If

End Sub
