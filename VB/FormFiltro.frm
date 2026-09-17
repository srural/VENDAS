VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormFiltro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8040
   ClientLeft      =   2280
   ClientTop       =   615
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8790
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -765
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   29
      Top             =   7695
      Width           =   11535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8790
      TabIndex        =   27
      Top             =   0
      Width           =   8790
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
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
         TabIndex        =   28
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.Frame FrmCriterio 
      Caption         =   "3º Critério"
      Height          =   975
      Index           =   2
      Left            =   180
      TabIndex        =   17
      Top             =   2880
      Width           =   6975
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         ItemData        =   "FormFiltro.frx":0000
         Left            =   5820
         List            =   "FormFiltro.frx":0019
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtValor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Width           =   2355
      End
      Begin VB.ComboBox CmbCampo 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton OptE 
         Caption         =   "E"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton OptOu 
         Caption         =   "Ou"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   585
         Width           =   555
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Index           =   2
         Left            =   3210
         TabIndex        =   23
         Top             =   240
         Width           =   360
      End
      Begin VB.Label LblCampo 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Data DatFiltro 
      Caption         =   "Lista"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5940
      Visible         =   0   'False
      Width           =   2715
   End
   Begin MSDBGrid.DBGrid DBGFiltro 
      Bindings        =   "FormFiltro.frx":0061
      Height          =   3195
      Left            =   180
      OleObjectBlob   =   "FormFiltro.frx":0086
      TabIndex        =   16
      Top             =   4080
      Width           =   8355
   End
   Begin VB.CommandButton CmdFiltrar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   11
      Top             =   3300
      Width           =   1155
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "&Reiniciar"
      Height          =   495
      Index           =   2
      Left            =   7440
      TabIndex        =   10
      Top             =   2460
      Width           =   1155
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Fil&trar"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   9
      Top             =   1620
      Width           =   1155
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "&Visualizar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   8
      Top             =   900
      Width           =   1155
   End
   Begin VB.Frame FrmCriterio 
      Caption         =   "2º Critério"
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1740
      Width           =   6975
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "FormFiltro.frx":0A6B
         Left            =   5820
         List            =   "FormFiltro.frx":0A84
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptOu 
         Caption         =   "Ou"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   630
         Width           =   555
      End
      Begin VB.OptionButton OptE 
         Caption         =   "E"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.ComboBox CmbCampo 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtValor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label LblCampo 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   14
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame FrmCriterio 
      Caption         =   "1º Critério"
      Height          =   915
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "0"
      Top             =   660
      Width           =   6975
      Begin VB.ComboBox CmbOperador 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "FormFiltro.frx":0ACC
         Left            =   5820
         List            =   "FormFiltro.frx":0AE5
         TabIndex        =   24
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox TxtValor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3180
         TabIndex        =   2
         Top             =   420
         Width           =   2415
      End
      Begin VB.ComboBox CmbCampo 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "FormFiltro.frx":0B2D
         Left            =   900
         List            =   "FormFiltro.frx":0B2F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1995
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Index           =   0
         Left            =   3150
         TabIndex        =   13
         Top             =   180
         Width           =   360
      End
      Begin VB.Label LblCampo 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   12
         Top             =   180
         Width           =   495
      End
   End
End
Attribute VB_Name = "FormFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbCampo_Click(Index As Integer)
Dim cmb As Variant

If CmbCampo(Index).ListIndex > 0 Then
    TxtValor(Index).Enabled = True
    CmbOperador(Index).Enabled = True
    CmbOperador(Index).ListIndex = 0
    CmdFiltrar(0).Enabled = True
    CmdFiltrar(1).Enabled = True
Else
    TxtValor(Index).Text = ""
    TxtValor(Index).Enabled = False
    For Each cmb In CmbCampo
        If cmb.ListIndex > 0 Then
            Exit Sub
        End If
    Next
    CmbOperador(Index).Enabled = False
    CmbOperador(Index).ListIndex = -1
    CmdFiltrar(0).Enabled = False
    CmdFiltrar(1).Enabled = False
End If
End Sub

Private Sub CmdFiltrar_Click(Index As Integer)
Dim Filtro As String
Dim cmb As Variant
Dim Criterio As String
Dim Operador As String
Dim Delimitador As String
Dim Valor As String
With Formulário
    If Index <= 2 Then ' Quando for os comandos que modificam o sql (visualizar, filtrar e reiniciar)
        If Index = 0 Then  'Botão visualizar
            SQL = .SelectFromFiltro
        Else
            SQL = .SelectFromSQL
        End If
        If .WhereSQL <> "" Then
            Filtro = " WHERE (" & .WhereSQL
        End If
        If Index < 2 Then
            For Each cmb In CmbCampo
                If cmb.ListIndex > 0 Then
                    Criterio = GeraCriterio(cmb.Index)
                    If Criterio <> "" Then
                        If Filtro <> "" Then
                           If cmb.Index = 0 Or OptE(cmb.Index - 1).value = True Then
                              Filtro = Filtro & " AND " & Criterio
                           Else
                              Filtro = Filtro & ") OR (" & Criterio
                           End If
                        Else
                            Filtro = " WHERE (" & Criterio
                        End If
                    End If
                End If
            Next
        End If
        If Filtro <> "" Then Filtro = Filtro & ")"
        SQL = SQL & Filtro
        If .OrderBySQL <> "" Then SQL = SQL & " ORDER BY " & .OrderBySQL
        If Index = 0 Then
            DatFiltro.RecordSource = SQL
            DatFiltro.Refresh
        Else
            Set .Registros = Banco.OpenRecordset(SQL)
            Inicio
            If Index = 1 Then
                .FiltroSQL = Mid(Filtro, 8)
            Else
                .FiltroSQL = ""
            End If
        End If
    End If
    If Index > 0 Then Unload Me
End With
End Sub

Private Sub Form_Load()
Dim Combo As Integer
DatFiltro.DatabaseName = Caminho & "\Dados.mdb"
With Formulário
    For Combo = 0 To CmbCampo.Count - 1
        ListaCampos CmbCampo(Combo)
        CmbCampo(Combo).AddItem "(Nenhum)"
        CmbCampo(Combo).ListIndex = 0
    Next
    SQL = .SelectFromFiltro
    If .FiltroSQL <> "" Then
        SQL = SQL & " WHERE " & .FiltroSQL
    ElseIf .WhereSQL <> "" Then
        SQL = SQL & " WHERE " & .WhereSQL
    End If
    If .OrderBySQL <> "" Then
        SQL = SQL & " ORDER BY " & .OrderBySQL
    End If
    DatFiltro.RecordSource = SQL
    DatFiltro.Refresh
End With
End Sub


Private Function GeraCriterio(cmb As Integer) As String
Dim Valor As String
With Formulário
    Select Case .Registros(CmbCampo(cmb)).Type
        Case dbText, dbChar, dbMemo
            Valor = """" & TxtValor(cmb)
            If CmbOperador(cmb) = "Contendo" Then
                Valor = """*" & TxtValor(cmb) & "*"""
            Else
                Valor = Valor + """"
            End If
            
        Case dbDate
            If IsDate(TxtValor(cmb)) Then
                If CmbOperador(cmb).Text = "Contendo" Then
                    MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                    Exit Function
                End If
                Valor = "#" & Format(TxtValor(cmb), "dd/mm/yyyy") & "#"
            Else
                MsgBox "Valor digitado não é uma data válida", vbCritical, App.Title
                Exit Function
            End If
        Case dbBoolean
            If CmbOperador(cmb).Text = "Contendo" Then
                MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                Exit Function
            End If
            If UCase(TxtValor(cmb)) = "ATIVO" Or UCase(TxtValor(cmb)) = "SIM" Or UCase(TxtValor(cmb)) = "VERDADEIRO" Then
                Valor = "true"
            ElseIf UCase(TxtValor(cmb)) = "INATIVO" Or UCase(TxtValor(cmb)) = "NÃO" Or UCase(TxtValor(cmb)) = "FALSO" Then
                Valor = "False"
            Else
                MsgBox "Valores Válidos são (Sim/Não), (Ativo/Inativo), (Verdadeiro/Falso)", vbCritical, App.Title
                Exit Function
            End If
        Case Else
            If IsNumeric(TxtValor(cmb)) Then
                If CmbOperador(cmb).Text = "Contendo" Then
                    MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                    Exit Function
                End If
                Valor = Numero2(Format(TxtValor(cmb), "0.00"))
            Else
                MsgBox "Valor não numérico", vbCritical, App.Title
                Exit Function
            End If
        End Select
        GeraCriterio = "[" & CmbCampo(cmb).Text & "]" & Operadores(CmbOperador(cmb).ListIndex) & Valor
End With
End Function
