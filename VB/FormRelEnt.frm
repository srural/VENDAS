VERSION 5.00
Begin VB.Form FormRelEnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5190
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8505
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   31
      Top             =   4815
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   29
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios de Entidades"
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
         Left            =   90
         TabIndex        =   30
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.TextBox TxtDia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2295
      TabIndex        =   25
      Top             =   3645
      Width           =   735
   End
   Begin VB.CommandButton ChkCarta 
      Caption         =   "&Carta"
      Height          =   330
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Imprimir"
      Top             =   3645
      Width           =   1320
   End
   Begin VB.ComboBox CmbAniver 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormRelEnt.frx":0000
      Left            =   225
      List            =   "FormRelEnt.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3630
      Width           =   1980
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   6510
      TabIndex        =   18
      Top             =   3750
      Width           =   1620
   End
   Begin VB.TextBox TxtFiltro 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   225
      TabIndex        =   17
      Top             =   4215
      Width           =   5805
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Filtrado por"
      Height          =   2235
      Left            =   4455
      TabIndex        =   12
      Top             =   945
      Width           =   2070
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Dia Bom Para"
         Height          =   330
         Index           =   5
         Left            =   195
         TabIndex        =   28
         Tag             =   "Cidade"
         Top             =   1845
         Width           =   1515
      End
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Novos Mês"
         Height          =   330
         Index           =   4
         Left            =   195
         TabIndex        =   27
         Tag             =   "Cidade"
         Top             =   1542
         Width           =   1110
      End
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Profissão"
         Height          =   330
         Index           =   3
         Left            =   195
         TabIndex        =   21
         Tag             =   "Cidade"
         Top             =   1239
         Width           =   1110
      End
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Uf (estado)"
         Height          =   330
         Index           =   2
         Left            =   195
         TabIndex        =   15
         Tag             =   "Cidade"
         Top             =   936
         Width           =   1110
      End
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Cidade"
         Height          =   375
         Index           =   1
         Left            =   195
         TabIndex        =   14
         Tag             =   "Nome"
         Top             =   588
         Width           =   885
      End
      Begin VB.OptionButton OptFiltro 
         Caption         =   "Inicial Nome"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Tag             =   "CodEntidade"
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opção"
      Height          =   2190
      Left            =   2475
      TabIndex        =   7
      Top             =   945
      Width           =   1815
      Begin VB.OptionButton Opt 
         Caption         =   "&Bloqueado"
         Height          =   465
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   1665
         Width           =   1500
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ativos"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   552
         Width           =   825
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Inativos"
         Height          =   375
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   924
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Aniversariantes"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1296
         Width           =   1500
      End
   End
   Begin VB.ListBox Lista 
      Height          =   2010
      ItemData        =   "FormRelEnt.frx":0091
      Left            =   225
      List            =   "FormRelEnt.frx":00AD
      TabIndex        =   6
      Top             =   1020
      Width           =   2130
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenado por"
      Height          =   2235
      Left            =   6690
      TabIndex        =   2
      Top             =   945
      Width           =   1380
      Begin VB.OptionButton OptOrdem 
         Caption         =   "DtNasc"
         Height          =   375
         Index           =   3
         Left            =   150
         TabIndex        =   22
         Tag             =   "day(DtNasc)"
         Top             =   1725
         Width           =   885
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Código"
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Tag             =   "CodEntidade"
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Nome"
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Tag             =   "Nome"
         Top             =   735
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Cidade"
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Tag             =   "Cidade"
         Top             =   1230
         Width           =   885
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir"
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dia"
      Height          =   240
      Left            =   2295
      TabIndex        =   26
      Top             =   3375
      Width           =   600
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aniversariante (Mês)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   225
      TabIndex        =   20
      Top             =   3375
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Descrição do Filtro"
      Height          =   195
      Left            =   225
      TabIndex        =   16
      Top             =   4005
      Width           =   1320
   End
   Begin VB.Label LblDados 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Index           =   56
      Left            =   225
      TabIndex        =   0
      Top             =   810
      Width           =   315
   End
End
Attribute VB_Name = "FormRelEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbCancel_Click()
Unload Me
End Sub

Private Sub ChkCarta_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula  As String
Dim Ordem As String
Dim Componentes As Variant
    

If CmbAniver.Text <> "" Then

    Select Case Lista.ListIndex
        Case 0
            Formula = "(Ent.Tipo = 1 OR Ent.Tipo = 8)"
        Case 1
            Formula = "Ent.Tipo = 2"
        Case 2
            Formula = "Ent.Tipo = 3"
        Case 3
            Formula = "Ent.Tipo = 4"
        Case 4
            Formula = "Ent.Tipo = 5"
        Case 5
            Formula = "Ent.Tipo = 6"
        Case 6
            Formula = "Ent.Tipo = 7"
        Case 7
            Formula = "Ent.Tipo = 8"
    
    End Select
    
    If Opt(3).value = True Then
        If TxtDia = "" Then
            Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1 & " and  Ent.Ativo = true "
        Else
            Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1 & " and Day(DtNasc) = " & TxtDia & " and Ent.Ativo = true "
        End If
    Else
        If Opt(1).value = True Then
                Formula = Formula + " and Ent.Ativo = true "
            ElseIf Opt(2).value = True Then
                Formula = Formula + " and Ent.Ativo = false "
        End If
    End If
    
    If TxtFiltro <> "" Then
        If OptFiltro(0).value = True Then
                Formula = Formula + "and Nome like " & """" & TxtFiltro & "*" & """"
            ElseIf OptFiltro(1).value = True Then
                    Formula = Formula + "and Cidade = " & """" & TxtFiltro & """"
                ElseIf OptFiltro(2).value = True Then
                        Formula = Formula + "and UF = " & """" & TxtFiltro & """"
                    ElseIf OptFiltro(3).value = True Then
                        Formula = Formula + "and Profissao = " & """" & TxtFiltro & """"
                    ElseIf OptFiltro(5).value = True Then
                        Formula = Formula + "and DiaBomPara = " & """" & TxtFiltro & """"
        
        End If
    End If
    
    For Each Componentes In OptOrdem
        If Componentes.value = True Then Ordem = Componentes.Tag
    Next
    
    Set varPesquisa(1) = Banco.OpenRecordset("Select * From Ent Where CodEntidade > 1 and " & Formula & " Order by " & Ordem)
    Set FrmRelCarta.Relatorio.Recordset = varPesquisa(1)
    FrmRelCarta.Config Caminho & "\Aniversariante.txt"
   
Else
    MsgBox "Você de deve escolher o mês !", vbInformation, App.Title
End If
End Sub


Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim Data1 As String
Dim Data2 As String
Dim Periodo As String
Dim Formula  As String
Dim Ordem As String
Dim Componentes As Variant

'MDIPrincipal.ComRel.ShowPrinter

'Data1 = "Date(" & Year(DtData1) & "," & Month(DtData1) & "," & Day(DtData1) & ")"

'If Opt(3).Value = True Then
'    MDIPrincipal.RptRel.ReportFileName = Caminho & "\relAni.RPT"
'Else
'    MDIPrincipal.RptRel.ReportFileName = Caminho & "\relENT.RPT"
'End If

Select Case Lista.ListIndex
    Case 0
        Formula = "(Ent.Tipo = 1 OR Ent.Tipo = 8)"
    Case 1
        Formula = "Ent.Tipo = 2"
    Case 2
        Formula = "Ent.Tipo = 3"
    Case 3
        Formula = "Ent.Tipo = 4"
    Case 4
        Formula = "Ent.Tipo = 5"
    Case 5
        Formula = "Ent.Tipo = 6"
    Case 6
        Formula = "Ent.Tipo = 7"
    Case 7
        Formula = "Ent.Tipo = 8"
End Select

If Opt(3).value = True Then
'    Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1 & " and  Ent.Ativo = true "
    If TxtDia = "" Then
        Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1 & " and  Ent.Ativo = true "
    Else
        Formula = Formula + " and Month(DtNasc) = " & CmbAniver.ListIndex + 1 & " and Day(DtNasc) = " & TxtDia & " and Ent.Ativo = true "
    End If

ElseIf Opt(1).value = True Then
        Formula = Formula + " and Ent.Ativo = true "
    ElseIf Opt(2).value = True Then
            Formula = Formula + " and Ent.Ativo = false "
    ElseIf Opt(4).value = True Then
            Formula = Formula + " and Ent.Ativo = true and Ent.status = 4 "
End If

If TxtFiltro <> "" Then
    If OptFiltro(0).value = True Then
            Formula = Formula + "and Nome like " & """" & TxtFiltro & "*" & """"
        ElseIf OptFiltro(1).value = True Then
                Formula = Formula + " and Cidade = " & """" & TxtFiltro & """"
            ElseIf OptFiltro(2).value = True Then
                Formula = Formula + " and UF = " & """" & TxtFiltro & """"
            ElseIf OptFiltro(3).value = True Then
                Formula = Formula + " and Profissao = " & """" & TxtFiltro & """"
            ElseIf OptFiltro(4).value = True Then
                Formula = Formula + " and month(DtCadastro) = " & """" & TxtFiltro & """" & " AND year(DtCadastro) = " & Year(Date)
            ElseIf OptFiltro(5).value = True Then
                    Formula = Formula + " and DiaBomPgto = " & TxtFiltro
    End If
End If

For Each Componentes In OptOrdem
    If Componentes.value = True Then Ordem = Componentes.Tag
Next
'Sql = "Select * From Ent Where CodEntidade > 1 and " & Formula & " Order by " & Ordem
Set varPesquisa(1) = Banco.OpenRecordset("Select * From Ent Where CodEntidade > 1 and " & Formula & " Order by " & Ordem)
Set FrmRelEnt.Relatorio.Recordset = varPesquisa(1)
FrmRelEnt.rpCabTitulo.Caption = "Relatório de " & Lista.Text
FrmRelEnt.rpCabSubTitulo.Caption = NomeEmpresa
FrmRelEnt.Config

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

Lista.ListIndex = 0
CmbAniver.Enabled = False
End Sub

Private Sub Opt_Click(Index As Integer)

If Index = 3 Then
    If Opt(3).value = True Then
        CmbAniver.Enabled = True
    End If
End If

End Sub

Private Sub OptFiltro_Click(Index As Integer)
If Index = 4 Then
    TxtFiltro = Month(Date)
End If
End Sub
