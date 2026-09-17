VERSION 5.00
Begin VB.Form FormLocalizar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   7515
   ClientTop       =   5280
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   9
      Top             =   3420
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
      TabIndex        =   7
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar"
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
         TabIndex        =   8
         Top             =   45
         Width           =   2445
      End
   End
   Begin VB.ListBox LstOperadores 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FormLocalizar.frx":0000
      Left            =   1980
      List            =   "FormLocalizar.frx":0019
      TabIndex        =   2
      Top             =   1395
      Width           =   975
   End
   Begin VB.CommandButton CmdLocalizar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Index           =   2
      Left            =   3180
      TabIndex        =   5
      Top             =   2655
      Width           =   1215
   End
   Begin VB.CommandButton CmdLocalizar 
      Caption         =   "&P&roximo"
      Height          =   375
      Index           =   1
      Left            =   3180
      TabIndex        =   4
      Top             =   1875
      Width           =   1215
   End
   Begin VB.CommandButton CmdLocalizar 
      Caption         =   "&Primeiro"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3180
      TabIndex        =   3
      Top             =   1395
      Width           =   1215
   End
   Begin VB.ListBox LstCampos 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FormLocalizar.frx":0061
      Left            =   360
      List            =   "FormLocalizar.frx":0063
      TabIndex        =   1
      Top             =   1395
      Width           =   1455
   End
   Begin VB.TextBox TxtCriterio 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   855
      Width           =   4035
   End
   Begin VB.Label LblValor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "FormLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdLocalizar_Click(Index As Integer)
Dim Valor As String
Dim Criterio As String
If Index = 2 Then ' Botão fechar
    Unload Me
Else
    With Formulário
        Select Case .Registros(LstCampos.Text).Type
            Case dbText, dbChar, dbMemo
                Valor = """" & TxtCriterio
                If LstOperadores = "Contendo" Then
                    Valor = """*" & TxtCriterio & "*"""
                Else
                    Valor = Valor + """"
                End If
                
            Case dbDate
                If IsDate(TxtCriterio) Then
                    If LstOperadores.Text = "Contendo" Then
                        MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                        Exit Sub
                    End If
                    Valor = "#" & Format(TxtCriterio, "dd/mm/yyyy") & "#"
                Else
                    MsgBox "Valor digitado não é uma data válida", vbCritical, App.Title
                    Exit Sub
                End If
            Case dbBoolean
                If LstOperadores.Text = "Contendo" Then
                    MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                    Exit Sub
                End If
                If UCase(TxtCriterio) = "ATIVO" Or UCase(TxtCriterio) = "SIM" Or UCase(TxtCriterio) = "VERDADEIRO" Then
                    Valor = "true"
                ElseIf UCase(TxtCriterio) = "INATIVO" Or UCase(TxtCriterio) = "NÃO" Or UCase(TxtCriterio) = "FALSO" Then
                    Valor = "False"
                Else
                    MsgBox "Valores Válidos são (Sim/Não), (Ativo/Inativo), (Verdadeiro/Falso)", vbCritical, App.Title
                    Exit Sub
                End If
            Case Else
                If IsNumeric(TxtCriterio) Then
                    If LstOperadores.Text = "Contendo" Then
                        MsgBox "Operador Contendo só pode ser usado com campos do tipo texto", vbCritical, App.Title
                        Exit Sub
                    End If
                    Valor = Numero2(Format(TxtCriterio, "0.00"))
                Else
                    MsgBox "Valor não numérico", vbCritical, App.Title
                    Exit Sub
                End If
            End Select
            Criterio = "[" & LstCampos.Text & "]" & Operadores(LstOperadores.ListIndex) & Valor
            If Index = 0 Then ' Botão primeiro
                .Registros.FindFirst Criterio
            Else
                .Registros.FindNext Criterio
            End If
            If .Registros.NoMatch = True Then
                MsgBox "Registro não Localizado", vbExclamation, App.Title
                CmdLocalizar(0).Default = True
            Else
                MostraDados
                CmdLocalizar(1).Default = True
            End If
    End With
End If
End Sub

Private Sub Form_Load()
ListaCampos LstCampos
LstCampos.ListIndex = 1
LstOperadores.ListIndex = 6
End Sub

