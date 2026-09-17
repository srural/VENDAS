VERSION 5.00
Begin VB.Form FormSenha 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   HelpContextID   =   140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11640
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   11910
      TabIndex        =   34
      Top             =   5985
      Width           =   11940
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
      TabIndex        =   32
      Top             =   0
      Width           =   13155
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Senhas"
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
         TabIndex        =   33
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.CommandButton BtCancela 
      Caption         =   "Ca&ncela"
      Height          =   330
      Left            =   9465
      TabIndex        =   31
      Top             =   900
      Width           =   780
   End
   Begin VB.CommandButton Bt2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10755
      TabIndex        =   30
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton Bt1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10395
      TabIndex        =   29
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton BtAltera 
      Caption         =   "&Altera"
      Height          =   330
      Left            =   7725
      TabIndex        =   28
      Top             =   900
      Width           =   690
   End
   Begin VB.CommandButton BtExclui 
      Caption         =   "&Exclui"
      Height          =   330
      Left            =   7035
      TabIndex        =   27
      Top             =   900
      Width           =   690
   End
   Begin VB.CommandButton BtNovo 
      Caption         =   "&Novo"
      Height          =   330
      Left            =   6345
      TabIndex        =   26
      Top             =   900
      Width           =   690
   End
   Begin VB.CommandButton BtGrava 
      Caption         =   "&Grava"
      Height          =   330
      Left            =   8685
      TabIndex        =   25
      Top             =   900
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu"
      Height          =   4440
      Left            =   180
      TabIndex        =   8
      Top             =   1350
      Width           =   11235
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   7
         Left            =   9855
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   660
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   6
         Left            =   8469
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   675
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   5
         Left            =   7085
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   663
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   4
         Left            =   5701
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   666
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   3
         Left            =   4317
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   669
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   2
         Left            =   2933
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   672
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   1
         Left            =   1549
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   675
         Width           =   1300
      End
      Begin VB.ListBox LstMenu 
         Appearance      =   0  'Flat
         Height          =   3630
         Index           =   0
         Left            =   165
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   678
         Width           =   1300
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Utilitários"
         Height          =   240
         Index           =   7
         Left            =   9855
         TabIndex        =   16
         Top             =   360
         Width           =   960
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Relatorio"
         Height          =   240
         Index           =   6
         Left            =   8469
         TabIndex        =   15
         Top             =   360
         Width           =   1005
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Caixa"
         Height          =   240
         Index           =   5
         Left            =   7085
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Contas(Pgto)"
         Height          =   240
         Index           =   4
         Left            =   5701
         TabIndex        =   13
         Top             =   360
         Width           =   1230
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Contas (Rec)"
         Height          =   240
         Index           =   3
         Left            =   4317
         TabIndex        =   12
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Pedidos"
         Height          =   240
         Index           =   2
         Left            =   2933
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Produtos"
         Height          =   240
         Index           =   1
         Left            =   1549
         TabIndex        =   10
         Top             =   360
         Width           =   990
      End
      Begin VB.CheckBox ChkMenu 
         Caption         =   "Cadastro"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.TextBox TxtConfirma 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   5130
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   900
      Width           =   1140
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3870
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   900
      Width           =   1140
   End
   Begin VB.TextBox TxtNome 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   1
      Top             =   900
      Width           =   2445
   End
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Con&firma"
      Height          =   195
      Left            =   5130
      TabIndex        =   7
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha"
      Height          =   195
      Left            =   3870
      TabIndex        =   6
      Top             =   675
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome"
      Height          =   195
      Left            =   1350
      TabIndex        =   5
      Top             =   675
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Código"
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   675
      Width           =   495
   End
End
Attribute VB_Name = "FormSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Senha As Recordset

Private Sub Limpa()
Dim x As Integer
TxtCodigo = ""
TxtNome = ""
TxtSenha = ""
TxtConfirma = ""
For x = 0 To 7
    ChkMenu(x).value = 0
    ChkMenu_Click x
Next
End Sub


Private Sub Bt1_Click()
If Not Senha.Bof Then
    Senha.MovePrevious
    AtuMenu
End If
End Sub

Private Sub Bt2_Click()
If Not Senha.Eof Then
    Senha.MoveNext
    AtuMenu
End If
End Sub

Private Sub BtAltera_Click()
If TxtNome <> "" Then
    TxtNome.Enabled = True
    TxtSenha.Enabled = True
    BtNovo.Enabled = False
    BtExclui.Enabled = False
    BtAltera.Enabled = False
    Bt1.Enabled = False
    Bt2.Enabled = False
    Senha.Edit
End If
End Sub

Private Sub BtCancela_Click()
    BtNovo.Enabled = True
    BtAltera.Enabled = True
    BtExclui.Enabled = True
    Bt1.Enabled = True
    Bt2.Enabled = True
    TxtNome.Enabled = False
    TxtSenha.Enabled = False

End Sub

Private Sub BtExclui_Click()
If MsgBox("Confirma Exclusão da Senha ? ", vbYesNo, App.Title) = vbYes Then
    Senha.Delete
    Limpa
End If
End Sub

Private Sub BtGrava_Click()
If TxtSenha = TxtConfirma Then
'    Senha.Edit
    Senha(1) = TxtNome
    Senha(2) = TxtSenha
    Senha(3) = MontaMenu(3)
    Senha(4) = MontaMenu(4)
    Senha(5) = MontaMenu(5)
    Senha(6) = MontaMenu(6)
    Senha(7) = MontaMenu(7)
    Senha(8) = MontaMenu(8)
    Senha(9) = MontaMenu(9)
    Senha(10) = MontaMenu(10)
    Senha.Update
    If Senha.Eof Then
        Senha.MovePrevious
    End If
    BtNovo.Enabled = True
    BtAltera.Enabled = True
    BtExclui.Enabled = True
    Bt1.Enabled = True
    Bt2.Enabled = True
    AtuMenu

Else
    MsgBox "Senha não confirma !", vbInformation, App.Title
End If
End Sub

Function MontaMenu(Indice As Integer)
Dim Cont As String
Dim x As Integer
Cont = ChkMenu(Indice - 3).value
For x = 0 To LstMenu(Indice - 3).ListCount - 1
    If LstMenu(Indice - 3).Selected(x) = True Then
        Cont = Cont & "1"
    Else
        Cont = Cont & "0"
    End If
Next
MontaMenu = Cont
End Function

Private Sub BtNovo_Click()
Limpa
Senha.AddNew
BtAltera.Enabled = False
BtExclui.Enabled = False
BtNovo.Enabled = False
Bt1.Enabled = False
Bt2.Enabled = False
TxtNome.Enabled = True
TxtSenha.Enabled = True
TxtNome.SetFocus

End Sub

Private Sub ChkMenu_Click(Index As Integer)
    For x = 0 To LstMenu(Index).ListCount - 1
        If ChkMenu(Index).value = 1 Then
            LstMenu(Index).Selected(x) = True
        Else
            LstMenu(Index).Selected(x) = False
        End If
    Next
End Sub

Private Sub Form_Activate()
AtuMenu
End Sub

Private Sub Form_Load()
Dim Componentes As Variant


Me.Left = 200
Me.Top = 50

Set Senha = Banco.OpenRecordset("Select * From Pwd Order By Codigo")

End Sub

Private Sub AtuMenu()
Dim x As Integer
Dim y As Integer

TxtNome.Enabled = False
TxtSenha.Enabled = False

If Senha.RecordCount > 0 And Not Senha.Eof And Not Senha.Bof Then
'    Senha.Requery
'    Senha.MoveFirst
    TxtCodigo = Senha(0)
    TxtNome = Senha(1)
    TxtSenha = Senha(2)

    For x = 0 To 7
        If Left(Senha(3 + x), 1) = "1" Then
            ChkMenu(x).value = 1
        Else
            ChkMenu(x).value = 0
        End If
    Next
    
    LstMenu(0).Clear
    For y = 0 To MDIPrincipal.MnuCad.Count - 1
        LstMenu(0).AddItem Replace(MDIPrincipal.MnuCad(y).Caption, "&", "")
        If Mid(Senha(3), y + 2, 1) = 1 Then
            LstMenu(0).Selected(y) = True
        Else
            LstMenu(0).Selected(y) = False
        End If
    Next
    
    LstMenu(1).Clear
    For y = 0 To MDIPrincipal.MnuPrecos.Count - 1
    
    LstMenu(1).AddItem Replace(MDIPrincipal.MnuPrecos(y).Caption, "&", "")
    If Mid(Senha(4), y + 2, 1) = 1 Then
         LstMenu(1).Selected(y) = True
    Else
         LstMenu(1).Selected(y) = False
    End If
    Next
    
    LstMenu(2).Clear
    For y = 0 To MDIPrincipal.MnuPed.Count - 1
       LstMenu(2).AddItem Replace(MDIPrincipal.MnuPed(y).Caption, "&", "")
       If Mid(Senha(5), y + 2, 1) = 1 Then
            LstMenu(2).Selected(y) = True
       Else
            LstMenu(2).Selected(y) = False
       End If
    Next
    
    LstMenu(3).Clear
    For y = 0 To MDIPrincipal.MnuPcl.Count - 1
       LstMenu(3).AddItem Replace(MDIPrincipal.MnuPcl(y).Caption, "&", "")
       If Mid(Senha(6), y + 2, 1) = 1 Then
            LstMenu(3).Selected(y) = True
       Else
            LstMenu(3).Selected(y) = False
       End If
    Next
    
    LstMenu(4).Clear
    For y = 0 To MDIPrincipal.MnuPag.Count - 1
       LstMenu(4).AddItem Replace(MDIPrincipal.MnuPag(y).Caption, "&", "")
       If Mid(Senha(7), y + 2, 1) = 1 Then
            LstMenu(4).Selected(y) = True
       Else
            LstMenu(4).Selected(y) = False
       End If
    Next
    
    LstMenu(5).Clear
    For y = 0 To MDIPrincipal.MnuCxa.Count - 1
       LstMenu(5).AddItem Replace(MDIPrincipal.MnuCxa(y).Caption, "&", "")
       If Mid(Senha(8), y + 2, 1) = 1 Then
            LstMenu(5).Selected(y) = True
       Else
            LstMenu(5).Selected(y) = False
       End If
    Next
    
    LstMenu(6).Clear
    For y = 0 To MDIPrincipal.MnuRel.Count - 1
       LstMenu(6).AddItem Replace(MDIPrincipal.MnuRel(y).Caption, "&", "")
       If Mid(Senha(9), y + 2, 1) = 1 Then
            LstMenu(6).Selected(y) = True
       Else
            LstMenu(6).Selected(y) = False
       End If
    Next
    
    LstMenu(7).Clear
    For y = 0 To MDIPrincipal.MnuUtil.Count - 1
       LstMenu(7).AddItem Replace(MDIPrincipal.MnuUtil(y).Caption, "&", "")
       If Mid(Senha(10), y + 2, 1) = 1 Then
            LstMenu(7).Selected(y) = True
       Else
            LstMenu(7).Selected(y) = False
       End If
    Next
        
End If

End Sub
