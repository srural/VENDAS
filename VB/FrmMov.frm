VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMov 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6570
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   8715
   Icon            =   "FrmMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8715
   Tag             =   "Lançamentos dos Movimentos"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   -675
      ScaleHeight     =   285
      ScaleWidth      =   14025
      TabIndex        =   39
      Top             =   6255
      Width           =   14055
      Begin VB.TextBox txtStatus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   990
         TabIndex        =   40
         Top             =   45
         Width           =   7905
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11085
      TabIndex        =   37
      Top             =   0
      Width           =   11085
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamentos dos Movimentos"
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
         TabIndex        =   38
         Top             =   45
         Width           =   6540
      End
   End
   Begin TabDlg.SSTab TabCad 
      Height          =   5190
      Left            =   90
      TabIndex        =   11
      Top             =   705
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   9155
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "P&esquisa (Ctrl + E)"
      TabPicture(0)   =   "FrmMov.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "M&ovimento  (Ctrl + O)"
      TabPicture(1)   =   "FrmMov.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "BtAcoes(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "BtAcoes(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ManterDbcDados(10)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CheckBox ManterDbcDados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check1"
         Height          =   225
         Index           =   10
         Left            =   4920
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Frame Frame6 
         Height          =   45
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   8085
      End
      Begin VB.CommandButton BtAcoes 
         Caption         =   "&Cancel. (F6)"
         Enabled         =   0   'False
         Height          =   765
         Index           =   3
         Left            =   7005
         Picture         =   "FrmMov.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3990
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         Height          =   825
         Left            =   150
         TabIndex        =   27
         Top             =   3900
         Width           =   3075
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Selecione a opção ao lado dos controles para manter os valores nos próximos lançamentos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   555
            Left            =   180
            TabIndex        =   28
            Top             =   180
            Width           =   2745
         End
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   120
         TabIndex        =   26
         Top             =   3630
         Width           =   8085
      End
      Begin VB.CommandButton BtAcoes 
         Caption         =   "&Salvar (F5)"
         Enabled         =   0   'False
         Height          =   765
         Index           =   2
         Left            =   6060
         Picture         =   "FrmMov.frx":0C94
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3990
         Width           =   950
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2355
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   960
         Width           =   7965
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   9
            Left            =   2490
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   3
            Left            =   7650
            TabIndex        =   34
            Top             =   450
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   2
            Left            =   2490
            TabIndex        =   15
            Top             =   1620
            Width           =   225
         End
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   1
            Left            =   2490
            TabIndex        =   14
            Top             =   1245
            Width           =   225
         End
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   0
            Left            =   2490
            TabIndex        =   13
            Top             =   855
            Width           =   225
         End
         Begin VB.CheckBox ManterDbcDados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check1"
            Height          =   225
            Index           =   7
            Left            =   2490
            MaskColor       =   &H00000000&
            TabIndex        =   12
            Top             =   90
            Width           =   225
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Index           =   1
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1950
            Width           =   2025
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Index           =   0
            Left            =   1710
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1950
            Width           =   3315
         End
         Begin MSMask.MaskEdBox MebDados 
            Height          =   315
            Index           =   3
            Left            =   5880
            TabIndex        =   3
            Tag             =   "O"
            Top             =   414
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            Height          =   315
            Index           =   10
            Left            =   3675
            TabIndex        =   2
            Tag             =   "O"
            Top             =   420
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CODTIT"
            Height          =   330
            Index           =   7
            Left            =   1710
            TabIndex        =   0
            Tag             =   "O"
            Top             =   30
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   30
            Top             =   30
            Width           =   5070
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   32
            Top             =   1182
            Width           =   5070
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   31
            Top             =   798
            Width           =   5070
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CODTIT"
            Height          =   330
            Index           =   0
            Left            =   1710
            TabIndex        =   4
            Tag             =   "O"
            Top             =   810
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CODTIT"
            Height          =   330
            Index           =   1
            Left            =   1710
            TabIndex        =   5
            Tag             =   "O"
            Top             =   1182
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CODTIT"
            Height          =   330
            Index           =   2
            Left            =   1710
            TabIndex        =   6
            Tag             =   "O"
            Top             =   1566
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtDados 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   33
            Top             =   1566
            Width           =   5070
         End
         Begin MSMask.MaskEdBox MebDados 
            Height          =   315
            Index           =   9
            Left            =   1710
            TabIndex        =   1
            Tag             =   "O"
            Top             =   414
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Doc.(Ch):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   1
            Left            =   5130
            TabIndex        =   25
            Top             =   1950
            Width           =   690
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   1950
            Width           =   960
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cta. de Controle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   14
            Left            =   150
            TabIndex        =   23
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Lançamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   22
            Top             =   414
            Width           =   885
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Dt. Oper."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   3
            Left            =   2910
            TabIndex        =   21
            Top             =   414
            Width           =   645
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Valor R$"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   5
            Left            =   5100
            TabIndex        =   20
            Top             =   414
            Width           =   750
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Conta Financeira"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   6
            Left            =   150
            TabIndex        =   19
            Top             =   798
            Width           =   1215
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Conta Pl. de Contas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   7
            Left            =   150
            TabIndex        =   18
            Top             =   1182
            Width           =   1410
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   8
            Left            =   150
            TabIndex        =   17
            Top             =   1566
            Width           =   630
         End
      End
   End
End
Attribute VB_Name = "FrmMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditarMovimento As Boolean
Public CodMovimento As Double
Public CodControle As Double
Dim Resp As String
Dim Senha As ADODB.Recordset
Dim Cancelar As Boolean

Private Sub BtAtualiza_Click()
    If IsNumeric(MebDados(5)) Then
        AtualizaGrid
    Else
        MsgBox "É necessário pelo menos informar a conta de controle!", vbInformation
        MebDados(5).SetFocus
        SelecionaCampo
    End If
End Sub

Private Sub Form_Activate()
    
    TabCad.TabEnabled(0) = False
    BtAcoes_Click 0

End Sub

Private Sub Form_Load()

    Set Registros = New ADODB.Recordset
    Registros.CursorLocation = adUseClient
    
    Set Consulta = New ADODB.Recordset
    Consulta.CursorLocation = adUseClient

    Set Formulário = Me
              
    CentralizarForm Me, 1
    'mebdados_Change 6

    Set Senha = New ADODB.Recordset
    Senha.Open "Select * From Pwd Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Senha.Eof = False And Senha.Bof = False And Not IsNull(Senha("DbMov")) Then
        For x = 1 To Len(Senha("DbMov"))
            ManterDbcDados(Mid(Senha("DbMov"), x, 1)).value = 1
        Next
    End If

'    If Banco.rsCtr.RecordCount = 1 Then
'        mebdados(6).Text = Banco.rsCtr("Descric")
'        lblCtr.Enabled = False
'        mebdados(6).Enabled = False
'    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 114 And BtAcoes(0).Enabled = True And TabCad.Tab <> 0 Then   'tecla F3
    '    BtAcoes_Click (0) 'novo
    'ElseIf KeyCode = 115 And BtAcoes(1).Enabled = True And TabCad.Tab <> 0 Then   'tecla F4
    '    BtAcoes_Click (1) 'editar
    If KeyCode = 116 And BtAcoes(2).Enabled = True And TabCad.Tab <> 0 Then   'tecla F5
        BtAcoes_Click (2) 'salvar
    ElseIf KeyCode = 117 And BtAcoes(3).Enabled = True And TabCad.Tab <> 0 Then   'tecla F6
        BtAcoes_Click (3) 'cancelar
    'ElseIf KeyCode = 118 And BtAcoes(4).Enabled = True And TabCad.Tab <> 0 Then   'tecla F7
    '    BtAcoes_Click (4) 'excluir
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'tecla enter
        SendKeys "{tab}"
    ElseIf KeyAscii = 5 Then
        TabCad.Tab = 0
    ElseIf KeyAscii = 15 Then
        TabCad.Tab = 1
    ElseIf KeyAscii = 27 Then  'tecla esc
        Unload Me 'fecha formulario
    'ElseIf KeyAscii = 27 And BtAcoes(0).Enabled = False Then 'tecla esc
        'BtAcoes_Click 3
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    FuncaoTeclas KeyCode, 2
End Sub

Private Sub BtAcoes_Click(Index As Integer)
Dim Excluiu As Boolean
Excluiu = False

If TabCad.Tab = 1 Then
    Select Case Index
        Case 0 'novo
            EditarMovimento = False
            'LimpaCampos
            HabilitaFrames True
            'SalvaMovimento
            Botoes
            MebDados(7).SetFocus
            
        Case 1 And MebDados(9).Enabled = False And MebDados(9) <> "" 'editar
            EditarMovimento = True
            HabilitaFrames True
            'SalvaMovimento
            Botoes
            MebDados(10).SetFocus
            SelecionaCampo
            'mebdados(7).SetFocus
        Case 2 'salvar
        
            If Not IsNumeric(MebDados(3)) Then
                MsgBox "É necessário ter um valor!", vbInformation
                MebDados(3).SetFocus
                Exit Sub
            End If
        
            For Each Componentes In MebDados
                If Componentes.Tag = "O" And Not (IsNumeric(Componentes.Text) Or IsDate(Componentes.Text)) Then
                    MsgBox "Campo obrigatório !", vbCritical
                    Componentes.SetFocus
                    Exit Sub
                End If
            Next

                SalvaMovimento
                HabilitaFrames False
                LimpaCampos
                Botoes
                Unload Me
            'End If
                       
            'BtAcoes_Click 0
            
            'mebdados(7).Enabled = False
            'mebdados(9).Enabled = False
            'MebDados_GotFocus 0
            'mebdados(10).SetFocus
            
        Case 3 'cancelar
            If EditarMovimento = True Then
                Resp = MsgBox("Deseja cancelar as alterações do movimento atual ???", 36)
            Else
                Resp = MsgBox("Deseja cancelar o movimento atual ???", 36)
            End If
            
            If Resp = vbYes Then
                LimpaCampos
                HabilitaFrames False
                Botoes
                Unload Me
            End If
            
        Case 4 And MebDados(9).Enabled = False And MebDados(9) <> "" 'excluir
            If MsgBox("Deseja excluir o movimento atual ???", 36) = vbYes Then
                ExcluiMovimento
                Excluiu = True
            End If
    End Select
    
       If Index = 1 Or Index = 4 Then
            If MebDados(9).Enabled = False And MebDados(9) = "" And Excluiu = False Then MsgBox "Operação não permitida, selecione um movimento!", vbInformation
       End If
       
End If
End Sub

Private Sub Botoes()
    For Each Componentes In BtAcoes
        Componentes.Enabled = Not Componentes.Enabled
    Next
    
    'If BtAcoes(1).Enabled = False Then
    '    TabCad.TabEnabled(0) = False
    'Else
        TabCad.TabEnabled(0) = False
    'End If
End Sub

Private Sub HabilitaFrames(Condicao As Boolean)
    For Each Componentes In Frame
        Componentes.Enabled = Condicao
    Next
    
    If EditarMovimento = True Then
        MebDados(7).Enabled = False
        MebDados(9).Enabled = False
    Else
        MebDados(7).Enabled = True
        MebDados(9).Enabled = True
    End If
    
End Sub

Private Sub LimpaCampos()
    For Each Componentes In MebDados
        If ManterDbcDados(Componentes.Index).value = False Then
            'If Componentes.Tag <> "N" Then
                Mascara = Componentes.Mask
                Componentes.Mask = ""
                Componentes.Text = ""
                Componentes.Mask = Mascara
            'End If
        End If
    Next
    
    CodMovimento = 0
    CodControle = 0
    EditarMovimento = False
    
    For Each Componentes In TxtDados
        If Componentes.Tag <> "N" Then
            Componentes.Text = ""
        End If
    Next
       
    'AtualizaObjeto Banco.rsPesqItm, "Select * From ITM Where ITMCodITM = 0", GridMovimento, 1
    
    HabilitaFrames False
  
End Sub

Public Function SalvaMovimento()
Dim NumCampos As Integer
NumCampos = 0

        'For Each Componentes In MebDados
        '    If Componentes.Tag = "O" And Not (IsNumeric(Componentes.Text) Or IsDate(Componentes.Text)) Then
        '        MsgBox "Campo obrigatório não numérico!", vbInformation
        '        Componentes.SetFocus
        '        SelecionaCampo
        '    End If
        'Next
              
'        If NumCampos = 0 Then
            BancoSRCaixa.Execute "Delete From Mov Where CodMov = " & MebDados(9) & " and CodCtr = " & MebDados(7)
            InsereBanco "MOV", "CodMov, CodCtr, CodCtf, CodCta, CodHis, Data, Valor, Complem, Docum", MebDados(9) & "," & MebDados(7) & "," & MebDados(0) & "," & MebDados(1) & "," & MebDados(2) & ",'" & Format(MebDados(10), "yyyy-mm-dd") & "'," & Replace(MebDados(3), ",", ".") & ",'" & TxtDados(0) & "','" & TxtDados(1) & "'"
            
            CodMovimento = MebDados(9)
            
            'SalvaMovimento = True
'        End If
        
End Function

Private Sub PreencheMovimento(CodigoMov As Currency, CodigoCtr As Currency)
'On Error Resume Next

LimpaCampos

CodMovimento = CodigoMov
CodControle = CodigoCtr

If Registros.State = 1 Then Registros.Close
Registros.Open "Select * From MOV Where CodMov = " & CodigoMov & " And CodCtr = " & CodigoCtr, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText

    If Registros.Eof = False Then
        'AtualizaObjeto Banco.rsPesquisa, "SELECT MOV.CODMOV, MOV.CODCTR, CTR.DESCRIC, HIS.DESCRIC FROM ((MOV INNER JOIN CTR ON MOV.CODCTR = CTR.CODCTR) INNER JOIN CTA ON MOV.CODCTA = CTA.CODCTA) INNER JOIN HIS ON MOV.CODHIS = HIS.CODHIS WHERE MOV.CODCTR=" & mebdados(6).text, DbGrid1, 1
            MebDados(9) = IIf(IsNull(Registros!CodMov), 0, Registros!CodMov)
            MebDados(7).Text = IIf(IsNull(Registros!CodCtr), 0, Registros!CodCtr)
            MebDados(0).Text = IIf(IsNull(Registros!CodCtf), 0, Registros!CodCtf)
            MebDados(1).Text = IIf(IsNull(Registros!CodCta), 0, Registros!CodCta)
            MebDados(2).Text = IIf(IsNull(Registros!CodHis), 0, Registros!CodHis)
            MebDados(10) = IIf(IsNull(Registros!Data), 0, Registros!Data)
            MebDados(3) = IIf(IsNull(Registros!Valor), 0, Registros!Valor)
            TxtDados(0) = IIf(IsNull(Registros!Complem), 0, Registros!Complem)
            TxtDados(1) = IIf(IsNull(Registros!Docum), 0, Registros!Docum)
            
            MebDados_KeyDown 7, 13, 0
            MebDados_KeyDown 0, 13, 0
            MebDados_KeyDown 1, 13, 0
            MebDados_KeyDown 2, 13, 0
            
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim varManterMov As String

'Set Senha = New ADODB.Recordset
'    Senha.Open "Select * From Pwd Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'", Banco, adOpenDynamic, adLockOptimistic, adCmdText

    
    For Each Componentes In ManterDbcDados
        If Componentes.value = 1 Then
            varManterMov = varManterMov & Componentes.Index
            'Debug.Print Componentes.Index
        End If
    Next
        
 '   If Senha.EOF = False And Senha.BOF = False And varManterMov <> "" Then
        BancoSRCaixa.Execute "UPDATE PWD SET DbMov = '" & varManterMov & "' Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'"
        
'        Senha("DbMov") = varManterMov
'        Senha.Update
 '   End If
    
End Sub

Private Sub ExcluiMovimento()
    
    BancoSRCaixa.Execute "Delete From MOV Where CodMov = " & CodMovimento & " And CodCtr = " & CodControle
    LimpaCampos
        
End Sub

Private Sub LblDados_Click(Index As Integer)

End Sub

Private Sub lvPesquisa_DblClick()
    PreencheMovimento lvPesquisa.SelectedItem.Text, MebDados(5).Text
    TabCad.Tab = 1
End Sub

Private Sub MebDados_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error Resume Next
Select Case Index
    Case 10
        If KeyAscii = 27 Then
            'mebdados(7).Enabled = True
            'mebdados(9).Enabled = True
            'mebdados(7).SetFocus
        End If
End Select
End Sub

Private Sub MebDados_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'If IsDate(MebDados(Index)) Then
'    AtualizaGrid
'End If
End Sub

Private Sub MebDados_LostFocus(Index As Integer)

Select Case Index
    Case 9
        If IsNumeric(MebDados(7)) Then
            SQL = "Select * From Mov Where CodCtr = " & MebDados(7).Text & " and CodMov = " & MebDados(9).Text
            
                Set Consulta = New ADODB.Recordset
                If MebDados(9).Text <> "" Then
                    Consulta.Open SQL, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    If Consulta.Bof = False Then
                        If MsgBox("Existe um lançamento com esse Código, deseja altera-lo?", 36) = vbYes Then
                            Botoes
                            PreencheMovimento MebDados(9), MebDados(7).Text
                            BtAcoes_Click 1
                        Else
                            LimpaCampos
                            BtAcoes_Click 0
                        End If
                    End If
                End If
        End If
    Case 10, 4
        If MebDados(Index) = "__/__/____" Then
        Else
            MebDados(Index) = Converte(MebDados(Index))
        End If
    Case 5, 6, 0, 1, 2, 7
        'AtualizaGrid
        Status -1
        
End Select
End Sub

Private Sub AtualizaGrid()
Dim SQL As String
Dim Ordenado As String
Dim Formula As String
Dim CondWhere(1 To 3) As String
Dim Data1 As String
Dim Data2 As String
Dim LstPesq As ListItem
Dim TotalReg As Integer
Dim Limite As Integer

DoEvents

If IsDate(MebDados(11)) And IsDate(MebDados(4)) Then
    Data1 = Format(MebDados(11), "mm/dd/yyyy")
    Data2 = Format(MebDados(4), "mm/dd/yyyy")
'ElseIf mebdados(11) <> "__/__/____" And MebDados(4) <> "__/__/____" Then
End If

CondWhere(1) = ""
CondWhere(2) = ""
CondWhere(3) = ""
x = 0
y = 0

If IsNumeric(MebDados(5)) Then CondWhere(1) = " MOV.CODCTR=" & MebDados(5).Text
If IsNumeric(MebDados(6)) Then CondWhere(2) = " MOV.CODCTF=" & MebDados(6).Text
If IsDate(Data1) And IsDate(Data2) Then CondWhere(3) = " MOV.DATA >= #" & Data1 & "# and MOV.DATA <= #" & Data2 & "#"

Formula = " Where Mov.CodMov <> 0 "

If CondWhere(1) <> "" Then Formula = Formula & " and " & CondWhere(1) & " "
If CondWhere(2) <> "" Then Formula = Formula & " and " & CondWhere(2) & " "
If CondWhere(3) <> "" Then Formula = Formula & " and " & CondWhere(3) & " "

For Each Componentes In opOrdem
    If Componentes.value = True Then
        Ordenado = Componentes.Tag
    End If
Next

SQL = "SELECT Count(mov.codmov) AS ContarDeCodigo From mov " & Formula

        If Consulta.State = 1 Then Consulta.Close
        Consulta.Open SQL, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
        
        TotalReg = Consulta("ContarDeCodigo")
        
        If ChLimite.value = 1 Then
            TotalReg = 100
            Limite = 100
        Else
            Limite = TotalReg
        End If
        
        If TotalReg > 100 Then
            Progresso.Max = (TotalReg / 100) + 2
        Else
            Progresso.Max = TotalReg + 2
        End If
                            
        If TotalReg > 2000 And ChLimite.value = False Then
            If MsgBox("O Sistema retornará " & TotalReg & " registros, isso pode demorar alguns segundos, deseja continuar?", 36) = vbNo Then
                Exit Sub
            End If
        End If
                
SQL = "SELECT MOV.CODMOV, MOV.CODCTR, CTR.DESCRIC, HIS.DESCRIC, CTA.CODCTA, CTA.DESCRIC, MOV.DATA, MOV.VALOR FROM ((MOV LEFT JOIN CTR ON MOV.CODCTR = CTR.CODCTR) LEFT JOIN CTA ON MOV.CODCTA = CTA.CODCTA) LEFT JOIN HIS ON MOV.CODHIS = HIS.CODHIS " & Formula & " order by " & Ordenado
        
        If Consulta.State = 1 Then Consulta.Close
        Consulta.Open SQL, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
        
           lvPesquisa.ListItems.Clear
           Progresso.Visible = True

           Do While Not Consulta.Eof
                Z = Z + 1
                If Z >= Limite Then Exit Do
               Set LstPesq = lvPesquisa.ListItems.Add(, , Consulta("CodMov"))
               LstPesq.SubItems(1) = SeCampoNulo(Consulta, 5)
               LstPesq.SubItems(2) = SeCampoNulo(Consulta, "Data")
               LstPesq.SubItems(3) = Format(SeCampoNulo(Consulta, "Valor"), "#,#0.00")
               LstPesq.SubItems(4) = Format(SeCampoNulo(Consulta, "CodCtr"), "#,#0.00")
               Consulta.MoveNext
                
                x = x + 1
                If TotalReg > 100 Then
                    If x = 100 Then
                        y = y + 1
                        Progresso.value = y
                        Me.Caption = Me.Tag & " --> " & y * 100 & " Registros carregados"
                        x = 0
                    End If
                Else
                    Progresso.value = x
                    Me.Caption = Me.Tag & " --> " & x & " Registros carregados"
                End If
           
           Loop
           
           Me.Caption = Me.Tag
           Progresso.Visible = False

End Sub

Private Sub OpOrdem_Click(Index As Integer)
    'AtualizaGrid
End Sub

Private Sub MebDados_Change(Index As Integer)
Select Case Index
'    Case 5, 6, 7, 0, 1
'        MebDados_KeyDown Index, 13, 0
End Select
End Sub

Private Sub MebDados_GotFocus(Index As Integer)
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Select Case Index
    Case 5, 6, 0, 1, 2, 7
        Status 4
    Case 9
        If IsNumeric(MebDados(7)) And CStr(MebDados(7)) <> "0" Then
            Rs.Open "Select * From CTR Where CodCtr = " & MebDados(7), BancoSRCaixa, adOpenStatic, adCmdText
        
            If IsNull(Rs("CodMov")) Then
                MebDados(Index).Text = 1
            Else
                MebDados(Index) = Rs("CodMov") + 1
            End If
            
            BancoSRCaixa.Execute "UPDATE CTR SET CodMov = " & MebDados(Index).Text & " Where CodCtr = " & MebDados(7)
            MebDados(Index).SelLength = 5000
                   
            Rs.Close
            Set Rs = Nothing
        Else
            MsgBox "É necessário selecionar uma conta de controle!", 16
            MebDados(7).SetFocus
            SelecionaCampo
        End If

    Case 10
        If Not IsDate(MebDados(10)) Then MebDados(10).Text = Date
    
End Select

End Sub

Private Sub MebDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 5, 7
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CODCTR"
            FrmPesquisaSRCaixa.Campos2 = "DESCRIC"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "CTR"
            FrmPesquisaSRCaixa.Condicao1 = "CODCTR > 1"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Contas de Controle"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
               
        If KeyCode = 13 Then
            If Busca("CODCTR", MebDados(Index), "CTR") = True Then
                If Index = 5 Then TxtDados(2).Text = SeCampoNulo(Consulta, "DESCRIC")
                If Index = 7 Then TxtDados(4).Text = SeCampoNulo(Consulta, "DESCRIC")
            End If
        End If
        
    Case 6, 0
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CODCTF"
            FrmPesquisaSRCaixa.Campos2 = "DESCRIC"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "CTF"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Contas Financeiras"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
               
        If KeyCode = 13 Then
            If Busca("CODCTF", MebDados(Index), "CTF") = True Then
                If Index = 6 Then TxtDados(3).Text = SeCampoNulo(Consulta, "DESCRIC")
                If Index = 0 Then TxtDados(5).Text = SeCampoNulo(Consulta, "DESCRIC")
            End If
        End If
    
    Case 1
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CODCTA"
            FrmPesquisaSRCaixa.Campos2 = "DESCRIC"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "CTA"
            'FrmPesquisaSRCaixa.Condicao1 = "CODCTA <> 1000 and CODCTA <> 2000 and Mid(CodCta, 1, 5)>0"
            FrmPesquisaSRCaixa.Condicao1 = "CODCTA <> 1000 and CODCTA <> 2000 and Mid(CodCta, 4, 1)>0"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Plano de Contas"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
               
        If KeyCode = 13 Then
            If Mid(MebDados(Index), 4, 1) = 0 Then
                MsgBox "Não pode haver lançamentos em Contas-chave!", vbInformation
                MebDados(Index).SetFocus
                SelecionaCampo
            Else
                If Busca("CODCTA", MebDados(Index), "CTA") = True Then
                    TxtDados(6).Text = SeCampoNulo(Consulta, "DESCRIC")
                End If
            End If
        End If
    
    Case 2
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CODHIS"
            FrmPesquisaSRCaixa.Campos2 = "DESCRIC"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "HIS"
            'FrmPesquisaSRCaixa.Condicao1 = "CODCTA <> 1000 and CODCTA <> 2000"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Plano de Contas"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
               
        If KeyCode = 13 Then
            If Busca("CODHIS", MebDados(Index), "HIS") = True Then
                TxtDados(7).Text = SeCampoNulo(Consulta, "DESCRIC")
            End If
        End If
    
    
End Select

End Sub
