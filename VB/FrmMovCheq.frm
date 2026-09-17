VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMovCheq 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5235
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9060
   Icon            =   "FrmMovCheq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9060
   Tag             =   "Lançamentos dos Movimentos"
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   -90
      ScaleHeight     =   285
      ScaleWidth      =   14025
      TabIndex        =   30
      Top             =   4950
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
         Left            =   45
         TabIndex        =   31
         Top             =   45
         Width           =   9255
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
      TabIndex        =   28
      Top             =   0
      Width           =   11085
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Movimentação de Cheques"
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
         TabIndex        =   29
         Top             =   45
         Width           =   6540
      End
   End
   Begin TabDlg.SSTab TabCad 
      Height          =   3885
      Left            =   360
      TabIndex        =   13
      Top             =   765
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   6853
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ch&eques (Ctrl + E)"
      TabPicture(0)   =   "FrmMovCheq.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "M&ovimento  (Ctrl + O)"
      TabPicture(1)   =   "FrmMovCheq.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "BtAcoes(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "BtAcoes(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Descontar (Ctrl + D)"
      TabPicture(2)   =   "FrmMovCheq.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton BtAcoes 
         Caption         =   "&Cancel. (F6)"
         Enabled         =   0   'False
         Height          =   765
         Index           =   3
         Left            =   6990
         Picture         =   "FrmMovCheq.frx":0496
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CommandButton BtAcoes 
         Caption         =   "&Incluir (F5)"
         Enabled         =   0   'False
         Height          =   765
         Index           =   2
         Left            =   5985
         Picture         =   "FrmMovCheq.frx":0CB0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2970
         Width           =   950
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2385
         Index           =   4
         Left            =   165
         TabIndex        =   14
         Top             =   480
         Width           =   7965
         Begin VB.Frame Frame7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00BB5B0D&
            Height          =   2400
            Left            =   15
            TabIndex        =   16
            Top             =   -60
            Width           =   7950
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00BB5B0D&
               Height          =   330
               Index           =   4
               Left            =   2415
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   27
               Top             =   1860
               Width           =   5385
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00BB5B0D&
               Height          =   315
               Index           =   0
               Left            =   3930
               MaxLength       =   30
               TabIndex        =   8
               Top             =   1470
               Width           =   3870
            End
            Begin VB.TextBox TxtDados 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00BB5B0D&
               Height          =   330
               Index           =   1
               Left            =   2415
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   12
               Top             =   270
               Width           =   5385
            End
            Begin MSMask.MaskEdBox MebDados 
               Height          =   315
               Index           =   3
               Left            =   6120
               TabIndex        =   6
               Tag             =   "O"
               Top             =   1074
               Width           =   1680
               _ExtentX        =   2963
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
               Left            =   3945
               TabIndex        =   5
               Tag             =   "O"
               Top             =   1074
               Width           =   1380
               _ExtentX        =   2434
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
               Index           =   1
               Left            =   1425
               TabIndex        =   0
               Tag             =   "O"
               Top             =   270
               Width           =   920
               _ExtentX        =   1614
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
               Index           =   7
               Left            =   1425
               TabIndex        =   1
               Top             =   672
               Width           =   915
               _ExtentX        =   1614
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
               Index           =   8
               Left            =   3030
               TabIndex        =   2
               Top             =   675
               Width           =   1785
               _ExtentX        =   3149
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
               Index           =   9
               Left            =   6120
               TabIndex        =   3
               Top             =   672
               Width           =   1680
               _ExtentX        =   2963
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
               Left            =   1425
               TabIndex        =   9
               Top             =   1860
               Width           =   915
               _ExtentX        =   1614
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
               Height          =   315
               Index           =   12
               Left            =   1425
               TabIndex        =   4
               Tag             =   "O"
               Top             =   1074
               Width           =   1380
               _ExtentX        =   2434
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
               Index           =   0
               Left            =   1425
               TabIndex        =   7
               Tag             =   "O"
               Top             =   1455
               Width           =   1740
               _ExtentX        =   3069
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
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Nome"
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
               Index           =   4
               Left            =   3330
               TabIndex        =   26
               Top             =   1470
               Width           =   405
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Cpf / CNPJ"
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
               Left            =   135
               TabIndex        =   25
               Top             =   1470
               Width           =   765
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Bom para"
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
               Index           =   9
               Left            =   135
               TabIndex        =   24
               Top             =   1074
               Width           =   690
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Chq. Terceiro"
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
               Left            =   135
               TabIndex        =   23
               Top             =   1860
               Width           =   975
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Banco"
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
               Left            =   135
               TabIndex        =   22
               Top             =   270
               Width           =   465
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
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
               Left            =   5505
               TabIndex        =   21
               Top             =   1050
               Width           =   390
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Data Oper."
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
               Left            =   2985
               TabIndex        =   20
               Top             =   1074
               Width           =   780
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Agência"
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
               Left            =   135
               TabIndex        =   19
               Top             =   672
               Width           =   600
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Conta"
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
               Left            =   2490
               TabIndex        =   18
               Top             =   672
               Width           =   420
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Cheque"
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
               Left            =   4995
               TabIndex        =   17
               Top             =   672
               Width           =   900
            End
         End
      End
   End
   Begin MSComctlLib.ProgressBar Progresso 
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "FrmMovCheq"
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

Private Sub BtAcoes_KeyPress(Index As Integer, KeyAscii As Integer)
    'If Index = 0 And Chr(KeyAscii) = "<" Then
    '    BtAcoes_Click 0
    '    SendKeys "<"
    'End If
End Sub

Private Sub chDesc_Click()
'    MebDados_KeyDown 4, 13, 0
End Sub

Private Sub Form_Activate()

    BtAcoes_Click 0

End Sub

Private Sub Form_Load()

    Set Registros = New ADODB.Recordset
    Registros.CursorLocation = adUseClient
    
    Set Consulta = New ADODB.Recordset
    Consulta.CursorLocation = adUseClient

    Set Formulário = Me
              
    CentralizarForm Me, 1
    
    'AtualizaGrid 1
        
    'TabCad_Click 0
    
    
    'mebdados_Change 6

'    Set Senha = New ADODB.Recordset
'    Senha.Open "Select * From Pwd Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'", Banco, adOpenDynamic, adLockOptimistic, adCmdText
    
'    If Senha.EOF = False And Senha.BOF = False And Not IsNull(Senha("DbMov")) Then
'        For x = 1 To Len(Senha("DbMov"))
'            ManterDbcDados(Mid(Senha("DbMov"), x, 1)).Value = 1
'        Next
'    End If

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
    ElseIf KeyAscii = 4 Then
        TabCad.Tab = 2
    ElseIf KeyAscii = 27 Then 'tecla esc
        Unload Me 'fecha formulario
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    FuncaoTeclas KeyCode, 2
End Sub

Private Sub BtAcoes_Click(Index As Integer)
Dim Excluiu As Boolean
Excluiu = False

'If TabCad.Tab = 1 Then
    Select Case Index
        Case 0 'novo
            EditarMovimento = False
            'LimpaCampos
            HabilitaFrames True
            'SalvaMovimento
            Botoes
            MebDados(1).SetFocus
            
        'Case 1 And MebDados(9).Enabled = False And MebDados(9) <> "" 'editar
        '    EditarMovimento = True
        '    HabilitaFrames True
        '    'SalvaMovimento
        '    Botoes
        '    MebDados(10).SetFocus
        '    SelecionaCampo
        '    'mebdados(7).SetFocus
        Case 2 'salvar
        
                For Each Componentes In MebDados
                    If Componentes.Tag = "O" And Not (IsNumeric(Componentes.Text) Or IsDate(Componentes.Text)) Then
                        MsgBox "Campo obrigatório não numérico!", vbInformation
                        Componentes.SetFocus
                        SelecionaCampo
                        Exit Sub
                    End If
                Next
        
                If MebDados(2).Text = "" Then MebDados(2) = 0
                SalvaMovimento
                Unload Me
                'HabilitaFrames False
                'LimpaCampos
                'AtualizaGrid 1
                'Botoes
                'BtAcoes(0).SetFocus
            'End If
                       
            'BtAcoes_Click 0
            
            'mebdados(7).Enabled = False
            'mebdados(9).Enabled = False
            'MebDados_GotFocus 0
            'mebdados(10).SetFocus
            
        Case 3 'cancelar
            
            Resp = MsgBox("Deseja cancelar o lançamento atual ???", 36)
            
            If Resp = vbYes Then
                LimpaCampos
                HabilitaFrames False
                Botoes
                BtAcoes(0).SetFocus
            End If
            
        Case 4, 5 'excluir
            If MsgBox("Deseja Excluir o Cheque Selecionado?", 36) = vbYes Then
                If TabCad.Tab = 0 Then
                    'BancoSRCaixa.Execute "Delete From Chq Where Lancamento = " & lvChqs2.SelectedItem.Text
                    'AtualizaGrid 2
                ElseIf TabCad.Tab = 1 Then
                    BancoSRCaixa.Execute "Delete From Chq Where Lancamento = " & lvChqs.SelectedItem.Text
                    'AtualizaGrid 1
                    BtAcoes(0).SetFocus
                End If

            End If
    
        Case 6
            
            Dim DtBomPara As Date
            Dim DataDesconto As Date
            Dim Dias As Integer
            Dim Juros As Currency
            Dim SomaCheq As Currency
            Dim SomaJuros As Currency
            
            If MebDados(14).Text = "" Then
                MsgBox "Selecione uma instituição financeira!", vbInformation
                MebDados(14).SetFocus
                SelecionaCampo
                Exit Sub
            End If
                
            If Not IsNumeric(MebDados(15)) Then
                MsgBox "Juros não numérico!", vbInformation
                MebDados(15).SetFocus
                SelecionaCampo
                Exit Sub
            End If
            
            If Not IsDate(MebDados(16)) Then
                MsgBox "Data Inválida!", vbInformation
                MebDados(16).SetFocus
                SelecionaCampo
                Exit Sub
            End If
                       
            SomaCheq = 0
            SomaJuros = 0
            DataDesconto = CDate(MebDados(16).Text)
            
            SQL = "SELECT CHQ.Lancamento, CHQ.BomPara, CHQ.CodMov, CHQ.NroChq, Nom.Nome, CHQ.Valor FROM Nom INNER JOIN CHQ ON Nom.CgcCpf = CHQ.ClienteCpf "
           
            
            If Consulta.State = 1 Then Consulta.Close
            Consulta.Open SQL, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
            
            For x = 1 To lvChqs3.ListItems.Count
                If lvChqs3.ListItems.Item(x).Checked = True Then
                    Consulta.Find "Lancamento = " & lvChqs3.ListItems(x).Text, , , 0
                        If Consulta.Eof = False Then
                            Dias = Consulta("BomPara") - DataDesconto
                            Juros = ((CCur(MebDados(15)) * Dias) / 100) * Consulta("Valor")
                            SomaCheq = SomaCheq + Consulta("Valor")
                            SomaJuros = SomaJuros + Juros
                            
                            If Dias < 0 Then Dias = 0
                            If Juros < 0 Then Juros = 0
                            If SomaJuros < 0 Then SomaJuros = 0
                            
                            BancoSRCaixa.Execute "UPDATE CHQ SET CHQ.Descontado = True, CHQ.DataDesc = #" & Format(DataDesconto, "mm/dd/yyyy") & "#, CHQ.DiasDesc = " & Dias & ", CHQ.JurosDesc = " & Replace(Juros, ",", ".") & ", CHQ.DescPor = '" & MebDados(14) & "' Where (((CHQ.Lancamento) = " & Consulta("Lancamento") & "))"
                        End If
                End If
            Next
            
            'If Dias <> 0 Then
                MsgBox "Cheques descontados com sucesso!" & Chr(10) & Chr(10) & "Total de Cheques: " & Format(SomaCheq, "#,#0.00") & Chr(10) & "Juros: " & Format(SomaJuros, "#,#0.00") & Chr(10) & "Valor Líquido: " & Format(SomaCheq - SomaJuros, "#,#0.00"), vbInformation
            'Else
            '    MsgBox "Não foi selecionado cheques a serem descontados!", vbInformation
            'End If
            
            'AtualizaGrid 3
            
            
    End Select
End Sub

Private Sub Botoes()
    For Each Componentes In BtAcoes
        If Componentes.Index <> 4 Then
            Componentes.Enabled = Not Componentes.Enabled
        End If
    Next
    
    'If BtAcoes(0).Enabled = False Then
    '    TabCad.TabEnabled(0) = False
    'Else
        
        TabCad.TabEnabled(0) = False
        TabCad.TabEnabled(1) = True
        TabCad.TabEnabled(2) = False
    'End If
End Sub

Private Sub HabilitaFrames(Condicao As Boolean)
    For Each Componentes In Frame
        Componentes.Enabled = Condicao
    Next
    
    If EditarMovimento = True Then
        'MebDados(7).Enabled = False
        'MebDados(9).Enabled = False
    Else
        'MebDados(7).Enabled = True
        'MebDados(9).Enabled = True
    End If
    
End Sub

Private Sub LimpaCampos()
    For Each Componentes In MebDados
'        If ManterDbcDados(Componentes.Index).Value = False Then
            'If Componentes.Tag <> "N" Then
                Mascara = Componentes.Mask
                Componentes.Mask = ""
                Componentes.Text = ""
                Componentes.Mask = Mascara
            'End If
'        End If
    Next
    
    CodMovimento = 0
    CodControle = 0
    EditarMovimento = False
    
    For Each Componentes In TxtDados
        If Componentes.Tag <> "N" Then
            Componentes.Text = ""
        End If
    Next
       
    'AtualizaObjeto BancoSRCaixa.rsPesqItm, "Select * From ITM Where ITMCodITM = 0", GridMovimento, 1
    
    HabilitaFrames False
  
End Sub

Public Function SalvaMovimento()
Dim NumCampos As Integer
NumCampos = 0

             
'        If NumCampos = 0 Then
            'BancoSRCaixa.Execute "Delete From Mov Where CodMov = " & MebDados(9) & " and CodCtr = " & MebDados(7)
            InsereBanco "Chq", "ClienteCpf, Banco, Agencia, Conta, NroChq, Bompara, DataOper, Valor, ChqTerc", "'" & MebDados(0) & "'," & MebDados(1) & ",'" & MebDados(7) & "','" & MebDados(8) & "','" & MebDados(9) & "','" & MebDados(12) & "','" & MebDados(10) & "'," & Replace(MebDados(3), ",", ".") & "," & MebDados(2)
            'CodMovimento = MebDados(9)
            
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
        'AtualizaObjeto BancoSRCaixa.rsPesquisa, "SELECT MOV.CODMOV, MOV.CODCTR, CTR.DESCRIC, HIS.DESCRIC FROM ((MOV INNER JOIN CTR ON MOV.CODCTR = CTR.CODCTR) INNER JOIN CTA ON MOV.CODCTA = CTA.CODCTA) INNER JOIN HIS ON MOV.CODHIS = HIS.CODHIS WHERE MOV.CODCTR=" & mebdados(6).text, DbGrid1, 1
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
'    Senha.Open "Select * From Pwd Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText

    
    'For Each Componentes In ManterDbcDados
    '    If Componentes.Value = 1 Then
    '        varManterMov = varManterMov & Componentes.Index
    '        'Debug.Print Componentes.Index
    '    End If
    'Next
        
 '   If Senha.EOF = False And Senha.BOF = False And varManterMov <> "" Then
'        BancoSRCaixa.Execute "UPDATE PWD SET DbMov = '" & varManterMov & "' Where Nome Like '" & varUsuario & "' and Senha Like '" & varSenha & "'"
        
'        Senha("DbMov") = varManterMov
'        Senha.Update
 '   End If
    
End Sub

Private Sub ExcluiMovimento()
    
    BancoSRCaixa.Execute "Delete From MOV Where CodMov = " & CodMovimento & " And CodCtr = " & CodControle
    LimpaCampos
        
End Sub

Private Sub lvChqs3_Click()
    SomaCheques
End Sub

Private Sub lvChqs3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    MsgBox Item.Checked
End Sub

Private Sub Label3_Click()

End Sub

Private Sub MebDados_KeyPress(Index As Integer, KeyAscii As Integer)
    'If KeyAscii = 27 And Index = 0 Then MebDados_KeyDown 0, 13, 0
End Sub

Private Sub MebDados_LostFocus(Index As Integer)

Select Case Index
    Case 10
        If MebDados(Index) = "__/__/____" Then
        Else
            MebDados(Index) = Converte(MebDados(Index))
        End If
    Case 0, 14
        Status -1
        'If Index = 0 Then MebDados_KeyDown 0, 13, 0
        
End Select
End Sub


Private Sub OpOrdem_Click(Index As Integer)
    'AtualizaGrid
End Sub

Private Sub MebDados_GotFocus(Index As Integer)
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Select Case Index
    Case 0:         Status 4
    Case 10, 12:    If Not IsDate(MebDados(Index)) Then MebDados(Index).Text = Date
    Case 14:        Status 10
    Case 16:        If Not IsDate(MebDados(Index).Text) Then MebDados(Index).Text = Date
End Select
End Sub

Private Sub MebDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    'Case 0
        'If KeyCode = 113 Or KeyCode = 114 Then
            
        '    FrmPesquisaSRCaixa.Campos1 = "CgcCpf"
        '    FrmPesquisaSRCaixa.Campos2 = "Nome"
        '    FrmPesquisaSRCaixa.QtdCampos = 1
        '    FrmPesquisaSRCaixa.TipoResultado = 2
        '    FrmPesquisaSRCaixa.Tabela = "NOM"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
        '    FrmPesquisaSRCaixa.Caption = "Pesquisa de Contas Financeiras"
        '    FrmPesquisaSRCaixa.Show vbModal
        
        '    MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        'End If
               
        'If KeyCode = 13 Then
        '    If Busca("CgcCpf", MebDados(Index), "Nom") = True Then
        '        TxtDados(0).Text = SeCampoNulo(Consulta, "Nome")
        '    Else
        '        TxtDados(0).SetFocus
        '    End If
        'End If
    Case 0
    
        If KeyCode = 13 Then
    
            If Consulta.State = 1 Then Consulta.Close
            Consulta.Open "Select * From Nom Where CgcCpf Like '" & ApenasNumeros(MebDados(0)) & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
           
            If Consulta.Bof Then
                If MsgBox("Cpf não cadastrado!, Deseja Cadastrar?", 36) = vbYes Then
                    'TxtDados(0).Locked = False
                    TxtDados(0).SetFocus
                    TxtDados(0).SelLength = 5000
                Else
                    'TxtDados(0).Locked = True
                    MebDados(0).SetFocus
                    MebDados(0).SelLength = 5000
                End If
            Else
                'TxtDados(0).Locked = True
                TxtDados(0).Text = Consulta("Nome")
                SendKeys "{tab}"
            End If
        End If
    
    Case 1
        If KeyCode = 13 Then
            If Len(MebDados(1)) > 15 Then
            
                Dim varCompleta As String
                Dim varBanco As String
                Dim varAgencia As String
                Dim varConta As String
                Dim varCompensacao As String
                Dim varCheque As String
                Dim varCpf As String
                
                varCompleta = ""
                varBanco = ""
                varAgencia = ""
                varConta = ""
                varCompensacao = ""
                varCheque = ""
                varCpf = ""
                
                varCompleta = MebDados(1)
                    
                varBanco = Mid(varCompleta, 2, 3)
                varAgencia = Mid(varCompleta, 5, 4)
                varConta = Mid(varCompleta, 27, 6)
                varComp = Mid(varCompleta, 11, 3)
                varCheque = Mid(varCompleta, 14, 6)
                
                 If Consulta.State = 1 Then Consulta.Close
                 Consulta.Open "Select * From Chq Where Banco = " & varBanco & " and Agencia Like '" & varAgencia & "' and Conta Like '" & varConta & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
                
                 If Consulta.Bof = False Then
                    varCpf = Consulta("ClienteCpf")
                 End If

                
                MebDados(1).Text = varBanco
                MebDados(7).Text = varAgencia
                MebDados(8).Text = varConta
                MebDados(9).Text = varCheque
                'MebDados(12).Text = varComp
                MebDados(0).Text = varCpf
                
                If varCpf <> "" Then MebDados_KeyDown 0, 13, 0
                MebDados_KeyDown 1, 13, 0
                MebDados(8).SetFocus 'é colocado o campo anterior pois o leitor da um enter

            Else
            
                If Busca("CodBan", MebDados(Index), "Ban") = True Then
                    TxtDados(1).Text = SeCampoNulo(Consulta, "Nome")
                End If
                
            End If
        End If
        
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CodBan"
            FrmPesquisaSRCaixa.Campos2 = "Nome"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "Ban"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Bancos"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
       
    Case 2
    
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CodTit"
            FrmPesquisaSRCaixa.Campos2 = "Nome"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Condicao1 = "Tipo=1"
            FrmPesquisaSRCaixa.Tabela = "Tit"
           
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Clientes"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
               
        If KeyCode = 13 Then
            If Busca("CodTit", MebDados(Index), "Tit") = True Then
                TxtDados(4).Text = SeCampoNulo(Consulta, "Nome")
            End If
        End If
       
    Case 4
    
        If KeyCode = 13 Then
            
            If Len(MebDados(Index)) > 15 Then
                           
                varCompleta = ""
                varBanco = ""
                varAgencia = ""
                varConta = ""
                varCpf = ""
                
                varCompleta = MebDados(Index)
                    
                varBanco = Mid(varCompleta, 2, 3)
                varAgencia = Mid(varCompleta, 5, 4)
                varConta = Mid(varCompleta, 27, 6)
                varComp = Mid(varCompleta, 11, 3)
                varCheque = Mid(varCompleta, 14, 6)
                
                 If Consulta.State = 1 Then Consulta.Close
                 Consulta.Open "Select * From Chq Where Banco = " & varBanco & " and Agencia Like '" & varAgencia & "' and Conta Like '" & varConta & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
                
                 If Consulta.Bof = False Then
                    varCpf = Consulta("ClienteCpf")
                 End If
            
                MebDados(0).Text = varCpf
                MebDados_KeyDown 0, 13, 0
            
            Else
                If Consulta.State = 1 Then Consulta.Close
                Consulta.Open "Select * From Nom Where CgcCpf Like '" & ApenasNumeros(MebDados(4)) & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
                
                If Consulta.Bof Then
                    MsgBox "Cpf não localizado!", vbInformation
                    MebDados(4).SetFocus
                    MebDados(4).SelLength = 5000
                    TxtDados(2).Text = ""
                Else
                    TxtDados(2).Text = IIf(IsNull(Consulta("Nome")), "", Consulta("Nome"))
                    'AtualizaGrid 2
                    SendKeys "{tab}"
                End If
            End If
        End If
    
    Case 8
        If KeyCode = 13 Then
            varBanco = ""
            varAgencia = ""
            varConta = ""
            varCompensacao = ""
            'varCheque = ""
            varCpf = ""
            
            'varCompleta = MebDados(1)
                
            varBanco = MebDados(1)
            varAgencia = MebDados(7)
            varConta = MebDados(8)
            'varComp = Mid(varCompleta, 11, 3)
            'varCheque = Mid(varCompleta, 14, 6)
             
            If Consulta.State = 1 Then Consulta.Close
            Consulta.Open "Select * From Chq Where Banco = " & varBanco & " and Agencia Like '" & varAgencia & "' and Conta Like '" & varConta & "'", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
        
            If Consulta.Bof = False Then
               varCpf = Consulta("ClienteCpf")
               MebDados(0).Text = varCpf
               MebDados(7).SetFocus 'é colocado o campo anterior pois o leitor da um enter
            Else
               MebDados(8).SetFocus
            End If
             
            If varCpf <> "" Then MebDados_KeyDown 0, 13, 0
            MebDados_KeyDown 1, 13, 0
        End If
    
    Case 13
        If KeyCode = 13 Then
            If Busca("CodBan", MebDados(Index), "Ban") = True Then
                TxtDados(3).Text = SeCampoNulo(Consulta, "Nome")
            End If
        End If
        
        If KeyCode = 113 Or KeyCode = 114 Then
            
            FrmPesquisaSRCaixa.Campos1 = "CodBan"
            FrmPesquisaSRCaixa.Campos2 = "Nome"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 2
            FrmPesquisaSRCaixa.Tabela = "Ban"
            'FrmPesquisaSRCaixa.condi = "Select GENCodigo, GENDescricao, GENControle From PRD Where PRDControle = 0"
            
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Bancos"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        
        End If
    
    
    Case 14
        
        If KeyCode = 113 Then
            FrmPesquisaSRCaixa.Campos1 = "CodTit"
            FrmPesquisaSRCaixa.Campos2 = "Nome"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 3
            FrmPesquisaSRCaixa.Tabela = "Tit"
           
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Clientes/Titulares"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
            
        ElseIf KeyCode = 114 Then
            FrmPesquisaSRCaixa.Campos1 = "CodBan"
            FrmPesquisaSRCaixa.Campos2 = "Nome"
            FrmPesquisaSRCaixa.QtdCampos = 1
            FrmPesquisaSRCaixa.TipoResultado = 3
            FrmPesquisaSRCaixa.Tabela = "Ban"
           
            FrmPesquisaSRCaixa.Caption = "Pesquisa de Bancos"
            FrmPesquisaSRCaixa.Show vbModal
        
            MebDados(Index).Text = FrmPesquisaSRCaixa.ResultadoPesquisa
        End If
               
        'If KeyCode = 13 Then
        '    If Busca("CodBan", MebDados(Index), "Ban") = True Then
        '        TxtDados(5).Text = SeCampoNulo(Consulta, "Nome")
        '    End If
        'End If

End Select
End Sub

Private Sub TabCad_Click(previoustab As Integer)
If TabCad.Tab = 0 Then MebDados(4).SetFocus
If TabCad.Tab = 1 Then BtAcoes(0).SetFocus
If TabCad.Tab = 2 Then MebDados(11).SetFocus
End Sub

Private Sub TxtDados_LostFocus(Index As Integer)
Select Case Index
    Case 0
        If TxtDados(0).Locked = False Then
            If Not IsNumeric(ApenasNumeros(MebDados(0))) Then
                MsgBox "Cpf/Cnpj não numérico!"
                MebDados(0).SetFocus
                SelecionaCampo
                Exit Sub
            End If
        
            If TxtDados(0).Text = "" Then
                MsgBox "É necessário que digite o nome referente ao Cpf", vbInformation
                TxtDados(0).SetFocus
                SelecionaCampo
                Exit Sub
            End If
        
            'TxtDados(0).Locked = True
            BancoSRCaixa.Execute "Delete From Nom Where CgcCpf Like '" & MebDados(0).Text & "'"
            InsereBanco "Nom", "CgcCpf, Nome", "'" & ApenasNumeros(MebDados(0)) & "','" & TxtDados(0).Text & "'"
        End If
End Select

End Sub

Private Sub SomaCheques()
Dim SomaChq As Currency
SomaChq = 0

    For x = 1 To lvChqs3.ListItems.Count
        If lvChqs3.ListItems.Item(x).Checked = True Then
            If Not IsNumeric(lvChqs3.ListItems(x).ListSubItems(3).Text) Then
                SomaChq = SomaChq + CCur(lvChqs3.ListItems(x).ListSubItems(4).Text)
            Else
                SomaChq = SomaChq + CCur(lvChqs3.ListItems(x).ListSubItems(3).Text)
            End If
        End If
    Next
    
    lblTotal2.Caption = Format(SomaChq, "#,#0.00")
    
End Sub
