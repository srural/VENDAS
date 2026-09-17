VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Entidades"
   ClientHeight    =   5490
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   9270
   Icon            =   "FrmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9270
   Begin TabDlg.SSTab TabCad 
      Height          =   5250
      Left            =   120
      TabIndex        =   66
      Top             =   120
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   9260
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "FrmClientes.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtPesquisa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DBGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DatDados(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cadastro"
      TabPicture(1)   =   "FrmClientes.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDados(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "PctFundo(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtDados(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DatDados(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Documentos"
      TabPicture(2)   =   "FrmClientes.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PctFundo(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "PctFundo(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "PctFundo(4)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Gerenciais"
      TabPicture(3)   =   "FrmClientes.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "PctFundo(5)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Complemento"
      TabPicture(4)   =   "FrmClientes.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "PctFundo(1)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Comunicação"
      TabPicture(5)   =   "FrmClientes.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Observação"
      TabPicture(6)   =   "FrmClientes.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "TxtDados(40)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.TextBox TxtDados 
         DataField       =   "Obs"
         Height          =   4485
         Index           =   40
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Tag             =   "S"
         Top             =   480
         Width           =   8535
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   -74640
         ScaleHeight     =   4695
         ScaleWidth      =   8535
         TabIndex        =   109
         Top             =   360
         Width           =   8535
         Begin VB.TextBox TxtDados 
            DataField       =   "Email"
            Height          =   315
            Index           =   33
            Left            =   4395
            MaxLength       =   40
            TabIndex        =   63
            Tag             =   "S"
            Top             =   4320
            Width           =   3420
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "HomePage"
            Height          =   315
            Index           =   32
            Left            =   360
            MaxLength       =   40
            TabIndex        =   62
            Tag             =   "S"
            Top             =   4320
            Width           =   3900
         End
         Begin VB.Frame Frame2 
            Caption         =   "Endereço de Entrega"
            Height          =   1575
            Left            =   120
            TabIndex        =   110
            Top             =   120
            Width           =   7935
            Begin VB.TextBox TxtDados 
               DataField       =   "Endereço_Entrega"
               Height          =   315
               Index           =   13
               Left            =   240
               MaxLength       =   40
               TabIndex        =   47
               Tag             =   "S"
               Top             =   480
               Width           =   5895
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Nro_Entrega"
               Height          =   315
               Index           =   12
               Left            =   6240
               MaxLength       =   40
               TabIndex        =   48
               Tag             =   "S"
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Bairro_Entrega"
               Height          =   315
               Index           =   11
               Left            =   240
               MaxLength       =   40
               TabIndex        =   49
               Tag             =   "S"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Cidade_Entrega"
               Height          =   315
               Index           =   10
               Left            =   2160
               MaxLength       =   40
               TabIndex        =   50
               Tag             =   "S"
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Uf_Entrega"
               Height          =   315
               Index           =   9
               Left            =   5280
               MaxLength       =   40
               TabIndex        =   51
               Tag             =   "S"
               Top             =   1080
               Width           =   495
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Entrega"
               Height          =   315
               Index           =   8
               Left            =   5880
               TabIndex        =   52
               Top             =   1080
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Index           =   23
               Left            =   240
               TabIndex        =   116
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Numero:"
               Height          =   195
               Index           =   22
               Left            =   6240
               TabIndex        =   115
               Top             =   240
               Width           =   600
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Index           =   21
               Left            =   240
               TabIndex        =   114
               Top             =   840
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Index           =   20
               Left            =   2160
               TabIndex        =   113
               Top             =   840
               Width           =   540
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Index           =   19
               Left            =   5280
               TabIndex        =   112
               Top             =   840
               Width           =   255
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Index           =   18
               Left            =   5880
               TabIndex        =   111
               Top             =   840
               Width           =   315
            End
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone2"
            Height          =   315
            Index           =   7
            Left            =   360
            TabIndex        =   59
            Top             =   3720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fax2"
            Height          =   315
            Index           =   19
            Left            =   2880
            TabIndex        =   60
            Top             =   3720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Celular2"
            Height          =   315
            Index           =   20
            Left            =   5640
            TabIndex        =   61
            Top             =   3720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame1 
            Caption         =   "Endereço de Cobrança"
            Height          =   1575
            Left            =   120
            TabIndex        =   117
            Top             =   1800
            Width           =   7935
            Begin VB.TextBox TxtDados 
               DataField       =   "Uf_Cob"
               Height          =   315
               Index           =   8
               Left            =   5280
               MaxLength       =   40
               TabIndex        =   57
               Tag             =   "S"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Cidade_Cob"
               Height          =   315
               Index           =   28
               Left            =   2160
               MaxLength       =   40
               TabIndex        =   56
               Tag             =   "S"
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Bairro_Cob"
               Height          =   315
               Index           =   29
               Left            =   240
               MaxLength       =   40
               TabIndex        =   55
               Tag             =   "S"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Nro_Cob"
               Height          =   315
               Index           =   30
               Left            =   6240
               MaxLength       =   40
               TabIndex        =   54
               Tag             =   "S"
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtDados 
               DataField       =   "Endereço_Cob"
               Height          =   315
               Index           =   31
               Left            =   240
               MaxLength       =   40
               TabIndex        =   53
               Tag             =   "S"
               Top             =   480
               Width           =   5895
            End
            Begin MSMask.MaskEdBox MebDados 
               DataField       =   "Cep_Cob"
               Height          =   315
               Index           =   5
               Left            =   5880
               TabIndex        =   58
               Top             =   1080
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Index           =   12
               Left            =   5880
               TabIndex        =   123
               Top             =   840
               Width           =   315
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Index           =   15
               Left            =   5280
               TabIndex        =   122
               Top             =   840
               Width           =   255
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Index           =   17
               Left            =   2160
               TabIndex        =   121
               Top             =   840
               Width           =   540
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Index           =   48
               Left            =   240
               TabIndex        =   120
               Top             =   840
               Width           =   450
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Numero:"
               Height          =   195
               Index           =   49
               Left            =   6240
               TabIndex        =   119
               Top             =   240
               Width           =   600
            End
            Begin VB.Label LblDados 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Index           =   50
               Left            =   240
               TabIndex        =   118
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
            Height          =   195
            Index           =   55
            Left            =   4440
            TabIndex        =   128
            Top             =   4080
            Width           =   465
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Home P."
            Height          =   195
            Index           =   54
            Left            =   360
            TabIndex        =   127
            Top             =   4080
            Width           =   615
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Celular 2:"
            Height          =   195
            Index           =   53
            Left            =   5640
            TabIndex        =   126
            Top             =   3480
            Width           =   660
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fone 2:"
            Height          =   195
            Index           =   52
            Left            =   360
            TabIndex        =   125
            Top             =   3480
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fax 2:"
            Height          =   195
            Index           =   51
            Left            =   2880
            TabIndex        =   124
            Top             =   3480
            Width           =   435
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   3975
         Index           =   5
         Left            =   -74520
         ScaleHeight     =   3975
         ScaleWidth      =   7695
         TabIndex        =   106
         Top             =   720
         Width           =   7695
         Begin VB.TextBox TxtDados 
            DataField       =   "Comissao"
            Height          =   315
            Index           =   35
            Left            =   4680
            MaxLength       =   40
            TabIndex        =   36
            Tag             =   "S"
            Top             =   3360
            Width           =   2775
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Refer3"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   24
            Left            =   360
            MaxLength       =   40
            TabIndex        =   35
            Tag             =   "S"
            Top             =   2760
            Width           =   7095
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Refer2"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   23
            Left            =   360
            MaxLength       =   40
            TabIndex        =   34
            Tag             =   "S"
            Top             =   2160
            Width           =   7095
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Refer1"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   22
            Left            =   360
            MaxLength       =   40
            TabIndex        =   33
            Tag             =   "S"
            Top             =   1560
            Width           =   7095
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   32
            Top             =   960
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "Status"
            Height          =   315
            Index           =   3
            ItemData        =   "FrmClientes.frx":00D0
            Left            =   1680
            List            =   "FrmClientes.frx":00E0
            TabIndex        =   27
            Tag             =   "SI"
            Top             =   360
            Width           =   1200
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "Condicao"
            Height          =   315
            Index           =   2
            ItemData        =   "FrmClientes.frx":010A
            Left            =   360
            List            =   "FrmClientes.frx":011A
            TabIndex        =   26
            Tag             =   "SI"
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Prazo"
            Height          =   315
            Index           =   7
            Left            =   3000
            MaxLength       =   40
            TabIndex        =   28
            Tag             =   "S"
            Top             =   360
            Width           =   1200
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Credito"
            Height          =   315
            Index           =   6
            Left            =   4320
            TabIndex        =   29
            Top             =   360
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Juros"
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Desconto"
            Height          =   315
            Index           =   22
            Left            =   1680
            TabIndex        =   31
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSDBCtls.DBCombo DbcDados 
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   141
            Top             =   3360
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comissões:"
            Height          =   195
            Index           =   60
            Left            =   4680
            TabIndex        =   138
            Top             =   3120
            Width           =   795
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Convênio"
            Height          =   195
            Index           =   59
            Left            =   405
            TabIndex        =   137
            Top             =   3120
            Width           =   675
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Referência 3"
            Height          =   195
            Index           =   34
            Left            =   360
            TabIndex        =   136
            Top             =   2520
            Width           =   915
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Referência 2"
            Height          =   195
            Index           =   33
            Left            =   360
            TabIndex        =   135
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Referência 1"
            Height          =   195
            Index           =   32
            Left            =   360
            TabIndex        =   134
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Transp."
            Height          =   195
            Index           =   16
            Left            =   2880
            TabIndex        =   133
            Top             =   720
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descontos:"
            Height          =   195
            Index           =   58
            Left            =   1680
            TabIndex        =   132
            Top             =   720
            Width           =   810
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   131
            Top             =   720
            Width           =   420
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Status:"
            Height          =   195
            Index           =   57
            Left            =   1680
            TabIndex        =   130
            Top             =   120
            Width           =   495
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Condição:"
            Height          =   195
            Index           =   56
            Left            =   360
            TabIndex        =   129
            Top             =   120
            Width           =   720
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Crédito:"
            Height          =   195
            Index           =   13
            Left            =   4320
            TabIndex        =   108
            Top             =   120
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   195
            Index           =   14
            Left            =   3000
            TabIndex        =   107
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   4
         Left            =   -74400
         ScaleHeight     =   1935
         ScaleWidth      =   8175
         TabIndex        =   97
         Top             =   2460
         Width           =   8175
         Begin VB.TextBox TxtDados 
            DataField       =   "Ibama"
            Height          =   315
            Index           =   27
            Left            =   480
            MaxLength       =   40
            TabIndex        =   24
            Tag             =   "S"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Cri"
            Height          =   315
            Index           =   26
            Left            =   2640
            MaxLength       =   40
            TabIndex        =   25
            Tag             =   "S"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Area"
            Height          =   315
            Index           =   19
            Left            =   4680
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "S"
            Top             =   840
            Width           =   2055
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "InscrProd"
            Height          =   315
            Index           =   14
            Left            =   480
            TabIndex        =   21
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   17
            Mask            =   "P-####-####.#/###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cnae"
            Height          =   315
            Index           =   15
            Left            =   2640
            TabIndex        =   22
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "####-#/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Incra"
            Height          =   315
            Index           =   16
            Left            =   480
            TabIndex        =   18
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            Mask            =   "###.###.###.###-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Itr"
            Height          =   315
            Index           =   17
            Left            =   2640
            TabIndex        =   19
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#######-#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cei"
            Height          =   315
            Index           =   18
            Left            =   4680
            TabIndex        =   20
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "##.###.#####/##"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cri"
            Height          =   195
            Index           =   47
            Left            =   2640
            TabIndex        =   105
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ibama"
            Height          =   195
            Index           =   46
            Left            =   480
            TabIndex        =   104
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Área"
            Height          =   195
            Index           =   45
            Left            =   4680
            TabIndex        =   103
            Top             =   600
            Width           =   330
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Incra:"
            Height          =   195
            Index           =   44
            Left            =   480
            TabIndex        =   102
            Top             =   0
            Width           =   405
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ITR."
            Height          =   195
            Index           =   43
            Left            =   2640
            TabIndex        =   101
            Top             =   0
            Width           =   315
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CEI."
            Height          =   195
            Index           =   42
            Left            =   4680
            TabIndex        =   100
            Top             =   0
            Width           =   300
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cnae"
            Height          =   195
            Index           =   41
            Left            =   2640
            TabIndex        =   99
            Top             =   600
            Width           =   375
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Incr. Prod."
            Height          =   195
            Index           =   40
            Left            =   480
            TabIndex        =   98
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   3
         Left            =   -74040
         ScaleHeight     =   855
         ScaleWidth      =   3615
         TabIndex        =   94
         Top             =   1560
         Width           =   3615
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CGC"
            Height          =   315
            Index           =   12
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "InscrEst"
            Height          =   315
            Index           =   13
            Left            =   1800
            TabIndex        =   17
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   17
            Mask            =   "##.###.###/###-##"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cgc:"
            Height          =   195
            Index           =   39
            Left            =   75
            TabIndex        =   96
            Top             =   120
            Width           =   330
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscr."
            Height          =   195
            Index           =   38
            Left            =   1800
            TabIndex        =   95
            Top             =   120
            Width           =   390
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   2
         Left            =   -74040
         ScaleHeight     =   855
         ScaleWidth      =   3735
         TabIndex        =   91
         Top             =   840
         Width           =   3735
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "RG"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "CPF"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   11
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
            Height          =   195
            Index           =   37
            Left            =   1800
            TabIndex        =   93
            Top             =   120
            Width           =   345
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "RG:"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   92
            Top             =   120
            Width           =   285
         End
      End
      Begin VB.PictureBox PctFundo 
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   1
         Left            =   -74040
         ScaleHeight     =   4215
         ScaleWidth      =   7215
         TabIndex        =   80
         Top             =   480
         Width           =   7215
         Begin VB.TextBox TxtDados 
            DataField       =   "TimeTorce"
            Height          =   315
            Index           =   25
            Left            =   4680
            MaxLength       =   40
            TabIndex        =   46
            Tag             =   "S"
            Top             =   3360
            Width           =   2040
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Uf_Trabalho"
            Height          =   315
            Index           =   20
            Left            =   2520
            MaxLength       =   40
            TabIndex        =   44
            Tag             =   "S"
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Endereco_Trabalho"
            Height          =   315
            Index           =   21
            Left            =   120
            MaxLength       =   40
            TabIndex        =   42
            Tag             =   "S"
            Top             =   2760
            Width           =   6615
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Cidade_Trabalho"
            Height          =   315
            Index           =   18
            Left            =   120
            MaxLength       =   40
            TabIndex        =   43
            Tag             =   "S"
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Trabalho"
            Height          =   315
            Index           =   17
            Left            =   2400
            MaxLength       =   40
            TabIndex        =   41
            Tag             =   "S"
            Top             =   2160
            Width           =   4335
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Profissao"
            Height          =   315
            Index           =   16
            Left            =   120
            MaxLength       =   40
            TabIndex        =   40
            Tag             =   "S"
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Filiacao"
            Height          =   315
            Index           =   15
            Left            =   120
            MaxLength       =   40
            TabIndex        =   39
            Tag             =   "S"
            Top             =   1560
            Width           =   6615
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Conjuge"
            Height          =   315
            Index           =   14
            Left            =   120
            MaxLength       =   40
            TabIndex        =   38
            Tag             =   "S"
            Top             =   960
            Width           =   6615
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "EstadoCivil"
            Height          =   315
            Index           =   1
            ItemData        =   "FrmClientes.frx":0150
            Left            =   120
            List            =   "FrmClientes.frx":0163
            TabIndex        =   37
            Tag             =   "S"
            Top             =   360
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "DtCasamento"
            Height          =   315
            Index           =   9
            Left            =   3120
            TabIndex        =   45
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data de Casamento"
            Height          =   195
            Index           =   36
            Left            =   3120
            TabIndex        =   90
            Top             =   3120
            Width           =   1410
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Time que Torce"
            Height          =   195
            Index           =   35
            Left            =   4680
            TabIndex        =   89
            Top             =   3120
            Width           =   1125
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Index           =   30
            Left            =   2520
            TabIndex        =   88
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   87
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   28
            Left            =   120
            TabIndex        =   86
            Top             =   3120
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Trababalho:"
            Height          =   195
            Index           =   27
            Left            =   2400
            TabIndex        =   85
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Profissão"
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   84
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Filiação:"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   83
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conjunge:"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "Estado Civil:"
            Height          =   255
            Left            =   135
            TabIndex        =   81
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Data DatDados 
         Connect         =   "Access"
         DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -67845
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ENT"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data DatDados 
         Caption         =   "Titulares"
         Connect         =   "Access"
         DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Index           =   0
         Left            =   6705
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ENT"
         Top             =   2715
         Visible         =   0   'False
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FrmClientes.frx":0194
         Height          =   3525
         Left            =   360
         OleObjectBlob   =   "FrmClientes.frx":01AE
         TabIndex        =   65
         Top             =   1440
         Width           =   8445
      End
      Begin VB.TextBox TxtPesquisa 
         Height          =   330
         Left            =   360
         TabIndex        =   13
         Top             =   990
         Width           =   5415
      End
      Begin VB.TextBox TxtDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "CodEntidade"
         DataSource      =   "DatDados(1)"
         Height          =   315
         Index           =   0
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   142
         Tag             =   "N"
         Top             =   840
         Width           =   975
      End
      Begin VB.PictureBox PctFundo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3240
         Index           =   0
         Left            =   -74520
         ScaleHeight     =   3240
         ScaleWidth      =   7920
         TabIndex        =   67
         Top             =   1200
         Width           =   7920
         Begin VB.ComboBox CmbDados 
            DataField       =   "Classificacao"
            Height          =   315
            Index           =   4
            ItemData        =   "FrmClientes.frx":0B95
            Left            =   2160
            List            =   "FrmClientes.frx":0BA2
            TabIndex        =   1
            Tag             =   "SI"
            Top             =   240
            Width           =   1200
         End
         Begin VB.ComboBox CmbDados 
            DataField       =   "Tipo"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmClientes.frx":0BC2
            Left            =   240
            List            =   "FrmClientes.frx":0BD8
            TabIndex        =   0
            Tag             =   "SI"
            Top             =   240
            Width           =   1800
         End
         Begin VB.CheckBox ChkDados 
            Caption         =   "Ativo"
            DataField       =   "Ativo"
            Height          =   255
            Index           =   0
            Left            =   6720
            TabIndex        =   3
            Tag             =   "S"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Uf"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   6
            Left            =   4080
            MaxLength       =   40
            TabIndex        =   8
            Tag             =   "S"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Cidade"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   5
            Left            =   1680
            MaxLength       =   40
            TabIndex        =   7
            Tag             =   "S"
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Bairro"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   4
            Left            =   240
            MaxLength       =   40
            TabIndex        =   6
            Tag             =   "S"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Nro"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   3
            Left            =   6360
            MaxLength       =   40
            TabIndex        =   5
            Tag             =   "S"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Endereco"
            DataSource      =   "DatDados(1)"
            Height          =   315
            Index           =   2
            Left            =   240
            MaxLength       =   40
            TabIndex        =   4
            Tag             =   "S"
            Top             =   1440
            Width           =   6015
         End
         Begin VB.TextBox TxtDados 
            DataField       =   "Nome"
            Height          =   315
            Index           =   1
            Left            =   240
            MaxLength       =   40
            TabIndex        =   2
            Tag             =   "S"
            Top             =   840
            Width           =   6255
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Cep"
            Height          =   315
            Index           =   1
            Left            =   4680
            TabIndex        =   9
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##-###-###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fone"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Top             =   2640
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Fax"
            Height          =   315
            Index           =   3
            Left            =   2400
            TabIndex        =   12
            Top             =   2640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MebDados 
            DataField       =   "Celular"
            Height          =   315
            Index           =   4
            Left            =   5880
            TabIndex        =   10
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "(###)-####-####"
            PromptChar      =   "_"
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Classificação"
            Height          =   195
            Index           =   62
            Left            =   2160
            TabIndex        =   140
            Top             =   0
            Width           =   930
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   61
            Left            =   240
            TabIndex        =   139
            Top             =   0
            Width           =   360
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Index           =   11
            Left            =   5880
            TabIndex        =   79
            Top             =   1800
            Width           =   525
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Index           =   10
            Left            =   2400
            TabIndex        =   78
            Top             =   2400
            Width           =   300
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fone:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   77
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Index           =   5
            Left            =   4680
            TabIndex        =   76
            Top             =   1800
            Width           =   315
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Index           =   4
            Left            =   4080
            TabIndex        =   75
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   9
            Left            =   1680
            TabIndex        =   74
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   73
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Numero:"
            Height          =   195
            Index           =   3
            Left            =   6360
            TabIndex        =   72
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   71
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label LblDados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   68
            Top             =   600
            Width           =   465
         End
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   -74280
         TabIndex        =   70
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Nome da Entidade"
         Height          =   240
         Left            =   360
         TabIndex        =   69
         Top             =   765
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset   'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Dim FlagAbertura As Boolean
Dim Cliente As Integer

Public CTextBox As Boolean      'Caixa de Texto
Public CMaskEditBox As Boolean  'Mascara
Public CComboBox As Boolean     'Combo
Public CListBox As Boolean      'Lista
Public CdbCombo As Boolean      'DBCombo
Public CDbList As Boolean       'DBLista
Public CCheckBox As Boolean     'Caixa de Checagem
Public CDbgrid As Boolean       'Grade
Public ComandBotom As Boolean   'Botão

Public SelectFromSQL  As String  ' Representa o Conteudo das Cláusulas Select e From do Sql Utilizado p/ acessar os registros deste formulário
Public WhereSQL As String        ' Representa o Conteudo da Cláusula Where do Sql Utilizado p/ acessar os registros deste formulário
Public FiltroSQL As String       ' Critério imposto pelo usuário
Public OrderBySQL As String      ' Representa a cláusula Order By do Sql do formulário Ativo
Public PesquisaSql As String     ' Representa a pesquisa F2
Public SelectFromFiltro As String  ' Representa a relação de campos que sera exibida na visualização no formulário de filtragem

Public LocalizacaoBasica As String ' Armazena o nome do campo que sera utilizado na ferramenta de localizaco basica

Private Sub DefineObjetos()
'Define quais objeto foram utilizados neste formulário

CTextBox = True         ' TxtDados
CMaskEditBox = True     ' MebDados
CComboBox = True        ' CmbDados
CListBox = False
CdbCombo = False        ' DbcDados
CDbList = False         '
CCheckBox = True        ' ChkDados
CDbgrid = False

End Sub

Private Sub DBGDados_BeforeInsert(Index As Integer, Cancel As Integer)
' antes de comercar incluir é armazenado o numero do pedido que será salvo
' Titular = Registros.Fields("CodTit")
End Sub

Private Sub DBGrid1_DBLClick()
Dim Criterio As String
Criterio = "CodEntidade = " & DBGrid1.Columns(0)
Registros.FindFirst Criterio
TabCad.Tab = 1 'direcina para o TAB seguinte

MostraDados
PosMovimentacao

End Sub

Private Sub Form_Activate()
'  Ocorre toda vez que o formulario é selecionado
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

AtivarMenu (True)
'Formulário.Left = 700
'Formulário.Top = 0

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

GerenteBarra IIf(Registros.EditMode = dbEditNone, False, True)

' testa se o banco está em modo de consulta ou não
If FlagAbertura = False Then
    FlagAbertura = True
    Inicio
End If
TabCad.Tab = 0
AbilitaBotoes (False)
TxtPesquisa.SetFocus

'PosMovimentação
' DbcPesquisa.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub

Private Sub Form_Load()
Dim Componentes As Variant
' O Evento Load Ocorre toda vez quando o formulario está sendo
' carregado p/ a memória
' Set Registros = Banco.OpenRecordset("SELECT * FROM CLI INNER JOIN ATU ON CLI.CODCLI = ATU.CODCLI ORDER BY CLI.NOME")
    
Set Registros = Banco.OpenRecordset("SELECT * From Ent Order by Nome")

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next

DatDados(0).RecordSource = "Select * From Ent Order by Nome"

LocalizacaoBasica = "Nome"
DefineObjetos
SelectFromSQL = "Select * From Ent"
SelectFromFiltro = "Select CodEntidade, Nome, Cidade, Uf from Ent"
PesquisaSql = "SELECT CodEntidade , Nome FROM Ent ORDER BY Nome"
FlagAbertura = False

OrderBySQL = "[Ent].[Nome]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' O Evento unLoad ocorre quando o formulario está sendo
' descarregado p/ a memória
If Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
   Confirmacoes (14)
End If

Registros.Close

AtivarMenu (False)

End Sub

Public Sub PosInclusao()
' sera executada logo depois que o usuario iniciar o processo de inclusao addnew

TxtDados(7) = 0
MebDados(6) = 0
TxtDados(19) = 0
TxtDados(35) = 0
MebDados(0) = 0
MebDados(22) = 0

CmbDados(0).SetFocus
TabCad.TabEnabled(0) = False

End Sub

Public Function PermiteInclusao() As Boolean
' Será executado antes do processo de inclusao (addnew)
' Aqui deve ser feitos os processo p/ que o usuário pode ou não incluir um novo registro
' MsgBox "Neste formulário não pode ser feitas Inclusão de novos registros", vbInformation, App.Title

PermiteInclusao = True
PosMovimentacao

End Function

Public Sub PosAlteracao()
' sera executada logo depois que o usuario iniciar o processo de alteracao
'TxtDados(0).BackColor = QBColor(8) 'Muda a Cor de fundo do código
CmbDados(0).SetFocus
TabCad.TabEnabled(0) = False
TabCad.TabEnabled(2) = True
'Atualizar
End Sub

Public Function PermiteAlteracao() As Boolean
' Será executado antes do processo de Alteracao

PermiteAlteracao = True

End Function

Public Sub PosMovimentacao()

End Sub

Public Sub PosExclusao()
' sera executada logo depois que o usuario iniciar o processo de exclusao

Inicio
Atualizar
End Sub

Public Function PermiteExclusao() As Boolean
' Será executado antes do processo de exclusao
Dim Mensagem As String
Mensagem = "Confirma Exclusão do Titular " & vbCr & Registros.Fields("Nome_Titular")

If MsgBox(Mensagem, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
' nesta posicao do codigo podemos incluir os procedimentos que deverao
' apagar todos os registro que dependem do registro que está sendo excluido, no caso de relacionamento entre tabelas
' ex : poderiamos excluir os itens de um pedido antes de excluir o cabeçalho
    PermiteExclusao = True
Else
    PermiteExclusao = False
End If

End Function

Public Sub PreConfirmacao() 'Será executada antes das rotinas de salvar ou cancelar

End Sub

Public Sub PosConfirmacao() 'Será executada deposi dos processos de salvar ou cancelar

'TxtDados(0).BackColor = QBColor(7) 'Muda a Cor de fundo do código
TxtDados(0).Locked = False
Atualizar
End Sub

Public Function PermiteCancelamento() As Boolean
' Será executada antes do processo de cancelamento ser executado (cancelupdate)
' podemos aqui impedir o usuario de cancelar uma inclusao ou uma alteracao

PermiteCancelamento = True

End Function

Public Function PermiteSalvar() As Boolean
' Será axecutada antes de atribuir os campos do formulário no bco de dados
' Podemos fazer qualquer tipo de consistencia que não possa ser feita pelo bco de  dados
' Ex: O estoque máximo sempre deverá ser maior que o estoque minimo

PermiteSalvar = True

End Function
Public Function PermiteContinuarMostrando() As Boolean

End Function

Public Function PermiteContinuarSalvar() As Boolean
' Será executada logo depois de atribuir os campos ao bco de dados
' Possui as mesma função do permitir salvar, Inclusive podendo atribuir valores a campos do bco de dados que não aparecem na tela
' Ex: Aqui vc pode atribuir um valor ao bco que não está no formulario
PermiteContinuarSalvar = True

End Function

Public Function PosSalvamento() As Boolean
'Será executado assim que todos os dados forem gravados no bco

PosSalvamento = True

TabCad.TabEnabled(0) = True
TabCad.TabEnabled(2) = True

End Function

Public Sub PosCancelamento()

'sera executado logo após do processo de cancelamento da aplicaçaõ

TabCad.TabEnabled(0) = True
TabCad.TabEnabled(2) = True

End Sub

Public Sub Atualizar()
'Registros.Requery
DatDados(0).Refresh
DatDados(1).Refresh
End Sub

Private Sub MebDados_LostFocus(Index As Integer)
Select Case Index
    Case 3
        TabCad.Tab = 2
End Select
End Sub

Private Sub Opt1_Click(Index As Integer)

'Ajustar Mascaras

Select Case Index
    Case 0
        MebDados(5).Mask = "##.###.###/####/##"
    Case 1
        MebDados(5).Mask = "###.###.###-##"
End Select

End Sub


Private Sub TabCad_Click(previoustab As Integer)
    If PctFundo(0).Enabled = True Then
        If CmbDados(4).Text = "FISICA" Then
            PctFundo(2).Enabled = True
            PctFundo(3).Enabled = False
            PctFundo(4).Enabled = False
        End If
        If CmbDados(4).Text = "JURIDICA" Then
            PctFundo(2).Enabled = False
            PctFundo(3).Enabled = True
            PctFundo(4).Enabled = False
        End If
        If CmbDados(4).Text = "PRODUTOR" Then
            PctFundo(2).Enabled = False
            PctFundo(3).Enabled = False
            PctFundo(4).Enabled = True
        End If
    End If

Select Case TabCad.Tab
    Case 0
        Atualizar
        AbilitaBotoes (False)
    Case 1
        AbilitaBotoes (True)
    Case 2
        AbilitaBotoes (True)
End Select

End Sub

Private Sub Txtdados_Change(Index As Integer)
Dim SomaGrid As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Pago = 0
Apagar = 0
If Index = 0 Then ' caso tenha ocorrido alguma alteração do numero do pedido
    If IsNumeric(TxtDados(0)) = False Then
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = 0 ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=0)) ORDER BY Bol.Data"

    Else
'        DatDados(2).RecordSource = "SELECT Trn.ENDERECO_LOTE, Trn.NRO_LOTE, Trn.BAIRRO_LOTE FROM For INNER JOIN Trn ON For.CODFOR = Trn.FOREIRO WHERE Trn.Foreiro = " & TxtDados(0) & " ORDER BY Trn.ENDERECO_LOTE"
'        DatDados(3).RecordSource = "SELECT Bol.TERRENO, Bol.NROPCL, Bol.CODREF, Bol.Data, Bol.VENC, Bol.PGT, Bol.VALOR, Bol.COND, Bol.Foreiro FROM For INNER JOIN Bol ON For.CODFOR = Bol.Foreiro WHERE (((Bol.Foreiro)=" & TxtDados(0) & " )) ORDER BY Bol.Data"
    End If
'    DatDados(2).Refresh
'    DatDados(3).Refresh
End If

End Sub

Private Sub TxtDados_LostFocus(Index As Integer)
Select Case Index
    Case 26
        TabCad.Tab = 3
    Case 35
        TabCad.Tab = 4
    Case 25
        TabCad.Tab = 5
    Case 33
        TabCad.Tab = 6
   
End Select
End Sub

Private Sub TxtPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If TxtPesquisa <> "" Then
    If IsNumeric(TxtPesquisa) Then
        Criterio = DatDados(0).Recordset.Fields(0).Name & " =  " & TxtPesquisa
    Else
        Criterio = DatDados(0).Recordset.Fields(1).Name & " LIKE " & """" & TxtPesquisa & "*"""
    End If
    
    DatDados(0).Recordset.FindFirst Criterio
    If DatDados(0).Recordset.NoMatch Then
        TxtPesquisa.SetFocus
    Else
        Registros.FindFirst Criterio
        MostraDados
        PosMovimentacao
    End If
End If


End Sub

Private Sub TxtPesquisa_LostFocus()
    DBGrid1_DBLClick
End Sub
