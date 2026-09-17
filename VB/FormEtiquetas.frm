VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form FormEtiquetas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   13935
   Begin VB.CommandButton BtEtq 
      Caption         =   "Promoção"
      Height          =   375
      Index           =   21
      Left            =   12780
      TabIndex        =   98
      Top             =   6210
      Width           =   930
   End
   Begin VB.CommandButton BtEtq 
      Caption         =   "Uma Coluna"
      Height          =   330
      Index           =   20
      Left            =   12420
      TabIndex        =   97
      Top             =   5805
      Width           =   1290
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   10620
      TabIndex        =   94
      Top             =   4905
      Width           =   2535
      Begin VB.CommandButton BtEtq 
         Caption         =   "Rio"
         Height          =   330
         Index           =   19
         Left            =   1395
         TabIndex        =   96
         Top             =   270
         Width           =   960
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Rio Kids"
         Height          =   330
         Index           =   18
         Left            =   180
         TabIndex        =   95
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.CommandButton CmdLista 
      Caption         =   "Listagem Produtos"
      Height          =   330
      Left            =   10530
      TabIndex        =   91
      Top             =   5805
      Width           =   1725
   End
   Begin VB.Data Datdados 
      Caption         =   "pesquisa"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   4
      Left            =   12105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormEtiquetas.frx":0000
      Top             =   3915
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton BtEtq 
      Caption         =   "Colombo"
      Height          =   375
      Index           =   17
      Left            =   11760
      TabIndex        =   88
      Top             =   6210
      Width           =   930
   End
   Begin VB.CommandButton BtEtq 
      Caption         =   "Borela"
      Height          =   375
      Index           =   16
      Left            =   10605
      TabIndex        =   87
      Top             =   6210
      Width           =   1080
   End
   Begin VB.CheckBox chkcadastro 
      Caption         =   "Cadastradas Hoje"
      Height          =   285
      Left            =   8280
      TabIndex        =   83
      Top             =   1125
      Width           =   1710
   End
   Begin VB.TextBox txtlote 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   12150
      TabIndex        =   82
      Top             =   6705
      Width           =   1455
   End
   Begin VB.CommandButton CmdTinta 
      Caption         =   "Tinta - Lote >>"
      Height          =   330
      Left            =   10485
      TabIndex        =   81
      Top             =   6705
      Width           =   1455
   End
   Begin VB.CommandButton BtEtq 
      Caption         =   "Dani"
      Height          =   330
      Index           =   13
      Left            =   6345
      TabIndex        =   80
      Top             =   6705
      Width           =   1320
   End
   Begin VB.CommandButton Acoff 
      Caption         =   "Acoff"
      Height          =   330
      Left            =   7875
      TabIndex        =   79
      Top             =   6705
      Width           =   960
   End
   Begin VB.CommandButton BtBalanca 
      Caption         =   "Balança"
      Height          =   330
      Left            =   10215
      TabIndex        =   78
      Top             =   1890
      Width           =   960
   End
   Begin VB.CommandButton BtPed 
      Caption         =   ">>"
      Height          =   330
      Left            =   11205
      TabIndex        =   76
      Top             =   1395
      Width           =   330
   End
   Begin VB.CommandButton CmdFicha 
      Caption         =   "Imprimir Ficha"
      Height          =   330
      Left            =   4815
      TabIndex        =   72
      Top             =   6705
      Width           =   1365
   End
   Begin VB.CommandButton EtqMercado 
      Caption         =   "Mercado"
      Height          =   330
      Left            =   9000
      TabIndex        =   71
      Top             =   6705
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   14475
      TabIndex        =   70
      Top             =   8820
      Width           =   14505
   End
   Begin VB.CheckBox ChkValor 
      Caption         =   "Imprimi Preço"
      Height          =   315
      Left            =   3240
      TabIndex        =   65
      Top             =   6750
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox ChkPreco 
      Caption         =   "Preço a vista"
      Height          =   315
      Left            =   1800
      TabIndex        =   62
      Top             =   6750
      Value           =   1  'Checked
      Width           =   1320
   End
   Begin VB.CheckBox chkprazo 
      Caption         =   "Preço a Prazo"
      Height          =   240
      Left            =   315
      TabIndex        =   60
      Top             =   6795
      Width           =   1410
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13920
      TabIndex        =   55
      Top             =   0
      Width           =   13920
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão de Etiquetas de Produto"
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
         TabIndex        =   56
         Top             =   45
         Width           =   6450
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2430
      Left            =   10485
      TabIndex        =   28
      Top             =   2280
      Width           =   3120
      Begin VB.ComboBox cmbModelo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   390
         Width           =   2520
      End
      Begin VB.TextBox txtAltura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   38
         Text            =   "25"
         Top             =   1995
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtLargura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   37
         Text            =   "95"
         Top             =   2355
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtMSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1755
         TabIndex        =   36
         Text            =   "17"
         Top             =   2745
         Width           =   885
      End
      Begin VB.TextBox txtMInf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   35
         Text            =   "5"
         Top             =   3075
         Width           =   885
      End
      Begin VB.TextBox txtMEsq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   34
         Text            =   "7"
         Top             =   3435
         Width           =   885
      End
      Begin VB.TextBox txtColunas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   33
         Text            =   "2"
         Top             =   3795
         Width           =   885
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   32
         Top             =   1995
         Width           =   945
      End
      Begin VB.CheckBox chkVisualizar 
         Caption         =   "Visualizar Impressão"
         Height          =   270
         Left            =   195
         TabIndex        =   31
         Top             =   1635
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.TextBox txtLinIni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         TabIndex        =   30
         Text            =   "1"
         Top             =   1140
         Width           =   465
      End
      Begin VB.TextBox txtColIni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1875
         TabIndex        =   29
         Text            =   "1"
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   54
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Altura"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   53
         Top             =   2055
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Index           =   2
         Left            =   2745
         TabIndex        =   52
         Top             =   2055
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Largura"
         Height          =   195
         Index           =   3
         Left            =   1140
         TabIndex        =   51
         Top             =   2415
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Index           =   4
         Left            =   2745
         TabIndex        =   50
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margem Sup."
         Height          =   195
         Index           =   5
         Left            =   735
         TabIndex        =   49
         Top             =   2775
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Index           =   6
         Left            =   2745
         TabIndex        =   48
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margem Inf."
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   47
         Top             =   3135
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Index           =   8
         Left            =   2745
         TabIndex        =   46
         Top             =   3135
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margem Esq."
         Height          =   195
         Index           =   9
         Left            =   750
         TabIndex        =   45
         Top             =   3495
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Index           =   10
         Left            =   2745
         TabIndex        =   44
         Top             =   3495
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Etiquetas por Linha"
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   43
         Top             =   3900
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Iniciar impressão na Etiqueta da"
         Height          =   195
         Index           =   12
         Left            =   195
         TabIndex        =   42
         Top             =   840
         Width           =   2250
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Linha"
         Height          =   195
         Index           =   13
         Left            =   195
         TabIndex        =   41
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Coluna"
         Height          =   195
         Index           =   14
         Left            =   1275
         TabIndex        =   40
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.CheckBox ChkEntradas 
      Caption         =   "Entradas de Hoje"
      Height          =   285
      Left            =   8280
      TabIndex        =   27
      Top             =   720
      Width           =   1710
   End
   Begin VB.CheckBox ChkDesconto 
      Caption         =   "&Desconto"
      Height          =   285
      Left            =   8295
      TabIndex        =   24
      Top             =   1890
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impressora Argox / Elgin-L42"
      Height          =   1185
      Left            =   270
      TabIndex        =   20
      Top             =   7155
      Width           =   13560
      Begin VB.CommandButton BtEtq 
         Caption         =   "Bijoteria"
         Height          =   375
         Index           =   15
         Left            =   4200
         TabIndex        =   86
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Bijoteria Zebra"
         Height          =   375
         Index           =   14
         Left            =   11160
         TabIndex        =   85
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta 3 Argox"
         Height          =   375
         Index           =   12
         Left            =   8865
         TabIndex        =   77
         Top             =   675
         Width           =   1500
      End
      Begin VB.CommandButton BtEtqTanca 
         Caption         =   "Tanca 3"
         Height          =   375
         Left            =   12510
         TabIndex        =   73
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta 3"
         Height          =   375
         Index           =   11
         Left            =   5400
         TabIndex        =   69
         Top             =   180
         Width           =   1020
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta 3 Argox C"
         Height          =   375
         Index           =   10
         Left            =   10530
         TabIndex        =   68
         Top             =   675
         Width           =   1500
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Roupa 3"
         Height          =   375
         Index           =   9
         Left            =   2025
         TabIndex        =   67
         Top             =   675
         Width           =   1500
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Peq 2"
         Height          =   375
         Index           =   8
         Left            =   12150
         TabIndex        =   66
         Top             =   675
         Width           =   1230
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta 2"
         Height          =   375
         Index           =   7
         Left            =   10155
         TabIndex        =   64
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Bijoteria 2"
         Height          =   375
         Index           =   6
         Left            =   9120
         TabIndex        =   63
         Top             =   180
         Width           =   960
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Roupa 3 Inv"
         Height          =   375
         Index           =   5
         Left            =   3600
         TabIndex        =   61
         Top             =   675
         Width           =   1725
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Roupa 3 Gde"
         Height          =   375
         Index           =   4
         Left            =   135
         TabIndex        =   59
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox TxtJuros 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         TabIndex        =   58
         Top             =   270
         Width           =   3870
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Bijoteria 1"
         Height          =   375
         Index           =   3
         Left            =   7920
         TabIndex        =   57
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Roupa (Juros)"
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   23
         Top             =   675
         Width           =   1770
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta Gde "
         Height          =   375
         Index           =   1
         Left            =   7305
         TabIndex        =   22
         Top             =   675
         Width           =   1410
      End
      Begin VB.CommandButton BtEtq 
         Caption         =   "Etiqueta 3 L-42"
         Height          =   375
         Index           =   2
         Left            =   6540
         TabIndex        =   21
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   225
      TabIndex        =   12
      Top             =   2415
      Width           =   7920
      Begin VB.CommandButton BtAdiciona 
         Caption         =   "&Adiciona"
         Height          =   315
         Left            =   6345
         TabIndex        =   4
         Top             =   225
         Width           =   1080
      End
      Begin MSMask.MaskEdBox MebQtd 
         Height          =   330
         Left            =   1110
         TabIndex        =   1
         Top             =   210
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPcl 
         Height          =   330
         Left            =   2775
         TabIndex        =   2
         Top             =   210
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebVlrPcl 
         Height          =   330
         Left            =   5070
         TabIndex        =   3
         Top             =   210
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
         Height          =   195
         Left            =   2055
         TabIndex        =   14
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Parcela"
         Height          =   195
         Left            =   3945
         TabIndex        =   13
         Top             =   225
         Width           =   945
      End
   End
   Begin VB.CheckBox Opt 
      Caption         =   "Parcelas com Juros"
      Height          =   285
      Left            =   8280
      TabIndex        =   11
      Top             =   1515
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.CommandButton BtEtqPrd 
      Height          =   690
      Left            =   12825
      Picture         =   "FormEtiquetas.frx":00BF
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Etiqueta Produto Grande (EtqPrd)"
      Top             =   1395
      Width           =   780
   End
   Begin VB.CommandButton BtBarra 
      Height          =   690
      Left            =   11745
      Picture         =   "FormEtiquetas.frx":0501
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "RelBarra"
      Top             =   1440
      Width           =   780
   End
   Begin VB.CommandButton BtImprimir 
      Height          =   690
      Left            =   12825
      Picture         =   "FormEtiquetas.frx":0943
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "RelEtq"
      Top             =   630
      Width           =   780
   End
   Begin VB.CommandButton BtLimpa 
      Height          =   690
      Left            =   11745
      Picture         =   "FormEtiquetas.frx":0D85
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   630
      Width           =   780
   End
   Begin VB.Data Datdados 
      Caption         =   "Etiquetas"
      Connect         =   "Access"
      DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormEtiquetas.frx":11C7
      Top             =   4530
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "FormEtiquetas.frx":1286
      Height          =   3510
      Left            =   240
      OleObjectBlob   =   "FormEtiquetas.frx":12A0
      TabIndex        =   9
      Top             =   3105
      Width           =   10170
   End
   Begin VB.Data Datdados 
      Caption         =   "Produtos"
      Connect         =   "Access"
      DatabaseName    =   "C:\AATESTE\Reta\Produtos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2520
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FormEtiquetas.frx":24E6
      Top             =   810
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisas"
      Height          =   1845
      Left            =   225
      TabIndex        =   10
      Top             =   555
      Width           =   7905
      Begin VB.CommandButton cmdcompl 
         Caption         =   ">>"
         Height          =   375
         Left            =   7440
         TabIndex        =   90
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtComplemento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5730
         TabIndex        =   89
         ToolTipText     =   "Coloque * antes e/ou depois da palavra para filtrar"
         Top             =   240
         Width           =   1590
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormEtiquetas.frx":259A
         Height          =   1530
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   2699
         _Version        =   393216
         Appearance      =   0
         Style           =   1
         ListField       =   "descrição_Produto"
         BoundColumn     =   "CodPrd"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MebCusto 
         Height          =   330
         Left            =   5730
         TabIndex        =   16
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebVenda 
         Height          =   330
         Left            =   5730
         TabIndex        =   17
         Top             =   1440
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebEst 
         Height          =   330
         Left            =   6795
         TabIndex        =   92
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         Height          =   195
         Left            =   6795
         TabIndex        =   93
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda"
         Height          =   195
         Left            =   5730
         TabIndex        =   19
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo"
         Height          =   195
         Left            =   5730
         TabIndex        =   18
         Top             =   660
         Width           =   645
      End
   End
   Begin MSMask.MaskEdBox MebDesc 
      Height          =   330
      Left            =   8280
      TabIndex        =   25
      Top             =   2445
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebPedido 
      Height          =   330
      Left            =   10170
      TabIndex        =   74
      Top             =   1410
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   315
      Left            =   10170
      TabIndex        =   84
      Top             =   720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   98631681
      CurrentDate     =   37501
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      Height          =   195
      Left            =   10170
      TabIndex        =   75
      Top             =   1170
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Perc. Desconto"
      Height          =   195
      Left            =   8280
      TabIndex        =   26
      Top             =   2205
      Width           =   1110
   End
End
Attribute VB_Name = "FormEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TEtiqueta
    Altura As Double
    Largura As Double
    MEsq As Double
    MSup As Double
    MInf As Double
    Colunas As Integer
End Type
'  Matriz para guardar informações das etiquetas
Private Etiqs() As TEtiqueta


Public Ind As Recordset

Private Sub Acoff_Click()
Dim Cont As Byte
Dim VarNome1(1 To 4) As String
Dim VarNome2(1 To 4) As String
Dim VarBarra(1 To 4) As String
Dim VarPreco(1 To 4) As String
Dim VarVista(1 To 4) As String
Dim VarPreco2(1 To 4) As String

Cont = 1
For Cont = 1 To 4
    VarNome1(Cont) = ""
    VarNome2(Cont) = ""
    VarBarra(Cont) = ""
    VarPreco(Cont) = ""
Next
Cont = 1
       
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = Trim(Datdados(0).Recordset("CodBar"))
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
                    
'                    Print #1, "1A3203000600005" & VarBarra(1)
'                    Print #1, "1A3203000600140" & VarBarra(2)
'                    Print #1, "1A3203000600280" & VarBarra(3)
                                        
'                   etiqueta 2
'                    Print #1, "1F1202000150015" & VarBarra(1)
'                    Print #1, "1F1202000150200" & VarBarra(2)

                    Print #1, "1F1202000150005" & VarBarra(1)
                    Print #1, "1F1202000150140" & VarBarra(2)
                    Print #1, "1F1202000150280" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
                    
                    
                    
               Wend
               Close #1
            End If
        End If

End Sub

Private Sub BtAdiciona_Click()
Dim X As Integer
If MebQtd.Text <> "" Then
    For X = 1 To CInt(MebQtd)
        Datdados(0).Recordset.AddNew
        Datdados(0).Recordset.Fields("Produto") = DbcPesquisa.BoundText
        Datdados(0).Recordset.Fields("Qtd") = MebQtd
        If ChkDesconto.value = 1 Then
            Datdados(0).Recordset.Fields("Preco") = (Datdados(1).Recordset.Fields("Venda") - (Datdados(1).Recordset.Fields("Venda") * (MebDesc / 100)))
        Else
            Datdados(0).Recordset.Fields("Preco") = Datdados(1).Recordset.Fields("Venda")
        End If
        Datdados(0).Recordset.Fields("Parcela") = MebPcl
        Datdados(0).Recordset.Fields("Valor") = MebVlrPcl
        Datdados(0).Recordset.Update
    Next
End If
'DbcPesquisa.Text = ""
MebQtd.Text = ""
MebPcl.Text = 0
MebVlrPcl.Text = 0
DbcPesquisa.SetFocus

End Sub

Private Sub BtBalanca_Click()
Dim X As Integer
'If MebPedido.Text <> "" Then
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto,  PRD.Custo, PRD.Margem, PRD.Venda, PRD.Chefe, Prd.Estoque From Prd WHERE Chefe = true ORDER BY CodPrd "
    Datdados(1).Refresh
    
    If Datdados(1).Recordset.RecordCount > 0 Then
        Datdados(1).Recordset.MoveFirst
        While Not Datdados(1).Recordset.Eof
            Datdados(0).Recordset.AddNew
            Datdados(0).Recordset.Fields("Produto") = Datdados(1).Recordset("CodPrd")
            Datdados(0).Recordset.Fields("Qtd") = 1
            Datdados(0).Recordset.Fields("Parcela") = 1
            Datdados(0).Recordset.Fields("Valor") = Datdados(1).Recordset.Fields("Venda")
            Datdados(0).Recordset.Fields("Preco") = Datdados(1).Recordset.Fields("Venda")
            Datdados(0).Recordset.Update
            Datdados(1).Recordset.MoveNext
        Wend
    
    End If
    
'End If

End Sub

Private Sub BtBarra_Click()

'MDIPrincipal.ComRel.ShowPrinter
MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelBarra.rpt")

Imprimir

End Sub

Private Sub BtEtq_Click(Index As Integer)
Dim Cont As Byte
Dim VarNome1(1 To 6) As String
Dim VarNome2(1 To 6) As String
Dim VarBarra(1 To 6) As String
Dim VarPreco(1 To 6) As String
Dim VarVista(1 To 6) As String
Dim VarPreco2(1 To 6) As String
Dim VarPrazo(1 To 6) As String

Cont = 1
For Cont = 1 To 6
    VarNome1(Cont) = ""
    VarNome2(Cont) = ""
    VarBarra(Cont) = ""
    VarPreco(Cont) = ""
Next
Cont = 1
Select Case Index
    Case 0
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    Print #1, "191200301600005" & Left(Cabecalho1, 15)
                    Print #1, "191200301600145" & Left(Cabecalho1, 15)
                    Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                VarVista(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                                VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2) & " =" & Alinhar(Datdados(0).Recordset("Valor") * Datdados(0).Recordset("Parcela"), 6, 2)
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarVista(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                    Print #1, "191200201300005" & VarNome1(1)
                    Print #1, "191200201100005" & VarNome2(1)
                    Print #1, "191200201300145" & VarNome1(2)
                    Print #1, "191200201100145" & VarNome2(2)
                    Print #1, "191200201300280" & VarNome1(3)
                    Print #1, "191200201100280" & VarNome2(3)
                    
                    Print #1, "1A3203000600005" & VarBarra(1)
                    Print #1, "1A3203000600140" & VarBarra(2)
                    Print #1, "1A3203000600280" & VarBarra(3)
                    
                    
                    Print #1, "191000100500010" & TxtJuros
                    Print #1, "191000100500145" & TxtJuros
                    Print #1, "191000100500290" & TxtJuros
                                                            
                    Print #1, "192200100300010" & "A Vista: " & VarVista(1)
                    Print #1, "192200100300145" & "A Vista: " & VarVista(2)
                    Print #1, "192200100300290" & "A Vista: " & VarVista(3)
                    
'                    Print #1, "191200300050010R$ " & Alinhar(VarPreco(1), 6, 2)
'                    Print #1, "191200300050155R$ " & Alinhar(VarPreco(2), 6, 2)
'                    Print #1, "191200300050290R$ " & Alinhar(VarPreco(3), 6, 2)
                    Print #1, "191200300050010 " & VarPreco(1)
                    Print #1, "191200300050153 " & VarPreco(2)
                    Print #1, "191200300050289 " & VarPreco(3)
                    Print #1, "Q001"
                    
                    Print #1, "E"
'                    If Not DatDados(0).Recordset.Eof Then
'                        DatDados(0).Recordset.MoveNext
'                    End If
               Wend
               Close #1
            End If
        End If
    Case 1
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    Print #1, "192200200750100" & Left(Cabecalho1, 15)
                    Print #1, "191200200430015" & Left(Datdados(0).Recordset("Descrição_Produto"), 40)
                    Print #1, "1a3202700160010" & psCompString(1, Left(Datdados(0).Recordset("Produto"), 8), 8)
                    Print #1, "192200100000010" & psCompString(1, Left(Datdados(0).Recordset("Produto"), 8), 8)
                    If ChkPreco.value = 1 Then
                        Print #1, "192300200050140R$ " & Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                    Else
                        Print #1, "192300200050140 "
                    End If
                    Print #1, "192200200350320" & Datdados(0).Recordset("Parcela") & " X"
                    Print #1, "192200200050290R$ " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                    Print #1, "Q001"
                    Print #1, "E"
                    Datdados(0).Recordset.MoveNext
               Wend
               Close #1
            End If
        End If
        
    Case 2
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If Datdados(0).Recordset("Parcela") <> 0 Then
                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                Else
'                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "191100100630000" & VarNome1(1)
'                    Print #1, "191100100630140" & VarNome1(2)
'                    Print #1, "191100100630280" & VarNome1(3)
                    
                    Print #1, "111300000440005R$ " & Alinhar(VarPreco(1), 6, 2)
                    Print #1, "111300000440145R$ " & Alinhar(VarPreco(2), 6, 2)
                    Print #1, "111300000440280R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    Print #1, "111300000440055" & VarPreco2(1)
                    Print #1, "111300000440195" & VarPreco2(2)
                    Print #1, "111300000440335" & VarPreco2(3)
                    
                    Print #1, "1A3103000030020" & VarBarra(1)
                    Print #1, "1A3103000030150" & VarBarra(2)
                    Print #1, "1A3103000030295" & VarBarra(3)
                    
'                    Print #1, "191200100000045" & VarBarra(1)
'                    Print #1, "191200100000195" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
                    For Cont = 1 To 3
                        VarNome1(Cont) = ""
                        VarNome2(Cont) = ""
                        VarBarra(Cont) = ""
                        VarPreco(Cont) = ""
                        VarPreco2(Cont) = ""
                    Next
               Wend
               Close #1
            End If
        End If
    Case 3
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"

'                    Print #1, "191200301600005" & Left(Cabecalho1, 15)
'                    Print #1, "191200301600145" & Left(Cabecalho1, 15)
'                    Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
'                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
'                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                VarPreco(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                            Else
'                                VarNome1(Cont) = ""
'                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
'                            VarNome1(Cont) = ""
'                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                    
                    Print #1, "191200200600015" & VarBarra(1)
                    Print #1, "191200200600110R$ " & Alinhar(VarPreco(1), 6, 2)
                    
                    Print #1, "191200200850190" & VarBarra(2)
                    Print #1, "191200200850290R$ " & Alinhar(VarPreco(2), 6, 2)
                    
'                    Print #1, "191200300600015" & VarBarra(3)
'                    Print #1, "191200300600110R$ " & Alinhar(VarPreco(3), 6, 2)
                    
'                    Print #1, "191200300850190" & VarBarra(4)
'                    Print #1, "191200300850290R$ " & Alinhar(VarPreco(4), 6, 2)
                    
                    
                    
'                    Print #1, "191200200040020" & VarBarra(1)
'                    Print #1, "191200200040100R$ " & Alinhar(VarPreco(1), 6, 2)
'                    Print #1, "191200200040020" & ""
'                    Print #1, "191200200040210" & VarBarra(2)
'                    Print #1, "191200200040280R$ " & Alinhar(VarPreco(2), 6, 2)


'                    If Not DatDados(0).Recordset.Eof Then
'                        DatDados(0).Recordset.MoveNext
'                    End If
                Print #1, "Q001"
                Print #1, "E"
               
               Wend
               
               Close #1
            End If
        End If
    Case 4
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
 '               Dim x As Integer
 '               x = 2
 '               If x = 1 Then
                   Datdados(0).Recordset.MoveFirst
                   Open VarImprBarra For Output As #1
                   While Not Datdados(0).Recordset.Eof
                        Print #1, Chr(2) + "O0220"
                        Print #1, Chr(2) + "M0300"
                        Print #1, Chr(2) + "c0000"
                        Print #1, Chr(2) + "f000"
                        Print #1, Chr(2) + "e"
                        Print #1, Chr(2) + "LC0000"
                        Print #1, "H09"
                        Print #1, "D11"
                        Print #1, "SC"
                        Print #1, "PC"
                        Print #1, "R0000"
                        Print #1, "z"
                        Print #1, "W"
                        Print #1, Chr(94) + "01"
                        Print #1, "191200301600005" & Left(Cabecalho1, 15)
                        Print #1, "191200301600145" & Left(Cabecalho1, 15)
                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                    VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    VarPreco2(Cont) = Datdados(0).Recordset("Preco")
                                    
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                        Next
                        Print #1, "191200201300005" & VarNome1(1)
                        Print #1, "191200201100005" & VarNome2(1)
                        Print #1, "191200201300145" & VarNome1(2)
                        Print #1, "191200201100145" & VarNome2(2)
                        Print #1, "191200201300280" & VarNome1(3)
                        Print #1, "191200201100280" & VarNome2(3)
                        
 '                       Print #1, "1A3103000300015" & VarBarra(1)
 '                       Print #1, "1A3103000300150" & VarBarra(2)
 '                       Print #1, "1A3103000300280" & VarBarra(3)
                        
                        Print #1, "1E2203000600005" & VarBarra(1)
                        Print #1, "1E2203000600145" & VarBarra(2)
                        Print #1, "1E2203000600292" & VarBarra(3)
                        
'                        Print #1, "192200100400010" & VarBarra(1)
'                        Print #1, "192200100400145" & VarBarra(2)
'                        Print #1, "192200100400290" & VarBarra(3)
                        
                        Print #1, "192200100400010R$ " & Alinhar(VarPreco2(1), 6, 2)
                        Print #1, "192200100400145R$ " & Alinhar(VarPreco2(2), 6, 2)
                        Print #1, "192200100400290R$ " & Alinhar(VarPreco2(3), 6, 2)
                        
                        
'                        Print #1, "191200201550025R$ " & Alinhar(VarPreco2(1), 6, 2)
'                        Print #1, "191200201550160R$ " & Alinhar(VarPreco2(2), 6, 2)
'                        Print #1, "191200201550295R$ " & Alinhar(VarPreco2(3), 6, 2)
                        
                        
    '                    Print #1, "191200300050010R$ " & Alinhar(VarPreco(1), 6, 2)
    '                    Print #1, "191200300050155R$ " & Alinhar(VarPreco(2), 6, 2)
    '                    Print #1, "191200300050290R$ " & Alinhar(VarPreco(3), 6, 2)
                        Print #1, "191200300050010 " & VarPreco(1)
                        Print #1, "191200300050155 " & VarPreco(2)
                        Print #1, "191200300050290 " & VarPreco(3)
                        
'                        Print #1, "191100100000010" & "Sem Juros"
'                        Print #1, "191100100000145" & "Sem Juros"
'                        Print #1, "191100100000290" & "Sem Juros"
                        
                        Print #1, "Q001"
                        
                        Print #1, "E"
    '                    If Not DatDados(0).Recordset.Eof Then
    '                        DatDados(0).Recordset.MoveNext
    '                    End If
                   Wend
                   Close #1
            End If
        End If
    Case 5
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
 '               Dim x As Integer
 '               x = 2
 '               If x = 1 Then
                   Datdados(0).Recordset.MoveFirst
                   Open VarImprBarra For Output As #1
                   While Not Datdados(0).Recordset.Eof
                        Print #1, Chr(2) + "O0220"
                        Print #1, Chr(2) + "M0300"
                        Print #1, Chr(2) + "c0000"
                        Print #1, Chr(2) + "f000"
                        Print #1, Chr(2) + "e"
                        Print #1, Chr(2) + "LC0000"
                        Print #1, "H09"
                        Print #1, "D11"
                        Print #1, "SC"
                        Print #1, "PC"
                        Print #1, "R0000"
                        Print #1, "z"
                        Print #1, "W"
                        Print #1, Chr(94) + "01"
'                        Print #1, "191200301600005" & Left(Cabecalho1, 15)
'                        Print #1, "191200301600145" & Left(Cabecalho1, 15)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                                                
                        
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 21, 20)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                    VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & "  X  " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2) ' & " s/juros"
                                    If ChkValor.value = 1 Then
                                        VarPreco2(Cont) = Datdados(0).Recordset("Preco")
                                    End If
                                    
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                    VarPreco2(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                        Next
                        
'                        Print #1, "191100201300020 " & VarPreco(1)
'                        Print #1, "191100201300155 " & VarPreco(2)
'                        Print #1, "191100201300290 " & VarPreco(3)
                        
                        Print #1, "191200201550025R$ " & Alinhar(VarPreco2(1), 6, 2)
                        Print #1, "191200201550160R$ " & Alinhar(VarPreco2(2), 6, 2)
                        Print #1, "191200201550295R$ " & Alinhar(VarPreco2(3), 6, 2)
                                                
                                                
                        Print #1, "191200100300005" & VarNome1(1)
                        Print #1, "191200100100005" & VarNome2(1)
                        Print #1, "191200100300145" & VarNome1(2)
                        Print #1, "191200100100145" & VarNome2(2)
                        Print #1, "191200100300280" & VarNome1(3)
                        Print #1, "191200100100280" & VarNome2(3)
                        
                        Print #1, "1E2203000600005" & VarBarra(1)
                        Print #1, "1E2203000600140" & VarBarra(2)
                        Print #1, "1E2203000600275" & VarBarra(3)
                        
                        
                        Print #1, "191200101050020" & Left(Cabecalho1, 15)
                        Print #1, "191200101050160" & Left(Cabecalho1, 15)
                        Print #1, "191200101050300" & Left(Cabecalho1, 15)
                        
                        
'                        Print #1, "192200100150010" & VarBarra(1)
'                        Print #1, "192200100150140" & VarBarra(2)
'                        Print #1, "192200100150285" & VarBarra(3)
                        
    '                    Print #1, "191200300050010R$ " & Alinhar(VarPreco(1), 6, 2)
    '                    Print #1, "191200300050155R$ " & Alinhar(VarPreco(2), 6, 2)
    '                    Print #1, "191200300050290R$ " & Alinhar(VarPreco(3), 6, 2)
                        Print #1, "Q001"
                        
                        Print #1, "E"
    '                    If Not DatDados(0).Recordset.Eof Then
    '                        DatDados(0).Recordset.MoveNext
    '                    End If
                   Wend
                   Close #1
            End If
        End If
    Case 6
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H12"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 35)
'                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                    
'                    Print #1, "191200200600015" & VarBarra(1)
'                    Print #1, "191200200850190" & VarBarra(2)
'                    Print #1, "191200200600110R$ " & Alinhar(VarPreco(1), 6, 2)
'                    Print #1, "191200200850290R$ " & Alinhar(VarPreco(2), 6, 2)
                    
                    
'                    Print #1, "191200301600005" & "  TT "
                    Print #1, "191100100200015" & VarNome1(1)
                    Print #1, "191100100100025" & VarBarra(1)
                    Print #1, "1a3101000010015" & VarBarra(1)
                    If ChkValor.value = 1 Then
                        Print #1, "191200100010125R$ " & VarPreco(1)
                    End If
                    
                    Print #1, "191100100200220" & VarNome1(2)
                    Print #1, "191100100100235" & VarBarra(2)
                    Print #1, "1a3101000010212" & VarBarra(2)
                    If ChkValor.value = 1 Then
                        Print #1, "191200100010315R$ " & VarPreco(2)
                    End If
                    
                    
                    Print #1, "Q001"
                    Print #1, "E"
               
               Wend
               
               Close #1
            End If
        End If
    
    
    
    Case 7
    
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
'                                VarNome1(Cont) = Trim(DatDados(0).Recordset("Produto")) & "-" & Left(DatDados(0).Recordset("Descrição_Produto"), 20)
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 25)
                                
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 26, 20)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "192200000850020" & Left(Trim(Cabecalho1), 18)
'                    Print #1, "192200000850230" & Left(Trim(Cabecalho1), 18)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    
                    
'                    Print #1, "191100100750040" & VarNome1(1)
'                    Print #1, "191100100750230" & VarNome1(2)
                    
                    Print #1, "191100200250010" & VarNome1(1)
                    Print #1, "191100200250215" & VarNome1(2)
                    
                    Print #1, "191100200100010" & VarNome2(1)
                    Print #1, "191100200100215" & VarNome2(2)
                    
                    
                    'Print #1, "191100100630285" & VarNome1(3)
                    If ChkValor.value = 1 Then
                        Print #1, "193200000800030R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "193200000800250R$ " & Alinhar(VarPreco(2), 6, 2)
                    End If
                    If ChkValor = 1 Then
'                        Print #1, "111300000550100R$ " & VarPreco2(1)
'                        Print #1, "111300000550310R$ " & VarPreco2(2)
'                        Print #1, "111300000150165" & "S/Juros"
'                        Print #1, "111300000150355" & "S/Juros"
                    End If
'                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    
                    Print #1, "1E2203000400020" & VarBarra(1)
                    Print #1, "1E2203000400250" & VarBarra(2)
'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
   Case 8
   
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 35)
'                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
'                                    VarPreco(Cont) = Alinhar(DatDados(0).Recordset("Preco"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                    
'                    Print #1, "191100100200015" & VarNome1(1)
                    Print #1, "193200100100015" & VarBarra(1)
                    Print #1, "1a3101000010015" & VarBarra(1)
                    If ChkValor.value = 1 Then
                        Print #1, "193200100250005R$ " & VarPreco(1)
                    End If
                    
'                    Print #1, "191100100200220" & VarNome1(2)
                    Print #1, "193200100100215" & VarBarra(2)
                    Print #1, "1a3101000010215" & VarBarra(2)
                    If ChkValor.value = 1 Then
                        Print #1, "193200100250195R$ " & VarPreco(2)
                    End If
                    
                    
                    Print #1, "Q001"
                    Print #1, "E"
               
               Wend
               
               Close #1
            End If
        End If
   
    Case 9
         If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
                If VarImprBarra <> "" Then
                                   
                   Datdados(0).Recordset.MoveFirst
                   Open VarImprBarra For Output As #1
                   While Not Datdados(0).Recordset.Eof
                        Print #1, Chr(2) + "O0220"
                        Print #1, Chr(2) + "M0300"
                        Print #1, Chr(2) + "c0000"
                        Print #1, Chr(2) + "f000"
                        Print #1, Chr(2) + "e"
                        Print #1, Chr(2) + "LC0000"
                        Print #1, "H09"
                        Print #1, "D11"
                        Print #1, "SC"
                        Print #1, "PC"
                        Print #1, "R0000"
                        Print #1, "z"
                        Print #1, "W"
                        Print #1, Chr(94) + "01"
'                        Print #1, "191200301600005" & Left(Cabecalho1, 15)
'                        Print #1, "191200301600145" & Left(Cabecalho1, 15)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                    If ChkPreco.value = 0 Then
                                        VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    End If
                                    VarPreco2(Cont) = Datdados(0).Recordset("Preco")
                                    
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                        Next
                        
                        Print #1, "191200201800025" & Left(Cabecalho1, 15)
                        Print #1, "191200201800160" & Left(Cabecalho1, 15)
                        Print #1, "191200201800290" & Left(Cabecalho1, 15)
                        
                        Print #1, "191200201400010" & VarNome1(1)
                        Print #1, "191200201200010" & VarNome2(1)
                        Print #1, "191200201400150" & VarNome1(2)
                        Print #1, "191200201200150" & VarNome2(2)
                        Print #1, "191200201400285" & VarNome1(3)
                        Print #1, "191200201200285" & VarNome2(3)
                        
                        
                        Print #1, "1E2203000700005" & VarBarra(1)
                        Print #1, "1E2203000700142" & VarBarra(2)
                        Print #1, "1E2203000700280" & VarBarra(3)
                        
                        
                        Print #1, "192200100600025" & VarBarra(1)
                        Print #1, "192200100600160" & VarBarra(2)
                        Print #1, "192200100600295" & VarBarra(3)
                        If chkprazo <> 0 Then
                            Print #1, "192200100200020R$" & Alinhar(VarPreco2(1), 6, 2)
                            Print #1, "192200100200155R$" & Alinhar(VarPreco2(2), 6, 2)
                            Print #1, "192200100200295R$" & Alinhar(VarPreco2(3), 6, 2)
                        Else
                            Print #1, "192200100200025R$ " & Alinhar(VarPreco2(1), 6, 2)
                            Print #1, "192200100200155R$ " & Alinhar(VarPreco2(2), 6, 2)
                            Print #1, "192200100200295R$ " & Alinhar(VarPreco2(3), 6, 2)
                            
'                            Print #1, "191100100000010" & "Sem Juros"
'                            Print #1, "191100100000145" & "Sem Juros"
'                            Print #1, "191100100000290" & "Sem Juros"
                        
                        End If
                        
'                        Print #1, "191200201550025R$ " & Alinhar(VarPreco2(1), 6, 2)
'                        Print #1, "191200201550160R$ " & Alinhar(VarPreco2(2), 6, 2)
'                        Print #1, "191200201550295R$ " & Alinhar(VarPreco2(3), 6, 2)
                        
                        
    '                    Print #1, "191200300050010R$ " & Alinhar(VarPreco(1), 6, 2)
    '                    Print #1, "191200300050155R$ " & Alinhar(VarPreco(2), 6, 2)
    '                    Print #1, "191200300050290R$ " & Alinhar(VarPreco(3), 6, 2)
                        If chkprazo <> 0 Then
                            Print #1, "191200100000015 " & VarPreco(1)
                            Print #1, "191200100000150 " & VarPreco(2)
                            Print #1, "191200100000295 " & VarPreco(3)
                        End If
                        Print #1, "Q001"
                        
                        Print #1, "E"
    '                    If Not DatDados(0).Recordset.Eof Then
    '                        DatDados(0).Recordset.MoveNext
    '                    End If
                   Wend
                   Close #1
               Else
                                   
                   Datdados(0).Recordset.MoveFirst
'                   Open VarImprBarra For Output As #1
                   While Not Datdados(0).Recordset.Eof
'                        Print #1, Chr(2) + "O0220"
'                        Print #1, Chr(2) + "M0300"
'                        Print #1, Chr(2) + "c0000"
'                        Print #1, Chr(2) + "f000"
'                        Print #1, Chr(2) + "e"
'                        Print #1, Chr(2) + "LC0000"
'                        Print #1, "H09"
'                        Print #1, "D11"
'                        Print #1, "SC"
'                        Print #1, "PC"
'                        Print #1, "R0000"
'                        Print #1, "z"
'                        Print #1, "W"
'                        Print #1, Chr(94) + "01"
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 17)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 18, 15)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 7)
                                    If ChkPreco.value = 0 Then
                                        VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    End If
                                    VarPreco2(Cont) = Datdados(0).Recordset("Preco")
                                    
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                        Next
                    AbreAcbr "ETQ.Ativar()"
                    TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.SetTemperatura(11)"
                    TrataAcbr "Etiqueta", 0
                        
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,3,43,4," & """" & Left(Cabecalho1, 15) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,3,43,39," & """" & Left(Cabecalho1, 15) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,3,43,73," & """" & Left(Cabecalho1, 15) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                                                            
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,34,2," & """" & VarNome1(1) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,29,2," & """" & VarNome2(1) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,34,37," & """" & VarNome1(2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,29,37," & """" & VarNome2(2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,34,72," & """" & VarNome1(3) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,29,72," & """" & VarNome2(3) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                                        
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,19,2," & VarBarra(1) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,19,37," & VarBarra(2) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,19,72," & VarBarra(3) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,10," & """" & VarBarra(1) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,45," & """" & VarBarra(2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,80," & """" & VarBarra(3) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    
                        
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,3,6," & """" & "R$ " & Alinhar(VarPreco2(1), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,3,42," & """" & "R$ " & Alinhar(VarPreco2(2), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,3,75," & """" & "R$ " & Alinhar(VarPreco2(3), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                                               
                        
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                        
                    
                   AbreAcbr "ETQ.Imprimir(1,0)"
                   TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.Desativar()"
                    TrataAcbr "Etiqueta", 0
                        
                        
'                        Print #1, "191200201800025" & Left(Cabecalho1, 15)
'                        Print #1, "191200201800160" & Left(Cabecalho1, 15)
'                        Print #1, "191200201800290" & Left(Cabecalho1, 15)
                        
'                        Print #1, "191200201400010" & VarNome1(1)
'                        Print #1, "191200201200010" & VarNome2(1)
'                        Print #1, "191200201400150" & VarNome1(2)
'                        Print #1, "191200201200150" & VarNome2(2)
'                        Print #1, "191200201400285" & VarNome1(3)
'                        Print #1, "191200201200285" & VarNome2(3)
                        
'                        Print #1, "1E2203000700005" & VarBarra(1)
'                        Print #1, "1E2203000700142" & VarBarra(2)
'                        Print #1, "1E2203000700280" & VarBarra(3)
                                                
'                        Print #1, "192200100600025" & VarBarra(1)
'                        Print #1, "192200100600160" & VarBarra(2)
'                        Print #1, "192200100600295" & VarBarra(3)
'                        If chkprazo <> 0 Then
'                            Print #1, "192200100200020R$" & Alinhar(VarPreco2(1), 6, 2)
'                            Print #1, "192200100200155R$" & Alinhar(VarPreco2(2), 6, 2)
'                            Print #1, "192200100200295R$" & Alinhar(VarPreco2(3), 6, 2)
'                        Else
'                            Print #1, "192200100200025R$ " & Alinhar(VarPreco2(1), 6, 2)
'                            Print #1, "192200100200155R$ " & Alinhar(VarPreco2(2), 6, 2)
'                            Print #1, "192200100200295R$ " & Alinhar(VarPreco2(3), 6, 2)
'
 '                       End If
                        
'                        If chkprazo <> 0 Then
'                            Print #1, "191200100000015 " & VarPreco(1)
'                            Print #1, "191200100000150 " & VarPreco(2)
'                            Print #1, "191200100000295 " & VarPreco(3)
'                        End If
                   
                   Wend
                   Close #1
               
               
               End If
            End If
        End If

    Case 10
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
                If VarImprBarra <> "" Then
                   Datdados(0).Recordset.MoveFirst
                   Open VarImprBarra For Output As #1
                   While Not Datdados(0).Recordset.Eof
                        Print #1, Chr(2) + "O0220"
                        Print #1, Chr(2) + "M0300"
                        Print #1, Chr(2) + "c0000"
                        Print #1, Chr(2) + "f000"
                        Print #1, Chr(2) + "e"
                        Print #1, Chr(2) + "LC0000"
                        Print #1, "H09"
                        Print #1, "D11"
                        Print #1, "SC"
                        Print #1, "PC"
                        Print #1, "R0000"
                        Print #1, "z"
                        Print #1, "W"
                        Print #1, Chr(94) + "01"
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 6)
                                    If ChkValor.value = 1 Then
                                        VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                    End If
                                    If Datdados(0).Recordset("Parcela") <> 0 Then
                                        VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    Else
    '                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    End If
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                    VarPreco2(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                        Next
                                                                                        
                        Print #1, "191100100850016" & VarNome1(1)
                        Print #1, "191100100850155" & VarNome1(2)
                        Print #1, "191100100850298" & VarNome1(3)
                        
                        Print #1, "112400000600035R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "112400000600172R$ " & Alinhar(VarPreco(2), 6, 2)
                        Print #1, "112400000600312R$ " & Alinhar(VarPreco(3), 6, 2)
                            
                        Print #1, "1E2202500250018" & VarBarra(1)
                        Print #1, "1E2202500250158" & VarBarra(2)
                        Print #1, "1E2202500250300" & VarBarra(3)
                        
                        Print #1, "191200100080045" & VarBarra(1)
                        Print #1, "191200100080185" & VarBarra(2)
                        Print #1, "191200100080325" & VarBarra(3)
                        Print #1, "Q001"
                        Print #1, "E"
                                                                                                
                        For Cont = 1 To 3
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        Next
                   Wend
                   Close #1
              Else
                   Datdados(0).Recordset.MoveFirst
     
                   While Not Datdados(0).Recordset.Eof
                        For Cont = 1 To 3
                            If Not Datdados(0).Recordset.Eof Then
                                If Not Datdados(0).Recordset.NoMatch Then
                                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                    VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 7)
                                    If ChkValor.value = 1 Then
                                        VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                    End If
                                    If Datdados(0).Recordset("Parcela") <> 0 Then
                                        VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    Else
    '                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    End If
                                Else
                                    VarNome1(Cont) = ""
                                    VarNome2(Cont) = ""
                                    VarBarra(Cont) = ""
                                    VarPreco(Cont) = ""
                                    VarPreco2(Cont) = ""
                                End If
                                Datdados(0).Recordset.MoveNext
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                        Next
                        
                    AbreAcbr "ETQ.Ativar()"
                    TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.SetTemperatura(11)"
                    TrataAcbr "Etiqueta", 0
    
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,1,1,3,17,2," & """" & VarNome1(1) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,1,1,3,17,37," & """" & VarNome1(2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,1,1,3,17,72," & """" & VarNome1(3) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,09,6," & """" & "R$ " & Alinhar(VarPreco(1), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,09,42," & """" & "R$ " & Alinhar(VarPreco(2), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,09,75," & """" & "R$ " & Alinhar(VarPreco(3), 6, 2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,1,1," & VarBarra(1) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,1,36," & VarBarra(2) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,1,2,1,71," & VarBarra(3) & ",6,1)"
                    TrataAcbr "Etiqueta", 0
                        
                    AbreAcbr "ETQ.ImprimirTexto(3,2,1,1,9,31," & """" & VarBarra(1) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(3,2,1,1,9,66," & """" & VarBarra(2) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(3,2,1,1,9,101," & """" & VarBarra(3) & """" & ",0,0)"
                    TrataAcbr "Etiqueta", 0
                        
                        
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                        
                    
                   AbreAcbr "ETQ.Imprimir(1,0)"
                   TrataAcbr "Etiqueta", 0
    
                    AbreAcbr "ETQ.Desativar()"
                    TrataAcbr "Etiqueta", 0
                                                                                        
                    For Cont = 1 To 3
                        VarNome1(Cont) = ""
                        VarNome2(Cont) = ""
                        VarBarra(Cont) = ""
                        VarPreco(Cont) = ""
                        VarPreco2(Cont) = ""
                    Next
                   Wend
                   Close #1
             End If
            End If
        End If
    Case 11
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If Datdados(0).Recordset("Parcela") <> 0 Then
                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                Else
'                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
                    Print #1, "191100100400020" & VarNome1(1)
                    Print #1, "191100100400160" & VarNome1(2)
                    Print #1, "191100100400310" & VarNome1(3)
                    
                                                                                                               
                    Print #1, "1E2201500170018" & VarBarra(1)
                    Print #1, "1E2201500170157" & VarBarra(2)
                    Print #1, "1E2201500170300" & VarBarra(3)
                    
                                        
                    Print #1, "191100100030020" & VarBarra(1)
                    Print #1, "191100100030160" & VarBarra(2)
                    Print #1, "191100100030310" & VarBarra(3)
                    
                    Print #1, "191200000000070R$ " & Alinhar(VarPreco(1), 6, 2)
                    Print #1, "191200000000210R$ " & Alinhar(VarPreco(2), 6, 2)
                    Print #1, "191200000000335R$ " & Alinhar(VarPreco(3), 6, 2)
                    
                    
                    Print #1, "Q001"
                    Print #1, "E"
                    For Cont = 1 To 3
                        VarNome1(Cont) = ""
                        VarNome2(Cont) = ""
                        VarBarra(Cont) = ""
                        VarPreco(Cont) = ""
                        VarPreco2(Cont) = ""
                    Next
               Wend
               Close #1
            
'                    Print #1, "191100100650010" & VarNome1(1)
'                    Print #1, "191100100650150" & VarNome1(2)
'                    Print #1, "191100100650290" & VarNome1(3)
'
'                    Print #1, "191200000470020R$ " & Alinhar(VarPreco(1), 6, 2)
'                    Print #1, "191200000470160R$ " & Alinhar(VarPreco(2), 6, 2)
'                    Print #1, "191200000470295R$ " & Alinhar(VarPreco(3), 6, 2)
'
'                    Print #1, "1E2201500170005" & VarBarra(1)
'                    Print #1, "1E2201500170143" & VarBarra(2)
'                    Print #1, "1E2201500170275" & VarBarra(3)
'
'                    Print #1, "191200100000050" & VarBarra(1)
'                    Print #1, "191200100000190" & VarBarra(2)
'                    Print #1, "191200100000315" & VarBarra(3)
            
            
            
            End If
        End If

    Case 12
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If Datdados(0).Recordset("Parcela") <> 0 Then
                                    chkprazo.value = 1
                                    VarPreco2(Cont) = "| " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                Else
'                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
                    Print #1, "191100100630005" & VarNome1(1)
                    Print #1, "191100100630140" & VarNome1(2)
                    Print #1, "191100100630280" & VarNome1(3)
                    
                    Print #1, "111300000440015R$ " & Alinhar(VarPreco(1), 6, 2)
                    Print #1, "111300000440145R$ " & Alinhar(VarPreco(2), 6, 2)
                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                                    
                    If chkprazo = 1 Then
                        Print #1, "191200000440075" & VarPreco2(1)
                        Print #1, "191200000440215" & VarPreco2(2)
                        Print #1, "191200000440350" & VarPreco2(3)
                    End If
                                                    
                                        
'                    Print #1, "111300000440055" & VarPreco2(1)
'                    Print #1, "111300000440195" & VarPreco2(2)
'                    Print #1, "111300000440335" & VarPreco2(3)
                    
'                    Print #1, "111300000440065" & "s/juros"
'                    Print #1, "111300000440205" & "s/juros"'
'                    Print #1, "111300000440345" & "s/juros"
                    
                    Print #1, "1E2201500170005" & VarBarra(1)
                    Print #1, "1E2201500170143" & VarBarra(2)
                    Print #1, "1E2201500170278" & VarBarra(3)
                    
                    Print #1, "191200100000045" & VarBarra(1)
                    Print #1, "191200100000185" & VarBarra(2)
                    Print #1, "191200100000310" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
                    For Cont = 1 To 3
                        VarNome1(Cont) = ""
                        VarNome2(Cont) = ""
                        VarBarra(Cont) = ""
                        VarPreco(Cont) = ""
                        VarPreco2(Cont) = ""
                    Next
               Wend
               Close #1
            End If
        End If
    Case 13
    
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
'                                VarNome1(Cont) = Trim(DatDados(0).Recordset("Produto")) & "-" & Left(DatDados(0).Recordset("Descrição_Produto"), 20)
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 22)
                                
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 23, 20)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "192200000850020" & Left(Trim(Cabecalho1), 18)
'                    Print #1, "192200000850230" & Left(Trim(Cabecalho1), 18)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    
                    
'                    Print #1, "191100100750040" & VarNome1(1)
'                    Print #1, "191100100750230" & VarNome1(2)
                    
                    Print #1, "191100200250042" & VarNome1(1)
                    Print #1, "191100200250215" & VarNome1(2)
                    
                    Print #1, "191100200100042" & VarNome2(1)
                    Print #1, "191100200100215" & VarNome2(2)
                    
                    
                    'Print #1, "191100100630285" & VarNome1(3)
                    If ChkValor.value = 1 Then
                        Print #1, "193200000800055R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "193200000800250R$ " & Alinhar(VarPreco(2), 6, 2)
                    End If
                    If ChkValor = 1 Then
'                        Print #1, "111300000550100R$ " & VarPreco2(1)
'                        Print #1, "111300000550310R$ " & VarPreco2(2)
'                        Print #1, "111300000150165" & "S/Juros"
'                        Print #1, "111300000150355" & "S/Juros"
                    End If
'                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    
                    Print #1, "1E2203000400050" & VarBarra(1)
                    Print #1, "1E2203000400235" & VarBarra(2)
'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
    Case 14
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
'               Open VarImprBarra For Output As #1
               
                AbreAcbr "ETQ.Ativar()"
                TrataAcbr "Etiqueta", 0


               While Not Datdados(0).Recordset.Eof
                    
                    VarBarra(1) = Datdados(0).Recordset("Produto")
                    VarPreco(1) = Int(Datdados(0).Recordset("Preco")) & " " & Format(Datdados(0).Recordset("Parcela"), "##") & "." & Int(Datdados(0).Recordset("Valor"))
                    VarNome1(1) = Left(Datdados(0).Recordset("Descrição_Produto"), 30)
'VarBarra(1) = "10005"
                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,2,8," & VarBarra(1) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,1,1,2,2,25," & VarNome1(1)
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,5,45," & """" & "00" & VarPreco(1) & """" & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                    
                    
                     Datdados(0).Recordset.MoveNext
               
               Wend
               
               AbreAcbr "ETQ.Imprimir(1,0)"
               TrataAcbr "Etiqueta", 0

                AbreAcbr "ETQ.Desativar()"
                TrataAcbr "Etiqueta", 0
               
'               Close #1
            End If
        End If
    Case 15
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H12"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "192200000850020" & Left(Trim(Cabecalho1), 18)
'                    Print #1, "192200000850230" & Left(Trim(Cabecalho1), 18)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    
                    
'                    Print #1, "191100100750040" & VarNome1(1)
'                    Print #1, "191100100750230" & VarNome1(2)
                    
'                    Print #1, "193200002000010" & "DUBAI STORE"
'                    Print #1, "193200002000210" & "DUBAI STORE"

'                    Print #1, "191100200800020" & "Prazo maximo p/ troca"
'                    Print #1, "191100200800220" & "Prazo maximo p/ troca"
                    
'                    Print #1, "191100200700020" & "15 dias apos a compra"
'                    Print #1, "191100200700220" & "15 dias apos a compra"
                    
                    
                    Print #1, "191100200460002" & VarNome1(1)
                    Print #1, "191100200460120" & VarNome1(2)
                    Print #1, "191100200460235" & VarNome1(3)
                    
                    
                    
                    'Print #1, "191100100630285" & VarNome1(3)
                    If ChkValor.value = 1 Then
                        Print #1, "191200000050020R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "191200000050140R$ " & Alinhar(VarPreco(2), 6, 2)
                        Print #1, "191200000050255R$ " & Alinhar(VarPreco(3), 6, 2)
                        
                    End If

''                     Print #1, "1E1203000150005" & VarBarra(1)
'                    Print #1, "1E1203000150090" & VarBarra(2)
                    Print #1, "191100100370035" & VarBarra(1)
                    Print #1, "191100100370155" & VarBarra(2)
                    Print #1, "191100100370255" & VarBarra(3)
                    
                    
                    Print #1, "1a3101000290005" & VarBarra(1)
                    Print #1, "1a3101000290135" & VarBarra(2)
                    Print #1, "1a3101000290242" & VarBarra(2)
                    

'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
    
    Case 16
    
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 24)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                
                                
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "192200000850020" & Left(Trim(Cabecalho1), 18)
'                    Print #1, "192200000850230" & Left(Trim(Cabecalho1), 18)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    
                    
'                    Print #1, "191100100750040" & VarNome1(1)
'                    Print #1, "191100100750230" & VarNome1(2)
                    
                    Print #1, "191100200200010" & VarNome1(1)
                    Print #1, "191100200200220" & VarNome1(2)
                    
                    'Print #1, "191100100630285" & VarNome1(3)
                    If ChkValor.value = 1 Then
                        Print #1, "193200000800030R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "193200000800250R$ " & Alinhar(VarPreco(2), 6, 2)
                    End If
                    If ChkValor = 1 Then
'                        Print #1, "111300000550100R$ " & VarPreco2(1)
'                        Print #1, "111300000550310R$ " & VarPreco2(2)
'                        Print #1, "111300000150165" & "S/Juros"
'                        Print #1, "111300000150355" & "S/Juros"
                    End If
'                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    
                    Print #1, "193200000400020" & VarBarra(1)
                    Print #1, "193200000400250" & VarBarra(2)
'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
    
    
   Case 17
   
           If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 2
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 25)
                                VarNome2(Cont) = Left(Datdados(0).Recordset("Complemento"), 20)
                                VarBarra(Cont) = Datdados(0).Recordset("Produto")
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
'                    Print #1, "192200000850020" & Left(Trim(Cabecalho1), 18)
'                    Print #1, "192200000850230" & Left(Trim(Cabecalho1), 18)
'                        Print #1, "191200301600280" & Left(Cabecalho1, 15)
                    
                    
'                    Print #1, "191100100750040" & VarNome1(1)
'                    Print #1, "191100100750230" & VarNome1(2)
                    
                    Print #1, "191100202000020" & VarNome2(1) & " ___/___/______"
                    Print #1, "191100202000210" & VarNome2(2) & " ___/___/______"


                    Print #1, "1E2203001500020" & VarBarra(1)
                    Print #1, "1E2203001500225" & VarBarra(2)

                    Print #1, "191100201350020" & VarNome1(1)
                    Print #1, "191100201350210" & VarNome1(2)


                    Print #1, "191100200900020" & "T-" & VarNome2(1)
                    Print #1, "191100200900210" & "T-" & VarNome2(2)


                    Print #1, "191100200700030" & "_______________________"
                    Print #1, "191100200700220" & "_______________________"
                    
                    Print #1, "191100200450020" & "V-" & VarNome2(1)
                    Print #1, "191100200450210" & "V-" & VarNome2(2)
                    
                    Print #1, "191100201100020" & "_______________________"
                    Print #1, "191100201100220" & "_______________________"
                    
                                       
  '                  Print #1, "191100201700010" & VarNome1(1)
  '                  Print #1, "191100201700210" & VarNome1(2)
                    
                    
                    'Print #1, "191100100630285" & VarNome1(3)
                    If ChkValor.value = 1 Then
                        Print #1, "193200000150030R$ " & Alinhar(VarPreco(1), 6, 2)
                        Print #1, "193200000150210R$ " & Alinhar(VarPreco(2), 6, 2)
                    End If
                    If ChkValor = 1 Then
'                        Print #1, "111300000550100R$ " & VarPreco2(1)
'                        Print #1, "111300000550310R$ " & VarPreco2(2)
'                        Print #1, "111300000150165" & "S/Juros"
'                        Print #1, "111300000150355" & "S/Juros"
                    End If
'                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    
 '                   Print #1, "1E1203001100020" & VarBarra(1)
 ''                   Print #1, "1E1203001100225" & VarBarra(2)
'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
    Case 18
    
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
'               Open VarImprBarra For Output As #1
               
                AbreAcbr "ETQ.Ativar()"
                TrataAcbr "Etiqueta", 0
                
               While Not Datdados(0).Recordset.Eof
                    
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
'                                    VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    VarPreco(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                


                    
'                    VarBarra(1) = Datdados(0).Recordset("Produto")
'                    VarPreco(1) = Int(Datdados(0).Recordset("Preco")) & " " & Format(Datdados(0).Recordset("Parcela"), "##") & "." & Int(Datdados(0).Recordset("Valor"))

                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,2," & """" & "LOJA RIO KIDS" & """" & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,33," & """" & "LOJA RIO KIDS" & """" & ")"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,62," & """" & "LOJA RIO KIDS" & """" & ")"
                    TrataAcbr "Etiqueta", 0


                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,0," & VarNome1(1) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,0," & VarNome2(1) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,28," & VarNome1(2) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,29," & VarNome2(2) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,60," & VarNome1(3) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,60," & VarNome2(3) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,1," & VarBarra(1) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,30," & VarBarra(2) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,60," & VarBarra(3) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    
                    
                    
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                    
                    

               
               Wend
               


               AbreAcbr "ETQ.Imprimir(1,0)"
               TrataAcbr "Etiqueta", 0

                AbreAcbr "ETQ.Desativar()"
                TrataAcbr "Etiqueta", 0
               
            End If
    
    Case 19
    
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
'               Open VarImprBarra For Output As #1
               
                AbreAcbr "ETQ.Ativar()"
                TrataAcbr "Etiqueta", 0
                
               While Not Datdados(0).Recordset.Eof
                    
                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 15)
                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                If ChkValor.value = 1 Then
'                                    VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                    VarPreco(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
                    Next
                
                    
'                    VarBarra(1) = Datdados(0).Recordset("Produto")
'                    VarPreco(1) = Int(Datdados(0).Recordset("Preco")) & " " & Format(Datdados(0).Recordset("Parcela"), "##") & "." & Int(Datdados(0).Recordset("Valor"))

                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,5," & """" & "LOJA RIO" & """" & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,36," & """" & "LOJA RIO" & """" & ")"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,32,65," & """" & "LOJA RIO" & """" & ")"
                    TrataAcbr "Etiqueta", 0


                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,0," & VarNome1(1) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,0," & VarNome2(1) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,28," & VarNome1(2) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,29," & VarNome2(2) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,20,60," & VarNome1(3) & ")"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,17,60," & VarNome2(3) & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,1," & VarBarra(1) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,30," & VarBarra(2) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirBarras(0,4,2,2,23,60," & VarBarra(3) & ",5,1)"
                    TrataAcbr "Etiqueta", 0
                    
                    
'                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,45,12," & VarBarra(1) & ")"
'                    TrataAcbr "Etiqueta", 0
'                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,45,35," & VarBarra(2) & ")"
'                    TrataAcbr "Etiqueta", 0
                    
'                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,1,45,75," & VarBarra(3) & ")"
'                    TrataAcbr "Etiqueta", 0
                    
                    
'                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,49,12," & """" & "R$" & VarPreco(1) & """" & ")"
'                    TrataAcbr "Etiqueta", 0
'                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,49,55," & """" & "R$" & VarPreco(2) & """" & ")"
'                    TrataAcbr "Etiqueta", 0
'
                    
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                    
                    

               
               Wend
               


               AbreAcbr "ETQ.Imprimir(1,0)"
               TrataAcbr "Etiqueta", 0

                AbreAcbr "ETQ.Desativar()"
                TrataAcbr "Etiqueta", 0
               
            End If
    
    Case 20
    
        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
'               Open VarImprBarra For Output As #1
               
                AbreAcbr "ETQ.Ativar()"
                TrataAcbr "Etiqueta", 0


               While Not Datdados(0).Recordset.Eof
                    
                    VarBarra(1) = Datdados(0).Recordset("Produto")
                    VarPreco(1) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                    VarNome1(1) = Left(Datdados(0).Recordset("Descrição_Produto"), 30)
'VarBarra(1) = "10005"
                    AbreAcbr "ETQ.IniciarEtiqueta()"
                    TrataAcbr "Etiqueta", 0

                    AbreAcbr "ETQ.ImprimirBarras(0,6,2,2,15,30," & VarBarra(1) & ",8,1)"
                    TrataAcbr "Etiqueta", 0
                    AbreAcbr "ETQ.ImprimirTexto(0,2,1,2,27,35," & VarBarra(1)
                    TrataAcbr "Etiqueta", 0
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,1,1,3,35,25," & VarNome1(1)
                    TrataAcbr "Etiqueta", 0
                    
                    
                    AbreAcbr "ETQ.ImprimirTexto(0,3,1,2,5,30," & """" & "R$ " & VarPreco(1) & """" & ")"
                    TrataAcbr "Etiqueta", 0
                    
                    
                    AbreAcbr "ETQ.FinalizarEtiqueta()"
                    TrataAcbr "Etiqueta", 0
                    
                    
                     Datdados(0).Recordset.MoveNext
               
               Wend
               
               AbreAcbr "ETQ.Imprimir(1,0)"
               TrataAcbr "Etiqueta", 0

                AbreAcbr "ETQ.Desativar()"
                TrataAcbr "Etiqueta", 0
               
'               Close #1
            End If
        End If
    
    
    Case 21
    
           If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0220"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H10"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    For Cont = 1 To 6
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 25)
                                VarNome2(Cont) = Left(Datdados(0).Recordset("Complemento"), 20)
                                VarBarra(Cont) = Datdados(0).Recordset("Produto")
                                If ChkValor.value = 1 Then
                                    VarPreco(Cont) = Datdados(0).Recordset("Preco")
                                End If
                                If chkprazo.value = 1 Then
                                    VarPreco2(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                                End If
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarPreco(Cont) = ""
                                VarPreco2(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                            VarPreco2(Cont) = ""
                        End If
                    Next
                    
'                    Print #1, "191100202000020" & VarNome2(1) & " ___/___/______"
'                    Print #1, "191100202000210" & VarNome2(2) & " ___/___/______"


                    Print #1, "1E2203001900040" & VarBarra(1)
                    Print #1, "1E2203001900225" & VarBarra(2)

'                    Print #1, "191100201700020" & VarNome1(1)
'                    Print #1, "191100201700210" & VarNome1(2)
                    
                    Print #1, "193200001600030R$ " & Alinhar(VarPreco(1), 6, 2)
                    Print #1, "193200001600210R$ " & Alinhar(VarPreco(2), 6, 2)

'                    Print #1, "191100200900030" & "_______________________"
'                    Print #1, "191100200900220" & "_______________________"

                    Print #1, "1E2203001100040" & VarBarra(3)
                    Print #1, "1E2203001100225" & VarBarra(4)

'                    Print #1, "191100201000020" & VarNome1(1)
'                    Print #1, "191100201000210" & VarNome1(2)
                    
                    Print #1, "193200000800030R$ " & Alinhar(VarPreco(3), 6, 2)
                    Print #1, "193200000800210R$ " & Alinhar(VarPreco(4), 6, 2)

'                    Print #1, "191100200900030" & "_______________________"
'                    Print #1, "191100200900220" & "_______________________"
                    
                    Print #1, "1E2203000400040" & VarBarra(5)
                    Print #1, "1E2203000400225" & VarBarra(6)

'                    Print #1, "191100200300020" & VarNome1(1)
'                    Print #1, "191100200300210" & VarNome1(2)
                    
                    Print #1, "193200000100030R$ " & Alinhar(VarPreco(5), 6, 2)
                    Print #1, "193200000100210R$ " & Alinhar(VarPreco(6), 6, 2)
                    


'                    Print #1, "191100200900020" & "T-" & VarNome2(1)
'                    Print #1, "191100200900210" & "T-" & VarNome2(2)


'                    Print #1, "191100200700030" & "_______________________"
'                    Print #1, "191100200700220" & "_______________________"
                    
'                    Print #1, "191100200450020" & "V-" & VarNome2(1)
'                    Print #1, "191100200450210" & "V-" & VarNome2(2)
'
'                    Print #1, "191100201100020" & "_______________________"
'                    Print #1, "191100201100220" & "_______________________"
                    
                                       
'                    If ChkValor.value = 1 Then
'                        Print #1, "193200000150030R$ " & Alinhar(VarPreco(1), 6, 2)
'                        Print #1, "193200000150210R$ " & Alinhar(VarPreco(2), 6, 2)
'                    End If
                    If ChkValor = 1 Then
'                        Print #1, "111300000550100R$ " & VarPreco2(1)
'                        Print #1, "111300000550310R$ " & VarPreco2(2)
'                        Print #1, "111300000150165" & "S/Juros"
'                        Print #1, "111300000150355" & "S/Juros"
                    End If
'                    Print #1, "111300000440285R$ " & Alinhar(VarPreco(3), 6, 2)
                                        
                    
 '                   Print #1, "1E1203001100020" & VarBarra(1)
 ''                   Print #1, "1E1203001100225" & VarBarra(2)
'                    Print #1, "1A3103000030280" & VarBarra(3)
                    
'                    Print #1, "191200100100075" & VarBarra(1)
'                    Print #1, "191200100100245" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
                    Print #1, "Q001"
                    Print #1, "E"
               Wend
               Close #1
            End If
        End If
    

    
End Select

End Sub

Private Sub BtEtqPrd_Click()

'MDIPrincipal.ComRel.ShowPrinter
MDIPrincipal.RptRel.ReportFileName = (Caminho & "\EtqPrd.rpt")
' MDIPrincipal.RptRel.ReportFileName = (Caminho & "\relzebra.rpt")
Imprimir

End Sub

Private Sub BtEtqTanca_Click()
Dim PrintLabels As String
Dim sImagem As String
                
Dim Cont As Byte
Dim VarNome1(1 To 4) As String
Dim VarNome2(1 To 4) As String
Dim VarBarra(1 To 4) As String
Dim VarPreco(1 To 4) As String
Dim VarVista(1 To 4) As String
Dim VarPreco2(1 To 4) As String

Cont = 1
For Cont = 1 To 4
    VarNome1(Cont) = ""
    VarNome2(Cont) = ""
    VarBarra(Cont) = ""
    VarPreco(Cont) = ""
Next
Cont = 1
                
        
If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
   Datdados(0).Recordset.MoveFirst
   
   While Not Datdados(0).Recordset.Eof
        Call openport(VarImprBarra)
        Call setup("96", "64", "3.0", "12", "0", "0", "0")

        Call clearbuffer
        
        For Cont = 1 To 3
            If Not Datdados(0).Recordset.Eof Then
                If Not Datdados(0).Recordset.NoMatch Then
                    VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 20)
                    VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 13, 12)
                    VarBarra(Cont) = "9" & psCompString(1, Datdados(0).Recordset("Produto"), 9)
                    If ChkValor.value = 1 Then
                        VarPreco(Cont) = Datdados(0).Recordset("Preco")
                    End If
                    If Datdados(0).Recordset("Parcela") <> 0 Then
                        VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                    Else
'                                    VarPreco2(Cont) = "- " & Datdados(0).Recordset("Parcela") & " X " & Alinhar(Datdados(0).Recordset("Valor"), 6, 2)
                    End If
                Else
                    VarNome1(Cont) = ""
                    VarNome2(Cont) = ""
                    VarBarra(Cont) = ""
                    VarPreco(Cont) = ""
                    VarPreco2(Cont) = ""
                End If
                Datdados(0).Recordset.MoveNext
            Else
                VarNome1(Cont) = ""
                VarNome2(Cont) = ""
                VarBarra(Cont) = ""
                VarPreco(Cont) = ""
                VarPreco2(Cont) = ""
            End If
        Next
'                    Print #1, "191100100630000" & VarNome1(1)
'                    Print #1, "191100100630140" & VarNome1(2)
'                    Print #1, "191100100630280" & VarNome1(3)
        
        


'        Call windowsfont(10, 10, 30, 90, 2, 0, "arial", VarNome1(1))
'        Call windowsfont(300, 10, 30, 90, 2, 0, "arial", VarNome1(2))
'        Call windowsfont(550, 10, 30, 90, 2, 0, "arial", VarNome1(3))
        
'        Call barcode("10", "100", "128", "100", "1", "0", "2", "2", "TESTE")
 '       Call downloadpcx("Logo.pcx", "Logo.pcx")
         Call barcode("40", "410", "128", "100", "1", "270", "2", "2", VarBarra(1))
         Call barcode("310", "410", "128", "100", "1", "270", "2", "2", VarBarra(2))
         Call barcode("570", "410", "128", "100", "1", "270", "2", "2", VarBarra(3))

        Call windowsfont(170, 400, 50, 90, 2, 0, "arial", "R$ " & Alinhar(VarPreco(1), 6, 2))
        Call windowsfont(450, 400, 50, 90, 2, 0, "arial", "R$ " & Alinhar(VarPreco(2), 6, 2))
        Call windowsfont(700, 400, 50, 90, 2, 0, "arial", "R$ " & Alinhar(VarPreco(3), 6, 2))

        Call windowsfont(40, 20, 50, 90, 2, 0, "arial", "POLLO")
        Call windowsfont(280, 20, 50, 90, 2, 0, "arial", "POLLO")
        Call windowsfont(530, 20, 50, 90, 2, 0, "arial", "POLLO")

        Call windowsfont(90, 5, 50, 90, 2, 0, "arial", "DO   ")
        Call windowsfont(330, 5, 50, 90, 2, 0, "arial", "DO   ")
        Call windowsfont(580, 5, 50, 90, 2, 0, "arial", "DO   ")

        Call windowsfont(130, 20, 50, 90, 2, 0, "arial", "JEANS")
        Call windowsfont(380, 20, 50, 90, 2, 0, "arial", "JEANS")
        Call windowsfont(630, 20, 50, 90, 2, 0, "arial", "JEANS")


'        Call sendcommand("PUTPCX 200, 80, ""Logo.pcx""")
        Call printlabel("1", "1")
'        Call formfeed
        
        
        
        
'        Print #1, "111300000440005R$ " & Alinhar(VarPreco(1), 6, 2)
'        Print #1, "111300000440145R$ " & Alinhar(VarPreco(2), 6, 2)
'        Print #1, "111300000440280R$ " & Alinhar(VarPreco(3), 6, 2)
                            
'        Print #1, "111300000440055" & VarPreco2(1)
'        Print #1, "111300000440195" & VarPreco2(2)
'        Print #1, "111300000440335" & VarPreco2(3)
        
'        Print #1, "1A3103000030020" & VarBarra(1)
'        Print #1, "1A3103000030150" & VarBarra(2)
'        Print #1, "1A3103000030295" & VarBarra(3)
        
'                    Print #1, "191200100000045" & VarBarra(1)
'                    Print #1, "191200100000195" & VarBarra(2)
'                    Print #1, "191200100000320" & VarBarra(3)
        For Cont = 1 To 3
            VarNome1(Cont) = ""
            VarNome2(Cont) = ""
            VarBarra(Cont) = ""
            VarPreco(Cont) = ""
            VarPreco2(Cont) = ""
        Next
    Call closeport
   Wend

End If
        
        
        
 
End Sub

Private Sub BtImprimir_Click()

'MDIPrincipal.ComRel.ShowPrinter
MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelEtq.rpt")
Imprimir

End Sub

Private Sub BtLimpa_Click()
If MsgBox("Limpar a Lista ?", vbYesNo, App.Title) = vbYes Then
    If Datdados(0).Recordset.RecordCount > 0 Then
        Datdados(0).Recordset.MoveFirst
        While Not Datdados(0).Recordset.Eof
            Datdados(0).Recordset.Delete
            Datdados(0).Recordset.MoveNext
        Wend
        Datdados(0).Refresh
    End If
End If
DbcPesquisa.Text = ""
MebQtd = ""
MebPcl = 0
MebVlrPcl = 0
DbcPesquisa.SetFocus

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtPed_Click()
Dim X As Integer
Dim VarQtd As Integer
Dim flag As Boolean
flag = True
If MebPedido.Text <> "" Then
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PED.Operacao, PED.DataEmiss , Itp.Qtd, Prd.Estoque FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE CodPed = " & MebPedido & " ORDER BY Itp.CodItp "
    Datdados(1).Refresh
    If Datdados(1).Recordset.RecordCount > 0 Then
        
        If MsgBox("Emitir Pela Quantidade ?", vbYesNo) = vbNo Then
            VarQtd = 1
            flag = False
        End If
        
        
        Datdados(1).Recordset.MoveFirst
        While Not Datdados(1).Recordset.Eof
            If flag Then
                VarQtd = Datdados(1).Recordset("Qtd")
            End If
            
            For X = 1 To VarQtd
                Datdados(0).Recordset.AddNew
                Datdados(0).Recordset.Fields("Produto") = Datdados(1).Recordset("CodPrd")
                Datdados(0).Recordset.Fields("Qtd") = Datdados(1).Recordset("Qtd")
                If ChkDesconto.value = 1 Then
                    Datdados(0).Recordset.Fields("Preco") = (Datdados(1).Recordset.Fields("Venda") - (Datdados(1).Recordset.Fields("Venda") * (MebDesc / 100)))
                Else
                    Datdados(0).Recordset.Fields("Preco") = Datdados(1).Recordset.Fields("Venda")
                End If
                Datdados(0).Recordset.Fields("Parcela") = 1
                Datdados(0).Recordset.Fields("Valor") = Datdados(1).Recordset.Fields("Venda")
                Datdados(0).Recordset.Update
            Next
            Datdados(1).Recordset.MoveNext
        Wend
    
    End If
    
End If
End Sub

Private Sub chkcadastro_Click()
Dim SQL As String
If chkcadastro.value = 1 Then
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PRD.DtAlt, prd.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE PRD.DtAlt =#" & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PRD.Descrição_Produto"
    Datdados(1).Refresh
Else
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, prd.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY PRD.CodPrd desc"
    Datdados(1).Refresh
End If

'sql = "SELECT PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PRD.DtAlt FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE PRD.DtAlt =#" & Format(DtData, "mm/dd/yyyy") & "# ORDER BY PRD.Descrição_Produto"
DbcPesquisa.Refresh

End Sub

Private Sub ChkEntradas_Click()
If ChkEntradas.value = 1 Then
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PED.Operacao, PED.DataEmiss, Prd.Estoque FROM (GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo) INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto WHERE (((PED.Operacao)=0) AND ((PED.DataEmiss)=Date())) ORDER BY PRD.Descrição_Produto "
    Datdados(1).Refresh
Else
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, Prd.EStoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY PRD.Descrição_Produto"
    Datdados(1).Refresh
End If
DbcPesquisa.Refresh
End Sub

Private Sub cmdcompl_Click()
If TxtComplemento <> "" Then
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PRD.DtAlt, Prd.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE Complemento = " & """" & TxtComplemento & """" & " ORDER BY PRD.Descrição_Produto"
    Datdados(1).Refresh
Else
    Datdados(1).RecordSource = "SELECT PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, PRD.DtAlt, PRD.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE PRD.Descrição_Produto LIKE " & """*" & DbcPesquisa.Text & "*""" & " ORDER BY PRD.Descrição_Produto"
    Datdados(1).Refresh

End If

End Sub

Private Sub CmdFicha_Click()
Dim y As Integer
Dim VarEvento As String
Dim VarChave As String

VarEvento = LerINI("Config", "Evento", App.Path & "\config.ini")

VarChave = EnCripta(VarEvento)

Iniciar (False)
    
    
    If Datdados(0).Recordset.RecordCount > 0 Then
        If MsgBox("Confirma impressão das fichas " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
           Datdados(0).Recordset.MoveFirst
'           Open VarImprBarra For Output As #1
           While Not Datdados(0).Recordset.Eof

                Print #1, "========================================"
                Print #1, Centralizar(Cabecalho1, 40)
                Print #1, Centralizar(VarEvento & " - " & Date, 40)
                    
                Print #1, "----------------------------------------"
                Print #1, ""
                Print #1, Tab(1); "1.00"; "  "; Left(Datdados(0).Recordset("Descrição_Produto"), 33)
                
                Print #1, ""
                Print #1, "========================================"
                Print #1, VarChave
                
                For y = 1 To 4
                    Print #1, ""
                Next
                Print #1, Chr(27) & "i" '& Chr(27) & "v50"
                Datdados(0).Recordset.MoveNext
           Wend
        End If
    End If
    Fim (False)

End Sub

Private Sub CmdLista_Click()

Dim y As Integer

        
    If Datdados(0).Recordset.RecordCount > 0 Then
        If MsgBox("Confirma impressão ?", vbYesNo, App.Title) = vbYes Then
           Iniciar (False)
           Datdados(0).Recordset.MoveFirst

           While Not Datdados(0).Recordset.Eof

                Print #1, Tab(1); Datdados(0).Recordset("Produto"); "-"; Left(Datdados(0).Recordset("Descrição_Produto"), 30) & "  " & Format(Datdados(0).Recordset("Preco"), "###,##0.00")
                
                Datdados(0).Recordset.MoveNext
           Wend
            For y = 1 To 4
                Print #1, ""
            Next
            Print #1, Chr(27) & "i" '& Chr(27) & "v50"
           
           
           Fim (False)
        End If
    End If



End Sub

Private Sub CmdTinta_Click()
Dim Cont As Byte
Dim VarNome1(1 To 4) As String
Dim VarNome2(1 To 4) As String
Dim VarBarra(1 To 4) As String
Dim VarPreco(1 To 4) As String
Dim VarVista(1 To 4) As String
Dim VarPreco2(1 To 4) As String

Cont = 1
For Cont = 1 To 4
    VarNome1(Cont) = ""
    VarNome2(Cont) = ""
    VarBarra(Cont) = ""
    VarPreco(Cont) = ""
Next
Cont = 1

        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0000"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
'                    Print #1, "191200301600005" & Left(Cabecalho1, 15)'
'                    Print #1, "191200301600145" & Left(Cabecalho1, 15)'
'                    Print #1, "191200301600280" & Left(Cabecalho1, 15)
'                    For Cont = 1 To 3
                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 40)
'                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("Produto"), 8)
                                VarVista(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
                                VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2) & " =" & Alinhar(Datdados(0).Recordset("Valor") * Datdados(0).Recordset("Parcela"), 6, 2)
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarVista(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
'                    Next
                                
                    Print #1, "191200500700030" & VarNome1(1)
                    Print #1, "1E2203000100280" & VarBarra(1) ' Codigo Barras
                    Print #1, "191200200450030" & Trim(txtlote)
                    Print #1, "191200400000030" & "AMPLACOR"
            '        If ChkDados(3).value = True Then
'                        Print #1, "142300000000195" & "AMPLACOR" ' Valor do Produto
            '        Else
            '            Print #1, "142300000000195" & Alinhar(MebDados(23), 6, 2) ' Valor do Produto
            '        End If
            '        Print #1, "191100100400020" & psCompString(3, NomeEmpresa, 20)
                    Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
                    Print #1, "E"
                                
                                
               Wend
               Close #1
            End If
        End If


End Sub

Private Sub DbcPesquisa_Change()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"


If DbcPesquisa <> "" And Not IsNumeric(DbcPesquisa.Text) Then
        Criterio = "Prd.descrição_produto Like " & """" & DbcPesquisa & "*"""
        Datdados(1).Recordset.FindFirst Criterio
        If Datdados(1).Recordset.NoMatch Then
'            Beep
'            DbcPesquisa.SetFocus
        Else
            If Len(DbcPesquisa) >= 1 Then
                Datdados(1).Recordset.FindFirst Criterio
                MebCusto = Datdados(1).Recordset("Custo")
                MebVenda = Datdados(1).Recordset("Venda")
                MebEst = Datdados(1).Recordset("Estoque")
            End If
        End If
End If

End Sub

Private Sub DbcPesquisa_Click(Area As Integer)
'DbcPesquisa_Change
End Sub

Private Sub DbcPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"
If KeyCode = 13 Then
    If IsNumeric(DbcPesquisa.Text) Then
        If Len(DbcPesquisa) > VarTamBarra Then
            Datdados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto ORDER BY BAR.BarBarra"
            Datdados(4).Refresh
            Criterio = "BarBarra = "
            Criterio = Criterio & Terminacao & DbcPesquisa.Text & Terminacao
            Datdados(4).Recordset.FindFirst Criterio
            If Datdados(4).Recordset.NoMatch Then
                DbcPesquisa.Text = ""
                DbcPesquisa.SetFocus
            Else
                Criterio = "CodPrd = " & Datdados(4).Recordset("Produto")
            
                Datdados(1).Recordset.FindFirst Criterio
                DbcPesquisa.Text = Datdados(1).Recordset.Fields("Descrição_Produto")
    '            EntraQtd
                MebQtd.SetFocus
            
    '            Criterio = "CodPrd = "
    '            Criterio = Criterio & DatDados(4).Recordset.Fields("Produto")
    '            DatDados(1).Recordset.FindFirst Criterio
    '            DatDados(4).RecordSource = "SELECT BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((BAR.Produto)=" & TxtDados(0) & ")) ORDER BY BAR.BarBarra"
    '            DatDados(4).Refresh
    '            MostraDados
    '            PosMovimentacao
            End If
        Else
    '        Criterio = "Prd.CodPrd = " & DbcPesquisa.BoundText
    '        DatDados(1).Recordset.FindFirst Criterio
    '        If DatDados(1).Recordset.NoMatch Then
    '            Beep
    '            DbcPesquisa.Enabled = True
    '            DbcPesquisa.Text = ""
    '            DbcPesquisa.SetFocus
    '        Else
    '            If Len(DbcPesquisa) >= 1 Then
    '                Registros.FindFirst Criterio
    '                MostraDados
    '                PosMovimentacao
    '            End If
    '        End If
        End If
    End If
End If
End Sub



Private Sub DbcPesquisa_LostFocus()
Dim Criterio As String
Dim Terminacao As String
Terminacao = """"

If IsNumeric(DbcPesquisa.Text) Then
        Criterio = "Prd.CodPrd = " & DbcPesquisa.BoundText
        Datdados(1).Recordset.FindFirst Criterio
        If Datdados(1).Recordset.NoMatch Then
            Beep
            DbcPesquisa.Text = ""
            DbcPesquisa.SetFocus
        Else
            If Len(DbcPesquisa) >= 1 Then
                Datdados(1).Recordset.FindFirst Criterio
                DbcPesquisa.Text = Datdados(1).Recordset.Fields("Descrição_Produto")
                EntraQtd
            End If
        End If
End If
End Sub

Private Sub EntraQtd()
MebQtd = ""
MebPcl.Text = 0
MebVlrPcl.Text = 0
MebQtd.SetFocus
MebQtd.SetFocus
End Sub

Private Sub DBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim Criterio As String
Dim Indice As Double
If ColIndex = 5 Then

    If IsNumeric(DBGrid.Columns(5)) And DBGrid.Columns(5) <> "0" Then
        If Ind.RecordCount > 0 Then
            If chkprazo = 0 Then
                Criterio = "NroIndice = " & DBGrid.Columns(5)
            Else
                Criterio = "NroIndice = 3"
            End If
            Ind.FindFirst Criterio
            If Not Ind.NoMatch Then
                Indice = Ind!Indice
            Else
                Indice = 1
            End If
        End If

    
        If Opt.value = 0 Then
'            MebVlrPcl.Text = CCur(Datdados(1).Recordset.Fields("Venda")) / CCur(MebPcl)
        Else
            If VarPzoProduto = True Then
                DBGrid.Columns(6) = (DBGrid.Columns(4) * CCur(Indice))
            Else
                If chkprazo = 0 Then
                    DBGrid.Columns(6) = (DBGrid.Columns(4) * CCur(Indice))
                Else
                    DBGrid.Columns(6) = (DBGrid.Columns(4) * CCur(Indice))
                End If
            End If
        End If
    
        DBGrid.Columns(4) = DBGrid.Columns(6) * DBGrid.Columns(5)
    End If
End If

If ColIndex = 6 Then

    If IsNumeric(DBGrid.Columns(6)) And DBGrid.Columns(6) <> "0" Then
    
            DBGrid.Columns(4) = DBGrid.Columns(6) * DBGrid.Columns(5)
    
    End If
End If
End Sub

Private Sub DbGrid_dblClick()
If Datdados(0).Recordset.RecordCount <> 0 Then
    Datdados(0).Recordset.Delete
    Datdados(0).Refresh
End If
End Sub

Private Sub EtqMercado_Click()
Dim Cont As Byte
Dim VarNome1(1 To 4) As String
Dim VarNome2(1 To 4) As String
Dim VarBarra(1 To 4) As String
Dim VarPreco(1 To 4) As String
Dim VarVista(1 To 4) As String
Dim VarPreco2(1 To 4) As String

Cont = 1
For Cont = 1 To 4
    VarNome1(Cont) = ""
    VarNome2(Cont) = ""
    VarBarra(Cont) = ""
    VarPreco(Cont) = ""
Next
Cont = 1

        If Datdados(0).Recordset.RecordCount > 0 Then
            If MsgBox("Confirma impressão das etiquetas na porta " & VarImprBarra & " ?", vbYesNo, App.Title) = vbYes Then
               Datdados(0).Recordset.MoveFirst
               Open VarImprBarra For Output As #1
               While Not Datdados(0).Recordset.Eof
                    Print #1, Chr(2) + "O0000"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
'                    Print #1, "191200301600005" & Left(Cabecalho1, 15)'
'                    Print #1, "191200301600145" & Left(Cabecalho1, 15)'
'                    Print #1, "191200301600280" & Left(Cabecalho1, 15)
'                    For Cont = 1 To 3


                        If Not Datdados(0).Recordset.Eof Then
                            If Not Datdados(0).Recordset.NoMatch Then
                                VarNome1(Cont) = Left(Datdados(0).Recordset("Descrição_Produto"), 40)
'                                VarNome2(Cont) = Mid(Datdados(0).Recordset("Descrição_Produto"), 16, 15)
                                VarBarra(Cont) = psCompString(1, Datdados(0).Recordset("codbar"), 13)
                                VarVista(Cont) = Alinhar(Datdados(0).Recordset("Preco"), 6, 2)
'                                VarPreco(Cont) = Format(Datdados(0).Recordset("Parcela"), "##") & " X" & Alinhar(Datdados(0).Recordset("Valor"), 6, 2) & " =" & Alinhar(Datdados(0).Recordset("Valor") * Datdados(0).Recordset("Parcela"), 6, 2)
                            Else
                                VarNome1(Cont) = ""
                                VarNome2(Cont) = ""
                                VarBarra(Cont) = ""
                                VarVista(Cont) = ""
                                VarPreco(Cont) = ""
                            End If
                            Datdados(0).Recordset.MoveNext
                        Else
                            VarNome1(Cont) = ""
                            VarNome2(Cont) = ""
                            VarBarra(Cont) = ""
                            VarPreco(Cont) = ""
                        End If
'                    Next
                                

'        Print #1, "191200500500005" & psCompString(3, txtDados(1), 40)
'        Print #1, "1F1202000100015" & psCompString(1, DBGrid1.Columns(1), 13)  ' Codigo Barras''

'        Print #1, "192300400000125" & "R$"
'        If ChkDados(3).value = True Then
'            Print #1, "142300000000195" & Alinhar(MebDados(16), 6, 2) ' Valor do Produto
'        Else
'            Print #1, "142300000000195" & Alinhar(MebDados(23), 6, 2) ' Valor do Produto
'        End If
'        Print #1, "191100100400020" & psCompString(3, NomeEmpresa, 20)
'        Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
'        Print #1, "E"
                                
                                
                    Print #1, "191100500800020" & VarNome1(1)
                    Print #1, "1F1202000200020" & VarBarra(1) ' Codigo Barras
                    Print #1, "192300400100160" & "R$"
            '        If ChkDados(3).value = True Then
                        Print #1, "142300000100210" & VarVista(1) ' Valor do Produto
            '        Else
            '            Print #1, "142300000000195" & Alinhar(MebDados(23), 6, 2) ' Valor do Produto
            '        End If
            '        Print #1, "191100100400020" & psCompString(3, NomeEmpresa, 20)
                    Print #1, "191100100500020" & psCompString(3, NomeEmpresa, 20)

                    Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
                    Print #1, "E"
                                
                                
               Wend
               Close #1
            End If
        End If


End Sub

Private Sub Form_Activate()
Set Formulário = Me ' Represente o formulário Atual, só pode ser usada dentro do próprio Formulário

Formulário.Left = 300 ' (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = 400 ' (MDIPrincipal.Height - Formulário.Height) / 6

DbcPesquisa.SetFocus

MebPcl = 0
MebVlrPcl = 0
MebDesc = 0

End Sub


Private Sub Form_Load()
Dim Componentes As Variant
Dim u As Integer
Dim Cn As String
'Dim Rs As ADODB.Recordset

For Each Componentes In Datdados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
Datdados(0).RecordSource = "SELECT ETQ.Codigo, ETQ.Produto, PRD.Descrição_Produto, Prd.Complemento, Prd.CodBar, ETQ.Qtd, Etq.Preco, Etq.Parcela, Etq.Valor FROM ETQ INNER JOIN PRD ON ETQ.Produto = PRD.CodPrd ORDER BY ETQ.Codigo"
Datdados(0).Refresh

Datdados(1).RecordSource = "SELECT PRD.PcoPrazo, PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda, Prd.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE PRD.Ativo=True ORDER BY PRD.Descrição_Produto"
Datdados(1).Refresh

Datdados(4).RecordSource = "SELECT PRD.PcoPrazo, PRD.CodPrd, PRD.Descrição_Produto, Prd.Complemento, GRU.Descrição_Grupo, PRD.Custo, PRD.Margem, PRD.Venda FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE PRD.Ativo=True ORDER BY PRD.Descrição_Produto"
Datdados(4).Refresh


Set Ind = Banco.OpenRecordset("SELECT * FROM Ind Order By NroIndice")

    
'Set Rs = New ADODB.Recordset

'ReDim Etiqs(0)
'u = UBound(Etiqs)

'cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Caminho & "\Dados.MDB"
' Carrega os modelos de etiquetas
'Rs.Open "SELECT * FROM Etiquetas ORDER BY Index", cn, adOpenStatic
'While Not Rs.Eof
'    ReDim Preserve Etiqs(u)
'    cmbModelo.AddItem Rs("Modelo")
'    With Etiqs(u)
'        .Altura = Rs("Altura")
'        .Largura = Rs("Largura")
'        .MEsq = Rs("MargemEsq")
 '       .MSup = Rs("MargemSup")
'        .MInf = Rs("MargemInf")
'        .Colunas = Rs("Colunas")
'    End With
'    u = u + 1
'    Rs.MoveNext
'Wend

'If cmbModelo.ListCount > 0 Then cmbModelo.ListIndex = 0

'Rs.Close
'Set Rs = Nothing
DtData = Date

End Sub


Private Sub cmbModelo_Click()
    
    If cmbModelo.ListIndex <> -1 Then
        With Etiqs(cmbModelo.ListIndex)
            txtAltura.Text = .Altura & ""
            txtLargura.Text = .Largura & ""
            txtMEsq.Text = .MEsq & ""
            txtMSup.Text = .MSup & ""
            txtMInf.Text = .MInf & ""
            txtColunas.Text = .Colunas
        End With
    End If
    
End Sub

Private Sub CmdImprimir_Click()

    With rlt_Etiquetas
        .Altura = CDbl(txtAltura.Text)
        .Largura = CDbl(txtLargura.Text)
        .MargemEsquerda = CDbl(txtMEsq.Text)
        .MargemSuperior = CDbl(txtMSup.Text)
        .MargemInferior = CDbl(txtMInf.Text)
        .Colunas = CInt(txtColunas.Text)
        .LinhaInicio = CInt(txtLinIni.Text)
        .ColunaInicio = CInt(txtColIni.Text)
        .Relatorio.Visualizar = (chkVisualizar.value = vbChecked)
        ' Indica que deverá selecionar a impressora
        ' antes de iniciar o relatorio, se nao passar pelo preview
        If Not .Relatorio.Visualizar Then
            .Relatorio.SelecionarImpressora = True
        End If
        .Config
    End With
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
'        KeyPress = 0
    End If
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

Private Sub MebPcl_LostFocus()
Dim Indice As Currency
Dim Criterio As String

If MebPcl.Text <> "" Then
    If MebPcl <> 0 Then
        If Ind.RecordCount > 0 Then
            If chkprazo = 0 Then
                Criterio = "NroIndice = " & MebPcl
            Else
                Criterio = "NroIndice = 3"
            End If
            Ind.FindFirst Criterio
            If Not Ind.NoMatch Then
                Indice = Ind!Indice
            Else
                Indice = 1
            End If
        End If
        If Opt.value = 0 Then
            MebVlrPcl.Text = CCur(Datdados(1).Recordset.Fields("Venda")) / CCur(MebPcl)
        Else
            If VarPzoProduto = True Then
                MebVlrPcl.Text = CCur(Datdados(1).Recordset.Fields("PcoPrazo")) * CCur(Indice)
            Else
                If chkprazo = 0 Then
                    MebVlrPcl.Text = CCur(Datdados(1).Recordset.Fields("Venda")) * CCur(Indice)
                Else
'                    If MebPcl > 1 Then
                        MebVlrPcl.Text = CCur(Datdados(1).Recordset.Fields("Venda") + (Datdados(1).Recordset.Fields("Venda") * CCur(Ind!Perc / 100)))
'                    Else
'                        MebVlrPcl.Text = CCur(DatDados(1).Recordset.Fields("Venda") + (DatDados(1).Recordset.Fields("Venda") * CCur(Ind!Perc / 100)))
'                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Public Function EnCripta(strText As String) As String
Dim i As Long, c As Long
Dim strBuff As String
Dim strPwd As String '
strBuff = ""
strPwd = "SR2090"

If Len(strPwd) Then
  For i = 1 To Len(strText)
  c = Asc(Mid$(strText, i, 1))
  c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
  strBuff = strBuff & c - 100 ' Chr$(c) ' And &HFF)
  Next i
Else
  strBuff = strText
End If

EnCripta = strBuff

End Function


