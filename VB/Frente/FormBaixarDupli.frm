VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FormBaixarDupli 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkvencidos 
      Caption         =   "&Somente os Vencidos"
      Height          =   285
      Left            =   9630
      TabIndex        =   65
      Top             =   5715
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -90
      ScaleHeight     =   330
      ScaleWidth      =   12045
      TabIndex        =   64
      Top             =   8190
      Width           =   12075
   End
   Begin VB.Frame Frame3 
      Height          =   2040
      Left            =   225
      TabIndex        =   53
      Top             =   5850
      Width           =   3480
      Begin VB.ComboBox LstTipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FormBaixarDupli.frx":0000
         Left            =   1845
         List            =   "FormBaixarDupli.frx":0016
         TabIndex        =   54
         Top             =   1560
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   360
         Left            =   135
         TabIndex        =   55
         Top             =   1530
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   99024897
         CurrentDate     =   36759
      End
      Begin MSMask.MaskEdBox MebDesconto 
         DataField       =   "Custo"
         Height          =   360
         Left            =   135
         TabIndex        =   56
         Tag             =   "S"
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebJuros 
         DataField       =   "Custo"
         Height          =   360
         Left            =   135
         TabIndex        =   57
         Tag             =   "S"
         Top             =   930
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPago 
         DataField       =   "Custo"
         Height          =   360
         Left            =   1845
         TabIndex        =   58
         Tag             =   "S"
         Top             =   930
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data "
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   63
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   62
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Juros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   61
         Top             =   750
         Width           =   375
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         Height          =   195
         Index           =   3
         Left            =   1845
         TabIndex        =   60
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pgto"
         Height          =   195
         Left            =   1845
         TabIndex        =   59
         Top             =   1350
         Width           =   690
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
      ScaleWidth      =   11940
      TabIndex        =   51
      Top             =   0
      Width           =   11940
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Baixar Conta"
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
         TabIndex        =   52
         Top             =   90
         Width           =   4020
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   8325
      TabIndex        =   26
      Top             =   6120
      Width           =   3300
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   27
         Tag             =   "S"
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebPagos 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   28
         Tag             =   "S"
         Top             =   870
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebAPagar 
         DataField       =   "Custo"
         Height          =   270
         Left            =   135
         TabIndex        =   29
         Tag             =   "S"
         Top             =   1410
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotJuros 
         DataField       =   "Custo"
         Height          =   270
         Left            =   1740
         TabIndex        =   30
         Tag             =   "S"
         Top             =   870
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotDesc 
         DataField       =   "Custo"
         Height          =   270
         Left            =   1740
         TabIndex        =   31
         Tag             =   "S"
         Top             =   1425
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotGeral 
         DataField       =   "Custo"
         Height          =   270
         Left            =   1740
         TabIndex        =   32
         Tag             =   "S"
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   38
         Top             =   135
         Width           =   540
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas pagas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   37
         Top             =   675
         Width           =   960
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas á pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   36
         Top             =   1245
         Width           =   1050
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   8
         Left            =   1680
         TabIndex        =   35
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Juros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   9
         Left            =   1680
         TabIndex        =   34
         Top             =   690
         Width           =   345
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   10
         Left            =   1680
         TabIndex        =   33
         Top             =   135
         Width           =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3240
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3270
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1620
         X2              =   1620
         Y1              =   120
         Y2              =   1800
      End
   End
   Begin VB.CommandButton CmdPedido 
      Caption         =   "&Imprimi &Pedido"
      Height          =   360
      Left            =   6630
      TabIndex        =   18
      Top             =   6630
      Width           =   1275
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7470
      Width           =   1275
   End
   Begin VB.Frame FrmEnt 
      Height          =   825
      Left            =   225
      TabIndex        =   11
      Top             =   720
      Width           =   11475
      Begin VB.TextBox TxtCod 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   4455
         TabIndex        =   15
         Top             =   390
         Width           =   690
      End
      Begin VB.Data DatDados 
         Caption         =   "Ent"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   1080
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox ChkPagos 
         Caption         =   "&Mostrar Pagos"
         Height          =   285
         Left            =   8955
         TabIndex        =   4
         Top             =   390
         Width           =   1410
      End
      Begin VB.CheckBox ChkJuros 
         Caption         =   "&Cobrar Juros"
         Height          =   285
         Left            =   7470
         TabIndex        =   3
         Top             =   390
         Width           =   1275
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&OK"
         Height          =   330
         Left            =   10575
         TabIndex        =   5
         Top             =   345
         Width           =   600
      End
      Begin MSMask.MaskEdBox MebDesc 
         DataField       =   "Custo"
         Height          =   315
         Left            =   5220
         TabIndex        =   1
         Tag             =   "S"
         Top             =   390
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Mebdias 
         DataField       =   "Custo"
         Height          =   315
         Left            =   6255
         TabIndex        =   2
         Tag             =   "S"
         Top             =   390
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormBaixarDupli.frx":006F
         DataField       =   "Nome"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   390
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
         Height          =   195
         Left            =   4455
         TabIndex        =   16
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entidade"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   165
         Width           =   630
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto(%)"
         Height          =   195
         Index           =   6
         Left            =   5220
         TabIndex        =   13
         Top             =   165
         Width           =   900
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "&Considerar Dias"
         Height          =   195
         Index           =   7
         Left            =   6255
         TabIndex        =   12
         Top             =   165
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   690
      Left            =   4095
      TabIndex        =   8
      Top             =   7095
      Width           =   2175
      Begin MSMask.MaskEdBox MebTotal 
         Height          =   510
         Left            =   45
         TabIndex        =   9
         Top             =   135
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483640
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton CmbImprimir 
      Caption         =   "&Imprimir Conta"
      Height          =   360
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Extrato"
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa Tela"
      Height          =   360
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7050
      Width           =   1275
   End
   Begin VB.CommandButton CmdPgto 
      Caption         =   "&Confirma Pgto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4140
      TabIndex        =   39
      Top             =   6210
      Width           =   2085
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   180
      TabIndex        =   19
      Top             =   1620
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   7170
      _Version        =   393216
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Parcelas"
      TabPicture(0)   =   "FormBaixarDupli.frx":0089
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbgrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BtPagarTodos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DatDados(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptOrdem(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "OptOrdem(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "OptOrdem(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "OptOrdem(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Itens Pedido"
      TabPicture(1)   =   "FormBaixarDupli.frx":00A5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDados(5)"
      Tab(1).Control(1)=   "LblDados(11)"
      Tab(1).Control(2)=   "MebTotalPed"
      Tab(1).Control(3)=   "MebDescPed"
      Tab(1).Control(4)=   "Lista"
      Tab(1).Control(5)=   "CmbProximo"
      Tab(1).Control(6)=   "CmdAnterior"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Relatório"
      TabPicture(2)   =   "FormBaixarDupli.frx":00C1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LstPedido"
      Tab(2).ControlCount=   1
      Begin VB.ListBox LstPedido 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   -74910
         TabIndex        =   50
         Top             =   480
         Width           =   11265
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pagamento"
         Height          =   285
         Index           =   3
         Left            =   3645
         TabIndex        =   49
         Top             =   3690
         Width           =   1365
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Vencimento"
         Height          =   285
         Index           =   2
         Left            =   2250
         TabIndex        =   48
         Top             =   3690
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pedido"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   47
         Top             =   3690
         Width           =   1050
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Emissão"
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   46
         Top             =   3690
         Width           =   1095
      End
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64110
         TabIndex        =   41
         ToolTipText     =   "Prôximo pedido"
         Top             =   1035
         Width           =   465
      End
      Begin VB.Data DatDados 
         Caption         =   "Dupli"
         Connect         =   "Access"
         DatabaseName    =   "C:\aateste\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   1530
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1770
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.CommandButton BtPagarTodos 
         Caption         =   "&Selecionar Todos"
         Height          =   285
         Left            =   9840
         TabIndex        =   23
         Top             =   3675
         Width           =   1500
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00C0FFFF&
         Height          =   825
         Left            =   4470
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   3120
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Aguarde . . . "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   330
            Left            =   810
            TabIndex        =   22
            Top             =   270
            Width           =   1755
         End
      End
      Begin VB.CommandButton CmbProximo 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64110
         TabIndex        =   20
         ToolTipText     =   "Prôximo pedido"
         Top             =   480
         Width           =   465
      End
      Begin MSDBGrid.DBGrid dbgrid1 
         Bindings        =   "FormBaixarDupli.frx":00DD
         Height          =   3225
         Left            =   90
         OleObjectBlob   =   "FormBaixarDupli.frx":00F7
         TabIndex        =   24
         Top             =   390
         Width           =   11265
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   3105
         Left            =   -74850
         TabIndex        =   40
         Top             =   420
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   5477
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CodPrd"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Embalagem"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qtd"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Unitário"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Valor"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Valor Dia"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSMask.MaskEdBox MebDescPed 
         DataField       =   "Custo"
         Height          =   315
         Left            =   -67080
         TabIndex        =   42
         Tag             =   "S"
         Top             =   3600
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebTotalPed 
         DataField       =   "Custo"
         Height          =   315
         Left            =   -65415
         TabIndex        =   44
         Tag             =   "S"
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Height          =   180
         Index           =   11
         Left            =   -65820
         TabIndex        =   45
         Top             =   3645
         Width           =   300
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
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
         Height          =   180
         Index           =   5
         Left            =   -67800
         TabIndex        =   43
         Top             =   3645
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Itens Selecionados:"
         Height          =   195
         Left            =   7650
         TabIndex        =   25
         Top             =   3705
         Width           =   1395
      End
   End
   Begin VB.Label LblPgto 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   66
      Top             =   5715
      Width           =   5505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4140
      TabIndex        =   10
      Top             =   6915
      Width           =   1170
   End
End
Attribute VB_Name = "FormBaixarDupli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas
Public TotJuros As Currency
Public TotDesc As Currency
Public VarCont As Integer
Public VarOrdem  As String
Private Sub BtPagarTodos_Click()

With DatDados(0).Recordset
    
    If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF
             DBGrid1_DBLClick
            .MoveNext
        Wend
    End If

End With

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub ChkPagos_Click()
If DbcPesquisa.Text <> "" Then
    Pesquisa DbcPesquisa.BoundText
End If
End Sub

Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub CmbImprimir_Click()
Dim Resp  As String
Dim x As Byte
Dim VarVenc As Boolean
VarVenc = True
Dim VarSoma As Currency
VarSoma = 0

Resp = MsgBox("Impressão resumida ? ", vbYesNoCancel, App.Title)
If Resp = vbYes Then
    Iniciar (False)
    
    
    Print #1, "" + Cond
    Print #1, ""
    Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
    Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 23)
    Print #1, "=========================================="
    Print #1, "Data:" & Date & " Hora:" & Time; Tab(33)
    Print #1, Tab(1); psCompString(1, TxtCod, 6); "-"; Left(DbcPesquisa.Text, 30)
    Print #1, Tab(1); "Ped."; Tab(9); "Pcl"; Tab(14); "Emiss."; Tab(26); "VENC."; Tab(38); "VALOR"
    Print #1, "------------------------------------------"
    If CmdPgto.Enabled = True Then
        DatDados(0).Recordset.MoveFirst
        While Not DatDados(0).Recordset.EOF
            If DatDados(0).Recordset("PgtTemp") = -1 Then
                Print #1, Tab(1); psCompString(1, DatDados(0).Recordset("Pedido"), 6); Tab(8); psCompString(1, DatDados(0).Recordset("NroPcl"), 3); Tab(12); DatDados(0).Recordset("DtEmiss"); Tab(24); DatDados(0).Recordset("DtVenc"); Tab(35); Alinhar(Format(DatDados(0).Recordset("Valor"), "##,##0.00"), 8, 2) '; Tab(45); DatDados(0).Recordset("Tipo_Dup")
                VarSoma = VarSoma + DatDados(0).Recordset("Valor")
            End If
            DatDados(0).Recordset.MoveNext
        Wend
        Print #1, Tab(33); "----------"
        Print #1, Tab(16); "TOTAL     => "; Tab(33); Alinhar(CStr(VarSoma), 10, 2)
    
    Else
        DatDados(0).Recordset.MoveFirst
        While Not DatDados(0).Recordset.EOF
            Print #1, Tab(1); psCompString(1, DatDados(0).Recordset("Pedido"), 6); Tab(8); psCompString(1, DatDados(0).Recordset("NroPcl"), 3); Tab(12); DatDados(0).Recordset("DtEmiss"); Tab(24); DatDados(0).Recordset("DtVenc"); Tab(35); Alinhar(Format(DatDados(0).Recordset("Valor"), "##,##0.00"), 8, 2) '; Tab(45); DatDados(0).Recordset("Tipo_Dup")
            DatDados(0).Recordset.MoveNext
        Wend
        Print #1, Tab(33); "----------"
        Print #1, Tab(16); "SUB-TOTAL =>"; Tab(33); Alinhar(CStr(MebSubTot.Text), 10, 2)
        Print #1, Tab(16); "JUROS     =>"; Tab(33); Alinhar(CStr(MebTotJuros.Text), 10, 2)
        Print #1, Tab(16); "DESCONTO  =>"; Tab(33); Alinhar(CStr(MebTotDesc.Text), 10, 2)
        Print #1, Tab(16); "TOTAL     =>"; Tab(33); Alinhar(CStr(MebTotGeral.Text), 10, 2)
    
    End If
    
    Print #1, "" & Nor
    For x = 1 To NroLinha
        Print #1, ""
    Next
'    Print #1, ""
    
    Fim (False)
ElseIf Resp = vbNo Then
    ImprimiDiscr TxtCod
End If
End Sub

Private Sub CmbProximo_Click()

If Not DatDados(0).Recordset.EOF Then
    DatDados(0).Recordset.MoveNext
    SSTab1_Click 1
Else
    MsgBox "Final do arquivo !", vbInformation, App.Title
End If

End Sub

Private Sub CmdAnterior_Click()

If Not DatDados(0).Recordset.BOF Then
    DatDados(0).Recordset.MovePrevious
    SSTab1_Click 1
Else
    MsgBox "Inicio do arquivo !", vbInformation, App.Title
End If

End Sub

Private Sub CmdFechar_Click()
Unload Me
End Sub

Private Sub CmdLimpa_Click()
MebPago = 0

DbcPesquisa = ""
MebSubTot = 0
MebPagos = 0
MebDesconto = 0
MebJuros = 0
MebAPagar = 0
MebTotal = 0
MebDesc = 0
Mebdias = 0
DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade =  0 ORDER BY " & VarOrdem
DatDados(0).Refresh
SomaGrid
FrmEnt.Enabled = True
ChkJuros.Value = 1
ChkPagos.Value = 0

DbcPesquisa.SetFocus

End Sub

Private Sub CmdOk_Click()
If DbcPesquisa.Text <> "" Then
    VarCont = 0
    If Not IsNumeric(DbcPesquisa.Text) Then
        If IsNumeric(DbcPesquisa.BoundText) Then
            Pesquisa DbcPesquisa.BoundText
        End If
    Else
        If MebTotal <> 0 Then
            CmdLimpa_Click
        End If
        Pesquisa DbcPesquisa.Text
    End If
End If
End Sub

Private Sub CmdPedido_Click()
'ImprimirPedido Right(CmdPedido.Caption, 7)
ImprimirPedido DBGrid1.Columns(1)
End Sub

Private Sub CmdPgto_Click()
Dim Criterio As String
Dim ValorConta As Currency
Dim TotalPago As Currency
Dim Desconto As Currency
Dim Juros As Currency
Dim Valor As Currency
Dim ValorPcl As Currency
Dim sql As String
Dim Dupli As Long
Dim DtEmissAnt As Date
Dim DtVencAnt As Date
Dim VarDoc As String
Dim VarParc As Double
Dim VarPed As Double
Dim VarSaldo As Boolean

If LerINI("PEDIDO", "NaoImprimirSaldo", App.Path & "\config.ini") = "" Then
    VarSaldo = True
Else
    If LerINI("PEDIDO", "NaoImprimirSaldo", App.Path & "\config.ini") = "0" Then
        VarSaldo = True
    Else
        VarSaldo = False
    End If
End If


If Not IsNumeric(MebDesconto) Then
    MsgBox "Campo desconto inválido", vbInformation
    MebDesconto = 0
    Exit Sub
End If

If Not IsNumeric(MebJuros) Then
    MsgBox "Campo juros inválido", vbInformation
    MebJuros = 0
    Exit Sub
End If

If Not IsNumeric(MebTotal) Then
    MsgBox "Campo valor pago inválido", vbInformation
    MebTotal = 0
    Exit Sub
End If


Dupli = 0
If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma pagamento ?", vbYesNo, App.Title) = vbYes Then
        
         DatDados(0).Recordset.MoveFirst
         ValorConta = MebTotal
         TotalPago = MebPago
         DtEmissAnt = Date
         Desconto = 0
         Juros = 0
         DtVencAnt = Date + 30
         If DatDados(0).Recordset.EOF Then
                MsgBox "Houve alguma falta de informação ! pagamento Cancelado", vbInformation, App.Title
                Exit Sub
         End If
         Do Until DatDados(0).Recordset.EOF
         
            If GetInputState() Then DoEvents
         
            If DatDados(0).Recordset("PgtTemp") = -1 Then
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("Cond") = -1
                DatDados(0).Recordset.Fields("DtPgto") = DTdata

                
                If MebDesconto.Enabled = False Then
                    ValorPcl = (DatDados(0).Recordset("Valor") + DatDados(0).Recordset("Juros")) - DatDados(0).Recordset("Desconto")
                    DatDados(0).Recordset.Fields("VlrPago") = ValorPcl
                    Juros = Juros + DatDados(0).Recordset("Juros")
                    
                Else
                    ValorPcl = (DatDados(0).Recordset("Valor") + MebJuros) - MebDesconto
                    If MebPago >= ValorPcl Then
                        DatDados(0).Recordset.Fields("VlrPago") = ValorPcl
                    Else
                        DatDados(0).Recordset.Fields("VlrPago") = MebPago
                    End If
                    Juros = Juros + MebJuros
                End If
                                                
                DatDados(0).Recordset.Update
                TotalPago = CCur(TotalPago) - CCur(ValorPcl)
                DtEmissAnt = DatDados(0).Recordset("DtEmiss")
                DtVencAnt = DatDados(0).Recordset("DtVenc")
                VarDoc = DatDados(0).Recordset("Pedido") & "/" & DatDados(0).Recordset("NroPcl")
                VarPed = DatDados(0).Recordset("Pedido")
                VarParc = DatDados(0).Recordset("NroPcl")
                
                
            End If
            DatDados(0).Recordset.MoveNext
         Loop
         
         
         If TotalPago < 0 Then
            TotalPago = TotalPago * (-1)
            If MsgBox("Faltou R$ " & Format(TotalPago, "#,##0.00;- #,##0.00") & " Sim para gerar saldo ou Não para dar desconto", vbYesNo, App.Title) = vbYes Then
'                sql = "INSERT INTO DUP ( DtEmiss, NroPcl, Entidade, DtVenc, Valor, Cond,Tipo_Dup  ) VALUES ( "
                If MebDesconto.Enabled = False Then

                    Criterio = "Pedido = " & VarPed & " and NroPcl = " & VarParc
                    DatDados(0).Recordset.FindFirst Criterio
                    DatDados(0).Recordset.Edit
                    DatDados(0).Recordset.Fields("VlrPago") = DatDados(0).Recordset.Fields("Valor") - TotalPago
                    DatDados(0).Recordset.Update
                End If
                Criterio = "Pedido = " & VarPed & " and NroPcl = " & VarParc
                DatDados(0).Recordset.FindFirst Criterio
     
     
                sql = "INSERT INTO DUP ( DtEmiss, Pedido, NroPcl, Entidade, DtVenc, Valor, Cond, Documento, Tipo_Dup  ) VALUES ( "
                
                sql = sql & "#" & Format(DatDados(0).Recordset.Fields("DtEmiss"), "mm/dd/yyyy") & "# , "
'                sql = sql & "#" & Format(Date, "mm/dd/yyyy") & "# , "
                
                sql = sql & DatDados(0).Recordset.Fields("Pedido") & " , "
                sql = sql & DatDados(0).Recordset.Fields("NroPcl") & " , "
                sql = sql & DbcPesquisa.BoundText & " , #"
                
                sql = sql & Format(DatDados(0).Recordset.Fields("DtVenc"), "mm/dd/yyyy") & "# , "
'                sql = sql & Format(Date + 30, "mm/dd/yyyy") & "# , "
                
                sql = sql & Numero(CStr(TotalPago), 2) & " ,"
                sql = sql & 0 & " ,"
                sql = sql & """" & VarDoc & """" & ", "
                sql = sql & """" & "SD" & """" & ")"
'                FrmFrente.Banco.Execute Sql
                
                
'                Sql = Sql & "#" & Format(DtEmissAnt, "mm/dd/yyyy") & "# , "
'                Sql = Sql & " 1 , "
'                Sql = Sql & DbcPesquisa.BoundText & " , #"
'                Sql = Sql & Format(Date, "mm/dd/yyyy") & "# , "
'                Sql = Sql & Numero(CStr(TotalPago), 2) & " ,"
'                Sql = Sql & 0 & " ,"
'                Sql = Sql & """" & "SD" & """" & ")"
                FrmFrente.Banco.Execute sql
                
                
              sql = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Juros, Maquina  ) VALUES ( "
             sql = sql & DbcPesquisa.BoundText & " , "
             sql = sql & "#" & Format(DTdata, "mm/dd/yyyy") & "# , "
             sql = sql & """" & Time & """" & " , "
    '         Sql = Sql & Numero(CStr(MebPago), 2) & " , "
             If CCur(MebPago) < CCur(MebTotal) Then
                sql = sql & Numero(CStr(MebPago), 2) & " , "
             Else
                sql = sql & Numero(CStr(MebTotal), 2) & " , "
             End If
             sql = sql & Numero(CStr(MebDesconto), 2) & " , "
             sql = sql & LstTipo.ListIndex + 1 & " ,"
             sql = sql & Dupli & " ,"
             sql = sql & Numero2(CStr(Juros)) & " ,"
             sql = sql & """" & LimpaCampo(Trim(FrmFrente.Maquina)) & """" & ")"
             FrmFrente.Banco.Execute sql
               
                
                
                If MsgBox("Imprimir comprovante de saldo ?", vbYesNo, App.Title) = vbYes Then
                   Iniciar (False)
                   Print #1, "======================================"
                   Print #1, "     COMPROVANTE DE SALDO             "
                   Print #1, "======================================"
                   Print #1, "Cliente: "; TxtCod; "            "; Format(DTdata, "dd/mm/yyyy")
                   Print #1, Cond & DbcPesquisa.Text
                   Print #1, Nor
                   Print #1, "Valor: "; Format(TotalPago, "#,##0.00;-#,##0.00")
                   Print #1, "Data.......: "; Format(DtEmissAnt, "dd/mm/yyyy")
                   Print #1, " "
                   Print #1, " "
                   Print #1, " "
                   Print #1, "Ass:------------------------------- "
                   Print #1, " "
                   Print #1, " "
                   Print #1, " "
                   Print #1, " "
                   Print #1, " "
                   Print #1, " "
                   Print #1, Nor;
                   Fim (False)
                End If
            Else
                 MebDesconto = CCur(MebDesconto) + CCur(TotalPago)
                  sql = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Juros, Maquina  ) VALUES ( "
                 sql = sql & DbcPesquisa.BoundText & " , "
                 sql = sql & "#" & Format(DTdata, "mm/dd/yyyy") & "# , "
                 sql = sql & """" & Time & """" & " , "
        '         Sql = Sql & Numero(CStr(MebPago), 2) & " , "
                 If CCur(MebPago) < CCur(MebTotal) Then
                    sql = sql & Numero(CStr(MebPago), 2) & " , "
                 Else
                    sql = sql & Numero(CStr(MebTotal), 2) & " , "
                 End If
                 sql = sql & Numero(CStr(MebDesconto), 2) & " , "
                 sql = sql & LstTipo.ListIndex + 1 & " ,"
                 sql = sql & Dupli & " ,"
                 sql = sql & Numero2(CStr(Juros)) & " ,"
                 sql = sql & """" & LimpaCampo(Trim(FrmFrente.Maquina)) & """" & ")"
                 FrmFrente.Banco.Execute sql
                
                
            End If
         Else
         
              sql = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Juros, Maquina  ) VALUES ( "
             sql = sql & DbcPesquisa.BoundText & " , "
             sql = sql & "#" & Format(DTdata, "mm/dd/yyyy") & "# , "
             sql = sql & """" & Time & """" & " , "
    '         Sql = Sql & Numero(CStr(MebPago), 2) & " , "
             If CCur(MebPago) < CCur(MebTotal) Then
                sql = sql & Numero(CStr(MebPago), 2) & " , "
             Else
                sql = sql & Numero(CStr(MebTotal), 2) & " , "
             End If
             sql = sql & Numero(CStr(MebDesconto), 2) & " , "
             sql = sql & LstTipo.ListIndex + 1 & " ,"
             sql = sql & Dupli & " ,"
             sql = sql & Numero2(CStr(Juros)) & " ,"
             sql = sql & """" & LimpaCampo(Trim(FrmFrente.Maquina)) & """" & ")"
             FrmFrente.Banco.Execute sql
         
            If TotalPago <> 0 Then
                MsgBox "Troco de R$ " & TotalPago, vbInformation, App.Title
            End If
         End If
         
         
         If MebDesconto.Enabled = True Then
            Dupli = VarPed
         End If
                
         If MsgBox("Imprimir comprovante de recebimento ?", vbYesNo, App.Title) = vbYes Then
            TotSaldo = MebTotGeral - CCur(MebPago)
         
            Iniciar (False)
'            Open "c:\lixo.txt" For Output As #1
            Print #1, "======================================"
            Print #1, "     COMPROVANTE DE RECEBIMENTO       "
            Print #1, "======================================"
            Print #1, "Cliente: "; TxtCod; "            "; Format(DTdata, "dd/mm/yyyy")
            Print #1, Cond & DbcPesquisa.Text
            Print #1, Nor
'            Print #1, "Valor Pago.: "; Format(MebPago, "#,##0.00;-#,##0.00")
            
            If MebJuros <> 0 Then
                Print #1, "Valor Pcl.: "; Format(CCur(MebTotal) - CCur(MebJuros), "#,##0.00;-#,##0.00")
                Print #1, " "
                Print #1, "Juros     : "; Format(MebJuros, "#,##0.00;-#,##0.00")
            Else
                Print #1, "Valor Pcl.: "; Format(MebTotal, "#,##0.00;-#,##0.00")
            End If
            Print #1, " "
                        
            If CCur(MebPago) < CCur(MebTotal) Then
                Print #1, "Valor Pago.: "; Format(MebPago, "#,##0.00;-#,##0.00")
            Else
                Print #1, "Valor Pago.: "; Format(MebTotal, "#,##0.00;-#,##0.00")
            End If
            If CCur(MebPago) > CCur(MebTotal) Then
                Print #1, "Troco.....: "; Format(TotalPago, "#,##0.00;-#,##0.00")
            End If
            
            Print #1, ""
            Print #1, "Tipo Pgto..: "; LstTipo.Text
            Print #1, ""
            If MebDesconto.Enabled = True Then
                Print #1, "Data Venc..: "; Format(DtVencAnt, "dd/mm/yyyy")
            Else
                Print #1, "Data.......: "; Format(DTdata, "dd/mm/yyyy")
            End If
            
            If VarSaldo = True Then
            
                    Print #1, "Total Conta.: "; Format(MebTotGeral, "#,##0.00;-#,##0.00")
    '                Print #1, ""
                    If CCur(MebPago) < CCur(MebTotal) Then
                        Print #1, "Valor Pago..: "; Format(MebPago, "#,##0.00;-#,##0.00")
    '                    Print #1, "  "
                        If MebDesconto = 0 Then
                            If TotSaldo > 0 Then
                                Print #1, "Saldo.......: "; Format(TotSaldo, "#,##0.00;-#,##0.00")
                            End If
                        Else
                            Print #1, "Desconto....: "; Format(MebDesconto, "#,##0.00;-#,##0.00")
                            If TotSaldo > 0 Then
                                Print #1, "Saldo.......: "; Format(TotSaldo, "#,##0.00;-#,##0.00")
                            End If
                        End If
                    Else
                        Print #1, "Valor Pago..: "; Format(MebTotal, "#,##0.00;-#,##0.00")
    '                    Print #1, "  "
                        If TotSaldo > 0 Then
                            Print #1, "Saldo.......: "; Format(TotSaldo, "#,##0.00;-#,##0.00")
                        End If
                    End If
            End If
            
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, " "
            Print #1, Nor;
'            Close #1
            Fim (False)
        End If
         
         MebTotal = 0
         MebPago = 0
         MebDesconto = 0
         MebJuros = 0
         CmdPgto.Enabled = False
         TxtDupl = ""
         DatDados(0).Refresh
         SomaGrid
         VarCont = 0
'         CmdLimpa_Click
    End If
End If

End Sub

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String

Frame.Visible = True
DoEvents

Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
Criterio = Criterio & Codigo

DatDados(1).Recordset.FindFirst Criterio
If DatDados(1).Recordset.NoMatch Then
    MsgBox "Entidade não Localizada", vbExclamation, App.Title
    DbcPesquisa.Text = ""
    DbcPesquisa.SetFocus
Else

    DatDados(0).RecordSource = "SELECT DUP.Entidade, Sum(DUP.VlrPago) AS SomaDeVlrPago, DUP.DtPgto FROM DUP WHERE (((DUP.Cond)=True)) GROUP BY DUP.Entidade, DUP.DtPgto HAVING DUP.Entidade =" & Val(Codigo) & " ORDER BY DUP.DtPgto"

    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount > 0 Then
        DatDados(0).Recordset.MoveLast
        LblPgto = "Ultimo pagamento dia " & DatDados(0).Recordset("DtPgto") & " no valor de R$ " & Format(DatDados(0).Recordset("SomaDeVlrPago"), "###,##0.00")
        
    End If
    DbcPesquisa.Text = DatDados(1).Recordset.Fields("Nome")
    If ChkPagos.Value = 0 Then
        If chkvencidos.Value = 1 Then
            DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & "  AND Cond = FALSE AND DtVenc <= #" & Format(Date + 1, "mm/dd/yyyy") & "# ORDER BY " & VarOrdem
        Else
            DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & "  AND Cond = FALSE  ORDER BY " & VarOrdem
        End If
    Else
        DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & " ORDER BY " & VarOrdem
    End If
    DatDados(0).Refresh
    If DatDados(0).Recordset.RecordCount <> 0 Then
        sql = "UPDATE DUP SET DUP.PgtTemp = False WHERE (((DUP.Entidade)= " & Val(Codigo) & ") AND ((DUP.Cond)=False))"
        FrmFrente.Banco.Execute sql
    
         DatDados(0).Recordset.MoveLast
    End If
    SomaGrid
    If DatDados(0).Recordset.RecordCount > 0 Then
        DatDados(0).Recordset.MoveFirst
    End If
End If

End Sub

Private Sub DatDados_Reposition(Index As Integer)
'If dbgrid1.Columns(1) <> "" Then
'    CmdPedido.Caption = "2ªVia.: " & psCompString(1, dbgrid1.Columns(1), 6)
'End If
End Sub

Private Sub DbcPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub DbcPesquisa_LostFocus()
If IsNumeric(DbcPesquisa.BoundText) Then
    TxtCod.Text = DbcPesquisa.BoundText
    Criterio = DatDados(1).Recordset.Fields("CodEntidade").Name & " = "
    Criterio = Criterio & DbcPesquisa.BoundText
    
    DatDados(1).Recordset.FindFirst Criterio
    
    If Not IsNull(DatDados(1).Recordset("Desconto")) Then
       If DatDados(1).Recordset("Desconto") <> "" Then
            If DatDados(1).Recordset("Desconto") <> 0 Then
                MebDesc = DatDados(1).Recordset("Desconto")
'                MebDesc.Enabled = False
            End If
        End If
    End If
End If

End Sub

Private Sub DBGrid1_DBLClick()
'TxtDupl.Text = ""
If Not Receber Then
'    If Not VarCaixa Then
        MsgBox "Você não tem permissão pra receber conta !", vbInformation
        Exit Sub
'    End If
End If
MebPago.Text = 0
If ChkPagos Then
    MsgBox "É preciso desmancar os pagos", vbInformation, App.Title
Else
    If MebTotal <> 0 Then
        MebDesconto.Enabled = False
        MebJuros.Enabled = False
    Else
        MebDesconto.Enabled = True
        MebJuros.Enabled = True
    End If
    FrmEnt.Enabled = False
    If DatDados(0).Recordset.RecordCount <> 0 Then
        If DBGrid1.Columns(16) = "Não" Then
            VarCont = VarCont + 1
            DatDados(0).Recordset.Edit
            DatDados(0).Recordset("PgtTemp") = True
            DatDados(0).Recordset.Update
            
            If DatDados(0).Recordset("Juros") <> "" Then
                MebJuros = MebJuros + CCur(DatDados(0).Recordset("Juros"))
            End If
            
            If DatDados(0).Recordset("Desconto") <> "" Then
                MebDesconto = MebDesconto + CCur(DatDados(0).Recordset("Desconto"))
            End If
            
            MebTotal.Text = CCur(MebTotal) + ((DatDados(0).Recordset.Fields("Valor") + CCur(DatDados(0).Recordset("Juros"))) - CCur(DatDados(0).Recordset("Desconto")))
            
    '        MebPago.Text =  (DatDados(0).Recordset.Fields("Valor") + CCur(DatDados(0).Recordset("Juros"))) - CCur(DatDados(0).Recordset("Desconto"))
            
        ElseIf DBGrid1.Columns(16) = "Sim" Then
            VarCont = VarCont - 1
            DatDados(0).Recordset.Edit
            DatDados(0).Recordset("PgtTemp") = False
            DatDados(0).Recordset.Update
    '        MebTotal.Text = CCur(MebTotal) - DatDados(0).Recordset.Fields("Valor")
            
            If DatDados(0).Recordset("Juros") <> "" Then
                MebJuros = MebJuros - CCur(DatDados(0).Recordset("Juros"))
            End If
            If DatDados(0).Recordset("Desconto") <> "" Then
                MebDesconto = MebDesconto - CCur(DatDados(0).Recordset("Desconto"))
            End If
            
            MebTotal.Text = CCur(MebTotal) - ((DatDados(0).Recordset.Fields("Valor") + CCur(DatDados(0).Recordset("Juros"))) - CCur(DatDados(0).Recordset("Desconto")))
    
    
    '        MebPago.Text = (DatDados(0).Recordset.Fields("Valor") + CCur(DatDados(0).Recordset("Juros"))) - CCur(DatDados(0).Recordset("Desconto"))
        End If
     End If
    MebPago = MebTotal
    If MebTotal <> 0 Then
        OptOrdem(0).Enabled = False
        OptOrdem(1).Enabled = False
        OptOrdem(2).Enabled = False
        OptOrdem(3).Enabled = False
    Else
        OptOrdem(0).Enabled = True
        OptOrdem(1).Enabled = True
        OptOrdem(2).Enabled = True
        OptOrdem(3).Enabled = True
    End If
'    CmdPgto.Enabled = True
'    MebDesconto.SetFocus
End If
Label6.Caption = "Itens Selecionados: " & VarCont

End Sub

Private Sub DtData_LostFocus()
'MebDesconto.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
'    If LstPedido.Visible = True Then
'        LstPedido.Visible = False
'    Else
        Unload Me
'    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()

'Me.Left = 100 '(MDIPrincipal.Width - Formulário.Width) / 2
'Me.Top = 100 '(MDIPrincipal.Height - Formulário.Height) / 4

DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " AND Cond = false AND Saldo = 0 ORDER BY " & VarOrdem
DatDados(0).Refresh

DatDados(1).RecordSource = "SELECT * FROM Ent  WHERE Tipo = 1 or Tipo >= 5 Order By Nome"
DatDados(1).Refresh

DTdata = Date
MebPago = 0
MebDesconto = 0
MebJuros = 0
MebTotal = 0
MebDesc = 0
Mebdias = 0
LstTipo.ListIndex = 0
CmdPgto.Enabled = False
ChkJuros.Value = 1


SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
Dim Componente As Variant

'Set Registros = FrmFrente.Banco.OpenRecordset("SELECT * FROM Dup")

For Each Componentes In DatDados
    Componentes.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
Next
If Receber Then
    CmdPgto.Enabled = True
Else
    CmdPgto.Enabled = False
End If

Flag = 1
VarOrdem = "DtVenc"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Flag = 0
End Sub

Private Sub MebDesc_LostFocus()
If MebDesc <> 0 Then
    Mebdias = 30
Else
    Mebdias = 0
End If
End Sub


Private Sub MebDesconto_lostfocus()
If MebTotal <> "" And MebDesconto <> "" And IsNumeric(MebDesconto) Then
    MebTotal = CCur(MebTotal) - CCur(MebDesconto)
End If
MebPago = Format(CCur(MebTotal), "#,##0.00;(#,##0.00)")
MebJuros.SetFocus
End Sub

Private Sub SomaGrid()
Dim CorrigePrd As Recordset
Dim Pago As Currency
Dim Apagar As Currency
Dim Saldo As Currency
Dim Prazo As Integer
Dim Criterio As String
Dim VarSomaPrd As Currency
Dim Contas As Recordset
Dim VarValor As Recordset
Pago = 0
Apagar = 0
Saldo = 0
TotJuros = 0
TotDesc = 0
VarSomaPrd = 0

'Set CorrigePrd = FrmFrente.Banco.OpenRecordset("SELECT ITP.*, PRD.*, ITP.Pedido, PED.Total FROM (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) INNER JOIN PED ON ITP.Pedido = PED.CodPed ORDER BY ITP.Pedido")
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveFirst
    Do Until DatDados(0).Recordset.EOF
'        Criterio = "Pedido = " & DatDados(0).Recordset.Fields("Pedido")
'        CorrigePrd.FindFirst Criterio
'        If CorrigePrd.RecordCount > 0 Then
'            While CorrigePrd!Pedido = DatDados(0).Recordset.Fields("Pedido") And Not CorrigePrd.EOF
'
'            Wend
'        End If
        If DatDados(0).Recordset.Fields("Cond") = True Then
            Pago = Pago + DatDados(0).Recordset.Fields("VlrPago")
        Else
            If VarReajConta = True Then
                sql = "SELECT ITP.Pedido, Sum(ITP.Valor) AS SomaDeValor, Sum(PED.Desconto) AS SomaDeDesconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN DUP ON ITP.Pedido = DUP.Pedido WHERE (((DUP.Cond)=False)) GROUP BY ITP.Pedido HAVING (((ITP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & ")) ORDER BY ITP.Pedido"
                
                Set VarValor = FrmFrente.Banco.OpenRecordset("SELECT ITP.Pedido, Sum(ITP.Valor) AS SomaDeValor, First(PED.Desconto) AS SomaDeDesconto FROM (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) INNER JOIN DUP ON ITP.Pedido = DUP.Pedido WHERE (((DUP.Cond)=False)) GROUP BY ITP.Pedido HAVING (((ITP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & ")) ORDER BY ITP.Pedido")

                VarValor.Requery
                If VarValor.RecordCount > 0 Then
                    If DatDados(0).Recordset.Fields("Valor") <> VarValor!SomaDeValor - VarValor!SomaDeDesconto Then
                        If DatDados(0).Recordset("Tipo_Dup") <> "SD" Then
                           If DatDados(0).Recordset("Tipo_Dup") <> "DV" Then
                                DatDados(0).Recordset.Edit
                                DatDados(0).Recordset.Fields("Valor") = VarValor!SomaDeValor - VarValor!SomaDeDesconto
                                DatDados(0).Recordset.Update
                           End If
                        End If
                    End If
                End If
                

                If VarPcoPrazo = True Then
                    Set Contas = FrmFrente.Banco.OpenRecordset("SELECT PED.Total, DUP.Pedido, DUP.Cond, Sum([itp].[qtd]*[prd].[PcoPrazo]) AS Expr1, Ped.Desconto FROM DUP INNER JOIN ((ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON DUP.Pedido = PED.CodPed GROUP BY PED.Total, DUP.Pedido, DUP.Cond, Ped.Desconto HAVING (((DUP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & " ) AND ((DUP.Cond)=False)) ORDER BY DUP.Pedido ")
                Else
                    Set Contas = FrmFrente.Banco.OpenRecordset("SELECT PED.Total, DUP.Pedido, DUP.Cond, Sum([itp].[qtd]*[prd].[venda]) AS Expr1, Ped.Desconto FROM DUP INNER JOIN ((ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON DUP.Pedido = PED.CodPed GROUP BY PED.Total, DUP.Pedido, DUP.Cond, Ped.Desconto HAVING (((DUP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & " ) AND ((DUP.Cond)=False)) ORDER BY DUP.Pedido ")
                End If
                Contas.Requery
                If Contas.RecordCount = 1 Then
                    If Contas!Expr1 > 0 Then
                        If Contas!Expr1 - Contas!Desconto > DatDados(0).Recordset.Fields("Valor") And DatDados(0).Recordset("Valor") > 0 Then
                            If DatDados(0).Recordset("Tipo_Dup") <> "SD" Then
                                If DatDados(0).Recordset("Tipo_Dup") <> "DV" Then
                                    DatDados(0).Recordset.Edit
                                    DatDados(0).Recordset.Fields("Valor") = Contas!Expr1 - Contas!Desconto
                                    DatDados(0).Recordset.Update
                                End If
                            End If
                        End If
                     End If
                End If
            End If
                
            If ChkJuros Then
                If DatDados(0).Recordset.Fields("DtVenc") < Date Then
                    Prazo = Date - DatDados(0).Recordset.Fields("DtVenc")
                    DatDados(0).Recordset.Edit
                    DatDados(0).Recordset.Fields("Juros") = DatDados(0).Recordset.Fields("Valor") * (((DatDados(1).Recordset("Juros") / 30) / 100) * Prazo)
                    DatDados(0).Recordset.Update
                    TotJuros = TotJuros + DatDados(0).Recordset("Juros")
                End If
            Else
                    DatDados(0).Recordset.Edit
                    DatDados(0).Recordset.Fields("Juros") = 0
                    DatDados(0).Recordset.Update
            End If
            
            If MebDesc <> 0 Then
                If DatDados(0).Recordset.Fields("DtVenc") > (Date - Mebdias) Then
                    DatDados(0).Recordset.Edit
                    DatDados(0).Recordset.Fields("Desconto") = DatDados(0).Recordset("Valor") * (MebDesc / 100)
                    DatDados(0).Recordset.Update
                    TotDesc = TotDesc + DatDados(0).Recordset("Desconto")
                End If
            Else
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("Desconto") = 0
                DatDados(0).Recordset.Update
            End If
                        
            Apagar = Apagar + DatDados(0).Recordset.Fields("Valor")
        End If
        If Not IsNull(DatDados(0).Recordset.Fields("Saldo")) Then
            Saldo = Saldo + DatDados(0).Recordset.Fields("Saldo")
        End If
        DatDados(0).Recordset.Edit
        DatDados(0).Recordset("PgtTemp") = 0
        DatDados(0).Recordset.Update
        DatDados(0).Recordset.MoveNext
    Loop
    MebPagos.Text = Pago + Saldo
    MebAPagar.Text = Apagar - Saldo
    MebSubTot = Pago + Apagar
    MebTotJuros = TotJuros
    MebTotDesc = TotDesc
    MebTotGeral = (MebSubTot + CCur(TotJuros)) - CCur(TotDesc)
End If
Frame.Visible = False

End Sub

Private Sub MebJuros_LostFocus()
If MebTotal <> "" And MebJuros <> "" And IsNumeric(MebJuros) Then
    MebTotal = CCur(MebTotal) + CCur(MebJuros)
End If
MebPago = CCur(MebTotal)

End Sub

Private Sub ImprimirPedido(Pedido As Double)
Dim SubTot As Currency
Dim Desconto As Currency
Dim Total As Currency
Dim Arq As Recordset
Dim CodEnt As Double
Dim x As Byte

Set Arq = FrmFrente.Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Uf, PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, Itp.Produto, PRD.Descrição_Produto FROM (ENT INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON ENT.CodEntidade = PED.Entidade) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE PED.CodPed = " & Pedido & " Order By PED.DataEmiss ")

If Arq.RecordCount > 0 Then

    If MsgBox("Confirma a Reimpressão do Pedido Nro " & Pedido, vbYesNo, App.Title) = vbYes Then
        Arq.MoveFirst
        Iniciar (False)
        CodEnt = Arq!CodEntidade
        Print #1, "" & Cond
        Print #1, ""
        Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
        Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 48)
        Print #1, "================================================"
        Print #1, "Data:" & Date & " Hora:" & Time; Tab(33); "Pedido.:"; psCompString(1, Pedido, 6)
        Print #1, Tab(1); psCompString(1, Arq!CodEntidade, 6); "-"; Left(Arq!Nome, 30)
        Print #1, Tab(1); Trim(Left(Arq!Endereco, 35)) & ", " & Arq!Nro
        Print #1, Tab(1); Trim(Left(Arq!Cidade, 30)) & " - " & Arq!Uf
        Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
        Print #1, Tab(1); "QTD"; Tab(20); "UNIT."; Tab(40); "TOTAL"
        Print #1, "------------------------------------------------"
        SubTot = 0
        Desconto = 0
        While Not Arq.EOF
            Print #1, Tab(1); psCompString(1, Arq!Produto, 6); "-"; psCompString(3, Left(Arq!Descrição_Produto, 39), 39)
            Print #1, Tab(3); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(12); "X"; Tab(17); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 9, 2); Tab(37); Alinhar(Format(Arq!Valor, "##,##0.00"), 10, 2)
            SubTot = SubTot + Arq!Valor
    '        Desconto = Desconto + Arq("Itp.Desconto")
    '        BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto")
            Arq.MoveNext
        Wend
        
        Print #1, ""
        Print #1, Tab(20); "SUB-TOTAL =>"; Tab(37); Alinhar(CStr(SubTot), 10, 2)
        Print #1, Tab(20); "DESCONTO  =>"; Tab(37); Alinhar(CStr(Desconto), 10, 2)
        Print #1, Tab(20); "TOTAL     =>"; Tab(37); Alinhar(CStr(SubTot - Desconto), 10, 2)
        If CodEnt = 1 Then
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
        Else
            Print #1, ""
            Print #1, ""
            Print #1, "         Ass.Cli-----------------------------"
            Print #1, "==============================================="
        End If
        Print #1, "" + Nor
        For x = 1 To NroLinha
            Print #1, ""
        Next

        Fim (False)
    Else
        Arq.MoveFirst
        CodEnt = Arq!CodEntidade
        LstPedido.Visible = True
        LstPedido.Clear
        With LstPedido
            .AddItem "================================================"
            .AddItem " Data:" & Date & "   Hora:" & Time & "  Pedido.:" & psCompString(1, Pedido, 6)
            .AddItem " " & psCompString(1, Arq!CodEntidade, 6) & "-" & Left(Arq!Nome, 30)
            .AddItem " " & Trim(Left(Arq!Endereco, 35)) & ", " & Arq!Nro
            .AddItem " " & Trim(Left(Arq!Cidade, 30)) & " - " & Arq!Uf
            .AddItem " CODIGO  DESCRICAO DO PRODUTO"
            .AddItem "   QTD                UNIT.               TOTAL"
            .AddItem "------------------------------------------------"
            SubTot = 0
            Desconto = 0
            While Not Arq.EOF
                .AddItem " " & psCompString(1, Arq!Produto, 6) & "-" & psCompString(3, Left(Arq!Descrição_Produto, 39), 39)
                .AddItem "   " & Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3) & "   " & "X" & "   " & Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 9, 2) & "         " & Alinhar(Format(Arq!Valor, "##,##0.00"), 10, 2)
                SubTot = SubTot + Arq!Valor
                Desconto = Desconto + Arq("Itp.Desconto")
'        '        BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto")
                Arq.MoveNext
            Wend
             .AddItem Space(30)
             .AddItem Space(30)
             .AddItem Space(20) & "SUB-TOTAL =>     " & Alinhar(CStr(SubTot), 10, 2)
             .AddItem Space(20) & "DESCONTO  =>     " & Alinhar(CStr(Desconto), 10, 2)
             .AddItem Space(20) & "TOTAL     =>     " & Alinhar(CStr(SubTot - Desconto), 10, 2)
             .AddItem Space(30)
             .AddItem Space(30)
             .AddItem Space(30)
        
        End With
    End If
Else
    MsgBox "Pedidos não se encontram nesse arquivo !", vbExclamation, App.Title
End If
End Sub

Private Sub ImprimiDiscr(CodEnt As Double)
Dim SubTot As Currency
Dim Desconto As Currency
Dim Total As Currency
Dim PedAnt As Double
Dim Pedido As Recordset
Dim Parcela As Recordset
Dim Juros As Currency
Dim TotJuros As Currency
Dim DtAnt As Date
Dim DescPed As Currency
Dim FlagDesc As Boolean
Dim Criterio As String
Dim VarObs As String

If chkvencidos.Value = 1 Then
    Set VarDup = FrmFrente.Banco.OpenRecordset("SELECT * FROM Dup  WHERE Entidade = " & Val(TxtCod) & " and Cond = False AND DtVenc <= #" & Format(Date + 1, "mm/dd/yyyy") & "# ORDER BY DtVenc, DtEmiss")
Else
    Set VarDup = FrmFrente.Banco.OpenRecordset("SELECT * FROM Dup  WHERE Entidade = " & Val(TxtCod) & " and Cond = False ORDER BY DtVenc, DtEmiss")
End If
VarDup.Requery


'Set Arq = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 ORDER BY  Ped.CodPed ")

'Sql = "SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 ORDER BY  Ped.CodPed "

'Set Pedido = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE PED.Entidade = " & CodEnt & " Order By CodPed ")

'Sql = "SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc, Ped.Operacao  FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 AND (Ped.Operacao = 1 or Ped.Operacao = 2 ) ORDER BY Ped.CodPed  "
Set Arq = FrmFrente.Banco.OpenRecordset("SELECT Ped.Obs, ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto,  Prd.PcoPrazo, Prd.Venda, ITP.Qtd, ITP.VlrVdaDia, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc, Ped.Operacao  FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 AND (Ped.Operacao = 1 or Ped.Operacao = 2 ) ORDER BY Ped.CodPed  ")


If Arq.RecordCount > 0 Then
        Arq.MoveFirst
'        Iniciar (False)
        Open App.Path & "\RelDiscr.txt" For Output As #1
        Print #1, ""
        Print #1, "" & Cond
        Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
        Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 45)
        Print #1, "=========================================="
        Print #1, "Data: " & Date & " Hora: " & Time; Tab(33)
        Print #1, Tab(1); psCompString(1, Arq!CodEntidade, 6); "-"; Left(Arq!Nome, 30)
        Print #1, Tab(1); Trim(Left(Arq!Endereco, 30)); ", "; Arq!Nro
        Print #1, "------------------------------------------"
        Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
        Print #1, Tab(1); "QTD"; Tab(20); "UNIT."; Tab(38); "TOTAL"
        Print #1, "------------------------------------------"
        SubTot = 0
        Desconto = 0
        TotJuros = 0
        Total = 0
        Juros = 0
        DescPed = 0
        VarDup.MoveFirst
        While Not VarDup.EOF
               Criterio = "CodPed = " & VarDup("Pedido")
               Arq.FindFirst Criterio
               If Not Arq.NoMatch And VarDup!Tipo_Dup <> "SD" Then
                   PedAnt = Arq!CodPed
                   Print #1, Tab(1); "Pedido ==>"; Tab(13); psCompString(1, Arq!CodPed, 6); Tab(22); Arq!DataEmiss
                   If Arq("Ped.Desconto") <> 0 Then
                        DescPed = Arq("Ped.Desconto")
                        Desconto = Desconto + DescPed
                   Else
                        DescPed = 0
                   End If
                   If Arq!Obs <> "" Then
                        VarObs = Arq!Obs
                   Else
                        VarObs = ""
                   End If

                   While Not Arq.EOF
                        If PedAnt = Arq!CodPed Then
                           If VarReajConta Then
                               If VarPcoPrazo = True Then
                                   If Arq!Qtd * Arq!PcoPrazo > Arq!Valor Then
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!PcoPrazo, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Qtd * Arq!PcoPrazo, "##,##0.00"), 8, 2)
                                        SubTot = SubTot + (Arq!Qtd * Arq!PcoPrazo)
                                        Total = Total + (Arq!Qtd * Arq!PcoPrazo)
                                   Else
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 8, 2)
                                        SubTot = SubTot + Arq!Valor
                                        Total = Total + Arq!Valor
                                   End If
                               Else
                                   If Arq!Qtd * Arq!Venda > Arq!Valor Then
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!Venda, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Venda * Arq!Qtd, "##,##0.00"), 8, 2)
                                        SubTot = SubTot + (Arq!Qtd * Arq!Venda)
                                        Total = Total + (Arq!Qtd * Arq!Venda)
                                   Else
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 8, 2)
                                        SubTot = SubTot + Arq!Valor
                                        Total = Total + Arq!Valor
                                   End If
                               End If
                           Else
                                If Arq!vlrvdadia = 0 Then
                                    Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 8, 2)
                                    SubTot = SubTot + Arq!Valor
                                    Total = Total + Arq!Valor
                                Else
                                    Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 6, 3); Tab(9); psCompString(3, Left(Arq!Descrição_Produto, 17), 19); Alinhar(Format(Arq!vlrvdadia / Arq!Qtd, "##,##0.00"), 7, 2); Alinhar(Format(Arq!vlrvdadia, "##,##0.00"), 8, 2)
                                    SubTot = SubTot + Arq!vlrvdadia
                                    Total = Total + Arq!vlrvdadia
                                End If
                           End If
                           DtAnt = Arq!DtVenc
                           Juros = Arq!Juros
                       End If
                       Arq.MoveNext
                   Wend
                   If Juros <> 0 Then
                        TotJuros = TotJuros + Juros
                        Print #1, Tab(10); "   Atraso .... "; tab30; Date - DtAnt & " dias  "; Tab(39); Alinhar(CStr(Juros), 10, 2)
                   End If
                   If DescPed <> 0 Then
                        Print #1, Tab(31); "Desconto"; Alinhar(CStr(DescPed), 10, 2)
                   End If
                   If VarObs <> "" Then
                        Print #1, Tab(1); VarObs
                   End If
                   Print #1, Tab(1); "================================" & Alinhar(CStr(SubTot + Juros), 10, 2)
                   SubTot = 0
                   Juros = 0
               Else
                   SubTot = SubTot + VarDup!Valor
                   Total = Total + VarDup!Valor
                   Juros = VarDup!Juros
                   Print #1, Tab(1); psCompString(1, VarDup("Pedido"), 6); Tab(9); psCompString(1, VarDup("NroPcl"), 3); Tab(14); VarDup("DtEmiss"); Tab(26); VarDup("DtVenc"); Tab(37); Alinhar(Format(VarDup("Valor"), "##,##0.00"), 9, 2); Tab(47); VarDup("Tipo_Dup")
                   If Juros <> 0 Then
                        TotJuros = TotJuros + Juros
                        Print #1, Tab(13); "   Atraso .... " & Date - VarDup!DtVenc & " dias  "; Tab(39); Alinhar(CStr(Juros), 10, 2)
                    End If
                    Print #1, Tab(1); "==============================" & Alinhar(CStr(SubTot + Juros), 10, 2)
                    SubTot = 0
                    Juros = 0
               
               End If
               VarDup.MoveNext
            Wend
            Print #1, ""
            Print #1, Tab(17); "SUB-TOTAL =>"; Tab(33); Alinhar(CStr(Total), 9, 2)
            Print #1, Tab(17); "ACRESCIMO =>"; Tab(33); Alinhar(CStr(TotJuros), 9, 2)
            Print #1, Tab(17); "DESCONTO  =>"; Tab(33); Alinhar(CStr(Desconto), 9, 2)
            Print #1, Tab(17); "TOTAL     =>"; Tab(33); Alinhar(CStr((Total + TotJuros) - Desconto), 9, 2)
'        End If
        Print #1, "" + Nor
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
            
'        Fim (False)
        Close #1
        If MsgBox("Enviar para impressora ?", vbYesNo, App.Title) = vbYes Then
            Set Fs2 = CreateObject("Scripting.FileSystemObject")
            
            Fs2.CopyFile App.Path & "\RelDiscr.txt", PortaNF
        
        Else
            SSTab1.Tab = 2
            On Error GoTo erro
            LstPedido.Clear
            Open App.Path & "\RelDiscr.txt" For Input As #1
            Do
                Line Input #1, Temp
                LstPedido.AddItem psCompString(3, Temp, 28)
        '        Lista.Selected(Lista.ListCount - 1) = True
            Loop Until Temp = "<>"
            Close #1
        End If

Else
    MsgBox "Pedidos não se encontram nesse arquivo !", vbExclamation, App.Title
End If

Exit Sub
erro:
If Err.Number = 62 Then
    Close #1
End If

End Sub

Private Sub MebTotal_Change()
If MebTotal = 0 Then
   CmdPgto.Enabled = False
   MebPago.Enabled = False
Else
    If VarCont > 0 Then
        CmdPgto.Enabled = True
        MebPago.Enabled = True
    Else
        CmdPgto.Enabled = False
        MebPago.Enabled = False
        MebTotal = 0
        MebPago = 0
    End If
End If
End Sub

Private Sub OptOrdem_Click(Index As Integer)
Select Case Index
    Case 0
        VarOrdem = "CodDup"
    Case 1
        VarOrdem = "DtEmiss"
    Case 2
        VarOrdem = "DtVenc"
    Case 3
        VarOrdem = "DtPgto"
End Select

If TxtCod <> "" Then
    Pesquisa TxtCod
End If

End Sub

Private Sub SSTab1_Click(previoustab As Integer)
Dim SqlPed As Recordset
Dim SqlTot As Recordset
Dim LstPesq As ListItem

Dim sql As String
Dim VarCodPed As Double
If SSTab1.Tab = 1 Then
    Lista.ListItems.Clear
    If IsNumeric(DBGrid1.Columns(1)) Then
        VarCodPed = DBGrid1.Columns(1)
    Else
'        DBGrid1.SetFocus
        VarCodPed = 0
        Exit Sub
    End If
    sql = "Select * from Ped WHERE Codped =  " & VarCodPed
    Set SqlTot = FrmFrente.Banco.OpenRecordset(sql)
    SqlTot.Requery
    If SqlTot.RecordCount > 0 Then
        MebTotalPed = SqlTot!Total
        MebDescPed = SqlTot!Desconto
    End If
        
    sql = "SELECT ITP.Produto, PRD.Descrição_Produto, PRD.Embalagem, ITP.Qtd, ITP.ValorUnit, ITP.VlrVdaDia, Itp.VendaDia, ITP.Pedido, ITP.CodItp FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido =  " & VarCodPed & " ORDER BY ITP.CodItp"
    Set SqlPed = FrmFrente.Banco.OpenRecordset(sql)
    SqlPed.Requery
    If SqlPed.RecordCount > 0 Then
        SqlPed.MoveFirst
        While Not SqlPed.EOF
            Set LstPesq = Lista.ListItems.Add(, , SqlPed("Produto"))
    
            LstPesq.SubItems(1) = psCompString(3, Left(SqlPed!Descrição_Produto, 40), 40)
            LstPesq.SubItems(2) = psCompString(3, Left(SqlPed!Embalagem, 10), 10)
            
            If Right(Format(Alinhar(SqlPed!Qtd, 10, 2), "##,###.00"), 2) = "00" Then
                LstPesq.SubItems(3) = Format(Alinhar(SqlPed!Qtd, 10, 2), "##,###,##0")
            Else
                LstPesq.SubItems(3) = Format(Alinhar(SqlPed!Qtd, 10, 2), "###,##0.0#")
            End If
            LstPesq.SubItems(4) = Alinhar(SqlPed!ValorUnit, 10, 3)
            LstPesq.SubItems(5) = Alinhar(SqlPed!vlrvdadia, 12, 2)
            LstPesq.SubItems(6) = Alinhar(SqlPed!VendaDia, 12, 2)
            SqlPed.MoveNext
        Wend
    End If
End If
End Sub

