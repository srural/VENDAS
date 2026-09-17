VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmTotal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DatGru 
      Caption         =   "Grupo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data DatPcl 
      Caption         =   "Parcelas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.Valor, DUP.Cond FROM DUP WHERE (((DUP.Pedido)=500)) ORDER BY DUP.NroPcl"
      Top             =   7245
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   2400
      ScaleHeight     =   3810
      ScaleWidth      =   7995
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   8025
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   8025
         TabIndex        =   33
         Top             =   0
         Width           =   8025
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Calculadora"
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
            Left            =   90
            TabIndex        =   34
            Top             =   45
            Width           =   3795
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Ok"
         Height          =   480
         Left            =   3285
         TabIndex        =   39
         Top             =   2835
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MebValor1 
         DataField       =   "Custo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   435
         Left            =   225
         TabIndex        =   35
         Tag             =   "S"
         Top             =   1080
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebValor2 
         DataField       =   "Custo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   435
         Left            =   225
         TabIndex        =   36
         Tag             =   "S"
         Top             =   1950
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   435
         Left            =   225
         TabIndex        =   37
         Tag             =   "S"
         Top             =   2850
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;- #,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3465
         TabIndex        =   43
         Top             =   600
         Width           =   4110
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor 2"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   1710
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor 1"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   2610
         Width           =   780
      End
      Begin VB.Label LblTroco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   3465
         TabIndex        =   38
         Top             =   1800
         Width           =   4110
      End
   End
   Begin VB.TextBox TotPgto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2970
      TabIndex        =   47
      Top             =   4545
      Width           =   3255
   End
   Begin VB.ListBox ListPgto 
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
      Height          =   1155
      ItemData        =   "FrmTotal.frx":0000
      Left            =   2970
      List            =   "FrmTotal.frx":0002
      TabIndex        =   46
      Top             =   3420
      Width           =   3255
   End
   Begin VB.CheckBox ChkComissao 
      Caption         =   "Check1"
      Height          =   420
      Left            =   2520
      TabIndex        =   45
      Top             =   2610
      Width           =   240
   End
   Begin VB.CommandButton Pgto 
      Caption         =   ">>"
      Height          =   285
      Left            =   2385
      TabIndex        =   44
      Top             =   4140
      Width           =   420
   End
   Begin VB.ComboBox txtpgto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "FrmTotal.frx":0004
      Left            =   225
      List            =   "FrmTotal.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4455
      Width           =   2625
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   29
      Top             =   8190
      Width           =   11355
   End
   Begin VB.CommandButton BtCalc 
      Caption         =   "Calculadora"
      Height          =   375
      Left            =   5220
      TabIndex        =   28
      Top             =   6120
      Width           =   1050
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   7035
      Left            =   6390
      ScaleHeight     =   6975
      ScaleWidth      =   4455
      TabIndex        =   26
      Top             =   630
      Width           =   4515
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   120
         Picture         =   "FrmTotal.frx":00B1
         ScaleHeight     =   3525
         ScaleWidth      =   4440
         TabIndex        =   27
         Top             =   3360
         Width           =   4440
      End
      Begin VB.Image Image3 
         Height          =   3750
         Left            =   360
         Picture         =   "FrmTotal.frx":2610
         Top             =   900
         Width           =   3750
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   0
         Picture         =   "FrmTotal.frx":53DC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4455
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
      ScaleWidth      =   11220
      TabIndex        =   24
      Top             =   0
      Width           =   11220
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   5280
         TabIndex        =   50
         Top             =   120
         Width           =   5685
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizando Pedido"
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
         TabIndex        =   25
         Top             =   45
         Width           =   4245
      End
   End
   Begin VB.TextBox TxtObs2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   225
      MaxLength       =   35
      TabIndex        =   7
      Top             =   6405
      Width           =   4860
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   225
      MaxLength       =   35
      TabIndex        =   6
      Top             =   6105
      Width           =   4860
   End
   Begin MSMask.MaskEdBox MebValorDig 
      Height          =   435
      Left            =   225
      TabIndex        =   3
      Top             =   3465
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "   #,##0.00; -  #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebValor 
      Height          =   540
      Left            =   225
      TabIndex        =   11
      Top             =   900
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   953
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   0
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00; -#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTroco 
      Height          =   435
      Left            =   2925
      TabIndex        =   12
      Top             =   2640
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "   #,##0.00; -  #,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTotal 
      Height          =   540
      Left            =   2925
      TabIndex        =   19
      Top             =   900
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   953
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14737632
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00; -#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDesc 
      Height          =   435
      Left            =   225
      TabIndex        =   0
      Top             =   1800
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   " #,##0.00; -#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDesconto 
      Height          =   435
      Left            =   2925
      TabIndex        =   1
      Top             =   1785
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "   #,##0.000; -  #,##0.000"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDoc 
      Height          =   405
      Left            =   225
      TabIndex        =   5
      Top             =   5355
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebComprove 
      DataField       =   "Custo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   435
      Left            =   3645
      TabIndex        =   8
      Tag             =   "S"
      Top             =   2340
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Enabled         =   0   'False
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
   Begin VB.CommandButton CmdFicha 
      Caption         =   "Ficha"
      Height          =   375
      Left            =   5220
      TabIndex        =   31
      Top             =   6705
      Width           =   1050
   End
   Begin MSMask.MaskEdBox MebComissao 
      DataField       =   "Custo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   435
      Left            =   225
      TabIndex        =   2
      Tag             =   "S"
      Top             =   2610
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDBGrid.DBGrid DBPcl 
      Bindings        =   "FrmTotal.frx":799D
      Height          =   1260
      Left            =   225
      OleObjectBlob   =   "FrmTotal.frx":79B2
      TabIndex        =   48
      Top             =   6885
      Width           =   4875
   End
   Begin MSMask.MaskEdBox MebNfe 
      Height          =   345
      Left            =   9855
      TabIndex        =   49
      Top             =   7785
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Comissão"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   30
      Top             =   2340
      Width           =   1530
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.F./C.N.P.J.  (somente numeros)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   225
      TabIndex        =   23
      Top             =   5100
      Width           =   3315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto (R$)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2925
      TabIndex        =   22
      Top             =   1530
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   21
      Top             =   1530
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2925
      TabIndex        =   20
      Top             =   630
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   18
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   8835
      TabIndex        =   17
      Top             =   11205
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   16
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   15
      Top             =   3225
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   14
      Top             =   630
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Troco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2925
      TabIndex        =   13
      Top             =   2370
      Width           =   630
   End
   Begin VB.Label LblCupom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cupon"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   7650
      TabIndex        =   10
      Top             =   11205
      Width           =   465
   End
   Begin VB.Label LblMaquina 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maq"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   10035
      TabIndex        =   9
      Top             =   11205
      Width           =   315
   End
End
Attribute VB_Name = "FrmTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Caixa As Recordset

Public VarDoc As New MSXML2.DOMDocument

Public Det As MSXML2.IXMLDOMNodeList

Public CfeProc As MSXML2.IXMLDOMElement
Public Cfe As MSXML2.IXMLDOMElement
Public VarinfCfe As MSXML2.IXMLDOMElement
Public infProt As MSXML2.IXMLDOMElement
Public Ide As MSXML2.IXMLDOMElement

Dim Flag As Boolean
Public VarLoja  As String
Public VarCxa As String
Public VarCartao As String
Public VarRetAmpla As String

Public FormaPgto As String
Public VarImposto As Currency
Public VarImpostoNac As Currency
Public VarImpostoEst As Currency
Public VarTotImpostoNac As Currency
Public VarTotImpostoEst  As Currency
Public Comissao As Currency
Public XPgto As Integer
Public VarDupTemp As Recordset
Public VarVdaPrazo As Currency
Public VarDescPgto As Boolean
Public VarDescTot As Currency



'Public TxtImposto As String


Private Sub BtAmpla_Click()
If TxtCartao <> "" Then
    VarCartao = TxtCartao
    FrameAmpla.Enabled = False
Else
    If MsgBox("Numero invalido ! Deseja redigita ? ", vbYesNo, App.Title) = vbYes Then
        TxtCartao.SetFocus
    Else
        FrameAmpla.Enabled = False
    End If
End If
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub BtCalc_Click()
Picture4.Visible = True

MebValor1 = MebTotal
MebValor2 = ""
MebValor2.SetFocus

End Sub

Private Sub ChkComissao_Click()
If ChkComissao.Value = 0 Then
    MebTotal = MebValor
    MebValorDig = MebTotal
End If
End Sub

Private Sub CmdFicha_Click()
Dim x As Integer
Iniciar (False)
    
    FrmFrente.DatCompra.Recordset.MoveFirst

    While Not FrmFrente.DatCompra.Recordset.EOF
        For x = 1 To FrmFrente.DatCompra.Recordset("Qtd")
            Print #1, "========================================"
            Print #1, Centralizar(Left(FrmFrente.Empresa("Cabecalho1"), 40), 40)
                
            Print #1, "----------------------------------------"
            Print #1, ""
            Print #1, Tab(1); "1.00"; "  "; Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 33)
            
            Print #1, ""
            Print #1, "========================================"
            For y = 1 To NroLinha
                Print #1, ""
            Next
            Print #1, Chr(27) & "i" '& Chr(27) & "v50"
        Next
        FrmFrente.DatCompra.Recordset.MoveNext
    Wend
Fim (False)
txtpgto.SetFocus
End Sub

Private Sub Command10_Click()
LblTotal = ""
MebValor1 = ""
MebValor2 = ""
Picture4.Visible = False
txtpgto.SetFocus
End Sub

Private Sub Form_Activate()
MebDesc = 0
MebDesconto = 0
'MebDesc.SetFocus
'Screen.ActiveControl.SelStart = 0
'Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
FrmFrente.DatCompra.RecordSource = "SELECT Prd.Grupo, Prd.Icm, Prd.SitTrib, Prd.CfopPrd, Classe, itp.loja, Itp.Desconto, Itp.VendaDia, Itp.CodItp, Itp.Produto, Prd.CodBar, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ITP.Entregue FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & FrmFrente.LblPedido & " ORDER BY ITP.CodItp"
sql = "SELECT Prd.Icm, Prd.SitTrib, Prd.CfopPrd, Classe, itp.loja, Itp.Desconto, Itp.VendaDia, Itp.CodItp, Itp.Produto, Prd.CodBar, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ITP.Entregue FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & FrmFrente.LblPedido & " ORDER BY ITP.CodItp"
FrmFrente.DatCompra.Refresh
Flag = True
FrmFrente.Timer2.Enabled = False

If VarObsTotal = False Then
    TxtObs.Enabled = False
    TxtObs2.Enabled = False
End If

MebValor = SomaPedido(LblTot)
MebTotal = MebValor


VarDesc5 = 0
VarDesc15 = 0
VarDescPromo = 0

If VarColombo Then
    FrmFrente.DatCompra.Recordset.MoveFirst
    While Not FrmFrente.DatCompra.Recordset.EOF
        If VarAniversariante = True Then
            VarDesc15 = VarDesc15 + (FrmFrente.DatCompra.Recordset("Valor") * 0.1)
            VarDesc15 = Round(VarDesc15, 2)
        End If
        Select Case FrmFrente.DatCompra.Recordset("Grupo")
            Case 110
                VarDescPromo = VarDescPromo + (FrmFrente.DatCompra.Recordset("Valor") * 0.1)
            Case 120
                VarDescPromo = VarDescPromo + (FrmFrente.DatCompra.Recordset("Valor") * 0.2)
            Case 130
                VarDescPromo = VarDescPromo + (FrmFrente.DatCompra.Recordset("Valor") * 0.3)
            Case 150
                VarDescPromo = VarDescPromo + (FrmFrente.DatCompra.Recordset("Valor") * 0.5)
            Case Else
                If FrmFrente.DatCompra.Recordset("Grupo") <> 27 Then
                    VarDesc5 = VarDesc5 + (FrmFrente.DatCompra.Recordset("Valor") * 0.05)
                    VarDesc5 = Round(VarDesc5, 5)
                End If
        End Select
        FrmFrente.DatCompra.Recordset.MoveNext
    Wend
End If
VarDescPromo = Round(VarDescPromo, 2)

If VarComissao <> 0 Then
    Comissao = (CCur(MebTotal.Text) * (VarComissao / 100))
    ChkComissao.Value = 1
    MebComissao = Comissao
    MebValorDig = MebTotal + CCur(MebComissao)
Else
    MebComissao.Enabled = False
End If

If VarDescFim = False Then
'  MebDesc.Enabled = False
 ' MebDesconto.Enabled = False
  MebValorDig.SetFocus
End If

If Impressora = 8 Then
    MebNfe = LerINI("Nfce", "NroNfce", App.Path & "\config.ini")
End If


End Sub

Private Sub Form_Load()
'MebValor = FrmFrente.MebTotal

If Impressora = 99 Then
'    Label11.Visible = False
'    Label19.Visible = False
'    MebDoc.Visible = False
    Picture1.Visible = False
    Image2.Visible = False
    Image3.Visible = True
'    Frame1.Visible = False
Else
    Image3.Visible = False
    Image2.Visible = True
    Label19.Visible = True
    Picture1.Visible = True
    
    If FrmFrente.Clientes("Classificacao") = 2 Then
           MebDoc = Format(FrmFrente.Clientes("Cgc"), "@@@.@@@.@@@/@@@@-@@")
       Else
           MebDoc = Format(FrmFrente.Clientes("Cpf"), "@@@.@@@.@@@-@@")
    End If
End If

If VarDescFim = False Then
  MebDesc.Enabled = False
  MebDesconto.Enabled = False
'  MebValorDig.SetFocus
End If

XPgto = 0
TotPgto.Text = Format(0, "###,##0.00")
DatPcl.DatabaseName = FrmFrente.Caminho & "\Dados.Mdb"
DatGru.DatabaseName = FrmFrente.Caminho & "\Dados.mdb"
If VarColombo Then
    VarDescPgto = True
End If


If Impressora = 8 Then
    AbreAcbr "NFE.CertificadoDataVencimento(" & App.Path & "\cfe.ini" & ")"
    TrataAcbr2 "Certificado", 1
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 9 Then
'    Screen.ActiveControl.SelStart = 0
'    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
End If
End Sub


Private Sub Timer1_Timer()
If Image1.Visible = True Then
    Image1.Visible = False
    Image2.Visible = True
Else
    Image2.Visible = False
    Image1.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmFrente.Timer2.Enabled = True
End Sub


Private Sub ListPgto_DblClick()
TotPgto = TotPgto - CCur(Right(ListPgto.Text, 9))
ListPgto.RemoveItem (ListPgto.ListIndex)
XPgto = XPgto - 1
MebValorDig = MebTotal - CCur(TotPgto)
txtpgto.Enabled = True


End Sub

Private Sub MebComissao_KeyPress(KeyAscii As Integer)
If IsNumeric(MebComissao) Then
    MebTotal = MebValor + CCur(MebComissao)
    MebValorDig = MebTotal
End If
End Sub

Private Sub MebDesc_GotFocus()
MebDesc.SelStart = 0
MebDesc.SelLength = 3
End Sub

Private Sub MebDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
If KeyAscii = 13 Then
    If IsNumeric(MebDesc) Then
        If MebDesc <> "" And MebValor <> "" Then
            MebDesconto = Format((MebValor - VarPrdDesc) * (MebDesc / 100), "##,##0.00")
            MebTotal = Format(MebValor - MebDesconto, "##,##0.00")
            If VarPrdDesc > 0 Then
'                MebDesconto.Enabled = False
            End If
        Else
            If MebValor <> "" Then
                MebDesconto = 0
                MebTotal = Format(MebValor - MebDesconto, "##,##0.00")
            End If
        End If
    Else
        MebDesc = ""
        MebDesc.SetFocus
    End If
End If
End Sub

Private Sub MebDesconto_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
If KeyAscii = 13 Then
    If IsNumeric(MebDesconto) Then
        If MebDesconto <> "" And MebValor <> "" Then
            MebTotal = CCur(MebValor) - CCur(MebDesconto)
        Else
            MebTotal = MebValor
        End If
    Else
        MebDesconto = ""
        MebDesconto.SetFocus
    End If
End If
End Sub


Private Sub MebDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MebDoc <> "" Then
        If IsNumeric(LIMPACampo2(MebDoc)) Then
            If Len(Trim(LIMPACampo2(MebDoc))) = 11 Then
                MebDoc = Format(Trim(LIMPACampo2(MebDoc)), "@@@.@@@.@@@-@@")
                If Not CalculaDacCPF(MebDoc) Then
                    MsgBox "CPF Inválido", vbInformation
                    MebDoc.SetFocus
                Else
'                    txtpgto.SetFocus
                End If
            ElseIf Len(Trim(LIMPACampo2(MebDoc))) = 14 Then
                MebDoc = Format(Trim(LIMPACampo2(MebDoc)), "@@.@@@.@@@/@@@@-@@")
                If Not VerificaCGC(MebDoc) Then
                    MsgBox "CNPJ Inválido.", vbInformation
                    MebDoc.SetFocus
                Else
'                    txtpgto.SetFocus
                End If
            Else
                MsgBox "Documento invalido", vbInformation
                MebDoc.SetFocus
            End If
        End If
    Else
        If VarObsTotal = False Then
            TxtObs2_KeyPress 13
        End If
'        txtpgto.SetFocus

    End If
End If

End Sub

Private Sub MebPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If IsNumeric(MebPago) Then
        LblTroco = Format(MebPago - CDec(LblTotal), "###,##0.00")
    End If
    Command10.SetFocus
End If

End Sub

Private Sub MebPago1_Change()
If MebPago1 <> "" Then
End If

End Sub

Private Sub MebValor2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If IsNumeric(MebValor1) And IsNumeric(MebValor2) Then
        LblTotal = Format(MebValor1 + CDec(MebValor2), "###,##0.00")
    End If
    MebPago.SetFocus
End If
End Sub


Private Sub MebValorDig_GotFocus()
If ChkComissao.Value = 0 Then
    MebValorDig = ""
End If
End Sub

Private Sub MebValorDig_LostFocus()
If IsNull(MebValorDig) Or MebValorDig = "" Then
    If ChkComissao.Value = 0 Then
        MebValorDig.Text = MebTotal - TotPgto
        MebTroco.Text = 0
    Else
        If MebValorDig <> "" Then
            MebTroco.Text = Format(CDec((MebValorDig)) - CCur(MebTotal), "  ##,##0.00")
        End If
    End If

End If
End Sub

Private Sub Pgto_Click()
If MebValorDig = "" Then
    MsgBox "Nenhum Valor Digitado", vbInformation
    Exit Sub
End If
    
If txtpgto.Text = "" Then
    MsgBox "Selecionar Tipo Pgto", vbInformation
    Exit Sub
End If

If TotPgto >= CCur(MebTotal) Then
    MsgBox "Total de Pgto Completo", vbInformation
    Exit Sub
End If
        
    ListPgto.List(XPgto) = psCompString(3, txtpgto.Text, 15) & "=>" & Alinhar(Format(MebValorDig, "###,##0.00"), 9, 2)
    XPgto = XPgto + 1
    TotPgto = TotPgto + CCur(MebValorDig)
    TotPgto = Format(TotPgto, "###,##0.00")
    MebValorDig = MebTotal - CCur(TotPgto)
    If MebValorDig > 0 Then
        MebValorDig.SetFocus
        MebValorDig.SetFocus
    Else
        MebDoc.SetFocus
    End If
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
If KeyAscii = 13 Then
'    TxtPgto.SetFocus
End If
End Sub

Function LimpaPrd() As String
Dim Produto As String
Dim PrdSujo As String
Dim x As Integer
Produto = ""
PrdSujo = Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 38)
For x = 1 To Len(PrdSujo)
    If Mid$(PrdSujo, x, 1) >= Chr(32) And Mid$(PrdSujo, x, 1) <= Chr(122) Then
            Produto = Produto + Mid$(PrdSujo, x, 1)
    End If
Next
LimpaPrd = psCompString(3, Produto, 38)
End Function


Private Sub TxtObs2_KeyPress(KeyAscii As Integer)
Dim Retorno As Integer
Dim retorno2 As String
Dim sql As String
Dim Cupom As String
Dim Troco As Currency
Dim erro As Integer
Dim pag As String
Dim buf_ret As String * 100


Dim Linha1 As String
Dim Linha2 As String
Dim Linha3 As String
Dim Linha4 As String
Dim Linha5 As String
Dim Linha6 As String
Dim Linha7 As String
Dim MensagemPromo As String
Dim Resp  As String
Dim Status As String
Dim TotalCupom As String
Dim Criterio As String
Dim Pago As Integer
Dim ConvertQtd As String
Dim ConvertVr As String
Dim CodigoBarra As String
Dim Dias As Integer
Dim x As Byte
Dim VarDesc As Currency
Dim VarValordesc As Currency
Dim VarValorNeg As Currency
Dim VarPopular As Recordset
Dim VarAutoriza As String
Dim VarEpharma As Recordset
Dim VarChavePop As String

Dim varSubTotal As Currency
Dim VarTotal As Currency
Dim VarVenc As String
Dim Doc1 As String
Dim Doc2 As String
Dim VarEst As String
Dim VarNroNota As String
Dim NroCupom2 As String
Dim VarCnpj As String
Dim VarDescIte As Currency
'Dim Comissao As Currency

Dim PercDesc As Double
Dim VarSaldoAPrazo As Currency
Dim VarMonoRet As String
Dim VarTotMonoRet As String
Dim VarTotQtdMono As String
Dim VarTpag As Integer

VarMonoRet = 0
VarTotMonoRet = 0
VarTotQtdMono = 0

VarSaldoAPrazo = MebTotal


Dim varSubTotal2 As String
If LerINI("Sistema", "loja", App.Path & "\Amplacard.ini") <> "" Then
    VarLoja = LerINI("sistema", "Loja", App.Path & "\Amplacard.ini")
End If

If LerINI("Sistema", "Caixa", App.Path & "\Amplacard.ini") <> "" Then
    VarCxa = LerINI("sistema", "Caixa", App.Path & "\Amplacard.ini")
End If

VarDescIte = 0

VarNroNota = 0
VarCodDev = 0
VarPedDec = 0
VarDescTot = 0

VarNroCupom = ""
Dim Linhas As Integer

VarEst = vbYes

varSubTotal = 0
VarTotal = 0
Linhas = 0

Cupom = "0"


If VarOrcamento = 0 Then
    If txtpgto.ListIndex = 7 Then
'        txtpgto = ""
        TxtObs2.SetFocus
        Exit Sub
    End If
End If

If Not IsNumeric(txtpgto.ListIndex) Then
'    txtpgto = ""
    TxtObs2.SetFocus
    Exit Sub
End If

If Impressora <> 9 Then
    If Impressora <> 99 Then
        If Impressora <> 8 Then
            If txtpgto.ListIndex > 5 Then
                MsgBox "Operação não aceita na impressora fiscal", vbInformation, App.Title
                Exit Sub
                txtpgto.SetFocus
            End If
        End If
    End If
End If

If TotPgto < CCur(Format(MebTotal, "###,##0.00")) - MebDesconto Then
    MsgBox "Total de Pgto incompleto", vbInformation
    Exit Sub
End If

If VarConveniencia = True Then
    If txtpgto.ListIndex > 5 Then
        MsgBox "Operação não aceita !! ", vbInformation, App.Title
        Exit Sub
        txtpgto.SetFocus
    End If
End If

If (CCur(MebTotal) - MebDesconto) > TotPgto Then
    MebTroco = TotPgto - (CCur(MebTotal) - MebDesconto)
End If

If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then
     Set VarPopular = FrmFrente.Banco.OpenRecordset("Select * From Popular Where Pedido = " & LblTot)
     VarPopular.Requery
     If VarPopular.RecordCount > 0 Then
        If txtpgto.ListIndex = 0 Then
            MsgBox "Operacao não aceita na popular !", vbInformation, App.Path
            txtpgto.SetFocus
            Exit Sub
'            TxtPgto = 6
        End If
     End If
End If

If VarBloqueado = True And txtpgto.ListIndex = 5 Then
    MsgBox "Cliente bloqueado pra venda a prazo !", vbInformation, App.Path
    txtpgto.SetFocus
    Exit Sub
End If


VarDesc = CCur(MebValor) - CCur(MebDesconto)
VarDesc = ((VarDesc / CCur(MebValor)) * 100)
If VarDesc < (100 - LimiteDesc) Then
    MsgBox "Desconto maior que o permitido", vbInformation, App.Title
    MebDesc = 0
    MebDesconto = 0
    MebValorDig = ""
    MebTroco = ""
    MebDesc.SetFocus
    Exit Sub
End If


If txtpgto.ListIndex = 1 Or txtpgto.ListIndex = 2 Then
    If ConsultaCh <> "calc" Then
        If MsgBox("Consultar Cheque ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
               Shell ConsultaCh, vbHide
        End If
    End If
End If

If VarDtVenc = "" Then
    If FrmFrente.Clientes!Prazo > 0 Then
        If FrmFrente.Clientes!Prazo <> "" Then
            VarDtVenc = Format(Date + FrmFrente.Clientes!Prazo, "dd/mm/yyyy")
        Else
            VarDtVenc = Format(Date + 30, "dd/mm/yyyy")
        End If
    Else
        VarDtVenc = Format(Date + 30, "dd/mm/yyyy")
    End If
    If IsNumeric(FrmFrente.LblPedido) Then
        If FrmFrente.LblPedido <> 0 Then
                sql = "DELETE DUP.Pedido, * From Dup Where Dup.Pedido  = " & FrmFrente.LblPedido
                FrmFrente.Banco.Execute sql
        End If
    End If
End If

If KeyAscii = 13 And Flag = True Then
    If txtpgto.ListIndex >= 0 And txtpgto.ListIndex <= 9 Then
        Resp = MsgBox("Confirma esta operação ?", vbQuestion + vbYesNoCancel + vbDefaultButton3)
        If Resp = vbYes Then
'            TxtPgto.Enabled = False
            Flag = False
            sql = "UPDATE PED SET PED.Entidade = " & FrmFrente.LblCodigo & " , ped.DataEmiss = #" & Format(FrmFrente.DtData, "mm/dd/yyyy") & "# WHERE PED.CodPed = " & Val(FrmFrente.CodTotalizador)
            FrmFrente.Banco.Execute sql
            
            If txtpgto.ListIndex < 5 Then
                Pago = -1
                Dias = 0
            Else
                Pago = 0
                Dias = 30
            End If
            With FrmFrente.Pedido
                If .RecordCount > 0 Then
                    Criterio = "CodPed = " & LblTot
                    .FindFirst Criterio
                    .Edit
                    .Fields("Total") = MebTotal.Text
                    .Fields("Desconto") = MebDesconto.Text
                    Select Case txtpgto.ListIndex
                        Case 0
                            .Fields("CondPgto") = "DINHEIRO"
                        Case 1
                            .Fields("CondPgto") = "CHEQUE"
                        Case 2
                            .Fields("CondPgto") = "PIX/TRANSF"
                        Case 3
                            .Fields("CondPgto") = "CARTAO DEBITO"
                        Case 4
                            .Fields("CondPgto") = "CARTAO CREDITO"
                        Case 5
                            .Fields("CondPgto") = "A PRAZO"
                        Case 6
                            .Fields("CondPgto") = "ORCAMENTO"
                        Case 7
                            .Fields("CondPgto") = "PARCELADO"
                        Case 8
                            .Fields("CondPgto") = "DEVOLUCAO"
                            VarCodDev = FrmFrente.LblCodigo
                            VarPedDec = FrmFrente.LblPedido
                        Case 9
                            .Fields("CondPgto") = "CONSIGNADO"
                    End Select
                    .Update
                End If
            End With
                                                                                                
            MebComprove = 0
            x = 0
            For x = 0 To XPgto - 1
                If Mid(ListPgto.List(x), 4, 10) <> "Consignado" Then
                    If CInt(Left(ListPgto.List(x), 1) - 1) < 5 Then
                        sql = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Maquina  ) VALUES ( "
                        sql = sql & FrmFrente.LblCodigo & " , "
                        sql = sql & "#" & Format(Date, "mm/dd/yyyy") & "# , "
                        sql = sql & """" & Time & """" & " , "
                        sql = sql & Numero(CCur(Right(ListPgto.List(x), 9)), 2) & " , "
                        sql = sql & Numero(CStr(MebDesconto), 2) & " , "
                        sql = sql & Left(ListPgto.List(x), 1) + 10 & " ,"
                        sql = sql & Val(FrmFrente.CodTotalizador) & " ,"
                        sql = sql & """" & Trim(LblMaquina) & """" & ")"
                        FrmFrente.Banco.Execute sql
                    End If
                End If
'                If FrmFrente.LblCodigo <> 1 Then
                    If CInt(Left(ListPgto.List(x), 1) - 1) = 8 Then
                            If MsgBox("Deseja devolver o valor na conta do cliente ?", vbYesNo, App.Title) = vbYes Then

                                sql = "INSERT INTO DUP ( Entidade, DtEmiss, DtVenc , Valor,  Pedido, Tipo_Dup, NroPcl, Cond ) VALUES ( "
                                sql = sql & FrmFrente.LblCodigo & " , #"
                                sql = sql & Format(FrmFrente.LblData, "mm/dd/yyyy") & "# , #"
                                sql = sql & Format(FrmFrente.LblData, "mm/dd/yyyy") & "# , "
                                sql = sql & Numero(CCur(Right(ListPgto.List(x), 9)) * (-1), 2) & " ,"
                                sql = sql & Val(FrmFrente.CodTotalizador) & " ,"
                                sql = sql & """" & "DV" & """" & " ,"
                                sql = sql & 1 & " ,"
                                sql = sql & Pago & ")"
                                FrmFrente.Banco.Execute sql
                           End If
                        ElseIf CInt(Left(ListPgto.List(x), 1) - 1) = 5 Then
                            sql = "INSERT INTO DUP ( Entidade, DtEmiss, DtVenc , Valor,  Pedido, Tipo_Dup, NroPcl, Cond ) VALUES ( "
                            sql = sql & FrmFrente.LblCodigo & " , #"
                            sql = sql & Format(FrmFrente.LblData, "mm/dd/yyyy") & "# , #"
                            sql = sql & Format(VarDtVenc, "mm/dd/yyyy") & "# , "
'                            If MebComprove <> "" Then
'                                sql = sql & Numero(MebComprove.Text, 2) & " ,"
'                            Else
                                sql = sql & Numero(CCur(Right(ListPgto.List(x), 9)), 2) & " ,"
                                MebComprove = MebComprove + (CCur(Right(ListPgto.List(x), 9)))
'                            End If
                            sql = sql & Val(FrmFrente.CodTotalizador) & " ,"
                            sql = sql & """" & "PC" & """" & " ,"
                            sql = sql & 1 & " ,"
                            sql = sql & 0 & ")"
                            FrmFrente.Banco.Execute sql
'                            MebComprove = MebComprove + CCur(Right(ListPgto.List(x), 9)) - CCur(MebTroco)
                            
                    End If
                    DatPcl.RecordSource = "SELECT DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.Valor, DUP.Cond FROM DUP WHERE DUP.Pedido = " & FrmFrente.CodTotalizador & " ORDER BY DUP.NroPcl"
                    DatPcl.Refresh
'                    if datpcl.Recordset.RecordCount
'                End If
                If CInt(Left(ListPgto.List(x), 1) - 1) = 8 And Vista = True Then
                    Set Caixa = FrmFrente.Banco.OpenRecordset("Select * From ContCxa Order by CodControle")
                    Caixa.Requery
                    If Caixa.RecordCount > 0 Then
                        Caixa.MoveLast
                        LblCaixa = Caixa!CodControle
                    End If
                
                    If MsgBox("Devolução de venda a vista, deseja lançar o valor como retirada do caixa ?", vbYesNo, App.Title) = vbYes Then
                        sql = "INSERT INTO CXA ( Maquina , Historico, Data, Hora, Valor, NroCaixa, D_C ) VALUES ( "
                        sql = sql + """" & LblMaquina & """" & ", "
                        sql = sql & """" & "DEVOLUÇÃO DE MERCADORIA" & """" & " , #"
                        sql = sql & Format(Date, "mm/dd/yyyy") & "# , #"
                        sql = sql & Format(Time, "hh:mm:ss") & "# , "
                        sql = sql & Numero(CCur(Right(ListPgto.List(x), 9)), 2) & " ,"
                        sql = sql & LblCaixa & " ,"
                        sql = sql & """" & "D" & """" & ")"
                        FrmFrente.Banco.Execute sql
                    End If
                                                            
                End If
            
            Next
                                                                                                                                                                                               
            If MebComprove = 0 Then
                MebComprove = ""
            End If
            Criterio = "CodPed = " & FrmFrente.CodTotalizador
            FrmFrente.Pedido.FindFirst Criterio
            If Not FrmFrente.Pedido.NoMatch Then
                FrmFrente.Pedido.Edit
                FrmFrente.Pedido.Fields("Total") = MebTotal
                If Pago Then
                    FrmFrente.Pedido.Fields("Operacao") = 1
                Else
                    FrmFrente.Pedido.Fields("Operacao") = 2
                End If
                If txtpgto.ListIndex = 6 Then
                    FrmFrente.Pedido.Fields("Operacao") = 4
                End If
                If txtpgto.ListIndex = 8 Then
                    FrmFrente.Pedido.Fields("Operacao") = 10
                End If
                
                If txtpgto.ListIndex = 9 Then
                    FrmFrente.Pedido.Fields("Operacao") = 5
                End If
                
                FrmFrente.Pedido.Fields("Obs") = TxtObs.Text & "  " & TxtObs2.Text
                FrmFrente.Pedido.Fields("ValorFrete") = 0
                FrmFrente.Pedido.Update
            End If
            

            Set VarDupTemp = FrmFrente.Banco.OpenRecordset("SELECT DUP.*, PED.Operacao, DUP.Pedido FROM PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE (PED.Operacao = 1 Or PED.Operacao = 4) AND DUP.Pedido= " & LblTot)
            VarDupTemp.Requery
            If VarDupTemp.RecordCount > 0 Then
                sql = "DELETE DUP.*, PED.Operacao, DUP.Pedido FROM PED INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE (PED.Operacao = 1 Or PED.Operacao =4) AND DUP.Pedido = " & LblTot
                FrmFrente.Banco.Execute sql
            End If

                                                                                                                                                                                                           
            If PortaNF = "NENHUMA" And Impressora = 99 Then
                Resp = vbNo
            ElseIf Impressora = 99 Then
                    Resp = MsgBox("Impressora pronta ?", vbYesNo, App.Title)
                Else
                    Resp = vbYes
            End If
            
            
            If txtpgto.ListIndex < 6 Or txtpgto.ListIndex = 7 Then
                 If VarBaixaAuto = "Nao" Then
                     VarEst = MsgBox("Produto entregue ?", vbYesNo, App.Title)
                 Else
                     VarEst = vbYes
                 End If
             End If
             
             If VarEst = vbYes Then
                      If txtpgto.ListIndex <> 6 And txtpgto.ListIndex <> 9 Then
                         VoltaEstoque LblTot
                      Else
                         If VarOrcaPrazo = True Then
                            VoltaEstoque LblTot
                         End If
                      End If
                 FrmFrente.DatCompra.Recordset.MoveFirst
                 On Error Resume Next
                 While Not FrmFrente.DatCompra.Recordset.EOF
                      If txtpgto.ListIndex <> 6 And txtpgto.ListIndex <> 9 Then
                             If txtpgto.ListIndex = 8 Then
                                 BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto"), FrmFrente.DatCompra.Recordset.Fields("Loja"), False
                             Else
                                 BaixarEstoque FrmFrente.DatCompra.Recordset.Fields("Qtd"), FrmFrente.DatCompra.Recordset.Fields("Produto"), FrmFrente.DatCompra.Recordset.Fields("Loja"), True
                             End If
                             
                             FrmFrente.DatCompra.Recordset.Edit
                             FrmFrente.DatCompra.Recordset("Entregue") = 1
                             FrmFrente.DatCompra.Recordset.Update
                             
                      Else
                          FrmFrente.DatCompra.Recordset.Edit
                          FrmFrente.DatCompra.Recordset("Entregue") = 0
                          FrmFrente.DatCompra.Recordset.Update
                      End If
                     FrmFrente.DatCompra.Recordset.MoveNext
                 Wend
             End If


            If NroCupom = "" Then NroCupom = 0
            
            If Impressora < 8 And IsNumeric(NroCupom) Then
                sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.NroNt = " & NroCupom & ", PED.Cfo = 'CUPOM' WHERE PED.CodPed = " & Val(FrmFrente.CodTotalizador)
                FrmFrente.Banco.Execute sql
            End If
            
            If Resp = vbYes Then
              Dim VarContPrd As Integer
 
               Select Case Impressora
                    Case 1
                    
                    Case 5
                    Case 6
                    
                    Case Is < 10
                        
                        ' AcbrMonitor
'                        Dim VarContPrd As Integer
                        VarContPrd = 1
                        ' Venda do Item
                        If dir(App.Path & "\" & Trim(Left(LimpaCampo(FrmFrente.Maquina), 8)) & ".TXT") <> "" Then
                           Kill App.Path & "\" & Trim(Left(LimpaCampo(FrmFrente.Maquina), 8)) & ".TXT"
                        End If
                        If dir(App.Path & "\Cfe.Ini") <> "" Then
                           Kill App.Path & "\Cfe.ini"
                        End If
                        
                        If dir(PortaAcbr & "\Sai.txt") <> "" Then
                            Kill PortaAcbr & "\Sai.txt"
                        End If
                        VarImprSat = ""
                        PercDesc = 0
                        VarDesc = 0
                        VarTotImpostoNac = 0
                        VarTotImpostoEst = 0
                        varSubTotal2 = 0
                        If Impressora = 9 Then
                                EscreveINI "InfCfe", "Versao", VarVersao, App.Path & "\cfe.ini"
            '                   Identificacao
                                EscreveINI "Identificacao", "Cnpj", VarCnpjDesenvolvedor, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "signAC", VarChave, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "NumeroCaixa", VarCaixaSat, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "NCfe", FrmFrente.LblPedido, App.Path & "\cfe.ini"
            '                   Emitente
                                EscreveINI "Emitente", "Cnpj", VarCnpjEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IE", VarInscEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IM", VarInscrMuni, App.Path & "\cfe.ini"
        '                        EscreveINI "Emitente", "CRT", "1", App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "indRatISSQN", "S", App.Path & "\cfe.ini"
                        Else
                        
                                Dim CodChave As Double
                                CodChave = Int(Mid(Int(CCur(Mid(Time, 7, 2)) * (Rnd * 100000)), 1, 6))
            '                   Identificacao
                                EscreveINI "Identificacao", "natOp", "Venda de Mercadoria", App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "mod", "65", App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "serie", VarSerieNfce, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "dEmi", Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS"), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "nNF", psCompString(1, MebNfe, 9), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "cNF", psCompString(1, CodChave, 9), App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "tpNF", 1, App.Path & "\cfe.ini"
        '                        EscreveINI "Identificacao", "indPag", TipoPgto, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "indFinal", 1, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "indPres", 1, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "tpImp", 4, App.Path & "\cfe.ini"
                                EscreveINI "Identificacao", "finNfe", 1, App.Path & "\cfe.ini"
            
            '                   Emitente
                                EscreveINI "Emitente", "Cnpj", VarCnpjEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "IE", VarInscEmite, App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xNome", psCompString(3, VarRazaoEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xFant", psCompString(3, VarFantasiaEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "Fone", psCompString(3, VarTelEmite, 10), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "CEP", psCompString(1, VarCepEmite, 8), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xLgr", psCompString(3, VarEndEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "nro", psCompString(3, VarNroEmite, 10), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "cMun", Trim(VarIbgeEmite), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xMun", psCompString(3, VarCidadeEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "xBairro", psCompString(3, VarBairroEmite, 60), App.Path & "\cfe.ini"
                                EscreveINI "Emitente", "UF", Left(VarUfEmite, 2), App.Path & "\cfe.ini"

                                
                                EscreveINI "Emitente", "CRT", 1, App.Path & "\cfe.ini"
                                                
                        End If
    '                   Destinatario
                        If Impressora <> 9 Then
                            EscreveINI "Destinatario", "indIEDest", "9", App.Path & "\cfe.ini"
                        End If
                        EscreveINI "Destinatario", "CNPJCPF", LIMPACampo2(MebDoc), App.Path & "\cfe.ini"
                        EscreveINI "Destinatario", "XNome", Left(FrmFrente.Clientes.Fields("Nome"), 40), App.Path & "\cfe.ini"
                        
                        If Trim(Left(FrmFrente.Clientes.Fields("Nome"), 40)) <> "CONSUMIDOR" Then
        '                   Entrega
                            If FrmFrente.Clientes.Fields("Endereco") <> "" Then
                                EscreveINI "Destinatario", "XLgr", RTrim(FrmFrente.Clientes.Fields("Endereco")), App.Path & "\cfe.ini"
                            End If
                            
                            If FrmFrente.Clientes.Fields("Nro") <> "" Then
                                EscreveINI "Destinatario", "nro", RTrim(FrmFrente.Clientes.Fields("Nro")), App.Path & "\cfe.ini"
                            End If
                            
    '                        EscreveINI "Entrega", "XCpl", "", App.Path & "\cfe.ini"
                            
                            If FrmFrente.Clientes.Fields("Bairro") <> "" Then
                                EscreveINI "Destinatario", "XBairro", RTrim(FrmFrente.Clientes.Fields("Bairro")), App.Path & "\cfe.ini"
                            Else
    '                            EscreveINI "Entrega", "XBairro", "", App.Path & "\cfe.ini"
                            End If
                            If FrmFrente.Clientes.Fields("Cidade") <> "" Then
                                EscreveINI "Destinatario", "XMun", RTrim(FrmFrente.Clientes.Fields("Cidade")), App.Path & "\cfe.ini"
                            Else
    '                            EscreveINI "Entrega", "XMun", "", App.Path & "\cfe.ini"
                            End If
                            If FrmFrente.Clientes.Fields("Uf") <> "" Then
                                EscreveINI "Destinatario", "UF", RTrim(FrmFrente.Clientes.Fields("Uf")), App.Path & "\cfe.ini"
                            Else
    '                            EscreveINI "Entrega", "UF", "", App.Path & "\cfe.ini"
                            End If
                            If FrmFrente.Clientes.Fields("Matricula") <> "" Then
                                EscreveINI "Destinatario", "cMun", RTrim(FrmFrente.Clientes.Fields("Matricula")), App.Path & "\cfe.ini"
                            Else
    '                            EscreveINI "Entrega", "UF", "", App.Path & "\cfe.ini"
                            End If
                            
                        End If
    '                   Produto
                        VarTotImposto = 0
                        
                        FrmFrente.DatCompra.Recordset.MoveFirst
                        While Not FrmFrente.DatCompra.Recordset.EOF
                            Criterio = "Produto = " & FrmFrente.DatCompra.Recordset("Produto")
                            FrmFrente.Registros.FindFirst Criterio
                            
                            Criterio = " Codigo = " & """" & Trim(LIMPACampo2(FrmFrente.DatCompra.Recordset.Fields("Classe"))) & """"
                            FrmFrente.DatMinutos.Recordset.FindFirst Criterio
                            If Not FrmFrente.DatMinutos.Recordset.NoMatch Then
                                VarImpostoNac = ((FrmFrente.DatCompra.Recordset.Fields("Qtd") * FrmFrente.DatCompra.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqNac") / 100))
                                VarImpostoEst = ((FrmFrente.DatCompra.Recordset.Fields("Qtd") * FrmFrente.DatCompra.Recordset.Fields("ValorUnit")) * (FrmFrente.DatMinutos.Recordset("AliqEst") / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            Else
                                VarImpostoNac = ((FrmFrente.DatCompra.Recordset.Fields("Qtd") * FrmFrente.DatCompra.Recordset.Fields("ValorUnit")) * (VarMediaImpostoNac / 100))
                                VarImpostoEst = ((FrmFrente.DatCompra.Recordset.Fields("Qtd") * FrmFrente.DatCompra.Recordset.Fields("ValorUnit")) * (VarMediaImpostoEst / 100))
                                VarTotImpostoNac = VarTotImpostoNac + VarImpostoNac
                                VarTotImpostoEst = VarTotImpostoEst + VarImpostoEst
                            End If
                                                                                    
                            ConvertQtd = Format(FrmFrente.DatCompra.Recordset.Fields("Qtd"), "0.000")
                            quantidade = FrmFrente.DatCompra.Recordset.Fields("Qtd")
                            
                            ConvertVr = Format(FrmFrente.DatCompra.Recordset.Fields("ValorUnit"), "0.000")
                            varSubTotal2 = varSubTotal2 + CCur(Format((ConvertQtd * ConvertVr), "0.00"))
                                                        
                                                        
                            If IsNull(FrmFrente.Registros("BarBarra")) Then
                                CodigoBarra = 1
                            Else
                                CodigoBarra = Left(FrmFrente.Registros("BarBarra"), 13)
                            End If
                            
                           If MebDesconto <> 0 Then
                                PercDesc = Round((CCur(MebDesconto) / CCur(MebValor)), 5)
                                VarValordesc = ConvertVr * ConvertQtd
                                VarDesc = Round(VarValordesc * (CCur(PercDesc)), 2)
                            Else
                                VarDesc = 0
                            End If
                                  
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cProd", FrmFrente.DatCompra.Recordset("Produto"), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "xProd", Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 29), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "cEan", "", App.Path & "\cfe.ini"
                            
                            If FrmFrente.DatCompra.Recordset("Classe") = "" Or Len(Trim(FrmFrente.DatCompra.Recordset("Classe"))) < 8 Then
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", "54023390", App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "NCM", Trim(LimpaNro(FrmFrente.DatCompra.Recordset("Classe"))), App.Path & "\cfe.ini"
                            End If
                            
                            If FrmFrente.DatCompra.Recordset("CfopPrd") = "" Or FrmFrente.DatCompra.Recordset("CfopPrd") = 0 Then
'                                If FrmFrente.DatCompra.Recordset("CfopPrd") <> "5102" Then
                                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", "5102", App.Path & "\cfe.ini"
'                                Else
'                                    EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", FrmFrente.DatCompra.Recordset("CfopPrd"), App.Path & "\cfe.ini"
'                                End If
                            Else
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "CFOP", FrmFrente.DatCompra.Recordset("CfopPrd"), App.Path & "\cfe.ini"
                            End If
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "uCom", IIf(FrmFrente.DatCompra.Recordset("Embalagem") = "", "UN", FrmFrente.DatCompra.Recordset("Embalagem")), App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "Combustivel", 0, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "qCom", ConvertQtd, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vUnCom", ConvertVr, App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "indRegra", "A", App.Path & "\cfe.ini"
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vDesc", Format(CCur(VarDesc), "####00.00"), App.Path & "\cfe.ini"
                            VarDescTot = VarDescTot + CCur(VarDesc)
                            If Impressora <> 9 Then
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vprod", CCur(Format((ConvertQtd * ConvertVr), "0.00")), App.Path & "\cfe.ini"
                            End If
                           
                            If ChkComissao.Value = 0 Then
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vOutro", 0, App.Path & "\cfe.ini"
                            Else
                                EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vOutro", FrmFrente.DatCompra.Recordset.Fields("Qtd") * FrmFrente.DatCompra.Recordset.Fields("ValorUnit") * (MebComissao / MebValor), App.Path & "\cfe.ini"
                            End If
                            EscreveINI "Produto" & psCompString(1, VarContPrd, 3), "vItem23741", CCur(VarImpostoNac + VarImpostoEst), App.Path & "\cfe.ini"
                            
                            
                            If FrmFrente.DatCompra.Recordset("CfopPrd") = "5656" Or FrmFrente.DatCompra.Recordset("CfopPrd") = "5.656" Or FrmFrente.DatCompra.Recordset("CfopPrd") = "5655" Or FrmFrente.DatCompra.Recordset("CfopPrd") = "5.655" Or FrmFrente.DatCompra.Recordset("CfopPrd") = "5661" Or FrmFrente.DatCompra.Recordset("CfopPrd") = "6661" Then
                                If Trim(LIMPACampo2(FrmFrente.DatCompra.Recordset.Fields("Classe"))) = "27111910" Then
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "210203001", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "GLP", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGLP", "40.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNn", "30.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "pGNi", "30.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "vPart", "5.00", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                                    VarMonoRet = 1.4139
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", Right("061", 2), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "qBCMonoRet", Numero2(FrmFrente.DatCompra.Recordset.Fields("Qtd")), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "adRemICMSRet", Numero2(VarMonoRet), App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "vICMSMonoRet", Numero2(FrmFrente.DatCompra.Recordset.Fields("Qtd") * VarMonoRet), App.Path & "\cfe.ini"
                                    VarTotMonoRet = VarTotMonoRet + FrmFrente.DatCompra.Recordset.Fields("Qtd") * VarMonoRet
                                    VarTotQtdMono = VarTotQtdMono + FrmFrente.DatCompra.Recordset.Fields("Qtd")
                                    
                                    
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "indImport", "0", App.Path & "\cfe.ini"
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "cUFOrig", "35", App.Path & "\cfe.ini"
                                    EscreveINI "ORIGCOMB" & psCompString(1, VarContPrd, 3) & "01", "pOrig", "100.00", App.Path & "\cfe.ini"
                                    
                                    
                                    
                                End If
                                
                                If Trim(LIMPACampo2(FrmFrente.DatCompra.Recordset.Fields("Classe"))) = "27101932" Then
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "CprodANP", "620505001", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "descANP", "OUTROS OLEOS LUBRIFICANTES AUTOMOTIVOS", App.Path & "\cfe.ini"
                                    EscreveINI "COMB" & psCompString(1, VarContPrd, 3), "UFCons", Left(VarUfEmite, 2), App.Path & "\cfe.ini"
                                End If
                            End If
                            EscreveINI "ObsFiscoDet" & psCompString(1, VarContPrd, 3) & psCompString(1, VarContPrd, 3), "xCampoDet", "", App.Path & "\cfe.ini"
                            EscreveINI "ObsFiscoDet" & psCompString(1, VarContPrd, 3) & psCompString(1, VarContPrd, 3), "xTextoDet", "", App.Path & "\cfe.ini"
    '                       Icms
                            EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "Orig", 0, App.Path & "\cfe.ini"
                            If VarRegime = 3 Then
                                If FrmFrente.DatCompra.Recordset("SitTrib") = "000" Then
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "00", App.Path & "\cfe.ini"
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "pICMS", FrmFrente.DatCompra.Recordset("Icm"), App.Path & "\cfe.ini"
                                
                                Else
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CST", "60", App.Path & "\cfe.ini"
                                End If
                                
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", "01", App.Path & "\cfe.ini"
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "vBC", (ConvertQtd * ConvertVr), App.Path & "\cfe.ini"
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "pPIS", VarAliqPis / 100, App.Path & "\cfe.ini"
        '                       Cofins
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", "01", App.Path & "\cfe.ini"
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "vBC", (ConvertQtd * ConvertVr), App.Path & "\cfe.ini"
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "pCOFINS", VarAliqCofins / 100, App.Path & "\cfe.ini"
                            
                            
                            
                            Else
                                If FrmFrente.DatCompra.Recordset("SitTrib") = "" Or FrmFrente.DatCompra.Recordset("SitTrib") = 0 Then
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", "102", App.Path & "\cfe.ini"
    
                                    EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "ICMSSN102", "", App.Path & "\cfe.ini"
                                Else
                                    If Trim(LIMPACampo2(FrmFrente.DatCompra.Recordset.Fields("Classe"))) = "27111910" Then

                                    Else
                                        EscreveINI "ICMS" & psCompString(1, VarContPrd, 3), "CSOSN", FrmFrente.DatCompra.Recordset("SitTrib"), App.Path & "\cfe.ini"
                                    End If
                                End If
                                
                                EscreveINI "PIS" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
        '                       Cofins
                                EscreveINI "Cofins" & psCompString(1, VarContPrd, 3), "CST", "49", App.Path & "\cfe.ini"
                                                      
                            End If
                                                                     
                            FrmFrente.DatCompra.Recordset.MoveNext
                            VarContPrd = VarContPrd + 1
                        Wend
                        
                         MebDesconto = VarDescTot
                         EscreveINI "DescAcrEntr", "VDescSubTotal", Format(MebDesconto, "####00.00"), App.Path & "\cfe.ini"
                         EscreveINI "DescAcrEntr", "VOutroSubTotal", MebComissao, App.Path & "\cfe.ini"
                         If Impressora <> 9 Then
                            EscreveINI "TOTAL", "vProd", Format(varSubTotal2, "####00.00"), App.Path & "\cfe.ini"
                            EscreveINI "TOTAL", "vDesc", Format(MebDesconto, "####00.00"), App.Path & "\cfe.ini"
                             
                            EscreveINI "TOTAL", "vnf", Format(varSubTotal2 - MebDesconto, "####00.00"), App.Path & "\cfe.ini"
                            EscreveINI "TRANSPORTADOR", "modFrete", 9, App.Path & "\cfe.ini"
                            If VarTotMonoRet <> 0 Then
                                EscreveINI "TOTAL", "vICMSMonoRet", Numero2(VarTotMonoRet), App.Path & "\cfe.ini"
                                EscreveINI "TOTAL", "qBCMonoRet", Numero2(VarTotQtdMono), App.Path & "\cfe.ini"
                            End If

                         End If
                        x = 0
                        ListPgto.Move Top
                        For x = 0 To XPgto - 1
                            Select Case CInt(Left(ListPgto.List(x), 1)) - 1
                                Case 0
                                    VarTpag = 1
                                Case 1
                                    VarTpag = 2
                                Case 2
                                    VarTpag = 17
                                Case 3
                                    VarTpag = 4
                                Case 4
                                    VarTpag = 3
                                Case 5
                                    VarTpag = 5
                                Case 7
                                    VarTpag = 21
                            End Select
                            EscreveINI "Pag" & psCompString(1, x + 1, 3), "tPag", psCompString(1, VarTpag, 2), App.Path & "\cfe.ini"
                            If VarTpag = 3 Or VarTpag = 4 Or VarTpag = 17 Then
                                EscreveINI "Pag" & psCompString(1, x + 1, 3), "tpIntegra", psCompString(1, "2", 1), App.Path & "\cfe.ini"
                            End If
                           MebTroco = MebTotal - (varSubTotal2 - Round(MebDesconto, 2))
                            
    
                            If ListPgto.ListCount > 1 Then
                               EscreveINI "Pag" & psCompString(1, x + 1, 3), "vPag", Format(CCur(Right(ListPgto.List(x), 9)), "####00.00"), App.Path & "\cfe.ini"
                            Else
'                               MebTroco = TotPgto - (varSubTotal2 - MebDesconto)
                               If MebTroco <> 0 Then
                                    If MebTroco < 0 Then
                                        EscreveINI "Pag001", "vPag", Format(varSubTotal2 - MebDesconto, "####00.00"), App.Path & "\cfe.ini"
                                        MebTroco = 0
                                    Else
                                        EscreveINI "Pag001", "vPag", Format(TotPgto, "####00.00"), App.Path & "\cfe.ini"
                                    End If
                                Else
                                    EscreveINI "Pag001", "vPag", Format(varSubTotal2, "####00.000"), App.Path & "\cfe.ini"
                                End If
                            End If
                         Next
                         If MebTroco <> 0 Then
                             EscreveINI "Pag" & "001", "vTroco", Format(MebTroco, "####00.00"), App.Path & "\cfe.ini"
                         End If
                         
                         TxtImposto = "Tributo R$ " & Format(VarTotImpostoNac, "#,##0.00") & "(" & Format(((VarTotImpostoNac / CCur(MebValor)) * 100), "##0.00") & "%) Federal e R$ " & Format(VarTotImpostoEst, "#,##0.00") & "(" & Format(((VarTotImpostoEst / CCur(MebValor)) * 100), "##0.00") & "%)" & " Estadual e R$ 0,00(0,00%) Municipal. Fonte:IBPT"
                         
                         EscreveINI "DadosAdicionais", "InfCpl", TxtImposto & " " & Trim(TxtObs) & " " & Trim(TxtObs2), App.Path & "\cfe.ini"
                         EscreveINI "ObsFiscoDet001", "xCampo", "", App.Path & "\cfe.ini"
                         EscreveINI "ObsFiscoDet001", "xTexto", "", App.Path & "\cfe.ini"
                        
                       If MsgBox("Imprimir Cupom ?", vbYesNo, App.Title) = vbYes Then
                            AbreAcbr "NFE.CriarEnviarNFe(" & App.Path & "\cfe.ini" & " , 1, 1, 1,,0)"
                       Else
                            AbreAcbr "NFE.CriarEnviarNFe(" & App.Path & "\cfe.ini" & " , 1, 0, 1,,0)"
                       End If
    
                       If TrataAcbr2("Nota", MebNfe) = 100 Then
                           MsgBox "NFc-e enviada com sucesso !", vbInformation, App.Title
                           
                           sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.Cfo = 'Nfc-e', Ped.PrevEntrega = ' ', Ped.NroNt = " & MebNfe & ", Ped.Sat = " & """" & VarChaveNfe & """" & "  WHERE PED.CodPed = " & LblTot
                           FrmFrente.Banco.Execute sql
                           EscreveINI "Nfce", "NroNfce", CCur(MebNfe) + 1, App.Path & "\config.ini"
                           VarNroCupom = CCur(MebNfe)
                           
                            If txtpgto.ListIndex = 5 And VarComprovante = True Then
                                If MsgBox("Imprimir comprovante ?", vbYesNo, App.Title) = vbYes Then
                    
                                    Iniciar (False)
                                    Print #1, "======================================="
                                    Print #1, " Pedido:"; Tab(10); psCompString(1, LblTot, 6); Tab(18); "Data: "; Tab(24); Date
                                    Print #1, ""
                                    Print #1, Tab(1); FrmFrente.DbcCliente
                                    Print #1, ""
                                    If MebComprove <> "" Then
                                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebComprove), 10, 2)
                                    ElseIf MebTotal <> "" Then
                                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebTotal), 10, 2)
                                     Else
                                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(FormConsultaPed.MebTotal), 10, 2)
                                    End If
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, "  Ass.Cli----------------------------"
                                    Print #1, "======================================="
                                    Print #1, ""
                                    
                                    Print #1, "" & Nor
                                    For x = 1 To NroLinha
                                        Print #1, " "
                                    Next
                                    Print #1, Chr(27) & "i"         '& Chr(27) & "v50"
                                                    
                                    Fim (False)
                '                    Unload FrmTotal
                
                                End If
                            End If
                                                              
                       Else
                           MsgBox "ATENÇÃO !!!! ERRO NO ENVIO DA NFC-e !", vbInformation, App.Title
                       End If
                                              
                      
                        
                    Case 99
                        
                        Set Empresa = FrmFrente.Banco.OpenRecordset("CFG")
                        Empresa.MoveFirst
                                                
                        If PortaNF = "TELA" Then
                                Dim RegPed As Recordset
                                Dim Sql2 As String

                                 Sql2 = "SELECT Ped.Ipi, Ped.PedSub,Ent.CodEntidade, ENT.Nome, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor, Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob , PED.ValorFrete, Ped.Desconto, ITP.CodItp, Ped.NroPedido, Prd.complemento FROM PRD INNER JOIN (ENT INNER JOIN (ITP INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON ENT.CodEntidade = PED.Entidade) ON PRD.CodPrd = ITP.Produto "
                                 Sql2 = Sql2 & " GROUP BY Ped.ipi, Ped.PedSub, Ent.CodEntidade, ENT.Nome, ENT.CPF, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Bairro, ENT.Uf, ENT.Fone, ENT.Fax, Ent.Celular, ENT.Cep, ENT.NomeVendedor , Ent.Complemento, Ent.InscrProd, PED.CodPed, PED.DataEmiss, PED.TipoFrete, ITP.Produto, Prd.Ipi, PRD.Descrição_Produto, ITP.Complemento, Prd.Embalagem, ITP.Entregue, ITP.Qtd, ITP.VlrVdaDia, ITP.Loja, ENT.Endereço_Entrega, ENT.Nro_Entrega, Ent.Cidade_Entrega, Ent.Cep_Entrega, Ent.Cep_Cob, Ent.Bairro_Entrega, Ent.Uf_Entrega, ENT.Endereço_Cob, ENT.Nro_Cob, Ent.Cidade_Cob, Ent.Bairro_Cob, Ent.Uf_Cob ,PED.ValorFrete,Ped.Desconto,PED.NroPedido,ITP.CodItp, Prd.complemento "
                                 Sql2 = Sql2 & " HAVING (((PED.CodPed)= " & LblTot & ")) ORDER BY ITP.CodItp"
                                 Set varPesquisa(1) = FrmFrente.Banco.OpenRecordset(Sql2)
                                 
                                 Set FrmRelPed.Relatorio.Recordset = varPesquisa(1)
                                     FrmRelPed.Obs2.Caption = TxtObs
                                     FrmRelPed.Obs3.Caption = TxtObs2
                                     FrmRelPed.ObsFinal.Caption = FrmFrente.Empresa("MensagemRodape")
                                     FrmRelPed.Vendedor.Caption = FrmFrente.DbcVendedor
                                     FrmRelPed.rpCabTitulo.Caption = FrmFrente.Empresa("Cabecalho1")
                                     FrmRelPed.rpCabInformacao.Caption = FrmFrente.Empresa("Cabecalho2")
                                     
                                     If FrmFrente.Clientes("Classificacao") = 2 Then
                                        FrmRelPed.RpCgc.Caption = Format(FrmFrente.Clientes("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                                        FrmRelPed.RpInscr.Caption = FrmFrente.Clientes("InscrEst")
                                    Else
                                        FrmRelPed.RpCgc.Caption = Format(FrmFrente.Clientes("Rg"), "@@.@@@.@@@-@")
                                        FrmRelPed.RpInscr.Caption = Format(FrmFrente.Clientes("Cpf"), "@@@.@@@.@@@.@@")
                                     End If
                                     FrmRelPed.CabGrupo(1).Mostrar = True
                                     FrmFrente.DatCompra.Recordset.MoveFirst
    
                                     
                                     If txtpgto.ListIndex = 6 Then
                                        FrmRelPed.ReportField10.Caption = "ORÇAMENTO"
                                     End If
                                     If txtpgto.ListIndex = 8 Then
                                        FrmRelPed.ReportField10.Caption = "DEVOLUÇÃO"

                                     End If
                                     
                                     
                                Me.Hide
                                FrmRelPed.Relatorio.Copias = VarCopias
                                FrmRelPed.Config
'                                Me.Show vbModal
                                
                        Else
                            Iniciar (False)
                                
                            If Left(PortaNF, 3) <> "COM" Then
                                If VarImpressoraNf = "TERMICA" Then
                                    If VarCozinha = True Then
                                        Print #1, "========================================"
                                        Print #1, "Data:" & FrmFrente.LblData & " Hora:" & Time; Tab(31); "Ped:"; psCompString(1, LblTot, 6)
                                        Print #1, "MESA/SENHA ===>> " & FrmFrente.TxtTmp2
                                        FrmFrente.DatCompra.Recordset.MoveFirst
                                        While Not FrmFrente.DatCompra.Recordset.EOF
                                            Print #1, Tab(1); Alinhar(Format(FrmFrente.DatCompra.Recordset("Qtd"), "##0"), 3, 0); Tab(4); Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 37)
                                            If Len(Trim(FrmFrente.DatCompra.Recordset("Complemento"))) > 0 Then
                                                Print #1, "   " & Trim(FrmFrente.DatCompra.Recordset("Complemento"))
'                                               Print #1, Tab(1); Alinhar(Format(FrmFrente.DatCompra.Recordset("Qtd"), "##0.000"), 7, 3); Tab(11); "X"; Tab(12); Alinhar(Format(FrmFrente.DatCompra.Recordset("VendaDia")), 7, 2); Tab(19); Alinhar(Format(FrmFrente.DatCompra.Recordset("VendaDia") * FrmFrente.DatCompra.Recordset("Qtd"), "##,##0.00"), 7, 2); Tab(29); "Desc."; Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 7, 2)
                                            End If
                                            FrmFrente.DatCompra.Recordset.MoveNext
                                        Wend
                                        Print #1, "========================================"
                                        Print #1, ""
                                        Print #1, ""
                                        Print #1, ""
                                        Print #1, Chr(27) & "i" ' & vbCrLf
                                        

                                    End If
                                    For x = 1 To VarVia
                                         Print #1, ""
                                         Print #1, "" ' & Cond
                                         Print #1, Centralizar(Left(FrmFrente.Empresa("Cabecalho1"), 40), 40)
                                         Print #1, Centralizar(Left(FrmFrente.Empresa("Cabecalho2"), 40), 40)
                                         Print #1, "========================================"
                                         Print #1, "Data:" & FrmFrente.LblData & " Hora:" & Time; Tab(31); "Ped:"; psCompString(1, LblTot, 6)
                                         If FrmFrente.TxtTmp2 <> "" Then
                                            Print #1, ""
                                            Print #1, "MESA/SENHA ===>> " & FrmFrente.TxtTmp2
                                            Print #1, ""
                                        End If
                                         
                                         Print #1, Tab(1); psCompString(1, FrmFrente.Clientes("CodEntidade"), 6); "-"; Left(FrmFrente.Clientes.Fields("Nome"), 34)
                                         If (FrmFrente.Clientes.Fields("Endereco")) <> "" Then
                                            Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Endereco"), 30)) & ", " & FrmFrente.Clientes("Nro")
                                         End If
                                         If FrmFrente.Clientes("Bairro") <> "" Then
                                            Print #1, Tab(1); FrmFrente.Clientes("Bairro")
                                         End If
                                         If FrmFrente.Clientes("CIDADE") <> "" Then
                                            Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Cidade"), 30)) & " - " & FrmFrente.Clientes("Uf")
                                         End If
                                        If VarLinhaDet = 1 Then
                                               Print #1, Tab(1); "QTD"; Tab(6); "DESCRICAO DO PRODUTO"; Tab(28); "UNIT."; Tab(36); "TOTAL"
                                         Else
                                             Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
                                             Print #1, Tab(1); "QTD"; Tab(17); "UNIT."; Tab(35); "TOTAL"
                                         End If
                                         Print #1, "----------------------------------------"
                                         
                                         FrmFrente.DatCompra.Recordset.MoveFirst
        
                                         While Not FrmFrente.DatCompra.Recordset.EOF
                                             Linhas = Linhas + 1
                                             If VarLinhaDet = 1 Then
                                                    Print #1, Tab(1); Format(FrmFrente.DatCompra.Recordset("Qtd"), " #######"); Tab(6); Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 18); Tab(26); Alinhar(Format(FrmFrente.DatCompra.Recordset("ValorUnit"), "##,##0.00"), 5, 2); Tab(33); Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 8, 2)
                                             Else
                                                 If Len(Trim(FrmFrente.DatCompra.Recordset("Complemento"))) > 0 Then
                                                    Print #1, Tab(1); psCompString(1, FrmFrente.DatCompra.Recordset("Produto"), 6); "-"; Left(FrmFrente.DatCompra.Recordset("Descrição_Produto") & "-" & Trim(FrmFrente.DatCompra.Recordset("Complemento")), 33)
                                                 Else
                                                    Print #1, Tab(1); psCompString(1, FrmFrente.DatCompra.Recordset("Produto"), 6); "-"; Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 33)
        '                                            Print #1, Tab(8); Mid(Trim(FrmFrente.DatCompra.Recordset("Descrição_Produto")) & "-" & Trim(FrmFrente.DatCompra.Recordset("Complemento")), 33, 33)
                                                 End If
                                                 If FrmFrente.DatCompra.Recordset("Desconto") = 0 Then
                                                       Print #1, Tab(1); Alinhar(Format(FrmFrente.DatCompra.Recordset("Qtd"), " ##,##0.000"), 7, 3); Tab(13); "X"; Tab(18); Alinhar(Format(FrmFrente.DatCompra.Recordset("ValorUnit"), "##,##0.00"), 9, 2); Tab(31); Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 8, 2)
                                                 Else
                                                       Print #1, Tab(1); Alinhar(Format(FrmFrente.DatCompra.Recordset("Qtd"), "##0.000"), 7, 3); Tab(11); "X"; Tab(12); Alinhar(Format(FrmFrente.DatCompra.Recordset("VendaDia")), 7, 2); Tab(19); Alinhar(Format(FrmFrente.DatCompra.Recordset("VendaDia") * FrmFrente.DatCompra.Recordset("Qtd"), "##,##0.00"), 7, 2); Tab(29); "Desc."; Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 7, 2)
                                                       VarDescIte = VarDescIte + (FrmFrente.DatCompra.Recordset("VendaDia") * FrmFrente.DatCompra.Recordset("Qtd") - FrmFrente.DatCompra.Recordset("Valor"))
                                                 End If
                                              End If
                                             FrmFrente.DatCompra.Recordset.MoveNext
                                         Wend
                                         
                                        If ChkComissao.Value = 0 Then
                                            If VarLinhaDet <> 1 Then
                                                Print #1, "ITENS : "; Linhas
                                            End If
                                            If (MebDesconto.Text + VarDescIte) <> 0 Then
                                                Print #1, Tab(17); "SUB-TOTAL =>"; Tab(31); Alinhar(CStr((MebValor.Text + CCur(VarDescIte))), 10, 2)
                                                Print #1, Tab(17); "DESCONTO  =>"; Tab(31); Alinhar(CStr(MebDesconto.Text + VarDescIte), 10, 2)
                                            End If
                                            Print #1, Tab(17); "TOTAL     =>"; Tab(31); Alinhar(CStr(MebTotal.Text), 10, 2)
                                         Else
    
                                            Print #1, Tab(17); "SUB-TOTAL =>"; Tab(31); Alinhar(CStr((MebValor.Text)), 10, 2)
                                            Print #1, Tab(17); "    10 %  =>"; Tab(31); Alinhar(CStr(MebComissao), 10, 2)
                                            Print #1, Tab(17); "TOTAL     =>"; Tab(31); Negi & Alinhar(CStr(MebTotal.Text), 10, 2) & NegF
                                            
                                         
                                         End If
                                         If VarLinhaDet <> 1 Then
                                            If txtpgto.ListIndex <> 5 Then
                                                If txtpgto.ListIndex = 6 Then
                                                    Print #1, Tab(1); "ORCAMENTO - " & FrmFrente.DbcVendedor
                                                ElseIf txtpgto.ListIndex = 7 Then
                                                        Print #1, Tab(1); "PARCELADO - " & FrmFrente.DbcVendedor
                                                    ElseIf txtpgto.ListIndex = 8 Then
                                                        Print #1, Tab(1); "DEVOLUCAO - " & FrmFrente.DbcVendedor
                                                    ElseIf txtpgto.ListIndex = 9 Then
                                                        Print #1, Tab(1); "CONSIGNADO - " & FrmFrente.DbcVendedor
                                                    Else
                                                        If MebValorDig <> MebTotal Then
                                                           Print #1, Tab(17); "RECEBIDO  =>"; Tab(31); Alinhar(CStr(MebValorDig), 10, 2)
                                                           If (MebTroco) <> 0 Then
                                                              Print #1, Tab(17); "TROCO     =>"; Tab(31); Alinhar(CStr(MebTroco), 10, 2)
                                                           End If
                                                        End If
                                                        Print #1, Tab(1); "A VISTA-" & FrmFrente.DbcVendedor
                                                End If
                                            ElseIf txtpgto.ListIndex = 5 Then
                                                Print #1, Tab(1); "A PRAZO - " & FrmFrente.DbcVendedor
                                            End If
                                         Else
                                            Print #1, ""
                                         End If
                                         If Not IsNull(TxtObs) And TxtObs <> "" Then
                                             Print #1, Tab(1); Left(TxtObs, 35)
                                         End If
                                         If Not IsNull(TxtObs2) And TxtObs <> "" Then
                                             Print #1, Tab(1); Left(TxtObs2, 35)
                                         End If
                                         
                                         Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 40)
                                         Print #1, ""
                                         FrmFrente.TxtTmp2 = ""
                                         FrmFrente.TxtSenha = ""
                                         If MebComprove <> "" Then
                                             Print #1, ""
                                             Print #1, "Saldo a Pagar ==> "; Tab(19); Alinhar(CStr(MebComprove.Text), 10, 2)
                                         End If
                                         
                                         Print #1, ""
                                         Print #1, Chr(29) & Chr(102) & Chr(0) & vbCrLf
    '                                     Print #1, Chr(29) & Chr(102) & Chr(0) & vbCrLf
                                    Next
                                     
                                End If
                                
                                If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Then
                                
                                     Print #1, "" & Cond
                                     Print #1, ""
                                     Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
                                     Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 48)
                                     Print #1, "================================================"
                                     Print #1, "Data:" & Date & " Hora:" & Time; Tab(33); "Pedido.:"; psCompString(1, LblTot, 6)
                                     If FrmFrente.TxtTmp2 <> "" Then
                                        Print #1, ""
                                        Print #1, "MESA ===>> " & FrmFrente.TxtTmp2
                                        Print #1, ""
                                    End If
                                     Print #1, Tab(1); psCompString(1, FrmFrente.Clientes("CodEntidade"), 6); "-"; Left(FrmFrente.Clientes.Fields("Nome"), 40)
                                     If (FrmFrente.Clientes.Fields("Endereco")) <> "" Then
                                        Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Endereco"), 35)) & ", " & FrmFrente.Clientes("Nro")
                                     End If
                                     If FrmFrente.Clientes("Bairro") <> "" Then
                                        Print #1, Tab(1); FrmFrente.Clientes("Bairro")
                                     End If
                                     If FrmFrente.Clientes("CIDADE") <> "" Then
                                        Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Cidade"), 30)) & " - " & FrmFrente.Clientes("Uf")
                                     End If
                                     If VarLinhaDet = 1 Then
                                           Print #1, Tab(1); "QTD"; Tab(9); "DESCRICAO DO PRODUTO"; Tab(33); "UNIT."; Tab(44); "TOTAL"
                                     Else
                                           Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
                                           Print #1, Tab(1); "QTD"; Tab(20); "UNIT."; Tab(40); "TOTAL"
                                     End If
                                     Print #1, "------------------------------------------------"
                                     
                                     FrmFrente.DatCompra.Recordset.MoveFirst
    
                                     While Not FrmFrente.DatCompra.Recordset.EOF
                                         If VarLinhaDet = 1 Then
                                                Print #1, Tab(1); Format(FrmFrente.DatCompra.Recordset("Qtd"), " #######"); Tab(9); Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 22); Tab(33); Alinhar(Format(FrmFrente.DatCompra.Recordset("ValorUnit"), "##,##0.00"), 8, 2); Tab(41); Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 8, 2)
                                         Else
                                                If Len(Trim(FrmFrente.DatCompra.Recordset("Complemento"))) > 0 Then
                                                   Print #1, Tab(1); psCompString(1, FrmFrente.DatCompra.Recordset("Produto"), 6); "-"; Left(FrmFrente.DatCompra.Recordset("Descrição_Produto") & "-" & Trim(FrmFrente.DatCompra.Recordset("Complemento")), 39)
                                                Else
                                                    
                                                    Criterio = "CodPrd = " & FrmFrente.DatCompra.Recordset("Produto")
                                                    FrmFrente.Registros.FindFirst Criterio
'                                                    If Not IsNumeric(FrmFrente.DatCompra.Recordset("Complemento")) Then
'                                                           Print #1, Tab(1); psCompString(1, FrmFrente.DatCompra.Recordset("Produto"), 6); "-"; Left(Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 27) & " " & Trim(FrmFrente.Registros("BarBarra")), 41)
'                                                       Else
                                                            Print #1, Tab(1); psCompString(1, FrmFrente.DatCompra.Recordset("Produto"), 6); "-"; Left(Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 40), 41)
'                                                       End If
                                                End If
                                                      Print #1, Tab(3); Alinhar(Format(FrmFrente.DatCompra.Recordset("Qtd"), " ##,##0.000"), 7, 3); Tab(15); "X"; Tab(18); Alinhar(Format(FrmFrente.DatCompra.Recordset("ValorUnit"), "##,##0.00"), 9, 2); Tab(37); Alinhar(Format(FrmFrente.DatCompra.Recordset("Valor"), "##,##0.00"), 12, 2)
                                         End If
                                         FrmFrente.DatCompra.Recordset.MoveNext
                                     Wend
                                     
                                     Print #1, ""
                                     
                                     If ChkComissao.Value = 0 Then
                                        If (MebDesconto.Text + VarDescIte) <> 0 Then
                                            Print #1, Tab(1); FrmFrente.DatCompra.Recordset.RecordCount; Tab(22); "SUB-TOTAL =>"; Tab(39); Alinhar(CStr((MebValor.Text + CCur(VarDescIte))), 10, 2)
                                            Print #1, Tab(22); "DESCONTO  =>"; Tab(39); Alinhar(CStr(MebDesconto.Text + VarDescIte), 10, 2)
                                            Print #1, Tab(1); Tab(22); "TOTAL     =>"; Tab(39); Alinhar(CStr(MebTotal.Text), 10, 2)
                                        Else
                                            Print #1, Tab(1); FrmFrente.DatCompra.Recordset.RecordCount; Tab(22); "TOTAL     =>"; Tab(39); Alinhar(CStr(MebTotal.Text), 10, 2)
                                        End If
                                     Else
'                                            Comissao = (CCur(MebValor.Text) * (VarComissao / 100))
                                            Print #1, Tab(1); FrmFrente.DatCompra.Recordset.RecordCount; Tab(22); "SUB-TOTAL =>"; Tab(39); Alinhar(CStr((MebValor.Text)), 10, 2)
                                            Print #1, Tab(22); "     10%  =>"; Tab(39); Alinhar(CStr(MebComissao), 10, 2)
                                            Print #1, Tab(22); "TOTAL     =>"; Tab(39); Alinhar(CStr(MebTotal.Text), 10, 2)
                                     End If
                                     If txtpgto.ListIndex <> 5 Then
                                         If txtpgto.ListIndex = 6 Then
                                             Print #1, Tab(1); "ORCAMENTO - " & FrmFrente.DbcVendedor
                                         ElseIf txtpgto.ListIndex = 7 Then
                                                 Print #1, Tab(1); "PARCELADO - " & FrmFrente.DbcVendedor
                                             ElseIf txtpgto.ListIndex = 8 Then
                                                 Print #1, Tab(1); "DEVOLUCAO - " & FrmFrente.DbcVendedor
                                             ElseIf txtpgto.ListIndex = 9 Then
                                                 Print #1, Tab(1); "CONSIGNADO - " & FrmFrente.DbcVendedor
                                                Else
                                                 If MebValorDig <> MebTotal Then
                                                    Print #1, Tab(22); "RECEBIDO  =>"; Tab(39); Alinhar(CStr(MebValorDig), 10, 2)
                                                    If (MebTroco) <> 0 Then
                                                       Print #1, Tab(22); "TROCO  =>"; Tab(39); Alinhar(CStr(MebTroco), 10, 2)
                                                    End If
                                                 End If
                                                 Print #1, ""
                                                 Print #1, Tab(1); "A VISTA - " & FrmFrente.DbcVendedor
                                         End If
                                     ElseIf txtpgto.ListIndex = 5 Then
                                         Print #1, Tab(1); "A PRAZO - " & FrmFrente.DbcVendedor
                                     End If
                                     If Not IsNull(TxtObs) Or TxtObs <> "" Then
                                         Print #1, Tab(1); Left(TxtObs, 35)
                                     End If
                                     If Not IsNull(TxtObs2) Or TxtObs <> "" Then
                                         Print #1, Tab(1); Left(TxtObs2, 35)
                                     End If
                                     
                                     Print #1, ""
                                     Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 48)
                                     Print #1, "" & Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
                                     Print #1, ""
                                     Print #1, ""
                                     FrmFrente.TxtTmp2 = ""
                                
                                
                                End If
                                                                                                
                                If VarImpressoraNf = "EPSON" Then
                                     If VarMeiaFolha = "Sim" Then
                                     
                                     
                                         For x = 1 To 2
                                             Print #1, Nor & Negi & Left(FrmFrente.Empresa("Cabecalho1"), 40); Tab(65); "Pedido: "; psCompString(1, LblTot, 6)
                                             Print #1, Left(FrmFrente.Empresa("Cabecalho2"), 80); Tab(70); Date ' & " - " & Time
 '                                            Print #1, ""
                                             Print #1, Nor & Negi & psCompString(1, FrmFrente.Clientes("CodEntidade"), 6) & "-" & Left(FrmFrente.Clientes("Nome"), 40) & NegF & Nor
                                             If FrmFrente.Clientes("Endereço_Entrega") <> "" Then
                                                Print #1, FrmFrente.Clientes("Endereço_entrega") & " " & FrmFrente.Clientes("Nro_Entrega")
                                                Print #1, FrmFrente.Clientes("Cidade_Entrega") & "  " & FrmFrente.Clientes("Uf_Entrega") & "  " & FrmFrente.Clientes("Bairro_Entrega") & "    " & Format(FrmFrente.Clientes("Cep_Entrega"), "@@.@@@-@@")
                                             Else
                                                Print #1, FrmFrente.Clientes("Endereco") & " " & FrmFrente.Clientes("Nro")
                                                Print #1, FrmFrente.Clientes("Cidade") & "  " & FrmFrente.Clientes("Uf") & "  " & FrmFrente.Clientes("Bairro") & "    " & Format(FrmFrente.Clientes("Cep"), "@@.@@@-@@")
                                             End If
                                             Print #1, "Fone:"; Tab(7); Format(FrmFrente.Clientes("Fone"), "(@@)@@@@-@@@@");
                                             If FrmFrente.Clientes("Classificacao") = 2 Then
                                                    Doc1 = Format(FrmFrente.Clientes("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                                                    Doc2 = FrmFrente.Clientes("InscrEst")
                                                    Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                                                    Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
                                                Else
                                                    Doc1 = Format(FrmFrente.Clientes("Rg"), "@@.@@@.@@@-@")
                                                    Doc2 = Format(FrmFrente.Clientes("Cpf"), "@@@.@@@.@@@.@@")
                                                    Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                                                    Print #1, Tab(52); "CPF:"; Tab(62); Doc2
                                             End If
'                                             Print #1, ""
                                             Print #1, "--------------------------------------------------------------------------------"
                                             Print #1, "CODIGO"; Tab(10); "QTD"; Tab(17); "UN"; Tab(25); "PRODUTO"; Tab(68); "UNIT."; Tab(76); "TOTAL"
                                             Print #1, "--------------------------------------------------------------------------------"
                                             Linhas = 9
                                             FrmFrente.DatCompra.Recordset.MoveFirst
                                             Do While Not FrmFrente.DatCompra.Recordset.EOF
                                                Print #1, Tab(1); psCompString(1, Trim(FrmFrente.DatCompra.Recordset("Produto")), 6); Tab(10); Alinhar(FrmFrente.DatCompra.Recordset("Qtd"), 6, 0); Tab(17); FrmFrente.DatCompra.Recordset("Embalagem"); Tab(20); Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 40); ' " "; Left(FrmFrente.DatCompra.Recordset("Complemento"), 20);
                                                Print #1, Tab(64); Alinhar(FrmFrente.DatCompra.Recordset("ValorUnit"), 8, 2); Tab(72); Alinhar(FrmFrente.DatCompra.Recordset("Valor"), 9, 2)
'                                                Print #1, ""
                                                Linhas = Linhas + 1
                                                varSubTotal = varSubTotal + FrmFrente.DatCompra.Recordset("Valor")
                                                If Linhas >= 30 Then
                                                        Print #1, ""
                                                        Print #1, ""
                                                        Print #1, ""
                                                        Print #1, ""
                                                        Print #1, ""
                                                        Linhas = 2
                                                End If
                                                FrmFrente.DatCompra.Recordset.MoveNext
                                             Loop
                                        
                                             VarTotal = varSubTotal
                                             If Linhas > 29 Then
'                                                SaltoPagina Linhas
                                             End If
                                             
'                                             Print #1, ""
                                             Print #1, Tab(55); " SUB-TOTAL =>   "; Alinhar(CStr(MebValor.Text), 10, 2)
                                             Print #1, Tab(57); "DESCONTO =>   "; Alinhar(CStr(MebDesconto.Text), 10, 2)
                                             Print #1, Tab(60); "TOTAL =>   "; Nor & Negi & Alinhar(CStr(MebTotal.Text), 10, 2) & NegF & Nor
                                             Linhas = Linhas + 3
                                                                              
                                             If Linhas > 28 Then
'                                                SaltoPagina Linhas
                                             End If
        
                                         
                                             If txtpgto.ListIndex <> 5 Then
                                                 If txtpgto.ListIndex = 6 Then
                                                     Print #1, Tab(1); "ORCAMENTO - " & FrmFrente.DbcVendedor
                                                     Linhas = Linhas + 1
                                                 ElseIf txtpgto.ListIndex = 7 Then
                                                         Print #1, Tab(1); "PARCELADO - " & FrmFrente.DbcVendedor
                                                         Linhas = Linhas + 1
                                                     ElseIf txtpgto.ListIndex = 8 Then
                                                         Print #1, Tab(1); "DEVOLUCAO - " & FrmFrente.DbcVendedor
                                                         Linhas = Linhas + 1
                                                     Else
                                                         Print #1, Tab(1); "A VISTA - " & psCompString(3, FrmFrente.DbcVendedor, 30);
                                                         Print #1, Tab(56); "RECEBIDO ==>   "; Alinhar(CStr(MebValorDig), 10, 2)
                                                         Print #1, Tab(59); "TROCO ==>   "; Alinhar(CStr(MebValorDig - MebTotal), 10, 2)
'                                                         Print #1, ""
                                                         Linhas = Linhas + 2
                                                 End If
                                              ElseIf txtpgto.ListIndex = 5 Then
                                                 Print #1, Tab(1); "A PRAZO - " & FrmFrente.DbcVendedor
'                                                 Print #1, ""
                                                 Print #1, Tab(1); "DATA VENCIMENTO - " '; VarDtVenc
                                                 Linhas = Linhas + 2
                                              End If
        '                                      Linhas = Linhas + 6
                                              If Not IsNull(TxtObs) Then
                                                 If TxtObs <> "" Then
                                                    Print #1, Tab(1); Left(TxtObs, 35)
                                                    Linhas = Linhas + 1
                                                 End If
                                              End If
                                              If Not IsNull(TxtObs2) Then
                                                 If TxtObs <> "" Then
                                                    Print #1, Tab(1); Left(TxtObs2, 35)
                                                    Linhas = Linhas + 1
                                                 End If
                                              End If
                                              If txtpgto.ListIndex = 5 Then
                                                    If Linhas > 27 Then
'                                                       SaltoPagina Linhas
                                                    End If
'                                                    Print #1, ""
'                                                    Print #1, ""
                                                    Print #1, Tab(30); " Ass.Cli-----------------------------------------"
                                                    Linhas = Linhas + 1
                                              End If
                                              If Linhas <= 30 Then
                                                SaltoPagina Linhas
                                              End If
'                                              x = 3
                                         Next
                                     Else
                                         Print #1, ""
                                         Print #1, "==============================================================================="
                                         Print #1, Gde & Centralizar(Left(FrmFrente.Empresa("Cabecalho1"), 40), 40) & Nor
                                         Print #1, Centralizar(Left(FrmFrente.Empresa("Cabecalho2"), 80), 80)
                                         Print #1, "==============================================================================="
                                         Print #1, "Data:"; Date; Tab(17); " - "; Time; Tab(66); "Pedido: "; psCompString(1, LblTot, 6)
                                         Print #1, "-------------------------------------------------------------------------------"
                                         Print #1, psCompString(1, FrmFrente.Clientes("CodEntidade"), 6) & "-" & Left(FrmFrente.Clientes("Nome"), 40)
                                         Print #1, FrmFrente.Clientes("Endereco") & " " & FrmFrente.Clientes("Nro")
                                         Print #1, FrmFrente.Clientes("Cidade") & "  " & FrmFrente.Clientes("Uf") & "  " & FrmFrente.Clientes("Bairro") & "    " & Format(FrmFrente.Clientes("Cep"), "@@.@@@-@@@")
                                         Print #1, "Fone:"; Tab(7); Format(FrmFrente.Clientes("Fone"), "(@@)@@@@-@@@@");
                                         If FrmFrente.Clientes("Classificacao") = 2 Then
                                                Doc1 = Format(FrmFrente.Clientes("Cgc"), "@@@.@@@.@@@/@@@@-@@")
                                                Doc2 = FrmFrente.Clientes("InscrEst")
                                                Print #1, Tab(24); "CNPJ:"; Tab(30); Doc1;
                                                Print #1, Tab(52); "InscrEst:"; Tab(62); Doc2
                                            Else
                                                Doc1 = Format(FrmFrente.Clientes("Rg"), "@@.@@@.@@@-@")
                                                Doc2 = Format(FrmFrente.Clientes("Cpf"), "@@@.@@@.@@@.@@")
                                                Print #1, Tab(24); "RG:"; Tab(30); Doc1;
                                                Print #1, Tab(52); "CPF:"; Tab(62); Doc2
                                         End If
                                         Print #1, "-------------------------------------------------------------------------------"
                                         Print #1, "CODIGO"; Tab(8); "PRODUTO"; Tab(44); "  "; Tab(53); "QTD"; Tab(61); "UNIT."; Tab(71); "TOTAL"
                                         Print #1, "-------------------------------------------------------------------------------"
                                         
                                         FrmFrente.DatCompra.Recordset.MoveFirst
    '                                  VarNroNota = FrmFrente.DatCompra.Recordset("NroNt")
                                         Do While Not FrmFrente.DatCompra.Recordset.EOF
                                            Print #1, Tab(1); psCompString(1, Trim(FrmFrente.DatCompra.Recordset("Produto")), 6); Tab(8); Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 38); Tab(48); Alinhar(FrmFrente.DatCompra.Recordset("Qtd"), 8, 2); Tab(58); Alinhar(FrmFrente.DatCompra.Recordset("ValorUnit"), 9, 2); Tab(67); Alinhar(FrmFrente.DatCompra.Recordset("Valor"), 9, 2)
                                            Linhas = Linhas + 1
                                            varSubTotal = varSubTotal + FrmFrente.DatCompra.Recordset("Valor")
                                            
                                            FrmFrente.DatCompra.Recordset.MoveNext
                                         Loop
                                    
                                         VarTotal = varSubTotal
                                         Print #1, ""
                                         Print #1, Tab(50); "SUB-TOTAL =>"; Tab(66); Alinhar(CStr(MebValor.Text), 10, 2)
                                         Print #1, Tab(50); "DESCONTO  =>"; Tab(66); Alinhar(CStr(MebDesconto.Text), 10, 2)
                                         Print #1, Tab(50); "TOTAL     =>"; Tab(66); Alinhar(CStr(MebTotal.Text), 10, 2)
                                         Print #1, ""
                                                                          
                                         If txtpgto.ListIndex <> 5 Then
                                             If txtpgto.ListIndex = 6 Then
                                                 Print #1, Tab(1); "ORCAMENTO - " & FrmFrente.DbcVendedor
                                             ElseIf txtpgto.ListIndex = 7 Then
                                                     Print #1, Tab(1); "PARCELADO - " & FrmFrente.DbcVendedor
                                                 ElseIf txtpgto.ListIndex = 8 Then
                                                     Print #1, Tab(1); "DEVOLUCAO - " & FrmFrente.DbcVendedor
                                                 Else
                                                     Print #1, Tab(50); "RECEBIDO ==>"; Tab(66); Alinhar(CStr(MebValorDig), 10, 2)
                                                     Print #1, Tab(50); "TROCO ==>"; Tab(66); Alinhar(CStr(MebValorDig - MebTotal), 10, 2)
                                                     Print #1, ""
                                                     Print #1, Tab(1); "A VISTA - " & FrmFrente.DbcVendedor
                                             
                                             End If
                                          ElseIf txtpgto.ListIndex = 5 Then
                                             Print #1, Tab(1); "A PRAZO - " & FrmFrente.DbcVendedor
                                          End If
                                          If Not IsNull(TxtObs) Or TxtObs <> "" Then
                                             Print #1, Tab(1); Left(TxtObs, 35)
                                          End If
                                          If Not IsNull(TxtObs2) Or TxtObs <> "" Then
                                             Print #1, Tab(1); Left(TxtObs2, 35)
                                          End If
                                          Print #1, ""
                                          Print #1, " Recebi os produtos acima - Ass.Cli ----------------------------------"
                                          Print #1, ""
                                          Print #1, ""
                                          Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 80)
                                          Print #1, ""
                                          Print #1, ""
                                          If txtpgto.ListIndex = 5 Then
                                                
                                                Print #1, "==============================================================================="
                                                Print #1, Gde & VarEmpresa & Nor
                                                Print #1, VarEnd; Tab(40); VarCidade
                                                Print #1, "Insc.Est.: " & VarInscEst; Tab(30); "CNPJ: " & VarCGC
                                                Print #1, "==============================================================================="
                                                Print #1, "         FATURA           |       DUPLICATA         |         DATA             "
                                                Print #1, "   Valor     |   Numero   |   Valor     |  Numero   | Emissão     | Vencimento "
                                                Print #1, "-------------------------------------------------------------------------------"
    '                                            If VarNroNota = 0 Then
                                                    Print #1, "             |            |"; Tab(30); Alinhar(CStr(MebTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, LblTot, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
    '                                            Else
    '                                                Print #1, Tab(2); Alinhar(CStr(MebTotal), 10, 2); Tab(12); "|"; Tab(15); Alinhar(CStr(MebTotal), 10, 2); Tab(27); "|"; Tab(30); Alinhar(CStr(MebTotal), 10, 2); Tab(41); "|"; Tab(45); psCompString(1, LblTot, 6); Tab(53); "|"; Tab(55); Date; Tab(67); "|"; Tab(69); VarDtVenc
    '                                            End If
                                                Print #1, "-------------------------------------------------------------------------------"
                                                Print #1, " Nome do Sacado.: "; FrmFrente.Clientes("Nome")
                                                Print #1, " Endereco.......: "; Trim(FrmFrente.Clientes("Endereco")) & ", " & FrmFrente.Clientes("Nro")
                                                Print #1, " Municipio......: "; Left(FrmFrente.Clientes("Cidade"), 20) & " - " & Left(FrmFrente.Clientes("Uf"), 2) & "  " & Format(FrmFrente.Clientes("Cep"), "@@.@@@-@@")
                                                Print #1, " CNPJ/CPF.......: "; Doc1; Tab(45); "Insc.Est./RG.: "; Doc2
                                                Print #1, "-------------------------------------------------------------------------------"
                                                Print #1, " Valor por Extenso: "; Cond & UCase(PassaExtenso(MebTotal)) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *" & Nor
                                                Print #1, "-------------------------------------------------------------------------------"
                                                Print #1, " Reconheco/emos a exatidao desta DUPLICATA de VENDA MERCANTIL, na importancia  "
                                                Print #1, " acima que pagarei/emos a " & Negi & VarEmpresa & NegF
                                                Print #1, " ou a sua ordem na praca e vencimento acima indicados."
                                                Print #1, " "
                                                Print #1, "           ____/____/_______            _______________________________"
                                                Print #1, "            Data do aceite                  Assinatura do Sacado       "
                                          End If
                                          Print #1, SaltoPag
                                    
     '                                     Fim (False)
     '                                     Close #1
                                     End If
                                End If
                            End If
        
                            
                            If txtpgto.ListIndex = 5 And (PortaNF <> "TELA" And VarImpressoraNf <> "EPSON") Then
                                If VarComprovante = True Then
                                    If MsgBox("Imprimir Comprovante ?", vbYesNo, App.Title) = vbYes Then
         '                               Iniciar (False)
                                
                                        If Left(PortaNF, 3) <> "COM" Then
                                            Print #1, ""
                                            Print #1, ""
                                            Print #1, "" & Chr(27) & "i"        '& Chr(27) & "v50"
                                            Fim (False)
                                            Iniciar (False)
                                            Print #1, "======================================="
                                            Print #1, " Pedido:"; Tab(10); psCompString(1, LblTot, 6); Tab(18); "Data: "; Tab(24); Date
                                            Print #1, ""
                                            Print #1, Tab(1); psCompString(1, FrmFrente.Clientes.Fields("CodEntidade"), 6); Tab(8); "-"; Tab(10); Left(FrmFrente.Clientes.Fields("Nome"), 30)
                                            Print #1, ""
                                            If MebComprove <> "" Then
                                                Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebComprove.Text), 10, 2)
                                            Else
                                                Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebTotal.Text), 10, 2)
                                            End If
                                            Print #1, ""
                                            Print #1, ""
                                            Print #1, "  Ass.Cli----------------------------"
                                            Print #1, "======================================="
                                            Print #1, ""
                                        End If
                                        
        '                                Fim (False)
                                    Else
                                        If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Or VarImpressoraNf = "TERMICA" Then
        '                                    Iniciar (False)
                                                Print #1, ""
                                                Print #1, ""
                                                Print #1, "   Ass.Cli---------------------------"
                                                Print #1, "======================================="
                                                Print #1, "" + Nor
        '                                    Fim (False)
                                        End If
                                    End If
                                Else
                                    If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Or VarImpressoraNf = "TERMICA" Then
    '                                    Iniciar (False)
                                            
                                            Print #1, ""
                                            Print #1, ""
                                            Print #1, ""
                                            Print #1, "   Ass.Cli---------------------------"
                                            Print #1, "======================================="
                                            Print #1, "" + Nor
    '                                    Fim (False)
                                    End If
                                
                                End If
                            End If
     '                       Iniciar (False)
                            If Left(PortaNF, 3) <> "COM" Then
                                If VarMeiaFolha = "Nao" And VarImpressoraNf <> "EPSON" Then
                                
                                    Print #1, "" & Nor
                                    For x = 1 To NroLinha
                                        Print #1, " "
                                    Next
        '                            Print #1, ""
        '                            Print #1, ""
                                End If
                            End If
                            Print #1, Chr(27) & "i"         '& Chr(27) & "v50"
                                                            
                            Fim (False)
                    End If
                End Select
                
                If IsNull(FrmFrente.Empresa("MensagemRodape")) Then
                    MensagemPromo = "           Agradecemos a Preferencia            "
                Else
                    MensagemPromo = psCompString(3, Centralizar(FrmFrente.Empresa("MensagemRodape"), 48), 48)
                End If
                TxtImposto = "Trib Aprox R$ " & Format(VarTotImpostoNac, "#,##0.00") & " Federal e R$ " & Format(VarImpostoEst, "##0.00") & " Estadual Fonte:IBPT"
                TxtImposto = psCompString(3, TxtImposto, 48)
                
                
                If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") <> "" Then
                    If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") <> 3 Then
                        MensagemPromo = "ICMS RECOLHIDO CONF.LC 123/2006-SIMPLES NACIONAL"
                    End If
                End If
                                
                
                If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then

                    Set VarPopular = FrmFrente.Banco.OpenRecordset("Select * From Popular Where Pedido = " & LblTot)
                    VarPopular.Requery
                    If VarPopular.RecordCount > 0 Then
                        If VarPopular!PreAutoriza <> "" Then
                            VarAutoriza = VarPopular!PreAutoriza
                            Dim VarPopular2 As Recordset
                            Set VarPopular2 = FrmFrente.Banco.OpenRecordset(" SELECT POPULAR.Pedido, Sum(POPULAR.ParcCli) AS SomaDeParcCli, Sum(POPULAR.ParcGov) AS SomaDeParcGov, First(Vendedor) as PVendedor, First(Senha) as RSenha FROM POPULAR GROUP BY POPULAR.Pedido HAVING POPULAR.Pedido= " & LblTot)
                            VarPopular2.Requery
                            Close #8
                            
                            If LerINI("Sistema", "Login", App.Path & "\Popular.ini") <> "" Then
                                VarUsuario = LerINI("sistema", "Login", App.Path & "\Popular.ini")
                            End If
                            
                            If LerINI("Sistema", "Senha", App.Path & "\Popular.ini") <> "" Then
                                VarSenha = LerINI("sistema", "Senha", App.Path & "\Popular.ini")
                            End If
                            
                            If LerINI("Sistema", "Chave", App.Path & "\Popular.ini") <> "" Then
                                VarChavePop = LerINI("sistema", "Chave", App.Path & "\Popular.ini")
                            End If
                            
                                            
                            If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
                                Kill "C:\FPOPULAR\FP_REQ.txt"
                            End If
                            If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
                                Kill "C:\FPOPULAR\FP_ENV.txt"
                            End If
                                
                            Open "c:\FPopular\FP_ENV.TXT" For Output As #8
                            Print #8, "01" & Trim(VarUsuario) & Trim(VarPopular2!pVendedor) & Trim(VarChavePop)
                            Print #8, "02" & Trim(VarSenha) & Trim(VarPopular2!rSenha)
                            If Impressora = 8 Then
                                Print #8, "05" & psCompString(1, LblTot, 6) & VarAutoriza & psCompString(1, VarNroCupom, 6)
                            Else
                                Print #8, "05" & psCompString(1, LblTot, 6) & VarAutoriza & psCompString(1, NroCupom, 6)
                            End If
                            Close #8
                                
                            TrataPopular2
                            
                            VarCnpj = LerINI("sistema", "cnpj", App.Path & "\Popular.ini")
                            VarCnpj = Format(VarCnpj, "@@.@@@.@@@/@@@@-@@")
                            On Error Resume Next
                            Dim VarTemp As String
                            
                            Select Case Impressora
                            
                                Case 1
                                Case 5
                                Case 8
                                    If MsgBox("Nfc-e Impresso com sucesso ? Imprimir comprovante da popular ?", vbYesNo) = vbYes Then
                                        Iniciar (False)
                                        
                                        Open "C:\FPOPULAR\FP_REQ.txt" For Input As #6
                                        
                                        If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Then
                                             Print #1, "" & Cond
                                        End If
                                        
                                        Print #1, "------------------------------------------"
                                        Print #1, "Data: " & Date & " Hora: " & Time()
                                        Print #1, psCompString(3, "Documento vinculado Sat Nro:  " & psCompString(1, VarNroCupom, 6), 46)
                                        Print #1, psCompString(3, "Valor da Compra R$ " & MebTotal, 46)
                                        Print #1, psCompString(3, "Valor Pago      R$ " & MebTotal, 46)
                                        Do
                                           Line Input #6, VarTemp
                                               If Left(VarTemp, 2) = 7 Then
                                                   Print #1, psCompString(3, Mid(VarTemp, 3, 45), 45)
                                            End If
                                        Loop Until EOF(6)
                                        Close #6
                                        For x = 1 To NroLinha
                                            Print #1, " "
                                        Next
                                        Print #1, ""
                                        Print #1, "" & Chr(27) & "i"        '& Chr(27) & "v50"
                                        Fim (False)
                                    End If
                            End Select
                            
                        End If
                    End If
                
                    If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
                        Kill "C:\FPOPULAR\FP_REQ.txt"
                    End If
                    If dir("C:\FPOPULAR\FP_ENV.txt") <> "" Then
                        Kill "C:\FPOPULAR\FP_ENV.txt"
                    End If
                
                
               End If
                
            End If
            
            If txtpgto.ListIndex = 7 And VarColombo = False Then
                If MsgBox("Gerar Parcelas Agora ?", vbYesNo, App.Title) = vbYes Then
                    FrmParcelas.CodDoc = LblTot
                    FrmParcelas.MebVlrPed = MebValor
                    Me.Hide
                    FrmParcelas.Show vbModal
                    Me.Show
                End If
            End If
            
            
            If Impressora = 99 Then
                VarImprEntrega = LerINI("Impressora", "ImprEntrega", LocalBanco & "\config.ini")
                If VarImprEntrega <> "" Then
                    If txtpgto.ListIndex <> 8 And txtpgto.ListIndex <> 6 Then
                        Me.Hide
                        FormEntrega.Show vbModal
                        Me.Show
                    End If
                End If
            End If
            
            If MebTroco = "" Then MebTroco = 0
            Troco = MebTroco
            MebValor = ""
            MebTotal = ""
            MebTroco = ""
            MebDesc = ""
            MebDesconto = ""
            MebValorDig = ""
'            txtpgto = ""
            MebComprove = ""
            FrmFrente.LimpaTela
            If VarConveniencia = True Or VarLanchonete = True Then
                sql = "UPDATE Conv SET  Totaliza = -1   WHERE Cliente = " & FrmFrente.CodCard
                FrmFrente.Banco.Execute sql
                FrmFrente.AtuMesa
            End If

            Unload Me
            FrmFrente.LblDescricao = "TROCO"
            FrmFrente.MebTotal = Format(Troco, "#,##0.00")
            FrmFrente.TxtTmp2 = ""
            If VarContMesa <> "" Then
'                FrmFrente.Mesa(VarContMesa).Caption = ""
                FrmFrente.Mesa(VarContMesa).BackColor = &HFFFFFF
            End If
            
        ElseIf Resp = vbNo Then
                MebTroco = ""
                MebValorDig = ""
'                txtpgto = ""
                MebDesc = 0
                MebDesconto = 0
                MebComprove = ""

                MebValorDig.SetFocus
        Else
                Unload Me
        End If
        
    Else
        MsgBox "Valor não numérico ou fora da faixa", vbInformation
'        txtpgto = ""
        txtpgto.SetFocus
    End If
    Flag = True
    
    If VarCodDev <> 0 Then
          If MsgBox("Deseja relacionar essa devolução a algum pedido ?", vbYesNo, App.Title) = vbYes Then
'                 Unload Me
                FormDevolucao.Show vbModal
          End If
    End If
    
End If

End Sub
Private Sub MebValorDig_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If

If KeyAscii = 13 Then
    MebComprove = ""
    If IsNumeric(MebValorDig.Text) Or MebValorDig.Text = "" Then
        If ChkComissao.Value = 0 Then
            If MebValorDig.Text = "" Then
                MebValorDig = MebTotal
            End If
            If CCur(MebValorDig.Text) < CCur(MebTotal) And MebValorDig <> 0 Then
                If VarMercado = False Then
'                    MebDesconto = MebValor - MebValorDig
                    MebTotal = MebValor - MebDesconto
                    MebDesc = 0
                    MebTroco = 0
                Else
                    MebComprove = MebValor - MebValorDig
                    MebDesc = 0
                    MebTroco = 0
                End If
                
            Else
                If MebValorDig = 0 Then
'                    MebValorDig.Text = MebTotal
                    MebTroco = 0
                    MebDesconto = MebTotal - TotPgto
                    MebTotal = MebTotal - MebDesconto
                    MebValorDig = 0
                    MebDoc.SetFocus
                Else
                    MebTroco.Text = Format(CDec((MebValorDig)) - MebTotal, "  ##,##0.00")
                End If
            End If
'            If MebValorDig = 0 Then
'                MebDesconto = MebDesconto + CCur(MebValorDig)
'                MebTotal = MebTotal - MebDesconto
'                MebValorDig = 0
'                MebDoc.SetFocus
'            End If
         Else
            If CDec((MebValorDig)) >= MebTotal Then
                MebTroco.Text = Format(CDec((MebValorDig)) - MebTotal, "  ##,##0.00")
            Else
'                MebValorDig = ""
'                MebValorDig.SetFocus
            End If
            
         End If
    Else
        MsgBox "Valor Inválido", vbInformation
        MebValorDig = ""
        MebValorDig.SetFocus
    End If
End If
'MebComprove = ""
End Sub
Private Sub TxtPgto_GotFocus()
'If VarConveniencia = True Or VarLanchonete = True Then
'     Vista = True
'End If

'If Vista = True Then
If FrmFrente.CodCli = 1 Then
    txtpgto.ListIndex = 0
Else
    txtpgto.ListIndex = 5
End If

'If VarBloqueado = True Then
'    txtpgto.ListIndex = 5
'End If

'txtpgto.SelStart = 0
'txtpgto.SelLength = 3
End Sub


Function TrataPopular2() As String
Dim VarPopular(1 To 6) As String
Dim VarTemp As String
Dim x As Double
Dim Var

On Error Resume Next
DoEvents
x = 0

For x = 1 To 6
    VarPopular(x) = ""
Next

While dir("C:\FPOPULAR\FP_REQ.txt") = ""

Wend

x = 0
LblAutoriza = ""

MsgBox "Autorização confirmada !", vbInformation


For x = 1 To 700000


Next

Open "C:\FPOPULAR\FP_REQ.txt" For Input As #5
    Do
        Input #5, VarTemp
        Select Case Left(VarTemp, 2)
            Case "01"
                VarPopular(1) = VarTemp
            Case "02"
                VarPopular(2) = VarTemp
            Case "03"
                VarPopular(3) = VarTemp
            Case "04"
                VarPopular(4) = VarTemp
            Case "05"
                VarPopular(5) = VarTemp
            Case "06"
                VarPopular(6) = VarTemp
        End Select
    Loop Until EOF(5)

    MsgBox VarPopular(1) & VarPopular(2) & VarPopular(3) & VarPopular(4), vbInformation, App.Title

Close #5




'txtPopular = VarPopular(1) + Chr(13) + VarPopular(2) + Chr(13) + VarPopular(3) + Chr(13) + VarPopular(4) + Chr(13) + VarPopular(5) + Chr(13) + VarPopular(6)

'Kill "C:\FPOPULAR\FP_REQ.txt"
'Kill "C:\FPOPULAR\FP_ENV.txt"

'For X = 1 To 100000

'Next

End Function


Function TrataAmpla(VarArq As String) As String
'Dim VarPopular(1 To 6) As String
Dim VarTemp As String
Dim sql As String
Dim x As Double
Dim Var
Dim VarQtd As Currency
Dim ValorGov As Currency
Dim ValorCli As Currency
Dim CodBarPop As String
Dim Criterio As String

If VarConveniencia = True Or VarLanchonete = True Then
    FrmFrente.Timer2.Enabled = False
End If
'On Error Resume Next
DoEvents
x = 0

'For X = 1 To 6
'    VarPopular(X) = ""
'Next

While dir("C:\Winguest\Ret\" & VarLoja & VarCxa & VarArq) = ""
 '   DoEvents
Wend

x = 0
LblAutoriza = ""

For x = 1 To 700000

Next

Open "C:\Winguest\Ret\" & VarLoja & VarCxa & VarArq For Input As #5
Input #5, VarTemp
Close #5

If dir("C:\Winguest\Ret\" & VarLoja & VarCxa & ".STS") <> "" Then
    Kill "C:\Winguest\Ret\" & VarLoja & VarCxa & ".STS"
End If

If VarArq = ".STS" Then
    If Mid$(VarTemp, 5, 2) <> "01" Then
        MsgBox "Servidor Amplacar off-line"
        TrataAmpla = "00"
    Else
        TrataAmpla = VarTemp
    End If
End If

If VarConveniencia = True Or VarLanchonete = True Then
    FrmFrente.Timer2.Enabled = True
End If
End Function




Public Sub VoltaEstoque(VarPedido As String)

Dim VarQtd As Currency
Dim VarPrd As Long
Dim VarLoja As String
Dim VarBaixa As Boolean

If dir(App.Path & "\" & Trim(VarPedido) & ".txt") <> "" Then

        Open App.Path & "\" & Trim(VarPedido) & ".txt" For Input As #6
        Do
            Line Input #6, VarTemp
            VarQtd = Mid(VarTemp, 1, 10) / 100
            VarPrd = Mid(VarTemp, 11, 8)
            VarLoja = Mid(VarTemp, 19, 20)
            VarBaixa = Mid(VarTemp, 39, 1)
            BaixarEstoque VarQtd, VarPrd, VarLoja, VarBaixa
            
        Loop Until EOF(6)
        Close #6
        Kill App.Path & "\" & Trim(VarPedido) & ".txt"

End If

End Sub


Public Sub ImprimirSat()
Dim VarContPrd As Integer
Dim VarDescPrd As String
'Dim TxtImposto As String
Dim ConvertQtd As String
Dim ConvertVr As String
Dim VarChave As String
Dim VarQRCode As String

Dim VarTotal As Currency
Dim VarPago As Currency
Dim VarDescIni As Currency
Dim VarNroPed As String
Dim FormaPgto As String
Dim sql As String
Dim VarCpf As String
Dim VarNroSerie As String
Dim VarObs As String
Dim VarOper As String
VarNroCupom = ""
VarObs = ""
'On Error Resume Next
  If PortaF <> "TELA" Then
  
       Iniciar (False)
            Print #1, Chr(29) + Chr(249) + Chr(53) + Chr(0)
        Fim (False)
   End If
   Set VarDoc = New MSXML2.DOMDocument
   VarDoc.async = False
   VarDoc.load VarImprSat
   
  If VarDoc.parseError.reason <> "" Then   'Verifica se Foi aberto com Sucesso
'      MsgBox VarDoc.parseError.reason & vbCrLf & _
'             "Line: " & xml.parseError.Line & vbCrLf & _
'             "Pos: " & xml.parseError.linepos
            MsgBox "Erro no xml", vbInformation
        Exit Sub
   End If
    '   Set CfeProc = VarDoc.selectSingleNode("Cfe")
       
       Set Cfe = VarDoc.selectSingleNode("CFe")
       Set VarinfCfe = Cfe.selectSingleNode("infCFe")
       
       ' Pega Chave da NF-e
       
    '   If VarinfCfe.Attributes.getNamedItem("versao") Is Nothing Then
    '      TxtId.Text = "Não encontrado!"
       
    '   ElseIf VarinfCfe.Attributes.getNamedItem("versao").Text = "0.03" Then
            
        VarChave = Format(Mid(VarinfCfe.Attributes.getNamedItem("Id").Text, _
              4, 44), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    '   End If
        Set Ide = VarinfCfe.selectSingleNode("ide")
    '    Set infProt = Ide.selectSingleNode("CNPJ")
        
        VarQRCode = Ide.selectSingleNode("assinaturaQRCODE").Text
        VarNroCupom = Ide.selectSingleNode("nCFe").Text
        VarNroSerie = Ide.selectSingleNode("nserieSAT").Text
        
    
    VarNroPed = LerINI("Identificacao", "NCfe", App.Path & "\cfe.ini")
    VarObs = LerINI("DadosAdicionais", "InfCpl", App.Path & "\cfe.ini")
    VarOper = CInt(LerINI("Pagto001", "cMP", App.Path & "\cfe.ini"))
    Select Case CInt(VarOper)
        Case 1
            FormaPgto = "Dinheiro"
        Case 2
            FormaPgto = "Cheque"
        Case 3
            FormaPgto = "Pix/Transf."
        Case 4
            FormaPgto = "Cartão Débito"
        Case 5
            FormaPgto = "Cartão Crédito"
        Case 6
            FormaPgto = "A Prazo"
    End Select
            
'    sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.Sat = " & """" & CStr(Mid(Trim(VarImprSat), Len(VarImprSat) - 47, 44)) & """" & " , PED.Cfo = 'SAT', Ped.PrevEntrega = ' ', Ped.NroNt = " & VarNroCupom & ", Ped.Operacao = " & IIf(CInt(VarOper) = "6", 2, CInt(VarOper)) & "  WHERE PED.CodPed = " & VarNroPed
    sql = "UPDATE PED SET PED.DtFat = #" & Format(Date, "mm/dd/yyyy") & "#, PED.Sat = " & """" & CStr(Mid(Trim(VarImprSat), Len(VarImprSat) - 47, 44)) & """" & " , PED.Cfo = 'SAT', Ped.PrevEntrega = ' ', Ped.NroNt = " & VarNroCupom & "  WHERE PED.CodPed = " & VarNroPed
    FrmFrente.Banco.Execute sql
            
    If PortaF = "TELA" Then
        If MsgBox("Visualizar/Imprimir Cupom ?", vbYesNo, App.Title) = vbYes Then
            AbreAcbr "SAT.ImprimirExtratoVenda(" & VarImprSat & ")"
'            TrataAcbr
            If CInt(VarOper) = 6 And VarComprovante = True Then
                If MsgBox("Imprimir comprovante ?", vbYesNo, App.Title) = vbYes Then
    
                    Iniciar (False)
                    Print #1, "======================================="
                    Print #1, " Pedido:"; Tab(10); psCompString(1, VarNroPed, 6); Tab(18); "Data: "; Tab(24); Date
                    Print #1, ""
                    Print #1, Tab(1); LerINI("Destinatario", "XNome", App.Path & "\cfe.ini")
                    Print #1, ""
                    If MebComprove <> "" Then
                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebComprove), 10, 2)
                    ElseIf MebTotal <> "" Then
                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(MebTotal), 10, 2)
                     Else
                        Print #1, " Valor ==> "; Tab(12); Alinhar(CStr(FormConsultaPed.MebTotal), 10, 2)
                    End If
                    Print #1, ""
                    Print #1, ""
                    Print #1, "  Ass.Cli----------------------------"
                    Print #1, "======================================="
                    Print #1, ""
                    
                    Print #1, "" & Nor
                    For x = 1 To NroLinha
                        Print #1, " "
                    Next
                    Print #1, Chr(27) & "i"         '& Chr(27) & "v50"
                                    
                    Fim (False)
'                    Unload FrmTotal

                End If
            End If
        End If
    Else
        
        If MsgBox("Impressora pronta ?", vbYesNo, App.Title) = vbYes Then
        
            VarInscMuni = "000.000"
            
            Comando = ConfiguraModeloImpressora(7)
            
            Comando = IniciaPorta(PortaF)
            
    
    '        Comando = PrinterReset()
            VarCpf = LerINI("Destinatario", "CNPJCPF", App.Path & "\cfe.ini")
            
        '                    Comando = FormataTX("" + Chr(10), 2, 0, 0, 0, 0)
    '        Comando = FormataTX(Chr(29) + Chr(249) + Chr(53) + Chr(0), 1, 0, 0, 0, 0)
            Comando = FormataTX(Centralizar(VarFantasiaEmite, 25) + Chr(10), 2, 0, 0, 1, 0)
            Comando = FormataTX(Centralizar(VarRazaoEmite, 48) + Chr(10), 2, 0, 0, 0, 1)
            Comando = FormataTX(Centralizar(VarEndEmite & ", " & VarNroEmite, 48) + Chr(10), 2, 0, 0, 0, 0)
            Comando = FormataTX(Centralizar(VarCidadeEmite & "-" & VarUfEmite, 48) + Chr(10), 2, 0, 0, 0, 0)
            
            Comando = FormataTX("CNPJ " & Format(VarCnpjEmite, "@@.@@@.@@@/@@@@-@@") & " ", 1, 0, 0, 0, 0)
            Comando = FormataTX("IE " & Format(VarInscEmite, "@@@.@@@.@@@.@@@") & " ", 1, 0, 0, 0, 0)
            Comando = FormataTX("IM " & Format(VarInscMuni, "@@@.@@@.@@@.@@@") + Chr(10), 1, 0, 0, 0, 0)
            Comando = FormataTX("================================================" & Chr(10), 2, 0, 0, 0, 0)
            Comando = FormataTX(Centralizar("Extrato Nro: " & VarNroCupom, 48) + Chr(10), 2, 0, 0, 0, 1)
            Comando = FormataTX(Centralizar("CUPOM FISCAL ELETRONICO - SAT", 48) + Chr(10), 2, 0, 0, 0, 1)
            Comando = FormataTX("------------------------------------------------" + Chr(10), 2, 0, 0, 0, 0)
            If VarCpf <> "" Then
                Comando = FormataTX("CPF/CNPJ do Consumidor: " & VarCpf + Chr(10), 2, 0, 0, 0, 0)
                Comando = FormataTX("-----------------------------------------------" + Chr(10), 2, 0, 0, 0, 0)
            End If
            Comando = FormataTX("#|COD|DESC|QTD|UN|VL UN R$|(VL TRR$)|VL ITEM R$" + Chr(10), 2, 0, 0, 0, 0)
            Comando = FormataTX("================================================" & Chr(10), 2, 0, 0, 0, 0)
            VarContPrd = 1
            VarDescPrd = ""
            Dim varteste As String
            VarTotal = 0
            VarDescIni = 0
            
            While LerINI("Produto" & psCompString(1, VarContPrd, 3), "cProd", App.Path & "\cfe.ini") <> ""
                ConvertQtd = LerINI("Produto" & psCompString(1, VarContPrd, 3), "qCom", App.Path & "\cfe.ini")
                ConvertVr = LerINI("Produto" & psCompString(1, VarContPrd, 3), "vUnCom", App.Path & "\cfe.ini")
                VarDescPrd = psCompString(1, VarContPrd, 3) & Space(1) & psCompString(1, LerINI("Produto" & psCompString(1, VarContPrd, 3), "cProd", App.Path & "\cfe.ini"), 6) & Space(1) & _
                             psCompString(3, Left(LerINI("Produto" & psCompString(1, VarContPrd, 3), "xProd", App.Path & "\cfe.ini"), 18), 18) & Space(1) & _
                             Alinhar(ConvertQtd, 7, 3) & Space(1) & _
                             Left(LerINI("Produto" & psCompString(1, VarContPrd, 3), "uCom", App.Path & "\cfe.ini"), 2) & " X " & _
                             Alinhar(ConvertVr, 5, 2) & "(" & _
                             Alinhar(LerINI("Produto" & psCompString(1, VarContPrd, 3), "vItem23741", App.Path & "\cfe.ini"), 5, 2) & ")" & _
                             Alinhar(ConvertQtd * ConvertVr, 9, 2)
                Comando = FormataTX(VarDescPrd + Chr(10), 1, 0, 0, 0, 0)
                
                VarContPrd = VarContPrd + 1
                VarTotal = VarTotal + (ConvertQtd * ConvertVr)
            Wend
            Comando = FormataTX("" + Chr(10), 1, 0, 0, 0, 0)
            
            VarDescIni = LerINI("DescAcrEntr", "VDescSubTotal", App.Path & "\cfe.ini")
            VarPago = LerINI("Pagto001", "vMP", App.Path & "\cfe.ini")
            If VarDescIni <> 0 Then
                Comando = FormataTX(psCompString(3, "SubTotal", 18) & Space(20) & Alinhar(CStr(VarTotal), 10, 2) + Chr(10), 2, 0, 0, 0, 0)
                Comando = FormataTX(psCompString(3, "Desconto", 18) & Space(20) & Alinhar(CStr(VarDescIni), 10, 2) + Chr(10), 2, 0, 0, 0, 0)
            End If
            
            Comando = FormataTX("TOTAL R$      " & Alinhar(CStr(VarTotal - VarDescIni), 10, 2) + Chr(10), 3, 0, 0, 1, 0)
        '                    Comando = FormataTX("" + Chr(10), 1, 0, 0, 0, 0)
            Comando = FormataTX(psCompString(3, FormaPgto, 18) & Space(20) & Alinhar(CStr(VarPago), 10, 2) + Chr(10), 2, 0, 0, 0, 0)
            If VarPago > VarTotal <> 0 Then
                Comando = FormataTX("Troco             " & Space(20) & Alinhar(CStr(VarPago - (VarTotal - VarDescIni)), 10, 2) + Chr(10), 2, 0, 0, 0, 0)
            End If
            If VarObs <> "" Then
                Comando = FormataTX(VarObs + Chr(10), 2, 0, 0, 0, 0)
            End If
            
            
        '                    Comando = FormataTX("" + Chr(10), 1, 0, 0, 0, 0)
        '    TxtImposto = "Val Aprox dos Trib.R$ " & Format(VarTotImposto, "#,##0.00") & "(" & Format(((VarTotImposto / CCur(MebValor)) * 100), "##0.00") & "%) Fonte: IBPT"
            Comando = FormataTX(TxtImposto + Chr(10), 1, 0, 0, 0, 1)
            Comando = FormataTX("------------------------------------------------" + Chr(10), 2, 0, 0, 0, 0)
            Comando = FormataTX(Centralizar("SAT Nro: " & psCompString(1, VarNroSerie, 6) & "  " & Date & "   " & Time(), 48) + Chr(10), 2, 0, 0, 0, 1)
        '                    Comando = FormataTX(Centralizar(Date & "   " & Time(), 48) + Chr(10), 2, 0, 0, 0, 1)
            
                                    
            
            Comando = FormataTX(Centralizar(Format(VarChave, "@@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@"), 62) + Chr(10), 1, 0, 0, 0, 0)
            Comando = ConfiguraCodigoBarras(35, 1, 0, 0, 15)
            Comando = ImprimeCodigoBarrasCODE128(Left(VarChave, 22))
            Comando = ImprimeCodigoBarrasCODE128(Mid(VarChave, 23, 22))
        '                    Comando = FormataTX("" + Chr(10), 1, 0, 0, 0, 0)
        '                    Comando = FormataTX(Space(20), 1, 0, 0, 0, 0)
            Comando = ConfiguraCodigoBarras(50, 1, 2, 0, 120)
            Comando = ImprimeCodigoQRCODE(1, 5, 0, 10, 1, Trim(VarChave) & "|" & Format(Date, "yyyymmdd") & Format(Time(), "hhmmss") & "|" & Numero(CCur(VarTotal) - CCur(VarDescIni), 2) & "|" & Trim(VarCpf) & "|" & VarQRCode)
                                       
            Comando = AcionaGuilhotina(1)
            
            If LerINI("Pagto001", "cMP", App.Path & "\cfe.ini") = "06" Then
            
                If MsgBox("Imprimir comprovante ?", vbYesNo, App.Title) = vbYes Then
                
                     Comando = FormataTX("================================================" & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX("Pedido: " & psCompString(1, VarNroPed, 6) & "  Data: " & Date & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(" " & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(LerINI("Destinatario", "XNome", App.Path & "\cfe.ini") & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(" " & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(" Valor ==> " & Alinhar(CStr(VarTotal - VarDescIni), 10, 2) & Chr(10), 2, 0, 0, 0, 1)
                     Comando = FormataTX(" " & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(" " & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX("  Ass.Cli ---------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX("================================================" & Chr(10), 2, 0, 0, 0, 0)
                     Comando = FormataTX(" " & Chr(10), 2, 0, 0, 0, 0)
                                           
                    Comando = AcionaGuilhotina(1)
               End If
            End If
        
            fecha = FechaPorta()
        End If
    End If
    
End Sub


Private Sub txtpgto_KeyPress(KeyAscii As Integer)

If MebValorDig = "" Then
    MsgBox "Nenhum Valor Digitado", vbInformation
    Exit Sub
End If
    
If txtpgto.Text = "" Then
    MsgBox "Selecionar Tipo Pgto", vbInformation
    Exit Sub
End If

'If TotPgto >= CCur(MebTotal) Then
'    MsgBox "Total de Pgto Completo", vbInformation
'    Exit Sub
'End If

If KeyAscii = 27 Then
    Unload Me
End If

'If MebValorDig < 0 Then
'    XPgto = 1
'End If

If KeyAscii = 13 Then
    If TotPgto < CCur(MebTotal) Then
    
         If VarBloqueado = True And txtpgto.ListIndex = 5 Then
            MsgBox "Cliente bloqueado pra venda a prazo !", vbInformation, App.Path
            txtpgto.SetFocus
            Exit Sub
        End If
        
        
        If VarAniversariante = True Then
            If MebDesconto = 0 Then
                MebDesconto = VarDesc15
                MebTotal = MebTotal - MebDesconto
                If MebValorDig = MebValor Then
                    MebValorDig = MebValorDig - CCur(MebDesconto)
                End If
            End If
        
        ElseIf VarDescPromo <> 0 Then
            If MebDesconto = 0 Then
                MebDesconto = VarDescPromo + VarDesc5
                MebTotal = MebTotal - MebDesconto
                If MebValorDig = MebValor Then
                    MebValorDig = MebValorDig - CCur(MebDesconto)
                End If
            End If
            
        End If
            Select Case txtpgto.ListIndex
                Case 0
                    If MebDesconto = 0 And VarDescPgto Then
                        MebDesconto = VarDesc5
                        MebTotal = MebTotal - MebDesconto
                        If MebValorDig = MebValor Then
                            MebValorDig = MebValorDig - CCur(MebDesconto)
                            VarDesc = True
                        End If
                    End If
                
                Case 1
                    If MebDesconto = 0 And VarDescPgto Then
                        MebDesconto = VarDesc5
                        MebTotal = MebTotal - MebDesconto
                        If MebValorDig = MebValor Then
                            MebValorDig = MebValorDig - CCur(MebDesconto)
                            VarDesc = True
                        End If
                    End If
                
                Case 2
                    If MebDesconto = 0 And VarDescPgto Then
                        MebDesconto = VarDesc5
                        MebTotal = MebTotal - MebDesconto
                        If MebValorDig = MebValor Then
                            MebValorDig = MebValorDig - CCur(MebDesconto)
                            VarDesc = True
                        End If
                    End If
                
                
                Case 3
                    If MebDesconto = 0 And VarDescPgto Then
                        MebDesconto = VarDesc5
                        MebTotal = MebTotal - MebDesconto
                        If MebValorDig = MebValor Then
                            MebValorDig = MebValorDig - CCur(MebDesconto)
                            VarDesc = True
                        End If
                    End If
                
                Case 4
                    If MebDesconto = 0 And VarDescPgto Then
                        MebDesconto = VarDesc5
                        MebTotal = MebTotal - MebDesconto
                        If MebValorDig = MebValor Then
                            MebValorDig = MebValorDig - CCur(MebDesconto)
                            VarDesc = True
                        End If
                    End If
                
                Case 5
                
                    If MebDesconto <> 0 Then
                        If TotPgto = MebValor Then
                            MsgBox "Esse forma de pgto não permite desconto", vbInformation
                            MebTotal = MebTotal + CCur(MebDesconto)
                            MebValorDig = MebValorDig + CCur(MebDesconto)
                            MebDesconto = 0
                        Else
                            MsgBox "Esse forma de pgto não permite desconto", vbInformation
                            MebTotal = MebTotal + CCur(MebDesconto)
'                            MebValorDig = MebTotal - TotPgto
                            MebDesconto = 0
                        End If
                     End If
                     VarDescPgto = False
                Case 7
                    If VarDescPromo = 0 And VarDesc15 = 0 Then
                        If MebDesconto <> 0 Then
                            If TotPgto = MebValor Then
                                MsgBox "Esse forma de pgto não permite desconto", vbInformation
                                MebTotal = MebTotal + CCur(MebDesconto)
                                MebValorDig = MebValorDig + CCur(MebDesconto)
                                MebDesconto = 0
                            Else
                                MsgBox "Esse forma de pgto não permite desconto", vbInformation
                                MebTotal = MebTotal + CCur(MebDesconto)
                                MebValorDig = MebValorDig + CCur(MebDesconto)
    '                            MebValorDig = MebTotal - TotPgto
                                MebDesconto = 0
                            End If
                         End If
                         VarDescPgto = False
                    Else
                                   
                        If VarDesc5 <> 0 And VarDesc15 = 0 Then
                            MebDesconto = CCur(MebDesconto) - VarDesc5
                        End If
                        
                        If MebValorDig = MebTotal Then
                            MebTotal = MebValor - CCur(MebDesconto)
                            MebValorDig = MebValor - CCur(MebDesconto)
                        Else
                            If VarDesc15 = 0 Then
                                MebTotal = MebValor - CCur(MebDesconto)
                                MebValorDig = MebValorDig - CCur(MebDesconto)
                            Else
                                MebTotal = MebValor - CCur(MebDesconto)
                                MebValorDig = MebValorDig
                            End If
                        End If
                    End If
                
                
            End Select
        
        MebTotal = Format(MebTotal, "###,##0.00")
            
        If CDec(MebValorDig) >= CCur(MebTotal) Then
            ListPgto.List(XPgto) = psCompString(3, Trim(txtpgto.Text), 15) & "=>" & Alinhar(Format(MebTotal, "###,##0.00"), 9, 2)
            TotPgto = Format(MebValorDig, "###,##0.00")
            XPgto = XPgto + 1
            txtpgto.Enabled = False

        Else
            ListPgto.List(XPgto) = psCompString(3, Trim(txtpgto.Text), 15) & "=>" & Alinhar(Format(MebValorDig, "###,##0.00"), 9, 2)
            XPgto = XPgto + 1
            TotPgto = TotPgto + CCur(MebValorDig)
            TotPgto = Format(TotPgto, "###,##0.00")
            MebValorDig = MebTotal - CCur(TotPgto)
            If MebValorDig > 0 Then
    '            If MebValorDig > (MebTotal * (10 / 100)) Then
                    MebValorDig.SetFocus
                    MebValorDig.SetFocus
    '            Else
    '                MebDesconto = MebDesconto + CCur(MebValorDig)
    '                MebValorDig = 0
    '                MebDoc.SetFocus
    '            End If
            Else
                txtpgto.Enabled = False
                MebDoc.SetFocus
            End If
        End If
    Else
'        If CCur(MebTotal) < 0 Then
            If CDec(MebValorDig) >= CCur(MebTotal) Then
            ListPgto.List(XPgto) = psCompString(3, Trim(txtpgto.Text), 15) & "=>" & Alinhar(Format(MebTotal, "###,##0.00"), 9, 2)
            TotPgto = Format(MebValorDig, "###,##0.00")
            XPgto = XPgto + 1
            txtpgto.Enabled = False

        Else
            MebDoc.SetFocus
            txtpgto.Enabled = False
        End If
    End If
End If

End Sub
