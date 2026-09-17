VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormMdfe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mdf-e"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   14025
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14025
      TabIndex        =   64
      Top             =   7830
      Width           =   14055
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
         Left            =   90
         TabIndex        =   65
         Top             =   45
         Width           =   5505
      End
   End
   Begin VB.CommandButton BtMarcados 
      Caption         =   "Im&primir Marcados"
      Height          =   330
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Imprimir Extrato"
      Top             =   7050
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14055
      TabIndex        =   60
      Top             =   0
      Width           =   14055
      Begin VB.Label LblTitulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   1440
         TabIndex        =   63
         Top             =   45
         Width           =   12390
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4920
      Left            =   90
      TabIndex        =   41
      Top             =   1350
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   8678
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Parcelas"
      TabPicture(0)   =   "FormMdfe.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbgrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BtPagarTodos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DatDados(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "OptOrdem(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "OptOrdem(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "OptOrdem(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "OptOrdem(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "BtBoleto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Itens Pedido"
      TabPicture(1)   =   "FormMdfe.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LblDados(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LblDados(5)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MebTotalPed"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MebDescPed"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Lista"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmbProximo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdAnterior"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Relatório"
      TabPicture(2)   =   "FormMdfe.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LstPedido"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton BtBoleto 
         Caption         =   "Boleto / Duplicata"
         Height          =   330
         Left            =   -69735
         TabIndex        =   61
         Top             =   4410
         Width           =   1950
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Emissão"
         Height          =   285
         Index           =   1
         Left            =   -73830
         TabIndex        =   58
         Top             =   4455
         Width           =   1095
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pedido"
         Height          =   285
         Index           =   0
         Left            =   -74865
         TabIndex        =   57
         Top             =   4455
         Width           =   1050
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Vencimento"
         Height          =   285
         Index           =   2
         Left            =   -72750
         TabIndex        =   56
         Top             =   4455
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pagamento"
         Height          =   285
         Index           =   3
         Left            =   -71400
         TabIndex        =   55
         Top             =   4455
         Width           =   1365
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
         Left            =   13140
         TabIndex        =   50
         ToolTipText     =   "Prôximo pedido"
         Top             =   1035
         Width           =   465
      End
      Begin VB.ListBox LstPedido 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74910
         TabIndex        =   49
         Top             =   540
         Width           =   13560
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
         Left            =   13140
         TabIndex        =   47
         ToolTipText     =   "Prôximo pedido"
         Top             =   480
         Width           =   465
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00C0FFFF&
         Height          =   825
         Left            =   -69360
         TabIndex        =   44
         Top             =   2190
         Visible         =   0   'False
         Width           =   3120
         Begin VB.Label Label5 
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
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   765
            TabIndex        =   45
            Top             =   270
            Width           =   1755
         End
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
         Left            =   -72120
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2205
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton BtPagarTodos 
         Caption         =   "&Selecionar Todos"
         Height          =   285
         Left            =   -62865
         TabIndex        =   43
         Top             =   4440
         Width           =   1500
      End
      Begin MSDBGrid.DBGrid dbgrid1 
         Bindings        =   "FormMdfe.frx":0054
         Height          =   3945
         Left            =   -74910
         OleObjectBlob   =   "FormMdfe.frx":006E
         TabIndex        =   42
         Top             =   360
         Width           =   13605
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   4050
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   12870
         _ExtentX        =   22701
         _ExtentY        =   7144
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
         NumItems        =   6
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
            Object.Width           =   1764
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
            Object.Width           =   2646
         EndProperty
      End
      Begin MSMask.MaskEdBox MebDescPed 
         DataField       =   "Custo"
         Height          =   315
         Left            =   10125
         TabIndex        =   51
         Tag             =   "S"
         Top             =   4680
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
         Left            =   11790
         TabIndex        =   52
         Tag             =   "S"
         Top             =   4680
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Margem:"
         Height          =   195
         Left            =   -67035
         TabIndex        =   59
         Top             =   4455
         Width           =   615
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
         Index           =   5
         Left            =   9405
         TabIndex        =   54
         Top             =   4725
         Width           =   690
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
         Index           =   11
         Left            =   11385
         TabIndex        =   53
         Top             =   4725
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Itens Selecionados:"
         Height          =   195
         Left            =   -65145
         TabIndex        =   48
         Top             =   4470
         Width           =   1395
      End
   End
   Begin VB.CommandButton BtMatricial 
      Caption         =   "Im&primir Matricial"
      Height          =   330
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Imprimir Extrato"
      Top             =   6690
      Width           =   1500
   End
   Begin VB.Frame FrmEnt 
      Height          =   735
      Left            =   90
      TabIndex        =   34
      Top             =   585
      Width           =   13830
      Begin VB.TextBox TxtCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5805
         TabIndex        =   39
         Top             =   315
         Width           =   960
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
         Left            =   1305
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From ENT WHERE CodEntidade = 1 or CodEntidade >= 5 Order by Nome"
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox ChkPagos 
         Caption         =   "&Mostrar Pagos"
         Height          =   285
         Left            =   11115
         TabIndex        =   4
         Top             =   315
         Width           =   1410
      End
      Begin VB.CheckBox ChkJuros 
         Caption         =   "&Cobrar Juros"
         Height          =   285
         Left            =   9765
         TabIndex        =   3
         Top             =   315
         Width           =   1275
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&OK"
         Height          =   330
         Left            =   12555
         TabIndex        =   5
         Top             =   270
         Width           =   1140
      End
      Begin MSMask.MaskEdBox Mebdias 
         DataField       =   "Custo"
         Height          =   315
         Left            =   8325
         TabIndex        =   2
         Tag             =   "S"
         Top             =   315
         Width           =   1230
         _ExtentX        =   2170
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
      Begin MSDBCtls.DBCombo DbcPesquisa 
         Bindings        =   "FormMdfe.frx":2341
         DataField       =   "Nome"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   5400
         _ExtentX        =   9525
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
      Begin MSMask.MaskEdBox MebDesc 
         DataField       =   "Custo"
         Height          =   315
         Left            =   6975
         TabIndex        =   1
         Tag             =   "S"
         Top             =   315
         Width           =   1050
         _ExtentX        =   1852
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
      Begin VB.Label Label3 
         Caption         =   "Código"
         Height          =   195
         Left            =   5805
         TabIndex        =   38
         Top             =   135
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entidade"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   135
         Width           =   630
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto(%)"
         Height          =   195
         Index           =   6
         Left            =   6975
         TabIndex        =   36
         Top             =   135
         Width           =   900
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "&Considerar Dias"
         Height          =   195
         Index           =   7
         Left            =   8325
         TabIndex        =   35
         Top             =   135
         Width           =   1110
      End
   End
   Begin VB.ComboBox LstTipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FormMdfe.frx":235B
      Left            =   1680
      List            =   "FormMdfe.frx":2371
      TabIndex        =   10
      Top             =   7155
      Width           =   2580
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valor à Pagar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4770
      TabIndex        =   21
      Top             =   6840
      Width           =   2355
      Begin MSMask.MaskEdBox MebTotal 
         Height          =   510
         Left            =   90
         TabIndex        =   22
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483640
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir Extrato"
      Top             =   6345
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   9225
      TabIndex        =   16
      Top             =   6435
      Width           =   4650
      Begin MSMask.MaskEdBox MebSubTot 
         DataField       =   "Custo"
         Height          =   270
         Left            =   3150
         TabIndex        =   13
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
         TabIndex        =   14
         Tag             =   "S"
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
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
         TabIndex        =   15
         Tag             =   "S"
         Top             =   870
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   1695
         TabIndex        =   23
         Tag             =   "S"
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
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
         Left            =   1650
         TabIndex        =   24
         Tag             =   "S"
         Top             =   885
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   3135
         TabIndex        =   27
         Tag             =   "S"
         Top             =   855
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
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   3015
         X2              =   3015
         Y1              =   90
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1485
         X2              =   1485
         Y1              =   120
         Y2              =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4635
         Y1              =   630
         Y2              =   630
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
         Left            =   3075
         TabIndex        =   28
         Top             =   675
         Width           =   300
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
         Height          =   270
         Index           =   9
         Left            =   1590
         TabIndex        =   26
         Top             =   150
         Width           =   345
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
         Left            =   1590
         TabIndex        =   25
         Top             =   675
         Width           =   600
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
         TabIndex        =   19
         Top             =   660
         Width           =   1050
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
         TabIndex        =   18
         Top             =   135
         Width           =   960
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
         Left            =   3090
         TabIndex        =   17
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "&Limpa Tela"
      Height          =   330
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7395
      Width           =   1500
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
      Height          =   420
      Left            =   4770
      TabIndex        =   11
      Top             =   6390
      Width           =   2355
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   360
      Left            =   150
      TabIndex        =   8
      Top             =   7125
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   81330177
      CurrentDate     =   36759
   End
   Begin MSMask.MaskEdBox MebDesconto 
      DataField       =   "Custo"
      Height          =   360
      Left            =   135
      TabIndex        =   6
      Tag             =   "S"
      Top             =   6525
      Width           =   1230
      _ExtentX        =   2170
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
      Left            =   1440
      TabIndex        =   7
      Tag             =   "S"
      Top             =   6525
      Width           =   1230
      _ExtentX        =   2170
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
      Left            =   2745
      TabIndex        =   9
      Tag             =   "S"
      Top             =   6525
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   150
      TabIndex        =   33
      Top             =   6915
      Width           =   390
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto"
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   32
      Top             =   6315
      Width           =   690
   End
   Begin VB.Label Juros 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juros"
      Height          =   195
      Index           =   5
      Left            =   1455
      TabIndex        =   31
      Top             =   6345
      Width           =   375
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago"
      Height          =   195
      Index           =   3
      Left            =   2790
      TabIndex        =   30
      Top             =   6345
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Pgto"
      Height          =   195
      Left            =   1680
      TabIndex        =   29
      Top             =   6945
      Width           =   690
   End
End
Attribute VB_Name = "FormMdfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotJuros As Currency
Public TotDesc As Currency
Public VarPedido As String
Public VarParcela As String
Public VarCont As Integer
Public VarOrdem  As String
Public VarJuros As Boolean
Public VarMulta As Boolean
Public VarCodDupli As Double


Private Sub BtBoleto_Click()

If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma emissão duplicata ?", vbYesNo, App.Title) = vbYes Then
        
         DatDados(0).Recordset.MoveFirst
         ValorConta = MebTotal
         DtEmissAnt = Date
'         Juros = 0
         Desconto = 0
         DtVencAnt = Date + 30
         If DatDados(0).Recordset.Eof Then
                MsgBox "Houve alguma falta de informação ! pagamento Cancelado", vbInformation, App.Title
                Exit Sub
         End If
         
         Do Until DatDados(0).Recordset.Eof
'            If GetInputState() Then DoEvents
         
            If DatDados(0).Recordset("PgtTemp") = -1 Then
                DatDados(0).Recordset.Edit
'                DatDados(0).Recordset.Fields("Tipo_Dup") = "DP"
'                DatDados(0).Recordset.Fields("DtPgto") = DtData
'                If MebDesconto.Enabled = False Then
'                    ValorPcl = (DatDados(0).Recordset("Valor") + DatDados(0).Recordset("Juros")) - DatDados(0).Recordset("Desconto")
'                    DatDados(0).Recordset.Fields("VlrPago") = ValorPcl
'                    Juros = Juros + DatDados(0).Recordset("Juros")
'                Else
'                    ValorPcl = (DatDados(0).Recordset("Valor") + MebJuros) - MebDesconto
'                    DatDados(0).Recordset.Fields("VlrPago") = MebPago
'                    Juros = Juros + MebJuros
'                End If
                
                DatDados(0).Recordset.Update
'                Flag = True
'                TotalPago = CCur(TotalPago) - CCur(ValorPcl)
'                DtEmissAnt = DatDados(0).Recordset("DtEmiss")
'                DtVencAnt = DatDados(0).Recordset("DtVenc")
                VarDoc = DatDados(0).Recordset("Pedido") & "/" & DatDados(0).Recordset("NroPcl")
                VarPed = DatDados(0).Recordset("Pedido")
                VarParc = DatDados(0).Recordset("NroPcl")
                DtVenc = DatDados(0).Recordset("DtVenc")
                 
                                
            End If
            DatDados(0).Recordset.MoveNext
         Loop
                  
         With FormBoleto
            .TxtEnt = TxtCod
        '    If Opt(2).value = True Then
        '        .TxtPedido = DBGrid1.Columns(0)
        '    Else
        '        .TxtPedido = DBGrid1.Columns(1)
        '    End If
        '    If DBGrid1.Columns(2) = "" Then
        '        MsgBox "Nenhuma parcela selecionada", vbInformation, App.Title
        '        CmdLimpa_Click
        '        Exit Sub
        '    End If
        '    .TxtPcl = DBGrid1.Columns(3)
        '    .TxtDoc = DBGrid1.Columns(14)
        '    .Opt(0) = Opt(0).value
        '    .Opt(1) = Opt(1).value
        '    .Opt(2) = Opt(2).value
        '    .MebBanco = MebBanco
            .TxtEnt = TxtCod
            .MebDtSaida = DtVenc            'MebDtSaida
            .MebEmiss = Date
            .MebVenc = Date + 30
            .MebValor = MebTotal
            .OPt(2).value = True
        '    If Opt(0).value = True Then
'                Set soma = Banco.OpenRecordset("SELECT DUP.Pedido,Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor  FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(1)))
        '    ElseIf Opt(1).value = True Then
'                    If Not IsNull(DBGrid1.Columns(13)) And DBGrid1.Columns(14) <> "" Then
'                        Set soma = Banco.OpenRecordset("SELECT Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor FROM DUP GROUP BY DUP.Documento HAVING DUP.Documento Like '" & Trim(DBGrid1.Columns(14)) & "'")
'                    Else
'                        MsgBox "Campo documento não está preenchido", vbInformation, App.Title
'                    End If
        '        Else
'                    Set soma = Banco.OpenRecordset("SELECT DUP.Pedido,Sum(DUP.Juros) AS SomaDeJuros, Sum(DUP.Valor) AS SomaDeValor  FROM DUP GROUP BY DUP.Pedido HAVING DUP.Pedido = " & CLng(DBGrid1.Columns(0)))
        '    End If
         '   If soma.RecordCount > 0 Then
'                .MebTotal = CCur(soma!SomaDeValor) + CCur(soma!somadejuros)
        '    Else
'                .MebTotal = CCur(DBGrid1.Columns(7)) + CCur(DBGrid1.Columns(8))
        '    End If
         End With
        
         FormBoleto.Show vbModal
         
         
         
    End If
End If



End Sub

Private Sub BtMarcados_Click()
Dim Formula As String
Dim flag As Boolean
Dim SQL As String
flag = True

     DatDados(0).Recordset.MoveFirst
     
     Formula = ""
     Do Until DatDados(0).Recordset.Eof
'            If GetInputState() Then DoEvents
        
        If DatDados(0).Recordset("PgtTemp") = -1 Then
            If flag Then
                Formula = " Ped.CodPed = " & DatDados(0).Recordset("Pedido")
                flag = False
            End If
            Formula = Formula & " or Ped.CodPed = " & DatDados(0).Recordset("Pedido")
        
        End If

       DatDados(0).Recordset.MoveNext
    Loop
'    Formula = Formula  & " and (Operacao = 1 or Operacao = 2 ) "
        
'    If varFormRelatorio.Name = "FrmRelConf" Then
'        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
'        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ENT.Nome, PED.Operacao "
        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto,Itp.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp, Ped.Desconto, Ped.Operacao FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto,PRD.Descrição_Produto, Itp.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ENT.Nome, PED.Operacao, Ped.Desconto "
'        Ordenado = " Order by CodPed, CodItp "
'        Formula = VarData & "  >= #" & Data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
        SQL = CamposSql & " Where " & Formula & CamposGroup '& Ordenado
'    End If
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelConf.Relatorio.Recordset = varPesquisa(1)
    FrmRelConf.rpCabTitulo.Caption = "Conferencia"
    FrmRelConf.rpCabPeriodo.Caption = Periodo
    FrmRelConf.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelConf.Config
        
End Sub

Private Sub BtMatricial_Click()
Dim Resp  As String

Resp = MsgBox("Impressão resumida ? ", vbYesNoCancel, App.Title)
If Resp = vbYes Then
    Set Arq = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro FROM Ent WHERE  Ent.CodEntidade  = " & TxtCod)

    Iniciar (False)
    Print #1, ""
    Print #1, "" & Cond
    Print #1, Gde & Centralizar(NomeEmpresa, 23) & Cond
 '   Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 48)
    Print #1, "================================================"
    Print #1, "Data: " & Date & " Hora: " & Time; Tab(33)
'    Print #1, Tab(1); psCompString(1, TxtCod, 6); "-"; Left(DbcPesquisa.Text, 30)
    Print #1, Tab(1); psCompString(1, Arq!CodEntidade, 6); "-"; Left(Arq!Nome, 30)
    Print #1, Tab(1); Trim(Left(Arq!Endereco, 30)); ", "; Arq!Nro
    Print #1, Tab(1); "Ped."; Tab(8); "Pcl"; Tab(11); "Emiss."; Tab(22); "VENC."; Tab(39); "VALOR"
    Print #1, "------------------------------------------------"
    
    DatDados(0).Recordset.MoveFirst
    While Not DatDados(0).Recordset.Eof
        Print #1, Tab(1); psCompString(1, DatDados(0).Recordset("Pedido"), 6); Tab(8); psCompString(1, DatDados(0).Recordset("NroPcl"), 2); Tab(11); DatDados(0).Recordset("DtEmiss"); Tab(22); DatDados(0).Recordset("DtVenc"); Tab(35); Alinhar(Format(DatDados(0).Recordset("Valor"), "##,##0.00"), 9, 2); Tab(45); DatDados(0).Recordset("Tipo_Dup"); Tab(48); IIf(DatDados(0).Recordset("Cond"), "P", "A")
        DatDados(0).Recordset.MoveNext
    Wend
    
    Print #1, Tab(35); "---------"
    Print #1, Tab(18); "SUB-TOTAL =>"; Tab(34); Alinhar(CStr(MebSubTot.Text), 10, 2)
    If ChkPagos.value = 1 Then
        Print #1, Tab(18); "PAGOS    =>"; Tab(34); Alinhar(CStr(MebPagos.Text), 10, 2)
    End If
    Print #1, Tab(18); "JUROS     =>"; Tab(34); Alinhar(CStr(MebTotJuros.Text), 10, 2)
    Print #1, Tab(18); "DESCONTO  =>"; Tab(34); Alinhar(CStr(MebTotDesc.Text), 10, 2)
    Print #1, Tab(18); "TOTAL     =>"; Tab(34); Alinhar(CStr(MebTotGeral.Text), 10, 2)
    Print #1, "" & Nor
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Fim (False)
ElseIf Resp = vbNo Then
    ImprimiDiscr TxtCod
End If

End Sub

Private Sub BtPagarTodos_Click()

With DatDados(0).Recordset
    
    If .RecordCount > 0 Then
        .MoveFirst
        While Not .Eof
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
MDIPrincipal.RptRel.ReportFileName = Caminho & "\relparcelas.RPT"
MDIPrincipal.RptRel.Formulas(0) = "Titulo = 'RELATÓRIO DE PARCELAS  " & "  -  " & NomeEmpresa & "'"
If ChkPagos.value = 0 Then
    MDIPrincipal.RptRel.SelectionFormula = "{DUP.Entidade} = " & DbcPesquisa.BoundText & " AND {Dup.Cond} = false "
Else
    MDIPrincipal.RptRel.SelectionFormula = "{DUP.Entidade} = " & DbcPesquisa.BoundText
End If
Imprimir

End Sub

Private Sub CmdExcluir_Click()
Dim Criterio As String
If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma Exclusão da Duplicata ?", vbYesNo, App.Title) = vbYes Then
        Criterio = "Numero  = " & """" & TxtDupl & """"
        DatDados(0).Recordset.FindFirst Criterio
        If Not DatDados(0).Recordset.NoMatch Then
            DatDados(0).Recordset.Delete
        End If
        DatDados(0).Refresh
        SomaGrid
        MebPago = 0
        TxtDupl = ""
    End If
End If

End Sub


Private Sub CmbProximo_Click()
If Not DatDados(0).Recordset.Eof Then
    DatDados(0).Recordset.MoveNext
    SSTab1_Click 1
Else
    MsgBox "Final do arquivo !", vbInformation, App.Title
End If
End Sub

Private Sub CmdAnterior_Click()

If Not DatDados(0).Recordset.Bof Then
    DatDados(0).Recordset.MovePrevious
    SSTab1_Click 1
Else
    MsgBox "Inicio do arquivo !", vbInformation, App.Title
End If

End Sub

Private Sub CmdLimpa_Click()
MebPago = 0
LblTitulo = ""
TxtCod = 0
DbcPesquisa = ""
MebSubTot = 0
MebPagos = 0
MebDesconto = 0
MebJuros = 0
MebAPagar = 0
MebTotal = 0
MebDesc = 0
Mebdias = 0
DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade = 0  ORDER BY " & VarOrdem
DatDados(0).Refresh

SomaGrid
FrmEnt.Enabled = True

If VarJuros = True Then
    ChkJuros.value = 1
Else
    ChkJuros = 0
End If

ChkPagos.value = 0

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

Private Sub CmdPgto_Click()
Dim Criterio As String
Dim ValorConta As Currency
Dim TotalPago As Currency
Dim Desconto As Currency
Dim Juros As Currency
Dim Valor As Currency
Dim ValorPcl As Currency
Dim SQL As String
Dim Dupli As Long
Dim DtEmissAnt As Date
Dim DtVencAnt As Date
Dim VarDoc As String
Dim VarParc As Double
Dim VarPed As Double
Dim flag As Boolean
Dim VarQtdPcl As Recordset
Dim VarSaldo As Boolean
Dim TotSaldo As Currency
Dim VarReciboGrande As Boolean
Dim VarTabSaldo As Recordset
Dim VarJuros As Currency


If LerINI("PEDIDO", "NaoImprimirSaldo", App.Path & "\config.ini") = "" Then
    VarSaldo = True
Else
    If LerINI("PEDIDO", "NaoImprimirSaldo", App.Path & "\config.ini") = "0" Then
        VarSaldo = True
    Else
        VarSaldo = False
    End If
End If


If LerINI("PEDIDO", "ReciboGrande", App.Path & "\config.ini") = "" Then
    VarReciboGrande = False
Else
    If LerINI("PEDIDO", "ReciboGrande", App.Path & "\config.ini") = "1" Then
        VarReciboGrande = True
        Set VarTabSaldo = Banco.OpenRecordset("SELECT DUP.Pedido, Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Count(DUP.NroPcl) AS ContarDeNroPcl FROM DUP GROUP BY DUP.Pedido, DUP.Cond HAVING (((DUP.Pedido)=999999) AND ((DUP.Cond)=False))")
    Else
        VarReciboGrande = False
    End If
End If


'On Error GoTo erro

VarDoc = ""


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

If Not IsNumeric(DbcPesquisa.BoundText) Then
    MsgBox "Entidade errada", vbInformation
    CmdLimpa_Click
    Exit Sub
End If

Dupli = 0

If DatDados(0).Recordset.RecordCount <> 0 And MebPago <> "" Then
    If MsgBox("Confirma pagamento ?", vbYesNo, App.Title) = vbYes Then
        If Dir(Caminho & "\Recebe.txt") <> "" Then
            Kill Caminho & "\Recebe.txt"
        End If
         DatDados(0).Recordset.MoveFirst
         ValorConta = MebTotal
         TotalPago = MebPago
         DtEmissAnt = Date
         Juros = 0
         Desconto = 0
         DtVencAnt = Date + 30
         If DatDados(0).Recordset.Eof Then
                MsgBox "Houve alguma falta de informação ! pagamento Cancelado", vbInformation, App.Title
                Exit Sub
         End If
         flag = False
         
         Do Until DatDados(0).Recordset.Eof
            VarJuros = 0
'            If GetInputState() Then DoEvents
         
            If DatDados(0).Recordset("PgtTemp") = -1 Then
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("Cond") = -1
                DatDados(0).Recordset.Fields("DtPgto") = DtData
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
                VarJuros = DatDados(0).Recordset("Juros")
                
                DatDados(0).Recordset.Update
                flag = True
                TotalPago = CCur(TotalPago) - CCur(ValorPcl)
                DtEmissAnt = DatDados(0).Recordset("DtEmiss")
                DtVencAnt = DatDados(0).Recordset("DtVenc")
                VarDoc = DatDados(0).Recordset("Pedido") & "/" & DatDados(0).Recordset("NroPcl")
                VarPed = DatDados(0).Recordset("Pedido")
                VarParc = DatDados(0).Recordset("NroPcl")
                Set VarQtdPcl = Banco.OpenRecordset("SELECT * FROM DUP WHERE DUP.Pedido= " & VarPed)
                VarQtdPcl.Requery
                VarQtdPcl.MoveLast
                VarQtdPcl.MoveFirst
                
                Open Caminho & "\recebe.txt" For Append As #1
'                Print #1, "Pedido; " & VarCodPed & "; " & DbcDados(0).Text & "; " & Date & "; " & Time & "; " & Maquina & "; " & CmbDados(3).Text & "; " & Alinhar(TxtDados(2), 10, 2)
                Print #1, psCompString(2, DatDados(0).Recordset.Fields("VlrPago"), 10) & psCompString(2, VarPed, 8) & psCompString(2, VarParc, 8) & psCompString(2, VarQtdPcl.RecordCount, 8) & DatDados(0).Recordset("DtVenc") & psCompString(2, VarJuros, 10)

                Close #1
                               
                
            End If
            DatDados(0).Recordset.MoveNext
         Loop
         
         If flag = True Then
            
             If MebDesconto.Enabled = True Then
                Dupli = VarCodDupli
             End If
                                                   
             If TotalPago < 0 Then
                TotalPago = TotalPago * (-1)
                If MsgBox("Faltou R$ " & Format(TotalPago, "#,##0.00;- #,##0.00") & " Sim para gerar saldo ou Não para dar desconto", vbYesNo, App.Title) = vbYes Then
    
                    If MebDesconto.Enabled = False Then
                        Criterio = "Pedido = " & VarPed & " and NroPcl = " & VarParc
                        DatDados(0).Recordset.FindFirst Criterio
                        DatDados(0).Recordset.Edit
                        DatDados(0).Recordset.Fields("VlrPago") = DatDados(0).Recordset.Fields("Valor") - TotalPago
                        DatDados(0).Recordset.Update
                    End If
    
                    Criterio = "Pedido = " & VarPed & " and NroPcl = " & VarParc
                    DatDados(0).Recordset.FindFirst Criterio
         
                    SQL = "INSERT INTO DUP ( DtEmiss, Pedido, NroPcl, Entidade, DtVenc, Valor, Cond, Documento, Tipo_Dup  ) VALUES ( "
                    
   '                     SQL = SQL & "#" & Format(DtEmissAnt, "mm/dd/yyyy") & "# , "
                    
                    If VarDtSaldoAtual = False Then
                        SQL = SQL & "#" & Format(DatDados(0).Recordset.Fields("DtEmiss"), "mm/dd/yyyy") & "# , "
                    Else
                        SQL = SQL & "#" & Format(Date, "mm/dd/yyyy") & "# , "
                    End If

                    SQL = SQL & DatDados(0).Recordset.Fields("Pedido") & " , "
                    SQL = SQL & DatDados(0).Recordset.Fields("NroPcl") & " , "
                    SQL = SQL & DbcPesquisa.BoundText & " , #"
                    If VarDtSaldoAtual = False Then
                        SQL = SQL & Format(DatDados(0).Recordset.Fields("DtVenc"), "mm/dd/yyyy") & "# , "
                    Else
                        SQL = SQL & Format(Date + 30, "mm/dd/yyyy") & "# , "
                    End If
                    
                    SQL = SQL & Numero2(CStr(TotalPago)) & " ,"
                    SQL = SQL & 0 & " ,"
                    SQL = SQL & """" & VarDoc & """" & ", "
                    SQL = SQL & """" & "SD" & """" & ")"
                    Banco.Execute SQL
                    
                    If MsgBox("Comprovante gerado com sucesso, Deseja Imprimir ?", vbYesNo, App.Title) = vbYes Then
                          Iniciar (False)
    '                       Open "c:\lixo.txt" For Output As #1
                           Print #1, "================================="
                           Print #1, "  COMPROVANTE DE SALDO DE CONTA  "
                           Print #1, "================================="
                           Print #1, "Cliente: "; TxtCod
    '                       Print #1, Chr(15) & DbcPesquisa.Text
    '                       Print #1, Trim(DatDados(1).Recordset.Fields("Endereço")) & ", " & DatDados(1).Recordset.Fields("Nro")
    '                       Print #1, Chr(18)
                           Print #1, "Total Conta.: "; Format(VarTotal, "#,##0.00;-#,##0.00")
                           Print #1, "Valor Pago..: "; Format(MebPago, "#,##0.00;-#,##0.00")
                           Print #1, ""
                           Print #1, "Data....: "; Format(DtData, "dd/mm/yyyy")
                           Print #1, "  "
                           Print #1, "Saldo...: "; Format(TotalPago, "#,##0.00;-#,##0.00")
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " ---------------------------"
                           Print #1, "        Assinatura             "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
                           Print #1, " "
'                           Print #1, Chr(18);
    '                       Close #1
                            Fim (False)
                        End If
                     
                Else
                    MebDesconto = CCur(MebDesconto) + CCur(TotalPago)
                End If
             Else
                If TotalPago <> 0 Then
                    MsgBox "Troco de R$ " & Format(TotalPago, "#,##0.00;-#,##0.00"), vbInformation, App.Title
                End If
             End If
             
            SQL = "INSERT INTO REC ( CodEnt, Data, Hora, Valor, Desconto, Tipo, CodDupli, Juros, Maquina  ) VALUES ( "
            SQL = SQL & DbcPesquisa.BoundText & " , "
            SQL = SQL & "#" & Format(DtData, "mm/dd/yyyy") & "# , "
            SQL = SQL & """" & Time & """" & " , "
            If CCur(MebPago) < CCur(MebTotal) Then
               SQL = SQL & Numero2(CStr(MebPago)) & " , "
            Else
               SQL = SQL & Numero2(CStr(MebTotal)) & " , "
            End If
            SQL = SQL & Numero2(CStr(MebDesconto)) & " , "
            SQL = SQL & LstTipo.ListIndex + 1 & " ,"
            SQL = SQL & Dupli & " ,"
            SQL = SQL & Numero2(CStr(Juros)) & " ,"
            SQL = SQL & """" & LIMPACampo(Trim(Maquina)) & """" & ")"
            Banco.Execute SQL
                                   
             If MsgBox("Imprimir comprovante de recebimento ?", vbYesNo, App.Title) = vbYes Then
                
                Iniciar (False)
'                Open "c:\lixo.txt" For Output As #1
                If VarReciboGrande = False Then
                    Print #1, "======================================"
                    Print #1, " COMPROVANTE DE RECEBIMENTO (CAIXA)   "
                    Print #1, "======================================"
                    
                    Print #1, "Cliente: "; TxtCod; "            "; Format(DtData, "dd/mm/yyyy")
                    Print #1, Cond & DbcPesquisa.Text
                End If
                Print #1, Nor
                
                If VarReciboGrande = False Then
                                
                    If VarCont = 1 Then
                        Print #1, "Pedido.....: "; VarPedido; " - " & VarParcela & "/"; VarQtdPcl.RecordCount
                        Print #1, ""
                        If MebJuros <> 0 Then
                            Print #1, "Valor Parcela " & VarParcela & " : "; Format(CCur(MebTotal) - CCur(MebJuros), "#,##0.00;-#,##0.00")
        '                    Print #1, " "
                            Print #1, "Juros     : "; Format(MebJuros, "#,##0.00;-#,##0.00")
                        Else
                            Print #1, "Valor Parcela " & VarParcela & " : "; Format(MebTotal, "#,##0.00;-#,##0.00")
                        End If
                    
                    Else
                    
                         Open Caminho & "\recebe.txt" For Input As #6
                         Do
                            Line Input #6, VarTemp
                            Print #1, "Pcl: "; Format(Mid(VarTemp, 1, 10), "#,##0.00"), Tab(20); " Ped: "; Trim(Mid(VarTemp, 11, 8)) & " - " & Trim(Mid(VarTemp, 19, 8)) & "/" & Trim(Mid(VarTemp, 27, 8))
                            
                        Loop Until Eof(6)
                        Close #6
                        If Dir(Caminho & "\Recebe.txt") <> "" Then
                            Kill Caminho & "\Recebe.txt"
                        End If
                                            
                    End If
                Else
                
                    If VarCont = 1 Then
                        
                        For x = 1 To 2
                        

                            Print #1, "======================================"
                            Print #1, " COMPROVANTE DE RECEBIMENTO (CAIXA)   "
                            Print #1, "======================================"
                            Print #1, "Cliente: "; TxtCod; "            "; Format(DtData, "dd/mm/yyyy")
                            Print #1, Cond & DbcPesquisa.Text
                            
'                            Print #1, "Pedido.....: "; VarPedido; " - " & VarParcela & "/"; VarQtdPcl.RecordCount
                            Print #1, ""
                            
                            Print #1, "Pcl  : "; Format(CCur(MebTotal) - CCur(MebJuros), "#,##0.00;-#,##0.00"), Tab(20); " Ped: "; Trim(VarPedido) & " - " & Trim(VarParcela) & "/" & Trim(VarQtdPcl.RecordCount)
                            Print #1, "Juros: "; Format(MebJuros, "#,##0.00")
                            Print #1, "Total: "; Format(MebTotal, "#,##0.00")
'                                Print #1, "Pcl  : "; Format(Mid(varTemp, 1, 10) - Mid(varTemp, 45, 10), "#,##0.00"), Tab(20); " Ped: "; Trim(Mid(varTemp, 11, 8)) & " - " & Trim(Mid(varTemp, 19, 8)) & "/" & Trim(Mid(varTemp, 27, 8))
'                                Print #1, "Juros: "; Format(Mid(varTemp, 45, 10), "#,##0.00")
'                                Print #1, "Total: "; Format(Mid(varTemp, 1, 10), "#,##0.00")

                            Print #1, ""
                            Print #1, "Data Venc..: "; Format(DtVencAnt, "dd/mm/yyyy")
                            Print #1,
                            
'                            If MebJuros <> 0 Then
'                                Print #1, "Valor Parcela " & VarParcela & " : "; Format(CCur(MebTotal) - CCur(MebJuros), "#,##0.00;-#,##0.00")
            '                    Print #1, " "
'                                Print #1, "Juros     : "; Format(MebJuros, "#,##0.00;-#,##0.00")
'                            Else
            '                    Print #1, "Valor Parcela " & VarParcela & " : "; Format(MebTotal, "#,##0.00;-#,##0.00")
'                            End If
                            
                            Print #1, "=========<< ----------- >> ============"
                            Print #1, ""
                            Print #1, ""
                        Next
                    Else
                        For x = 1 To 2
                             Open Caminho & "\recebe.txt" For Input As #6
                             Do
                                Line Input #6, VarTemp
                                Print #1, "======================================"
                                Print #1, " COMPROVANTE DE RECEBIMENTO (CAIXA)   "
                                Print #1, "======================================"
                                
                                Print #1, ""
                                Print #1, "Cliente: "; TxtCod; "            "; Format(DtData, "dd/mm/yyyy")
                                Print #1, Cond & DbcPesquisa.Text
                                Print #1, ""
                                Print #1, "Pcl  : "; Format(Mid(VarTemp, 1, 10) - Mid(VarTemp, 45, 10), "#,##0.00"), Tab(20); " Ped: "; Trim(Mid(VarTemp, 11, 8)) & " - " & Trim(Mid(VarTemp, 19, 8)) & "/" & Trim(Mid(VarTemp, 27, 8))
                                Print #1, "Juros: "; Format(Mid(VarTemp, 45, 10), "#,##0.00")
                                Print #1, "Total: "; Format(Mid(VarTemp, 1, 10), "#,##0.00")
                                Print #1,
                                Print #1, "Data Venc..: "; Format(Mid(VarTemp, 35, 10), "dd/mm/yyyy")
                                Set VarTabSaldo = Banco.OpenRecordset("SELECT DUP.Pedido, Sum(DUP.Valor) AS SomaDeValor, DUP.Cond, Count(DUP.NroPcl) AS ContarDeNroPcl FROM DUP GROUP BY DUP.Pedido, DUP.Cond HAVING (((DUP.Pedido)= " & CDbl(Mid(VarTemp, 12, 7)) & ") AND ((DUP.Cond)=False))")
                                VarTabSaldo.Requery
    '                            If VarTabSaldo.RecordCount > 0 Then
    '                                VarTabSaldo.MoveLast
    '                                VarTabSaldo.MoveFirst
    '                                Print #1, "Falta " & VarTabSaldo!ContardeNroPcl & " parc., Saldo do carne:" & Format(VarTabSaldo!SomaDeValor, "#,##0.00")
    '                            End If
                                Print #1, "--------------------------------------"
                                Print #1, ""
                                Print #1, ""
                                Print #1, ""
                                
                            Loop Until Eof(6)
                            Close #6
'                            If Dir(Caminho & "\Recebe.txt") <> "" Then
'                                Kill Caminho & "\Recebe.txt"
'                            End If
                            
                            Print #1, "=========<< ----------- >> ============"
                            Print #1, ""
                            Print #1, ""
'                            MsgBox "Imprimindo segunda via !", vbInformation
                        Next
                        If Dir(Caminho & "\Recebe.txt") <> "" Then
                            Kill Caminho & "\Recebe.txt"
                        End If

                    End If
                                                
                End If
                
'                Print #1, " "
                If CCur(MebPago) < CCur(MebTotal) Then
                    Print #1, "Valor Pago.: "; Format(MebPago, "#,##0.00;-#,##0.00")
                Else
                    Print #1, "Valor Pago.: "; Format(MebTotal, "#,##0.00;-#,##0.00")
                End If
                Print #1, ""
                TotSaldo = MebPago
                If CCur(MebPago) > CCur(MebTotal) Then
                    Print #1, "Troco.....: "; Format(TotalPago, "#,##0.00;-#,##0.00")
                    TotSaldo = MebPago - TotalPago
                End If
'                Print #1, ""
                Print #1, "Tipo Pgto..: "; LstTipo.Text
'                Print #1, ""
                If VarCont = 1 Then
                    Print #1, "Data Venc..: "; Format(DtVencAnt, "dd/mm/yyyy")
                Else
                    Print #1, "Data.......: "; Format(DtData, "dd/mm/yyyy")
                End If
'                Print #1, " "
                Print #1, " "
                MebTotGeral = Format(MebTotGeral, "#,##0.00;-#,##0.00")
'                MebPago = Format(MebPago, "#,##0.00;-#,##0.00")
                MebTotal = Format(MebTotal, "#,##0.00;-#,##0.00")
                
                If VarSaldo = True Then
                
                        Print #1, "Total Conta.: "; Format(MebTotGeral, "#,##0.00;-#,##0.00")
        '                Print #1, ""
                        If CCur(MebPago) < CCur(MebTotal) Then
                            Print #1, "Valor Pago..: "; Format(MebPago, "#,##0.00;-#,##0.00")
        '                    Print #1, "  "
                            If MebDesconto = 0 Then
                                Print #1, "Saldo.......: "; Format(MebTotGeral - TotSaldo, "#,##0.00;-#,##0.00")
                            Else
                                Print #1, "Desconto....: "; Format(MebDesconto, "#,##0.00;-#,##0.00")
                                Print #1, "Saldo.......: "; Format(MebTotGeral - TotSaldo, "#,##0.00;-#,##0.00")
                            End If
                        Else
                            Print #1, "Valor Pago..: "; Format(MebTotal, "#,##0.00;-#,##0.00")
        '                    Print #1, "  "
                            Print #1, "Saldo.......: "; Format(MebTotGeral - TotSaldo, "#,##0.00;-#,##0.00")
                        End If
                End If
                
                
                
'                Print #1, "Data....: "; Format(DtData, "dd/mm/yyyy")
'                Print #1, " "
'                Print #1, " "
'                Print #1, " "
                Print #1, " "
                Print #1, " "
                Print #1, " "
                Print #1, " "
                Print #1, " "
                Print #1, Nor;
    '            Close #1
                Fim (False)
            End If
        Else
            MsgBox "Houve uma falha na baixa, favor fechar essa tela e refazer a operação", vbInformation
        End If
        If LstTipo.ListIndex = 1 Or LstTipo.ListIndex = 2 Then
           If InterFinanc = True Then
                FrmMovCheq.MebDados(3) = MebPago
                FrmMovCheq.MebDados(10) = DtData
                FrmMovCheq.Show
           End If
        End If
        
        MebTotal = 0
        MebPago = 0
        MebDesconto = 0
        MebJuros = 0
        CmdPgto.Enabled = False
        MebPago.Enabled = False
        BtBoleto.Enabled = False
'         TxtDupl = ""
        DatDados(0).Refresh
        SomaGrid
        VarCont = 0
'         CmdLimpa_Click
    End If
End If

Exit Sub
Erro:

If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If

End Sub

Private Sub Pesquisa(Codigo As Integer)
Dim Criterio As String
Frame.Visible = True
Dim SQL As String
'DoEvents

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
    LblTitulo = DbcPesquisa.Text

    If ChkPagos.value = 0 Then
        DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & " AND Cond = FALSE ORDER BY " & VarOrdem
    Else
        DatDados(0).RecordSource = "SELECT * FROM Dup WHERE Entidade = " & Val(Codigo) & " ORDER BY " & VarOrdem
    End If
    'If GetInputState() Then DoEvents
    
    DatDados(0).Refresh
    
    If DatDados(0).Recordset.RecordCount <> 0 Then
        SQL = "UPDATE DUP SET DUP.PgtTemp = False WHERE (((DUP.Entidade)= " & Val(Codigo) & ") AND ((DUP.Cond)=False))"
        Banco.Execute SQL
         DatDados(0).Recordset.MoveLast
    End If
    SomaGrid
End If

End Sub


Private Sub DbcPesquisa_LostFocus()
If IsNumeric(DbcPesquisa.BoundText) Then
    TxtCod.Text = DbcPesquisa.BoundText
End If
End Sub

Private Sub DBGrid1_DBLClick()
'TxtDupl.Text = ""
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
'        CmdPgto.Enabled = False
    End If
    VarPedido = dbgrid1.Columns(1)
    VarParcela = dbgrid1.Columns(2)
    VarCodDupli = dbgrid1.Columns(0)
    FrmEnt.Enabled = False
    If DatDados(0).Recordset.RecordCount <> 0 Then
        If dbgrid1.Columns(16) = "Não" Then
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
            
        ElseIf dbgrid1.Columns(16) = "Sim" Then
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
    
End If
Label6.Caption = "Itens Selecionados: " & VarCont

End Sub

Private Sub DTData_LostFocus()
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'        If KeyCode = 13 Then
'            Screen.ActiveControl.SelStart = 0
'            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
'        End If
'        If KeyCode = 9 Then
'            Screen.ActiveControl.SelStart = 0
'            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
'        End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
Set Formulário = Me

'Me.Left = (MDIPrincipal.Width - Me.Width) / 2
'Me.Top = (MDIPrincipal.Height - Me.Height) / 6

Me.Left = 50
Me.Top = 30


'DatDados(0).RecordSource = "SELECT * FROM Dup  WHERE Entidade = " & Val(DbcPesquisa.BoundText) & " AND Cond = false AND Saldo = 0 ORDER BY DtVenc"
'DatDados(0).Refresh

DatDados(1).RecordSource = "SELECT * FROM Ent  WHERE (Tipo = 1 or Tipo >= 5) Order By Nome"
DatDados(1).Refresh


DtData = Date
MebPago = 0
MebDesconto = 0
MebJuros = 0
MebTotal = 0
MebDesc = 0
Mebdias = 0
LstTipo.ListIndex = 0
CmdPgto.Enabled = False
MebPago.Enabled = False

SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
Dim Componente As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

If Admin Then
  DtData.Enabled = True
End If

VarOrdem = "DtVenc"

If LerINI("Financeiro", "CobrarJuros", App.Path & "\config.ini") <> "" Then
    VarJuros = LerINI("Financeiro", "CobrarJuros", App.Path & "\config.ini")
Else
    VarJuros = False
End If

If VarJuros Then
    ChkJuros = 1
Else
    ChkJuros = 0
End If

If LerINI("Financeiro", "CobrarMulta", App.Path & "\config.ini") <> "" Then
    VarMulta = LerINI("Financeiro", "CobrarMulta", App.Path & "\config.ini")
Else
    VarMulta = False
End If

End Sub

Private Sub MebDesc_LostFocus()
If IsNumeric(MebDesc) Then
    If MebDesc <> 0 Then
        Mebdias = 30
    Else
        Mebdias = 0
    End If
Else
    MebDesc = 0
End If
End Sub

'Private Sub LstTipo_LostFocus()
'If LstTipo.ListIndex = 1 Or LstTipo.ListIndex = 2 Then
'    MebChPre.SetFocus
'End If
'End Sub

Private Sub MebDesconto_lostfocus()
If MebTotal <> "" And MebDesconto <> "" And IsNumeric(MebDesconto) Then
    MebTotal = CCur(MebTotal) - CCur(MebDesconto)
End If
MebPago = Format(CCur(MebTotal), "#,##0.00;(#,##0.00)")
MebJuros.SetFocus
End Sub

Private Sub SomaGrid()
Dim SomaGrid As Recordset
Dim Contas As Recordset
Dim SqlPed As Recordset

Dim Pago As Currency
Dim Apagar As Currency
Dim Saldo As Currency
Dim Prazo As Integer
Dim VarMargem As Currency
Dim VarMedia As Currency
Dim VarCont As Currency
Pago = 0
Apagar = 0
Saldo = 0
TotJuros = 0
TotDesc = 0
VarMedia = 0
VarCont = 0
'If GetInputState() Then DoEvents

'Set SomaGrid = DatDados(0).Recordset.Clone
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveFirst
    Do Until DatDados(0).Recordset.Eof
        If DatDados(0).Recordset.Fields("Cond") = True Then
            Pago = Pago + DatDados(0).Recordset.Fields("VlrPago")
        Else
            If VarReajConta = True Then
                If VarPcoPrazo = True Then
                    Set Contas = Banco.OpenRecordset("SELECT PED.Total, DUP.Pedido, DUP.Cond, Sum([itp].[qtd]*[prd].[PcoPrazo]) AS Expr1, Ped.Desconto FROM DUP INNER JOIN ((ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON DUP.Pedido = PED.CodPed GROUP BY PED.Total, DUP.Pedido, DUP.Cond, Ped.Desconto HAVING (((DUP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & " ) AND ((DUP.Cond)=False)) ORDER BY DUP.Pedido ")
                Else
                    Set Contas = Banco.OpenRecordset("SELECT PED.Total, DUP.Pedido, DUP.Cond, Sum([itp].[qtd]*[prd].[venda]) AS Expr1, Ped.Desconto FROM DUP INNER JOIN ((ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) INNER JOIN PED ON ITP.Pedido = PED.CodPed) ON DUP.Pedido = PED.CodPed GROUP BY PED.Total, DUP.Pedido, DUP.Cond, Ped.Desconto HAVING (((DUP.Pedido)= " & DatDados(0).Recordset.Fields("Pedido") & " ) AND ((DUP.Cond)=False)) ORDER BY DUP.Pedido ")
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
                If VarMulta = True Then
                    If DatDados(0).Recordset.Fields("DtVenc") < (Date - 10) Then
'                        Prazo = Date - DatDados(0).Recordset.Fields("DtVenc")
                        DatDados(0).Recordset.Edit
                        DatDados(0).Recordset.Fields("Juros") = DatDados(0).Recordset.Fields("Valor") * ((DatDados(1).Recordset("Juros") / 100))
                        DatDados(0).Recordset.Update
                        TotJuros = TotJuros + DatDados(0).Recordset("Juros")
                        If DatDados(0).Recordset.Fields("DtVenc") < Date Then
                            Prazo = Date - DatDados(0).Recordset.Fields("DtVenc")
                            DatDados(0).Recordset.Edit
                            DatDados(0).Recordset.Fields("Juros") = DatDados(0).Recordset.Fields("Juros") + (DatDados(0).Recordset.Fields("Valor") * (((DatDados(1).Recordset("Juros") / 30) / 100) * Prazo))
                            DatDados(0).Recordset.Update
                            TotJuros = TotJuros + DatDados(0).Recordset("Juros")
                        End If
                    
                    End If
                Else
                    If DatDados(0).Recordset.Fields("DtVenc") < Date Then
                        Prazo = Date - DatDados(0).Recordset.Fields("DtVenc")
                        DatDados(0).Recordset.Edit
                        DatDados(0).Recordset.Fields("Juros") = DatDados(0).Recordset.Fields("Valor") * (((DatDados(1).Recordset("Juros") / 30) / 100) * Prazo)
                        If Len(Trim(DatDados(0).Recordset.Fields("Obs"))) = 0 Or IsNull(DatDados(0).Recordset.Fields("Obs")) Or Left(DatDados(0).Recordset.Fields("Obs"), 2) = "R$" Then
                            DatDados(0).Recordset.Fields("Obs") = Format(DatDados(0).Recordset.Fields("Valor") + DatDados(0).Recordset.Fields("Valor") * (((DatDados(1).Recordset("Juros") / 30) / 100) * Prazo), "R$  ##0.00")
                        End If
                        DatDados(0).Recordset.Update
                        TotJuros = TotJuros + DatDados(0).Recordset("Juros")
                    End If
                End If
            Else
                    DatDados(0).Recordset.Edit
                    DatDados(0).Recordset.Fields("Juros") = 0
                    If Left(DatDados(0).Recordset.Fields("Obs"), 2) = "R$" Then
                        DatDados(0).Recordset.Fields("Obs") = ""
                    End If
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
            
            If VarCustoPed = True Then
'                Sql = "SELECT ITP.Produto, Prd.Custo, ITP.Qtd, ITP.ValorUnit, ITP.VlrVdaDia, ITP.Pedido, ITP.CodItp FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido =  " & Datdados(0).Recordset("Pedido") & " ORDER BY ITP.CodItp"
                SQL = "SELECT ITP.Pedido, Avg((( [itp].[vlrvdadia]/ iif([ITP].[custoDIA]*[itp].[qtd] =0,[itp].[vlrvdadia],([ITP].[custoDIA]*[itp].[qtd]))-1)*100)) AS Expr1 FROM ITP GROUP BY ITP.Pedido HAVING Itp.Pedido <> 0 and ITP.Pedido = " & DatDados(0).Recordset("Pedido")
                Set SqlPed = Banco.OpenRecordset(SQL)
                SqlPed.Requery
                If SqlPed.RecordCount > 0 Then
                    VarMedia = VarMedia + SqlPed!Expr1
                    VarCont = VarCont + 1
                End If
            End If
                        
        End If
        If Not IsNull(DatDados(0).Recordset.Fields("Saldo")) Then
            Saldo = Saldo + DatDados(0).Recordset.Fields("Saldo")
        End If
        DatDados(0).Recordset.Edit
        DatDados(0).Recordset("PgtTemp") = 0
        DatDados(0).Recordset.Update
        DatDados(0).Recordset.MoveNext
    Loop
    If VarCustoPed = True Then
        If VarCont > 0 Then
            Label7 = "Margem:  " & Format(VarMedia / VarCont, "##0.00") & " %"
        End If
    End If
    MebPagos.Text = Pago + Saldo
    MebAPagar.Text = Apagar - Saldo
    MebSubTot = Pago + Apagar
    MebTotJuros = TotJuros
    MebTotDesc = TotDesc
    MebTotGeral = (MebSubTot + CCur(TotJuros)) - CCur(TotDesc)
    DatDados(0).Recordset.MoveFirst
End If
Frame.Visible = False
End Sub

Private Sub MebJuros_LostFocus()
If MebTotal <> "" And MebJuros <> "" And IsNumeric(MebJuros) Then
    MebTotal = CCur(MebTotal) + CCur(MebJuros)
End If
MebPago = CCur(MebTotal)

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

If MsgBox("Imprimir só os marcados ?", vbYesNo) = vbYes Then
    Set VarDup = Banco.OpenRecordset("SELECT * FROM Dup  WHERE Entidade = " & Val(TxtCod) & " and Cond = False and  Dup.PgtTemp = true  ORDER BY DtVenc, DtEmiss")
    VarDup.Requery
Else
    Set VarDup = Banco.OpenRecordset("SELECT * FROM Dup  WHERE Entidade = " & Val(TxtCod) & " and Cond = False ORDER BY DtVenc, DtEmiss")
    VarDup.Requery
End If
'Set Arq = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 ORDER BY  Ped.CodPed ")

'Sql = "SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 ORDER BY  Ped.CodPed "

'Set Pedido = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido WHERE PED.Entidade = " & CodEnt & " Order By CodPed ")

'Sql = "SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc, Ped.Operacao  FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 AND (Ped.Operacao = 1 or Ped.Operacao = 2 ) ORDER BY Ped.CodPed  "
Set Arq = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Ent.Endereco, Ent.Nro,  PED.CodPed, PED.DataEmiss, PED.Desconto, ITP.Qtd, ITP.VlrVdaDia, ITP.ValorUnit, ITP.Valor, ITP.Desconto, ITP.Produto, PRD.Descrição_Produto, Prd.PcoPrazo, Prd.Venda, PED.Entidade, DUP.NroPcl, DUP.CodDup, DUP.Cond, Dup.Juros, Dup.DtVenc, Ped.Operacao  FROM ((ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN DUP ON PED.CodPed = DUP.Pedido WHERE  PED.Entidade  = " & CodEnt & " AND DUP.Cond = False and  NroPcl = 1 AND (Ped.Operacao = 1 or Ped.Operacao = 2 ) ORDER BY Ped.CodPed  ")

If Arq.RecordCount > 0 Then
        Arq.MoveFirst
'        Iniciar (False)
        Open App.Path & "\RelDiscr.txt" For Output As #1
        Print #1, ""
        Print #1, "" & Cond
        Print #1, Gde & Centralizar(NomeEmpresa, 26) & Cond
'        Print #1, Cond & Centralizar(Empresa("Cabecalho2"), 48)
        Print #1, "================================================"
        Print #1, "Data: " & Date & " Hora: " & Time; Tab(33)
        Print #1, Tab(1); psCompString(1, Arq!CodEntidade, 6); "-"; Left(Arq!Nome, 30)
        Print #1, Tab(1); Trim(Left(Arq!Endereco, 30)); ", "; Arq!Nro
        Print #1, "------------------------------------------------"
        Print #1, Tab(1); "CODIGO"; Tab(8); "DESCRICAO DO PRODUTO"
        Print #1, Tab(1); "QTD"; Tab(20); "UNIT."; Tab(40); "TOTAL"
        Print #1, "------------------------------------------------"
        SubTot = 0
        Desconto = 0
        TotJuros = 0
        Total = 0
        Juros = 0
        DescPed = 0
        VarDup.MoveFirst
        While Not VarDup.Eof
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
                   While Not Arq.Eof
                        If PedAnt = Arq!CodPed Then
                           If VarReajConta Then
                               If VarPcoPrazo = True Then
                                   If Arq!Qtd * Arq!PcoPrazo > Arq!Valor Then
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!PcoPrazo, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Qtd * Arq!PcoPrazo, "##,##0.00"), 10, 2)
                                        SubTot = SubTot + (Arq!Qtd * Arq!PcoPrazo)
                                        Total = Total + (Arq!Qtd * Arq!PcoPrazo)
                                   Else
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 10, 2)
                                        SubTot = SubTot + Arq!Valor
                                        Total = Total + Arq!Valor
                                   End If
                               Else
                                   If Arq!Qtd * Arq!Venda > Arq!Valor Then
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!Venda, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Venda * Arq!Qtd, "##,##0.00"), 10, 2)
                                        SubTot = SubTot + (Arq!Qtd * Arq!Venda)
                                        Total = Total + (Arq!Qtd * Arq!Venda)
                                   Else
                                        Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 10, 2)
                                        SubTot = SubTot + Arq!Valor
                                        Total = Total + Arq!Valor
                                   End If
                               End If
                           Else
                                If Arq!VlrVdaDia = 0 Then
                                    Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!ValorUnit, "##,##0.00"), 7, 2); Alinhar(Format(Arq!Valor, "##,##0.00"), 10, 2)
                                    SubTot = SubTot + Arq!Valor
                                    Total = Total + Arq!Valor
                                Else
                                    Print #1, Tab(1); Alinhar(Format(Arq!Qtd, " ##,##0.000"), 9, 3); Tab(11); psCompString(3, Left(Arq!Descrição_Produto, 17), 21); Alinhar(Format(Arq!VlrVdaDia / Arq!Qtd, "##,##0.00"), 7, 2); Alinhar(Format(Arq!VlrVdaDia, "##,##0.00"), 10, 2)
                                    SubTot = SubTot + Arq!VlrVdaDia
                                    Total = Total + Arq!VlrVdaDia
                                End If
                           End If
                           DtAnt = Arq!DtVenc
                           Juros = Arq!Juros
                       End If
                       Arq.MoveNext
                   Wend
                   If Juros <> 0 Then
                        TotJuros = TotJuros + Juros
                        Print #1, Tab(13); "   Atraso .... "; tab30; Date - DtAnt & " dias  "; Tab(39); Alinhar(CStr(Juros), 10, 2)
                   End If
                   If DescPed <> 0 Then
                        Print #1, Tab(31); "Desconto"; Alinhar(CStr(DescPed), 10, 2)
                   End If
                   Print #1, Tab(1); "======================================" & Alinhar(CStr(SubTot + Juros), 10, 2)
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
                    Print #1, Tab(1); "======================================" & Alinhar(CStr(SubTot + Juros), 10, 2)
                    SubTot = 0
                    Juros = 0
               
               End If
               VarDup.MoveNext
            Wend
            Print #1, ""
            Print #1, Tab(23); "SUB-TOTAL =>"; Tab(39); Alinhar(CStr(Total), 10, 2)
            Print #1, Tab(23); "ACRESCIMO =>"; Tab(39); Alinhar(CStr(TotJuros), 10, 2)
            Print #1, Tab(23); "DESCONTO  =>"; Tab(39); Alinhar(CStr(Desconto), 10, 2)
            Print #1, Tab(23); "TOTAL     =>"; Tab(39); Alinhar(CStr((Total + TotJuros) - Desconto), 10, 2)
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
            
            Fs2.CopyFile App.Path & "\RelDiscr.txt", Porta
        
        Else
            SSTab1.Tab = 2
            On Error GoTo Erro
            LstPedido.Clear
            Open App.Path & "\RelDiscr.txt" For Input As #1
            Do
                Line Input #1, Temp
                LstPedido.AddItem Temp
        '        Lista.Selected(Lista.ListCount - 1) = True
            Loop Until Temp = "<>"
            Close #1
        End If

Else
    MsgBox "Pedidos não se encontram nesse arquivo !", vbExclamation, App.Title
End If

Exit Sub
Erro:
If Err.Number = 62 Then
    Close #1
End If

End Sub


Private Sub MebTotal_Change()
If MebTotal = 0 Then
   CmdPgto.Enabled = False
   MebPago.Enabled = False
   BtBoleto.Enabled = False
Else
    If VarCont > 0 Then
        CmdPgto.Enabled = True
        MebPago.Enabled = True
        BtBoleto.Enabled = True
    Else
        CmdPgto.Enabled = False
        MebPago.Enabled = False
        BtBoleto.Enabled = False
        MebTotal = 0
        MebPago = 0
    End If
End If
End Sub

Private Sub SSTab1_Click(previoustab As Integer)
Dim SqlPed As Recordset
Dim lstPesq As ListItem

Dim SQL As String
Dim VarCodPed As Double
If SSTab1.Tab = 1 Then
    Lista.ListItems.Clear
    If IsNumeric(dbgrid1.Columns(1)) Then
        VarCodPed = dbgrid1.Columns(1)
    Else
        VarCodPed = 0
    End If
    
    SQL = "Select * from Ped WHERE Codped =  " & VarCodPed
    Set SqlTot = Banco.OpenRecordset(SQL)
    SqlTot.Requery
    If SqlTot.RecordCount > 0 Then
        MebTotalPed = SqlTot!Total
        If Not IsNull(SqlTot!Desconto) Then
            MebDescPed = SqlTot!Desconto
        End If
    End If
    
    SQL = "SELECT ITP.Produto, PRD.Descrição_Produto, PRD.Embalagem, ITP.Qtd, ITP.ValorUnit, ITP.VlrVdaDia, ITP.Pedido, ITP.CodItp FROM ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd WHERE ITP.Pedido =  " & VarCodPed & " ORDER BY ITP.CodItp"
    Set SqlPed = Banco.OpenRecordset(SQL)
    If SqlPed.RecordCount > 0 Then
        SqlPed.MoveFirst
        While Not SqlPed.Eof
            Set lstPesq = Lista.ListItems.Add(, , SqlPed("Produto"))
    
            lstPesq.SubItems(1) = psCompString(3, Left(SqlPed!Descrição_Produto, 40), 40)
            lstPesq.SubItems(2) = psCompString(3, Left(SqlPed!Embalagem, 10), 10)
            
            If Right(Format(Alinhar(SqlPed!Qtd, 10, 2), "##,###.00"), 2) = "00" Then
                lstPesq.SubItems(3) = Format(Alinhar(SqlPed!Qtd, 10, 2), "##,###,##0")
            Else
                lstPesq.SubItems(3) = Format(Alinhar(SqlPed!Qtd, 10, 2), "###,##0.0#")
            End If
            lstPesq.SubItems(4) = Alinhar(SqlPed!ValorUnit, 10, 3)
            lstPesq.SubItems(5) = Alinhar(SqlPed!VlrVdaDia, 12, 2)
            SqlPed.MoveNext
        Wend
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




