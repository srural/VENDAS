VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPclRes 
   Caption         =   "Relatório de Contas a Receber / Recebidas"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   69.85
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   90
      TabIndex        =   0
      Top             =   5280
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   1380
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2434
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   30
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   582
         Caption         =   ""
         TipoCampo       =   7
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpCabSubTitulo 
         Height          =   270
         Left            =   6300
         TabIndex        =   5
         Top             =   90
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   476
         Caption         =   ""
         TipoCampo       =   7
         Alignment       =   1
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   9915
         TabIndex        =   6
         Top             =   765
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Ultima Compra"
         TipoCampo       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpCabPeriodo 
         Height          =   225
         Left            =   60
         TabIndex        =   7
         Top             =   330
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   397
         Caption         =   ""
         TipoCampo       =   7
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   765
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   "Cod. Ent."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   210
         Left            =   960
         TabIndex        =   9
         Top             =   765
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   370
         Caption         =   "Nome"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   990
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   370
         Caption         =   "Endereço"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField21 
         Height          =   210
         Left            =   5085
         TabIndex        =   11
         Top             =   765
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   370
         Caption         =   "Fone"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   7050
         TabIndex        =   31
         Top             =   765
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Valor"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField34 
         Height          =   210
         Left            =   7950
         TabIndex        =   49
         Top             =   765
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Juros"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField36 
         Height          =   210
         Left            =   8880
         TabIndex        =   50
         Top             =   765
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "VlrCorrig."
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField38 
         Height          =   210
         Left            =   4680
         TabIndex        =   51
         Top             =   990
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   370
         Caption         =   "Cidade"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   90
         X2              =   11430
         Y1              =   1275
         Y2              =   1275
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   240
      Index           =   2
      Left            =   0
      Top             =   1785
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   165
         TabIndex        =   15
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Campo           =   "Pedido"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   6975
         TabIndex        =   16
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Campo           =   "Valor"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   210
         Left            =   4755
         TabIndex        =   17
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "VlrPago"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField28 
         Height          =   210
         Left            =   855
         TabIndex        =   18
         Top             =   0
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   370
         Campo           =   "NroPcl"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField30 
         Height          =   210
         Left            =   1635
         TabIndex        =   19
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "DtEmiss"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField31 
         Height          =   210
         Left            =   2715
         TabIndex        =   20
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   370
         Campo           =   "DtVenc"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField37 
         Height          =   210
         Left            =   5895
         TabIndex        =   25
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "DtPgto"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField txtcorrecao 
         Height          =   210
         Left            =   8880
         TabIndex        =   33
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "forCorrecao"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   7950
         TabIndex        =   34
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "Dup.Juros"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportSection Sec2 
         Height          =   240
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   423
         Mostrar         =   0   'False
         Begin ReportX.ReportField ReportField5 
            Height          =   210
            Left            =   165
            TabIndex        =   39
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   370
            Campo           =   "Pedido"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField13 
            Height          =   210
            Left            =   6975
            TabIndex        =   40
            Top             =   0
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   370
            Campo           =   "Valor"
            Formato         =   "Standard"
            Caption         =   ""
            TipoCampo       =   1
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField15 
            Height          =   210
            Left            =   4755
            TabIndex        =   41
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   370
            Campo           =   "VlrPago"
            Formato         =   "Standard"
            Caption         =   ""
            TipoCampo       =   1
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField20 
            Height          =   210
            Left            =   855
            TabIndex        =   42
            Top             =   0
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   370
            Campo           =   "NroPcl"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField22 
            Height          =   210
            Left            =   1635
            TabIndex        =   43
            Top             =   0
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   370
            Campo           =   "DtEmiss"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField23 
            Height          =   210
            Left            =   2715
            TabIndex        =   44
            Top             =   0
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   370
            Campo           =   "DtVenc"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField24 
            Height          =   210
            Left            =   5895
            TabIndex        =   45
            Top             =   0
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
            Campo           =   "DtPgto"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField25 
            Height          =   210
            Left            =   8880
            TabIndex        =   46
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   370
            Campo           =   "forCorrecao"
            Formato         =   "Standard"
            Caption         =   ""
            TipoCampo       =   1
            Formula         =   -1  'True
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ReportX.ReportField ReportField29 
            Height          =   210
            Left            =   7950
            TabIndex        =   47
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   370
            Campo           =   "Dup.Juros"
            Formato         =   "Standard"
            Caption         =   ""
            TipoCampo       =   1
            Alignment       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   90
            X2              =   11430
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            DataField       =   "forSoma"
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
            Left            =   180
            TabIndex        =   48
            Top             =   630
            Width           =   45
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00E0E0E0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   210
            Left            =   10770
            Top             =   30
            Width           =   255
         End
      End
      Begin VB.Shape shpZebra 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   10770
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         DataField       =   "forSoma"
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
         Left            =   180
         TabIndex        =   1
         Top             =   630
         Width           =   45
      End
   End
   Begin ReportX.ReportSection Rod 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      Top             =   2820
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   953
      Tipo            =   7
      Begin ReportX.ReportField rpfRod 
         Height          =   210
         Left            =   8520
         TabIndex        =   2
         Top             =   60
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Campo           =   "=Página [Pagina]"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfCab 
         Height          =   210
         Index           =   7
         Left            =   255
         TabIndex        =   3
         Top             =   60
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   370
         Campo           =   "=Emitido em [Hoje] as [Hora]"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line lneRod 
         X1              =   10980
         X2              =   240
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   0
      Left            =   0
      Top             =   1380
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      Tipo            =   3
      Mostrar         =   0   'False
      Ordem           =   1
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "CodEntidade"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   930
         TabIndex        =   13
         Top             =   120
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   370
         Campo           =   "Nome"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField26 
         Height          =   210
         Left            =   3645
         TabIndex        =   14
         Top             =   120
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   370
         Campo           =   "Endereco"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   6390
         TabIndex        =   21
         Top             =   135
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   370
         Campo           =   "Nro"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   210
         Left            =   7155
         TabIndex        =   22
         Top             =   135
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         Campo           =   "Fone"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   11400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   11490
         Y1              =   90
         Y2              =   90
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   450
      Index           =   0
      Left            =   0
      Top             =   2025
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   794
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   6900
         TabIndex        =   26
         Top             =   210
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "forTotal"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   210
         Left            =   45
         TabIndex        =   27
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "forCodEnt"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   210
         Left            =   900
         TabIndex        =   28
         Top             =   0
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   370
         Campo           =   "forNome"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   210
         Left            =   885
         TabIndex        =   29
         Top             =   210
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   370
         Campo           =   "forEnd"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField39 
         Height          =   210
         Left            =   5085
         TabIndex        =   30
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "forFone"
         Formato         =   "(@@)@@@@-@@@@"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   4050
         TabIndex        =   32
         Top             =   210
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   370
         Campo           =   "Cidade"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField40 
         Height          =   210
         Left            =   8910
         TabIndex        =   35
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "forTotCorrecao"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField41 
         Height          =   210
         Left            =   7980
         TabIndex        =   36
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "forTotJuros"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField44 
         Height          =   210
         Left            =   9870
         TabIndex        =   52
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "forData"
         Formato         =   "Short Date"
         Caption         =   ""
         TipoCampo       =   2
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   11385
         Y1              =   435
         Y2              =   435
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   1
      Left            =   0
      Top             =   2475
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField ReportField33 
         Height          =   210
         Left            =   6930
         TabIndex        =   23
         Top             =   15
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "forTotalG"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField35 
         Height          =   210
         Left            =   5640
         TabIndex        =   24
         Top             =   15
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Caption         =   "Total Geral:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField42 
         Height          =   210
         Left            =   8910
         TabIndex        =   37
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "forTotGCorr"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField43 
         Height          =   210
         Left            =   7980
         TabIndex        =   38
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "forTotGJuros"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line6 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   60
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
   End
End
Attribute VB_Name = "FrmRelPclRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalValor As Currency
Dim TotalValorG As Currency
Dim TotalValorPgto As Currency
Dim TotalValorPgtoG As Currency
Dim VarCodEnt As String
Dim varNome As String
Dim VarEnd As String
Dim VarFone As String
Dim VarTotCorrecao As Currency
Dim TotalJuros As Currency
Dim TotalGJuros As Currency
Dim TotalCorrecao As Currency
Dim TotalGCorr As Currency
Dim VarData As String

'Private Rs As ADODB.Recordset

' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()

    'Dim Cn As String
        
    ' Uso do ADO nesse exemplo
    'Set Rs = New ADODB.Recordset
    
    ' Define a conexão. Pode ser usado uma conexao
    ' ja aberta pelo sistema.
    'Cn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\..\Exemplo.MDB"

    ' Abre o recordset com os dados
    'Rs.Open "SELECT * FROM ", Cn, adOpenStatic
    
    '  Associa o recordset ao relatorio
    'Set Relatorio.Recordset = Pesquisa(1)
    ' Inicia a geração do relatório.
    Relatorio.Ativar
    
    ' Fecha o recordset
    'Rs.Close
    'Set Rs = Nothing
    
    ' Retira o formulário de relatorio da memória
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    ' O componente trabalha no modo silencioso para
    ' erros. Ele dispara esse evento Erro e sai. Caso
    ' o seu relatório esteja iniciando e saindo sem
    ' apresentar erro, verifique se você colocou algum
    ' código nesse evento.

    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' Configura o tamanho ideal para o Shape que servirá para
    ' fazer um relatório do tipo "zebrado"
'    With shpZebra
'        .Left = 0
'        .Top = 0
'        .Width = Relatorio.LarguraPapel - Relatorio.MargemDireita - Relatorio.MargemEsquerda
'        .Height = Det.Height
'    End With
    'lneCab.X1 = 0
    'lneCab.X2 = Relatorio.LarguraUtil
    'lneRod.X1 = 0
    'lneRod.X2 = Relatorio.LarguraUtil
    
    TotalValor = 0
    TotalValorG = 0
    TotalValorPgto = 0
    TotalValorPgtoG = 0
    VarTotCorrecao = 0
    TotalJuros = 0
    TotalGJuros = 0
    TotalCorrecao = 0
    TotalGCorr = 0
    VarData = ""

End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        'shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
        shpZebra.Visible = False
    End If
       
End Sub

' Sub para apresentar mensagens de erro para o Visual ReportX
' Utilize sempre uma rotina no evento Erro do componente.
Private Sub Rpx_MsgErro(Numero As Long)

    Dim Msg$
    
    If Numero < 0 Then
    
        ' Mensagens de erro previstas
        Select Case Numero - vbObjectError
            Case 1001: Msg = "É necessário existir uma impressora instalada no Windows"
            Case 1002: Msg = "Não há registros a imprimir"
            Case 1003: Msg = "Não foi definida a seção de detalhe do relatório"
            Case 1004: Msg = "A configuração das seções de grupos está incorreta"
            Case 1005: Msg = "Foi definido um cursor do tipo Forward-Only para o recordset do relatório."
            Case 1006: Msg = "A página configurada para o relatório não possuí espaço suficiente para a impressão"
            Case 1007: Msg = "Já existe um relatório em andamento"
        End Select
        
        MsgBox Msg, vbInformation, "Impressão"
        
    Else
        
        ' Mensagens não previstas. Isso pode significar um erro
        ' interno no ReportX. Se isso acontecer, por favor reporte isso
        ' através de e-mail para ser corrigido.
        MsgBox "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impressão"
        
    End If
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        Case "forTotal"
            Valor = TotalValor
            TotalValor = 0
        
        Case "forTotalG"
            Valor = TotalValorG
        
        Case "forTotalPgto"
            Valor = TotalValorPgto
            TotalValorPgto = 0
    
        Case "forTotalPgtoG"
            Valor = TotalValorPgtoG
        Case "forNome"
            Valor = varNome
        Case "forCodEnt"
            Valor = VarCodEnt
        Case "forEnd"
            Valor = VarEnd
        Case "forFone"
            Valor = VarFone
        Case "forCorrecao"
            If rpCabTitulo.Tag = "1" Then
                Valor = Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
                VarTotCorrecao = Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
            Else
                Valor = Relatorio.Recordset("Valor") + Relatorio.Recordset("Dup.Juros")
                VarTotCorrecao = Relatorio.Recordset("Valor") + Relatorio.Recordset("Dup.Juros")
            End If
        Case "forTotJuros"
            Valor = TotalJuros
            TotalJuros = 0
        Case "forTotGJuros"
            Valor = TotalGJuros
        Case "forTotCorrecao"
            Valor = TotalCorrecao
            TotalCorrecao = 0
        Case "forTotGCorr"
            Valor = TotalGCorr
        Case "forData"
            Valor = VarData
        Case "forValor"
            If rpCabTitulo.Tag = "1" Then
                Valor = Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
            Else
                Valor = Relatorio.Recordset("Valor")
            End If
    End Select
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    ' Indica que o grupo irá quebrar a cada novo codigo
    ' do pedido. Verifica se é formula para o grupo 1 ou
    ' para o grupo 2. Nesse caso o grupo 2 nao tem quebra
    ' apenas o grupo 1.
    If Ordem = 1 Then
        Valor = Relatorio.Recordset("ent.CodEntidade")
    End If
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
    ' A cada impressão de registro na linha de detalhe,
    ' soma as qtdes e os totais para apresentar no rodapé
    ' do grupo
    VarEnd = ""
    VarFone = ""
    If rpCabTitulo.Tag = "1" Then
        VarTotCorrecao = Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
        TotalValor = TotalValor + Relatorio.Recordset("VlrPago")
        TotalValorG = TotalValorG + Relatorio.Recordset("VlrPago")
    Else
        VarTotCorrecao = Relatorio.Recordset("Valor") + Relatorio.Recordset("Dup.Juros")
        TotalValor = TotalValor + Relatorio.Recordset("Valor")
        TotalValorG = TotalValorG + Relatorio.Recordset("Valor")
               
    End If
    TotalValorPgto = TotalValorPgto + Relatorio.Recordset("VlrPago")
    TotalValorPgtoG = TotalValorPgtoG + Relatorio.Recordset("VlrPago")
    TotalJuros = TotalJuros + Relatorio.Recordset("Dup.Juros")
    TotalGJuros = TotalGJuros + Relatorio.Recordset("Dup.Juros")
    TotalCorrecao = TotalCorrecao + VarTotCorrecao
    TotalGCorr = TotalGCorr + VarTotCorrecao
        
    VarCodEnt = Relatorio.Recordset("Ent.CodEntidade")
    varNome = Relatorio.Recordset("Ent.Nome")
    If Not IsNull(Relatorio.Recordset("Endereco")) Or Not IsNull(Relatorio.Recordset("Nro")) Then
        VarEnd = Trim(Relatorio.Recordset("Endereco")) & ", " & Relatorio.Recordset("Nro")
    End If
    If Not IsNull(Relatorio.Recordset("Fone")) Then
        VarFone = Relatorio.Recordset("Fone")
    End If
    If Not IsNull(Relatorio.Recordset("DtEmiss")) Then
        VarData = Relatorio.Recordset("DtEmiss")
    End If
    
End Sub

