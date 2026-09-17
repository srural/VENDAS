VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.5#0"; "ReportX.ocx"
Begin VB.Form FrmRelPcl 
   Caption         =   "Relatório de Contas a Receber / Recebidas"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   75.935
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   3300
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      MargemEsquerda  =   3.21733665466309
      MargemDireita   =   3.55599522590637
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   1275
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2249
      Tipo            =   2
      Begin ReportX.ReportField rpCabTitulo 
         Height          =   330
         Left            =   510
         TabIndex        =   4
         Top             =   30
         Width           =   5685
         _ExtentX        =   10028
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
         Width           =   4995
         _ExtentX        =   8811
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
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   615
         TabIndex        =   6
         Top             =   930
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Caption         =   "Pedido"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   3555
         TabIndex        =   7
         Top             =   930
         Width           =   765
         _ExtentX        =   1349
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
      Begin ReportX.ReportField ReportField13 
         Height          =   210
         Left            =   6165
         TabIndex        =   8
         Top             =   930
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "VlrPago"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   9105
         TabIndex        =   9
         Top             =   930
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Obs."
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
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   7995
         TabIndex        =   10
         Top             =   930
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Tipo"
         TipoCampo       =   1
         Alignment       =   2
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
         Left            =   510
         TabIndex        =   11
         Top             =   360
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
         Left            =   600
         TabIndex        =   12
         Top             =   720
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
         Left            =   1530
         TabIndex        =   13
         Top             =   720
         Width           =   825
         _ExtentX        =   1455
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
         Left            =   4365
         TabIndex        =   14
         Top             =   720
         Width           =   825
         _ExtentX        =   1455
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
      Begin ReportX.ReportField ReportField20 
         Height          =   210
         Left            =   7635
         TabIndex        =   15
         Top             =   720
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         Caption         =   "Nro"
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
         Left            =   9870
         TabIndex        =   16
         Top             =   720
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
      Begin ReportX.ReportField ReportField22 
         Height          =   210
         Left            =   1275
         TabIndex        =   17
         Top             =   930
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   370
         Caption         =   "Pcl"
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
      Begin ReportX.ReportField ReportField23 
         Height          =   210
         Left            =   8430
         TabIndex        =   18
         Top             =   930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   370
         Caption         =   "Doc."
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
      Begin ReportX.ReportField ReportField24 
         Height          =   210
         Left            =   1680
         TabIndex        =   19
         Top             =   930
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Caption         =   "Dt. Emiss"
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
      Begin ReportX.ReportField ReportField25 
         Height          =   210
         Left            =   2580
         TabIndex        =   20
         Top             =   930
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   "Dt. Vencto"
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
         Left            =   7080
         TabIndex        =   42
         Top             =   930
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "DtPgto"
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
      Begin ReportX.ReportField ReportField16 
         Height          =   210
         Left            =   4380
         TabIndex        =   45
         Top             =   930
         Width           =   735
         _ExtentX        =   1296
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
      Begin ReportX.ReportField ReportField39 
         Height          =   210
         Left            =   5190
         TabIndex        =   46
         Top             =   930
         Width           =   885
         _ExtentX        =   1561
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
      Begin ReportX.ReportField ReportField45 
         Height          =   210
         Left            =   10095
         TabIndex        =   52
         Top             =   930
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Vendedor"
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
      Begin ReportX.ReportField ReportField46 
         Height          =   210
         Left            =   8415
         TabIndex        =   53
         Top             =   720
         Width           =   660
         _ExtentX        =   1164
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
         X1              =   510
         X2              =   11430
         Y1              =   660
         Y2              =   675
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   510
         X2              =   11430
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   210
      Index           =   2
      Left            =   0
      Top             =   1680
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   370
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   645
         TabIndex        =   24
         Top             =   0
         Width           =   525
         _ExtentX        =   926
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
         Left            =   3465
         TabIndex        =   25
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6135
         TabIndex        =   26
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Campo           =   "forVlrPago"
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
         Left            =   9120
         TabIndex        =   27
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "Obs"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
         Left            =   1185
         TabIndex        =   28
         Top             =   0
         Width           =   405
         _ExtentX        =   714
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
      Begin ReportX.ReportField ReportField29 
         Height          =   210
         Left            =   8460
         TabIndex        =   29
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   370
         Campo           =   "Documento"
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
         TabIndex        =   30
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
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
         Left            =   2565
         TabIndex        =   31
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
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
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   8100
         TabIndex        =   32
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   370
         Campo           =   "Tipo_Dup"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField37 
         Height          =   210
         Left            =   7095
         TabIndex        =   41
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   4350
         TabIndex        =   43
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
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
      Begin ReportX.ReportField txtcorrecao 
         Height          =   210
         Left            =   5220
         TabIndex        =   44
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
      Begin ReportX.ReportField ReportField44 
         Height          =   210
         Left            =   10080
         TabIndex        =   51
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "Vendedores.ENT.Nome"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
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
      Top             =   2610
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
         Left            =   600
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
         X2              =   570
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection Sec2 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   0
      Left            =   0
      Top             =   1275
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField ReportField11 
         Height          =   210
         Left            =   600
         TabIndex        =   21
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Ent.CodEntidade"
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
         Left            =   1530
         TabIndex        =   22
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   370
         Campo           =   "Ent.Nome"
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
         Left            =   4365
         TabIndex        =   23
         Top             =   120
         Width           =   3180
         _ExtentX        =   5609
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
         Left            =   7605
         TabIndex        =   33
         Top             =   120
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
         Left            =   9870
         TabIndex        =   34
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
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
      Begin ReportX.ReportField ReportField47 
         Height          =   210
         Left            =   8370
         TabIndex        =   54
         Top             =   120
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   370
         Campo           =   "Cidade"
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
         X1              =   510
         X2              =   11400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   510
         X2              =   11400
         Y1              =   90
         Y2              =   90
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   315
      Index           =   0
      Left            =   0
      Top             =   1890
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField ReportField27 
         Height          =   210
         Left            =   3120
         TabIndex        =   35
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   210
         Left            =   6015
         TabIndex        =   36
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "forTotalPgto"
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
      Begin ReportX.ReportField ReportField36 
         Height          =   210
         Left            =   2010
         TabIndex        =   40
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Caption         =   "SubTotal:"
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
      Begin ReportX.ReportField ReportField40 
         Height          =   210
         Left            =   5130
         TabIndex        =   47
         Top             =   60
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField41 
         Height          =   210
         Left            =   4200
         TabIndex        =   48
         Top             =   60
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line5 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   4200
         X2              =   5070
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   5130
         X2              =   6000
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line4 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   6060
         X2              =   6990
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line9 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   14  'Copy Pen
         X1              =   3120
         X2              =   4230
         Y1              =   270
         Y2              =   270
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   1
      Left            =   0
      Top             =   2205
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   714
      Tipo            =   6
      Ordem           =   1
      Begin ReportX.ReportField ReportField33 
         Height          =   210
         Left            =   3150
         TabIndex        =   37
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ReportX.ReportField ReportField34 
         Height          =   210
         Left            =   6045
         TabIndex        =   38
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   370
         Campo           =   "forTotalPgtoG"
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
         Left            =   2010
         TabIndex        =   39
         Top             =   0
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
         Left            =   5160
         TabIndex        =   49
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
         Left            =   4260
         TabIndex        =   50
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
         X1              =   570
         X2              =   11400
         Y1              =   270
         Y2              =   270
      End
   End
End
Attribute VB_Name = "FrmRelPcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalValor As Currency
Dim TotalValorG As Currency
Dim TotalValorPgto As Currency
Dim TotalValorPgtoG As Currency
Dim TotalJuros As Currency
Dim TotalGJuros As Currency
Dim TotalCorrecao As Currency
Dim TotalGCorr As Currency
Dim VarCorrecao As Currency
Dim VarTotCorrecao As Currency
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
    VarCorrecao = 0
    VarTotCorrecao = 0
    TotalJuros = 0
    TotalGJuros = 0
    TotalCorrecao = 0
    TotalGCorr = 0
    
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
            
        Case "forVlrPago"
            Valor = Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
            
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
'    VarCorrecao = Relatorio.Recordset("Valor") + Relatorio.Recordset("Dup.Juros")
    TotalValor = TotalValor + Relatorio.Recordset("Valor")
    TotalValorG = TotalValorG + Relatorio.Recordset("Valor")
    TotalValorPgto = TotalValorPgto + Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
    TotalValorPgtoG = TotalValorPgtoG + Relatorio.Recordset("VlrPago") + Relatorio.Recordset("Dup.Juros")
    TotalJuros = TotalJuros + Relatorio.Recordset("Dup.Juros")
    TotalGJuros = TotalGJuros + Relatorio.Recordset("Dup.Juros")
    TotalCorrecao = TotalCorrecao + VarTotCorrecao
    TotalGCorr = TotalGCorr + VarTotCorrecao
    
End Sub
