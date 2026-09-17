VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form FrmRelReciboProdutor 
   Caption         =   "Relatório de CFO"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   202.671
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   336.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   150
      TabIndex        =   0
      Top             =   6150
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   4920
      Left            =   0
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   8678
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField rpExtenso 
         Height          =   270
         Left            =   150
         TabIndex        =   5
         Top             =   3090
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   476
         Campo           =   "forExtenso"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField rpCpf 
         Height          =   270
         Left            =   660
         TabIndex        =   4
         Top             =   2250
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   476
         Campo           =   "forCpf"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField rpNome 
         Height          =   270
         Left            =   660
         TabIndex        =   3
         Top             =   1920
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
         Caption         =   ""
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
      Begin ReportX.ReportField rpTotal 
         Height          =   330
         Left            =   8940
         TabIndex        =   2
         Top             =   1350
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         Campo           =   "forTotal"
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   1
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
      Begin ReportX.ReportField ReportField2 
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   1920
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   450
         Caption         =   "Eu:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   2250
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   450
         Caption         =   "Cpf:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   2790
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         Caption         =   "declaro que recebi nesta, data a importância acima de:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpTitulo 
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   390
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   1085
         Caption         =   "RECIBO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   510
         TabIndex        =   9
         Top             =   3930
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   476
         Campo           =   "forEmpresa"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField12 
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   3930
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "da"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   270
         Left            =   5100
         TabIndex        =   11
         Top             =   3930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   476
         Campo           =   "forCpf"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField8 
         Height          =   255
         Left            =   4350
         TabIndex        =   12
         Top             =   3930
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   450
         Caption         =   "CNPJ:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   255
         Left            =   7200
         TabIndex        =   17
         Top             =   3660
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "Referente pagamento das seguintes notas:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField36 
         Height          =   270
         Left            =   570
         TabIndex        =   35
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Nro. da Nota"
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
      Begin ReportX.ReportField ReportField37 
         Height          =   270
         Left            =   2370
         TabIndex        =   36
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Cod. Pedido"
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
      Begin ReportX.ReportField ReportField38 
         Height          =   270
         Left            =   4230
         TabIndex        =   37
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Data"
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
      Begin ReportX.ReportField ReportField39 
         Height          =   270
         Left            =   6240
         TabIndex        =   38
         Top             =   4500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Caption         =   "Total"
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
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   180
         X2              =   11310
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillColor       =   &H00CECECE&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   8820
         Top             =   1230
         Width           =   2625
      End
      Begin VB.Line Line10 
         BorderStyle     =   2  'Dash
         X1              =   180
         X2              =   11310
         Y1              =   4470
         Y2              =   4470
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   2175
      Index           =   1
      Left            =   0
      Top             =   4920
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   3836
      Tipo            =   5
      QuebraDepois    =   -1  'True
      Ordem           =   1
      Begin ReportX.ReportField ReportField17 
         Height          =   255
         Left            =   7140
         TabIndex        =   18
         Top             =   1110
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "Por ser verdade, assino e dou fé."
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   255
         Left            =   7155
         TabIndex        =   42
         Top             =   1485
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "____/_____/_________"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line6 
         BorderStyle     =   2  'Dash
         X1              =   180
         X2              =   11310
         Y1              =   60
         Y2              =   60
      End
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   11310
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   609
      Ordem           =   1
      Begin ReportX.ReportField ReportField13 
         Height          =   240
         Left            =   570
         TabIndex        =   13
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "NroNt"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   240
         Left            =   2370
         TabIndex        =   14
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "CodPed"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   240
         Left            =   4230
         TabIndex        =   15
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "DtFat"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   240
         Left            =   6210
         TabIndex        =   16
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "calTotal"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ReportX.ReportSection cabFormiga 
      Align           =   1  'Align Top
      Height          =   4215
      Left            =   0
      Top             =   7095
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   7435
      Tipo            =   1
      Mostrar         =   0   'False
      QuebraDepois    =   -1  'True
      Ordem           =   1
      Begin ReportX.ReportField titExtenso 
         Height          =   270
         Left            =   3150
         TabIndex        =   19
         Top             =   1080
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   476
         Caption         =   ""
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
      Begin ReportX.ReportField titCpf 
         Height          =   270
         Left            =   6720
         TabIndex        =   20
         Top             =   780
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   476
         Caption         =   ""
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
      Begin ReportX.ReportField titNome 
         Height          =   270
         Left            =   600
         TabIndex        =   21
         Top             =   780
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
         Caption         =   ""
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
      Begin ReportX.ReportField titValor 
         Height          =   330
         Left            =   8940
         TabIndex        =   22
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField ReportField23 
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   780
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   450
         Caption         =   "Eu:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField24 
         Height          =   255
         Left            =   6150
         TabIndex        =   24
         Top             =   780
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   450
         Caption         =   "Cpf:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField25 
         Height          =   255
         Left            =   8970
         TabIndex        =   25
         Top             =   780
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   450
         Caption         =   "declaro que recebi nesta,"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField26 
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   1085
         Caption         =   "RECIBO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField27 
         Height          =   270
         Left            =   510
         TabIndex        =   27
         Top             =   1380
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   476
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField28 
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   1380
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "da"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField29 
         Height          =   270
         Left            =   5100
         TabIndex        =   29
         Top             =   1380
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   476
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField ReportField30 
         Height          =   255
         Left            =   4350
         TabIndex        =   30
         Top             =   1380
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   450
         Caption         =   "CNPJ:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField31 
         Height          =   255
         Left            =   7200
         TabIndex        =   31
         Top             =   1380
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "Referente pagamento dos fretes seguintes:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField32 
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   1080
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         Caption         =   "data a importância acima de:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField titNome2 
         Height          =   270
         Left            =   240
         TabIndex        =   33
         Top             =   3840
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField ReportField34 
         Height          =   255
         Left            =   7200
         TabIndex        =   34
         Top             =   2430
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "Por ser verdade, assino e dou fé."
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpObs 
         Height          =   1080
         Left            =   330
         TabIndex        =   39
         Top             =   1830
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1905
         Linhas          =   4
         Caption         =   ""
         TipoCampo       =   7
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
      Begin ReportX.ReportField ReportField3 
         Height          =   255
         Left            =   7200
         TabIndex        =   41
         Top             =   2880
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   "____/_____/_________"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   11460
         X2              =   30
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   30
         X2              =   30
         Y1              =   450
         Y2              =   4560
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   11460
         X2              =   11460
         Y1              =   510
         Y2              =   4560
      End
      Begin VB.Line Line11 
         X1              =   240
         X2              =   5730
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   30
         Top             =   60
         Width           =   8295
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillColor       =   &H00CECECE&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   8850
         Top             =   30
         Width           =   2625
      End
   End
   Begin ReportX.ReportSection RodGrupo 
      Align           =   1  'Align Top
      Height          =   795
      Index           =   0
      Left            =   0
      Top             =   11655
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   1402
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField rpNome2 
         Height          =   270
         Left            =   60
         TabIndex        =   40
         Top             =   300
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   476
         Caption         =   ""
         Formula         =   -1  'True
         Alignment       =   2
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
      Begin VB.Line Line2 
         X1              =   60
         X2              =   5550
         Y1              =   210
         Y2              =   210
      End
   End
End
Attribute VB_Name = "FrmRelReciboProdutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SomaValor As Currency
Public FunRural As Boolean
Public SqlCondicao As String
Public Tabela As String
Dim NomeAnt As String

Public Sub Config()

    
    rpNome.Campo = Tabela & ".Nome"
    rpNome2.Campo = "for" & Tabela & ".Nome"
    
    Relatorio.Ativar
        
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    Rpx_MsgErro Numero
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    SomaValor = 0
        
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    If Secao = secDetalhe Then
        'shpZebra.Visible = Not shpZebra.Visible
    ElseIf Secao = secCabecalho Then
        ' Inicia a primeira zebra de detalhe sempre igual
        ' em todas as folhas
        'shpZebra.Visible = False
    End If

End Sub

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
        
        MsgBox "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impressão"
        
    End If
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
    
        
    SomaValor = SomaValor + Relatorio.Recordset("Total")
    rpTotal.Caption = "R$ " & Format(SomaValor, "##,#0.00")
    
End Sub


Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    
    If Ordem = 1 Then
        Valor = Relatorio.Recordset(Tabela & ".Nome")
    End If
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)
Dim Completa As String
Dim valorTemp As Currency
Completa = " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
    Select Case Campo
        Case "forExtenso"
            NomeAnt = Relatorio.Recordset(Tabela & ".Nome")
            varPesquisa(2).MoveFirst
            varPesquisa(2).FindFirst IIf((Tabela = "Ent"), "Entidade", "Vendedor") & " = " & Relatorio.Recordset(Tabela & ".CodEntidade")
            
            If FunRural = True Then
                valorTemp = varPesquisa(2)("SomaDeTotal") - (varPesquisa(2)("SomaDeTotal") * "0,023")
            Else
                valorTemp = varPesquisa(2)("SomaDeTotal")
            End If
            
            Valor = "((" & Format(PassaExtenso(CStr(valorTemp)), ">") & "))" & Completa
        Case "forCpf"
            Valor = Format(Relatorio.Recordset(Tabela & ".Cpf"), "@@@.@@@.@@@-@@")
        Case "forEmpresa"
            Valor = NomeEmpresa
        Case "forTotal"
            varPesquisa(2).MoveFirst
            varPesquisa(2).FindFirst IIf((Tabela = "Ent"), "Entidade", "Vendedor") & " = " & Relatorio.Recordset(Tabela & ".CodEntidade")
            
            If FunRural = True Then
                valorTemp = varPesquisa(2)("SomaDeTotal") - (varPesquisa(2)("SomaDeTotal") * "0,023")
            Else
                valorTemp = varPesquisa(2)("SomaDeTotal")
            End If
            
            If varPesquisa(2).NoMatch = False Then
                Valor = "R$ " & Format(valorTemp, "#,#0.00")
            End If
        Case "forDataExtenso"
            Valor = DataExtenso(Date)
        Case "calTotal"
            If FunRural = True Then
                Valor = Relatorio.Recordset("Total") - (Relatorio.Recordset("Total") * "0,023")
            Else
                Valor = Relatorio.Recordset("Total")
            End If
        Case ("for" & Tabela & ".Nome")
            Valor = NomeAnt
    End Select
End Sub

