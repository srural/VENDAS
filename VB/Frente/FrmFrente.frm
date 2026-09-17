VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmFrente 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11010
   ClientLeft      =   -4635
   ClientTop       =   -855
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   32.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MskData 
      Height          =   285
      Left            =   11115
      TabIndex        =   275
      Top             =   7380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preço"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   11070
      TabIndex        =   272
      Top             =   6300
      Width           =   1140
      Begin VB.OptionButton OptPreco 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prazo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   274
         Top             =   495
         Width           =   915
      End
      Begin VB.OptionButton OptPreco 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   273
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.CommandButton CmdCli 
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14445
      TabIndex        =   271
      Top             =   4095
      Width           =   465
   End
   Begin VB.TextBox TxtSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14040
      TabIndex        =   270
      Top             =   3735
      Width           =   870
   End
   Begin VB.TextBox CliCartao 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -90
      TabIndex        =   59
      Top             =   10035
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton BtCusto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8820
      Width           =   465
   End
   Begin VB.Data DatMinutos 
      Caption         =   "Minutos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1620
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox TxtUltimo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15300
      TabIndex        =   35
      Top             =   1665
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox DbcLoja 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9045
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   8820
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9045
      ScaleHeight     =   345
      ScaleWidth      =   1560
      TabIndex        =   21
      Top             =   9570
      Width           =   1590
      Begin VB.TextBox TxtEstoque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.TextBox TxtComplemento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6030
      MaxLength       =   20
      TabIndex        =   20
      Top             =   8820
      Width           =   2610
   End
   Begin VB.TextBox TxtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5265
      TabIndex        =   19
      Top             =   9570
      Width           =   3630
   End
   Begin VB.TextBox TxtDados 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   8820
      Width           =   2535
   End
   Begin VB.TextBox TxtDados 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3015
      TabIndex        =   18
      Top             =   8820
      Width           =   2715
   End
   Begin VB.CommandButton BtAtu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Atualiza Tabelas"
      Top             =   9540
      Width           =   465
   End
   Begin VB.Timer Timer3 
      Interval        =   65000
      Left            =   11565
      Top             =   8775
   End
   Begin VB.CommandButton BtEcf 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUPOM / NFC-e  [F10]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8280
      Width           =   2580
   End
   Begin VB.CommandButton BtCancela 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCELA [F8]"
      DownPicture     =   "FrmFrente.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9585
      Width           =   2580
   End
   Begin VB.CommandButton BtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTALIZA [F9]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8925
      Width           =   2580
   End
   Begin VB.TextBox TxtTmp2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16155
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   15300
      Top             =   1125
   End
   Begin VB.TextBox TxtTmp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   16560
      TabIndex        =   12
      Top             =   1710
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Data DatVend 
      Caption         =   "Vendedor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   17100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade,Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
      Top             =   2205
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DatCli 
      Caption         =   "Cliente"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   17010
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FrmFrente.frx":1E53
      Top             =   1215
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6165
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde . . . "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   1080
         TabIndex        =   11
         Top             =   270
         Width           =   2880
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1058
      ButtonWidth     =   2884
      ButtonHeight    =   1005
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sair  [F11]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Configura Terminal"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Data e Hora"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "&Mesa"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Movimento Mesa"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Inicializa Senha"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixar &Conta [F5]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lançar Conta"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Preços [F6]"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emissão Pedido"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Consulta Pedido"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pe&didos [F7]"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ecf"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Leitura X"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Redução Z"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Abertura Gaveta"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C&aixa"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "acc_type"
                  Text            =   "Movimentação do Caixa"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "users"
                  Text            =   "Fechamento do Caixa"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Consulta Cheque"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Farmacia Popular"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "e-Pharma"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "VidaLink"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cancela Popular"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "SevenPdv"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Data DatLoja 
      Caption         =   "Loja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2340
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4950
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   17910
      Top             =   630
   End
   Begin VB.Data DatCompra 
      Caption         =   "Compra"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4380
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSComctlLib.ListView ListVda 
      Height          =   5790
      Left            =   225
      TabIndex        =   6
      Top             =   2520
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   10213
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Qtd"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CodPrd"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CodBar"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Produto"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Unitário"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CodItp"
         Object.Width           =   0
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   15705
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMask.MaskEdBox MebDesconto 
      Height          =   360
      Left            =   225
      TabIndex        =   24
      Top             =   9570
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebPreco 
      Height          =   360
      Left            =   1620
      TabIndex        =   25
      Top             =   9570
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Format          =   "#,##0.000;($#,##0.000)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebVlrTot 
      Height          =   360
      Left            =   3300
      TabIndex        =   26
      Top             =   9570
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSDBCtls.DBCombo DbcCliente 
      Bindings        =   "FrmFrente.frx":1EE2
      Height          =   345
      Left            =   10980
      TabIndex        =   36
      Top             =   4095
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
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
   Begin MSDBCtls.DBCombo DbcVendedor 
      Bindings        =   "FrmFrente.frx":1EF7
      Height          =   315
      Left            =   11025
      TabIndex        =   41
      Top             =   5805
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTdata 
      Height          =   330
      Left            =   13410
      TabIndex        =   42
      Top             =   5445
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   40304641
      CurrentDate     =   41051
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11475
      Top             =   7965
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin TabDlg.SSTab Frame5 
      Height          =   9825
      Left            =   15030
      TabIndex        =   64
      Top             =   585
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   17330
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "01"
      TabPicture(0)   =   "FrmFrente.frx":1F0D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Mesa(99)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Mesa(98)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Mesa(97)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Mesa(96)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Mesa(95)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Mesa(94)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Mesa(93)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Mesa(92)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Mesa(91)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Mesa(90)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Mesa(89)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Mesa(88)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Mesa(87)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Mesa(86)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Mesa(85)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Mesa(84)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Mesa(83)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Mesa(82)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Mesa(81)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Mesa(80)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Mesa(79)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Mesa(78)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Mesa(77)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Mesa(76)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Mesa(75)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Mesa(74)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Mesa(73)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Mesa(72)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Mesa(71)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Mesa(70)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Mesa(69)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Mesa(68)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Mesa(67)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Mesa(66)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Mesa(65)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Mesa(64)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Mesa(63)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Mesa(62)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Mesa(61)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Mesa(60)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Mesa(59)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Mesa(58)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Mesa(57)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Mesa(56)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Mesa(55)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Mesa(54)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Mesa(53)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Mesa(52)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Mesa(51)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Mesa(50)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Mesa(49)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Mesa(48)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Mesa(0)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Mesa(1)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Mesa(2)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Mesa(3)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Mesa(4)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Mesa(5)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Mesa(6)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Mesa(7)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Mesa(8)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Mesa(9)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Mesa(10)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Mesa(11)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Mesa(12)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Mesa(13)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Mesa(14)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Mesa(15)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Mesa(16)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Mesa(17)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Mesa(18)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Mesa(19)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Mesa(20)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Mesa(21)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Mesa(22)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "Mesa(23)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Mesa(24)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Mesa(25)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Mesa(26)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "Mesa(27)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "Mesa(28)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Mesa(29)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "Mesa(30)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "Mesa(31)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Mesa(32)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Mesa(33)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Mesa(34)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Mesa(35)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Mesa(36)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "Mesa(37)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "Mesa(38)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "Mesa(39)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "Mesa(40)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "Mesa(41)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "Mesa(42)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "Mesa(43)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "Mesa(44)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "Mesa(45)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "Mesa(46)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "Mesa(47)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "Mesa(100)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "Mesa(101)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).ControlCount=   102
      TabCaption(1)   =   "02"
      TabPicture(1)   =   "FrmFrente.frx":1F29
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Mesa(201)"
      Tab(1).Control(1)=   "Mesa(200)"
      Tab(1).Control(2)=   "Mesa(199)"
      Tab(1).Control(3)=   "Mesa(198)"
      Tab(1).Control(4)=   "Mesa(197)"
      Tab(1).Control(5)=   "Mesa(196)"
      Tab(1).Control(6)=   "Mesa(195)"
      Tab(1).Control(7)=   "Mesa(194)"
      Tab(1).Control(8)=   "Mesa(193)"
      Tab(1).Control(9)=   "Mesa(192)"
      Tab(1).Control(10)=   "Mesa(191)"
      Tab(1).Control(11)=   "Mesa(190)"
      Tab(1).Control(12)=   "Mesa(189)"
      Tab(1).Control(13)=   "Mesa(188)"
      Tab(1).Control(14)=   "Mesa(187)"
      Tab(1).Control(15)=   "Mesa(186)"
      Tab(1).Control(16)=   "Mesa(185)"
      Tab(1).Control(17)=   "Mesa(184)"
      Tab(1).Control(18)=   "Mesa(183)"
      Tab(1).Control(19)=   "Mesa(182)"
      Tab(1).Control(20)=   "Mesa(181)"
      Tab(1).Control(21)=   "Mesa(180)"
      Tab(1).Control(22)=   "Mesa(179)"
      Tab(1).Control(23)=   "Mesa(178)"
      Tab(1).Control(24)=   "Mesa(177)"
      Tab(1).Control(25)=   "Mesa(176)"
      Tab(1).Control(26)=   "Mesa(175)"
      Tab(1).Control(27)=   "Mesa(174)"
      Tab(1).Control(28)=   "Mesa(173)"
      Tab(1).Control(29)=   "Mesa(172)"
      Tab(1).Control(30)=   "Mesa(171)"
      Tab(1).Control(31)=   "Mesa(170)"
      Tab(1).Control(32)=   "Mesa(169)"
      Tab(1).Control(33)=   "Mesa(168)"
      Tab(1).Control(34)=   "Mesa(167)"
      Tab(1).Control(35)=   "Mesa(166)"
      Tab(1).Control(36)=   "Mesa(165)"
      Tab(1).Control(37)=   "Mesa(164)"
      Tab(1).Control(38)=   "Mesa(163)"
      Tab(1).Control(39)=   "Mesa(162)"
      Tab(1).Control(40)=   "Mesa(161)"
      Tab(1).Control(41)=   "Mesa(160)"
      Tab(1).Control(42)=   "Mesa(159)"
      Tab(1).Control(43)=   "Mesa(158)"
      Tab(1).Control(44)=   "Mesa(157)"
      Tab(1).Control(45)=   "Mesa(156)"
      Tab(1).Control(46)=   "Mesa(155)"
      Tab(1).Control(47)=   "Mesa(154)"
      Tab(1).Control(48)=   "Mesa(153)"
      Tab(1).Control(49)=   "Mesa(152)"
      Tab(1).Control(50)=   "Mesa(151)"
      Tab(1).Control(51)=   "Mesa(150)"
      Tab(1).Control(52)=   "Mesa(102)"
      Tab(1).Control(53)=   "Mesa(103)"
      Tab(1).Control(54)=   "Mesa(104)"
      Tab(1).Control(55)=   "Mesa(105)"
      Tab(1).Control(56)=   "Mesa(106)"
      Tab(1).Control(57)=   "Mesa(107)"
      Tab(1).Control(58)=   "Mesa(108)"
      Tab(1).Control(59)=   "Mesa(109)"
      Tab(1).Control(60)=   "Mesa(110)"
      Tab(1).Control(61)=   "Mesa(111)"
      Tab(1).Control(62)=   "Mesa(112)"
      Tab(1).Control(63)=   "Mesa(113)"
      Tab(1).Control(64)=   "Mesa(114)"
      Tab(1).Control(65)=   "Mesa(115)"
      Tab(1).Control(66)=   "Mesa(116)"
      Tab(1).Control(67)=   "Mesa(117)"
      Tab(1).Control(68)=   "Mesa(118)"
      Tab(1).Control(69)=   "Mesa(119)"
      Tab(1).Control(70)=   "Mesa(120)"
      Tab(1).Control(71)=   "Mesa(121)"
      Tab(1).Control(72)=   "Mesa(122)"
      Tab(1).Control(73)=   "Mesa(123)"
      Tab(1).Control(74)=   "Mesa(124)"
      Tab(1).Control(75)=   "Mesa(125)"
      Tab(1).Control(76)=   "Mesa(126)"
      Tab(1).Control(77)=   "Mesa(127)"
      Tab(1).Control(78)=   "Mesa(128)"
      Tab(1).Control(79)=   "Mesa(129)"
      Tab(1).Control(80)=   "Mesa(130)"
      Tab(1).Control(81)=   "Mesa(131)"
      Tab(1).Control(82)=   "Mesa(132)"
      Tab(1).Control(83)=   "Mesa(133)"
      Tab(1).Control(84)=   "Mesa(134)"
      Tab(1).Control(85)=   "Mesa(135)"
      Tab(1).Control(86)=   "Mesa(136)"
      Tab(1).Control(87)=   "Mesa(137)"
      Tab(1).Control(88)=   "Mesa(138)"
      Tab(1).Control(89)=   "Mesa(139)"
      Tab(1).Control(90)=   "Mesa(140)"
      Tab(1).Control(91)=   "Mesa(141)"
      Tab(1).Control(92)=   "Mesa(142)"
      Tab(1).Control(93)=   "Mesa(143)"
      Tab(1).Control(94)=   "Mesa(144)"
      Tab(1).Control(95)=   "Mesa(145)"
      Tab(1).Control(96)=   "Mesa(146)"
      Tab(1).Control(97)=   "Mesa(147)"
      Tab(1).Control(98)=   "Mesa(148)"
      Tab(1).Control(99)=   "Mesa(149)"
      Tab(1).Control(100)=   "Mesa(202)"
      Tab(1).Control(101)=   "Mesa(203)"
      Tab(1).ControlCount=   102
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   101
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   268
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   100
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   47
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   266
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   46
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   265
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   45
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   44
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   263
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   43
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   262
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   42
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   261
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   41
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   260
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   40
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   259
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   39
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   258
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   38
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   257
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   37
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   256
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   36
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   35
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   34
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   33
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   252
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   32
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   31
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   250
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   30
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   249
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   29
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   248
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   28
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   247
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   27
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   26
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   25
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   244
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   24
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   23
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   242
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   21
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   20
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   19
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   18
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   17
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   16
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   15
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   14
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   13
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   232
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   11
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   10
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   48
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   49
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   50
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   51
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   52
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   53
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   54
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   55
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   56
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   57
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   58
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   59
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   60
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   61
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   62
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   63
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   64
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   65
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   66
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   67
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   68
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   69
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   70
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   71
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   72
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   73
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   74
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   75
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   76
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   77
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   78
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   79
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   80
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   81
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   82
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   83
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   84
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   85
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   86
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   87
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   88
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   89
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   90
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   91
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   92
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   93
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   94
         Left            =   2625
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   95
         Left            =   3225
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   96
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   97
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   98
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   99
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   203
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   202
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   149
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   148
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   147
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   146
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   145
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   144
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   4320
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   143
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   142
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   141
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   140
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   139
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   138
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   3780
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   137
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   136
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   135
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   134
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   133
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   132
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   131
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   130
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   129
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   128
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   127
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   126
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   2700
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   125
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   124
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   123
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   122
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   121
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   120
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   2160
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   119
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   118
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   117
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   116
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   115
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   114
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   1620
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   113
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   112
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   111
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   110
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   109
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   108
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1080
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   107
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   106
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   105
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   104
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   103
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   102
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   150
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   151
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   152
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   153
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   154
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   155
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   4860
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   156
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   157
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   158
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   159
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   160
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   161
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   5400
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   162
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   163
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   164
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   165
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   166
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   167
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   5940
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   168
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   169
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   170
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   171
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   172
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   173
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   6480
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   174
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   175
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   176
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   177
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   178
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   179
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   7020
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   180
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   181
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   182
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   183
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   184
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   185
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   7560
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   186
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   187
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   188
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   189
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   190
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   191
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   8100
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   192
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   193
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   194
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   195
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   196
         Left            =   -72375
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   197
         Left            =   -71775
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   8640
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   198
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   199
         Left            =   -74175
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   200
         Left            =   -73575
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   9180
         Width           =   555
      End
      Begin VB.CommandButton Mesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   201
         Left            =   -72975
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   9180
         Width           =   555
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa/Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   12645
      TabIndex        =   269
      Top             =   3825
      Width           =   1320
   End
   Begin VB.Label LblNroCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7695
      TabIndex        =   63
      Top             =   10035
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label LblNsu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      TabIndex        =   62
      Top             =   10035
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label LblBPms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2700
      TabIndex        =   61
      Top             =   10080
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label LblCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   225
      TabIndex        =   60
      Top             =   10035
      Width           =   2265
   End
   Begin VB.Label LblCredito 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13365
      TabIndex        =   58
      Top             =   6300
      Width           =   45
   End
   Begin VB.Label LblAberto 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13365
      TabIndex        =   57
      Top             =   6630
      Width           =   45
   End
   Begin VB.Label LblSaldo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13365
      TabIndex        =   56
      Top             =   7260
      Width           =   45
   End
   Begin VB.Label LblPrazo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13365
      TabIndex        =   55
      Top             =   6945
      Width           =   45
   End
   Begin VB.Label LblDias 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13365
      TabIndex        =   54
      Top             =   7575
      Width           =   45
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12660
      TabIndex        =   53
      Top             =   7260
      Width           =   585
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12645
      TabIndex        =   52
      Top             =   7575
      Width           =   600
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12555
      TabIndex        =   51
      Top             =   6945
      Width           =   690
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aberto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12555
      TabIndex        =   50
      Top             =   6630
      Width           =   690
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12555
      TabIndex        =   49
      Top             =   6300
      Width           =   690
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   14625
      Shape           =   3  'Circle
      Top             =   6300
      Width           =   285
   End
   Begin VB.Label MebUnit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   12060
      TabIndex        =   48
      Top             =   765
      Width           =   2895
   End
   Begin VB.Label LblDescricao 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1680
      Left            =   270
      TabIndex        =   47
      Top             =   720
      Width           =   10410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   10845
      TabIndex        =   46
      Top             =   2385
      Width           =   690
   End
   Begin VB.Label MebTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   54.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1185
      Left            =   10800
      TabIndex        =   45
      Top             =   2475
      Width           =   4155
   End
   Begin VB.Shape Shape2 
      Height          =   960
      Left            =   10980
      Top             =   4410
      Width           =   3930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   11025
      TabIndex        =   44
      Top             =   5580
      Width           =   1035
   End
   Begin VB.Label totlista 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   43
      Top             =   7965
      Width           =   555
   End
   Begin VB.Label LblEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10980
      TabIndex        =   39
      Top             =   4500
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Label LblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10980
      TabIndex        =   38
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label LblCidade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10980
      TabIndex        =   37
      Top             =   4665
      Width           =   3915
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9045
      TabIndex        =   34
      Top             =   8580
      Width           =   1590
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9045
      TabIndex        =   33
      Top             =   9300
      Width           =   870
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   6015
      TabIndex        =   32
      Top             =   8580
      Width           =   2310
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Obs Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5265
      TabIndex        =   31
      Top             =   9300
      Width           =   1305
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3255
      TabIndex        =   30
      Top             =   9315
      Width           =   555
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unitário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1635
      TabIndex        =   29
      Top             =   9315
      Width           =   825
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.(%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   225
      TabIndex        =   28
      Top             =   9300
      Width           =   960
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd/Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   225
      TabIndex        =   27
      Top             =   8550
      Width           =   1200
   End
   Begin VB.Label LblMaquina 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "maquina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   14535
      TabIndex        =   9
      Tag             =   "|"
      Top             =   8820
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   16065
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   14175
      TabIndex        =   5
      Top             =   10170
      Width           =   75
   End
   Begin VB.Label LblItens 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   14310
      TabIndex        =   4
      Top             =   10170
      Width           =   750
   End
   Begin VB.Label LblData 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   14535
      TabIndex        =   3
      Top             =   8415
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblPedido 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13320
      TabIndex        =   2
      Top             =   10170
      Width           =   795
   End
   Begin VB.Label Mens 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   10980
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   9960
      Left            =   45
      Stretch         =   -1  'True
      Top             =   540
      Width           =   19290
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   10545
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   19050
   End
End
Attribute VB_Name = "FrmFrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Banco As database
Public Formulario As Form
Public Registros As Recordset
Public Clientes As Recordset
Public Pedido As Recordset
Public RecVendas As Recordset
Public Ofertas As Recordset
Public Empresa As Recordset
Public Loja As Recordset
Public CLoja As Recordset
Public Linha As Integer
Public ContPro As Integer
Public Soma As Currency
Public Caminho As String
Public CodCli As Integer
Public Flag2 As Boolean
Public Maquina As String
Public NroCupom As String
Public Itens As Integer
Public CodTotalizador As Long
Public Vendedor As Integer
Public Indices As Recordset
Public Valor As Currency
Public quantidade As Currency
Public Tempo As Byte
Public LstPesq As ListItem
Public VarItem As String
Public ContList As Integer
Public Senha As Boolean
Public RecSenha As Recordset
Public VarSomadeCusto As Currency
Public DatSenha As Recordset
Public CodCard  As Double
Public VarMesa As Recordset
Public VarGelada As String
Public PrecoAnt As Currency
Public VarLista1 As Recordset
Public Barra As String
Public VarDiversos As Currency
Public VarBloqueio As Boolean



Private Sub Bomba_Click(Index As Integer)
MsgBox "Teste", vbInformation
End Sub

Private Sub BtAtu_Click()
If CodCli = 0 Then
        Loja.Requery
        Clientes.Requery
        Registros.Requery
        DatCli.Recordset.Requery
End If
End Sub

Private Sub BtCancela_Click()
TxtDados_KeyDown 0, 119, 0
End Sub


Private Sub BtCusto_Click()

MsgBox "Custo: R$ " & Format(VarSomadeCusto, " ##,##0.00") & Chr(13) & "Margem: " & Format(((MebTotal / VarSomadeCusto) - 1) * 100, " ##,##0.00") & " %", vbInformation, App.Title


End Sub


Private Sub BtEcf_Click()
TxtDados_KeyDown 0, 120, 1
End Sub

Private Sub BtTotal_Click()
TxtDados_KeyDown 0, 120, 0
End Sub

Private Sub DbgCompra_dblClick()
Dim Criterio As String
Dim Resp As Byte
Dim Flag As Boolean
Resp = 7
Flag = True
If DatCompra.Recordset.RecordCount > 0 Then
    If EditarFrente = True Then
        Resp = MsgBox("Sim para Alterar, Não para Excluir", vbYesNoCancel + vbDefaultButton3, App.Title)
        If Resp = 6 Then
            Flag = False
'            Pic1.Visible = False
'            LblItens.Caption = DatCompra.Recordset.Fields("CodItp")
            
            Criterio = "CodItp = " & DatCompra.Recordset.Fields("CodItp")
            RecVendas.FindFirst Criterio
            If Not RecVendas.NoMatch Then
                FormEditar.LblProduto.Caption = DbgCompra.Columns(2)
                FormEditar.MebQtd = RecVendas("Qtd")
                FormEditar.MebDesc = RecVendas("Desconto")
                FormEditar.MebUnitario = RecVendas("ValorUnit")
                FormEditar.MebFina = RecVendas("VlrVdaDia")
                FormEditar.MebTot = RecVendas("Valor")
                Soma = CCur(Soma) - CCur(RecVendas("VlrVdaDia"))
                MebTotal = Format(Soma, "##,##0.00")
                FormEditar.Show vbModal
            End If
        End If
    End If
    If Flag = True And Resp = 7 Then
        If DatCompra.Recordset.RecordCount > 1 Then
            If MsgBox("Confirma exclusão do " & DbgCompra.Columns(2), vbYesNo, App.Title) = vbYes Then
                Soma = CCur(Soma) - DatCompra.Recordset("Valor")
                MebTotal = Format(Soma, "##,##0.00")
                DatCompra.Recordset.Delete
                DatCompra.Refresh
                If Not DatCompra.Recordset.EOF Then
                    DatCompra.Recordset.MoveLast
                End If
            End If
        Else
            MsgBox "Exclui o pedido inteiro", vbInformation, App.Title
        End If
    End If
End If

End Sub

Private Sub CmdPopular_Click()
If Soma <> 0 Then
    FormAutoriza.Show vbModal
Else
    MsgBox "Não tem venda pra autorizar !", vbInformation, App.Title
End If

End Sub

Private Sub CmdAmplacard_Click()

If Soma <> 0 Then
    FormAmpla.Show vbModal
Else
    MsgBox "Não tem venda pra autorizar !", vbInformation, App.Title
End If
End Sub


Private Sub CmdCli_Click()
FrmPesqCli.Show modal

End Sub

Private Sub DbcCliente_Change()
Dim Criterio As String

LblCodigo = DbcCliente.BoundText
If LblCodigo <> "" Then
    Criterio = "CodEntidade = " & LblCodigo
    DatCli.Recordset.FindFirst Criterio
    If Not IsNull(DatCli.Recordset("Endereco")) Then
        LblEnd = DatCli.Recordset("Endereco")
    End If
    If Not IsNull(DatCli.Recordset("Cidade")) Then
        LblCidade = DatCli.Recordset("Cidade")
    End If
    Clientes.FindFirst Criterio
    If LblCodigo = 1 Then
        Vista = True
    Else
        CodCli = LblCodigo
        Vista = False
    End If
End If
End Sub

Private Sub DbcCliente_LostFocus()
   CodCli = DbcCliente.BoundText

If Clientes!Status = 4 Then
    If VarBloqueio = False Then
         MsgBox "Cliente Bloqueado !  Você só pode vender a vista " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
        
            FormSenha = True
            If VarSenhaLibera <> "" Then
                VarTelaSenha = VarSenhaLibera
                FormSenha2.Show vbModal
            Else
                FormSenha = False
            End If
            If FormSenha = False Then
                 VarBloqueado = True
            End If
            Vista = True
     Else
        MsgBox "Cliente Bloqueado ! " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
        TxtDados(0).Text = ""
        TxtDados(2) = ""
        TxtDados(0).SetFocus
        CodCli = 0
        Exit Sub
    End If
            

End If
If CodCli <> 1 And Clientes.Fields("Prazo") <> "" Then
        Set Conta = Banco.OpenRecordset("SELECT DUP.* FROM DUP WHERE  DUP.Cond=False AND DUP.Entidade = " & CodCli & " Order By Dup.DtVenc ")
        Conta.Requery
        LblCredito = Format(Clientes!Credito, "#,##0.00;- #,##0.00")
        LblPrazo = Clientes!Prazo
        LblDias = 0
        LblSaldo = Format(0, "#,##0.00;- #,##0.00")
        LblAberto = 0
        Shape1.FillColor = &HFF00&
        If Conta.RecordCount <> 0 Then
            Dim SomaCta As Currency
            SomaCta = 0
            Conta.MoveFirst
            If Date > Conta!DtVenc Then
                LblDias = Date - Conta!DtVenc
                If IsNumeric(Clientes!Prazo) Then
                    If Conta("DtVenc") < (Date - (Clientes!Prazo * 2)) Then
                        Shape1.FillColor = &H80000008
                    ElseIf Conta("DtVenc") < (Date - Clientes!Prazo) Then
                            Shape1.FillColor = &HFF&
                        Else
                            Shape1.FillColor = &HFF00&
                    End If
                End If
            End If
            While Not Conta.EOF
                SomaCta = SomaCta + CCur(Conta!Valor)
                Conta.MoveNext
            Wend
        Else
            Shape1.FillColor = &HFF00&
        End If
        LblSaldo = Format(Clientes!Credito - SomaCta, "#,##0.00;- #,##0.00")
        LblAberto = Format(SomaCta, "#,##0.00;- #,##0.00")
        If Shape1.FillColor <> -2147483640 Then
            If SomaCta > (Clientes!Credito * 2) Then
                  Shape1.FillColor = &H80000008
            ElseIf SomaCta > Clientes!Credito Then
                   Shape1.FillColor = &HFF&
                Else
                    If Shape1.FillColor = &HFF00& Then
                        Shape1.FillColor = &HFF00&
                    End If
            End If
        End If
        Vista = False
Else
        Vista = True
End If
If Clientes!Status = 5 Then
'    If Shape1.FillColor = 255 Or Shape1.FillColor = -2147483640 Then
        If VarBloqueio = False Then
            MsgBox "Cliente Bloqueado !  Você só pode vender a vista " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
            
            
            FormSenha = True
            If VarSenhaLibera <> "" Then
                VarTelaSenha = VarSenhaLibera
                FormSenha2.Show vbModal
            Else
                FormSenha = False
            End If
            If FormSenha = False Then
                 VarBloqueado = True
            End If
         Else
            MsgBox "Cliente Bloqueado ! " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
            TxtDados(0).Text = ""
            TxtDados(2) = ""
            TxtDados(0).SetFocus
            CodCli = 0
            Exit Sub
        End If
            
'    End If
End If
VarDiaBom = 0

If Not IsNull(Clientes.Fields("Prazo")) Then
    If IsNumeric(Clientes.Fields("Prazo")) Then
        Dias = Clientes.Fields("Prazo")
        If Dias = "0" Then
            Dias = 30
        End If
    End If
End If

If Not IsNull(Clientes.Fields("DiaBomPgto")) Then
    If IsNumeric(Clientes.Fields("DiaBomPgto")) Then
        VarDiaBom = Clientes.Fields("DiaBomPgto")
    End If
End If
LblCodigo.Caption = Clientes.Fields("CodEntidade")
Criterio = "CodEntidade = " & LblCodigo
DatCli.Recordset.FindFirst Criterio
DbcCliente.Text = DatCli.Recordset("Nome")
'                            LblNome.Caption = Clientes.Fields("Nome")
If IsNull(Clientes.Fields("Endereco")) Then
    LblEnd.Caption = ""
Else
    LblEnd.Caption = Trim(Clientes.Fields("Endereco")) & ", " & Clientes.Fields("Nro")
End If
If IsNull(Clientes.Fields("Cidade")) Then
    LblCidade.Caption = ""
Else
    LblCidade.Caption = Clientes.Fields("Cidade")
End If
'                            If CodCli <> 1 Then
    If Not IsNull(Clientes!PesAut) Then
        FormVista.TxtPesAut = Clientes!PesAut
    End If
    If Not IsNull(Clientes!Obs) Then
        FormVista.TxtObs = Clientes!Obs
    End If
    If Dias = "" Then
        Dias = 30
    End If
    If CodCli <> 1 Then
        If VarDiaBom <> 0 Then
            If VarDiaBom <= 28 Then
                If Month(Date) + 1 > 12 Then
                    VarDtVenc = Format(VarDiaBom & "/01/" & Year(Date) + 1, "dd/mm/yyyy")
                Else
                    VarDtVenc = Format(VarDiaBom & "/" & Month(Date) + 1 & "/" & Year(Date), "dd/mm/yyyy")
                End If
            Else
                VarDtVenc = Format(Date + Dias, "dd/mm/yyyy")
            End If
        Else
            VarDtVenc = Format(Date + Dias, "dd/mm/yyyy")
        End If
    Else
        VarDtVenc = Date
    End If


End Sub

Private Sub DbcLoja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CfgDesc = "Sim" Or CfgDesc = 1 Then
         MebDesconto.SetFocus
'         Venda
    Else
        MebPreco = Round(CCur(Valor), 2)
        Venda
    End If
End If

End Sub

Private Sub DTdata_Change()
LblData = DtData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub


Private Sub ListVda_DblClick()
Dim Criterio As String
Dim Resp As Byte
Dim Flag As Boolean

Resp = 7
Flag = True

If ListVda.ListItems.Count > 0 Then

    If EditarFrente = True Then
        Resp = MsgBox("Sim para Alterar, Não para Excluir", vbYesNoCancel + vbDefaultButton3, App.Title)
        If Resp = 6 Then
            Flag = False
            VarItem = ListVda.SelectedItem.Index
            Criterio = "CodItp = " & ListVda.SelectedItem.SubItems(7)
            RecVendas.FindFirst Criterio
            If Not RecVendas.NoMatch Then
                FormEditar.LblProduto.Caption = ListVda.ListItems(CInt(VarItem)).ListSubItems(4).Text
                FormEditar.TxtBarra = ListVda.ListItems(CInt(VarItem)).ListSubItems(3).Text
                FormEditar.MebQtd = RecVendas("Qtd")
                FormEditar.MebDesc = RecVendas("Desconto")
                FormEditar.MebUnitario = RecVendas("ValorUnit")
                FormEditar.MebFina = RecVendas("VlrVdaDia")
                FormEditar.MebTot = RecVendas("Valor")
                FormEditar.txtLblItens = ListVda.SelectedItem.SubItems(7)
                Soma = CCur(Soma) - CCur(RecVendas("VlrVdaDia"))
                MebTotal = Format(Soma, "##,##0.00")
                FormEditar.Show vbModal
                Soma = SomaPedido(LblPedido)
                MebTotal = Format(Soma, "##,##0.00")
                
            End If
        End If
    End If
    If Flag = True And Resp = 7 Then
        If ListVda.ListItems.Count > 1 Then
            VarItem = ListVda.SelectedItem.Index
            If MsgBox("Confirma exclusão do " & ListVda.ListItems(CInt(VarItem)).ListSubItems(4).Text, vbYesNo, App.Title) = vbYes Then
                Criterio = "CodItp = " & ListVda.SelectedItem.SubItems(7)
                RecVendas.FindFirst Criterio
                If Not RecVendas.NoMatch Then
                    Soma = CCur(Soma) - CCur(RecVendas("VlrVdaDia"))  ' Format((ListVda.ListItems(CInt(VarItem)).ListSubItems(6).Text), "#,##0.000")
                    MebTotal = Format(Soma, "##,##0.00")
                    ListVda.ListItems.Remove CDec(VarItem)
'                    ListVda.ListItems(CInt(VarItem)).ListSubItems(3).Text = "(Cancelado)"
                    RecVendas.Edit
                    RecVendas.Delete
                    ContList = ContList - 1
'                    RecVendas.Update
                End If
                Soma = SomaPedido(LblPedido)
                MebTotal = Format(Soma, "##,##0.00")
            End If
        Else
            MsgBox "Exclui o pedido inteiro", vbInformation, App.Title
        End If
    End If
        
End If

End Sub

Private Sub MebDesconto_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10
End Sub

Private Sub MebDesconto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MebPreco.SetFocus
End If
End Sub

Private Sub MebDesconto_lostfocus()
If MebDesconto <> "" Then
    MebPreco = Round(CCur(Valor) - ((CCur(Valor) * (CCur(MebDesconto)) / 100)), 2)
Else
    MebDesconto.SetFocus
End If
End Sub

Private Sub MebPreco_GotFocus()
    PrecoAnt = MebPreco
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 10
End Sub

Private Sub MebPreco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MebPreco <> "" And IsNumeric(MebPreco) Then
        If LimiteDesc = 0 Then
            If MebPreco >= PrecoAnt Then
                Venda
            Else
                MsgBox "Valor nao pode ser menor que a tabela", vbInformation, App.Title
                TxtDados(0).SetFocus
            End If
        Else
            Venda
        End If
    Else
        TxtDados(0).SetFocus
    End If
End If
End Sub


Private Sub Mesa_Click(Index As Integer)
If Soma = 0 Then
    VarContMesa = Index
    Banco.Execute "UPDATE CONV INNER JOIN PRD ON CONV.Produto = PRD.CodPrd SET CONV.Valor = [prd].[venda] * Conv.Qtd  WHERE CONV.Totaliza=False AND CONV.Cliente = " & Trim(Mesa(VarContMesa).Caption)
    If VarConveniencia = True And VarLanchonete = True Then
            FrmLanca.TxtCartao = Trim(Mesa(VarContMesa).Caption)
            SendKeys ("{ENTER}")
            FrmLanca.Show vbModal
    ElseIf VarLanchonete = True Then
            FrmLanca.TxtCartao = Trim(Mesa(VarContMesa).Caption)
            SendKeys ("{ENTER}")
            FrmLanca.Show vbModal
        ElseIf VarConveniencia = True Then
            TxtDados(0).SetFocus
            TxtDados(0) = "/" & Trim(Mesa(VarContMesa).Caption)
            SendKeys ("{ENTER}")
    End If
Else
    MsgBox "Cupom nao totalizado", vbInformation, App.Title
End If

End Sub

Private Sub MSComm_OnComm()
Dim Teclado As String

Teclado = MSComm.Input

SendKeys Teclado

End Sub


Private Sub Form_Activate()
Set Formulário = Me
'Vista = True
MebPreco = 0
'MebDesconto = 0
MebVlrTot = 0
Mens = ""
'TxtDados(0).SetFocus
If Frame5.Visible = True And TxtTmp2 <> "" Then
    TxtDados(0).SetFocus
End If

'If Left(TxtDados(0).Text, 1) = "/" Then
'    SendKeys ("{ENTER}")
'End If

'Me.Left = 10 '(1024 - Me.Width) / 2
'Me.Top = 10 '(768 - Me.Height) / 2

'            If Registros.RecordCount = 1 And CodCli = 0 Then
'                If KeyAscii = 13 And TxtDados(0) = "" Then
'                    TxtDados(0).SetFocus
'                    TxtDados(0).Text = 1
'                    SendKeys ("{Enter}")
'               End If
'            End If



End Sub

Private Sub Form_Load()
Dim PesquisaSql As String
Dim Retorno As Integer
Dim Status As String
Dim P As Long
Dim DescImp As String

FlagFrame = True
VarPrdDesc = 0

If App.PrevInstance Then
    MsgBox "Existe uma aplicação do " & App.ProductName & " aberta nesta máquina, um ou mais programas abertos pode causar lentidão no sistema!", 16
    End
End If

Dim Criterio As String
CodTotalizador = 0
CodCli = 0
Tempo = 0

VarBaixaAuto = "Nao"
VarMostraCusto = "Nao"

On Error Resume Next

If dir(App.Path & "\Config.ini") <> "" Then
'        Open App.Path & "\Config.txt" For Input As #1
'        Input #1, LocalBanco
'        Input #1, Impressora
'        Input #1, PortaF
'        Input #1, PortaNF
'        Input #1, Leitor
'        Input #1, ConsultaCh
'        Input #1, CfgDesc
'        Input #1, Impr40
'        Input #1, NroLinha
'        Input #1, VarLocalEst
'        Input #1, VarMeiaFolha
'        Input #1, VarComplemento
'        Input #1, VarBaixaAuto
'        Input #1, VarMostraCusto
'        Close #1
VarDescFim = True
VarTamBarra = 0
VarLinhaDet = 0
VarComissao = 0
VarOrcamento = True
VarTelaPreco = True
VarCaixa = True
VarMostraEstoque = True
Gelado = False
VarBalanca = ""
VarCopias = 1
VarMercado = 0
VarImprCozinha = ""
VarImprBar = ""
VarEnter = 0
VarComprovante = 0
VarObsTotal = True
VarCaixaSeparado = False
VarPacote = False
VarCozinha = 0
VarVia = 1
VarBloqueio = False
VarRapido = 0

    LocalBanco = LerINI("Config", "LocalBco", App.Path & "\config.ini")
    Impressora = LerINI("Config", "Ecf", App.Path & "\config.ini")
    PortaF = LerINI("Config", "Porta", App.Path & "\config.ini")
    Leitor = LerINI("Config", "Leitor", App.Path & "\config.ini")
    ConsultaCh = LerINI("Config", "Consulta", App.Path & "\config.ini")
    VarMp4000 = LerINI("Config", "MP4000", App.Path & "\config.ini")
    PortaNF = LerINI("Config", "PortaNf", App.Path & "\config.ini")
            
    CfgDesc = LerINI("Config", "CfgDesc", App.Path & "\config.ini")
    NroLinha = LerINI("Config", "QtdLinhas", App.Path & "\config.ini")
    VarLocalEst = LerINI("Config", "LocalEst", App.Path & "\config.ini")
    VarMeiaFolha = LerINI("Config", "MeiaFolha", App.Path & "\config.ini")
    VarComplemento = LerINI("Config", "Complemento", App.Path & "\config.ini")
    VarBaixaAuto = LerINI("Config", "BaixaAuto", App.Path & "\config.ini")
    VarMostraCusto = LerINI("Config", "MostraCusto", App.Path & "\config.ini")
    VarImpressoraNf = LerINI("Config", "ImprNf", App.Path & "\config.ini")
    VarConveniencia = LerINI("Config", "Conveniencia", App.Path & "\config.ini")
    VarDescFim = LerINI("Config", "Desconto", App.Path & "\config.ini")
    VarLanchonete = LerINI("Config", "Lanchonete", App.Path & "\config.ini")
    VarTamBarra = LerINI("Config", "TamBarra", App.Path & "\config.ini")
    VarLinhaDet = LerINI("Config", "LinhaDet", App.Path & "\config.ini")
    VarComissao = LerINI("Config", "Comissao", App.Path & "\config.ini")
    VarOrcamento = LerINI("Config", "Orcamento", App.Path & "\config.ini")
    VarTelaPreco = LerINI("Config", "Preco", App.Path & "\config.ini")
    VarCaixa = LerINI("Config", "Caixa", App.Path & "\config.ini")
    VarMostraEstoque = LerINI("Config", "MostraEstoque", App.Path & "\config.ini")
    Gelado = LerINI("Config", "Gelado", App.Path & "\config.ini")
    VarBalanca = LerINI("Config", "Balanca", App.Path & "\config.ini")
    VarCopias = LerINI("Config", "Copias", App.Path & "\config.ini")
    VarCozinha = LerINI("Config", "ViaCozinha", App.Path & "\config.ini")
    VarVia = LerINI("Config", "SegundaVia", App.Path & "\config.ini")
    VarColombo = False
    VarColombo = LerINI("Config", "Colombo", App.Path & "\config.ini")

    VarObsTotal = LerINI("Config", "Obs", App.Path & "\config.ini")
    VarCaixaSeparado = LerINI("Config", "CaixaSeparado", App.Path & "\config.ini")
    
    VarBanco = LerINI("Boleto", "Banco", App.Path & "\config.ini")
    VarAgencia = LerINI("Boleto", "Agencia", App.Path & "\config.ini")
    VarConta = LerINI("Boleto", "Conta", App.Path & "\config.ini")
    VarConvenio = LerINI("Boleto", "Convenio", App.Path & "\config.ini")
    
    VarMercado = LerINI("Config", "Mercado", App.Path & "\config.ini")
    VarImprEntrega = LerINI("Impressora", "ImprEntrega", LocalBanco & "\config.ini")
    VarImprCozinha = LerINI("Config", "ImprCozinha", App.Path & "\config.ini")
    VarImprBar = LerINI("Config", "ImprBar", App.Path & "\config.ini")
    
    VarEnter = LerINI("Config", "Enter", App.Path & "\config.ini")
    VarComprovante = LerINI("Config", "Comprovante", App.Path & "\config.ini")
    VarPacote = LerINI("Config", "Pacote", App.Path & "\config.ini")
    VarBloqueio = LerINI("Config", "CliBloqueio", App.Path & "\config.ini")
    VarRapido = LerINI("Config", "CxaRapido", App.Path & "\config.ini")
    
        
    If VarTamBarra = 0 Then
        VarTamBarra = 6
    End If
   
    
    If LerINI("Sat", "CnpjDesenvolvedor", App.Path & "\config.ini") <> "" Then
        VarCnpjDesenvolvedor = LerINI("Sat", "CnpjDesenvolvedor", App.Path & "\config.ini")
        VarChave = LerINI2("Sat", "Chave", App.Path & "\config.ini")
'        VarChave2 = Mid$(LerINI("Sat", "Chave", App.Path & "\config.ini"), 201, 344)
        VarCaixaSat = LerINI("Sat", "NroCaixa", App.Path & "\config.ini")
        VarVersao = LerINI("Sat", "Versao", App.Path & "\config.ini")
        VarSerieNfce = LerINI("Nfce", "SerieNfce", App.Path & "\config.ini")
        VarNroNfce = LerINI("Nfce", "NroNfce", App.Path & "\config.ini")
    End If
        
    If LerINI("EMITENTE", "Cnpj", LocalBanco & "\nfe.ini") <> "" Then
        VarCnpjEmite = LerINI("EMITENTE", "Cnpj", LocalBanco & "\nfe.ini")
        VarUfEmite = LerINI("EMITENTE", "Uf", LocalBanco & "\nfe.ini")
        VarIbgeEmite = LerINI("EMITENTE", "CodigoIbge", LocalBanco & "\nfe.ini")
        VarRazaoEmite = LerINI("EMITENTE", "RazaoSocial", LocalBanco & "\nfe.ini")
        VarFantasiaEmite = LerINI("EMITENTE", "Fantasia", LocalBanco & "\nfe.ini")
        VarEndEmite = LerINI("EMITENTE", "Logradouro", LocalBanco & "\nfe.ini")
        VarNroEmite = LerINI("EMITENTE", "Nro", LocalBanco & "\nfe.ini")
        VarBairroEmite = LerINI("EMITENTE", "Bairro", LocalBanco & "\nfe.ini")
        VarCidadeEmite = LerINI("EMITENTE", "Cidade", LocalBanco & "\nfe.ini")
        VarInscEmite = LerINI("EMITENTE", "InscEst", LocalBanco & "\nfe.ini")
        VarCepEmite = LerINI("EMITENTE", "Cep", LocalBanco & "\nfe.ini")
    End If
    
    If LerINI("EMITENTE", "Pis", LocalBanco & "\nfe.ini") <> "" Then
        VarPis = LerINI("EMITENTE", "Pis", LocalBanco & "\nfe.ini")
    Else
        VarPis = ""
    End If
    
    If LerINI("EMITENTE", "Cofins", LocalBanco & "\nfe.ini") <> "" Then
        VarCofins = LerINI("EMITENTE", "Cofins", LocalBanco & "\nfe.ini")
    Else
        VarCofins = ""
    End If
    
    If LerINI("EMITENTE", "AliqPis", LocalBanco & "\nfe.ini") <> "" Then
        VarAliqPis = LerINI("EMITENTE", "AliqPis", LocalBanco & "\nfe.ini")
    Else
        VarAliqPis = ""
    End If
    
    If LerINI("EMITENTE", "AliqCofins", LocalBanco & "\nfe.ini") <> "" Then
        VarAliqCofins = LerINI("EMITENTE", "AliqCofins", LocalBanco & "\nfe.ini")
    Else
        VarAliqCofins = ""
    End If
    
                
    If LerINI("EMITENTE", "InscrMuni", App.Path & "\nfe.ini") <> "" Then
        VarInscrMuni = LerINI("EMITENTE", "InscrMuni", App.Path & "\nfe.ini")
    Else
        VarInscrMuni = Space(15)
    End If
    VarRegime = 1
    
    If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") <> "" Then
        If LerINI("EMITENTE", "RegimeTrib", LocalBanco & "\nfe.ini") = 3 Then
            VarRegime = 3
        Else
            VarRegime = 1
        End If
    End If
    
    
    Select Case VarImpressoraNf
        Case 1
            VarImpressoraNf = "EPSON"
        Case 2
            VarImpressoraNf = "CUPOM"
        Case 3
            VarImpressoraNf = "MECAF"
        Case 4
            VarImpressoraNf = "TERMICA"
     End Select
            
'    LblLocal.Caption = LocalBanco
    
'    If CfgDesc = 1 Then
'        CfgDesc = "Sim"
'    Else
'        CfgDesc = "Nao"
'    End If
            
    If VarMeiaFolha = 1 Then
        VarMeiaFolha = "Sim"
    Else
        VarMeiaFolha = "Nao"
    End If
    
    If VarComplemento = 1 Then
        VarComplemento = "Sim"
    Else
        VarComplemento = "Nao"
    End If
    
    If VarBaixaAuto = 1 Then
        VarBaixaAuto = "Sim"
    Else
        VarBaixaAuto = "Nao"
    End If
    
    If VarMostraCusto = 1 Then
        VarMostraCusto = "Sim"
    Else
        VarMostraCusto = "Nao"
    End If
           
Else
    FormConfigura.Show vbModal
End If

VarBloqueado = False

If VarBalanca <> "" Then

    If VarBalanca = 99 Then
        
         AbreAcbr "Bal.Ativar"
'        MSComm1.CommPort = 1
'        MSComm1.Settings = "9600,N,8,1"
'        MSComm1.Handshaking = comNone
'        MSComm1.InputLen = 0
'        MSComm1.SThreshold = 1
'        MSComm1.RThreshold = 1
'        MSComm1.InputMode = comInputModeText
'        MSComm1.EOFEnable = True
'        MSComm1.PortOpen = True
    
    Else
        RetornoBal = AbrePorta(CInt(VarBalanca), 0, 0, 2)
        If RetornoBal = 1 Then
    '      Retorno = MsgBox("Abriu!", vbOKOnly, "P05 - VB", 0, 0)
        Else
          RetornoBal = MsgBox("Erro! Abertura Porta Balança", vbOKOnly, "P05 - VB", 0, 0)
        End If
    End If
End If

Nor = Chr(18)
Cond = Chr(15)
Gde = Chr(14)
Negi = Chr(27) + Chr(69)
NegF = Chr(27) + Chr(70)
SaltoPag = Chr(12)


Select Case Impressora
    Case 1
        DescImp = "MECAF"
    Case 2
        DescImp = "CORISCO"
    Case 3
        DescImp = "GENERAL"
    Case 4
        DescImp = "URANO"
    Case 5
        DescImp = "BEMATECH"
'        If LerINI("Sistema", "Mp4000", App.Path & "\Bemafi32.ini") <> "" Then
'            VarMp4000 = LerINI("Sistema", "Mp4000", App.Path & "\Bemafi32.ini")
'        Else
'            VarMp4000 = False
'        End If

    Case 6
        DescImp = "DARUMA"
    Case 7
        DescImp = "ELGIN FIT"
    Case 8
        DescImp = "ZPN"
    Case 10
        DescImp = "NENHUMA"
End Select
        
LimiteDesc = LerINI("DescontoFrente", "Valor", LocalBanco & "\config.ini")
EditarFrente = LerINI("DescontoFrente", "Editar", LocalBanco & "\config.ini")
Receber = LerINI("DescontoFrente", "Receber", LocalBanco & "\config.ini")
VarPcoPrazo = LerINI("Pedido", "PzoProduto", LocalBanco & "\config.ini")
VarReajConta = LerINI("Pedido", "ReajustaConta", LocalBanco & "\config.ini")
VarCnpj = LerINI("Config", "Cnpj", LocalBanco & "\config.ini")


If Impr40 = "Nao" Then
    VarEmpresa = LerINI("Config", "Empresa", LocalBanco & "\config.ini")
    VarCGC = LerINI("Config", "Cnpj", LocalBanco & "\config.ini")
    VarInscEst = LerINI("Config", "Inscr", LocalBanco & "\config.ini")
    VarEnd = LerINI("Config", "Endereco", LocalBanco & "\config.ini")
    VarCidade = LerINI("Config", "CidadeEmpresa", LocalBanco & "\config.ini")
End If

VarSenhaBco = LerINI("LOCALBCO", "SenhaBco", LocalBanco & "\config.ini")

DtData.Value = Date

'Impressora = 99
RemoveMenus
 
If dir(App.Path & "\FUNDO.jpg") <> "" Then
   Image1.Picture = LoadPicture(App.Path & "\FUNDO.jpg")
   Shape2.Visible = True
End If
 
If VarSenhaBco <> "" Then
    BtPedidos.Visible = False
    BtConta.Visible = False
    LblAberto.Visible = False
    LblSaldo.Visible = False
End If

If VarMostraCusto = "Sim" Then
    BtCusto.Visible = True
End If

Flag = 0
Flag2 = False
P = NameOfPC(Maquina)
Caminho = LocalBanco
Maquina = LimpaCampo(Trim(Maquina))
Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Dados.mdb")

Me.Caption = Trim(LimpaCampo(Maquina)) & " | " & LocalBanco & " | " & PortaNF & " | " & DescImp & " | " & PortaF

DatVend.DatabaseName = Caminho & "\Dados.Mdb"
DatCli.DatabaseName = Caminho & "\Dados.Mdb"
DatLoja.DatabaseName = Caminho & "\Dados.Mdb"
DatVend.RecordSource = "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7 Order by Nome"
DatVend.Refresh

DatCompra.DatabaseName = Caminho & "\Dados.Mdb"
DatCompra.RecordSource = "SELECT Prd.Icm, Prd.CfopPrd, Prd.SitTRib,  Prd.Classe, Itp.VlrVdaDia, Itp.CodItp, Itp.Produto, Prd.CodBar, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem,  Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ITP.Entregue FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = 0 ORDER BY ITP.CodItp"


Set Registros = Banco.OpenRecordset("SELECT PRD.* , BAR.BarBarra, BAR.Produto FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE Ativo = True ORDER BY PRD.CodPrd")
'Set Registros = Banco.OpenRecordset("SELECT PRD.*, GRU.*, PRD.CodPrd FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo ORDER BY PRD.CodPrd")
Set Clientes = Banco.OpenRecordset("SELECT *  FROM Ent Where Ativo = True and  Tipo = 1 or Tipo = 7  or Tipo = 5 or Tipo = 8 ORDER BY Ent.Nome")
Set RecVendas = Banco.OpenRecordset("SELECT * FROM Itp Order By CodItp ")
Set Pedido = Banco.OpenRecordset("SELECT * FROM Ped Order By CodPed")
Set Loja = Banco.OpenRecordset("SELECT * FROM Loj")
Set CLoja = Banco.OpenRecordset("SELECT * FROM CLoj")
Set DatSenha = Banco.OpenRecordset("Select * from Pwd Where Len(mnuutil)=11")
DatSenha.Requery
DatSenha.MoveLast
DatSenha.MoveFirst

CLoja.Requery
If CLoja.RecordCount > 0 Then
    CLoja.MoveLast
End If
Set Indices = Banco.OpenRecordset("SELECT * FROM Ind WHERE NroIndice = 3")

Registros.Requery
'Registros.MoveLast
'If dir(App.Path & "\frente.vbp") = "" Then
'    Pedido.Requery
'    If Pedido.RecordCount > 0 Then
'        Pedido.MoveLast
'        While Pedido!Operacao <> 1 Or Pedido!Operacao <> 2
'            Pedido.MoveLast
'        Wend
'        If Pedido!dataEmiss <> Date Then
'            If Pedido!dataEmiss <= Date - 2 Then
'                MsgBox "O Sistema não pode prosseguir, favor corrigir a data e a hora", vbInformation
'                End
'            End If
'        End If
'        Pedido.MoveFirst
'    End If
'End If
'Set Ofertas = Banco.OpenRecordset("SELECT PRD.Descrição_Produto, PRD.Preco FROM PRD WHERE(((PRD.Promocao)= -1) OR ((PRD.Concorrencia)= -1)) ORDER BY Descrição_Produto ")
Set Empresa = Banco.OpenRecordset("CFG")

'If Dir(App.Path & "\Mesa.jpg") <> "" Then
   Frame5.Picture = LoadPicture(App.Path & "\Mesa.jpg")
'End If

If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then
'    CmdPopular.Visible = True
Else
'    CmdPopular.Visible = False
'    Toolbar1.Buttons.Item(19).Enabled = False
End If

If dir("C:\Comunix\Win_guest\Win_guest.exe") <> "" Then
'    CmdAmplacard.Visible = True
Else
'    CmdAmplacard.Visible = False
End If

If dir("C:\e-pharma\PBMS_PDV.EXE") <> "" Then
'    CmdPopular.Visible = True
Else
'    CmdPopular.Visible = False
End If

If ConsultaCh = "" Then
    Toolbar1.Buttons(10).Enabled = False

End If

If Registros.RecordCount = 1 Then
    Toolbar1.Buttons(16).Enabled = False
End If

Set RecSenha = Banco.OpenRecordset("SELECT * FROM PWD WHERE Nome = 'LIBERA'")

If RecSenha.RecordCount > 0 Then
    VarSenhaLibera = RecSenha!Senha
Else
    Set RecSenha = Banco.OpenRecordset("SELECT * FROM PWD WHERE Nome = 'USUARIO'")
    If RecSenha.RecordCount > 0 Then
        VarSenhaLibera = "USUARIO"
        Set RecSenha = Banco.OpenRecordset("SELECT * FROM PWD")
    Else
        VarSenhaLibera = ""
    End If
End If

If CfgDesc = "Sim" Or CfgDesc = 1 Then
    MebDesconto.Enabled = True
    MebPreco.Enabled = True
Else
    MebDesconto.Enabled = False
    MebPreco.Enabled = False
End If

If VarComplemento = "Sim" Then
    txtcomplemento.Enabled = True
Else
    txtcomplemento.Enabled = False
End If


' Configuração da Impressora Fiscal
TxtDados(2).Text = 2  'FrmPath.TxtSenha.Text

'If TxtDados(2).Text = 1 Then
    LblCupom = "ECF"
    Select Case Impressora
        Case 1
            Impressora = 1
            TxtDados(2).Text = Impressora
        Case 2
'            Impressora = 2
'            Set CTV3 = CreateObject("CTV3.PosDLLCorisco")
'            Status = CTV3.AbreComunicacao(Mid$(PortaF, 4, 1), "19200,n,8,1", "OFF", Mensagem)
        Case 5
            Retorno = Bematech_FI_AbrePortaSerial()
            Mens.Caption = Retorno
        Case 6
            Retorno = Daruma_FI_AbrePortaSerial()
            Mens.Caption = Retorno
        
        Case 7
            NomeImpressora = "ELGIN"
            Retorno = Elgin.Elgin_AbrePortaSerial()
            TrataRetorno2 Retorno
            Retorno = Elgin.Elgin_LinhasEntreCupons(0)
            TrataRetorno2 Retorno
        
    '       ELGIN TERMICA
        Case 8
        
            PortaAcbr = LerINI("Sat", "LocalAcbr", App.Path & "\config.ini")
        
            If dir(PortaAcbr & "\Sai.txt") <> "" Then
                Kill PortaAcbr & "\Sai.txt"
            End If
        
        Case 9
        
            PortaAcbr = LerINI("Sat", "LocalAcbr", App.Path & "\config.ini")
        
            If dir(PortaAcbr & "\Sai.txt") <> "" Then
                Kill PortaAcbr & "\Sai.txt"
            End If
            
    '        PortaAcbr = "\\sidival-pc\ACBrMonitorPLUS"
            
    '        AbreAcbr "SAT.ConsultarSAT"
    '       TrataAcbr
            
            
            AbreAcbr "Sat.Inicializar"
 '           TrataAcbr
             
 '            AbreAcbr "SAT.ConsultarSAT"
 '            TrataAcbr
             
    '        AbreAcbr "SAT.ImprimirExtratoVenda(" & "C:\ACBrMonitorPLUS\Arqs\SAT\Vendas\CFe35150611111111111111591234567890000279333438" & ")"
    '        TrataAcbr
    '        AbreAcbr "SAT.Ativar(" & """" & "11111111111111" & """" & ",35)"
    '        TrataAcbr
    '        AbreAcbr "SAT.CriarEnviarCfe(" & App.Path & "\cfe.ini" & ")"
    '        TrataAcbr
            
    '        AbreAcbr "SAT.ImprimirExtratoVenda(" & "C:\ACBrMonitorPLUS\Arqs\SAT\Vendas\CFe35150511111111111111591234567890000031015331.xml" & ")"
    '        TrataAcbr
     
    
    ' SAT.AssociarAssinatura
                
            
                        
        Case Else
             LblCupom = ""
             BtEcf.Visible = False
             Toolbar1.Buttons.Item(13).Enabled = False
             Impressora = 99
             
'            MsgBox "Impressora fiscal não configurada", vbInformation, App.Title
'            Unload Me
'            Exit Sub
    End Select
'Else


'End If

If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then
    
'    Retorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("5", "", "")
    If dir("C:\FPOPULAR\FP_REQ.txt") <> "" Then
        If MsgBox("Ouve erro o retorno da farmacia popular, pode ser que o comprovante não tenha sido impresso, deseja reimprimir ?", vbYesNo, App.Title) = vbYes Then
            If Impressora = 5 Then
               Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
                Open "C:\FPOPULAR\FP_REQ.txt" For Input As #6
                Do
                   Line Input #6, VarTemp
                       If Left(VarTemp, 2) = 7 Then
            '                Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(psCompString(3, Mid(VarTemp, 3, 46), 48))
                             Retorno = Bematech_FI_UsaRelatorioGerencialMFD(psCompString(3, Mid(VarTemp, 3, 46), 48))
            
                    End If
                Loop Until EOF(6)
                Close #6
                    
            '    Retorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
                 Retorno = Bematech_FI_FechaRelatorioGerencial()
            
            End If
            
            If impresspra = 9 Then
               Iniciar (False)
               Open "C:\FPOPULAR\FP_REQ.txt" For Input As #6
               Do
                  Line Input #6, VarTemp
                      If Left(VarTemp, 2) = 7 Then
                          Print #1, psCompString(3, Mid(VarTemp, 3, 46), 48)
                   End If
               Loop Until EOF(6)
               Close #6
               Print #1, "" & Chr(27) & "i"        '& Chr(27) & "v50"
               Fim (False)
            End If
            
        End If
    End If
Else
    Toolbar1.Buttons.Item(19).Enabled = False
End If
' Configuração do Leitor de Código de Barra

If Leitor <> "TECLADO" Then

    MSComm.CommPort = Mid$(Leitor, 7, 1)
    MSComm.Settings = "9600,N,8,1"
    MSComm.Handshaking = comNone
    MSComm.InputLen = 0
    MSComm.SThreshold = 1
    MSComm.RThreshold = 1
    MSComm.InputMode = comInputModeText
    MSComm.EOFEnable = True
    MSComm.PortOpen = True

End If
    
If dir(App.Path & "\Livre.jpg") <> "" Then
    Timer1.Enabled = True
    Timer1.Interval = 65000
Else
    Timer1.Enabled = False
End If
    

If Pedido.RecordCount <> 0 Then
    Pedido.MoveLast
    Pedido.MovePrevious
    If Not Pedido.BOF Then
        Vendedor = Pedido("Vendedor")
        Criterio = "CodEntidade = " & Vendedor
        DatVend.Recordset.FindFirst Criterio
        DbcVendedor = DatVend.Recordset("Nome")
    End If
Else
    Vendedor = "(Nenhum)"
    Criterio = "Nome = " & """" & Vendedor & """"
    DatVend.Recordset.FindFirst Criterio
    DbcVendedor = DatVend.Recordset("Nome")
End If

FormSenha = False

If VarConveniencia = True Or VarLanchonete = True Then
'    For x = 0 To 203
'        Mesa(x).Visible = False
'    Next
    sql = "UPDATE Conv SET Conv.Cancelado = -1 , Funcionario = 0  WHERE Flag = 0  " 'AND Maquina = " & """" & Trim(LimpaCampo(Left(FrmFrente.Maquina, 15))) & """"
    Banco.Execute sql
    
    AtuMesa
'    Frame5.Visible = True
End If

DatMinutos.DatabaseName = Caminho & "\Minutos.mdb"
DatMinutos.RecordSource = "SELECT Avg(Ncm.aliqNac) AS MediaDealiqNac, Avg(Ncm.aliqEst) AS MediaDealiqEst FROM Ncm"
DatMinutos.Refresh
VarMediaImpostoNac = DatMinutos.Recordset.Fields("MediaDealiqNac")
VarMediaImpostoEst = DatMinutos.Recordset.Fields("MediaDealiqEst")

DatMinutos.RecordSource = "Select codigo, AliqNac, AliqEst From Ncm Order By Codigo"
DatMinutos.Refresh


If VarConveniencia = False Then
    Timer2.Enabled = False
End If

If VarLanchonete = False Then
    Timer2.Enabled = False
'    Toolbar1.Buttons.Item(22).Visible = False
End If


If dir(LocalBanco & "\Vendas.Vbp") = "" Then

    If DatSenha.RecordCount > 0 Then
        FormSenhaInicial.Show vbModal
        Criterio = "CodEntidade = " & DatSenha!MnuCad
        DatVend.Recordset.FindFirst Criterio
        DbcVendedor = DatVend.Recordset("Nome")
        DbcVendedor.Enabled = False
            
        If FormSenha = False Then
            Banco.Close
            Set Banco = Nothing
            Unload Me
        End If
    End If
End If

End Sub

Function LimpaTela()
TxtDados(0) = ""
LblDescricao = "CAIXA LIVRE"
MebUnit = ""
MebTotal = ""
Soma = 0
ContPro = 0
Valor = 0
txtpgto = ""
CodCli = 0
LblCodigo.Caption = ""
LblNome.Caption = ""
LblEnd.Caption = ""
LblCidade.Caption = ""
LblPedido.Caption = ""
LblItens.Caption = ""
DbcCliente.Text = ""
VarPrdDesc = 0
'LblProduto.Caption = ""
'MebQtd = ""
Shape1.FillColor = &HFFFFFF
LblSaldo = ""
LblDias = ""
LblCredito = ""
LblPrazo = ""
LblAberto = ""
CodTotalizador = 0
TxtObs = ""
txtcomplemento = ""
ListVda.ListItems.Clear
TxtEstoque = ""
TxtUltimo = ""
VarSomadeCusto = 0
DbcCliente.Enabled = False
CmdCli.Enabled = False


'MebUnitario = ""
'MebTot = ""
DatCompra.RecordSource = "SELECT Prd.Icm, Prd.CfopPrd, Prd.SitTRib, Classe, VlrVdaDia,  Itp.CodItp, Itp.Produto, Prd.CodBar, ITP.pedido, ITP.Complemento, ITP.Qtd,  Prd.Embalagem, Prd.Tributo, Prd.Aliquota, PRD.Descrição_Produto,  ITP.ValorUnit, ITP.Valor, ITP.CodItp,  ITP.Entregue  FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = 0 ORDER BY ITP.CodItp"
'TxtDados(2).SetFocus
'TxtDados(0).SetFocus
DatCompra.Refresh
'TxtDados(0).SetFocus
End Function

Private Sub SubWizard1_GotFocus()

End Sub

Private Sub OptPreco_Click(Index As Integer)
Dim sql As String

If MsgBox("Deseja mudar o valor dos produtos ?", vbYesNo) = vbYes Then

   If OptPreco(0).Value = True Then
       sql = "UPDATE ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd SET ITP.ValorUnit = [prd].[Venda], ITP.Valor = [prd].[Venda]*[itp].[qtd], ITP.VlrVdaDia = [prd].[Venda]*[itp].[qtd], ITP.VendaDia = [prd].[Venda] WHERE ITP.Pedido = " & LblPedido
       Vista = True
   Else
       sql = "UPDATE ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd SET ITP.ValorUnit = [prd].[pcoprazo], ITP.Valor = [prd].[pcoprazo]*[itp].[qtd], ITP.VlrVdaDia = [prd].[pcoprazo]*[itp].[qtd], ITP.VendaDia = [prd].[pcoprazo] WHERE ITP.Pedido = " & LblPedido
       Vista = False
   End If
   Banco.Execute sql
    
   With FrmFrente
    .ContList = 0
    .DbcCliente.Enabled = True
       .DatCompra.RecordSource = "SELECT Prd.CodBar,  Itp.Entregue, Itp.Produto, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp , Itp.Loja FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & LblPedido & " ORDER BY ITP.CodItp"
       .DatCompra.Refresh
               
       If .DatCompra.Recordset.RecordCount > 0 Then
            
            Criterio = "CodPed = " & LblPedido
            .Pedido.FindFirst Criterio
            If .Pedido.NoMatch Then
                MsgBox "Erro ao localizar pedido ! Sai do programa e edita de novo ", vbInformation, App.Title
                Exit Sub
            End If
            
            
            .MebTotal = 0
            .ListVda.ListItems.Clear
            While Not .DatCompra.Recordset.EOF
                .MebTotal = .MebTotal + (CCur(.DatCompra.Recordset("Qtd")) * CCur(.DatCompra.Recordset("ValorUnit")))
                .MebTotal = Format(.MebTotal, "##,##0.00")
                
                Set .LstPesq = .ListVda.ListItems.Add(, , .DatCompra.Recordset("CodItp"))
                .LstPesq.SubItems(1) = Format(.DatCompra.Recordset.Fields("Qtd"), " ##,##0.000")  'SeCampoNulo(Consulta, "Agencia")
                .LstPesq.SubItems(2) = .DatCompra.Recordset.Fields("Produto")
                .LstPesq.SubItems(3) = .DatCompra.Recordset.Fields("CodBar")
                .LstPesq.SubItems(4) = .DatCompra.Recordset("Descrição_Produto")    'SeCampoNulo(Consulta, "Agencia")
                .LstPesq.SubItems(5) = Format(.DatCompra.Recordset("ValorUnit"), " ##,##0.00")  'SeCampoNulo(Consulta, "Conta")
                .LstPesq.SubItems(6) = Format(Format(CCur(.DatCompra.Recordset("Qtd")), "##,##0.000") * Format(CCur(.DatCompra.Recordset("ValorUnit")), "##,##0.000"), "##,##0.00")  'SeCampoNulo(Consulta, "DataOper")
                .LstPesq.SubItems(7) = .DatCompra.Recordset("CodItp")
                .ContList = ContList + 1
                .DatCompra.Recordset.MoveNext
            Wend
            Soma = .MebTotal
       End If
       
    End With
               
End If

End Sub

Private Sub Timer1_Timer()
If Soma = 0 And Flag = 0 And CodCli = 0 Then
    If Tempo > 5 Then
        FrmLivre.Show vbModal
        Tempo = 0
    Else
        Tempo = Tempo + 1
    End If
End If
End Sub

Private Sub Timer2_Timer()
If VarLanchonete = True And TxtTmp2 = "" Then
    AtuMesa
Else
'    Frame5.Visible = False
End If
End Sub

Private Sub Timer3_Timer()
 BtAtu_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        TxtDados_KeyDown 0, 122, 0
    Case 4
        TxtDados_KeyDown 0, 116, 0
    Case 7
       TxtDados_KeyDown 0, 117, 0
    Case 10
        TxtDados_KeyDown 0, 118, 0
    Case 19
        If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then
            If Soma <> 0 Then
                FormAutoriza.Show vbModal
            Else
                MsgBox "Não tem venda pra autorizar !", vbInformation, App.Title
            End If
        Else
                MsgBox "Não tem autorizador nessa maquina !", vbInformation, App.Title
        End If
    Case 22
'        If Frame5.Visible = False Then
'            Frame5.Visible = True
'            FlagFrame = True
'            Timer2.Enabled = True
'        Else
'            Frame5.Visible = False
'            FlagFrame = False
'            Timer2.Enabled = False
'            TxtDados(0).SetFocus
'        End If
         If MsgBox("Confirma Inclusão do Pacote", vbYesNo) = vbYes Then
              DatCompra.RecordSource = "SELECT Prd.CodBar,  Itp.Entregue, Itp.Produto, ITP.pedido, ITP.Complemento ,ITP.Qtd, PRD.Descrição_Produto, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, ITP.ValorUnit, ITP.Valor, ITP.CodItp , Itp.Loja FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & LblPedido & " ORDER BY ITP.CodItp"

         
              If DatCompra.Recordset.RecordCount > 0 Then
                DatCompra.Refresh
                DatCompra.Recordset.MoveFirst
                While Not DatCompra.Recordset.EOF
                    sql = "INSERT INTO Conv (Cliente, Maquina, Data, Hora, Cupom, Qtd, Produto, Flag,  Valor ) Values ("
                    sql = sql & LblCodigo & " ,"
                    sql = sql & """" & LimpaCampo(Left(DatCompra.Recordset("Complemento"), 25)) & """" & " , #"
                    sql = sql & Format(Date, "mm/dd/yyyy") & "# , "
                    sql = sql & """" & CStr(Time) & """" & ", "
                    sql = sql & LblPedido & " , "
                    sql = sql & DatCompra.Recordset("Qtd") & " , "
                    sql = sql & DatCompra.Recordset("Produto") & " , "
                    sql = sql & -1 & " , "
                    sql = sql & Numero(DatCompra.Recordset("Valor"), 2) & ")"
                    Banco.Execute sql
                    DatCompra.Recordset.MoveNext
                Wend
              Else
                    MsgBox "Não tem pedido lançado !", vbInformation
              End If
              BtTotal_Click

         End If
    '    Sair
End Select
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
    Case "Configura Terminal"
        FormConfigura.Show vbModal
    Case "Data e Hora"
        FormData.Show vbModal
    Case "Lançar Conta"
        FormLancaCta.Show vbModal
    Case "Leitura X"
        FrmFrente.Soma = 1
        If MsgBox("Confirma leitura X ?", vbYesNo) = vbYes Then
            Select Case Impressora
                Case 1
                    Retorno = LEITURAX("0")
                    MensErro (Retorno)
                Case 2
                    Status = CTV3.LEITURAX("0", " ", Mensagem)
                    Mens.Caption = Status
                Case 5
                    Retorno = Bematech_FI_LeituraX()
                    FrmFrente.Mens.Caption = Retorno
                
                Case 6
                    Retorno = Daruma_FI_LeituraX()
                    FrmFrente.Mens.Caption = Retorno
                
                Case 7
                    Retorno = Elgin.Elgin_LeituraX()
                    TrataRetorno2 Retorno
                                        
                Case 8
                    DLLG2_ExecutaComando Handle, "EmiteLeituraX"
                    coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                    circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                    If coderro <> "OK" Then
                        MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                        
                    End If
                                            
                Case 99
                    Entradas = 0
                    Saidas = 0
'                    On Error GoTo ManutençãoErro
                    Set TotVda = FrmFrente.Banco.OpenRecordset("SELECT Sum(TOT.Total) AS SomaDeTotal, TOT.Pgto FROM TOT WHERE TOT.Data = #" & Format(Date, "mm/dd/yyyy") & "# AND TOT.Maquina = " & """" & LimpaCampo(Left(FrmFrente.Maquina, 15)) & """" & " AND TOT.Cancelado = False  GROUP BY TOT.Pgto ORDER BY TOT.Pgto")
                    TotVda.MoveFirst
                    
                    Set TotCxa = FrmFrente.Banco.OpenRecordset("SELECT * FROM CXA WHERE Data = #" & Format(Date, "mm/dd/yyyy") & "# AND Maquina = " & """" & LimpaCampo(Left(FrmFrente.Maquina, 15)) & """" & " AND Tipo = 'M' ORDER BY  D_C")
                    If TotCxa.RecordCount <> 0 Then
                        TotCxa.MoveFirst
                    End If
                    
                    Iniciar (False)
                    If Left(PortaNF, 3) <> "COM" Then

                        Print #1, ""
                        Print #1, Centralizar("LEITURA X", 47)
                        Print #1, "==============================================="
                        Print #1, "Data: " & Date & "   Hora: " & Time

                        Print #1, "-----------------------------------------------"
                        Print #1, Centralizar("VENDAS", 48)
                        Print #1, "-----------------------------------------------"
                        While Not TotVda.EOF
                            Select Case TotVda("Pgto")
                                Case 1
                                    Print #1, "Dinheiro ....................(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                Case 2
                                    Print #1, "Cheque ......................(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                Case 3
                                    Print #1, "Cheque Pré ..................(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                Case 4
                                    Print #1, "Cartão ......................(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                Case 5
                                    Print #1, "Ticket ......................(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                Case 6
                                    Print #1, "Venda a Prazo ...............(+) "; Alinhar(str(TotVda("SomaDeTotal")), 12, 2)
                                    VdaPrazo = VdaPrazo + TotVda("SomaDeTotal")
                            End Select
                            If TotVda("Pgto") < 6 Then
                                Entradas = Entradas + TotVda("SomaDeTotal")
                            End If
                            TotVda.MoveNext
                        Wend
                        Print #1, "-----------------------------------------------"
                        Print #1, Centralizar("MOVIMENTACAO CAIXA", 47)
                        Print #1, "-----------------------------------------------"
                        Print #1, ""
                        
                        If TotCxa.RecordCount <> 0 Then
                            While Not TotCxa.EOF
                                If TotCxa("D_C") = "C" Then
                                    Print #1, Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))); "(+)"; Alinhar(str(TotCxa("valor")), 12, 2)
                                    Entradas = Entradas + TotCxa("Valor")
                                ElseIf TotCxa("D_C") = "D" Then
                                    Print #1, Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))); "(-)"; Alinhar(str(TotCxa("valor")), 12, 2)
                                    Saidas = Saidas + TotCxa("Valor")
                                End If
                                TotCxa.MoveNext
                            Wend
                        End If
                        Print #1, "-----------------------------------------------"
                        Print #1, Centralizar("TOTALIZACAO", 47)
                        Print #1, "-----------------------------------------------"
                        Print #1, ""
                        Print #1, "Total do Caixa...............(=) "; Alinhar(CCur(Entradas) + CCur(VdaPrazo), 12, 2)
                        Print #1, "Saidas do Caixa..............(-) "; Alinhar(CCur(Saidas), 12, 2)
                        Print #1, "Vendas a Prazo...............(-) "; Alinhar(CCur(VdaPrazo), 12, 2)
                        Print #1, "Total Liquido................(=) "; Alinhar(CCur(Entradas) - CCur(Saidas), 12, 2)
                        Print #1, "" & Nor
                        Print #1, ""
                    
                    Else
                    
                        Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX(Centralizar("LEITURA X", 47) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("===============================================" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("Data: " & Date & "   Hora: " & Time & Chr(10), 2, 0, 0, 0, 0)
                        
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX(Centralizar("VENDAS", 48) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        
                        While Not TotVda.EOF
                            Select Case TotVda("Pgto")
                                Case 1
                                    Comando = FormataTX("Dinheiro ....................(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                Case 2
                                    Comando = FormataTX("Cheque ......................(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                Case 3
                                    Comando = FormataTX("Cheque Pré ..................(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                Case 4
                                    Comando = FormataTX("Cartão ......................(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                Case 5
                                    Comando = FormataTX("Ticket ......................(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                Case 6
                                    Comando = FormataTX("Venda a Prazo ...............(+) " & Alinhar(str(TotVda("SomaDeTotal")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    VdaPrazo = VdaPrazo + TotVda("SomaDeTotal")
                            End Select
                            If TotVda("Pgto") < 6 Then
                                Entradas = Entradas + TotVda("SomaDeTotal")
                            End If
                            TotVda.MoveNext
                        Wend
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX(Centralizar("MOVIMENTACAO CAIXA", 47) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                       
                        If TotCxa.RecordCount <> 0 Then
                            While Not TotCxa.EOF
                                If TotCxa("D_C") = "C" Then
                                    Comando = FormataTX(Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))) & "(+)" & Alinhar(str(TotCxa("valor")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Entradas = Entradas + TotCxa("Valor")
                                ElseIf TotCxa("D_C") = "D" Then
                                    Comando = FormataTX(Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))) & "(-)" & Alinhar(str(TotCxa("valor")), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Saidas = Saidas + TotCxa("Valor")
                                End If
                                TotCxa.MoveNext
                            Wend
                        End If
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX(Centralizar("TOTALIZACAO", 47) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                        
                        Comando = FormataTX("Total do Caixa...............(=) " & Alinhar(CCur(Entradas) + CCur(VdaPrazo), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("Saidas do Caixa..............(-) " & Alinhar(CCur(Saidas), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("Vendas a Prazo...............(-) " & Alinhar(CCur(VdaPrazo), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("Total Liquido................(=) " & Alinhar(CCur(Entradas) - CCur(Saidas), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                        
                        Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                        Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                                                
                    End If
                    Fim (False)
            End Select
        End If
        FrmFrente.Soma = 0
    Case "Redução Z"
            FrmFrente.Soma = 1
            MsgBox "A T E N Ç Ã O . . .  Está operação fecha o caixa por hoje !", vbExclamation, App.Title
            If MsgBox("Confirma redução Z ?", vbYesNo + vbDefaultButton2) = vbYes Then
                Select Case Impressora
                    Case 1
                        Retorno = REDUCAOZ("0")
                        MensErro (Retorno)
                    Case 2
                        Status = CTV3.FecharEcf("0", Chr(13), Mensagem)
                        Mens.Caption = Status
                    Case 5
                        Retorno = Bematech_FI_ReducaoZ("", "")
                        FrmFrente.Mens.Caption = Retorno
                    Case 6
                        Retorno = Daruma_FI_ReducaoZ("", "")
                        FrmFrente.Mens.Caption = Retorno
                        
                    Case 7
                        Retorno = Elgin.Elgin_ReducaoZ("", "")
                        TrataRetorno2 Retorno
                        
                    Case 8
                        
                        DLLG2_LimpaParams Handle
                        DLLG2_DefineTimeout Handle, 30
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                                                                        
                        DLLG2_LimpaParams Handle
                        DLLG2_ExecutaComando Handle, "EmiteReducaoZ"
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                        
                        DLLG2_LimpaParams Handle
                        DLLG2_DefineTimeout Handle, 10
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                        
                        
                End Select
            End If
            FrmFrente.Soma = 0
    
    Case "Abertura de Gaveta"
            Select Case Impressora
                Case 1
                    Retorno = AcionarGaveta() ' "0", "32", "38")
                    MensErro (Retorno)
                Case 2
                    Status = CTV3.ABRIRGAVETA(Mensagem)
                    Mens.Caption = Status
                Case 99
                    If Left(PortaNF, 3) <> "COM" Then
                        If Impr40 = "Sim" Then
                            Iniciar (False)
                                Print #1, "" & Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
                            Fim (False)
                        End If
                    End If
            End Select
    Case "Movimentação do Caixa"
        FormCaixa.Show vbModal
    Case "Fechamento do Caixa"
        FormFchCxanovo.Show
    Case "Consulta Cheque"
'        FrmPedidos.Show
    Case "e-Pharma"
        If dir("C:\e-pharma\PBMS_PDV.EXE") <> "" Then
'            Formepharma.Show

                Dim VarPrd As Recordset
                Dim sql As String
                Dim VarLinha As String
                
                If MsgBox("Confirma exportação de produtos ? ", vbYesNo, App.Title) = vbYes Then
                
                        sql = "SELECT Prd.CodPrd, Prd.Descrição_Produto, Prd.Complemento, Prd.Venda, BAR.BarBarra FROM Prd INNER JOIN BAR ON Prd.CodPrd = BAR.Produto ORDER BY BAR.BarBarra"
                        Set VarPrd = FrmFrente.Banco.OpenRecordset(sql)
                        VarPrd.Requery
                        If VarPrd.RecordCount > 0 Then
                           VarPrd.MoveFirst
                           Open "c:\e-pharma\Tabela.txt" For Output As #1
                           While Not VarPrd.EOF
                               
                               If Len(Trim(VarPrd!BarBarra)) >= 8 And Len(Trim(VarPrd!BarBarra)) <> "" Then
                                    VarLinha = psCompString(3, Trim(VarPrd!BarBarra), 13)
                                    VarLinha = VarLinha + psCompString(3, Left(VarPrd!Descrição_Produto, 40), 40)
                                    VarLinha = VarLinha + psCompString(1, LimpaNro(VarPrd!Venda), 7)
                                    VarLinha = VarLinha + psCompString(1, LimpaNro(VarPrd!Venda), 7)
                                    VarLinha = VarLinha + psCompString(1, 0, 8)
                                    Print #1, VarLinha & Chr(13)
                                    VarLinha = ""
                               End If
                               VarPrd.MoveNext
                           Wend
                           MsgBox "Arquivo exportado com sucesso!", vbInformation, App.Title
                           Close #1
                      End If
                End If
        
        Else
            MsgBox "e-Pharma não instalado !", vbInformation
        End If
    Case "VidaLink"
'        FormVidaLink.Show
    Case "Cancela Popular"
        FrmCancelaPopular.Show vbModal
    Case "SevenPdv"
'        FormSevenPdv.Show vbModal
    Case "&Mesa"
        If FlagFrame = False Then
            Frame5.Visible = True
            FlagFrame = True
        Else
            Frame5.Visible = False
            FlagFrame = False
        End If

    Case "Movimento Mesa"
        FormContMesa.Show vbModal
    Case "Imprimir Ficha"
        FormFicha.Show vbModal
    Case "Inicializa Senha"
        EscreveINI "Config", "Senha", 1, App.Path & "\config.ini"
 End Select
 
End Sub
 

Private Sub TxtComplemento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtDados(2) <> "" Then
     LblDescricao.Caption = LblDescricao.Caption & " " & txtcomplemento
    If CfgDesc = "Sim" Or CfgDesc = 1 Then
         MebDesconto.SetFocus
'         Venda
    Else
        MebPreco = Round(CCur(Valor), 2)
        Venda
    End If
End If

End Sub

Private Sub TxtDados_Change(Index As Integer)
If Len(TxtDados(Index)) >= 13 And VarEnter = True Then
    If Mid$(TxtDados(Index), 14, 1) <> Chr(13) Then
              SendKeys ("{Enter}")
    End If
End If

End Sub

Private Sub TxtDados_GotFocus(Index As Integer)
Select Case Index
    Case 0
        If CodCli = 0 Then
             Label6.Caption = "Codigo Cliente ou F2"
             If VarRapido = True Then
                TxtDados(0).Text = 1
                TxtDados_KeyPress 0, 13
             End If
        Else
             Label6.Caption = "EAN 13 ou Qtd"
        End If
    Case 2
        Label6.Caption = "EAN 13/CodPrd ou F2 "
        If Registros.RecordCount = 1 Then
'             If KeyAscii = 13 And TxtDados(2) = "" Then
'                  TxtDados(2).SetFocus
                  TxtDados(2).Text = 1
'                  SendKeys ("{Enter}")
 '             End If
        End If
End Select

End Sub

Private Sub TxtDados_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Status As String
Dim sql As String
Dim Criterio As String
Dim VarNroPed As String
'Dim Criterio As String
Dim VarDup As Recordset
Dim ImprAnt As Byte
Dim Cancela As String
Dim ValorBal As String

Select Case Index
    Case 0
        If CodCli = 0 Then
            If KeyCode = 113 Then
                FrmPesqCli.Show vbModal
                TxtDados(0).SetFocus
                TxtDados_KeyPress 0, 13
                TxtDados(0).SetFocus
            End If
'            If KeyCode = 9 Or KeyCode = 13 Then
'                If TxtDados(0).Text = "" Then
'                    FrmPesqCli.Show vbModal
'                    TxtDados(0).SetFocus
'                    TxtDados_KeyPress 0, 13
'                End If
'            End If
            
        End If
     Case 2
        If KeyCode = 113 Then
            FrmPesquisa.Show vbModal
            TxtDados(2).SetFocus
            TxtDados_KeyPress 2, 13
        End If
        
        If KeyCode = 9 Or KeyCode = 13 Then
            If TxtDados(2).Text = "" Then
                FrmPesquisa.Show vbModal
                TxtDados(2).SetFocus
                TxtDados_KeyPress 2, 13
            End If
        
        End If
        
End Select

If KeyCode = 119 Then  ' And Impressora <> 99 Then
    
'        TxtDados(2).Text = "AGUARDE..."
        Frame.Visible = True
        DoEvents
        If Impressora > 19 Then
                If MebTotal = "" Or MebTotal = "0" Then
                    If MsgBox("Confirma o cancelamento do ultimo cupom ?", vbYesNo, App.Title) = vbYes Then
                        Select Case Impressora
                            Case 1
                                Retorno = CANCELACUPOMFISCAL()
                                MensErro (Retorno)
                            Case 2
                                Status = CTV3.ChkCupomAberto(Mensagem)
                                If Status = "00" Then
                                    Status = CTV3.CancelarCupom("0", Mensagem)
                                Else
                                    Status = CTV3.CancelarCupom("1", Mensagem)
                                End If
                                Mens.Caption = Status
                             Case 5
                                Retorno = Bematech_FI_CancelaCupom()
                                Mens.Caption = Retorno
                             Case 6
                                Retorno = Daruma_FI_CancelaCupom()
                                Mens.Caption = Retorno
                             
                             Case 7
                                Retorno = Elgin.Elgin_CancelaCupom()
                                TrataRetorno2 Retorno
                                                                
                             Case 8
'                                Cupom = Space(6)
                                
                                DLLG2_LimpaParams Handle
                                DLLG2_AdicionaParam Handle, "NomeInteiro", "CCF", 7
                                DLLG2_ExecutaComando Handle, "LeInteiro"
                                Cupom = LimpaNumero(DLLG2_ObtemRetornos(Handle, "valorinteiro", 0))
                                coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                                circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                                If coderro <> "OK" Then
                                    MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                                End If
                                     
                                DLLG2_LimpaParams Handle
                                DLLG2_ExecutaComando Handle, "CancelaCupom"
                                coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                                circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                                If coderro <> "OK" Then
                                    MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                                End If
                                
                        End Select
                    End If
            End If
        End If
        If CodTotalizador <> 0 And DatSenha.RecordCount <= 0 Then
                If Impressora >= 10 Then
                    Cancela = MsgBox("Confirma cancelamento ?", vbYesNo)
                ElseIf MebTotal = "0" Or MebTotal = "" Then
                    Cancela = vbYes
                Else
                    Cancela = MsgBox("Confirma cancelamento ?", vbYesNo)
                End If

                If Cancela = vbYes Then
                   If dir("C:\FPOPULAR\Autorizador_Farmacia_Popular.EXE") <> "" Then
                    
    '                    sql = "DELETE POPULAR.*, PEDIDO FROM POPULAR WHERE PEDIDO  = " & CodTotalizador
    '                    Banco.Execute sql
                    End If
                    
 '                   sql = "DELETE ITP.*, ITP.PEDIDO FROM ITP WHERE ITP.PEDIDO  = " & CodTotalizador
 '                   Banco.Execute sql
                    
                    FrmTotal.VoltaEstoque LblPedido
                    
                    If LblPedido <> "" Then
                        If LblPedido > 0 Then
                            sql = "DELETE Dup.*, DUP.PEDIDO FROM DUP WHERE DUP.PEDIDO  = " & LblPedido
                            Banco.Execute sql
                         End If
                     End If
                     
                    sql = "UPDATE PED SET PED.Operacao = 4, PED.DtSaida = #" & Format(Date, "mm/dd/yyyy") & "#, PED.NroPedido = '999999', PED.PrevEntrega = 'CANCELADO' WHERE PED.CodPed = " & CodTotalizador
                     
'                    sql = "DELETE PED.*, PED.CodPed FROM PED WHERE PED.CodPed  = " & CodTotalizador
                    Banco.Execute sql
                     
                            
                    Open Caminho & "\Delete.txt" For Append As #1
                    Print #1, "FRENTE; Pedido; " & CodTotalizador & "; " & Date & "; " & Time & "; " & LblMaquina; "; R$ " & Format(MebTotal, "#,##0.00;- #,##0.00")
                    Close #1
            
                    TxtDados(2).Text = ""
                                        
                    LimpaTela

                    If VarConveniencia = True Or VarLanchonete = True And CodCard <> 0 Then
                        TxtTmp2 = ""
                        If MsgBox("Deseja cancelar movimento mesa ? ", vbYesNo, App.Title) = vbYes Then
                            sql = "UPDATE Conv SET Cancelado = -1, Totaliza = -1, Funcionario = 0    WHERE Cliente = " & CodCard
                            Banco.Execute sql
                            If TxtTmp3 <> "" Then
                                Mesa(VarContMesa).Caption = ""
                                Mesa(VarContMesa).BackColor = &HFFFFFF
                            End If
                            
                            AtuMesa
                         Else
                            sql = "UPDATE Conv SET  Funcionario = 0    WHERE Cliente = " & CodCard
                            Banco.Execute sql
                            AtuMesa
                            
                         End If
                    End If

                End If
        End If
        Frame.Visible = False
        TxtDados(0).SetFocus
End If


If KeyCode = 118 Then
Senha = False

    If Soma = 0 Then
        
        If VarSenhaBco <> "" Then
            VarTelaSenha = VarSenhaBco
            FormSenha2.Show vbModal
            If Senha = True Then
                FormConsultaPed.Show vbModal
            End If
         Else
            FormConsultaPed.Show vbModal
         
         End If
    Else
        MsgBox "Venda não finalizada", vbInformation, App.Title
    End If

'        MsgBox "Clique duas vezes em cima do item que deseja exluir !", vbInformation, App.Title
End If


If KeyCode = 117 Then
        FrmPesquisa.Show vbModal
End If

If KeyCode = 120 Then
    If Soma <> 0 Then
        ImprAnt = Impressora
        If Shift = 0 Then
            Impressora = 99
        End If
        FrmTotal.LblCupom = NroCupom
        FrmTotal.LblMaquina = Maquina
        FrmTotal.LblTot = CodTotalizador
        FrmTotal.Show vbModal
        Impressora = ImprAnt
    Else
        MsgBox "Cupom Zerado", vbInformation, App.Title
    End If
End If

If KeyCode = 114 Then
'    FrmLanca.TxtCartao = Trim(Mesa(Index).Caption)
'    SendKeys ("{ENTER}")
   If (VarConveniencia Or VarLanchonete) Then
        If Soma = 0 Then
            FrmLanca.Show
        Else
            MsgBox "Cupom nao totalizado", vbInformation, App.Title
        End If
    Else
        FrmCancelaItem.Show
    End If
End If


If KeyCode = 115 And VarBalanca <> "" Then

    If VarBalanca = 99 Then
'        Dim Teclado As String * 5
        
'        Teclado = MSComm1.Input
'        MsgBox Teclado
'        SendKeys Teclado
            If dir(PortaAcbr & "\Sai.txt") <> "" Then
                Kill PortaAcbr & "\Sai.txt"
            End If
    
    
             AbreAcbr "Bal.LePeso"
             ValorBal = Mid$(TrataAcbr, 4, 8)
             If IsNumeric(ValorBal) Then

                 TxtDados(0).Text = ValorBal
    '         MsgBox TxtDados(0), vbInformation
                SendKeys ("{Enter}")
             End If
'       TxtDados(0).Text = 0.32
       If CodCli = 0 Then
            Barra = TxtDados(0).Text
            TxtDados(0) = 1
        Else
            Barra = ""
        End If
         
    Else

        RetornoBal = PegaPeso(0, PesoBal, "C:\")
    '    MsgBox RetornoBal & " " & PesoBal
        If RetornoBal = 1 Then
          ValorBal = PesoBal
          If IsNumeric(ValorBal) Then
            TxtDados(0) = CDbl(Format(ValorBal, "##,000")) / 1000
            SendKeys ("{Enter}")
          End If
        Else
        '      lbPeso.Caption = "Erro"
        End If
    End If
End If
    
If KeyCode = 121 Then
    BtEcf_Click
End If


If KeyCode = 122 Then
    If Soma = 0 Then
        Banco.Close
        Set Banco = Nothing
        If VarBalanca <> "" Then
'            RetornoBal = FechaPortaBal()
        End If
        Unload Me
       
        End
    Else
        MsgBox "Venda não finalizada", vbInformation, App.Title
    End If

End If

If KeyCode = 116 Then
    If Soma = 0 Then
        FormBaixarDupli.Show vbModal
    Else
        MsgBox "Venda não finalizada", vbInformation, App.Title
    End If
End If


End Sub

Private Sub MSComm1_OnComm()
'Dim Teclado As String

'Teclado = MSComm1.Input

'SendKeys Teclado

End Sub


Private Sub TxtDados_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Criterio As String
Dim sql As String
Dim Status As String
Dim buf_ret As String * 100
Dim Conta As Recordset
Dim Controle As Boolean

Controle = True
Dim Cartao  As String
'Barra = ""
Cartao = ""
VarGelada = ""

'On Error GoTo erro

Select Case Index
       
    Case 0
        If CodCli <> 0 Then
            TxtEstoque = ""
            
            If Registros.RecordCount = 1 Then
                If KeyAscii = 13 And TxtDados(0) = "" Then
                              
 '                 TxtDados(0).SetFocus
                  TxtDados(0).Text = 1
'                  SendKeys ("{Enter}")
                                    
 '                 TxtDados(2).SetFocus
 '                 TxtDados(2).Text = 1
 '                 SendKeys ("{Enter}")
                End If
                  
            End If
            If KeyAscii = 13 And TxtDados(0) <> "" Then
            
                If Left(TxtDados(0), 1) = "*" Then
                    LerSeven
                End If
                    
                If VarConveniencia = True Or VarLanchonete = True Then
               
                    If Left(TxtDados(0), 1) = "/" Then
                        LeCartao TxtDados(0)
'                             Vista = True
                        Exit Sub
                    End If
                End If
                If TxtDados(0).Text = "9999" Then
                     TxtDados(0).Text = ""
                     FormListaVda.Show vbModal
'                         Totaliza
                End If
                    
                If Format(Left(TxtDados(0).Text, 1), ">") = "G" Or Format(Left(TxtDados(0).Text, 1), ">") = "/" Then
                    VarGelada = Format(Left(TxtDados(0).Text, 1), ">")
                    TxtDados(0).Text = Mid(TxtDados(0), 2, Len(TxtDados(0)) - 1)
                End If
                           
                If Format(Left(TxtDados(0).Text, 1), ">") = "B" Then
                    VarGelada = Format(Left(TxtDados(0).Text, 1), ">")
                    TxtDados(0).Text = Mid(TxtDados(0), 2, Len(TxtDados(0)) - 1)
                End If
                                                      
                If Len(TxtDados(0)) >= 7 Then
                    If IsNumeric(TxtDados(0)) Then
                        EfetuaVenda (TxtDados(0))
                    Else
                        TxtDados(0) = ""
                        TxtDados(0).SetFocus
                    End If
                Else
                    If IsNumeric(TxtDados(0)) Then
'                        If Len(TxtDados(0)) > 4 Then
'                            If MsgBox("Quantidade muito grande... Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
'                                TxtDados(2).SetFocus
'                            Else
'                                TxtDados(0) = ""
'                                TxtDados(0).SetFocus
'                            End If
'                        Else
                            TxtDados(2).SetFocus
'                        End If
                    Else
                        TxtDados(0) = ""
                        TxtDados(0).SetFocus
                    End If
                End If
            End If
        Else
            
            If Registros.RecordCount = 1 Then
                If KeyAscii = 13 And TxtDados(0) = "" Then
'                    TxtDados(0).SetFocus
                    TxtDados(0).Text = 1
'                    SendKeys ("{Enter}")
                    
               End If
                  
            End If
            
            If KeyAscii = 13 And TxtDados(0) <> "" Then
                    ContList = 0
                    CodCard = 0
                                       
                    If VarConveniencia = True Or VarLanchonete = True Then

                            If Left(TxtDados(0), 1) = "/" Then
                                Cartao = TxtDados(0)
                                If CliCartao <> "" Then
                                    TxtDados(0).Text = CliCartao
                                Else
                                    TxtDados(0).Text = 1
                                End If
                                CliCartao = ""
                            Else
                                Cartao = ""
                            End If
                            Vista = True
                    End If
                    If Not IsNumeric(TxtDados(0)) Then
                        TxtDados(0).Text = ""
                        Exit Sub
                    End If
                    
                    If DbcVendedor = "" Then
                        MsgBox "Vendedor não identificado !", vbExclamation, App.Title
                        TxtDados(0).SetFocus
                        Exit Sub
                    End If
                    LblDescricao.Caption = ""
                    MebTotal = ""
                    If Len(TxtDados(0)) > 6 Then
                        Barra = TxtDados(0).Text
                        TxtDados(0).Text = 1
                    Else
'                        Barra = ""
                    End If
                    VarBloqueado = False
                    VarOrcaPrazo = False
                    If IsNumeric(TxtDados(0)) Then
                        Criterio = "CodEntidade = " & Int(TxtDados(0))
                        Clientes.FindFirst Criterio
                        If Clientes.NoMatch Then
                            Beep
                            TxtDados(0).Text = ""
                            TxtDados(0).SetFocus
                            SendKeys "{BACKSPACE}"
                            MsgBox "Cliente não cadastrado !", vbInformation, App.Title
                        Else
'                            If (Clientes!Tipo <> 1 Or Clientes!Tipo <> 7) Then
'                            End If

                            If Not IsNull(Clientes!DtNasc) Then
                                If Month(Clientes!DtNasc) = Month(Date) And VarColombo = True Then
                                    VarAniversariante = True
                                    MsgBox ("FELIZ ANIVERSARIO !!! VC TERÁ UM DESCONTO DE 10% "), vbInformation
                                Else
                                    VarAniversariante = False
                                End If
                            Else
                                VarAniversariante = False
                            End If

                            If TxtDados(0) = "" Then
                                CodCli = 0
                            Else
                                CodCli = TxtDados(0)
                                Vista = True
                            End If
                            TxtDados(2) = "Aguarde . . ."
                            If Clientes!Status = 4 Then
                                If VarBloqueio = False Then
                            
                                     MsgBox "Cliente Bloqueado !  Você só pode vender a vista " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
                                    
                                        FormSenha = True
                                        If VarSenhaLibera <> "" Then
                                            VarTelaSenha = VarSenhaLibera
                                            FormSenha2.Show vbModal
                                        Else
                                            FormSenha = False
                                        End If
                                        If FormSenha = False Then
                                             VarBloqueado = True
    '                                        TxtDados(0).Text = ""
    '                                        TxtDados(2) = ""
    '                                        TxtDados(0).SetFocus
    '                                        CodCli = 0
    '                                        Exit Sub
                                        End If
                                 Else
                                    MsgBox "Cliente Bloqueado ! " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
                                    TxtDados(0).Text = ""
                                    TxtDados(2) = ""
                                    TxtDados(0).SetFocus
                                    CodCli = 0
                                    Exit Sub
                                End If
                            End If
                            If CodCli <> 1 And Clientes.Fields("Prazo") <> "" Then
'                                    Set Conta = Banco.OpenRecordset("SELECT DUP.*, DUP.Cond, DUP.DtVenc FROM DUP WHERE  DUP.Cond=False AND DUP.DtVenc <= #" & Format(Date - Clientes.Fields("Prazo"), "mm/dd/yyyy") & "# and DUP.Entidade = " & CodCli)
                                    Set Conta = Banco.OpenRecordset("SELECT DUP.* FROM DUP WHERE  DUP.Cond=False AND DUP.Entidade = " & CodCli & " Order By Dup.DtVenc ")
                                    Conta.Requery
                                    LblCredito = Format(Clientes!Credito, "#,##0.00;- #,##0.00")
                                    LblPrazo = Clientes!Prazo
                                    LblDias = 0
                                    LblSaldo = Format(0, "#,##0.00;- #,##0.00")
                                    LblAberto = 0
                                    Shape1.FillColor = &HFF00&
                                    If Conta.RecordCount <> 0 Then
                                        Dim SomaCta As Currency
                                        SomaCta = 0
                                        Conta.MoveFirst
                                        If Date > Conta!DtVenc Then
                                            LblDias = Date - Conta!DtVenc
                                            If IsNumeric(Clientes!Prazo) Then
                                                If Conta("DtVenc") < (Date - (Clientes!Prazo * 2)) Then
                                                    Shape1.FillColor = &H80000008
                                                ElseIf Conta("DtVenc") < (Date - Clientes!Prazo) Then
                                                        Shape1.FillColor = &HFF&
                                                    Else
                                                        Shape1.FillColor = &HFF00&
                                                End If
                                            End If
                                        End If
                                        While Not Conta.EOF
                                            SomaCta = SomaCta + CCur(Conta!Valor)
                                            Conta.MoveNext
                                        Wend
                                    Else
                                        Shape1.FillColor = &HFF00&
                                    End If
                                    LblSaldo = Format(Clientes!Credito - SomaCta, "#,##0.00;- #,##0.00")
                                    LblAberto = Format(SomaCta, "#,##0.00;- #,##0.00")
                                    If Shape1.FillColor <> -2147483640 Then
                                        If SomaCta > (Clientes!Credito * 2) Then
                                              Shape1.FillColor = &H80000008
                                        ElseIf SomaCta > Clientes!Credito Then
                                               Shape1.FillColor = &HFF&
                                            Else
                                                If Shape1.FillColor = &HFF00& Then
                                                    Shape1.FillColor = &HFF00&
                                                End If
                                        End If
                                    End If
                                    Vista = False
                                    VarDtVenc = CDate(Date + LblPrazo)
                            Else
                                    Vista = True
                            End If
                            If Clientes!Status = 5 Then
                                If Shape1.FillColor = 255 Or Shape1.FillColor = -2147483640 Then
                                    If VarBloqueio = False Then
                                        MsgBox "Cliente Bloqueado !  Você só pode vender a vista " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
                                        
                                        FormSenha = True
                                        If VarSenhaLibera <> "" Then
                                            VarTelaSenha = VarSenhaLibera
                                            FormSenha2.Show vbModal
                                        Else
                                            FormSenha = False
                                        End If
                                        If FormSenha = False Then
                                             VarBloqueado = True
    '                                        TxtDados(0).Text = ""
    '                                        TxtDados(2) = ""
    '                                        TxtDados(0).SetFocus
    '                                        CodCli = 0
    '                                        Exit Sub
                                        End If
                                     Else
                                        MsgBox "Cliente Bloqueado ! " + Chr(13) + Chr(13) + Clientes!Obs, vbExclamation, App.Title
                                        TxtDados(0).Text = ""
                                        TxtDados(2) = ""
                                        TxtDados(0).SetFocus
                                        CodCli = 0
                                        Exit Sub
                                    End If
                                        
                                End If
                            End If
                            VarDiaBom = 0

                            If Not IsNull(Clientes.Fields("Prazo")) Then
                                If IsNumeric(Clientes.Fields("Prazo")) Then
                                    Dias = Clientes.Fields("Prazo")
                                    If Dias = "0" Then
                                        Dias = 30
                                    End If
                                End If
                            End If
                            
                            If Not IsNull(Clientes.Fields("DiaBomPgto")) Then
                                If IsNumeric(Clientes.Fields("DiaBomPgto")) Then
                                    VarDiaBom = Clientes.Fields("DiaBomPgto")
                                End If
                            End If
                            LblCodigo.Caption = Clientes.Fields("CodEntidade")
                            Criterio = "CodEntidade = " & LblCodigo
                            DatCli.Recordset.FindFirst Criterio
                            DbcCliente.Text = DatCli.Recordset("Nome")
'                            LblNome.Caption = Clientes.Fields("Nome")
                            If IsNull(Clientes.Fields("Endereco")) Then
                                LblEnd.Caption = ""
                            Else
                                LblEnd.Caption = Trim(Clientes.Fields("Endereco")) & ", " & Clientes.Fields("Nro")
                            End If
                            If IsNull(Clientes.Fields("Cidade")) Then
                                LblCidade.Caption = ""
                            Else
                                LblCidade.Caption = Clientes.Fields("Cidade")
                            End If
'                            If CodCli <> 1 Then
                                If Not IsNull(Clientes!PesAut) Then
                                    FormVista.TxtPesAut = Clientes!PesAut
                                End If
                                If Not IsNull(Clientes!Obs) Then
                                    FormVista.TxtObs = Clientes!Obs
                                End If
                                If Dias = "" Then
                                    Dias = 30
                                End If
                                If CodCli <> 1 Then
                                    If VarDiaBom <> 0 Then
                                        If VarDiaBom <= 28 Then
                                            If Month(Date) + 1 > 12 Then
                                                VarDtVenc = Format(VarDiaBom & "/01/" & Year(Date) + 1, "dd/mm/yyyy")
                                            Else
                                                VarDtVenc = Format(VarDiaBom & "/" & Month(Date) + 1 & "/" & Year(Date), "dd/mm/yyyy")
                                            End If
                                        Else
                                            VarDtVenc = Format(Date + Dias, "dd/mm/yyyy")
                                        End If
                                    Else
                                        VarDtVenc = Format(Date + Dias, "dd/mm/yyyy")
                                    End If
                                Else
                                    VarDtVenc = Date
                                End If
                                If Barra = "" And Cartao = "" Then
                                    If VarMercado = True Then
'                                        If MsgBox("Identificar cliente para nota fiscal paulista ?", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                                        
                                         Label6.Caption = "EAN 13 ou Qtd"

'                                        End If
                                    Else
                                        FormVista.Show vbModal
                                    End If
                                End If
'                            End If
                            If LblPedido <> "" Then
                                If Vista = True Then
                                    OptPreco(0).Value = True
                                Else
                                    OptPreco(1).Value = True
                                End If
                            End If
                            NroCupom = 0
                            
                            Pedido.AddNew
                            Pedido("Entidade") = CodCli
                            Pedido("Operacao") = 4
                            Pedido("Vendedor") = DbcVendedor.BoundText
                            Pedido("DataEmiss") = Date
                            Pedido("Hora") = Time
                            Pedido("DtSaida") = Date
                            Pedido("NroPedido") = ""
                            Pedido("Contato") = LimpaCampo(Trim(Maquina))
'                            Pedido("NroPedido") = txtPedidoRef.Text
'                           Pedido("Desconto") = MebDados(3).Text
                            Pedido.Update
                            Pedido.Requery
                            Pedido.MoveLast
                            Controle = True
                            While Controle And Not Pedido.BOF
                                If Pedido.Fields("Vendedor") = DbcVendedor.BoundText Then
                                    CodTotalizador = Pedido.Fields("CodPed")
                                    Controle = False
                                End If
                                Pedido.MovePrevious
                            Wend
                            Set VarLista1 = Banco.OpenRecordset("SELECT HISTEST.Nome, HISTEST.CodPrd, HISTEST.Valor, HISTEST.Qtd, HISTEST.Ordem FROM HISTEST WHERE HISTEST.Nome = " & """" & TxtDados(0) & """" & " ORDER BY HISTEST.CodPrd")
                            VarLista1.Requery
                            

'                            DatCompra.RecordSource = "SELECT  Itp.Produto, ITP.pedido, ITP.Complemento, ITP.Qtd, Prd.Embalagem, Prd.Tributo, Prd.Aliquota, PRD.Descrição_Produto, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ITP.Entregue FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto WHERE Itp.Pedido = " & CodTotalizador & " ORDER BY ITP.CodItp"
                                                        
                            LblPedido.Caption = CodTotalizador
                            DbcCliente.Enabled = True
                            CmdCli.Enabled = True
                            LblData.Caption = Date
                            DtData.Value = Date
                            TxtSenha = ""
                            TxtSenha = LerINI("Config", "Senha", App.Path & "\config.ini")
                            If TxtSenha <> "" Then
                                If TxtTmp2 = "" Then
                                    TxtTmp2 = TxtSenha
                                    EscreveINI "Config", "Senha", TxtSenha + 1, App.Path & "\config.ini"
                                End If
                            End If
'                            DtData   Month(Date)
                            
 '                           LblHora.Caption = Time
                            LblMaquina.Caption = Maquina
                            
                            If Barra = "" Then
                                TxtDados(0).Text = ""
                            Else
                                TxtDados(0).Text = Barra
                                SendKeys ("{ENTER}")
                                Barra = ""
                            End If
                            If Cartao <> "" Then
                                LeCartao Cartao
                                Exit Sub
                            End If
                                                        
                            TxtDados(2).Text = ""
                            TxtDados(0).SetFocus
                        End If
                    Else
                        TxtDados(0) = ""
                        TxtDados(0).SetFocus
                    End If
            End If
        End If
    Case 2
        If KeyAscii = 27 Then
            TxtDados(2) = ""
            TxtDados(0) = ""
            TxtDados(0).SetFocus
        End If
        
        
        If KeyAscii = 13 And TxtDados(2) <> "" Then
            If Format(Left(TxtDados(2).Text, 1), ">") = "G" Or Format(Left(TxtDados(2).Text, 1), ">") = "/" Then
                VarGelada = Format(Left(TxtDados(2).Text, 1), ">")
                TxtDados(2).Text = Mid(TxtDados(2), 2, Len(TxtDados(2)) - 1)
            End If
            
            If Format(Left(TxtDados(2).Text, 1), ">") = "B" Then
                VarGelada = Format(Left(TxtDados(2).Text, 1), ">")
                TxtDados(2).Text = Mid(TxtDados(2), 2, Len(TxtDados(2)) - 1)
            End If
                        
            If IsNumeric(TxtDados(2)) Then
                If VarPacote = True Then
                    If LePacote(TxtDados(2)) = False Then
                        EfetuaVenda TxtDados(2)
                    End If
                Else
                    EfetuaVenda (TxtDados(2))
                End If
            Else
                TxtDados(2) = ""
                TxtDados(2).SetFocus
            End If
        End If
End Select

Exit Sub
erro:
If Err.Number <> 0 Then
    MsgBox Err.Description & Err.Number, vbCritical, Caminho
End If

End Sub

Private Sub IncluirLinha(ValorVenda As String)

Dim Retorno As Integer
Dim TESTE As String
Dim TESTE2 As String
Dim Produto As String
Dim CodigoBarra As String
Dim Status As String
Dim sql As String
Dim Qtd As String
Dim Vlr As String
Dim VlrApurado As Currency
Dim Criterio As String
Dim Flag As Boolean
Dim VarSeven As Recordset
Dim VlrSeven As Currency
Dim DescSeven As Currency
'Dim Criterio As String


'On Error GoTo erro1

If LerINI("DescontoFrente", "UltimoPreco", LocalBanco & "\config.ini") = 1 Then
    
    If Pedido.RecordCount > 0 Then
        Pedido.MoveLast
        Criterio = "Entidade = " & CodCli
        Pedido.FindPrevious Criterio
        If Not Pedido.NoMatch Then
            Flag = False
            While Not Pedido.BOF And Flag = False
                If RecVendas.RecordCount > 0 Then
                    Criterio = "Pedido = " & Pedido!CodPed & " and Produto = " & TxtDados(2)
                    RecVendas.FindFirst Criterio
                    If Not RecVendas.NoMatch Then
                        TxtUltimo = Format(RecVendas!ValorUnit, "#,##0.00")
                        Flag = True
                    End If
                End If
                Pedido.MovePrevious
            Wend
        End If
    End If
End If

Set Empresa = Banco.OpenRecordset("CFG")
If TxtDados(0) <> "" Then
    If Len(TxtDados(0)) >= 7 Then
        quantidade = 1
    Else
        quantidade = CDec(TxtDados(0))
        If quantidade = 0 Then
            quantidade = 1
        End If
    End If
Else
    quantidade = 1
End If

'MebPreco = Round(Valor, 3)
If IsNull(Registros("Obs")) Then
    TxtObs.Text = ""
Else
    TxtObs.Text = Registros("Obs")
End If

If Len(Trim(ValorVenda)) > 6 And Registros.Fields("Venda") <> 0 Then
    If Left(ValorVenda, 1) = 2 Then
            VlrApurado = (CCur(Mid(ValorVenda, 8, 5)) / 100)
            quantidade = Format(CCur(VlrApurado), "#,##0.00") / Format(Registros.Fields("Venda"), "#,##0.00")
            quantidade = Format(quantidade, "#,##0.000")

'            quantidade = (CCur(Mid(ValorVenda, 8, 5)) / 1000)
'            VlrApurado = Format(CCur(quantidade), "#,##0.00") * Format(Registros.Fields("Venda"), "#,##0.00")
'            quantidade = Format(quantidade, "#,##0.000")


'            Quantidade = Format(CStr((CCur(Mid(ValorVenda, 8, 5)) / 100) / Registros.Fields("Preco")), "#,##0.000")
'            MebUnit.Caption = Truncar(Registros.Fields("Preco") * Quantidade, "##,##0.00", 2)
'            MebUnit.Caption = Format(Truncar(Registros.Fields("Preco") * Quantidade, 2), "##,##0.00")
'            MebUnit = Truncar(MebUnit, 2)
    End If
End If

MebDesconto = 0

If Vista Then
    If Registros("Promocao") Then
'        MebDesconto = CCur(Registros("Margem_promocao"))
        Valor = CCur(Registros("Venda_Promocao"))
    Else
        Valor = CCur(Registros("Venda"))
    End If
Else
    If VarPcoPrazo = True Then
            If IsNull(Registros("PcoPrazo")) Then
                Valor = CCur(Registros("Venda"))
            Else
                Valor = CCur(Registros("PcoPrazo"))
'                MebDesconto = CCur(Registros("Margem_promocao"))
            End If
    Else
        If Indices.RecordCount > 0 Then
            Valor = CCur(Registros("Venda")) + (CCur(Registros("Venda")) * (CCur(Indices("Perc"))) / 100)
        Else
            Valor = CCur(Registros("Venda"))
        End If
    End If
    
End If

If Not IsNull(Registros("Dtalt")) Then
    MskData.Text = Registros("DtAlt")
Else
    MskData.Text = ""
End If

If LblCartao = "BPMS" Then
       
       Set VarSeven = Banco.OpenRecordset("SELECT * from SevenPdv Where Pedido = " & """" & LblPedido & """")
       VarSeven.Requery
       If VarSeven.RecordCount > 0 Then
            Criterio = "CodBar = " & """" & Registros!BarBarra & """"
            VarSeven.FindFirst Criterio
            If Not VarSeven.NoMatch Then
                DescSeven = Registros!Venda * (VarSeven!Desconto / 100)
                VlrSeven = Registros!Venda - DescSeven
                If VlrSeven < VarSeven!Liquido Then
                    Valor = VlrSeven
                Else
                    Valor = VarSeven("Liquido")
                End If
            End If
       End If
End If


If Gelado = False Then
    If VarGelada = "G" Or VarGelada = "/" Then
        Valor = IIf(IsNull(Round(Registros.Fields("Venda_Concorrencia"), 2)), 0, Format(Registros.Fields("Venda_Concorrencia"), "###,##0.00"))
    End If
Else
    If VarGelada = "" Then
        Valor = IIf(IsNull(Round(Registros.Fields("Venda_Concorrencia"), 2)), 0, Format(Registros.Fields("Venda_Concorrencia"), "###,##0.00"))
    End If
End If

If VarGelada = "B" Then
       Valor = IIf(IsNull(Round(Registros.Fields("Venda_PROMOCAO"), 2)), 0, Format(Registros.Fields("Venda_PROMOCAO"), "###,##0.00"))
End If

If NatOper = 3 Then
       Valor = IIf(IsNull(Round(Registros.Fields("Margem_Concorrencia"), 2)), Format(Registros.Fields("Venda"), "###,##0.00"), Format(Registros.Fields("Margem_Concorrencia"), "###,##0.00"))
       If Valor = 0 Then
            Valor = Format(Registros.Fields("Venda"), "###,##0.00")
       End If
            
End If


If VarLista1.RecordCount > 0 Then
    VarLista1.FindFirst "CodPrd = " & ValorVenda
    If Not VarLista1.NoMatch Then
        Valor = Round(VarLista1!Valor, 3)
    End If
End If

MebPreco = Round(CCur(Valor), 2)

'TxtComplemento = ""
If VarComplemento = "Sim" Then
      DbcLoja = VarLocalEst

     txtcomplemento.SetFocus
     Exit Sub
End If

If CLoja.RecordCount > 1 And VarLocalEst = "" Then
    DbcLoja.Enabled = True
    DatLoja.RecordSource = "SELECT * FROM LOJ  WHERE Produto = " & Registros.Fields("CodPrd") & " ORDER BY Descricaoloja"
    DatLoja.Refresh
    If DatLoja.Recordset.RecordCount <> 0 Then
        DatLoja.Recordset.MoveFirst
        DbcLoja.Clear
        While Not DatLoja.Recordset.EOF
            DbcLoja.AddItem DatLoja.Recordset.Fields("DescricaoLoja") & "  -  " & DatLoja.Recordset("Estoque")
            DatLoja.Recordset.MoveNext
        Wend
        DbcLoja.ListIndex = 0
       DbcLoja.SetFocus
        Exit Sub
    End If
Else
    DbcLoja = VarLocalEst
    DbcLoja.Enabled = False
End If

DbcLoja = VarLocalEst
'DbcLoja.Enabled = False

If Valor = 0 Then
    If VarDiversos = 0 Then
        FrmValor.Show vbModal
        If ValorUnit <> "" Then
        Valor = ValorUnit
        End If
    Else
        Valor = VarDiversos / 100
    End If
End If


If CfgDesc = "Sim" Or CfgDesc = 1 Then
     MebDesconto.SetFocus
'    Venda
Else
    MebPreco = Round(CCur(Valor), 2)
    Venda
End If


End Sub

Private Sub Venda()
Dim ConvertVr As String
Dim ConvertQtd As String
Dim VarLoja As String
Dim F As Byte
'If (Registros("Estoque") - TxtDados(0)) < Registros("Mínimo") Then
'    MsgBox "Estoque abaixo mínimo. Estoque = " & Registros("Estoque") & "  Mínimo = " & Registros("Mínimo"), vbInformation, App.Title
'End If

Valor = MebPreco

MebUnit.Caption = Format(MebPreco, "##,##0.00")

MebVlrTot = MebPreco * quantidade

If TxtDados(0).Text <> "" And LblDescricao.Caption <> "" Then

' ***************
'    If IsNull(Registros.Fields("BarBarra")) Then
'        CodigoBarra = "0"
'    Else
'        CodigoBarra = Registros.Fields("BarBarra")
'    End If
    
'    Produto = LimpaPrd()
    ConvertQtd = Format(quantidade, "000.000")
    ConvertVr = Format(Valor, "0,000,000.000")

'*************************
'    Soma = Soma + Truncar(Format(Valor, "##,##0.00") * Quantidade, 2)
    Soma = Soma + Arredondamento((Format(Valor, "##,##0.000") * quantidade), 2)
'*************************
    
    MebTotal = Format(Soma, "##,##0.00")
        
    RecVendas.AddNew
    RecVendas("Pedido") = CodTotalizador

    RecVendas("Produto") = Registros.Fields("CodPrd")
    RecVendas("Qtd") = Format(quantidade, " ##,##0.000")
    RecVendas("ValorUnit") = Valor
    RecVendas("Valor") = Arredondamento(Format(Valor, "##,##0.000") * Format(quantidade, "##,##0.000"), 2)
    RecVendas("VlrVdaDia") = Arredondamento(Format(Valor, "##,##0.000") * Format(quantidade, "##,##0.000"), 2)
    RecVendas("CustoDia") = Registros("Custo")
    RecVendas("VendaDia") = Registros("Venda")
    If Round((Registros("Venda") * quantidade), 2) > Round(Valor * quantidade, 2) Then
        RecVendas("Desconto") = (Registros("Venda") * quantidade) - (Valor * quantidade)
        VarPrdDesc = VarPrdDesc + (Valor * quantidade)
    End If
    
    If Registros.Fields("Comissao") <> 0 Then
        RecVendas("Comissao") = RecVendas("Comissao") * (Format(Valor, "##,##0.000") * Format(quantidade, "##,##0.000"))
    End If
    
    VarLoja = ""
    For F = 1 To Len(DbcLoja.Text)
        If Mid$(DbcLoja, F, 1) <> Chr(32) Then
            VarLoja = VarLoja + Mid$(DbcLoja, F, 1)
        Else
            F = Len(DbcLoja.Text)
        End If
    Next
    
    RecVendas("Loja") = VarLoja
    RecVendas("Entregue") = 0
    RecVendas("Complemento") = txtcomplemento
    If Not IsNull(Registros("Venda_Concorrencia")) Then
        If IsNumeric(Registros("Venda_Concorrencia")) Then
            RecVendas("Comissao") = Format((Valor * quantidade) * (Registros("Venda_Concorrencia") / 100), "##,##0.00")
        End If
    End If
    LblItens.Caption = RecVendas("CodItp")
    RecVendas("Comissao") = (Valor * quantidade) * (Registros("Comissao") / 100)
    RecVendas.Update
        
    VarSomadeCusto = VarSomadeCusto + (Registros("Custo") * Format(quantidade, "##,##0.000"))
    
'    DatCompra.Refresh
'    DatCompra.Recordset.MoveLast
    ContList = ContList + 1
    Set LstPesq = ListVda.ListItems.Add(, , ContList)
    LstPesq.SubItems(2) = ContList
    LstPesq.SubItems(1) = Format(quantidade, " ##,##0.000") 'SeCampoNulo(Consulta, "Agencia")
    LstPesq.SubItems(2) = Registros.Fields("CodPrd")
    If IsNull(Registros.Fields("BarBarra")) Then
        LstPesq.SubItems(3) = 0
    Else
        LstPesq.SubItems(3) = Registros.Fields("BarBarra")
    End If
    LstPesq.SubItems(4) = LblDescricao  'SeCampoNulo(Consulta, "Agencia")
    LstPesq.SubItems(5) = Format(Valor, " ##,##0.000")  'SeCampoNulo(Consulta, "Conta")
    LstPesq.SubItems(6) = Format(Format(Valor, "##,##0.000") * Format(quantidade, "##,##0.000"), "##,##0.00") 'SeCampoNulo(Consulta, "DataOper")
    LstPesq.SubItems(7) = LblItens
    totlista.Caption = ContList
    
'    For X = 1 To ListVda.ListItems.Count
'        DoEvents
    
    ListVda.ListItems(ContList).EnsureVisible
'    ListVda.ListItems(ContList).Bold = True ' = &HFF&
    If (Registros("Estoque") - quantidade) < 0 Then
        TxtEstoque.ForeColor = &HFF&
    Else
        TxtEstoque.ForeColor = &H80000008
'        MsgBox "Estoque abaixo mínimo. Estoque = " & Registros("Estoque") & "  Mínimo = " & Registros("Mínimo"), vbInformation, App.Title
    End If
    TxtEstoque = (Registros("Estoque"))

End If

TxtDados(0) = ""
TxtDados(2) = ""
'TxtEstoque = ""
MebDesconto = 0
MebPreco = 0
MebVlrTot = 0
txtcomplemento = ""
TxtObs = ""
'TxtEstoque = ""
TxtUltimo = ""
'Valor = 0
TxtDados(0).SetFocus
VarDiversos = 0

End Sub

Private Sub EfetuaVenda(Valor As String)
Dim Terminação As String
Terminação = """"
TxtEstoque = ""
Valor = LimpaNro(Valor)

If Len(Trim(Valor)) >= 7 Then
    If Left(Valor, 1) = 2 Then
        Criterio = Registros.Fields("BarBarra").Name & " = "
        If VarTamBarra = 4 Then
            Criterio = Criterio & Terminação & CDbl(Mid(Valor, 2, VarTamBarra)) & Terminação
        ElseIf VarTamBarra = 5 Then
            Criterio = Criterio & Terminação & CDbl(Mid(Valor, 2, VarTamBarra)) & Terminação
        Else
            Criterio = Criterio & Terminação & CDbl(Mid(Valor, 2, VarTamBarra)) & Terminação
        End If
    ElseIf Left(Valor, 1) = "0" And Len(Valor) < 9 Then
            Criterio = Registros.Fields(1).Name & " = "
            Criterio = Criterio & Valor
        Else
            Criterio = Registros.Fields("BarBarra").Name & " = "
            Criterio = Criterio & Terminação & (Valor) & Terminação
    End If
    
    If Left(Valor, 1) = 9 And Len(Valor) < 9 Then
        Valor = CDbl(Right(Valor, 9))
        Criterio = Registros.Fields(1).Name & " = "
        Criterio = Criterio & Valor
    End If
    
    If CDbl(Mid(Valor, 2, 6)) = "000000" And CDbl(Mid(Valor, 1, 1)) <> "1" Then
        VarDiversos = CDbl(Mid(Valor, 8, 5))
        Criterio = Registros.Fields(1).Name & " = "
        Criterio = Criterio & 1
    Else
        VarDiversos = 0
    End If
    
Else
    Criterio = Registros.Fields(1).Name & " = "
    Criterio = Criterio & CDbl(Valor)
End If
Registros.FindFirst Criterio
If Registros.NoMatch Then
    MsgBox "Informação não Localizada", vbExclamation, App.Title
Else
    If Registros!complemento <> "" Then
        txtcomplemento = Registros!complemento
    End If

    LblDescricao.Caption = Registros.Fields("Descrição_Produto").Value
    If txtcomplemento <> "" Then
        LblDescricao.Caption = LblDescricao.Caption & "-" & txtcomplemento
    End If
'    MebUnit.Caption = Registros.Fields("Preco").Value

    IncluirLinha (Valor)
End If
End Sub

Private Sub LeCartao(Cliente As String)
Dim VarVenda As Recordset

       CodCard = Mid(Cliente, 2, 10)
       
'       Set VarVenda = Banco.OpenRecordset("Select * From Conv Where Cliente = " & Mid(Cliente, 2, 10) & " AND Totaliza = 0 and Flag = -1 and Cancelado = 0 ")
       Set VarVenda = Banco.OpenRecordset("SELECT CONV.Produto, CONV.Cliente, CONV.Flag, CONV.Totaliza, Sum(CONV.Qtd) AS SomaDeQtd FROM CONV GROUP BY CONV.Produto, CONV.Cliente, CONV.Flag, CONV.Totaliza, CONV.Cancelado HAVING (((CONV.Cliente)=" & Mid(Cliente, 2, 10) & ") AND ((CONV.Flag)=-1) AND ((CONV.Totaliza)=0) AND ((CONV.Cancelado)=0))")
       VarVenda.Requery
       If VarVenda.RecordCount > 0 Then
            VarVenda.MoveFirst
            While Not VarVenda.EOF
       
                Criterio = "CodPrd = " & VarVenda!Produto
                Registros.FindFirst Criterio
                TxtTmp = Registros("Venda")
                TxtDados(0).Text = VarVenda!SomaDeQtd
                TxtDados(0).SetFocus
                TxtDados_KeyPress 0, 13
        
                TxtDados(2).Text = VarVenda!Produto
                TxtDados(2).SetFocus
                TxtDados_KeyPress 2, 13
                VarVenda.MoveNext
             Wend
        Else
            MsgBox "Cartão não tem movimento !", vbInformation, App.Title
        End If
        TxtDados(0) = ""
        TxtDados(0).SetFocus
End Sub

Public Sub AtuMesa()
Dim x As Integer
Frame5.Visible = True
If FlagFrame = True And (MebTotal = "" Or MebTotal = "0" Or LblDescricao = "TROCO") Then
            
        Set VarMesa = Banco.OpenRecordset("SELECT CONV.Funcionario, CONV.Cliente, CONV.Totaliza, CONV.Cancelado, flag FROM CONV GROUP BY CONV.Funcionario, CONV.Cliente, CONV.Totaliza, CONV.Cancelado, flag HAVING (((CONV.Totaliza)=False) AND ((CONV.Cancelado)=False)) and Flag = true ORDER BY CONV.Cliente")
        VarMesa.Requery
        For x = 0 To 203
'            If Mesa(x).BackColor = &HC0C0C0 Then
                Mesa(x).Caption = x + 1
                Mesa(x).BackColor = &HFFFFFF
                Mesa(x).Visible = True
'            Else
'                Mesa(x).Caption = ""
'                Mesa(x).BackColor = &HFFFFFF
'                x = 299
'            End If
        Next
        x = 0
        If VarMesa.RecordCount > 0 Then
            VarMesa.MoveFirst
            x = 0
            While Not VarMesa.EOF
'                Mesa(VarMesa!Cliente + 1).Visible = True
                Mesa(VarMesa!Cliente - 1).Caption = VarMesa!Cliente
                If VarMesa!Funcionario = 99999 Then
                    Mesa(VarMesa!Cliente - 1).BackColor = &HFF&
                Else
                    Mesa(VarMesa!Cliente - 1).BackColor = &HC0C0C0
                End If
'                x = x + 1
                VarMesa.MoveNext
            Wend
        End If
'        LblOcupadas.Caption = x
'        LblLivre.Caption = 300 - x
        If VarConveniencia = True Then
            Frame5.Visible = True
'            Frame5.Left = 0
'            Frame5.Top = 10300
        Else
            If VarLanchonete = True Then
                Frame5.Visible = True
'                Frame5.Left = 0
'                Frame5.Top = 540
            End If
        End If
        
    '    TxtDados(0).SetFocus
        Timer2.Enabled = True
Else
'    Frame5.Visible = False
End If
End Sub


Private Sub RemoveMenus()
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION

End Sub

Private Sub LerSeven()

Dim VarSeven As Recordset
Dim Criterio As String
       
       Set VarSeven = Banco.OpenRecordset("SELECT * from SevenPdv Where Pedido = " & """" & LblPedido & """")
       VarSeven.Requery
       If VarSeven.RecordCount > 0 Then
            VarSeven.MoveFirst
            While Not VarSeven.EOF
       
                Criterio = "BarBarra = " & """" & VarSeven!Codbar & """"
                Registros.FindFirst Criterio
                If Not Registros.NoMatch Then
                    TxtTmp = Registros("Venda")
                    TxtDados(0).Text = VarSeven!Qtd
                    TxtDados(0).SetFocus
                    TxtDados_KeyPress 0, 13
            
                    TxtDados(2).Text = Registros("Produto")
                    TxtDados(2).SetFocus
                    TxtDados_KeyPress 2, 13
                End If
                VarSeven.MoveNext
             Wend
        Else
            MsgBox "Não tem produtos !", vbInformation, App.Title
        End If
        TxtDados(0) = ""
        TxtDados(0).SetFocus

End Sub

Function LePacote(VarProduto As String) As Boolean
Dim VarVenda As Recordset
Dim y As Byte
y = 0


'       Set VarVenda = Banco.OpenRecordset("Select * From Conv Where Cliente = " & Mid(Cliente, 2, 10) & " AND Totaliza = 0 and Flag = -1 and Cancelado = 0 ")
       Set VarVenda = Banco.OpenRecordset("SELECT * FROM Comp WHERE Produto = " & VarProduto)
       VarVenda.Requery
'       TxtComplemento = ""
       If VarVenda.RecordCount > 0 Then
            VarVenda.MoveFirst
            While Not VarVenda.EOF
       
                Criterio = "CodPrd = " & VarVenda!PrdComp
                Registros.FindFirst Criterio
                For y = 0 To VarVenda!Qtd
                    txtcomplemento = VarVenda("Complemento")
                    
                    TxtTmp = Registros("Venda")
                    TxtDados(0).Text = 1
                    TxtDados(0).SetFocus
                    TxtDados_KeyPress 0, 13
    
                    TxtDados(2).Text = VarVenda!PrdComp
                    TxtDados(2).SetFocus
                    TxtDados_KeyPress 2, 13
                    y = y + 1
                Next
                
                VarVenda.MoveNext

             Wend
             LePacote = True
        Else
'            MsgBox "Produto Sem Composição !", vbInformation, App.Title
             LePacote = False
        End If
'        TxtDados(0) = ""
'        TxtDados(0).SetFocus
End Function

