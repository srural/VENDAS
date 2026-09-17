VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmPedTran 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transportadora"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmbCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5670
      TabIndex        =   20
      Top             =   2025
      Width           =   1215
   End
   Begin VB.CommandButton BtOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4275
      TabIndex        =   19
      Top             =   2025
      Width           =   1215
   End
   Begin VB.ComboBox CmbDados 
      DataField       =   "Status"
      Height          =   315
      Index           =   3
      ItemData        =   "FrmPedTran.frx":0000
      Left            =   5640
      List            =   "FrmPedTran.frx":000A
      TabIndex        =   5
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6735
      Begin VB.Data DatDados 
         Connect         =   "Access"
         DatabaseName    =   "C:\Minha Pasta\Programas\Vendas\Dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   2025
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select * From Ent Where Tipo = 1 or Tipo = 5 Or Tipo = 4 Or Tipo = 7 Order by Nome"
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBCtls.DBCombo DbcDados 
         Bindings        =   "FrmPedTran.frx":0014
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   1
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
   End
   Begin MSMask.MaskEdBox MebDados 
      Height          =   315
      Index           =   4
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "CCC-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDados 
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDados 
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDados 
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDados 
      Height          =   315
      Index           =   3
      Left            =   5280
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      Caption         =   "Marca"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   18
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Espécie"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Qtd"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo do Frete"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "UF"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Placa"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "FrmPedTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Extra As Recordset

Private Sub BtOK_Click()

If Extra.RecordCount <> 0 Then
'    FrmPedidos.Tab
    Extra.Edit
    Extra("PlacaVeiculo") = MebDados(4).Text
    Extra("QtdVolume") = MebDados(1).Text
    Extra("Especie") = MebDados(2).Text
    Extra("Marca") = MebDados(3).Text
    Extra("Uf") = MebDados(0).Text
    Extra("Tipo") = CmbDados(3).Text
    Extra("Transportadora") = DbcDados(0).BoundText
    Extra.Update
Else
    Unload Me
'    FrmPedidos.Show
End If

End Sub

Private Sub CmbCancelar_Click()
Unload Me
End Sub

Private Sub dbcdados_Change(Index As Integer)
Dim Criterio As String
If Index = 0 Then
    Criterio = "CodEntidade = " & DbcDados(0).BoundText
    DatDados(0).Recordset.FindFirst Criterio
    Label(0).Caption = DatDados(0).Recordset.Fields("Endereco") & " - " & DatDados(0).Recordset.Fields("Nro")
    Label(1).Caption = DatDados(0).Recordset.Fields("Cidade")
End If
End Sub

Private Sub Form_Activate()
Dim Criterio As String

Set Extra = Banco.OpenRecordset("Select * From  PEd Where CodPed = " & varCodPed)

End Sub

Private Sub Form_Load()
For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\dados.mdb"
Next
DatDados(0).RecordSource = "Select * From Ent Where Tipo = 3 or Tipo = 5 Order by Nome"



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
        If KeyCode = 13 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
        If KeyCode = 9 Then
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 20
        End If
End Sub

