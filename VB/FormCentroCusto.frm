VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormCentroCusto 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   10
      Top             =   7695
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11715
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Custo"
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
         TabIndex        =   9
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exporta LayOut"
      Height          =   735
      Left            =   4185
      TabIndex        =   2
      Top             =   6840
      Width           =   7170
      Begin VB.CommandButton CmdExporta 
         Caption         =   "&Exporta"
         Height          =   330
         Left            =   5130
         TabIndex        =   7
         Top             =   315
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker DtData 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   300
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52953089
         CurrentDate     =   36759
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   315
         Left            =   3105
         TabIndex        =   5
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52953089
         CurrentDate     =   36759
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Final"
         Height          =   195
         Index           =   0
         Left            =   2655
         TabIndex        =   6
         Top             =   330
         Width           =   330
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   795
      End
   End
   Begin VB.CommandButton CmbExclui 
      Caption         =   "E&xclui"
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Top             =   7020
      Width           =   1140
   End
   Begin VB.Data Datdados 
      Caption         =   "Centro de Custo"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Desenvolvimento\Vendas_1_8\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * FROM CCu ORDER BY Codigo"
      Top             =   3690
      Visible         =   0   'False
      Width           =   4290
   End
   Begin MSDBGrid.DBGrid DbgIndices 
      Bindings        =   "FormCentroCusto.frx":0000
      Height          =   6045
      Left            =   180
      OleObjectBlob   =   "FormCentroCusto.frx":001A
      TabIndex        =   0
      Top             =   765
      Width           =   11220
   End
End
Attribute VB_Name = "FormCentroCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbExclui_Click()
If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
    If DatDados(0).Recordset.RecordCount <> 0 Then
        DatDados(0).Recordset.Delete
    End If
End If
End Sub


Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

If DatDados(0).Recordset.RecordCount > 0 Then
    DatDados(0).Recordset.MoveLast
End If

DtData = Date
DtData2 = Date

End Sub

Private Sub Form_Load()
Dim Componentes As Variant

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next
End Sub

Private Sub CmdExporta_Click()
Dim VarCusto As Recordset

Set VarCusto = Banco.OpenRecordset("Select * From CCU Where Data >= #" & Format(DtData, "mm/dd/yyyy") & "# AND Data <= #" & Format(DtData2, "mm/dd/yyyy") & "# ORDER BY Codigo")
VarCusto.Requery

If VarCusto.RecordCount > 0 Then
    
    If MsgBox("Confirma exportação ?", vbYesNo, App.Title) = vbYes Then
        VarCusto.MoveFirst
        Open Caminho & "\CentroCusto.txt" For Output As #1
            While Not VarCusto.Eof
                
                Print #1, VarCusto!Codigo & ";" & VarCusto!Filial & ";" & VarCusto!Data & ";" & _
                          VarCusto!Tipo & ";" & VarCusto!Historico & ";" & VarCusto!CentroCusto & ";" & _
                          VarCusto!Conta & ";" & VarCusto!NrOp & ";" & Numero2(VarCusto!Valor)
                
                VarCusto.MoveNext
            Wend
        Close #1
        MsgBox "Arquivo gerado com sucesso !", vbInformation, App.Title
    End If
Else
    MsgBox "Arquivo sem movimento !", vbInformation, App.Title
    
End If

End Sub

