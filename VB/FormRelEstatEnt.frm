VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormRelEstatEnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6600
   Begin VB.Data DatDados 
      Caption         =   "Entidade"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *  FROM ENT WHERE  Tipo =  3 Or Tipo = 6 ORDER BY Nome"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -90
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   14
      Top             =   4140
      Width           =   11535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13155
      TabIndex        =   12
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatística de Clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   135
         TabIndex        =   13
         Top             =   45
         Width           =   6270
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ordem"
      Height          =   1620
      Left            =   810
      TabIndex        =   7
      Top             =   2385
      Width           =   2085
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Código"
         Height          =   225
         Index           =   3
         Left            =   270
         TabIndex        =   11
         Tag             =   "CodPrd"
         Top             =   255
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Descrição"
         Height          =   225
         Index           =   4
         Left            =   270
         TabIndex        =   10
         Tag             =   "Descrição_Produto"
         Top             =   562
         Width           =   1305
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Valor"
         Height          =   225
         Index           =   5
         Left            =   270
         TabIndex        =   9
         Tag             =   "Sum(ITP.Qtd)"
         Top             =   869
         Width           =   1485
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Ultima Compra"
         Height          =   225
         Index           =   6
         Left            =   270
         TabIndex        =   8
         Tag             =   "Sum(ITP.VlrVdaDia)"
         Top             =   1176
         Width           =   1605
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   420
      Left            =   3885
      TabIndex        =   6
      Top             =   2655
      Width           =   1300
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   420
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir"
      Top             =   3270
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   705
      Left            =   795
      TabIndex        =   0
      Top             =   750
      Width           =   5055
      Begin MSComCtl2.DTPicker DtData1 
         Height          =   330
         Left            =   1095
         TabIndex        =   1
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   87687169
         CurrentDate     =   36641
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   330
         Left            =   3495
         TabIndex        =   2
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   87687169
         CurrentDate     =   36641
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   255
         Left            =   2625
         TabIndex        =   4
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   795
      End
   End
   Begin MSDBCtls.DBCombo DbcEnt 
      Bindings        =   "FormRelEstatEnt.frx":0000
      Height          =   360
      Left            =   765
      TabIndex        =   15
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor"
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
      Left            =   765
      TabIndex        =   16
      Top             =   1530
      Width           =   900
   End
End
Attribute VB_Name = "FormRelEstatEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CamposSql As String
Dim CondicaoWhere As String
Dim SqlGroup As String
Dim Ordenado As String
Dim SQL As String
Dim Dif As Currency

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim Periodo As String

Periodo = "Período de " & DtData1 & " a " & DtData2
    If DbcEnt = "" Then
        CamposSql = " SELECT  PED.Entidade, ENT.Nome, Ent.Fone, Ent.UF, Sum(PED.Total) AS SomaDeTotal, ENT.Cidade, Last(PED.DataEmiss) AS ÚltimoDeDataEmiss FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade  "
    Else
        CamposSql = " SELECT Ped.Vendedor, PED.Entidade, ENT.Nome, Ent.Fone, Ent.UF, Sum(PED.Total) AS SomaDeTotal, ENT.Cidade, Last(PED.DataEmiss) AS ÚltimoDeDataEmiss FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade  "
    End If

    If OptOrdem(3).value = True Then
        Ordenado = " ORDER BY Ped.Entidade "
    ElseIf OptOrdem(4).value = True Then
            Ordenado = " ORDER BY Ent.Nome "
        ElseIf OptOrdem(5).value = True Then
                Ordenado = " Order by Sum(PED.Total) DESC "
            ElseIf OptOrdem(6).value = True Then
                Ordenado = " Order by Last(PED.DataEmiss) "
    End If
    
'    CondicaoWhere = " WHERE (((ENT.Tipo)=1 Or (ENT.Tipo)=5 Or (ENT.Tipo)=6) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY PED.Entidade, ENT.Nome, ENT.Cidade HAVING (((Last(PED.DataEmiss))>=#" & Format(DtData1, "mm/dd/yyyy") & "# And (Last(PED.DataEmiss))<=#" & Format(DtData2, "mm/dd/yyyy") & "#))"
    If DbcEnt = "" Then
        CondicaoWhere = " Where (((Ped.DataEmiss) >= #" & Format(DtData1, "mm/dd/yyyy") & "# And (Ped.DataEmiss) <= #" & Format(DtData2, "mm/dd/yyyy") & "#) And ((ENT.tipo) = 1 Or (ENT.tipo) = 5 Or (ENT.tipo) = 6) And ((Ped.Operacao) = 1 Or (Ped.Operacao) = 2 Or (Ped.Operacao) = 3)) GROUP BY  PED.Entidade, ENT.Nome, ENT.Cidade, Ent.Fone, Ent.UF "
    Else
        CondicaoWhere = " Where (((Ped.DataEmiss) >= #" & Format(DtData1, "mm/dd/yyyy") & "# And (Ped.DataEmiss) <= #" & Format(DtData2, "mm/dd/yyyy") & "#) And ((ENT.tipo) = 1 Or (ENT.tipo) = 5 Or (ENT.tipo) = 6) And ((Ped.Operacao) = 1 Or (Ped.Operacao) = 2 Or (Ped.Operacao) = 3)) and Ped.Vendedor = " & DbcEnt.BoundText & "  GROUP BY Ped.Vendedor, PED.Entidade, ENT.Nome, ENT.Cidade, Ent.Fone, Ent.UF "
    End If
   
    SQL = CamposSql & CondicaoWhere & Ordenado
    
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set FrmRelEstatEnt.Relatorio.Recordset = varPesquisa(1)
    FrmRelEstatEnt.rpCabTitulo.Caption = "Relatório de Estatistica de Clientes"
    FrmRelEstatEnt.rpCabPeriodo.Caption = Periodo
    FrmRelEstatEnt.rpCabSubTitulo.Caption = NomeEmpresa
    FrmRelEstatEnt.Config

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DtData1.value = Date - 30
DtData2.value = Date

For Each Componentes In DatDados
    Componentes.DatabaseName = Caminho & "\Dados.mdb"
Next

End Sub

