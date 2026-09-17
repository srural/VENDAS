VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormEstatVend 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9180
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -360
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   8
      Top             =   6840
      Width           =   14370
   End
   Begin MSFlexGridLib.MSFlexGrid flxGrid 
      Height          =   4290
      Left            =   180
      TabIndex        =   7
      Top             =   1665
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   7567
      _Version        =   393216
      Rows            =   12
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9195
      TabIndex        =   4
      Top             =   0
      Width           =   9195
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatistica de Vendedores"
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
         TabIndex        =   5
         Top             =   45
         Width           =   6630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedor"
      Height          =   780
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   8955
      Begin VB.Data DatVend 
         Caption         =   "Vendedores"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1395
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Select CodEntidade, Nome, Tipo From Ent Where Tipo = 3 or Tipo = 7"
         Top             =   270
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton BtRel 
         Caption         =   "&Imprimir"
         Height          =   330
         Left            =   7110
         TabIndex        =   3
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton BtVisualiza 
         Caption         =   "&Visualiza"
         Height          =   330
         Left            =   5535
         TabIndex        =   0
         Top             =   225
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo DbcVend 
         Bindings        =   "FormEstatVend.frx":0000
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nome"
         BoundColumn     =   "CodEntidade"
         Text            =   ""
      End
   End
   Begin VB.ListBox LstVend 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   6165
      Visible         =   0   'False
      Width           =   6120
   End
End
Attribute VB_Name = "FormEstatVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Venda As Recordset

Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub BtCalcula_Click()
Dim MesAnt As String
Dim AnoAnt As String
Dim VarData As Date


VarData = Date - 1095

If DbcVend.Text <> "" Then
                    
    SQL = "SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Year(ped.dataemiss) AS Expr1, ENT.CodEntidade, Month(ped.dataemiss) AS Expr2 FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)>=#" & Format(VarData, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(Date, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ENT.CodEntidade, ENT.Nome, Year(ped.dataemiss), ENT.CodEntidade, Month(ped.dataemiss) HAVING (((ENT.CodEntidade)= " & DbcVend.BoundText & " )) ORDER BY ENT.Nome, Month(ped.dataemiss), Year(ped.dataemiss)"
    Set Venda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Year(ped.dataemiss) AS Expr1, ENT.CodEntidade, Month(ped.dataemiss) AS Expr2 FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)>=#" & Format(VarData, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(Date, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2) GROUP BY ENT.CodEntidade, ENT.Nome, Year(ped.dataemiss), ENT.CodEntidade, Month(ped.dataemiss) HAVING (((ENT.CodEntidade)= " & DbcVend.BoundText & " )) ORDER BY ENT.Nome, Month(ped.dataemiss), Year(ped.dataemiss)")

    Venda.Requery
    If Venda.RecordCount > 0 Then
        Venda.MoveFirst
        Open App.Path & "\EstatVen.txt" For Output As #1
        Print #1, ""
        Print #1, Centralizar("ESTATISTICA DE VENDEDOR", 80)
        Print #1, "======================================================================================"
        Print #1, "Nome: " & Venda!Nome
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------------"
        Print #1, "  Mês/Ano          "; Year(Date); Space(15); Year(Date) - 1; Space(15); Year(Date) - 2
        Print #1, "--------------------------------------------------------------------------------------"
        MesAnt = Venda!Expr2
        AnoAnt = Venda!Expr1
'        Print #1, Alinhar(Venda("SomaDeVlrVdaDia"), 12, 2)
        Print #1, Space(4); Venda!Expr2;
        While Not Venda.Eof
            If MesAnt = Venda!Expr2 Then
                Print #1, Space(7); Alinhar(Venda("SomaDeVlrVdaDia"), 12, 2); Space(3);
            Else
                Print #1, " "
                Print #1, Space(4); Venda!Expr2; IIf(Len(Trim(Venda!Expr2)) < 2, Space(7), Space(6)); Alinhar(Venda("SomaDeVlrVdaDia"), 12, 2); Space(3);
                MesAnt = Venda!Expr2
                AnoAnt = Venda!Expr1
            End If
            Venda.MoveNext
        Wend
        Print #1, " "
        Print #1, "------------------------------------------------------------------------------------------------"
        
        Print #1, "" & Nor
        Print #1, "<>"
    
        Close #1
        
        On Error GoTo Erro
        LstVend.Clear
        Open App.Path & "\EstatVen.txt" For Input As #1
        Do
            Line Input #1, Temp
            LstVend.AddItem Temp
    '        Lista.Selected(Lista.ListCount - 1) = True
        Loop Until Temp = "<>"
        
        Close #1
    End If
Else
    MsgBox "Vendedor não seleciondo ", vbInformation, App.Title
End If

Erro:
If Err.Number = 62 Then
    Close #1
End If

End Sub

Private Sub BtRel_Click()

    Dim rptGrid As rlt_FlexGrid
    
    If flxGrid.Rows > 0 Then
        Set rptGrid = New rlt_FlexGrid
        Set rptGrid.Grid = flxGrid
'        If opt_Ori(2).Value Then
'            rptGrid.Relatorio.Orientacao = oriPaisagem
'        End If
        rptGrid.Relatorio.Escala = scmCentimetros
        rptGrid.Relatorio.MargemEsquerda = 1
        rptGrid.Relatorio.MargemDireita = 1
        rptGrid.Titulo = "Vendedor: " & DbcVend.Text
        ' Configura a distancia em twips que o texto
        ' fica da linha do grid
        ' rptGrid.GapColunas = 30
        rptGrid.Executar
    End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbImprimir_Click()
    
    If MsgBox("Confirma impressão?", 36) = vbYes Then
    
        Set Fs2 = CreateObject("Scripting.FileSystemObject")
        
        Fs2.CopyFile App.Path & "\caixa.txt", Porta
        
'        Fs2.Close
        
    End If

End Sub

Private Sub BtVisualiza_Click()
Dim VarData As Date
If DbcVend.Text <> "" Then

  With flxGrid
    .Redraw = False
'    .Font = Me.Font
    .Clear
    .Rows = 13
    .FixedRows = 1
    .FixedCols = 1
    .FormatString = "^Mês|>" & Year(Date) & "|>" & Year(Date) - 1 & "|>" & Year(Date) - 2 & "|>" & Year(Date) - 3
    .ColWidth(0) = TextWidth("X") * 10
    .ColWidth(1) = TextWidth("X") * 15
    .ColWidth(2) = TextWidth("X") * 15
    .ColWidth(3) = TextWidth("X") * 15
    .ColWidth(4) = TextWidth("X") * 15
    
    .TextMatrix(1, 0) = "jan"
    .TextMatrix(2, 0) = "Fev"
    .TextMatrix(3, 0) = "Mar"
    .TextMatrix(4, 0) = "Abr"
    .TextMatrix(5, 0) = "Mai"
    .TextMatrix(6, 0) = "Jun"
    .TextMatrix(7, 0) = "Jul"
    .TextMatrix(8, 0) = "ago"
    .TextMatrix(9, 0) = "Set"
    .TextMatrix(10, 0) = "Out"
    .TextMatrix(11, 0) = "Nov"
    .TextMatrix(12, 0) = "Dez"
    
    VarData = Date - 1095
    
    Set Venda = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, Year(ped.dataemiss) AS Expr1, ENT.CodEntidade, Month(ped.dataemiss) AS Expr2 FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)>=#" & Format(VarData, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(Date, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY ENT.CodEntidade, ENT.Nome, Year(ped.dataemiss), ENT.CodEntidade, Month(ped.dataemiss) HAVING (((ENT.CodEntidade)= " & DbcVend.BoundText & " )) ORDER BY ENT.Nome, Month(ped.dataemiss), Year(ped.dataemiss)")

    Venda.Requery
    If Venda.RecordCount > 0 Then
        Venda.MoveFirst
        While Not Venda.Eof
        
           .TextMatrix(Venda!Expr2, (Year(Date) - Venda!Expr1) + 1) = Alinhar(Venda("SomaDeVlrVdaDia"), 12, 2)
            Venda.MoveNext
        Wend
    End If
    
    .WordWrap = False
    
    .Redraw = True
    
    
    End With
Else
    MsgBox "Vendedor não seleciondo ", vbInformation, App.Title
End If


End Sub

Private Sub Command1_Click()

    FrmRelTXT.Config App.Path & "\EstatVen.Txt"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim Componentes As Variant

'MebData.Text = Date
'MebData2.Text = Date


Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 200 '(MDIPrincipal.Height - Me.Height) / 6

DatVend.DatabaseName = Caminho & "\Dados.Mdb"

End Sub

Private Sub MebDif_LostFocus()
BtCalcula_Click
End Sub


Private Sub MebProximo_LostFocus()
BtCalcula_Click
End Sub


Private Sub MebValor_LostFocus()
BtCalcula_Click
End Sub
