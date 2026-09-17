VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormPegasus 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5280
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -360
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   12
      Top             =   3150
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
      TabIndex        =   10
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Integração Pegasus Bayer"
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
         Left            =   90
         TabIndex        =   11
         Top             =   45
         Width           =   4335
      End
   End
   Begin VB.CommandButton BtCancela 
      Caption         =   "&Cancela"
      Height          =   405
      Left            =   2745
      TabIndex        =   5
      Top             =   2475
      Width           =   1665
   End
   Begin VB.CommandButton Btimprimir 
      Caption         =   "&Visualiza"
      Height          =   405
      Left            =   720
      TabIndex        =   4
      Top             =   2475
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estoque"
      Height          =   1350
      Left            =   2565
      TabIndex        =   0
      Top             =   810
      Width           =   1995
      Begin VB.OptionButton OptTipo 
         Caption         =   "Posição"
         Height          =   345
         Index           =   2
         Left            =   270
         TabIndex        =   3
         Top             =   945
         Width           =   1065
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "&Saida"
         Height          =   345
         Index           =   1
         Left            =   270
         TabIndex        =   2
         Top             =   630
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "&Entrada"
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   285
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker DtData1 
      Height          =   330
      Left            =   720
      TabIndex        =   6
      Top             =   990
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin MSComCtl2.DTPicker DtData2 
      Height          =   330
      Left            =   720
      TabIndex        =   7
      Top             =   1665
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   52953089
      CurrentDate     =   36641
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   765
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   720
   End
End
Attribute VB_Name = "FormPegasus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim oExcel As Object
Dim objExlSht As Object
Dim db As Database
Dim Sn As Recordset   ' Recordset do tipo Snapshot

Private Type ExlCell
   row As Long
   col As Long
End Type

Private Sub BtCancela_Click()
Unload Me
End Sub

Private Sub BtImprimir_Click()
Dim stCell As ExlCell

If OptTipo(0).value = True Then

    
    MousePointer = vbHourglass ' Muda o ponteiro do mouse
    
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Add   'inclui o workbook
    Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
    
    Set db = OpenDatabase(Caminho & "\dados.mdb")
    Set Sn = db.OpenRecordset("SELECT ENT.Nome, ENT.Cidade, PED.Cfo, PED.NroPedido, PRD.Descrição_Produto, ITP.Qtd,  PRD.Custo, PED.DtSaida FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PRD.Grupo)=1) AND ((PED.Operacao)=0 Or (PED.Operacao)=9 Or (PED.Operacao)=10) AND DTSAIDA >= #" & Format(DtData1, "MM/DD/YYYY") & "# AND DTSAIDA <= #" & Format(DtData2, "MM/DD/YYYY") & "# )", dbOpenSnapshot)
    
    ' Inclui os dados a partir da celula A1
    stCell.row = 1
    stCell.col = 1
    CopiarTabelaExcel Sn, objExlSht, stCell
    
    ' Salva a planilha
    objExlSht.SaveAs Caminho & "\Entrada.xls"
    
    oExcel.Visible = True
  '  frmexcelvb.Show


ElseIf OptTipo(1).value = True Then
        MousePointer = vbHourglass ' Muda o ponteiro do mouse
        
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Add   'inclui o workbook
        Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
        
        Set db = OpenDatabase(Caminho & "\dados.mdb")
        Set Sn = db.OpenRecordset("SELECT ENT.Nome, ENT.Cidade, PED.Cfo, PED.NroNt, PRD.Descrição_Produto, ITP.Qtd, PRD.Custo, PED.DataEmiss, PRD.Grupo FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido WHERE (((PED.NroNt)<>0) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=8 Or (PED.Operacao)=11) AND ((PRD.Grupo)=1)AND DTSAIDA >= #" & Format(DtData1, "MM/DD/YYYY") & "# AND DTSAIDA <= #" & Format(DtData2, "MM/DD/YYYY") & "# )", dbOpenSnapshot)
        
        ' Inclui os dados a partir da celula A1
        stCell.row = 1
        stCell.col = 1
        CopiarTabelaExcel Sn, objExlSht, stCell
        
        ' Salva a planilha
        objExlSht.SaveAs Caminho & "\Saida.xls"
        
        oExcel.Visible = True
      '  frmexcelvb.Show
    

    Else

      MousePointer = vbHourglass ' Muda o ponteiro do mouse
      
      Set oExcel = CreateObject("Excel.Application")
      oExcel.Workbooks.Add   'inclui o workbook
      Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
      
      Set db = OpenDatabase(Caminho & "\dados.mdb")
      Set Sn = db.OpenRecordset("SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Estoque FROM PRD WHERE Grupo = 1 order by Codprd ")
      ' Inclui os dados a partir da celula A1
      stCell.row = 1
      stCell.col = 1
      CopiarTabelaExcel Sn, objExlSht, stCell
      
      ' Salva a planilha
      objExlSht.SaveAs Caminho & "\Estoque.xls"
      
      oExcel.Visible = True
    '  frmexcelvb.Show

End If



End Sub


Private Sub CopiarTabelaExcel(Rs As Recordset, ws As Worksheet, StartingCell As ExlCell)

Dim Vetor() As Variant
Dim row As Long, col As Long
Dim fd As Field

Rs.MoveLast
ReDim Vetor(Rs.RecordCount + 1, Rs.Fields.Count)

' Copia as colunas do cabecalho para um vetor
col = 0
For Each fd In Rs.Fields
  Vetor(0, col) = fd.Name
  col = col + 1
Next

' copia o rs par um vetor
Rs.MoveFirst
For row = 1 To Rs.RecordCount
   For col = 0 To Rs.Fields.Count - 1
       Vetor(row, col) = Rs.Fields(col).value
       'O Excel não suporta valores NULL em uma célula.
       If IsNull(Vetor(row, col)) Then Vetor(row, col) = ""
   Next
   Rs.MoveNext
Next
ws.Range(ws.Cells(StartingCell.row, StartingCell.col), ws.Cells(StartingCell.row + Rs.RecordCount + 1, _
StartingCell.col + Rs.Fields.Count)).value = Vetor

End Sub


Private Sub Form_Activate()
DtData1 = Date - 90
DtData2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Label1.Caption = "Encerrando o Excel"
Label1.Refresh
'objExlSht.Application.Quit

Set objExlSht = Nothing   ' remove a variavel objeto
Set oExcel = Nothing       ' remove a variavel objeto
Set Sn = Nothing             ' reomove a variavel objeto
Set db = Nothing             ' reomove a variavel objeto

MousePointer = vbDefault     ' Restaura o ponteiro do mouse.
Label1.Caption = "Muito bem, deu certo ! "
Label1.Refresh

End Sub
