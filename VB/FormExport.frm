VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormExport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7455
   Begin VB.CommandButton ExportaExcel 
      Caption         =   "Exporta Excell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2430
      TabIndex        =   10
      Top             =   5175
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -225
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   9
      Top             =   6210
      Width           =   11535
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   330
      Left            =   585
      TabIndex        =   2
      Top             =   2295
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7485
      TabIndex        =   7
      Top             =   0
      Width           =   7485
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportação de Dados"
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
         TabIndex        =   8
         Top             =   45
         Width           =   5010
      End
   End
   Begin VB.Data FamDestino 
      Caption         =   "FamDestino"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5580
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data GruDestino 
      Caption         =   "GruDestino"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data LojDestino 
      Caption         =   "LojDestino"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4545
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data BarDestino 
      Caption         =   "BarDestino"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data PrdDestino 
      Caption         =   "PrdDestino"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3465
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data FamOrigem 
      Caption         =   "FamOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5535
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data GruOrigem 
      Caption         =   "GruOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4995
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data LojOrigem 
      Caption         =   "LojOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4500
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data BarOrigem 
      Caption         =   "BarOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3915
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data PrdOrigem 
      Caption         =   "PrdOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3420
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton CmdImp 
      Caption         =   "Importa Produtos"
      Height          =   465
      Left            =   3825
      TabIndex        =   6
      Top             =   3960
      Width           =   1590
   End
   Begin VB.CommandButton CmdExp 
      Caption         =   "Exporta Produtos"
      Height          =   465
      Left            =   1980
      TabIndex        =   5
      Top             =   3960
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DtData 
      Height          =   330
      Left            =   630
      TabIndex        =   3
      Top             =   1755
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Format          =   83689473
      CurrentDate     =   37732
   End
   Begin VB.Data DatItpDest 
      Caption         =   "ItpDest"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistemas em Desenvolvimento\Desenvolvimento\Reta\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4770
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ITP"
      Top             =   2745
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data DatPedDest 
      Caption         =   "PedDest"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistemas em Desenvolvimento\Desenvolvimento\Reta\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4770
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PED"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data DatPedOrigem 
      Caption         =   "PedOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2745
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Data DatItpOrigem 
      Caption         =   "ItpOrigem"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2745
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton CmbOk 
      Caption         =   "&Inicia"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1755
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7470
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Limite"
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
      Left            =   630
      TabIndex        =   4
      Top             =   1485
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção, está operação pode demorar, procure executa-la  quando outras estações  não estiverem ligados."
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
      Height          =   735
      Left            =   630
      TabIndex        =   0
      Top             =   675
      Width           =   6090
   End
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LocalDrive As String


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmbOk_Click()
Dim Ped As Double
Dim Itp As Double
Dim PedAnt As Double
Dim Criterio As String

'On Error GoTo TrataErro

CmbOk.Caption = "Aguarde ..."
CmbOk.Enabled = False

DatPedOrigem.RecordSource = "SELECT Ped.* FROM Ped WHERE DataEmiss < #" & Format(DtData, "dd/mm/yyyy") & "# Order By CodPed "
DatItpOrigem.RecordSource = "SELECT Itp.* FROM Itp ORDER BY Pedido"

DatPedOrigem.Refresh

If DatPedOrigem.Recordset.RecordCount <> 0 Then
    DatItpOrigem.Refresh
    
    DatPedOrigem.Recordset.MoveLast
    DoEvents
    Bar1.Min = 0
    Bar1.Max = DatPedOrigem.Recordset.RecordCount
    DatPedOrigem.Recordset.MoveFirst
    While Not DatPedOrigem.Recordset.Eof
        DatPedDest.Recordset.AddNew
        PedAnt = DatPedOrigem.Recordset.Fields(0)
        For Ped = 0 To DatPedOrigem.Recordset.Fields.Count - 1
            DatPedDest.Recordset.Fields(Ped) = DatPedOrigem.Recordset.Fields(Ped)
        Next
        DatPedDest.Recordset.Update
        Criterio = "Pedido = " & PedAnt
        DatItpOrigem.Recordset.FindFirst Criterio
        If Not DatItpOrigem.Recordset.NoMatch Then
            While DatItpOrigem.Recordset.Fields("Pedido") = PedAnt And Not DatItpOrigem.Recordset.Eof
                DatItpDest.Recordset.AddNew
                For Itp = 0 To DatItpOrigem.Recordset.Fields.Count - 1
                    DatItpDest.Recordset.Fields(Itp) = DatItpOrigem.Recordset.Fields(Itp)
                Next
                DatItpDest.Recordset.Update
                DatItpOrigem.Recordset.MoveNext
            Wend
        End If
        Bar1.value = DatPedOrigem.Recordset.AbsolutePosition
        DatPedOrigem.Recordset.MoveNext
    Wend
    
    MsgBox "Operação concluída com sucesso !", vbInformation, App.Title
    
    Bar1.value = 0
        
Else
    MsgBox "Arquivo sem movimento !", vbInformation, App.Title

End If
CmbOk.Caption = "Inicia"
CmbOk.Enabled = True

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdExpPrd_Click()
Dim X As Integer
Dim VarLinha As String
Dim VarPrd As Recordset
VarLinha = ""

    Sql = "Select * FROM Prd Order By CodPrd "
    Set VarPrd = Banco.OpenRecordset(Sql)
    VarPrd.Requery
    If VarPrd.RecordCount > 0 Then
       VarPrd.MoveFirst
       Open App.Path & "\Produtos.txt" For Output As #1
       While Not VarPrd.Eof
           DoEvents
           VarLinha = VarPrd.Fields(0) & "|"
           For X = 1 To 62
               If VarPrd.Fields(X).Name = "Obs" Then
                   If IsNull(VarPrd.Fields(X)) Then
                       VarLinha = VarLinha & "" & "|"
                   Else
                       VarLinha = VarLinha & LimpaString(VarPrd.Fields(X)) & "|"
                   End If
               Else
                   VarLinha = VarLinha & VarPrd.Fields(X) & "|"
               End If
           Next
           Print #1, VarLinha
           VarPrd.MoveNext
       Wend
       MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
       Close #1
    End If

End Sub

Private Sub CmdImpPrd_Click()

Dim X As Integer
Dim y As Integer
Dim VarLinha As String
Dim VarPrd As Recordset
Dim VarCampo(1 To 63) As String
Dim VarItem As String
VarLinha = ""


    Sql = "Select * FROM Prd1 Order By CodPrd "
    Set VarPrd = Banco.OpenRecordset(Sql)
    VarPrd.Requery
       If VarPrd.RecordCount > 0 Then
            VarPrd.MoveFirst
       End If
       Open App.Path & "\Produtos.txt" For Input As #1
       Do
            y = 1
            For X = 1 To 63
                VarCampo(X) = ""
            Next
            Input #1, VarLinha
            For X = 1 To Len(VarLinha)
                 VarItem = Mid$(VarLinha, X, 1)
                 If VarItem <> "|" Then
                     VarCampo(y) = VarCampo(y) & VarItem
                 Else
                     y = y + 1
                 End If
            Next
            
            Criterio = "CodPrd = " & VarCampo(2)
            VarPrd.FindFirst Criterio
            If Not VarPrd.NoMatch Then
                 VarPrd.Edit
                 For X = 3 To 63
                     If VarCampo(X) <> "" Then
                         VarPrd.Fields(X - 1) = VarCampo(X)
                     End If
                 Next
                 VarPrd.Update
            Else
                 VarPrd.AddNew
                 For X = 2 To 63
                     If VarCampo(X) <> "" Then
                         VarPrd.Fields(X - 1) = VarCampo(X)
                     End If
                 Next
                 VarPrd.Update
            End If
                   
       Loop Until Eof(1)
       MsgBox "Arquivo importado com sucesso!", vbInformation, App.Title
       Close #1


End Sub

Private Sub CmdExp_Click()

On Error Resume Next

PrdOrigem.DatabaseName = Caminho & "\Dados.Mdb"
BarOrigem.DatabaseName = Caminho & "\Dados.Mdb"
LojOrigem.DatabaseName = Caminho & "\Dados.Mdb"
GruOrigem.DatabaseName = Caminho & "\Dados.Mdb"
FamOrigem.DatabaseName = Caminho & "\Dados.Mdb"


PrdDestino.DatabaseName = LocalDrive & "\Export.Mdb"
BarDestino.DatabaseName = LocalDrive & "\Export.Mdb"
LojDestino.DatabaseName = LocalDrive & "\Export.Mdb"
GruDestino.DatabaseName = LocalDrive & "\Export.Mdb"
FamDestino.DatabaseName = LocalDrive & "\Export.Mdb"


CmdExp.Caption = "Aguarde ..."
CmdExp.Enabled = False
CmdImp.Enabled = False

PrdOrigem.RecordSource = "SELECT * FROM Prd ORDER BY CodPrd"
BarOrigem.RecordSource = "SELECT * FROM Bar ORDER BY BarCodBar"
LojOrigem.RecordSource = "SELECT * FROM Loj ORDER BY Produto"
GruOrigem.RecordSource = "SELECT * FROM Gru ORDER BY CodGru"
FamOrigem.RecordSource = "SELECT * FROM Fam ORDER BY CodFam"

PrdDestino.RecordSource = "SELECT * FROM Prd ORDER BY CodPrd"
BarDestino.RecordSource = "SELECT * FROM Bar ORDER BY BarCodBar"
LojDestino.RecordSource = "SELECT * FROM Loj ORDER BY Produto"
GruDestino.RecordSource = "SELECT * FROM Gru ORDER BY CodGru"
FamDestino.RecordSource = "SELECT * FROM Fam ORDER BY CodFam"


FamOrigem.Refresh
GruOrigem.Refresh
LojOrigem.Refresh
BarOrigem.Refresh
PrdOrigem.Refresh

FamDestino.Refresh
GruDestino.Refresh
LojDestino.Refresh
BarDestino.Refresh
PrdDestino.Refresh

' ******** Familia

If FamOrigem.Recordset.RecordCount <> 0 Then
    
    DoEvents
    While Not FamOrigem.Recordset.Eof
        
        Criterio = "CodFam = " & FamOrigem.Recordset("CodFam")
        FamDestino.Recordset.FindFirst Criterio
        If Not FamDestino.Recordset.NoMatch Then
            FamDestino.Recordset.Edit
            For X = 1 To FamOrigem.Recordset.Fields.Count - 1
                FamDestino.Recordset.Fields(X) = FamOrigem.Recordset.Fields(X)
            Next
            FamDestino.Recordset.Update
        Else
            FamDestino.Recordset.AddNew
            For X = 0 To FamOrigem.Recordset.Fields.Count - 1
                FamDestino.Recordset.Fields(X) = FamOrigem.Recordset.Fields(X)
            Next
            FamDestino.Recordset.Update
        End If
        FamOrigem.Recordset.MoveNext
    
    Wend
    
End If

' ******** Grupo

If GruOrigem.Recordset.RecordCount <> 0 Then
    
    While Not GruOrigem.Recordset.Eof
        Criterio = "CodGru = " & GruOrigem.Recordset("CodGru")
        GruDestino.Recordset.FindFirst Criterio
        If Not GruDestino.Recordset.NoMatch Then
            GruDestino.Recordset.Edit
            For X = 1 To GruOrigem.Recordset.Fields.Count - 1
                GruDestino.Recordset.Fields(X) = GruOrigem.Recordset.Fields(X)
            Next
            GruDestino.Recordset.Update
        Else
            GruDestino.Recordset.AddNew
            For X = 0 To GruOrigem.Recordset.Fields.Count - 1
                GruDestino.Recordset.Fields(X) = GruOrigem.Recordset.Fields(X)
            Next
            GruDestino.Recordset.Update
        End If
        GruOrigem.Recordset.MoveNext
    
    Wend
    
End If

' ******** Bar

If BarOrigem.Recordset.RecordCount <> 0 Then
    
    While Not BarOrigem.Recordset.Eof
        Criterio = "BarBarra = " & """" & BarOrigem.Recordset("BarBarra") & """" & " and  Produto = " & BarOrigem.Recordset.Fields("Produto")
        BarDestino.Recordset.FindFirst Criterio
        If Not BarDestino.Recordset.NoMatch Then
            BarDestino.Recordset.Edit
            For X = 1 To BarOrigem.Recordset.Fields.Count - 1
                BarDestino.Recordset.Fields(X) = BarOrigem.Recordset.Fields(X)
            Next
            BarDestino.Recordset.Update
        Else
            BarDestino.Recordset.AddNew
            For X = 0 To BarOrigem.Recordset.Fields.Count - 1
                BarDestino.Recordset.Fields(X) = BarOrigem.Recordset.Fields(X)
            Next
            BarDestino.Recordset.Update
        End If
        BarOrigem.Recordset.MoveNext
    
    Wend
    
End If


' ******** Loja

If LojOrigem.Recordset.RecordCount <> 0 Then
    
    While Not LojOrigem.Recordset.Eof
        Criterio = "DescricaoLoja = " & """" & LojOrigem.Recordset("DescricaoLoja") & """" & " and Produto = " & LojOrigem.Recordset("Produto")
        LojDestino.Recordset.FindFirst Criterio
        If Not LojDestino.Recordset.NoMatch Then
            LojDestino.Recordset.Edit
            For X = 1 To LojOrigem.Recordset.Fields.Count - 1
                LojDestino.Recordset.Fields(X) = LojOrigem.Recordset.Fields(X)
            Next
            LojDestino.Recordset.Update
        Else
            LojDestino.Recordset.AddNew
            For X = 0 To LojOrigem.Recordset.Fields.Count - 1
                LojDestino.Recordset.Fields(X) = LojOrigem.Recordset.Fields(X)
            Next
            LojDestino.Recordset.Update
        End If
        LojOrigem.Recordset.MoveNext
    
    Wend
    
End If


' ******** Produto

If PrdOrigem.Recordset.RecordCount <> 0 Then
    
    While Not PrdOrigem.Recordset.Eof
        Criterio = "CodPrd = " & PrdOrigem.Recordset("Codprd")
        PrdDestino.Recordset.FindFirst Criterio
        If Not PrdDestino.Recordset.NoMatch Then
            PrdDestino.Recordset.Edit
            For X = 1 To PrdOrigem.Recordset.Fields.Count - 1
                PrdDestino.Recordset.Fields(X) = PrdOrigem.Recordset.Fields(X)
            Next
            PrdDestino.Recordset.Update
        Else
            PrdDestino.Recordset.AddNew
            For X = 0 To PrdOrigem.Recordset.Fields.Count - 1
                If PrdOrigem.Recordset.Fields(X).Name = "Fabricante" Then
                    PrdDestino.Recordset.Fields(X) = 0
                ElseIf PrdOrigem.Recordset.Fields(X).Name = "SitTrib" Then
                        If PrdOrigem.Recordset.Fields(X) = "000" Then
                            PrdDestino.Recordset.Fields(X) = "102"
                        ElseIf PrdOrigem.Recordset.Fields(X) = "060" Then
                            PrdDestino.Recordset.Fields(X) = "500"
                        End If
                    Else
                        PrdDestino.Recordset.Fields(X) = PrdOrigem.Recordset.Fields(X)
                End If
            Next
            PrdDestino.Recordset.Update
        End If
        PrdOrigem.Recordset.MoveNext
    
    Wend
    
End If


CmdExp.Caption = "Exporta Produtos"
CmdImp.Enabled = True
CmdExp.Enabled = True

MsgBox "Arquivo exportado com sucesso!", vbInformation, App.Title

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub CmdImp_Click()

On Error Resume Next


PrdOrigem.DatabaseName = LocalDrive & "\Export.Mdb"
BarOrigem.DatabaseName = LocalDrive & "\Export.Mdb"
LojOrigem.DatabaseName = LocalDrive & "\Export.Mdb"
GruOrigem.DatabaseName = LocalDrive & "\Export.Mdb"
FamOrigem.DatabaseName = LocalDrive & "\Export.Mdb"


PrdDestino.DatabaseName = Caminho & "\Dados.Mdb"
BarDestino.DatabaseName = Caminho & "\Dados.Mdb"
LojDestino.DatabaseName = Caminho & "\Dados.Mdb"
GruDestino.DatabaseName = Caminho & "\Dados.Mdb"
FamDestino.DatabaseName = Caminho & "\Dados.Mdb"

CmdImp.Caption = "Aguarde ..."
CmdImp.Enabled = False
CmdExp.Enabled = False

PrdOrigem.RecordSource = "SELECT * FROM Prd ORDER BY CodPrd"
BarOrigem.RecordSource = "SELECT * FROM Bar ORDER BY BarCodBar"
LojOrigem.RecordSource = "SELECT * FROM Loj ORDER BY Produto"
GruOrigem.RecordSource = "SELECT * FROM Gru ORDER BY CodGru"
FamOrigem.RecordSource = "SELECT * FROM Fam ORDER BY CodFam"

PrdDestino.RecordSource = "SELECT * FROM Prd ORDER BY CodPrd"
BarDestino.RecordSource = "SELECT * FROM Bar ORDER BY BarCodBar"
LojDestino.RecordSource = "SELECT * FROM Loj ORDER BY Produto"
GruDestino.RecordSource = "SELECT * FROM Gru ORDER BY CodGru"
FamDestino.RecordSource = "SELECT * FROM Fam ORDER BY CodFam"


FamOrigem.Refresh
GruOrigem.Refresh
LojOrigem.Refresh
BarOrigem.Refresh
PrdOrigem.Refresh

FamDestino.Refresh
GruDestino.Refresh
LojDestino.Refresh
BarDestino.Refresh
PrdDestino.Refresh

' ******** Familia

If FamOrigem.Recordset.RecordCount <> 0 Then
    DoEvents
    While Not FamOrigem.Recordset.Eof
        Criterio = "CodFam = " & FamOrigem.Recordset("CodFam")
        FamDestino.Recordset.FindFirst Criterio
        If Not FamDestino.Recordset.NoMatch Then
            FamDestino.Recordset.Edit
            For X = 1 To FamOrigem.Recordset.Fields.Count - 1
                FamDestino.Recordset.Fields(X) = FamOrigem.Recordset.Fields(X)
            Next
            FamDestino.Recordset.Update
        Else
            FamDestino.Recordset.AddNew
            For X = 0 To FamOrigem.Recordset.Fields.Count - 1
                FamDestino.Recordset.Fields(X) = FamOrigem.Recordset.Fields(X)
            Next
            FamDestino.Recordset.Update
        End If
        FamOrigem.Recordset.MoveNext
    
    Wend
    
End If

' ******** Grupo

If GruOrigem.Recordset.RecordCount <> 0 Then
    
    While Not GruOrigem.Recordset.Eof
        Criterio = "CodGru = " & GruOrigem.Recordset("CodGru")
        GruDestino.Recordset.FindFirst Criterio
        If Not GruDestino.Recordset.NoMatch Then
            GruDestino.Recordset.Edit
            For X = 1 To GruOrigem.Recordset.Fields.Count - 1
                GruDestino.Recordset.Fields(X) = GruOrigem.Recordset.Fields(X)
            Next
            GruDestino.Recordset.Update
        Else
            GruDestino.Recordset.AddNew
            For X = 0 To GruOrigem.Recordset.Fields.Count - 1
                GruDestino.Recordset.Fields(X) = GruOrigem.Recordset.Fields(X)
            Next
            GruDestino.Recordset.Update
        End If
        GruOrigem.Recordset.MoveNext
    
    Wend
    
End If

' ******** Bar

If BarOrigem.Recordset.RecordCount <> 0 Then
    
    While Not BarOrigem.Recordset.Eof
        Criterio = "BarBarra = " & """" & BarOrigem.Recordset("BarBarra") & """" & " and  Produto = " & BarOrigem.Recordset.Fields("Produto")
        BarDestino.Recordset.FindFirst Criterio
        If Not BarDestino.Recordset.NoMatch Then
            BarDestino.Recordset.Edit
            For X = 1 To BarOrigem.Recordset.Fields.Count - 1
                BarDestino.Recordset.Fields(X) = BarOrigem.Recordset.Fields(X)
            Next
            BarDestino.Recordset.Update
        Else
            BarDestino.Recordset.AddNew
            For X = 0 To BarOrigem.Recordset.Fields.Count - 1
                BarDestino.Recordset.Fields(X) = BarOrigem.Recordset.Fields(X)
            Next
            BarDestino.Recordset.Update
        End If
        BarOrigem.Recordset.MoveNext
    
    Wend
    
End If


' ******** Loja

If LojOrigem.Recordset.RecordCount <> 0 Then
    
    While Not LojOrigem.Recordset.Eof
        Criterio = "DescricaoLoja = " & """" & LojOrigem.Recordset("DescricaoLoja") & """" & " and Produto = " & LojOrigem.Recordset("Produto")
        LojDestino.Recordset.FindFirst Criterio
        If Not LojDestino.Recordset.NoMatch Then
'            LojDestino.Recordset.Edit
'            For x = 1 To LojOrigem.Recordset.Fields.Count - 1
'                LojDestino.Recordset.Fields(x) = LojOrigem.Recordset.Fields(x)
'            Next
'            LojDestino.Recordset.Update
        Else
            LojDestino.Recordset.AddNew
            For X = 0 To LojOrigem.Recordset.Fields.Count - 1
                If LojOrigem.Recordset.Fields(X).Name <> "Estoque" Then
                    LojDestino.Recordset.Fields(X) = LojOrigem.Recordset.Fields(X)
                Else
                    LojDestino.Recordset.Fields(X) = 0
                End If
            Next
            LojDestino.Recordset.Update
        End If
        LojOrigem.Recordset.MoveNext
    
    Wend
    
End If


' ******** Produto

If PrdOrigem.Recordset.RecordCount <> 0 Then
    
    While Not PrdOrigem.Recordset.Eof
        Criterio = "CodPrd = " & PrdOrigem.Recordset("Codprd")
        PrdDestino.Recordset.FindFirst Criterio
        If Not PrdDestino.Recordset.NoMatch Then
            PrdDestino.Recordset.Edit
            For X = 2 To PrdOrigem.Recordset.Fields.Count - 1
                If PrdOrigem.Recordset.Fields(X).Name <> "Estoque" Then
                    If PrdOrigem.Recordset.Fields(X).Name <> "DtCompra" Then
                        If PrdOrigem.Recordset.Fields(X).Name <> "DtVenda" Then
                            If PrdOrigem.Recordset.Fields(X).Name <> "Fabricante" Then
                                PrdDestino.Recordset.Fields(X) = PrdOrigem.Recordset.Fields(X)
                            End If
                        End If
                    End If
                End If
            Next
            PrdDestino.Recordset.Update
        Else
            PrdDestino.Recordset.AddNew
            For X = 0 To PrdOrigem.Recordset.Fields.Count - 1
                PrdDestino.Recordset.Fields(X) = PrdOrigem.Recordset.Fields(X)
            Next
            PrdDestino.Recordset.Update
        End If
        PrdOrigem.Recordset.MoveNext
    
    Wend
    
End If

CmdImp.Caption = "Importa Produtos"
CmdImp.Enabled = True
CmdExp.Enabled = True

MsgBox "Arquivo importado com sucesso!", vbInformation, App.Title

Exit Sub

TrataErro:
    
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub ExportaExcel_Click()

Dim stCell As ExlCell
Dim CamposSql As String
Dim Ordenado As String
Dim Sql As String
Dim SqlWhere As String
    
              
              MousePointer = vbHourglass ' Muda o ponteiro do mouse
              
              Set oExcel = CreateObject("Excel.Application")
              oExcel.Workbooks.Add   'inclui o workbook
              Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
              
              Set db = OpenDatabase(App.Path & "\Dados.mdb")
              Set Sn = db.OpenRecordset("SELECT PED.CodPed, ENT.CodEntidade, ENT.Nome, PRD.CodPrd, PRD.Descrição_Produto, ITP.Referencia, ITP.Qtd FROM ENT INNER JOIN (PRD INNER JOIN (PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) ON PRD.CodPrd = ITP.Produto) ON ENT.CodEntidade = PED.Entidade WHERE (((cdate(ITP.Referencia))<=" & DtData & " )) ORDER BY ITP.Referencia DESC ", dbOpenSnapshot)
              
'              CamposSql = " SELECT CodPrd, Descrição_Produto, Troca, Falta From Prd "
'              Ordenado = " ORDER BY Descrição_Produto "
'              Sql = CamposSql & SqlWhere & Ordenado
              
                  
              ' Inclui os dados a partir da celula A1
              stCell.Row = 1
              stCell.Col = 1
              CopiarTabelaExcel Sn, objExlSht, stCell
              objExlSht.Application.Columns(1).ColumnWidth = 12
              objExlSht.Application.Columns(2).ColumnWidth = 40
              objExlSht.Application.Columns(3).ColumnWidth = 15
              objExlSht.Application.Columns(4).ColumnWidth = 12
              
              objExlSht.Application.Cells(1, 4).value = "Preço"
              objExlSht.Application.Cells(1, 5).value = " <= Condições"
              
              
              
              ' Salva a planilha
              objExlSht.SaveAs App.Path & "\PrdVend.xls"
              
              oExcel.Visible = True
'              oExcel.
            '  frmexcelvb.Show
              MsgBox "Arquivo gerado com sucesso !", vbInformation, App.Title

End Sub

Private Sub Form_Activate()
Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6
End Sub

Private Sub Form_Load()
'Dim LocalDrive As String

LocalDrive = LerINI("Produto", "SkyDrive", Caminho & "\config.ini")


DatPedOrigem.DatabaseName = Caminho & "\Dados.Mdb"
DatItpOrigem.DatabaseName = Caminho & "\Dados.Mdb"

DatPedDest.DatabaseName = Caminho & "\Morto.Mdb"
DatItpDest.DatabaseName = Caminho & "\Morto.Mdb"



DtData.value = Date

End Sub
