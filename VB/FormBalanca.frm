VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormBalanca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exporta Preços para o PGR (Balança Filizola)"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9390
   Begin VB.CommandButton CmdCancela 
      Caption         =   "&Cancela"
      Height          =   375
      Left            =   5085
      TabIndex        =   2
      Top             =   4815
      Width           =   1770
   End
   Begin VB.CommandButton CmdAtualiza 
      Caption         =   "&Atualiza"
      Height          =   375
      Left            =   2565
      TabIndex        =   1
      Top             =   4815
      Width           =   1770
   End
   Begin VB.Data DatDados 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   "C:\filizola\pgr"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Produtos"
      Top             =   3285
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormBalanca.frx":0000
      Height          =   4335
      Left            =   180
      OleObjectBlob   =   "FormBalanca.frx":001A
      TabIndex        =   0
      Top             =   270
      Width           =   8970
   End
End
Attribute VB_Name = "FormBalanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registros As Recordset

Private Sub CmdAtualiza_Click()

Dim Criterio As String
If DatDados(0).Recordset.RecordCount <> 0 Then
    DatDados(0).Recordset.MoveFirst
    While Not DatDados(0).Recordset.EOF
        Criterio = "CodPrd = " & DatDados(0).Recordset.Fields("Cod")
        Registros.FindFirst Criterio
        If Not Registros.NoMatch Then
            If DatDados(0).Recordset.Fields("PU") = CInt(Registros.Fields("Preco") * 100) Then
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("ALT_P") = 0
                DatDados(0).Recordset.Update
            Else
                DatDados(0).Recordset.Edit
                DatDados(0).Recordset.Fields("PU") = CInt(Registros.Fields("Preco") * 100)
                DatDados(0).Recordset.Fields("ALT_P") = 1
                DatDados(0).Recordset.Update
            End If
        End If
        DoEvents
        DatDados(0).Recordset.MoveNext
    Wend
    DatDados(0).Refresh
    DBGrid1.Refresh
    MsgBox "Preços exportados ! Não esqueça de dar carga nas balanças", vbInformation, App.Title
End If

End Sub

Private Sub CmdCancela_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Componente As Variant

FormBalanca.Left = 100
FormBalanca.Top = 200

Set Registros = Banco.OpenRecordset("SELECT PRD.CodPrd, PRD.Preco FROM PRD ORDER BY PRD.CodPrd")

End Sub
