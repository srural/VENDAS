VERSION 5.00
Begin VB.Form FormEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Entrega"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -45
      ScaleHeight     =   330
      ScaleWidth      =   11325
      TabIndex        =   5
      Top             =   5940
      Width           =   11355
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8610
      TabIndex        =   3
      Top             =   0
      Width           =   8610
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Entrega"
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
         TabIndex        =   4
         Top             =   45
         Width           =   5325
      End
   End
   Begin VB.CommandButton BtOk 
      Appearance      =   0  'Flat
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   420
      Left            =   135
      TabIndex        =   2
      Top             =   5310
      Width           =   1995
   End
   Begin VB.CommandButton CmdMarcar 
      Appearance      =   0  'Flat
      Caption         =   "&Desmarcar"
      Height          =   420
      Left            =   6480
      TabIndex        =   1
      Top             =   5265
      Width           =   1905
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "FormEntrega.frx":0000
      Left            =   135
      List            =   "FormEntrega.frx":0007
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   675
      Width           =   8295
   End
End
Attribute VB_Name = "FormEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtOk_Click()
Dim y As Double
Dim Criterio As String
        
If VarImpressoraNf = "CUPOM" Or VarImpressoraNf = "MECAF" Then
        
        Open VarImprEntrega For Output As #1

        Print #1, "" & Cond
        Print #1, ""
        Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 23) & Nor
        Print #1, Cond & Centralizar(FrmFrente.Empresa("Cabecalho2"), 48)
        Print #1, "================================================"
        Print #1, "Data:" & Date & " Hora:" & Time; Tab(33)
        Print #1, Tab(1); psCompString(1, FrmFrente.Clientes("CodEntidade"), 6); "-"; Left(FrmFrente.Clientes.Fields("Nome"), 40)
        If (FrmFrente.Clientes.Fields("Endereco")) <> "" Then
           Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Endereco"), 35)) & ", " & FrmFrente.Clientes("Nro")
        End If
        If FrmFrente.Clientes("Bairro") <> "" Then
           Print #1, Tab(1); FrmFrente.Clientes("Bairro")
        End If
        If FrmFrente.Clientes("CIDADE") <> "" Then
           Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Cidade"), 30)) & " - " & FrmFrente.Clientes("Uf")
        End If
        Print #1, Tab(1); "QTD"; Tab(10); "DESCRICAO DO PRODUTO"
        Print #1, "------------------------------------------------"
        
'        FrmFrente.DatCompra.Recordset.MoveFirst

'        While Not FrmFrente.DatCompra.Recordset.EOF

         For y = 0 To Lista.ListCount - 1
            If Lista.Selected(y) = True Then

                Print #1, Tab(1); Lista.List(y)
            End If
        Next
        
        Print #1, ""
        Print #1, ""
        Print #1, Centralizar(FrmFrente.Empresa("MensagemRodape"), 48)
        Print #1, "" & Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
        Print #1, ""
        Print #1, ""
        
        Close #1
        Unload Me
Else

        Open VarImprEntrega For Output As #1

        Print #1, "" & Cond
        Print #1, ""
        Print #1, Gde & Centralizar(FrmFrente.Empresa("Cabecalho1"), 40) & Nor
        Print #1, Centralizar(FrmFrente.Empresa("Cabecalho2"), 79)
'        Print #1, "================================================"
        Print #1, "============================================================================="
        
        Print #1, "Data:" & Date & " Hora:" & Time; Tab(33)
        Print #1, Tab(1); psCompString(1, FrmFrente.Clientes("CodEntidade"), 6); "-"; Left(FrmFrente.Clientes.Fields("Nome"), 40)
        If (FrmFrente.Clientes.Fields("Endereco")) <> "" Then
           Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Endereco"), 35)) & ", " & FrmFrente.Clientes("Nro")
        End If
        If FrmFrente.Clientes("Bairro") <> "" Then
           Print #1, Tab(1); FrmFrente.Clientes("Bairro")
        End If
        If FrmFrente.Clientes("CIDADE") <> "" Then
           Print #1, Tab(1); Trim(Left(FrmFrente.Clientes.Fields("Cidade"), 30)) & " - " & FrmFrente.Clientes("Uf")
        End If
        Print #1, Tab(1); "QTD"; Tab(10); "DESCRICAO DO PRODUTO"
'        Print #1, "------------------------------------------------"
        Print #1, "-------------------------------------------------------------------------------"
        
        
'        FrmFrente.DatCompra.Recordset.MoveFirst

'        While Not FrmFrente.DatCompra.Recordset.EOF

         For y = 0 To Lista.ListCount - 1
            If Lista.Selected(y) = True Then

                Print #1, Tab(1); Lista.List(y)
                Print #1, "-------------------------------------------------------------------------------"
        
            End If
        Next
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, "   FAVOR CONFERIR A MERCADORIA NO ATO DA ENTREGA "
        Print #1, ""
        Print #1, ""
        Print #1, "   --------------------------------------------"
        Print #1, "          Confirmacao de recebimento "
        Print #1, "" '& Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
        Print #1, ""
        Print #1, SaltoPag
        
        Close #1
        
        Unload Me
        
        If VarImprEntrega = "C:\TELA.TXT" Then
            FrmRelTXT.Config "C:\TELA.Txt"
        End If
        
        

End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

        Lista.Clear
        FrmFrente.DatCompra.Recordset.MoveFirst
        While Not FrmFrente.DatCompra.Recordset.EOF
            Lista.AddItem Alinhar(FrmFrente.DatCompra.Recordset("Qtd"), 8, 3) & " - " & Left(FrmFrente.DatCompra.Recordset("Descrição_Produto") & " " & FrmFrente.DatCompra.Recordset("Complemento"), 37)
            Lista.Selected(Lista.ListCount - 1) = True
            FrmFrente.DatCompra.Recordset.MoveNext
        Wend

End Sub


Private Sub CmdMarcar_Click()
Dim x As Integer
For x = 1 To Lista.ListCount
    If CmdMarcar.Caption = "&Desmarcar" Then
        Lista.Selected(x - 1) = False
    Else
        Lista.Selected(x - 1) = True
    End If
Next
If CmdMarcar.Caption = "&Desmarcar" Then
    CmdMarcar.Caption = "&Marcar"
Else
    CmdMarcar.Caption = "&Desmarcar"
End If
End Sub

