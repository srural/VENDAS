VERSION 5.00
Begin VB.Form FormZebra 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -450
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   8
      Top             =   6390
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
      ScaleWidth      =   14655
      TabIndex        =   6
      Top             =   0
      Width           =   14655
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão de Etiqueta de Código de Barras"
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
         TabIndex        =   7
         Top             =   45
         Width           =   7800
      End
   End
   Begin VB.TextBox TxtCont 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4455
      TabIndex        =   4
      Top             =   5895
      Width           =   870
   End
   Begin VB.CommandButton CmbDesmarcar 
      Caption         =   "&Desmarcar Todos"
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
      Left            =   360
      TabIndex        =   3
      Top             =   5850
      Width           =   1770
   End
   Begin VB.CommandButton CmbImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   6390
      TabIndex        =   2
      Top             =   5850
      Width           =   1320
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   360
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   7395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Etiquetas Impressas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2565
      TabIndex        =   5
      Top             =   5895
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os produtos a serem impressos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   1
      Top             =   795
      Width           =   2910
   End
End
Attribute VB_Name = "FormZebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbDesmarcar_Click()
Dim y As Integer
    For y = 0 To List.ListCount - 1
          List.Selected(y) = False
    Next
End Sub

Private Sub CmbImprimir_Click()
Dim y As Integer
Dim Zebra As Boolean
Dim Formula As String
Dim Prod As Recordset
Dim Busca As String
Dim Fs, Arq
Dim VarArq As String
Dim VarPorta As String
Dim Cont As Integer
Cont = 0
TxtCont = 0


If VarImprBarra = "" Then
'        FormCopia.CommonDialog.ShowPrinter
        MDIPrincipal.RptRel.ReportFileName = (Caminho & "\RelZebra.rpt")
        MDIPrincipal.RptRel.WindowTitle = "Etiquetas"
        Formula = " ({Prd.CodPrd} = " & Left(List.List(y), 6) & " and length(totext({Bar.BarBarra})) >= 8) "
        For y = 0 To List.ListCount - 1
            If List.Selected(y) = True Then
                Formula = Formula & " or ({Prd.CodPrd} = " & Left(List.List(y), 6) & " and length(totext({Bar.BarBarra})) >= 8) "
            End If
        Next
        MDIPrincipal.RptRel.SelectionFormula = Formula
        
        'MDIPrincipal.RptRel.SelectionFormula = FormAtuCxa.SqlPrinter
        
        MDIPrincipal.RptRel.WindowState = crptMaximized
        MDIPrincipal.RptRel.RetrieveDataFiles
        MDIPrincipal.RptRel.PrintReport
        Exit Sub

End If


If Dir(Caminho & "\Zebra") = "" Then

        Set Prod = Banco.OpenRecordset("SELECT PRD.Descrição_Produto, PRD.CodPrd, PRD.Embalagem, BAR.BarBarra, PRD.Preco, PRD.Venda, PRD.Venda_Promocao, PRD.Promocao FROM BAR INNER JOIN PRD ON BAR.Produto = PRD.CodPrd")
    
        VarArq = Caminho & "\110x90.pbt"
        VarPorta = VarImprBarra
    
'        Open "110x90.pbt" For Output As #1
        Open VarImprBarra For Output As #1
        
        For y = 0 To List.ListCount - 1
            If List.Selected(y) = True Then
                Busca = "CodPrd = " & Left(List.List(y), 6)
                Prod.FindFirst Busca
                If Not Prod.NoMatch Then
                    Cont = Cont + 1
                    TxtCont = TxtCont + 1
                    DoEvents
                    If Cont = 20 Then
                        Close #1
                        MsgBox "Impressas " & Cont & " etiquetas", vbInformation, App.Title
                        Open VarImprBarra For Output As #1
                        Cont = 0
                    End If
'                    If Cont = 30 Then
'                        Close #1
'                        Open VarImprBarra For Output As #1
'                        Cont = 0
'                    End If
                    Print #1, Chr(2) + "O0000"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    Print #1, "191200500500000" & psCompString(3, Left(Prod!Descrição_Produto, 25), 25)
                    If Not IsNull(Prod!BarBarra) Then
                        Print #1, "1F1202000100010" & psCompString(1, Prod!BarBarra, 13) ' Codigo Barras
                    End If
'                    Print #1, "192300500000125" & "R$"
                    Print #1, "192300400000125" & "R$"
                    If Prod!Promocao Then
                        Print #1, "142300000000195" & Alinhar(Prod!Venda_Promocao, 6, 2) ' Valor do Produto
                    Else
                        Print #1, "142300000000195" & Alinhar(Prod!Venda, 6, 2) ' Valor do Produto
                    End If
                    Print #1, "191100100400020" & psCompString(3, NomeEmpresa, 20)
                    Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
                    Print #1, "E"
                    
'                    Close #1
                    
                End If
            End If
        Next
        Close #1
                            
    
Else
'   FormCopia.CommonDialog.ShowPrinter
    
    Set Prod = Banco.OpenRecordset("SELECT PRD.Descrição_Produto, PRD.CodPrd, PRD.Embalagem, BAR.BarBarra, PRD.Preco , PRD.Venda, PRD.Venda_Promocao, PRD.Promocao FROM BAR INNER JOIN PRD ON BAR.Produto = PRD.CodPrd")
    
'    For y = 0 To List.ListCount - 1
            
            
        Open VarImprBarra For Output As #1
        
        For y = 0 To List.ListCount - 1
            If List.Selected(y) = True Then
                Busca = "CodPrd = " & Left(List.List(y), 6)
                Prod.FindFirst Busca
                If Not Prod.NoMatch Then
                    Cont = Cont + 1
                    TxtCont = TxtCont + 1
                    DoEvents
                    If Cont = 20 Then
                        Close #1
                        MsgBox "Impressas " & Cont & " etiquetas", vbInformation, App.Title
                        Open VarImprBarra For Output As #1
                        Cont = 0
                    End If
'                    If Cont = 30 Then
'                        Close #1
'                        Open VarImprBarra For Output As #1
'                        Cont = 0
'                    End If
                    Print #1, Chr(2) + "O0000"
                    Print #1, Chr(2) + "M0300"
                    Print #1, Chr(2) + "c0000"
                    Print #1, Chr(2) + "f000"
                    Print #1, Chr(2) + "e"
                    Print #1, Chr(2) + "LC0000"
                    Print #1, "H09"
                    Print #1, "D11"
                    Print #1, "SC"
                    Print #1, "PC"
                    Print #1, "R0000"
                    Print #1, "z"
                    Print #1, "W"
                    Print #1, Chr(94) + "01"
                    Print #1, "191200500500010" & psCompString(3, Left(Prod!Descrição_Produto, 33), 33) & "-" & Prod!CodPrd
                    If Not IsNull(Prod!BarBarra) Then
                        Print #1, "1F1202000100010" & psCompString(1, Prod!BarBarra, 13) ' Codigo Barras
                    End If
'                    Print #1, "192300500000125" & "R$"
                    Print #1, "192300400000130" & "R$"
                    If Prod!Promocao Then
                        Print #1, "142300000000200" & Alinhar(Prod!Venda_Promocao, 6, 2) ' Valor do Produto
                    Else
                        Print #1, "142300000000200" & Alinhar(Prod!Venda, 6, 2) ' Valor do Produto
                    End If
                    Print #1, "191100100400010" & psCompString(3, NomeEmpresa, 20)
                    Print #1, "Q" & "0001"  ' Quantidade de Etiquetas
                    Print #1, "E"
                    
'                    Close #1
                    
                End If
            End If
        Next
        Close #1


            
            
            
            
'            If List.Selected(y) = True Then
'                Printer.Font.Name = "Arial"
'                Printer.Font.Size = 14
'                Printer.Font.Bold = True
'                Busca = "CodPrd = " & Left(List.List(y), 6)
'                Prod.FindFirst Busca
'                If Not Prod.NoMatch Then
'                    Printer.Print psCompString(2, Left(Trim(Prod!Descrição_Produto), 35), 45)
'                    Printer.Font.Size = 10
'                    Printer.Print Space(30); Prod!Embalagem & Space(10);
'                    Printer.Print Prod!CodPrd
'                    Printer.Font.Size = "6"
'                    Printer.Print Space(50)
'                    Printer.Font.Name = "c39hrp48dhtt"
'    '                Printer.Font.Name = "EAN-13B Half Height"
'                    Printer.Font.Size = "42"
'                    Printer.Font.Bold = False
'                    Printer.Print Space(3) & Prod!BARBARRA;
'                    Printer.Font.Name = "Arial"
'                    Printer.Font.Italic = True
'                    Printer.Font.Size = "18"
'                    Printer.Print Space(20 - Len(Trim(Prod!BARBARRA))) & "R$  ";
'                    Printer.Font.Size = 32
'                    Printer.Font.Bold = True
'                    Printer.Font.Italic = True
'                    If Prod!Promocao Then
'                        Printer.Print Alinhar(Prod!Venda_Promocao, 6, 2)
'                    Else
'                        Printer.Print Alinhar(Prod!Preco, 6, 2)
'                    End If
'                    Printer.Font.Italic = False
'                    Printer.EndDoc
'                End If
'            End If
'    Next
    
End If
Exit Sub

Erro:
    MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub Form_Activate()
FormAtuCxa.DatDados(0).Refresh
If FormAtuCxa.DatDados(0).Recordset.RecordCount <> 0 Then
    FormAtuCxa.DatDados(0).Recordset.MoveFirst
    List.Clear
    While Not FormAtuCxa.DatDados(0).Recordset.Eof
        List.AddItem psCompString(1, FormAtuCxa.DatDados(0).Recordset.Fields("CodPrd"), 6) & " - " & FormAtuCxa.DatDados(0).Recordset.Fields("Descrição_Produto")
        If FormAtuCxa.DatDados(0).Recordset.Fields("Promocao") Then
            List.Selected(List.ListCount - 1) = False
        Else
            List.Selected(List.ListCount - 1) = True
        End If
        FormAtuCxa.DatDados(0).Recordset.MoveNext
    Wend
End If

End Sub

