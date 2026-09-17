VERSION 5.00
Begin VB.Form FormImportaTab 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6870
   Begin VB.Data DatDados 
      Caption         =   "abcfarma"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   810
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6900
      TabIndex        =   10
      Top             =   0
      Width           =   6900
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Importa Tabela Preços Remédios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   135
         TabIndex        =   11
         Top             =   45
         Width           =   6045
      End
   End
   Begin VB.CommandButton BtAtualiza 
      Caption         =   "&Atualiza Produtos"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1980
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   6840
      TabIndex        =   4
      Top             =   2880
      Width           =   6900
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "AGUARDE . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   2160
         TabIndex        =   5
         Top             =   45
         Width           =   2475
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Está operação importa a tabela de preços de remédio, pode demorar alguns minutos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   735
         Left            =   135
         TabIndex        =   8
         Top             =   0
         Width           =   6615
      End
   End
   Begin VB.TextBox TxtCont 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1890
      TabIndex        =   3
      Top             =   945
      Width           =   1410
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   315
      TabIndex        =   2
      Top             =   1350
      Width           =   5685
   End
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Top             =   945
      Width           =   1365
   End
   Begin VB.CommandButton BtImporta 
      Caption         =   "&Importa Tabela Nova"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Produtos"
      Height          =   195
      Left            =   1935
      TabIndex        =   9
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "FormImportaTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabela As Recordset

Private Sub BtAtualiza_Click()
Dim SQL As String
Dim VarTabPrd As Recordset
Dim VarProdutos As Recordset
Dim Criterio As String
Dim VarCodPrd As Double

'On Error GoTo erro

If MsgBox("Confirma Atualização ?", vbYesNo, App.Title) = vbYes Then

    Set VarTabPrd = Banco.OpenRecordset("SELECT * FROM TABELA ORDER BY MED_ABC")
    
    Set VarProdutos = Banco.OpenRecordset("SELECT BAR.BarBarra, PRD.CodPrd, PRD.Custo, PRD.Venda, Prd.Custo, PRD.Margem, Prd.MargemPrazo, Prd.PcoPrazo FROM PRD INNER JOIN BAR ON PRD.CodPrd = BAR.Produto ORDER BY BAR.BarBarra ")
    
    BtAtualiza.Enabled = False
    Label2.Visible = False
    Label1.Visible = True
'    DoEvents
'    Sql = "UPDATE TABELA INNER JOIN PRD ON TABELA.MED_ABC = PRD.Classe SET PRD.Custo = [tabela].[med_pla1], PRD.Venda = IIf([Prd].[Fator]>1,[tabela].[med_pco1],[tabela].[med_pco1]/[Prd].[Fator]) WHERE (((TABELA.MED_PLA1)>[prd].[custo]) or ((TABELA.MED_PCO1)>[prd].[venda]))"
    
    BeginTrans
    
    If VarTabPrd.RecordCount > 0 Then
        VarTabPrd.MoveFirst
        While Not VarTabPrd.Eof
            Criterio = "Bar.BarBarra = " & """" & VarTabPrd!Med_barra & """"
            VarProdutos.FindFirst Criterio
'            If VarProdutos!CodPrd = 4385 Then
'               MsgBox " dlkflkd"
'            End If
            If VarTabPrd!Med_pla1 <> 0 And VarTabPrd!Med_Pco1 <> 0 Then
                VarCodPrd = VarProdutos!CodPrd
                VarProdutos.Edit
'                If VarTabPrd!Fator > 1 Then
'                    VarProdutos!Custo = VarTabPrd!Med_pla1 / VarTabPrd!Fator
'                    VarProdutos!Venda = VarTabPrd!Med_Pco1 / VarTabPrd!Fator
'                Else
                    VarProdutos!Custo = VarTabPrd!Med_pla1
                    VarProdutos!Venda = VarTabPrd!Med_Pco1
'                End If
                VarProdutos.Update
                Criterio = "Prd.CodPrd = " & VarCodPrd
                VarProdutos.FindFirst Criterio
                If Not VarProdutos.NoMatch Then

                    VarProdutos.Edit
                    
                    If VarProdutos!Custo <> 0 Then
                        VarProdutos!Margem = ((VarProdutos!Venda / VarProdutos!Custo) - 1) * 100
                    End If
                    If IsNumeric(VarProdutos!margemprazo) Then
                        VarProdutos!PcoPrazo = VarProdutos!Venda + (VarProdutos!Venda * (VarProdutos!margemprazo / 100))
                    Else
                        VarProdutos!PcoPrazo = VarTabPrd!Venda
                    End If
                    VarProdutos.Update
                End If
            End If
            VarTabPrd.MoveNext
        Wend
            
    Else
        MsgBox "Tabela vazia", vbInformation, App.Title
    End If
    
'        Banco.Execute Sql
    
    CommitTrans
    
    BtAtualiza.Enabled = True
    Label2.Visible = True
    Label1.Visible = False
    MsgBox "Tabela de preços atualizada", vbInformation, App.Title

End If

Exit Sub

Erro:
    
    BtAtualiza.Enabled = True
    Label2.Visible = True
    Label1.Visible = False
    
    Rollback
    MsgBox "Tabela de preços não pode ser atualizada ! Erro: " & Err.Description & Err.Number, vbCritical, Caminho
End Sub

Private Sub BtImporta_Click()
Dim Linha As String
Dim Cont As Double
Dim X As Integer
Dim SQL As String
Dim Campo(1 To 20) As String
On Error GoTo Erro

If Dir(App.Path & "\tabela.txt") = "" Then
    MsgBox "Arquivo da tabela de preços não está presente", vbInformation, App.Title
Else
    If MsgBox("Confirma Importação ? ", vbYesNo, App.Title) = vbYes Then
        BtImporta.Enabled = False
        Label2.Visible = False
        Label1.Visible = True
        DoEvents
        SQL = "Delete * from Tabela"
        Banco.Execute SQL
        Cont = 0
        Open App.Path & "\TABELA.TXT" For Input As #1
        Do
            Line Input #1, Linha
            Campo(1) = Mid$(Linha, 1, 6) ' Codigo
            TxtCod = Campo(1)
'            Campo(2) = Mid$(Linha, 10, 1) 'Controle
'            Campo(3) = Mid$(Linha, 11, 6) ' Cod. do Lab
            Campo(4) = Mid$(Linha, 76, 18) ' Laboratorio
            Campo(5) = Mid$(Linha, 20, 32) ' Descr. Med.
            txtNome = Campo(5)
            Campo(6) = " " ' Apresentação Med.
            Campo(7) = Mid$(Linha, 53, 11) ' Preco do Lab.
            Campo(8) = Mid$(Linha, 64, 11) ' Preco Max Venda
'            Campo(9) = Mid$(Linha, 159, 11) ' Preco Fracao
'            Campo(10) = Mid$(Linha, 170, 3) ' Unidade
'            Campo(11) = Mid$(Linha, 173, 5) ' Perc Ipi
'            Campo(12) = Mid$(Linha, 178, 8) ' Data Vigencia
'            Campo(13) = Mid$(Linha, 186, 1) ' Campo Logico True = Novo, False = Existente
            Campo(14) = Mid$(Linha, 7, 13) ' Cod. Barra
'            Campo(15) = Mid$(Linha, 200, 3) ' Med. Genérico
'            Campo(16) = Mid$(Linha, 203, 3) ' Lista Positiva/negativa
'            Campo(17) = Mid$(Linha, 206, 3) ' Lista Neutra
'            Campo(18) = Mid$(Linha, 209, 130) ' Principio Ativo
            With Tabela
                    .AddNew
                    For X = 1 To 18
                        If Not IsNull(Campo(X)) And Campo(X) <> "" Then
                            If X = 7 Or X = 8 Or X = 9 Or X = 11 Then
                                .Fields(X - 1) = Format(Campo(X) / 1000, "###,#00.00")
                            ElseIf X = 12 Then
                                    If Campo(X) <> "00000000" Then
                                        .Fields(X - 1) = Format(Campo(X), "##/##/####")
                                    End If
                                Else
                                    If Campo(X) = "F" Or Campo(X) = "T" Then
                                        If Campo(X) = "F" Then
                                            .Fields(X - 1) = 0
                                        Else
                                            .Fields(X - 1) = 1
                                        End If
                                    Else
                                        .Fields(X - 1) = Campo(X)
                                    End If
                            End If
                        End If
                    Next
                    .Update
                    Cont = Cont + 1
                    TxtCont = Cont
                    DoEvents
            End With
            DoEvents
        Loop Until Linha = ""
        Close #1
    End If
End If

BtImporta.Enabled = True
Label1.Visible = False
Label2.Visible = True

Exit Sub
Erro:
If Err.Number = 62 Then
    Close #1
    MsgBox "Importação Concluida !", vbInformation, App.Title
    BtImporta.Enabled = True
    Label1.Visible = False
    Label2.Visible = True
End If

End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

Label1.Visible = False

'DatDados.DatabaseName = Caminho & "\AbcFarma.mdb"
'DatDados.RecordSource = "SELECT * FROM Dados "


Set Tabela = Banco.OpenRecordset("SELECT *  From Tabela  ORDER BY Med_Des")


End Sub
