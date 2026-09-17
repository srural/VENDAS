VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormLimpaCampos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDestino 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10680
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox TxtOrigem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8400
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton TransfEnt 
      Caption         =   "Tranfere Entidade"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   2355
   End
   Begin VB.CheckBox ChkTotal 
      Caption         =   "Loja => Estoque"
      Height          =   330
      Left            =   6570
      TabIndex        =   24
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1680
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LOJ"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -540
      ScaleHeight     =   345
      ScaleWidth      =   13890
      TabIndex        =   23
      Top             =   8100
      Width           =   13920
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9270
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select CodEntidade, Nome, Tipo From Ent  Where Tipo = 2 Order bY  Nome"
      Top             =   1350
      Visible         =   0   'False
      Width           =   2490
   End
   Begin MSDBCtls.DBCombo DBCombo 
      Bindings        =   "FormLimpaCampos.frx":0000
      DataSource      =   "Data4"
      Height          =   315
      Left            =   8415
      TabIndex        =   19
      Top             =   3780
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descrição_Grupo"
      BoundColumn     =   "CodGru"
      Text            =   ""
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From Gru Order bY  Descrição_Grupo"
      Top             =   1935
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13380
      TabIndex        =   17
      Top             =   0
      Width           =   13380
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Correção de Arquivos"
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
         TabIndex        =   18
         Top             =   90
         Width           =   5010
      End
   End
   Begin VB.ListBox LstLista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   270
      TabIndex        =   16
      Top             =   5895
      Width           =   12780
   End
   Begin VB.CommandButton CmdVerifica 
      Caption         =   "&Verifica Intergridade"
      Height          =   375
      Left            =   270
      TabIndex        =   14
      Top             =   5400
      Width           =   2355
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Grava Estoque"
      Height          =   375
      Left            =   255
      TabIndex        =   12
      Top             =   4365
      Width           =   2355
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   9135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2565
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Corrige estoque"
      Height          =   375
      Left            =   255
      TabIndex        =   10
      Top             =   3225
      Width           =   2355
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Zerar Estoque"
      Height          =   375
      Left            =   255
      TabIndex        =   5
      Top             =   3795
      Width           =   2355
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ativa &Prd"
      Height          =   375
      Left            =   255
      TabIndex        =   4
      Top             =   2655
      Width           =   2355
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Corrigi &Prd"
      Height          =   375
      Left            =   255
      TabIndex        =   2
      Top             =   2100
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Corrigi &Cnpj"
      Height          =   375
      Left            =   255
      TabIndex        =   1
      Top             =   960
      Width           =   2355
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   5175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PRD"
      Top             =   2610
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Corrigi C&pf"
      Height          =   375
      Left            =   255
      TabIndex        =   0
      Top             =   1530
      Width           =   2355
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   7245
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ENT"
      Top             =   2565
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "FormLimpaCampos.frx":0014
      DataSource      =   "Data5"
      Height          =   315
      Left            =   8415
      TabIndex        =   21
      Top             =   4365
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   "CodEntidade"
      Text            =   ""
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      Height          =   255
      Left            =   10680
      TabIndex        =   31
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Origem"
      Height          =   255
      Left            =   8400
      TabIndex        =   30
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   29
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Essa rotina transfere todo o moviemnto de uma entidade pra outra"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3120
      TabIndex        =   26
      Top             =   5040
      Width           =   4650
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   8415
      TabIndex        =   22
      Top             =   4140
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
      Height          =   195
      Left            =   8415
      TabIndex        =   20
      Top             =   3555
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Essa rotina verifica e intergridade do arquivo de entidades"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   15
      Top             =   5445
      Width           =   4095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Essa rotina grava o estoques. Certifique-se de estar no final do dia."
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   13
      Top             =   4455
      Width           =   4710
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Corrigi divergência de estoque entre as lojas"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   11
      Top             =   3315
      Width           =   3120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muito Cuidado.. essa rotina zera todos o estoque de produto"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   9
      Top             =   3885
      Width           =   4260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ativa todos os produto"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   8
      Top             =   2745
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esta rotina é usada para limpa aspas do campo descrição do produto"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   7
      Top             =   2190
      Width           =   4890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esta rotina é usada para limpar pontos, barras  e traços dos campos  Cpf"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   6
      Top             =   1620
      Width           =   5115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esta rotina é usada para limpar pontos, barras  e traços do campo  Cgc"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3150
      TabIndex        =   3
      Top             =   1050
      Width           =   5010
   End
End
Attribute VB_Name = "FormLimpaCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub CmdVerifica_Click()
Dim PortaTemp As String
'Dim Porta As String
Dim Arq As Recordset
Dim varCpf As String
Dim VarCgc As String
Dim VarCep As String
Dim VarUf As String

    PortaTemp = Porta
    Porta = Caminho & "\ChkEnt.txt"
        
    Set Arq = Banco.OpenRecordset("SELECT Ent.CodEntidade, Ent.Nome, Fone, Celular, Ent.CPF, Ent.CGC, Ent.Uf, Ent.Cep, Classificacao FROM Ent WHERE (((Ent.Classificacao)<=2)) ORDER BY Ent.Classificacao, Ent.Nome")
    Arq.Requery
                
    
    If Arq.RecordCount > 0 Then
        Iniciar False 'inicia o registro onde sera gravado as informacoes
        While Not Arq.Eof
                varCpf = "    "
                VarCgc = "    "
                VarUf = "    "
                VarCep = "    "
        
                If Arq!Classificacao = 1 And Trim(Arq!Cpf) <> "" Then
                    If CalculaDacCPF(Format(Trim(Arq!Cpf), "@@@.@@@.@@@-@@")) = True Then
                        varCpf = "    "
                    Else
                        varCpf = "CPF "
                    End If
                End If
                If Arq!Classificacao = 2 And Trim(Arq!CGC) <> "" Then
                    If VerificaCGC(Format(Trim(Arq!CGC), "@@.@@@.@@@/@@@@-@@")) = True Then
                        VarCgc = "    "
                    Else
                        VarCgc = "CGC "
                    End If
                End If
'                If Arq!Uf = "" Then
'                    If Len(Arq!Uf) <> 2 Then
'                        VarUf = "    "
'                    Else
'                        VarUf = "UF  "
'                    End If
'                End If
'                If Arq!cep <> "" Then
'                    If Len(Arq!cep) <> 8 Then
'                        VarCep = "    "
'                    Else
'                        VarCep = "CEP "
'                    End If
'                End If
'
                If Len(Trim(varCpf & VarCgc & VarUf & VarCep)) <> 0 Then
                    Print #1, Arq!CodEntidade; Tab(10); Arq!Nome; Tab(55); Format(Arq!Fone, "## ####-####"); Tab(68); Format(Arq!celular, "## ####-####"); Tab(83); varCpf; VarCgc; VarUf; VarCep
                End If
                Arq.MoveNext
                               
        Wend
        Fim False
    End If
    On Error GoTo Erro
    
    LstLista.Visible = True
    LstLista.Clear
    Porta = PortaTemp
    If Dir(Caminho & "\ChkEnt.Txt") <> "" Then
        Open Caminho & "\ChkEnt.Txt" For Input As #1
        Do
            Line Input #1, PortaTemp
            LstLista.AddItem PortaTemp
        Loop Until PortaTemp = "<>"
        Close #1
    End If

Erro:
If Err.Number = 62 Then
    Close #1
'    SSTabNota.Tab = 2
End If

End Sub

Private Sub Command1_Click()
DesabilitaBotoes

Data1.Refresh
Data1.Recordset.MoveFirst
While Not Data1.Recordset.Eof
    If Not IsNull(Data1.Recordset.Fields("Cpf")) Then
        Data1.Recordset.Edit
        Data1.Recordset.Fields("Cpf") = LIMPACampo(Data1.Recordset.Fields("Cpf"))
        Data1.Recordset.Update
    End If
    Data1.Recordset.MoveNext
Wend

MsgBox "Campo CPF limpo"

AbilitaBotoes

End Sub
Function LIMPACampo(Campo As String) As String
Dim X As Integer
Dim CampoLimpo As String

For X = 1 To Len(Campo)
    If Mid$(Campo, X, 1) >= Chr(32) And Mid$(Campo, X, 1) <= Chr(122) And Mid$(Campo, X, 1) <> "." And Mid$(Campo, X, 1) <> "/" And Mid$(Campo, X, 1) <> "-" Then
            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
    End If
Next
LIMPACampo = psCompString(3, CampoLimpo, Len(Campo))
End Function

Function LIMPACampo2(Campo As String) As String
Dim X As Integer
Dim CampoLimpo As String

For X = 1 To Len(Campo)
    If Mid$(Campo, X, 1) <> "'" And Mid$(Campo, X, 1) <> Chr(34) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
    End If
Next
LIMPACampo2 = psCompString(3, CampoLimpo, Len(Campo))
End Function


Public Function psCompString(ByVal liTipoComplementar As Integer, _
                                                ByVal psString As String, _
                                                ByVal liTamanho As Integer) As String
'Objetivo: Completa uma string com zeros ou brancos a esquerda
'Entradas: liTipoComplementar define se zero ou branco.
'   1 = Zeros a esquerda; _
    2 = Brancos a esquerda; _
    3 = Brancos a direira. _
                  psString é a String passada para a função. _
                  liTamanho define a quantidade de Brancos ou espaços a acrescentar.
'Saídas: A string montada.
'Autor: Ernani Medeiros
'Data: 15/07/1997
    
    Dim lsStringFinal As String
    Dim liQtdZeros As Integer
    Dim liQtdBrancos As Integer
    
    Select Case liTipoComplementar
        Case 1
        '1 - Zeros a esquerda
            lsStringFinal = ""
            'Verifica quantos zeros são necessários
            liQtdZeros = liTamanho - Len(psString)
            'Se for necessário mais do que 1
            'a lsStringFinal é completada
            If liQtdZeros > 0 Then lsStringFinal = String(liQtdZeros, "0")
            psCompString = lsStringFinal & psString
        Case 2
        '2 - Brancos a esquerda
            lsStringFinal = ""
            'Verifica quantos zeros são necessários
            liQtdBrancos = liTamanho - Len(psString)
            'Se for necessário mais do que 1
            'a lsStringFinal é completada
            If liQtdBrancos > 0 Then lsStringFinal = String(liQtdBrancos, 32)
            psCompString = lsStringFinal & psString
        Case 3
        '3 - Brancos a direita
            lsStringFinal = ""
            'Verifica quantos Brancos são necessários
            liQtdBrancos = liTamanho - Len(psString)
            'Se for necessário mais do que 1
            'a lsStringFinal é completada
            If liQtdBrancos > 0 Then lsStringFinal = String(liQtdBrancos, 32)
            psCompString = psString & lsStringFinal
    End Select
End Function


Private Sub Command2_Click()
DesabilitaBotoes

Data1.Refresh
Data1.Recordset.MoveFirst
While Not Data1.Recordset.Eof
    If Not IsNull(Data1.Recordset.Fields("Cgc")) Then
        Data1.Recordset.Edit
        Data1.Recordset.Fields("Cgc") = LIMPACampo(Data1.Recordset.Fields("Cgc"))
        If IsNull(Data1.Recordset.Fields("Juros")) Then
            Data1.Recordset.Fields("Juros") = 0
        End If
        If IsNull(Data1.Recordset.Fields("Prazo")) Then
            Data1.Recordset.Fields("Prazo") = 0
        End If
        Data1.Recordset.Update
    End If
    Data1.Recordset.MoveNext
Wend

MsgBox "Campo CGC limpo"

AbilitaBotoes

End Sub

Private Sub Command3_Click()
DesabilitaBotoes

Data2.Refresh
Data2.Recordset.MoveFirst
While Not Data2.Recordset.Eof
    If Not IsNull(Data2.Recordset.Fields("Descrição_Produto")) Then
        Data2.Recordset.Edit
        Data2.Recordset.Fields("Descrição_Produto") = LIMPACampo2(Data2.Recordset.Fields("Descrição_Produto"))
        Data2.Recordset.Update
    End If
    Data2.Recordset.MoveNext
Wend

MsgBox "Campo Descrição do Produto limpo"

AbilitaBotoes

End Sub

Private Sub Command4_Click()
DesabilitaBotoes

If MsgBox("Atenção esta rotina torna todos os produtos ativos. Deseja continuar", vbYesNo, App.Title) = vbYes Then
    Data2.Refresh
    Data2.Recordset.MoveFirst
    While Not Data2.Recordset.Eof
        If Not IsNull(Data2.Recordset.Fields("Descrição_Produto")) Then
            Data2.Recordset.Edit
            Data2.Recordset.Fields("Ativo") = True
            Data2.Recordset.Update
        End If
        Data2.Recordset.MoveNext
    Wend
    
    MsgBox "Todos os Produtos estão Ativo"
End If
AbilitaBotoes

End Sub

Private Sub Command5_Click()
Dim Sql As String
DesabilitaBotoes

If DBCombo.Text = "" Then

    If DBCombo1.Text = "" Then
        If MsgBox("Atenção...  esta rotina zera o estoque de produtos. Deseja continuar", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Command7_Click
            Sql = "UPDATE PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto SET LOJ.Estoque = 0, PRD.Estoque = 0"
            Banco.Execute Sql
            MsgBox "Estoque Zerado"
        End If
    Else
        If MsgBox("Atenção...  esta rotina zera o estoque de produtos do Fornecedor " & DBCombo1.Text & ". Deseja continuar", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Command7_Click
            Sql = "UPDATE PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto SET LOJ.Estoque = 0, PRD.Estoque = 0 WHERE  Prd.Fabricante = " & DBCombo1.BoundText
            Banco.Execute Sql
            MsgBox "Estoque Zerado"
        End If
    
    End If
Else
    If DBCombo1.Text = "" Then
        If MsgBox("Atenção...  esta rotina zera o estoque de produtos do grupo " & DBCombo.Text & ". Deseja continuar", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Command7_Click
            Sql = "UPDATE PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto SET LOJ.Estoque = 0, PRD.Estoque = 0 WHERE PRD.Grupo = " & DBCombo.BoundText
            Banco.Execute Sql
            MsgBox "Estoque Zerado"
        End If
    Else
        If MsgBox("Atenção...  esta rotina zera o estoque de produtos do grupo " & DBCombo.Text & " e do Fornecedor " & DBCombo1.Text & ". Deseja continuar", vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Command7_Click
            Sql = "UPDATE PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto SET LOJ.Estoque = 0, PRD.Estoque = 0 WHERE PRD.Grupo = " & DBCombo.BoundText & " and Prd.Fabricante = " & DBCombo1.BoundText
            Banco.Execute Sql
            MsgBox "Estoque Zerado"
        End If
    
    End If
End If
AbilitaBotoes


End Sub

Private Sub Command6_Click()
Dim Criterio As String
DesabilitaBotoes

If MsgBox("Atenção esta rotina corrige os estoques. Deseja continuar", vbYesNo, App.Title) = vbYes Then
    Data3.RecordSource = "SELECT Sum(LOJ.Estoque) AS SomaDeEstoque, PRD.CodPrd, First(PRD.Estoque) AS PrimeiroDeEstoque FROM PRD INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto GROUP BY PRD.CodPrd HAVING (((First(PRD.Estoque))<>Sum([LOJ].[Estoque])))"
    Data3.Refresh
    If Data3.Recordset.RecordCount > 0 Then
        Data3.Recordset.MoveFirst
        While Not Data3.Recordset.Eof
            If ChkTotal.value = 1 Then
                Criterio = "CodPrd = " & Data3.Recordset.Fields("CodPrd")
                Data2.Recordset.FindFirst Criterio
                If Not IsNull(Data2.Recordset.Fields("Descrição_Produto")) Then
                        Data2.Recordset.Edit
                        Data2.Recordset.Fields("Estoque") = Data3.Recordset.Fields("SomaDeEstoque")
                        Data2.Recordset.Update
                End If
            Else
                Criterio = "Produto = " & Data3.Recordset.Fields("CodPrd") & " and DescricaoLoja =  " & """" & "LOJA" & """"
                Data6.Recordset.FindFirst Criterio
                If Not Data6.Recordset.NoMatch Then
                        Data6.Recordset.Edit
                        Data6.Recordset.Fields("Estoque") = Data3.Recordset.Fields("PrimeiroDeEstoque")
                        Data6.Recordset.Update
                End If
            
            End If
            Data3.Recordset.MoveNext
        Wend
    End If
    MsgBox "Estoque corrgido"
End If
AbilitaBotoes

End Sub

Private Sub Command7_Click()
Dim Sql As String
Dim BcoHist As Recordset
Dim VarPrd As Recordset
'Sql = " INSERT INTO HISTEST ( CodPrd, QtdEst, DTEstoque ) SELECT PRD.CodPrd, PRD.Estoque, Date() AS Expr1 FROM PRD"
'DesabilitaBotoes
'BcoEst.Execute Sql
'MsgBox "Estoque gravado com sucesso !", vbInformation, App.Title

Set BcoEst = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Estoque.mdb")


Set BcoHist = BcoEst.OpenRecordset("Select * FROM HistEst WHERE DtEstoque = #" & Format(Date, "mm/dd/yyyy") & "#")
BcoHist.Requery
If BcoHist.RecordCount = 0 Then
    Set VarEstoque = BcoEst.OpenRecordset("SELECT * FROM HistEst")
    Set VarPrd = Banco.OpenRecordset("SELECT CodPrd, Estoque FROM Prd ORDER BY CodPrd")
    VarPrd.Requery
    If VarPrd.RecordCount > 0 Then
        VarPrd.MoveFirst
        While Not VarPrd.Eof
            VarEstoque.AddNew
            VarEstoque!DtEstoque = Date
            VarEstoque!CodPrd = VarPrd!CodPrd
            VarEstoque!QtdEst = VarPrd!estoque
            VarEstoque.Update
            VarPrd.MoveNext
        Wend
    End If

'            Sql = " INSERT INTO HISTEST ( CodPrd, QtdEst, DTEstoque ) SELECT PRD.CodPrd, PRD.Estoque, Date() AS Expr1 FROM PRD"
'            BcoEst.Execute Sql
    MsgBox "Historico de estoque gravado com sucesso", vbInformation, App.Title
End If

AbilitaBotoes
End Sub

Private Sub Form_Load()
Data1.DatabaseName = Caminho & "\Dados.Mdb"
Data2.DatabaseName = Caminho & "\Dados.Mdb"
Data3.DatabaseName = Caminho & "\Dados.Mdb"
Data4.DatabaseName = Caminho & "\Dados.Mdb"
Data5.DatabaseName = Caminho & "\Dados.Mdb"
Data6.DatabaseName = Caminho & "\Dados.Mdb"
End Sub

Private Sub AbilitaBotoes()

Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True

End Sub
Private Sub DesabilitaBotoes()

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False

End Sub

Private Sub TransfEnt_Click()

If TxtOrigem <> "" And TxtDestino <> "" Then

    If MsgBox("Atenção esta rotina transfere todo o movimento. Deseja continuar", vbYesNo, App.Title) = vbYes Then
        Dim Sql As String
        
        Sql = "UPDATE PED SET PED.Entidade = " & TxtDestino & "  WHERE PED.Entidade = " & TxtOrigem
        Banco.Execute Sql
        
        Sql = "UPDATE DUP SET DUP.Entidade = " & TxtDestino & "  WHERE DUP.Entidade = " & TxtOrigem
        Banco.Execute Sql
        
        Sql = "UPDATE CONV SET CONV.Cliente = " & TxtDestino & "  WHERE CONV.Cliente = " & TxtOrigem
        Banco.Execute Sql
        
        Sql = "UPDATE REC SET REC.CodEnt = " & TxtDestino & "  WHERE REC.CodEnt = " & TxtOrigem
        Banco.Execute Sql
    
    
        MsgBox "Tranferencia Completa !", vbInformation
    End If
End If
End Sub
