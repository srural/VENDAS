VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormLivros 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10080
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   24
      Top             =   7290
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
      TabIndex        =   22
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Módulo de Geração de Arquivos Textos"
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
         Left            =   225
         TabIndex        =   23
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.Frame framePeriodo 
      Caption         =   "Periodo"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6150
      TabIndex        =   19
      Top             =   1545
      Width           =   1545
      Begin VB.ComboBox cbPeriodo 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "FormLivros.frx":0000
         Left            =   240
         List            =   "FormLivros.frx":0010
         TabIndex        =   20
         Text            =   "Diario"
         Top             =   210
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   645
      Left            =   210
      TabIndex        =   14
      Top             =   1515
      Width           =   5835
      Begin MSMask.MaskEdBox MebDtI 
         Height          =   285
         Left            =   1470
         TabIndex        =   15
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDtF 
         Height          =   285
         Left            =   4080
         TabIndex        =   16
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   3090
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame FrameSerasa 
      Caption         =   "Serasa"
      Enabled         =   0   'False
      Height          =   705
      Left            =   6735
      TabIndex        =   5
      Top             =   765
      Width           =   3090
      Begin VB.OptionButton opSerasa 
         Caption         =   "&Pagamentos"
         Height          =   285
         Index           =   1
         Left            =   1650
         TabIndex        =   13
         Top             =   375
         Width           =   1275
      End
      Begin VB.OptionButton opSerasa 
         Caption         =   "&Vencidos"
         Height          =   285
         Index           =   0
         Left            =   1650
         TabIndex        =   12
         Top             =   150
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.ComboBox CmbTipo 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "FormLivros.frx":0038
         Left            =   570
         List            =   "FormLivros.frx":0054
         TabIndex        =   8
         Text            =   "PC"
         Top             =   255
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   315
      End
   End
   Begin VB.Frame FrameLivros 
      Caption         =   "Livros Fiscais"
      Height          =   705
      Left            =   5040
      TabIndex        =   4
      Top             =   765
      Width           =   1590
      Begin VB.OptionButton OPt 
         Caption         =   "Saida"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   435
         Width           =   855
      End
      Begin VB.OptionButton OPt 
         Caption         =   "Entrada"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Arquivo"
      Height          =   705
      Left            =   210
      TabIndex        =   3
      Top             =   765
      Width           =   4755
      Begin VB.OptionButton opTipo 
         Caption         =   "Contimatic"
         Height          =   195
         Index           =   2
         Left            =   1755
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton opTipo 
         Caption         =   "Serasa"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.OptionButton opTipo 
         Caption         =   "Livros Fiscais"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   330
         Width           =   1425
      End
   End
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Minha Pasta\Programas\Vendas-DAO\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4515
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton BtGera 
      Caption         =   "&Gera Arquivo"
      Height          =   450
      Left            =   8865
      TabIndex        =   1
      Top             =   1665
      Width           =   960
   End
   Begin VB.CommandButton BtOK 
      Caption         =   "&Pesquisa"
      Height          =   450
      Left            =   7830
      TabIndex        =   0
      Top             =   1665
      Width           =   915
   End
   Begin MSDBGrid.DBGrid DbGrid 
      Bindings        =   "FormLivros.frx":0078
      Height          =   4770
      Left            =   210
      OleObjectBlob   =   "FormLivros.frx":008B
      TabIndex        =   2
      Top             =   2235
      Width           =   9630
   End
End
Attribute VB_Name = "FormLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Texto As String
Dim tabDuplicatas As Recordset

Private Sub BtGera_Click()
    If opTipo(0).value = True Then
        PesquisaLivros
        GeraLivros
    ElseIf opTipo(1).value = True Then
            GeraSerasa
        Else
            PesquisaLivros
            GeraContimatic
    End If
End Sub

Private Sub BtOK_Click()
    If opTipo(0).value = True Then
        PesquisaLivros
    ElseIf opTipo(1).value = True Then
        PesquisaSerasa
    End If
End Sub

Private Sub Form_Activate()

MebDtI.Text = Date
MebDtF.Text = Date

Texto = Space(768)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
'Dim Componentes As Variant

Me.Left = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6

'For Each Componentes In DatDados
    Data.DatabaseName = Caminho & "\Dados.mdb"
'Next
'CmbTipo.ListIndex = 0
End Sub

Function psString(Tipo As Currency, Valor As String, Posicao As Variant, Tamanho As Currency)
Dim Temp As String
    If Tipo = 0 Then 'espacos em branco
        Mid(Texto, Posicao, Tamanho) = Space(Tamanho)
    ElseIf Tipo = 1 Then
        Mid(Texto, Posicao, Tamanho) = Valor
    ElseIf Tipo = 2 Then 'com zeros a esquerda
        Mid(Texto, Posicao, Tamanho) = Completa(Tamanho - Len(Valor)) & Valor
    ElseIf Tipo = 3 Then 'valor
        If Valor <> 0 Then
            If opTipo(0).value = True Then
                Temp = Replace(Format(Valor, "###########.00"), ",", ".")
            Else
                Temp = Replace(Format(Valor, "##########.00"), ",", ".")
            End If
            Mid(Texto, Posicao, Tamanho) = Completa(Tamanho - Len(Temp)) & Temp
        End If
    End If
End Function

Function LimpaString(Valor As String) As String
Dim VarTemp As String
    VarTemp = Replace(Valor, ".", "")
    VarTemp = Replace(VarTemp, ",", "")
    VarTemp = Replace(VarTemp, "-", "")
    VarTemp = Replace(VarTemp, "/", "")
    VarTemp = Replace(VarTemp, "\", "")
    
    LimpaString = VarTemp
    
End Function

Function Completa(Qtd As Currency) As String
    If Qtd >= 0 Then
        For x = 1 To Qtd
            Temp = Temp & "0"
            Completa = Temp
        Next
    Else
        Completa = ""
    End If
End Function

Private Sub opTipo_Click(Index As Integer)
Select Case Index
    Case 0
        FrameLivros.Enabled = True
        FrameSerasa.Enabled = False
        framePeriodo.Enabled = False
    Case 1
        FrameLivros.Enabled = False
        FrameSerasa.Enabled = True
        framePeriodo.Enabled = True
End Select
End Sub

Private Sub PesquisaSerasa()

If opSerasa(0).value = True Then
    Data.RecordSource = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond, ENT.DtCadastro, DUP.DtPgto FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON DUP.Pedido = PED.CodPed WHERE DUP.DtVenc >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And DUP.DtVenc <= #" & Format(MebDtF, "mm/dd/yyyy") & "# AND DUP.Cond = False AND Tipo_Dup = " & """" & CmbTipo.Text & """" & " GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond, ENT.DtCadastro, DUP.DtPgto ORDER BY ENT.CodEntidade, DUP.Pedido"
Else
    Data.RecordSource = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond, ENT.DtCadastro, DUP.DtPgto FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON DUP.Pedido = PED.CodPed WHERE DUP.DtPgto >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And DUP.DtPgto <= #" & Format(MebDtF, "mm/dd/yyyy") & "# AND DUP.Cond = true AND Tipo_Dup = " & """" & CmbTipo.Text & """" & " GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond, ENT.DtCadastro, DUP.DtPgto ORDER BY Ent.CodEntidade, DUP.Pedido"
End If

Data.Refresh

If Data.Recordset.RecordCount = 0 Then
    MsgBox "Não existe dados nesse intervalo", vbInformation, App.Title
    Exit Sub
End If


End Sub

Private Sub GeraSerasa()
Dim TipoEnvio As String
Dim TipoCliente As Currency '1 = antigo, 2=menos de um ano, 3=inativo
Dim Dif As Currency
Dim CliAnt As String
Dim QtdRegRelac As Currency
Dim QtdRegDupli As Currency

Select Case cbPeriodo.ListIndex + 1
    Case 0, 1
        TipoEnvio = "D" 'diario
    Case 2
        TipoEnvio = "S" 'semanal
    Case 3
        TipoEnvio = "Q" 'quinzenal
    Case 4
        TipoEnvio = "M" 'mensal
End Select

NomeArq = App.Path & "\" & "Serasa.txt"
CliAnt = ""
QtdRegRelac = 0
QtdRegDupli = 0

If MsgBox("Confirma geração do arquivo texto ?", vbYesNo, App.Title) = vbYes Then
    
    Open NomeArq For Output As #1
    
    'header
    
    Texto = Space(100)
    psString 1, "00", 1, 2
    psString 1, "RELATO COMP NEGOCIOS", 3, 20
    psString 1, LimpaString(Cnpj), 23, 14
    psString 1, LimpaString(Format(MebDtI, "yyyy/mm/dd")), 37, 8
    psString 1, LimpaString(Format(MebDtF, "yyyy/mm/dd")), 45, 8
    psString 1, TipoEnvio, 53, 1
    psString 1, "000", 69, 3
    psString 0, "", 72, 29
    
    Print #1, Texto
    
    'detalhe
    
    With Data.Recordset
        .MoveFirst

            While Not .Eof
                
                If IsDate(Data.Recordset("DtCadastro")) Then
                    Dif = Date - CDate(Data.Recordset("DtCadastro"))
                    If Dif > 365 Then
                        TipoCliente = 1
                    ElseIf Dif < 365 Then
                        TipoCliente = 2
                    End If
                Else
                    TipoCliente = 3
                End If
                    
                'tempo de relacionamento
                If CliAnt <> Data.Recordset("CodEntidade") Then
                    CliAnt = Data.Recordset("CodEntidade")
                    QtdRegRelac = QtdRegRelac + 1
                    Texto = Space(101)
                    psString 1, "01", 1, 2
                    psString 1, LimpaString(Data.Recordset("Cgc")), 3, 14
                    psString 1, "01", 17, 2
                    psString 1, IIf(IsDate(Data.Recordset("DtCadastro")), LimpaString(Format(Data.Recordset!DtCadastro, "yyyy/mm/dd")), ""), 19, 8
                    psString 1, CCur(TipoCliente), 27, 1
                    psString 0, "", 28, 38
                    psString 0, "", 66, 34
                    psString 0, "", 28, 38
                    psString 0, "", 100, 1
                    Print #1, Texto
                End If
                
                'duplicatas e titulos
                    Texto = Space(101)
                    QtdRegDupli = QtdRegDupli + 1
                    psString 1, "01", 1, 2
                    psString 1, LimpaString(Data.Recordset("Cgc")), 3, 14
                    psString 1, "05", 17, 2
                    psString 1, Data.Recordset("Pedido"), 19, 10
                    psString 1, IIf(IsDate(Data.Recordset("DataEmiss")), LimpaString(Format(Data.Recordset!DataEmiss, "yyyy/mm/dd")), ""), 29, 8
                    psString 3, Data.Recordset("Valor"), 37, 13
                    psString 1, IIf(IsDate(Data.Recordset("DtVenc")), LimpaString(Format(Data.Recordset!DtVenc, "yyyy/mm/dd")), ""), 50, 8
                    psString 1, IIf(IsDate(Data.Recordset("DtPgto")), LimpaString(Format(Data.Recordset!Dtpgto, "yyyy/mm/dd")), ""), 58, 8
                    psString 0, "", 66, 34
                    psString 0, "", 100, 1
                
                Print #1, Texto
               .MoveNext

            Wend
    End With
    
    'trailler
    
    Texto = Space(100)
    psString 1, "99", 1, 2
    psString 2, CStr(QtdRegRelac), 3, 11
    psString 0, "", 14, 44
    psString 2, CStr(QtdRegDupli), 58, 11
    psString 0, "", 69, 32
    
    Print #1, Texto
    Close #1
    
    MsgBox "Arquivo gerado com sucesso !", vbInformation
    
End If
End Sub

Private Sub PesquisaLivros()
Dim Condicao As String

If Opt(0).value = True Then
    Condicao = "SELECT ENT.Rg, ENT.Nome,  ENT.CPF, PED.CFO, PED.Operacao, Sum(ITP.Valor) AS Valor, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, PED.DtSaida FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido GROUP BY PED.Operacao, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, ENT.CPF, PED.CFO, ENT.Rg, ENT.Nome, PED.DtSaida HAVING (PED.Operacao=0 Or PED.Operacao=9 Or PED.Operacao=10) AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#)"
Else
    Condicao = "SELECT ENT.Rg, ENT.Nome,  ENT.CPF, PED.CFO, PED.Operacao, Sum(ITP.Valor) AS Valor, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, PED.DtSaida FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido GROUP BY PED.Operacao, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, ENT.CPF, PED.CFO, ENT.Rg, ENT.Nome,PED.DtSaida HAVING (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3 Or (PED.Operacao)=6 Or (PED.Operacao)=7 Or (PED.Operacao)=8 Or (PED.Operacao)=11) AND ((([PED].[DtFat]))>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (([PED].[DtFat]))<= #" & Format(MebDtF, "mm/dd/yyyy") & "#))"
End If

Data.RecordSource = Condicao
Data.Refresh

If Data.Recordset.RecordCount <= 0 Then
    MsgBox "Não existe dados nesse intervalo", vbInformation, App.Title
    Exit Sub
End If

End Sub

Private Sub GeraLivros()
Dim Tipo As String
Dim CgcTmp As String
Dim NomeArq As String
Dim x, y, Z As Currency
Dim PedAnt, Contador As Currency

If Data.Recordset.RecordCount <= 0 Then
    Exit Sub
End If


If Opt(0).value = True Then
    NomeArq = App.Path & "\" & "Entrada.txt"
Else
    NomeArq = App.Path & "\" & "Saida.txt"
End If

If MsgBox("Confirma geração do arquivo texto ?", vbYesNo, App.Title) = vbYes Then
    
    Open NomeArq For Output As #1
    
    With Data.Recordset
        .MoveFirst
        
        If Opt(0).value = True Then
        
            While Not .Eof
                                '  psstring = 0 - espacos a direita | 2 = zeros a esquerda
                Texto = Space(768)
                psString 0, "", 1, 4
                If (Data.Recordset!CGC) <> "" Then
                    psString 2, LimpaString(Data.Recordset!CGC), 5, 14
                Else
                    psString 2, LimpaString(Data.Recordset!Cpf), 5, 14
                End If
                psString 1, LimpaString(Format(Data.Recordset!DataEmiss, "dd/mm/yy")), 19, 6
                psString 1, LimpaString(Format(Data.Recordset!DtFat, "dd/mm/yy")), 25, 6
                psString 2, Data.Recordset!NroNt, 31, 6
                psString 0, "NF", 37, 3
                psString 0, "U", 40, 3
                
                If PedAnt <> Data.Recordset!CodPed Then
                    PedAnt = Data.Recordset!CodPed
                    Contador = 1
                ElseIf PedAnt = Data.Recordset!CodPed Then
                    Contador = Contador + 1
                End If
                
                psString 1, CStr(Contador), 43, 2 'item de desdobramento (fazer rotina)
                
                psString 0, "", 44, 5
                psString 0, "", 49, 3
                psString 3, Data.Recordset!Valor, 52, 14
                psString 3, Data.Recordset!Valor, 66, 14
                psString 1, Data.Recordset!Icms, 80, 14
                psString 0, "", 94, 14
                psString 0, "", 108, 14
                psString 0, "", 122, 14
                psString 2, Data.Recordset!Icm, 136, 5
                psString 3, Data.Recordset!Valor, 141, 14
                psString 3, Data.Recordset("Ped.Ipi"), 155, 14
                psString 3, Data.Recordset("Ped.Ipi") / 2, 169, 14
                psString 0, "", 183, 14
                psString 0, "", 197, 14
                psString 0, "", 211, 14
                psString 0, "", 225, 1
                psString 0, "", 226, 2
                psString 0, "", 228, 3
                psString 0, "", 231, 12
                psString 0, "", 243, 12
                psString 0, "", 255, 12
                psString 0, "", 267, 12
                psString 0, "", 279, 100
                
                Set tabDuplicatas = Banco.OpenRecordset("SELECT DUP.*, DUP.Pedido from Dup WHERE DUP.Pedido=" & Data.Recordset!CodPed)
                x = 359
                y = 365
                Z = 0
                          
                Do While Not tabDuplicatas.Eof
                               
                    x = x + 20
                    y = y + 20
                    Z = Z + 1
                
                    psString 1, LimpaString(Format(tabDuplicatas!DtVenc, "dd/mm/yy")), x, 6
                    psString 3, tabDuplicatas!Valor, y, 14
                                                    
                    If Z >= 6 Then Exit Do
                
                    tabDuplicatas.MoveNext
                Loop
               
                psString 1, LimpaString(Data.Recordset!InscrEst), 499, 18
                psString 1, Data.Recordset!Uf, 517, 2
                psString 1, LimpaString(Data.Recordset!Cfo), 624, 4
               
               .MoveNext
               
                'Texto = Texto & Chr(13)
                Print #1, Texto
                Texto = ""
               
            Wend
        Else
            
            While Not .Eof
        
                If PedAnt <> Data.Recordset!CodPed Then
                    PedAnt = Data.Recordset!CodPed
                    Contador = 1
                ElseIf PedAnt = Data.Recordset!CodPed Then
                    Contador = Contador + 1
                End If
                
                Texto = Space(874)
                psString 3, Data.Recordset!NroNt, 7, 6
                psString 3, Data.Recordset!NroNt, 13, 6
                psString 3, CStr(Contador), 19, 1
                psString 1, LimpaString(Format(Data.Recordset!DtSaida, "dd/mm/yy")), 20, 6
                psString 3, Data.Recordset!Valor, 34, 14
                psString 3, Data.Recordset!Valor, 48, 14
                psString 3, Data.Recordset!Valor, 62, 14
                psString 2, Data.Recordset("Icms"), 76, 14
                psString 3, Data.Recordset!Valor, 118, 14
                psString 3, Data.Recordset("Ped.Ipi"), 132, 14
                psString 3, Data.Recordset!Icm, 188, 5
                psString 1, LimpaString(Data.Recordset!CGC), 254, 14
                psString 1, Data.Recordset!Uf, 273, 2
                psString 3, Data.Recordset!NroNt, 533, 6
                
                
                Set tabDuplicatas = Banco.OpenRecordset("SELECT DUP.*, DUP.Pedido from Dup WHERE DUP.Pedido=" & Data.Recordset!CodPed)
                x = 550
                y = 564
                Z = 0
                          
                Do While Not tabDuplicatas.Eof
                               
                    x = x + 20
                    y = y + 20
                    Z = Z + 1
                
                    psString 1, LimpaString(Format(tabDuplicatas!DtVenc, "dd/mm/yy")), x, 6
                    psString 3, tabDuplicatas!Valor, y, 14
                                                    
                    If Z >= 6 Then Exit Do
                
                    tabDuplicatas.MoveNext
                Loop


                Print #1, Texto
                Texto = ""
                .MoveNext
               
            Wend
                        
        End If
    End With
    
    Close #1
    
    MsgBox "Arquivo gerado com sucesso !", vbInformation
    
End If
End Sub


Private Sub GeraContimatic()
Dim Tipo As String
Dim CgcTmp As String
Dim NomeArq As String
Dim x, y, Z As Currency
Dim PedAnt, Contador As Currency

If Data.Recordset.RecordCount <= 0 Then
    Exit Sub
End If


If Opt(0).value = True Then
    NomeArq = App.Path & "\" & "Entrada.txt"
Else
    NomeArq = App.Path & "\" & "Saida.txt"
End If

If MsgBox("Confirma geração do arquivo texto ?", vbYesNo, App.Title) = vbYes Then
    
    Open NomeArq For Output As #1
    
    With Data.Recordset
        .MoveFirst
        
        If Opt(0).value = True Then
        
            While Not .Eof
                                '  psstring = 0 - espacos a direita | 2 = zeros a esquerda
                Texto = Space(768)
                psString 0, "", 1, 4
                If (Data.Recordset!CGC) <> "" Then
                    psString 2, LimpaString(Data.Recordset!CGC), 5, 14
                Else
                    psString 2, LimpaString(Data.Recordset!Cpf), 5, 14
                End If
                psString 1, LimpaString(Format(Data.Recordset!DataEmiss, "dd/mm/yy")), 19, 6
                psString 1, LimpaString(Format(Data.Recordset!DtFat, "dd/mm/yy")), 25, 6
                psString 2, Data.Recordset!NroNt, 31, 6
                psString 0, "NF", 37, 3
                psString 0, "U", 40, 3
                
                If PedAnt <> Data.Recordset!CodPed Then
                    PedAnt = Data.Recordset!CodPed
                    Contador = 1
                ElseIf PedAnt = Data.Recordset!CodPed Then
                    Contador = Contador + 1
                End If
                
                psString 1, CStr(Contador), 43, 2 'item de desdobramento (fazer rotina)
                
                psString 0, "", 44, 5
                psString 0, "", 49, 3
                psString 3, Data.Recordset!Valor, 52, 14
                psString 3, Data.Recordset!Valor, 66, 14
                psString 1, Data.Recordset!Icms, 80, 14
                psString 0, "", 94, 14
                psString 0, "", 108, 14
                psString 0, "", 122, 14
                psString 2, Data.Recordset!Icm, 136, 5
                psString 3, Data.Recordset!Valor, 141, 14
                psString 3, Data.Recordset("Ped.Ipi"), 155, 14
                psString 3, Data.Recordset("Ped.Ipi") / 2, 169, 14
                psString 0, "", 183, 14
                psString 0, "", 197, 14
                psString 0, "", 211, 14
                psString 0, "", 225, 1
                psString 0, "", 226, 2
                psString 0, "", 228, 3
                psString 0, "", 231, 12
                psString 0, "", 243, 12
                psString 0, "", 255, 12
                psString 0, "", 267, 12
                psString 0, "", 279, 100
                
                Set tabDuplicatas = Banco.OpenRecordset("SELECT DUP.*, DUP.Pedido from Dup WHERE DUP.Pedido=" & Data.Recordset!CodPed)
                x = 359
                y = 365
                Z = 0
                          
                Do While Not tabDuplicatas.Eof
                               
                    x = x + 20
                    y = y + 20
                    Z = Z + 1
                
                    psString 1, LimpaString(Format(tabDuplicatas!DtVenc, "dd/mm/yy")), x, 6
                    psString 3, tabDuplicatas!Valor, y, 14
                                                    
                    If Z >= 6 Then Exit Do
                
                    tabDuplicatas.MoveNext
                Loop
               
                psString 1, LimpaString(Data.Recordset!InscrEst), 499, 18
                psString 1, Data.Recordset!Uf, 517, 2
                psString 1, LimpaString(Data.Recordset!Cfo), 624, 4
               
               .MoveNext
               
                'Texto = Texto & Chr(13)
                Print #1, Texto
                Texto = ""
               
            Wend
        Else
            
            .MoveFirst
            
                While Not .Eof
                                    '  psstring = 0 - espacos a direita | 2 = zeros a esquerda
                    Texto = Space(768)
                    psString 1, "R1", 1, 2
                    psString 1, psCompString(1, LimpaString(Day(Data.Recordset!DtFat)), 2) & psCompString(1, LimpaString(Month(Data.Recordset!DtFat)), 2), 3, 4
                    psString 1, psCompString(1, LimpaString(Day(Data.Recordset!DtSaida)), 2) & psCompString(1, LimpaString(Month(Data.Recordset!DtSaida)), 2), 7, 4
                    psString 1, "00", 11, 2
                    psString 1, "NF", 13, 2
                    psString 1, "001", 16, 2
                    psString 2, psCompString(1, Data.Recordset!NroNt, 6), 19, 6
                    psString 1, "000000", 25, 6
                    psString 1, LimpaString(Data.Recordset!Cfo), 31, 5
                    psString 3, Data.Recordset!Valor, 36, 12
                    
                    psString 3, Data.Recordset!Valor, 48, 12  '1
                    psString 2, "000000", 60, 6 '1
                    psString 1, "000000000000", 66, 12 '1
                    psString 1, "000000000000", 78, 12 '1
                    psString 1, "000000000000", 90, 12 '1
                    
                    psString 3, Data.Recordset!Valor, 102, 12  '2
                    psString 2, "000000", 114, 6 '2
                    psString 1, "000000000000", 120, 12 '2
                    psString 1, "000000000000", 132, 12 '2
                    psString 1, "000000000000", 144, 12 '2
                    
                    psString 3, Data.Recordset!Valor, 156, 12  '3
                    psString 2, "000000", 168, 6 '3
                    psString 1, "000000000000", 174, 12 '3
                    psString 1, "000000000000", 186, 12 '3
                    psString 1, "000000000000", 198, 12 '3
                    
                    psString 3, Data.Recordset!Valor, 210, 12  '4
                    psString 2, "000000", 222, 6 '4
                    psString 1, "000000000000", 228, 12 '4
                    psString 1, "000000000000", 240, 12 '4
                    psString 1, "000000000000", 252, 12 '4
                    
                    psString 3, Data.Recordset!Valor, 264, 12  '5
                    psString 2, "000000", 276, 6 '5
                    psString 1, "000000000000", 282, 12 '5
                    psString 1, "000000000000", 294, 12 '5
                    psString 1, "000000000000", 306, 12 '5
                    
                    psString 1, "000000000000", 318, 12 ' ipi
                    psString 1, "000000000000", 330, 12
                    psString 1, "000000000000", 342, 12
                    psString 1, "000000000000", 354, 12
                    psString 1, "000000000000", 366, 12
                    psString 1, "000000000000", 378, 12
                    psString 1, "000000000000", 390, 12
                    psString 1, "000000000000", 402, 12
                    psString 1, "000000000000", 414, 12
                    psString 1, "0", 438, 1
                    psString 1, "00", 439, 2
                    psString 1, "00", 441, 2
                    psString 1, Space(14), 443, 14
                    If (Data.Recordset!CGC) <> "" Then
                        psString 1, LimpaString(Data.Recordset!CGC), 457, 14
                    Else
                        psString 1, LimpaString(Data.Recordset!Cpf), 457, 14
                    End If
                    If (Data.Recordset!InscrEst) <> "" Then
                        psString 1, LimpaString(Data.Recordset!InscrEst), 471, 16
                    Else
                        psString 1, LimpaString(Data.Recordset!Rg), 471, 16
                    End If
                    psString 1, LimpaString(Left(Data.Recordset!Nome, 35)), 487, 35
                    psString 1, Space(18), 522, 18
                    psString 1, LimpaString(Data.Recordset!Uf), 540, 2
                    psString 1, "0000", 542, 4
                    psString 1, Space(50), 546, 50

                   .MoveNext
                   
                    'Texto = Texto & Chr(13)
                    Print #1, Texto
                    Texto = ""
                   
                Wend
                        
        End If
    End With
    
    Close #1
    
    MsgBox "Arquivo gerado com sucesso !", vbInformation
    
End If
End Sub


