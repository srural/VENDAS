VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormSintegra 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14445
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -495
      ScaleHeight     =   345
      ScaleWidth      =   14925
      TabIndex        =   23
      Top             =   7830
      Width           =   14955
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14475
      TabIndex        =   18
      Top             =   0
      Width           =   14475
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos Sintegra / R.E.D.F. / Nota Fiscal Paulista / S.P.E.D."
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
         TabIndex        =   19
         Top             =   45
         Width           =   11445
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1920
      Left            =   9180
      TabIndex        =   11
      Top             =   585
      Width           =   5115
      Begin VB.Label LblInscr 
         Caption         =   "Insc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1035
         TabIndex        =   22
         Top             =   1260
         Width           =   3900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Est"
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.N.P.J."
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label LblSerie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1290
         TabIndex        =   15
         Top             =   1605
         Width           =   2190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo/Serie:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1635
         Width           =   1005
      End
      Begin VB.Label LblCnpj 
         Caption         =   "Cnpj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1020
         TabIndex        =   13
         Top             =   960
         Width           =   3900
      End
      Begin VB.Label LblEmpresa 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4710
      End
   End
   Begin VB.CommandButton BtOK 
      Caption         =   "&Pesquisa"
      Height          =   450
      Left            =   5895
      TabIndex        =   10
      Top             =   990
      Width           =   1545
   End
   Begin VB.CommandButton BtGera 
      Caption         =   "&Gera Arquivo"
      Height          =   450
      Left            =   5940
      TabIndex        =   9
      Top             =   1755
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Arquivo"
      Height          =   1710
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2310
      Begin VB.OptionButton opTipo 
         Caption         =   "S.P.E.D."
         Height          =   420
         Index           =   2
         Left            =   165
         TabIndex        =   20
         Top             =   855
         Width           =   1875
      End
      Begin VB.CheckBox ChkConsistir 
         Caption         =   "Consistir Arquivo"
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   1305
         Width           =   1815
      End
      Begin VB.OptionButton opTipo 
         Caption         =   "R.E.D.F."
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton opTipo 
         Caption         =   "Nota Fiscal Paulista"
         Enabled         =   0   'False
         Height          =   420
         Index           =   1
         Left            =   165
         TabIndex        =   7
         Top             =   525
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   1710
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1920
      Begin MSMask.MaskEdBox MebDtI 
         Height          =   375
         Left            =   210
         TabIndex        =   2
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebDtF 
         Height          =   375
         Left            =   210
         TabIndex        =   3
         Top             =   1170
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   285
         Width           =   795
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
      Left            =   4770
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2085
   End
   Begin MSDBGrid.DBGrid DbGrid 
      Bindings        =   "FormSintegra.frx":0000
      Height          =   5025
      Left            =   120
      OleObjectBlob   =   "FormSintegra.frx":0013
      TabIndex        =   0
      Top             =   2610
      Width           =   14175
   End
End
Attribute VB_Name = "FormSintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Texto As String
Dim tabDuplicatas As Recordset
Public VarIbge As String
Public VarCnpj As String
Public Reg0150 As Recordset
Public Reg0190 As Recordset
Public Reg0200 As Recordset
Public Reg0400 As Recordset
Public RegC100 As Recordset
Public RegC170 As Recordset
Public Linhas0  As Double
Public LinhasC  As Double

Private Sub BtGera_Click()
    If opTipo(0).value = True Then
        PesquisaLivros
        GeraLivros
    ElseIf opTipo(1).value = True Then
'        GeraSerasa
    Else
        PesquisaLivros
        GeraLivros

    End If
End Sub

Private Sub BtOK_Click()
    If opTipo(0).value = True Then
        PesquisaLivros
    ElseIf opTipo(1).value = True Then
'        PesquisaSerasa
    Else
        PesquisaLivros
    End If
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

MebDtI.Text = Date - 30
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
LblEmpresa = NomeEmpresa
LblCnpj = Cnpj
If VarSerieNota = "0" Then
'        LblSerie = "Série D"
'    ElseIf varserinota = "2" Then
'        LblSerie = "Série D única"
'    Else
        LblSerie = "Série única"
Else
    LblSerie = VarModeloNota & "/" & VarSerieNota
End If

If Dir(App.Path & "\nfe.ini") <> "" Then

    If LerINI("EMITENTE", "CNPJ", App.Path & "\nfe.ini") <> "" Then
        VarCnpj = LerINI("EMITENTE", "CNPJ", App.Path & "\nfe.ini")
    End If


    If LerINI("EMITENTE", "INSCEST", App.Path & "\nfe.ini") <> "" Then
                LblInscr = LerINI("EMITENTE", "INSCEST", App.Path & "\nfe.ini")
    End If

    If LerINI("EMITENTE", "CODIGOIBGE", App.Path & "\nfe.ini") <> "" Then
        VarIbge = LerINI("EMITENTE", "CODIGOIBGE", App.Path & "\nfe.ini")
    End If
End If
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



Private Sub PesquisaLivros()
Dim Condicao As String

If Not IsDate(MebDtI) Or Not IsDate(MebDtF) Then
    MsgBox "Data Inválida", vbInformation, App.Title
    Exit Sub
End If

If opTipo(0).value = True Then
'    Condicao = "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, Ent.Classificacao, Ped.CodPed FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, Ent.Classificacao, Ped.CodPed  HAVING PED.NroNt <> 0  AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#) ORDER BY Ped.DtFat "
    Condicao = "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, Ped.Operacao, PED.CodPed, PED.Cfo FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.CodPed,  Ped.Operacao, PED.Cfo HAVING PED.NroNt <> 0  AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#) ORDER BY Ped.DtFat "

ElseIf opTipo(1) Then
'    Condicao = "SELECT ENT.CPF, PED.CFO, PED.Operacao, Sum(ITP.Valor) AS Valor, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, PED.DtSaida FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido GROUP BY PED.Operacao, ENT.CGC, PED.DataEmiss, PED.DtFat, PED.NroNt, PED.Icms, ENT.InscrEst, PED.CodPed, PRD.Icm, PRD.Ipi, PED.Ipi, ENT.Uf, ENT.CPF, PED.CFO, PED.DtSaida HAVING (((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3 Or (PED.Operacao)=6 Or (PED.Operacao)=7 Or (PED.Operacao)=8 Or (PED.Operacao)=11) AND ((([PED].[DtFat]))>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (([PED].[DtFat]))<= #" & Format(MebDtF, "mm/dd/yyyy") & "#))"
        Condicao = "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, Ped.Operacao, PED.CodPed, PED.Cfo FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.CodPed,  Ped.Operacao, PED.Cfo HAVING PED.NroNt <> 0  AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#) ORDER BY Ped.DtFat "
    Else
    
        Condicao = "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.Operacao, PED.CodPed, PED.Cfo, NOTA.Modelo  FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.Operacao, PED.CodPed, PED.Cfo, NOTA.Modelo HAVING PED.NroNt <> 0  AND PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "# AND Nota.Modelo  = '55' ORDER BY PED.DtFat "

        
        
  '      "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, Ped.Operacao, PED.CodPed, PED.Cfo FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.CodPed,  Ped.Operacao, PED.Cfo HAVING PED.NroNt <> 0  AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#) ORDER BY Ped.DtFat "
End If

If ChkConsistir.value = 1 Then
    Condicao = "SELECT PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, Ped.Operacao, PED.CodPed, PED.Cfo FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade GROUP BY PED.NroNt, PED.DtFat, PED.DtSaida, ENT.CPF, ENT.CGC, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, PED.Total, PED.Desconto, PED.ValorFrete, PED.TipoFrete, ENT.Classificacao, PED.CodPed,  Ped.Operacao, PED.Cfo HAVING PED.NroNt <> 0  AND (PED.DtFat >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And PED.DtFat <= #" & Format(MebDtF, "mm/dd/yyyy") & "#) ORDER BY Ped.DtFat "
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
Dim VarReg As String
Dim DtIni As Date
Dim DtFim  As Date
Dim RegItens As Recordset
Dim RegNota As Recordset
Dim Criterio As String
Dim Cont As Integer
Dim Cont20 As Integer
Dim Cont30 As Integer
Dim Cont40 As Integer
Dim Cont50 As Integer
Dim Cont60 As Integer
Dim PathRedf As String
Dim VarTot As Currency

Dim VarTotal As Currency
Dim Linha As String
Dim VarTipo As Byte

VarTotal = 0
VarReg = ""

VarTot = 0

If Data.Recordset.RecordCount <= 0 Then
    Exit Sub
End If

Set RegItens = Banco.OpenRecordset("SELECT ITP.*, ITP.Pedido, ITP.CodItp, PRD.Descrição_Produto, PRD.ClasseFiscal, PRD.Embalagem FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto ORDER BY ITP.Pedido, ITP.CodItp")

Set RegNota = Banco.OpenRecordset("Select * From  Nota Order by Codped")


If LerINI("Nota", "CaminhoRedf", App.Path & "\config.ini") <> "" Then
    PathRedf = LerINI("Nota", "CaminhoRedf", App.Path & "\config.ini")
Else
    PathRedf = App.Path
End If

If MsgBox("Confirma geração do arquivo texto ?", vbYesNo, App.Title) = vbYes Then
    
    If opTipo(0).value = True Then
    
        With Data.Recordset
            If .RecordCount > 0 Then
                RegItens.Requery
                RegItens.MoveFirst
                
                .MoveLast
                DtFim = .Fields("DtFat")
                .MoveFirst
                DtIni = .Fields("DtFat")
                    
                NomeArq = PathRedf & "\" & "REDF_" & Format(DtIni, "DDMMYYYY") & "_" & Format(DtFim, "DDMMYYYY") & ".txt"
                    
                Open NomeArq For Output As #1
                
                VarReg = "10" & "|" & "1,00" & "|" & LimpaString(Cnpj) & "|" & Format(DtIni, "DD/MM/YYYY") & "|" & Format(DtFim, "DD/MM/YYYY")
                
                Print #1, UTF8_Encode(VarReg)
                
                Cont20 = 0
                Cont30 = 0
                Cont40 = 0
                Cont50 = 0
                Cont60 = 0
                
                While Not .Eof
                
                    Criterio = "CodPed = " & .Fields("CodPed")
                    RegNota.FindFirst Criterio
                    
                    If Not RegNota.NoMatch And RegNota!Modelo <> "55" Then
                    
                        Select Case .Fields("Operacao")
                                Case 0
                                    VarTipo = 0
                                Case 1
                                    VarTipo = 1
                                Case 2
                                    VarTipo = 1
                                Case 3
                                    VarTipo = 1
                                Case 4
                                    VarTipo = 1
                                Case 5
                                    VarTipo = 1
                                Case 6
                                    VarTipo = 1
                                Case 7
                                    VarTipo = 1
                                Case 8
                                    VarTipo = 1
                                Case 9
                                    VarTipo = 0
                                Case 10
                                    VarTipo = 0
                                Case 11
                                    VarTipo = 1
                                Case 12
                                    VarTipo = 0
                                Case 13
                                    VarTipo = 1
                        End Select
                                        
        '               psstring = 0 - espacos a direita | 2 = zeros a esquerda
        '               Texto = Space(768)
                        VarReg = "20" & "|" & "I" & "|" & "|" & RegNota("NaturezaOp") & "|" & VarSerieNota & "|" & .Fields("NroNt") & "|"
                        VarReg = VarReg & Format(.Fields("DtFat"), "DD/MM/YYYY HH:MM:SS") & "|" & Format(.Fields("DtSaida"), "DD/MM/YYYY HH:MM:SS") & "|"
                        VarReg = VarReg & VarTipo & "|" & Left(LIMPACampo2(.Fields("Cfo")), 4) & "|" & "|" & "|" & Trim(LimpaString(RegNota!Documento)) & "|"
                        
                        VarReg = VarReg & IIf(RegNota!RazaoSocial <> "", Left(RegNota!RazaoSocial, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Logradouro <> "", Left(RegNota!Logradouro, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Nro <> "", Left(RegNota!Nro, 60), "NAO IDENTIFICADO") & "|" & "|"
                        VarReg = VarReg & IIf(RegNota!Bairro <> "", Left(RegNota!Bairro, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Cidade <> "", Left(RegNota!Cidade, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & RegNota!UfNota & "|" & RegNota!Cep & "|" & "|" & "|" & RegNota!IEDestinatario
                        Cont20 = Cont20 + 1
                        Print #1, UTF8_Encode(VarReg)
                                        
                        Criterio = "Itp.Pedido = " & .Fields("CodPed")
                        RegItens.FindFirst Criterio
                        If Not RegItens.NoMatch Then
                            While Not RegItens.Eof
                                If RegItens("Itp.Pedido") = .Fields("CodPed") Then
                                    VarReg = "30" & "|" & RegItens!Produto & "|" & Left(RegItens!Descrição_Produto, 120) & "||" & IIf(RegItens!Embalagem = "", "UN", Left(RegItens!Embalagem, 6)) & "|"
                                    VarReg = VarReg & Format(RegItens!Qtd, "#######0.0000") & "|"
                                    VarReg = VarReg & Format(RegItens!ValorUnit, "###########0.0000") & "|"
                                    VarReg = VarReg & Format(RegItens!Valor, "############0.00") & "|" & "000" & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!AliqIcms), 0, RegItens!AliqIcms), "##0.00") & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!AliqIPI), 0, RegItens!AliqIPI), "##0.00") & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!ValorIPI), 0, RegItens!ValorIPI), "############0.00")
                                    Print #1, UTF8_Encode(VarReg)
                                    VarReg = ""
                                    Cont30 = Cont30 + 1
                                    VarTot = VarTot + RegItens!Valor
                                 End If
                                 RegItens.MoveNext
                            Wend
                        End If
                        
'                        VarReg = "40" & "|" & Format(RegNota!BaseIcms, "############0.00") & "|"
                        VarReg = "40" & "|" & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIcms, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!BaseIcmsSub, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIcmsSub, "############0.00") & "|"
'                        VarReg = VarReg & Format(RegNota!ValorTotalPrd, "############0.00") & "|"
                        VarReg = VarReg & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorFrete, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorSeguro, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorDesconto, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotIpi, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotOutras, "############0.00") & "|"
'                        VarReg = VarReg & Format(RegNota!TotNota, "############0.00") & "|"
                        VarReg = VarReg & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotServicos, "##0.00") & "|"
                        VarReg = VarReg & Format(RegNota!AliqIss, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIss, "############0.00")
                        Print #1, UTF8_Encode(VarReg)
                        Cont40 = Cont40 + 1
                        VarTot = 0
                        
                        
                        VarReg = "50" & "|" & RegNota!TipoFrete - 1 & "|" & LimpaString(SeCampoNulo(RegNota, "CnpjTransp")) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "RazaoTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "IETransp"), 14) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "EndTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "CidadeTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "UfTransp"), 2) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "PlacaVeiculo"), 8) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Uf"), 2) & "|"
                        VarReg = VarReg & Format(RegNota!QtdVolume, "##############0") & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Especie"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Marca"), 60) & "|" & "|"
                        VarReg = VarReg & Format(RegNota!PesoLiquido, "###########0.000") & "|"
                        VarReg = VarReg & Format(RegNota!PesoBruto, "###########0.000")
                        Print #1, UTF8_Encode(VarReg)
                        Cont50 = Cont50 + 1
                    
                        VarReg = "60" & "|" & RegNota!DadosFatura & "|" & RegNota!InformFisco & "|" & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Cont60 = Cont60 + 1
                    End If
                    .MoveNext
                Wend
                VarReg = "90" & "|" & Format(Cont20, "00000") & "|" & Format(Cont30, "00000") & "|" & Format(Cont40, "00000") & "|" & Format(Cont50, "00000") & "|" & Format(Cont60, "00000") '& "|"
                Print #1, UTF8_Encode(VarReg)
            Else
                MsgBox "Arquivo sem movimento !", vbInformation
            End If
            
        End With
    ElseIf opTipo(1) Then
    
        With Data.Recordset
            If .RecordCount > 0 Then
                RegItens.Requery
                RegItens.MoveFirst
                
                .MoveLast
                DtFim = .Fields("DtFat")
                .MoveFirst
                DtIni = .Fields("DtFat")
                
                NomeArq = PathRedf & "\" & "REDF_" & Format(DtIni, "DDMMYYYY") & "_" & Format(DtFim, "DDMMYYYY") & ".txt"
                Open NomeArq For Output As #1
                                        
                Linha = UTF8_Encode("10" & "|" & "1,00" & "|" & LimpaString(Cnpj) & "|" & Format(DtIni, "DD/MM/YYYY") & "|" & Format(DtFim, "DD/MM/YYYY"))
                
                Print #1, Linha & vbCr
                
                Cont21 = 0
                Cont22 = 0
                
                While Not .Eof
                                    '  psstring = 0 - espacos a direita | 2 = zeros a esquerda
    '                Texto = Space(768)
                    VarReg = "20" & "|" & "I" & "|" & "|" & .Fields("Descricao") & "|" & "0" & "|" & .Fields("NroNt") & "|"
                    VarReg = VarReg & Format(.Fields("DtFat"), "DD/MM/YYYY HH:MM:SS") & "|" & Format(.Fields("DtSaida"), "DD/MM/YYYY HH:MM:SS") & "|"
                    VarReg = VarReg & "1|" & Left(LIMPACampo2(.Fields("Cfo")), 4) & "|" & "|" & "|"
                    If .Fields("Classificacao") = 1 Then
                        If IsNull(.Fields("Cpf")) Then
                            VarReg = VarReg & "|"
                        Else
                            VarReg = VarReg & LimpaString(.Fields("Cpf")) & "|"
                        End If
                    ElseIf .Fields("Classificacao") = 2 Then
                            If IsNull(.Fields("Cgc")) Then
                                VarReg = VarReg & "|"
                            Else
                                VarReg = VarReg & LimpaString(.Fields("Cgc")) & "|"
                            End If
                        Else
                            VarReg = VarReg & "|"
                    End If
                    VarReg = VarReg & Left(.Fields("Nome"), 60) & "|" & Left(.Fields("Endereco"), 60) & "|" & Left(.Fields("Nro"), 60) & "|"
                    VarReg = VarReg & Left(.Fields("Complemento"), 60) & "|" & Left(.Fields("Bairro"), 60) & "|"
                    VarReg = VarReg & Left(.Fields("Cidade"), 60) & "|" & Left(.Fields("Uf"), 2) & "|" & Left(LimpaString(.Fields("Cep")), 8) & "|"
                    VarReg = VarReg & LimpaString(Left(.Fields("Fone"), 10)) & "|"
                    VarReg = VarReg & Replace(Format(.Fields("Total") + .Fields("Desconto"), "##############0,00"), ",", ".") & "|"
                    VarReg = VarReg & Replace(Format(.Fields("Desconto"), "##############0,00"), ",", ".") & "|"
                    VarReg = VarReg & Replace(Format(.Fields("ValorFrete"), "##############0,00"), ",", ".") & "|" & "0,00" & "|" & "0,00" & "|" & "|"
                    VarReg = VarReg & Replace(Format(.Fields("Total") + .Fields("ValorFrete"), "##############0,00"), ",", ".") & "|"
                    VarReg = VarReg & "|" & "|" & "1" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "1" & "|" & "|" & "|" & "|"
                    VarTotal = VarTotal + (.Fields("Total") + .Fields("ValorFrete"))
                    
                    Print #1, UTF8_Encode(VarReg) & vbCr
                    VarReg = ""
                    
                    Criterio = "Itp.Pedido = " & .Fields("CodPed")
                    RegItens.FindFirst Criterio
                    Cont = 0
                    If Not RegItens.NoMatch Then
        
                        While RegItens("Itp.Pedido") = .Fields("CodPed") And Not RegItens.Eof
                            VarReg = "21" & "|" & Cont & "|" & RegItens!Produto & "|" & Left(RegItens!Descrição_Produto, 120) & "|" & Left(RegItens!Embalagem, 6) & "|"
                            VarReg = VarReg & Replace(Format(RegItens!Qtd, "##########0,000"), ",", ".") & "|"
                            VarReg = VarReg & Replace(Format(RegItens!ValorUnit, "##############0,000"), ",", ".") & "|"
                            VarReg = VarReg & Replace(Format(RegItens!VlrVdaDia, "##############0,000"), ",", ".")
                            Cont = Cont + 1
                            RegItens.MoveNext
                            Print #1, UTF8_Encode(VarReg) & vbCr
                            VarReg = ""
                            Cont21 = Cont21 + 1
                        Wend
                    End If
                    .MoveNext
                    Cont20 = Cont20 + 1
                Wend
                VarReg = "90" & "|" & Cont20 & "|" & Cont21 & "|" & "0" & "|" & Replace(Format(VarTotal, "##############0.000"), ",", ".") & vbCr
                Print #1, UTF8_Encode(VarReg) & vbCr

    '
            Else
                MsgBox "Arquivo sem movimento !", vbInformation
            End If
            
        End With
            
    Else

        With Data.Recordset
            If .RecordCount > 0 Then
                RegItens.Requery
                RegItens.MoveFirst
                
                .MoveLast
                DtFim = .Fields("DtFat")
                .MoveFirst
                DtIni = .Fields("DtFat")
                    
                NomeArq = PathRedf & "\" & "SPED_" & Format(DtIni, "DDMMYYYY") & "_" & Format(DtFim, "DDMMYYYY") & ".txt"
                    
                Open NomeArq For Output As #1
                
                VarReg = "|0140" & "|" & "1" & "|" & Trim(LblEmpresa) & "|SP|" & LimpaString(VarCnpj) & "|" & Trim(LblInscr) & "|" & Trim(VarIbge) & "|||"
                
                Print #1, UTF8_Encode(VarReg)
                
                Linhas0 = 1
                
                Cont20 = 0
                Cont30 = 0
                Cont40 = 0
                Cont50 = 0
                Cont60 = 0
                
                '  ************************  0150
                
                Set Reg0150 = Banco.OpenRecordset("SELECT ENT.CodEntidade, ENT.Nome, ENT.CGC, NOTA.Modelo, ENT.CPF, ENT.InscrEst, ENT.Matricula, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Cidade FROM NOTA INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON NOTA.CodPed = PED.CodPed WHERE (((PED.DtFat)>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (PED.DtFat)<=#" & Format(MebDtF, "mm/dd/yyyy") & "#)) GROUP BY ENT.CodEntidade, ENT.Nome, ENT.CGC, NOTA.Modelo, ENT.CPF, ENT.InscrEst, ENT.Matricula, ENT.Endereco, ENT.Nro, ENT.Complemento, ENT.Cidade HAVING (((NOTA.Modelo)='55'))  ORDER BY ENT.CodEntidade")
                Reg0150.Requery
                If Reg0150.RecordCount > 0 Then
                    Reg0150.MoveFirst
                    While Not Reg0150.Eof
                        VarReg = "|0150" & "|" & Reg0150!CodEntidade & "|" & Trim(Reg0150!Nome) & "|1058|" & Trim(LimpaString(Reg0150!CGC)) & "|" & Trim(LimpaString(Reg0150!Cpf)) & "|"
                        VarReg = VarReg & Trim(Reg0150!matricula) & "|" & Trim(LimpaString(Reg0150!InscrEst)) & "||"
                        VarReg = VarReg & Trim(Reg0150!Endereco) & "|" & Reg0150!Nro & "|" & Trim(Reg0150!Complemento) & "|" & Trim(Reg0150!Bairro) & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Linhas0 = Linhas0 + 1
                        Reg0150.MoveNext
                    Wend
                End If
                
                '  ************************  0190
                
                Set Reg0190 = Banco.OpenRecordset("SELECT NOTA.Modelo, PRD.Embalagem FROM ((PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido) INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed) INNER JOIN PRD ON ITP.Produto = PRD.CodPrd  WHERE (((PED.DtFat)>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (PED.DtFat)<=#" & Format(MebDtF, "mm/dd/yyyy") & " #)) GROUP BY NOTA.Modelo, PRD.Embalagem HAVING (((NOTA.Modelo)='55')) ")
                Reg0190.Requery
                If Reg0190.RecordCount > 0 Then
                    Reg0190.MoveFirst
                    While Not Reg0190.Eof
                        VarReg = "|0190" & "|" & Trim(Reg0190!Embalagem) & "|" & Trim(Reg0190!Embalagem) & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Linhas0 = Linhas0 + 1
                        Reg0190.MoveNext
                    Wend
                End If
                
                '  ************************** 0200
                
                Set Reg0200 = Banco.OpenRecordset("SELECT NOTA.Modelo, PRD.Embalagem, PRD.CodPrd, PRD.Descrição_Produto, BAR.BarBarra, PRD.Classe, PRD.Icm FROM ((PED INNER JOIN NOTA ON PED.CodPed = NOTA.CodPed) INNER JOIN (ITP INNER JOIN PRD ON ITP.Produto = PRD.CodPrd) ON PED.CodPed = ITP.Pedido) INNER JOIN BAR ON PRD.CodPrd = BAR.Produto WHERE (((PED.DtFat)>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (PED.DtFat)<=#" & Format(MebDtF, "mm/dd/yyyy") & "#)) GROUP BY NOTA.Modelo, PRD.Embalagem, PRD.CodPrd, PRD.Descrição_Produto, BAR.BarBarra, PRD.Classe, PRD.Icm HAVING (((NOTA.Modelo)='55'))")
                Reg0200.Requery
                If Reg0200.RecordCount > 0 Then
                    Reg0200.MoveFirst
                    While Not Reg0200.Eof
                        VarReg = "|0200" & "|" & Reg0200!CodPrd & "|" & Trim(Reg0200.Fields("Descrição_Produto")) & "|" & Trim(Reg0200!BarBarra) & "||"
                        VarReg = VarReg & Trim(Reg0200!Embalagem) & "|00|" & Trim(Reg0200!Classe) & "||||" & CInt(Reg0200!Icm) & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Linhas0 = Linhas0 + 1
                        Reg0200.MoveNext
                    Wend
                End If
                
                
                '  ************************** 0400
                
                Set Reg0400 = Banco.OpenRecordset("SELECT NOTA.Modelo, CFO.Descricao, PED.Cfo FROM (NOTA INNER JOIN PED ON NOTA.CodPed = PED.CodPed) INNER JOIN CFO ON PED.Cfo = CFO.Codigo WHERE (((PED.DtFat)>=#1/1/2013# And (PED.DtFat)<=#1/31/2013#)) GROUP BY NOTA.Modelo, CFO.Descricao, PED.Cfo HAVING (((NOTA.Modelo)='55'))")
                Reg0400.Requery
                If Reg0400.RecordCount > 0 Then
                    Reg0400.MoveFirst
                    While Not Reg0400.Eof
                        VarReg = "|0400" & "|" & Left(LIMPACampo2(Reg0400.Fields("Cfo")), 4) & "|" & Trim(Reg0400.Fields("Descricao")) & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Linhas0 = Linhas0 + 1
                        Reg0400.MoveNext
                    Wend
                End If
                
                Print #1, UTF8_Encode("|0990|" & Linhas0 & "|")
                
                ' ********************************** Bloco A
                Print #1, UTF8_Encode("|A001|1|")
                Print #1, UTF8_Encode("|A990|2|")
                
                ' ********************************** Bloco C
                
                Print #1, UTF8_Encode("|C001|0|")
                Print #1, UTF8_Encode("|C010|" & LimpaString(VarCnpj) & "|2|")
                LinhasC = 2
                
                '  ************************  C100
                
                Set RegC100 = Banco.OpenRecordset("SELECT PED.DtFat, NOTA.Modelo, PED.Operacao, PED.Entidade, PED.NroNt, PED.Desconto, NOTA.BaseIcms, PED.Total FROM NOTA INNER JOIN PED ON NOTA.CodPed = PED.CodPed WHERE (((PED.DtFat)>=#1/1/2013# And (PED.DtFat)<=#1/31/2013#) AND ((NOTA.Modelo)='55')) ORDER BY PED.DtFat")
                RegC100.Requery
                If RegC100.RecordCount > 0 Then
                    RegC100.MoveFirst
                    While Not RegC100.Eof
                    
                        Select Case .Fields("Operacao")
                                Case 0
                                    VarTipo = 0
                                Case 1
                                    VarTipo = 1
                                Case 2
                                    VarTipo = 1
                                Case 3
                                    VarTipo = 1
                                Case 4
                                    VarTipo = 1
                                Case 5
                                    VarTipo = 1
                                Case 6
                                    VarTipo = 1
                                Case 7
                                    VarTipo = 1
                                Case 8
                                    VarTipo = 1
                                Case 9
                                    VarTipo = 0
                                Case 10
                                    VarTipo = 0
                                Case 11
                                    VarTipo = 1
                                Case 12
                                    VarTipo = 0
                                Case 13
                                    VarTipo = 1
                        End Select
                    
                    
                        VarReg = "|C100" & "|" & VarTipo & "|0|" & RegC100!CodEntidade & "|" & Trim(RegC100!Nome) & "|1058|" & Trim(LimpaString(RegC100!CGC)) & "|" & Trim(LimpaString(RegC100!Cpf)) & "|"
                        VarReg = VarReg & Trim(RegC100!matricula) & "|" & Trim(LimpaString(RegC100!InscrEst)) & "||"
'                        VarReg = VarReg & Trim(RegC100!Endereco) & "|" & RegC100!Nro & "|" & Trim(RegC100!Complemento) & "|" & Trim(RegC100!Bairro) & "|"
'                        Print #1, UTF8_Encode(VarReg)
'                        LinhasC = LinhasC + 1
'                        RegC100.MoveNext
'                    Wend
                    Wend
                End If
                While Not .Eof
                
                    Criterio = "CodPed = " & .Fields("CodPed")
                    RegNota.FindFirst Criterio
                    
                    If Not RegNota.NoMatch And RegNota!Modelo = "55" Then
                    
                        Select Case .Fields("Operacao")
                                Case 0
                                    VarTipo = 0
                                Case 1
                                    VarTipo = 1
                                Case 2
                                    VarTipo = 1
                                Case 3
                                    VarTipo = 1
                                Case 4
                                    VarTipo = 1
                                Case 5
                                    VarTipo = 1
                                Case 6
                                    VarTipo = 1
                                Case 7
                                    VarTipo = 1
                                Case 8
                                    VarTipo = 1
                                Case 9
                                    VarTipo = 0
                                Case 10
                                    VarTipo = 0
                                Case 11
                                    VarTipo = 1
                                Case 12
                                    VarTipo = 0
                                Case 13
                                    VarTipo = 1
                        End Select
                                        
        '               psstring = 0 - espacos a direita | 2 = zeros a esquerda
        '               Texto = Space(768)
                        VarReg = "20" & "|" & "I" & "|" & "|" & RegNota("NaturezaOp") & "|" & VarSerieNota & "|" & .Fields("NroNt") & "|"
                        VarReg = VarReg & Format(.Fields("DtFat"), "DD/MM/YYYY HH:MM:SS") & "|" & Format(.Fields("DtSaida"), "DD/MM/YYYY HH:MM:SS") & "|"
                        VarReg = VarReg & VarTipo & "|" & Left(LIMPACampo2(.Fields("Cfo")), 4) & "|" & "|" & "|" & Trim(LimpaString(RegNota!Documento)) & "|"
                        
                        VarReg = VarReg & IIf(RegNota!RazaoSocial <> "", Left(RegNota!RazaoSocial, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Logradouro <> "", Left(RegNota!Logradouro, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Nro <> "", Left(RegNota!Nro, 60), "NAO IDENTIFICADO") & "|" & "|"
                        VarReg = VarReg & IIf(RegNota!Bairro <> "", Left(RegNota!Bairro, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & IIf(RegNota!Cidade <> "", Left(RegNota!Cidade, 60), "NAO IDENTIFICADO") & "|"
                        VarReg = VarReg & RegNota!UfNota & "|" & RegNota!Cep & "|" & "|" & "|" & RegNota!IEDestinatario
                        Cont20 = Cont20 + 1
                        Print #1, UTF8_Encode(VarReg)
                                        
                        Criterio = "Itp.Pedido = " & .Fields("CodPed")
                        RegItens.FindFirst Criterio
                        If Not RegItens.NoMatch Then
                            While Not RegItens.Eof
                                If RegItens("Itp.Pedido") = .Fields("CodPed") Then
                                    VarReg = "30" & "|" & RegItens!Produto & "|" & Left(RegItens!Descrição_Produto, 120) & "||" & IIf(RegItens!Embalagem = "", "UN", Left(RegItens!Embalagem, 6)) & "|"
                                    VarReg = VarReg & Format(RegItens!Qtd, "#######0.0000") & "|"
                                    VarReg = VarReg & Format(RegItens!ValorUnit, "###########0.0000") & "|"
                                    VarReg = VarReg & Format(RegItens!Valor, "############0.00") & "|" & "000" & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!AliqIcms), 0, RegItens!AliqIcms), "##0.00") & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!AliqIPI), 0, RegItens!AliqIPI), "##0.00") & "|"
                                    VarReg = VarReg & Format(IIf(IsNull(RegItens!ValorIPI), 0, RegItens!ValorIPI), "############0.00")
                                    Print #1, UTF8_Encode(VarReg)
                                    VarReg = ""
                                    Cont30 = Cont30 + 1
                                    VarTot = VarTot + RegItens!Valor
                                 End If
                                 RegItens.MoveNext
                            Wend
                        End If
                        
'                        VarReg = "40" & "|" & Format(RegNota!BaseIcms, "############0.00") & "|"
                        VarReg = "40" & "|" & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIcms, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!BaseIcmsSub, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIcmsSub, "############0.00") & "|"
'                        VarReg = VarReg & Format(RegNota!ValorTotalPrd, "############0.00") & "|"
                        VarReg = VarReg & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorFrete, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorSeguro, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorDesconto, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotIpi, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotOutras, "############0.00") & "|"
'                        VarReg = VarReg & Format(RegNota!TotNota, "############0.00") & "|"
                        VarReg = VarReg & Format(VarTot, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!TotServicos, "##0.00") & "|"
                        VarReg = VarReg & Format(RegNota!AliqIss, "############0.00") & "|"
                        VarReg = VarReg & Format(RegNota!ValorIss, "############0.00")
                        Print #1, UTF8_Encode(VarReg)
                        Cont40 = Cont40 + 1
                        VarTot = 0
                        
                        
                        VarReg = "50" & "|" & RegNota!TipoFrete - 1 & "|" & LimpaString(SeCampoNulo(RegNota, "CnpjTransp")) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "RazaoTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "IETransp"), 14) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "EndTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "CidadeTransp"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "UfTransp"), 2) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "PlacaVeiculo"), 8) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Uf"), 2) & "|"
                        VarReg = VarReg & Format(RegNota!QtdVolume, "##############0") & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Especie"), 60) & "|"
                        VarReg = VarReg & Left(SeCampoNulo(RegNota, "Marca"), 60) & "|" & "|"
                        VarReg = VarReg & Format(RegNota!PesoLiquido, "###########0.000") & "|"
                        VarReg = VarReg & Format(RegNota!PesoBruto, "###########0.000")
                        Print #1, UTF8_Encode(VarReg)
                        Cont50 = Cont50 + 1
                    
                        VarReg = "60" & "|" & RegNota!DadosFatura & "|" & RegNota!InformFisco & "|" & "|"
                        Print #1, UTF8_Encode(VarReg)
                        Cont60 = Cont60 + 1
                    End If
                    .MoveNext
                Wend
                VarReg = "90" & "|" & Format(Cont20, "00000") & "|" & Format(Cont30, "00000") & "|" & Format(Cont40, "00000") & "|" & Format(Cont50, "00000") & "|" & Format(Cont60, "00000") '& "|"
                Print #1, UTF8_Encode(VarReg)
            Else
                MsgBox "Arquivo sem movimento !", vbInformation
            End If
            
        End With
    
    
    
    
    End If
    
    Close #1
    
    MsgBox "Arquivo gerado com sucesso !", vbInformation
    
End If
End Sub
