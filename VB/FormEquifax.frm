VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormEquifax 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11160
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -585
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   14
      Top             =   7515
      Width           =   14370
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11175
      TabIndex        =   12
      Top             =   0
      Width           =   11175
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Geração do arquivo texto da Equifax"
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
         TabIndex        =   13
         Top             =   45
         Width           =   7530
      End
   End
   Begin VB.ComboBox CmbTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "FormEquifax.frx":0000
      Left            =   3465
      List            =   "FormEquifax.frx":001C
      TabIndex        =   10
      Top             =   1215
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   5715
      TabIndex        =   7
      Top             =   675
      Width           =   1590
      Begin VB.OptionButton OPt 
         Caption         =   "&Pagamentos"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   495
         Width           =   1275
      End
      Begin VB.OptionButton OPt 
         Caption         =   "&Vencidos"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\(Rede)\Sistemas\Vendas\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4185
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4545
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton BtGera 
      Caption         =   "&Gera Arquivo"
      Height          =   330
      Left            =   7650
      TabIndex        =   4
      Top             =   1260
      Width           =   1455
   End
   Begin VB.CommandButton BtOK 
      Caption         =   "&Pesquisa"
      Height          =   330
      Left            =   7650
      TabIndex        =   3
      Top             =   765
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MebDtI 
      Height          =   330
      Left            =   225
      TabIndex        =   0
      Top             =   1215
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14737632
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebDtF 
      Height          =   330
      Left            =   1845
      TabIndex        =   1
      Top             =   1215
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14737632
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid DbGrid 
      Bindings        =   "FormEquifax.frx":0040
      Height          =   5280
      Left            =   180
      OleObjectBlob   =   "FormEquifax.frx":0053
      TabIndex        =   6
      Top             =   1845
      Width           =   10815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   3465
      TabIndex        =   11
      Top             =   1035
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   9360
      Picture         =   "FormEquifax.frx":2CF5
      ToolTipText     =   "Progama desenvolvido pela SidComp Informática"
      Top             =   675
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final"
      Height          =   195
      Left            =   1845
      TabIndex        =   5
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Inicial"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   990
      Width           =   795
   End
End
Attribute VB_Name = "FormEquifax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtGera_Click()
Dim Tipo As String
Dim Texto As String
Dim CgcTmp As String

If Opt(0).value = True Then
    Tipo = "U"
Else
    Tipo = "J"
End If

If MsgBox("Confirma geração do arquivo texto ?", vbYesNo, App.Title) = vbYes Then
    Open "ArqTexto.Txt" For Output As #1
    With Data.Recordset
        .MoveFirst
        While Not .Eof
            
            Texto = Tipo
            If Not IsNull("CGC") Then
                CgcTmp = LimpaMask(.Fields("Cgc"))
                If Len(CgcTmp) > 12 Then
                    Texto = Texto & psCompString(1, CgcTmp, 14)
                    Texto = Texto & psCompString(3, .Fields("Nome"), 55)
                    If Not IsNull(.Fields("Fantasia")) Then
                        Texto = Texto & psCompString(1, .Fields("Fantasia"), 55)
                    Else
                        Texto = Texto & Space(55)
                    End If
                Else
                    Texto = Texto & psCompString(1, CgcTmp, 11)
                    Texto = Texto & psCompString(3, .Fields("InscrEst"), 20)
                    Texto = Texto & "    "
                    Texto = Texto & psCompString(3, .Fields("Nome"), 55)
                End If
                Texto = Texto & "M"
                Texto = Texto & psCompString(3, Left(.Fields("Endereco"), Len(.Fields("Endereco"))) & .Fields("Nro"), 70)
                Texto = Texto & psCompString(3, .Fields("Cidade"), 30)
                Texto = Texto & psCompString(3, .Fields("Uf"), 2)
                If Not IsNull(.Fields("Fone")) Then
                    Texto = Texto & psCompString(3, Left(.Fields("Fone"), 2), 4)
                    Texto = Texto & psCompString(3, Mid(.Fields("Fone"), 2, 8), 10)
                Else
                    Texto = Texto & Space(14)
                End If
                If Not IsNull(.Fields("Fax")) Then
                    Texto = Texto & psCompString(3, Left(.Fields("Fax"), 2), 4)
                    Texto = Texto & psCompString(3, Mid(.Fields("Fax"), 2, 8), 10)
                Else
                    Texto = Texto & Space(14)
                End If
                Texto = Texto & Space(50)
                Texto = Texto & Space(6)
                Texto = Texto & psCompString(1, .Fields("Pedido"), 10)
                Texto = Texto & "C"
                Texto = Texto & psCompString(1, CCur(.Fields("Valor")), 11)
                Texto = Texto & psCompString(1, Right(Format(.Fields("Valor"), "##,##0.00"), 2), 2)
                Texto = Texto & psCompString(1, CCur(.Fields("Valor")), 11)
                Texto = Texto & psCompString(1, Right(Format(.Fields("Valor"), "##,##0.00"), 2), 2)
                Texto = Texto & Format(.Fields("DataEmiss"), "ddmmyyyy")
                Texto = Texto & Format(.Fields("DtVenc"), "ddmmyyyy")
                Texto = Texto & Format(.Fields("DtPgto"), "ddmmyyyy")
                If Len(CgcTmp) < 12 Then
                    Texto = Texto & Space(34)
                End If
                Texto = Texto & Chr(13)
                Print #1, Texto
            End If
            .MoveNext
        Wend
    End With
    Close #1
    MsgBox "Arquivo gerado com sucesso !", vbInformation
End If
End Sub

Private Sub BtOK_Click()

'DatPed.RecordSource = "Select PED.* , PED.DATA  from ped where Ped.Data >= #" & Format(MebDtI, "mm/dd/yyyy") & "# and Ped.Data <= #" & Format(MebDtF, "mm/dd/yyyy") & "#"
'DatPed.RecordSource = "SELECT CLI.CODCLI, CLI.NOME, CLI.ENDERECO, CLI.NRO, CLI.BAIRRO, CLI.DTNASC, CLI.CIDADE, CLI.UF, CLI.CEP, CLI.FONE, CLI.FAX, CLI.CPF_CGC, CLI.RG_INSC, PCL.CODPED, PCL.NROPCL, PCL.VENC, PCL.PGT, PCL.VALOR, PCL.COND, PCL.NROBCO, PED.NOTA, PED.DATA FROM CLI INNER JOIN (PED INNER JOIN PCL ON PED.CODPED = PCL.CODPED) ON CLI.CODCLI = PED.CODCLI WHERE (((PED.DATA)>#1/10/2001# And (PED.DATA)<#1/20/2001#))"
'DatPed.Refresh
'DatPcl.RecordSource = "SELECT PCL.*, PCL.CODPED FROM PCL WHERE PCL.CODPED = 0 ORDER BY PCL.CODPED"
'DatPcl.Refresh
'DatCli.RecordSource = "Select * from Cli Order by CodCli"
'DatCli.Refresh

'    Data.RecordSource = "SELECT CLI.CODCLI, CLI.NOME, CLI.ENDERECO, CLI.NRO, CLI.BAIRRO, CLI.DTNASC, CLI.CIDADE, CLI.UF, CLI.CEP, CLI.FONE, CLI.FAX, CLI.CPF_CGC, CLI.RG_INSC, PCL.CODPED, PCL.NROPCL, PCL.VENC, PCL.PGT, PCL.VALOR, PCL.COND, PCL.NROBCO, PED.NOTA, PED.DATA FROM CLI INNER JOIN (PED INNER JOIN PCL ON PED.CODPED = PCL.CODPED) ON CLI.CODCLI = PED.CODCLI WHERE Ped.Data >= #" & Format(MebDtI, "mm/dd/yyyy") & "# and Ped.Data <= #" & Format(MebDtF, "mm/dd/yyyy") & "# Order by Ped.CodPed"



If Opt(0).value = True Then
    Data.RecordSource = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON DUP.Pedido = PED.CodPed WHERE DUP.DtVenc >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And DUP.DtVenc <= #" & Format(MebDtF, "mm/dd/yyyy") & "# AND DUP.Cond = False AND Tipo_Dup = " & """" & CmbTipo.Text & """" & " Order By Dup.Pedido"
Else
    Data.RecordSource = "SELECT ENT.CodEntidade, ENT.Nome, ENT.Fantasia, ENT.Endereco, ENT.Nro, ENT.Bairro, ENT.Cidade, ENT.Uf, ENT.Cep, ENT.Fone, ENT.Fax, ENT.InscrEst, ENT.CGC, ENT.DtNasc, DUP.Pedido, DUP.NroPcl, DUP.DtVenc, DUP.DtPgto, DUP.Valor, DUP.Documento, DUP.Tipo_Dup, PED.DataEmiss, PED.NroNt, DUP.Cond FROM (ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade) INNER JOIN PED ON DUP.Pedido = PED.CodPed WHERE DUP.DtPgto >=#" & Format(MebDtI, "mm/dd/yyyy") & "# And DUP.DtPgto <= #" & Format(MebDtF, "mm/dd/yyyy") & "# AND DUP.Cond = true AND Tipo_Dup = " & """" & CmbTipo.Text & """" & " Order By Dup.Pedido"
End If


'If OPt(0).Value = True Then
'    Data.RecordSource = "SELECT CLI.CODCLI, CLI.NOME, CLI.APELIDO, CLI.ENDERECO, CLI.NRO, CLI.BAIRRO, CLI.DTNASC, CLI.CIDADE, CLI.UF, CLI.CEP, CLI.FONE, CLI.FAX, CLI.CPF_CGC, CLI.RG_INSC, PCL.CODPED, PCL.NROPCL, PCL.VENC, PCL.PGT, PCL.VALOR, PCL.COND, PCL.NROBCO, PED.NOTA, PED.DATA FROM CLI INNER JOIN (PED INNER JOIN PCL ON PED.CODPED = PCL.CODPED) ON CLI.CODCLI = PED.CODCLI WHERE Pcl.Venc >= #" & Format(MebDtI, "mm/dd/yyyy") & "# and Pcl.Venc <= #" & Format(MebDtF, "mm/dd/yyyy") & "# and Pcl.Cond = 'A'  Order by Ped.CodPed"
'Else
'    Data.RecordSource = "SELECT CLI.CODCLI, CLI.NOME, CLI.APELIDO, CLI.ENDERECO, CLI.NRO, CLI.BAIRRO, CLI.DTNASC, CLI.CIDADE, CLI.UF, CLI.CEP, CLI.FONE, CLI.FAX, CLI.CPF_CGC, CLI.RG_INSC, PCL.CODPED, PCL.NROPCL, PCL.VENC, PCL.PGT, PCL.VALOR, PCL.COND, PCL.NROBCO, PED.NOTA, PED.DATA FROM CLI INNER JOIN (PED INNER JOIN PCL ON PED.CODPED = PCL.CODPED) ON CLI.CODCLI = PED.CODCLI WHERE Pcl.Pgt >= #" & Format(MebDtI, "mm/dd/yyyy") & "# and Pcl.Pgt <= #" & Format(MebDtF, "mm/dd/yyyy") & "#  and Pcl.Cond = 'P' Order by Ped.CodPed"
'End If
Data.Refresh



End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

MebDtI.Text = Date
MebDtF.Text = Date

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
'Dim Componentes As Variant

Me.Left = 300 ' (MDIPrincipal.Width - Me.Width) / 2
Me.Top = 300 '(MDIPrincipal.Height - Me.Height) / 6

'For Each Componentes In DatDados
    Data.DatabaseName = Caminho & "\Dados.mdb"
'Next
CmbTipo.ListIndex = 0
End Sub

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


Function LIMPACampo(Campo As String) As String
Dim x As Integer
Dim CampoLimpo As String

For x = 1 To Len(Campo)
    If Mid$(Campo, x, 1) >= Chr(32) And Mid$(Campo, x, 1) <= Chr(122) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, x, 1)
    End If
Next
LIMPACampo = psCompString(3, CampoLimpo, Len(CampoLimpo))
End Function

Function LimpaMask(Campo As String) As String
Dim x As Integer
Dim CampoLimpo As String

For x = 1 To Len(Campo)
    If Mid$(Campo, x, 1) <> "/" And Mid$(Campo, x, 1) <> "." And Mid$(Campo, x, 1) <> "-" Then
            CampoLimpo = CampoLimpo + Mid$(Campo, x, 1)
    End If
Next
LimpaMask = psCompString(3, CampoLimpo, Len(CampoLimpo))
End Function



