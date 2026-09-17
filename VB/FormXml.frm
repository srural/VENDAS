VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormXml 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14685
   Begin VB.CommandButton BtBaixar 
      Caption         =   "Baixar X.M.L"
      Height          =   330
      Left            =   7695
      TabIndex        =   47
      Top             =   3420
      Width           =   2175
   End
   Begin MSMask.MaskEdBox MebCusto 
      Height          =   285
      Left            =   9630
      TabIndex        =   44
      Top             =   7380
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtEan 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12780
      TabIndex        =   43
      Top             =   7110
      Width           =   1680
   End
   Begin VB.CommandButton CmdSalvaPrd 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   13275
      TabIndex        =   42
      Top             =   7470
      Width           =   1185
   End
   Begin VB.TextBox DtSaida 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3420
      TabIndex        =   38
      Top             =   3465
      Width           =   1410
   End
   Begin VB.TextBox DtFat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1845
      TabIndex        =   37
      Top             =   3465
      Width           =   1410
   End
   Begin VB.TextBox NroNfe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   270
      TabIndex        =   36
      Top             =   3465
      Width           =   1410
   End
   Begin VB.TextBox TxtNcm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10665
      TabIndex        =   35
      Top             =   7110
      Width           =   1095
   End
   Begin VB.TextBox TxtUn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11385
      TabIndex        =   34
      Top             =   6570
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   4830
   End
   Begin VB.TextBox TxtCsT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11835
      TabIndex        =   32
      Top             =   7110
      Width           =   780
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      TabIndex        =   31
      Top             =   7650
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox Txtcfop 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      TabIndex        =   30
      Top             =   7110
      Width           =   960
   End
   Begin VB.TextBox TxtCodPrd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      TabIndex        =   29
      Top             =   6570
      Width           =   1410
   End
   Begin VB.TextBox TxtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      TabIndex        =   28
      Top             =   6840
      Width           =   4830
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   13545
      TabIndex        =   24
      Top             =   585
      Width           =   1095
      Begin VB.OptionButton Opt 
         Caption         =   "Saida"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   540
         Width           =   870
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Entrada"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   225
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdEnt 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   13590
      TabIndex        =   23
      Top             =   2835
      Width           =   960
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   22
      Top             =   675
      Width           =   5955
   End
   Begin VB.TextBox TxtProt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   21
      Top             =   945
      Width           =   5955
   End
   Begin VB.TextBox TxtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   20
      Top             =   1530
      Width           =   5955
   End
   Begin VB.TextBox TxtCnpj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   19
      Top             =   1260
      Width           =   2715
   End
   Begin VB.TextBox TxtIe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10755
      TabIndex        =   18
      Top             =   1260
      Width           =   2715
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   17
      Top             =   2880
      Width           =   5955
   End
   Begin VB.TextBox TxtEnd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   16
      Top             =   1800
      Width           =   4470
   End
   Begin VB.TextBox TxtNro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12285
      TabIndex        =   15
      Top             =   1800
      Width           =   1185
   End
   Begin VB.TextBox TxtBairro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   14
      Top             =   2340
      Width           =   3255
   End
   Begin VB.TextBox txtCidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   13
      Top             =   2070
      Width           =   4470
   End
   Begin VB.TextBox TxtUf 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12285
      TabIndex        =   12
      Top             =   2070
      Width           =   1185
   End
   Begin VB.TextBox TxtCep 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10845
      TabIndex        =   11
      Top             =   2340
      Width           =   1275
   End
   Begin VB.TextBox TxtIbge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12195
      TabIndex        =   10
      Top             =   2340
      Width           =   1275
   End
   Begin VB.TextBox TxtPais 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      TabIndex        =   9
      Top             =   2610
      Width           =   3255
   End
   Begin VB.TextBox TxtFone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10845
      TabIndex        =   8
      Top             =   2610
      Width           =   2625
   End
   Begin VB.ListBox LsCob 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   225
      TabIndex        =   7
      Top             =   7200
      Width           =   7125
   End
   Begin VB.ListBox Lsitens 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   225
      TabIndex        =   6
      Top             =   3960
      Width           =   14190
   End
   Begin VB.CommandButton BtImporta 
      Caption         =   "Ler  X.M.L."
      Height          =   330
      Left            =   5130
      TabIndex        =   5
      Top             =   3420
      Width           =   2175
   End
   Begin VB.FileListBox File 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   3690
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.DriveListBox Drive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   225
      TabIndex        =   3
      Top             =   720
      Width           =   3345
   End
   Begin VB.DirListBox Dir 
      Appearance      =   0  'Flat
      Height          =   2115
      Left            =   225
      TabIndex        =   2
      Top             =   1035
      Width           =   3345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Importa XML"
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
         TabIndex        =   1
         Top             =   45
         Width           =   4650
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   285
      Left            =   11385
      TabIndex        =   45
      Top             =   7650
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin VB.Label LblCodPrd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   14355
      TabIndex        =   46
      Top             =   6525
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Saida"
      Height          =   195
      Left            =   3420
      TabIndex        =   41
      Top             =   3285
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissao"
      Height          =   195
      Left            =   1845
      TabIndex        =   40
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Nfe"
      Height          =   195
      Left            =   270
      TabIndex        =   39
      Top             =   3285
      Width           =   555
   End
   Begin VB.Label LblCod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   13590
      TabIndex        =   27
      Top             =   1980
      Width           =   90
   End
End
Attribute VB_Name = "FormXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Doc As DOMDocument

Public xml As New MSXML2.DOMDocument

Public Det As MSXML2.IXMLDOMNodeList

Public nfeProc As MSXML2.IXMLDOMElement
Public Nfe As MSXML2.IXMLDOMElement
Public infNFe As MSXML2.IXMLDOMElement
Public infProt As MSXML2.IXMLDOMElement

Public Prod As MSXML2.IXMLDOMElement
Public imposto As MSXML2.IXMLDOMElement

Public cobr As MSXML2.IXMLDOMElement
Public fat As MSXML2.IXMLDOMElement

Public Dup As MSXML2.IXMLDOMElement

Public Cab As MSXML2.IXMLDOMElement
Public Cab2 As MSXML2.IXMLDOMElement

Public Ide As MSXML2.IXMLDOMElement

Public VarEnt As Recordset
Public VarPrd As Recordset
Public Criterio As String




Private Sub BtBaixar_Click()
'NFe.DistribuicaoDFePorChaveNFe(cUF, cCNPJ, aChNFe)

' AbreAcbr "NFe.DistribuicaoDFePorChaveNFe(SP,23103347000165,35230219511493000152550010000003041000004806)"

AbreAcbr "NFe.DistribuicaoDFePorUltNSU(SP,23103347000165,72)"

'EscreveINI "EVENTO", "IdLote", "1", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "cOrgao", "91", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "Cnpj", "23103347000165", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "chNFe", "35230219511493000152550010000003041000004806", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "dhEvento", Date & " " & Time, App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "tpEvento", "210240", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "nSeqEvento", "1", App.Path & "\cfe.ini"        ' NUMERO SEQUENCIA DO EVENTO'
'EscreveINI "EVENTO001", "versaoEvento", "4.00", App.Path & "\cfe.ini"
'EscreveINI "EVENTO001", "xJust", "TESTE DE PROGRAMACAO", App.Path & "\cfe.ini"

               
'AbreAcbr "NFE.EnviarEvento(" & App.Path & "\cfe.ini" & ")"

MsgBox "Comando enviado", vbInformation

End Sub

Private Sub BtImporta_Click()
Dim Temp(10) As String
Dim VarForPrd As Recordset
Dim Criterio As String


'Set Doc = New DOMDocument
'Doc.Load Dir.Path & "\" & File.filename

'Temp(0) = Doc.documentElement.ChildNodes(0).ChildNodes(0).Text
'Temp(1) = Doc.documentElement.ChildNodes(0).ChildNodes(1).Text
'Temp(2) = Doc.documentElement.ChildNodes(0).ChildNodes(2).Text


'Text1 = Temp(0) & vbCrLf & Temp(1) & vbCrLf & Temp(2) & vbCrLf


'Botão do FORM que faz a leitura:

   Dim X As Integer
   
'   Dlg1.Filter = "Arquivos XML|*.XML"
'   Dlg1.ShowOpen
   
'   If Dlg1.filename = "" Then
'      Exit Sub
'   End If
On Error Resume Next

      
   Set xml = New MSXML2.DOMDocument
   xml.async = False
   xml.Load Dir.Path & "\" & File.filename ' pega o nome e o path de um CommonDialog
   
  If xml.parseError.reason <> "" Then   'Verifica se Foi aberto com Sucesso
      MsgBox xml.parseError.reason & vbCrLf & _
             "Line: " & xml.parseError.Line & vbCrLf & _
             "Pos: " & xml.parseError.linepos
      Exit Sub
   End If
   
   Set nfeProc = xml.SelectSingleNode("nfeProc")
   
   If nfeProc Is Nothing Then
      MsgBox "NFe não tem protocolo de autorização", vbCritical, "Ler XML NF-e"
      Exit Sub
   End If
   
   Set VarEnt = Banco.OpenRecordset("Select * From Ent Order By CodEntidade")
   VarEnt.Requery
   
   Set VarPrd = Banco.OpenRecordset("Select * From Prd Order By CodPrd")
   VarPrd.Requery
   
   
   Set Nfe = nfeProc.SelectSingleNode("NFe")
   Set infNFe = Nfe.SelectSingleNode("infNFe")
   
   ' Pega Chave da NF-e
   If infNFe.Attributes.getNamedItem("versao") Is Nothing Then
      TxtId.Text = "Não encontrado!"
   
   ElseIf infNFe.Attributes.getNamedItem("versao").Text = "1.10" Then
      TxtId.Text = Format(Mid(infNFe.Attributes.getNamedItem("Id").Text, _
          4, 44), "@@.@@@@.@@@@@@@@@@@@@@.@@.@@@.@@@@@@@@@.@@@@@@@@@-@")
   
   ElseIf infNFe.Attributes.getNamedItem("versao").Text = "2.00" Then
      TxtId.Text = Format(Mid(infNFe.Attributes.getNamedItem("Id").Text, _
          4, 44), "@@.@@@@.@@@@@@@@@@@@@@.@@.@@@.@@@@@@@@@.@.@@@@@@@@-@")

   End If
   
   If nfeProc.ChildNodes.Length < 2 Then
      
      TxtProt.Text = "NFe não tem protocolo de autorização"
      Exit Sub
      
   Else
   
      Set infProt = nfeProc.SelectSingleNode("protNFe/infProt")
      TxtProt.Text = infProt.SelectSingleNode("nProt").Text & _
                     "-" & infProt.SelectSingleNode("dhRecbto").Text
      
   End If
   
   ' Pega totas as TAG's Cab do XML
    Set Ide = nfeProc.SelectSingleNode("NFe/infNFe/ide")
    NroNfe = Ide.SelectSingleNode("cNf").Text
    DtFat = Ide.SelectSingleNode("dEmi").Text
    DtSaida = Ide.SelectSingleNode("dSaiEnt").Text

   If OPt(0).value = True Then
        Set Cab = nfeProc.SelectSingleNode("NFe/infNFe/emit")
   Else
        Set Cab = nfeProc.SelectSingleNode("NFe/infNFe/dest")
   End If
        TxtCnpj = Format(Cab.SelectSingleNode("CNPJ").Text, "@@.@@@.@@@/@@@@-@@")
        TxtNome = Cab.SelectSingleNode("xNome").Text
        TxtIe = Cab.SelectSingleNode("IE").Text
        TxtEmail = Cab.SelectSingleNode("email").Text
   If OPt(0).value = True Then
        Set Cab2 = nfeProc.SelectSingleNode("NFe/infNFe/emit/enderEmit")
    Else
        Set Cab2 = nfeProc.SelectSingleNode("NFe/infNFe/dest/enderDest")
    End If
        TxtEnd = Cab2.SelectSingleNode("xLgr").Text
        TxtNro = Cab2.SelectSingleNode("nro").Text
        TxtBairro = Cab2.SelectSingleNode("xBairro").Text
        txtCidade = Cab2.SelectSingleNode("xMun").Text
        TxtUf = Cab2.SelectSingleNode("UF").Text
        TxtCep = Format(Cab2.SelectSingleNode("CEP").Text, "@@.@@@-@@@")
        TxtIbge = Cab2.SelectSingleNode("cMun").Text
        TxtPais = Cab2.SelectSingleNode("xPais").Text
        TxtFone = Format(Cab2.SelectSingleNode("fone").Text, "(@@)@@@@-@@@@")
        TxtEnd = Cab2.SelectSingleNode("xLgr").Text
        
    Criterio = "Cgc = " & """" & LIMPACampo2(TxtCnpj) & """"
    VarEnt.FindFirst Criterio
    If Not VarEnt.NoMatch Then
        LblCod = VarEnt!CodEntidade
    End If
    
    CmdEnt.Enabled = True
    
   
   ' Pega totas as TAG's det do XML
   
   Set Det = infNFe.SelectNodes("det")
   
   Lsitens.Clear
'    Lsitens.ListItems.Clear
   Set VarForPrd = Banco.OpenRecordset("SELECT PRDFOR.*, PRDFOR.CodFor FROM PRDFOR WHERE PRDFOR.CodFor = " & LblCod)
    
   For X = 0 To Det.Length - 1
   
        Set Prod = Det.Item(X).SelectSingleNode("prod")
        Set imposto = Det.Item(X).SelectSingleNode("imposto")
        
        Criterio = "CodFor = " & LblCod & " CodPrdFor = " & Prod.SelectSingleNode("cProd").Text
        VarForPrd.FindFirst Criterio
        If VarForPrd.NoMatch Then
        ' Montagem dos itens na lista
               Lsitens.AddItem _
                       psCompString(1, X, 5) & " | " & _
                       psCompString(2, Prod.SelectSingleNode("cEAN").Text, 13) & " | " & _
                       psCompString(2, Prod.SelectSingleNode("cProd").Text, 12) & " | " & _
                       psCompString(3, Left(Prod.SelectSingleNode("xProd").Text, 40), 40) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("uCom").Text, 5) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("NCM").Text, 8) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("CFOP").Text, 4) & " | "
        Else
               Lsitens.AddItem _
                       psCompString(1, X, 5) & " | " & _
                       psCompString(2, Prod.SelectSingleNode("cEAN").Text, 13) & " | " & _
                       psCompString(2, Prod.SelectSingleNode("cProd").Text, 12) & " | " & _
                       psCompString(3, Left(Prod.SelectSingleNode("xProd").Text, 40), 40) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("uCom").Text, 5) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("NCM").Text, 8) & " | " & _
                       psCompString(3, Prod.SelectSingleNode("CFOP").Text, 4) & " | " & psCompString(1, VarForPrd!CodPrd, 6)
        End If
   Next
   
   Set cobr = infNFe.SelectSingleNode("cobr")
   'Set dup = cobr.selectNodes("dup")
   
   LsCob.Clear
   
   If Not (cobr Is Nothing) Then
   
      For X = 0 To cobr.ChildNodes.Length - 1
         
         Set fat = cobr.ChildNodes(X)
         
         LsCob.AddItem _
            fat.SelectSingleNode("nDup").Text & "|" & _
            Format(fat.SelectSingleNode("dVenc").Text, "dd/mm/yyyy") & "|" & _
            Format(Val(fat.SelectSingleNode("vDup").Text), "#,###,##0.00")
            
      Next
   Else
   
      LsCob.AddItem "Não tem cobrança..."
      
   End If
   
End Sub

Private Sub BtSair_Click()
Unload Me
End Sub

'Private Sub CmbExclui_Click()
'If MsgBox("Confirma exclusão ?", vbYesNo, App.Title) = vbYes Then
'    If Datdados(0).Recordset.RecordCount <> 0 Then
'        Datdados(0).Recordset.Delete
'    End If
'End If
'End Sub


Private Sub CmdEnt_Click()
If LblCod = "" Then
    VarEnt.AddNew
    VarEnt!Ativo = 1
Else
    VarEnt.Edit
End If
VarEnt!Nome = Left(TxtNome, 40)
VarEnt!CGC = LIMPACampo2(TxtCnpj)
VarEnt!InscrEst = TxtIe
VarEnt!Endereco = TxtEnd
VarEnt!Nro = TxtNro
VarEnt!Cidade = txtCidade
VarEnt!Uf = TxtUf
VarEnt!Bairro = TxtBairro
VarEnt!Cep = Trim(LIMPACampo2(TxtCep))
VarEnt!matricula = TxtIbge
VarEnt!pais = "BRASIL"
VarEnt!Fone = Trim(LIMPACampo2(TxtFone))
VarEnt!Tipo = 2
VarEnt!Classificacao = 2

VarEnt.Update
If LblCod = "" Then
    VarEnt.Bookmark = VarEnt.LastModified
    LblCod = VarEnt!CodEntidade
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Dir_Change()
    File.Path = Dir.Path
End Sub

Private Sub Drive_Change()
    Dir.Path = Drive.Drive
End Sub

Private Sub Form_Activate()
Dim Formulario As Form

Set Formulario = Me

'Me.Left = 200 ' (MDIPrincipal.Width - Me.Width) / 2
'Me.Top = 300  '(MDIPrincipal.Height - Me.Height) / 6


End Sub

Private Sub Form_Load()
Dim Componentes As Variant

'For Each Componentes In DatDados
'    Componentes.DatabaseName = Caminho & "\Dados.mdb"
'Next

'Dir.Path = "S:\Trabalho\Xml joframa\Xml recebidos"
'File.Path = "S:\Trabalho\Xml joframa\Xml recebidos"

Dir.Path = Drive.Drive
File.Path = Dir.Path


End Sub

Private Sub Lsitens_Click()
   Set Det = infNFe.SelectNodes("det")
   
   Set Prod = Det.Item(Left(Lsitens, 5)).SelectSingleNode("prod")
   Set imposto = Det.Item(Left(Lsitens, 5)).SelectSingleNode("imposto/ICMS/ICMS40")
     
   TxtCodPrd = psCompString(2, Prod.SelectSingleNode("cProd").Text, 13)
   TxtUn = psCompString(3, Prod.SelectSingleNode("uCom").Text, 5)
   TxtDesc = psCompString(3, Prod.SelectSingleNode("xProd").Text, 60)
   Txtcfop = psCompString(2, Prod.SelectSingleNode("CFOP").Text, 13)
   TxtNcm = psCompString(2, Prod.SelectSingleNode("NCM").Text, 13)
   TxtEan = psCompString(2, Prod.SelectSingleNode("cEAN").Text, 13)
   MebCusto = psCompString(2, Prod.SelectSingleNode("vUnCom").Text, 3)
   
End Sub

Private Sub CmdSalvaPrd_Click()

If LblCodPrd = "" Then
    VarPrd.AddNew
    VarPrd!Ativo = 1
Else
    VarPrd.Edit
End If
VarPrd!Descrição_Produto = Left(TxtDesc, 40)
VarPrd!Embalagem = TxtUn
VarPrd!Cfop = Txtcfop
VarPrd!Classe = TxtNcm
VarEnt!Nro = TxtNro
VarEnt!Cidade = txtCidade
VarEnt!Uf = TxtUf
VarEnt!Bairro = TxtBairro
VarEnt!Cep = Trim(LIMPACampo2(TxtCep))
VarEnt!matricula = TxtIbge
VarEnt!pais = "BRASIL"
VarEnt!Fone = Trim(LIMPACampo2(TxtFone))
VarEnt!Tipo = 2
VarEnt!Classificacao = 2

VarEnt.Update
If LblCod = "" Then
    VarEnt.Bookmark = VarEnt.LastModified
    LblCod = VarEnt!CodEntidade
End If

End Sub
