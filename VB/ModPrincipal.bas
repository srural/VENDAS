Attribute VB_Name = "ModPrincipal"


Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetInputState Lib "user32" () As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, _
                                                                     ByVal lpOperation As String, _
                                                                     ByVal lpFile As String, _
                                                                     ByVal lpParameters As String, _
                                                                     ByVal lpDirectory As String, _
                                                                     ByVal nShowCmd As Long) As Long
 
Global oExcel As Object
Global objExlSht As Object
Global db As Database
Global Sn As Recordset   ' Recordset do tipo Snapshot

Public Type ExlCell
   Row As Long
   Col As Long
End Type


'variaveis do SRCaixa
Global BancoSRCaixa As New ADODB.Connection 'banco de dados do SRCaixa
Global BcoAutomasis As New ADODB.Connection 'Banco site automasis
Global Pesquisa(1 To 2) As ADODB.Recordset 'variavel ultilizada para os relatorios na impressora matricial
'------


Declare Sub openport Lib "C:\windows\system32\tsclib.dll" (ByVal PrinterName As String)
Declare Sub closeport Lib "C:\windows\system32\tsclib.dll" ()
Declare Sub sendcommand Lib "C:\windows\system32\tsclib.dll" (ByVal command As String)
Declare Sub setup Lib "C:\windows\system32\tsclib.dll" (ByVal LabelWidth As String, ByVal LabelHeight As String, ByVal Speed As String, ByVal Density As String, ByVal Sensor As String, ByVal Vertical As String, ByVal Offset As String)
Declare Sub downloadpcx Lib "C:\windows\system32\tsclib.dll" (ByVal filename As String, ByVal ImageName As String)
Declare Sub barcode Lib "C:\windows\system32\tsclib.dll" (ByVal X As String, ByVal y As String, ByVal CodeType As String, ByVal Height As String, ByVal Readable As String, ByVal rotation As String, ByVal Narrow As String, ByVal Wide As String, ByVal Code As String)
Declare Sub printerfont Lib "C:\windows\system32\tsclib.dll" (ByVal X As String, ByVal y As String, ByVal FontName As String, ByVal rotation As String, ByVal Xmul As String, ByVal Ymul As String, ByVal Content As String)
Declare Sub clearbuffer Lib "C:\windows\system32\tsclib.dll" ()
Declare Sub printlabel Lib "C:\windows\system32\tsclib.dll" (ByVal NumberOfSet As String, ByVal NumberOfCopy As String)
Declare Sub formfeed Lib "C:\windows\system32\tsclib.dll" ()
Declare Sub nobackfeed Lib "C:\windows\system32\tsclib.dll" ()
Declare Sub windowsfont Lib "C:\windows\system32\tsclib.dll" (ByVal X As Integer, ByVal y As Integer, ByVal fontheight As Integer, ByVal rotation As Integer, ByVal fontstyle As Integer, ByVal fontunderline As Integer, ByVal FaceName As String, ByVal TextContent As String)




Public Banco As Database ' Representa o Banco de Dados ( Arquivo.mdb)
Public BcoEst As Database
Global Extra As Recordset
Global Itens As Recordset
Global Entidade As Recordset
'Global Pedidos As Recordset
Global Parcelas  As Recordset
Global Empresa As Recordset
Global Nota As Recordset
Global Configuracao As Recordset
Global Lojas As Recordset
Global varPesquisa(1 To 2) As Recordset
Global ContCompra  As Boolean
Global VarPzoProduto As Boolean
Global VarPcoPrazo As Boolean
Global QtdLinhaFatura As Byte
Global InterFinanc As Boolean
Global LocalBcoFin As String
Global CasasDec As String
Global LocalEstoque As String
Global VarCustoPed As Boolean
Global Admin As Boolean
Global VarObsAutomatica As Boolean ' gera parcela automatica no vendas
Global VarMostraPgtoCaixa As Boolean ' mostra tipo de pgto no fechamento do caixa
Global VarSenhaBco ' Senha para acessar o banco de dados remoto
Global VarFunRural As Boolean
Global VarPrecoCima As Currency
Global VarCompObs As Byte
Global VarImprBarra As String
Global VarTranspNt As Boolean
Global VarTamCod As Byte
Global VarTamBarra As Byte
Global VarSenhaPedido As String
Global VarHistEst As Boolean
Global VarReajConta As Boolean
Global VarPedCompleto As Boolean
Global VarFinanceiro As Boolean
Global VarSerieNota As String
Global VarSimples As Boolean
Global VarModeloNota As String
Global VarAssistencia As Boolean
Global VarOrdemPed As Boolean
Global VarQtdFixa As Integer
Global VarBloquearEst As Boolean
Global VarTelEmite As String
Global VarPromissoriaBca As Boolean
Global VarValorPedidoVista As Currency
Global VarDtSaldoAtual As Boolean
Global VarLimiteNaParcela As Boolean
Global VarArquivoNfe As String
Global VarChaveNfe As String
Global VarComposicao As Boolean
Global VarPrecoPedido As Boolean



Global soma As Recordset

Global Linha As Integer 'Qtd de Linhas a serem impressas no txt
Global Ant As Integer 'Linha anterior impressa no txt

Public Formulário As Form ' Representa o formulário de cadastro que estiver ativo
Global Caminho As String
Public Operadores(7) As String
Public NomeDoBanco As String
Public Senha As Boolean
Public NomeEmpresa As String
Public NomeInscr As String
Public VarInscr As String
Public VarEndEmpresa As String
Public CidadeEmpresa As String
Public Cabecalho1 As String
Public Cabecalho2 As String
Public Cnpj As String
Public SqlDiv(0 To 3) As String

Global VarCodPed As Currency

Global Nor As String 'Impressao em Modo Normal
Global Cond As String 'Impressao em Modo Condicionado
Global Gde As String 'Impressao em Modo Grande
Global Negi As String 'Impressao em Modo Negrito
Global NegF As String 'Desliga Negrito
Global SaltoPag As String 'Salto da Página
Global Porta As String 'Porta onde se localiza a impressora
Global UltimaLinhaImpressa As Currency
Global Variavel(1 To 40) As Variant  'variaveis de somas

Global Arq 'Impressao em Linhas
Global Fs2 'Impressao em Linhas

Global VarPermitePrd As Boolean
Global VarPermiteEnt As Boolean
Global VarLojaUnica As Boolean

Public Maquina As String
Public LocalDrive As String

Public PortaAcbr As String

Public Declare Function Extenso Lib "Extens32.dll" Alias "extenso" (ByVal Valor As String, ByVal Retorno As String) As Integer

Public Sub AtivarMenu(Ativar As Boolean)

' Quando o Ativar = Verdadeiro os menus do formMdi serão verdadeiros

With MDIPrincipal
    .TlbPrincipal.Visible = Ativar
'    .MnuPrincipal(3).Visible = Ativar
'    .MnuFim(1).Visible = Ativar
'    .MnuPrincipal(2).Visible = Ativar ' Menu Ferramentas
'    .MnuPrincipal(3).Visible = Ativar ' Menu Janela
End With

End Sub

Public Sub LimpaTela()
Dim Mascara As String
Dim Componente As Variant

With Formulário
    If .CTextBox Then
        For Each Componente In .TxtDados
            ' Componente ira representar uma das caixas de texto cujo nome é TxtDados a cada execução do looping
            Componente.Text = ""
        Next
    End If
        ' O comando next faz com que a variavel componente represente a próxima caixa de texto
        
        ' Quando o Objeto possuir a propriedade promptinclude true e uma mascara
        ' ele não poderá ser limpo com "". Portando devemos armazenar a mascara em uma
        ' variavel temporaria, apagar a mascara do objeto, apagar o conteudo do objeto com "",
        ' e devolver a mascara ao objeto
    If .CMaskEditBox Then
        For Each Componente In .MebDados
            Mascara = Componente.Mask
            Componente.Mask = ""
            Componente.Text = ""
            Componente.Mask = Mascara
        Next
    End If
    
    If .CComboBox Then
        For Each Componente In .CmbDados
            ' O ListIndex identifica qual dos itens de uma combo foi selecionada começando de 0 e indo até
            ' o total de itens - 1. Quando o valor -1 é atribuido a propriedade text será limpa.
            
            Componente.ListIndex = -1
        Next
    End If
        
    If .CdbCombo Then
        For Each Componente In .DbcDados
        
            ' BoundText representa o conteudo do campo referenciado por boundcolum,
            ' quando o usuario selecionar uma opção válida o boundtext recebera um valor,
            ' caso contrario ele é igual a "".
            ' Por tanto se atribuirmos "" o text será limpo automaticamnente
            
            Componente.BoundText = ""
        Next
    End If
    
    If .CCheckBox Then
        For Each Componente In .ChkDados
            Componente.value = 0
        Next
    End If
End With

End Sub

Public Sub MostraDados()
Dim Mascara As String
Dim Componente As Variant
With Formulário
    If .CTextBox Then
        For Each Componente In .TxtDados
            Componente.Text = IIf(IsNull(.Registros.Fields(Componente.DataField)), "", .Registros.Fields(Componente.DataField))
        Next
     End If
     If .CMaskEditBox Then
        For Each Componente In .MebDados
               If .Registros.Fields(Componente.DataField).Type = dbDate Then
                    ' a propriedade type de um campo represente o tipo do campo dentro do banco de dados
                    If IsNull(.Registros.Fields(Componente.DataField)) Then
                        Mascara = Componente.Mask
                        Componente.Mask = ""
                        Componente.Text = ""
                        Componente.Mask = Mascara
                    Else
                        Componente.Text = Format(.Registros.Fields(Componente.DataField), "dd/mm/yyyy")
                    End If
                Else
                    Componente.Text = IIf(IsNull(.Registros.Fields(Componente.DataField)), "", .Registros.Fields(Componente.DataField))
                End If
        Next
     End If
     If .CComboBox Then
        For Each Componente In .CmbDados
            If Mid(Componente.Tag, 2, 1) = "I" Then
                Componente.ListIndex = IIf(IsNull(.Registros.Fields(Componente.DataField)), -1, .Registros.Fields(Componente.DataField)) - 1
            Else
                Componente.Text = IIf(IsNull(.Registros.Fields(Componente.DataField)), "", .Registros.Fields(Componente.DataField))
            End If
        Next
     End If
     
     If .CdbCombo Then
        For Each Componente In .DbcDados
            If Mid(Componente.Tag, 1, 1) = "S" Then
                Componente.BoundText = IIf(IsNull(.Registros.Fields(Componente.DataField)), "", .Registros.Fields(Componente.DataField))
                Componente.Refresh
            End If
        Next
     End If
     
     If .CCheckBox Then
        For Each Componente In .ChkDados
            If Componente.Tag = "S" Then
                Componente.value = IIf(.Registros.Fields(Componente.DataField) = True, 1, 0)
            Else
                Componente.value = IIf(.Registros.Fields(Componente.DataField) = "SIM", 1, 0)
            End If
        Next
     End If
     If .ComandBotom Then
        For Each Componente In .CmdDados
            Componente.Enabled = Not .Registros.Fields(Componente.Tag)
        Next
     End If
     
End With

End Sub

Public Sub Movimentos(Botao As Integer)
' Quando o usuário selecionar um botão, e este estiver entre o 5 e o 8, está rotina será invocado e a variavel botão receberá o numero do
' botão clicado

With Formulário
    If .Registros.RecordCount = 0 Then
        MsgBox "Não existem dados cadastrados", vbInformation, App.Title
    Else
        Select Case Botao
            Case 6
                .Registros.MoveFirst
            Case 7
               .Registros.MovePrevious
               If .Registros.Bof = True Then
                   .Registros.MoveFirst
               End If
            Case 8
                .Registros.MoveNext
                If .Registros.Eof = True Then
                    .Registros.MovePrevious
                End If
            Case 9
                .Registros.MoveLast  'dbRunAsync  posicionamento do fim dos registros fica mais rápido
        End Select
        MostraDados
        .PosMovimentacao
    End If
End With

End Sub

Public Sub Manutencoes(Botao As Integer)
' Esta rotina é resposavel pelo incio do processo de inclusão, pelo inicio do processo de alteração e por todo o precesso de exclusão

On Error GoTo ManutençãoErro

With Formulário
'    On Error GoTo ManutençãoErro
    Select Case Botao
        Case 1 ' Inclusão
            If .PermiteInclusao() = True Then
                LimpaTela
                GerenteBarra True
                .Registros.AddNew  ' Cria um Registro em Branco no final do Arquivo
                .PosInclusao
            End If
        Case 2 ' Alteração
            If .Registros.Eof = False And .Registros.Bof = False Then
                'permite editar um registro somente se voce não estiver no comerco ou no fim do arquivo
                If .PermiteAlteracao() Then
                     GerenteBarra True ' Ajusta os menus e a picture de fundo
                     .Registros.Edit  ' Permite Editar um registro ativo
                     .PosAlteracao
                End If
            End If
        Case 3 ' Exclusao
            If .Registros.Eof = False And .Registros.Bof = False Then
                'permite excluir o registro somente se voce não estiver no comerco ou no fim do arquivo
                If .PermiteExclusao() Then
                     .Registros.Delete  ' Permite Deletetar um registro ativo
                     .PosExclusao
                End If
            End If
    End Select
End With
Exit Sub
ManutençãoErro:
    MsgBox Err.Description & Err.Number
End Sub

Public Sub GerenteBarra(InclusaoEdicao As Boolean)
Dim X As Integer
' Está rotina irá gerenciar o estado da barra de ferramentas dos menus e da pctfundo dos formularios
' O parametro ediçaão exclusão será verdadeiro quando estivermos em modo de inclusão ou edição
Dim Componente As Variant
Dim Botao As Variant

With MDIPrincipal
    
    For Each Botao In .TlbPrincipal.Buttons
        If Botao.Index >= 14 Then 'Botao salvar e cancelar
            Botao.Visible = InclusaoEdicao
        Else
            Botao.Visible = Not InclusaoEdicao
        End If
        
    Next
'    .TlbPrincipal.Buttons(5).Visible = False
'    .TlbPrincipal.Buttons(10).Visible = False
    
'    .MnuArquivo(0).Enabled = Not InclusaoEdicao ' Novo Registro do Menu
'    .MnuArquivo(4).Enabled = InclusaoEdicao ' Salvar
'    .MnuArquivo(6).Enabled = Not InclusaoEdicao ' Relatório
'    .mnuEditar(0).Enabled = InclusaoEdicao  ' Cancelar
'   .mnuEditar(6).Enabled = Not InclusaoEdicao  ' Excluir
'    .mnuEditar(7).Enabled = Not InclusaoEdicao  ' Localizar
'    .MnuFerramentas(0).Enabled = Not InclusaoEdicao  ' Classificar
'    .MnuFerramentas(1).Enabled = Not InclusaoEdicao  ' Filtrar
    
With Formulário

'    For x = 0 To .PctFundo.Count - 1
'        .PctFundo(x).Enabled = InclusaoEdicao
'    Next
    
    For Each Componente In .PctFundo
        Componente.Enabled = InclusaoEdicao
    Next

End With

               
End With

End Sub

Public Sub Confirmacoes(Botao As Integer)

With Formulário
    .PreConfirmacao
    If Botao = 15 Then 'Botão Cancelar
        If .PermiteCancelamento() Then
            If .Registros.EditMode <> dbEditNone Then ' caso o usuário esteja fechando o formulário sem salvar
                .Registros.CancelUpdate 'Cancela o ultimo comando de inclusão ou edição
            End If
            Inicio
            .PosCancelamento
        End If
    Else
        If Salvar() = False Then ' a função irá tentar gravar os dados no banco
            Exit Sub
        End If
        .Registros.Bookmark = .Registros.LastModified
        ' BookMark identifica o nro físico do registro
        ' A propriedade LastModified é o numero fisico do ultimo registro que foi salvo
        MostraDados
    End If
    GerenteBarra False
    .PosConfirmacao
End With

End Sub

Public Function Salvar() As Boolean

On Error GoTo Campos

Dim Mascara As String
Dim Componente As Variant
With Formulário
    If .Registros.EditMode = dbEditNone Then
        .Registros.Edit
    End If

    If .PermiteSalvar() Then
        If .CTextBox Then
            For Each Componente In .TxtDados
                If Componente.Tag = "S" Then
                     If .Registros.Fields(Componente.DataField).Name = "CodPrd" Then
                        If .Registros.EditMode = 2 Then
                            .Registros.Fields(Componente.DataField) = Componente.Text
                        End If
                     Else
                        .Registros.Fields(Componente.DataField) = Componente.Text
                     End If
                End If
            Next
        End If
        If .CComboBox Then
            For Each Componente In .CmbDados
                If Mid(Componente.Tag, 1, 1) = "S" Then
                    If Mid(Componente.Tag, 2, 1) = "I" Then
                        .Registros.Fields(Componente.DataField) = Componente.ListIndex + 1
                    Else
                        .Registros.Fields(Componente.DataField) = Componente.Text
                    End If
                End If
            Next
        End If
            
        If .CMaskEditBox Then
            For Each Componente In .MebDados
'                If Componente.Tag = "S" Then
'                     .Registros.Fields(Componente.DataField) = Componente.Text
'               End If
 
        
                If .Registros.Fields(Componente.DataField).Type = dbDate Then
                    ' a propriedade type de um campo represente o tipo do campo dentro do banco de dados
                    If Not IsDate(Componente.Text) Then
                        .Registros.Fields(Componente.DataField) = Null
                    Else
                        .Registros.Fields(Componente.DataField) = Componente.Text
                    End If
                Else
                       .Registros.Fields(Componente.DataField) = Componente.Text
                End If
        
            Next
        
        End If
        
        If .CdbCombo Then
            For Each Componente In .DbcDados
                If Mid(Componente.Tag, 1, 1) = "S" Then
                    If Mid(Componente.Tag, 2, 1) = "I" Then
                         If IsNumeric(Componente.BoundText) Then
                            .Registros.Fields(Componente.DataField) = CDbl(Componente.BoundText)
                         End If
                    Else
                         .Registros.Fields(Componente.DataField) = Componente.BoundText
                    End If
                 End If
             Next
        End If
        
        If .CCheckBox Then
             For Each Componente In .ChkDados
                If Componente.Tag = "S" Then
                    .Registros.Fields(Componente.DataField) = IIf(Componente.value = 1, True, False)
                Else
                    .Registros.Fields(Componente.DataField) = IIf(Componente.value = 1, Left(Componente.Tag, 3), Right(Componente.Tag, 3))
                End If
             Next
        End If
        
        If .PermiteContinuarSalvar() Then
            .Registros.Update
            Salvar = True
            .PosSalvamento
        End If
    End If
End With
Exit Function
    
Campos:
'    If Err.Number <> 424 Then
'        MsgBox Err.Description & " no campo <<" & Componente.DataField & ">> Nro: " & Err.Number, vbCritical, Caminho
'    Else
        MsgBox Err.Description & Err.Number, vbCritical, Caminho
'    End If

End Function

Public Sub Inicio()

With Formulário

    If .Registros.RecordCount = 0 Then 'Se quantidade de registros = 0
        LimpaTela
    Else
        .Registros.MoveFirst
        MostraDados
    End If
    
End With

End Sub

Public Sub ListaCampos(Campos As Control)
'  Uma Variavel do tipo control pode representar (ponteiro) qualquer controle desenhado em um formulario
Dim NomeCampo As Variant
With Formulário
    If .CTextBox Then
        For Each NomeCampo In .TxtDados
            Campos.AddItem NomeCampo.DataField
        Next
    End If
    
    If .CMaskEditBox Then
        For Each NomeCampo In .MebDados
            Campos.AddItem NomeCampo.DataField
        Next
    End If
    
    If .CComboBox Then
        For Each NomeCampo In .CmbDados
            Campos.AddItem NomeCampo.DataField
        Next
    End If
    
    If .CCheckBox Then
        For Each NomeCampo In .ChkDados
            Campos.AddItem NomeCampo.DataField
        Next
    End If
    
    If .CdbCombo Then
        For Each NomeCampo In .DbcDados
            Campos.AddItem NomeCampo.DataField
        Next
    End If
    
End With

End Sub

Public Function Numero2(Valor As String) As String
Valor = Format(Valor, "0.00")
Numero2 = Left(Valor, Len(Valor) - 3) & "." & Right(Valor, 2)

End Function

Public Function Numero(Valor As String, Decimais As Byte) As String
If Decimais = 2 Then
    Valor = Format(Valor, "00.00")
    Numero = Left(Valor, Len(Valor) - 3) & "." & Right(Valor, 2)
Else
    Valor = Format(Valor, "0.000")
    Numero = Left(Valor, Len(Valor) - 4) & "." & Right(Valor, 3)
End If

End Function

Public Function Numero4(Valor As String) As String
Valor = Format(Valor, "0.0000")
Numero4 = Left(Valor, Len(Valor) - 5) & "." & Right(Valor, 4)

End Function

Public Function Numero3(Valor As String) As String
Valor = Format(Valor, "0.000")
Numero3 = Left(Valor, Len(Valor) - 4) & "." & Right(Valor, 3)

End Function



Public Sub CompactarReparar()
Dim Segurança As Workspace
Set Segurança = DBEngine.CreateWorkspace("Compactação", "admin", "")
' MDISistema.PctFundoMDI.Visible = True
' MDISistema.LblTitulo.Caption = "Compactação"
DoEvents
FrmCompac.Progresso.Max = 6
' MDISistema.MousePointer = vbHourglass
Banco.Close
FrmCompac.Progresso.value = 1
On Error GoTo ObCompact
Set Banco = Segurança.OpenDatabase(Caminho & "\" & NomeDoBanco, True)
On Error GoTo 0
Banco.Close
FrmCompac.Progresso.value = FrmCompac.Progresso.value + 1

DBEngine.RepairDatabase Caminho & "\" & NomeDoBanco
FrmCompac.Progresso.value = FrmCompac.Progresso.value + 1
DBEngine.CompactDatabase Caminho & "\" & NomeDoBanco, Caminho & "\Temporário.mdb"
FrmCompac.Progresso.value = FrmCompac.Progresso.value + 1
Kill Caminho & "\" & NomeDoBanco
FrmCompac.Progresso.value = FrmCompac.Progresso.value + 1
Name Caminho & "\Temporário.mdb" As Caminho & "\" & NomeDoBanco
FrmCompac.Progresso.value = FrmCompac.Progresso.value + 1

' SaveSetting App.Title, "Engines", "UltimaCompactação", Date
Segurança.Close
Set Segurança = Nothing
Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\" & NomeDoBanco)
Exit Sub
ObCompact:
    Segurança.Close
    Set Segurança = Nothing
    MsgBox Err.Description & Err.Number
    Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\" & NomeDoBanco)
    Exit Sub
End Sub

Public Function NameOfPC(MachineName As String) As Long
Dim NameSize As Long
Dim X As Long


MachineName = Space$(16)
NameSize = Len(MachineName)
X = GetComputerName(MachineName, NameSize)


End Function

Public Function Get_User_name(UserName As String) As Long
Dim lpBuff As String * 25
Dim ret As Long

ret = GetUserName(lpBuff, 25)
UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function

Function LimpaString(Campo As String) As String
Dim X As Integer
Dim CampoLimpo As String

For X = 1 To Len(Campo)
    If Mid$(Campo, X, 1) >= Chr(32) And Mid$(Campo, X, 1) <= Chr(122) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
    End If
Next
LimpaString = psCompString(3, CampoLimpo, Len(Campo))
End Function

Function AbilitaBotoes(Controle)

Dim Componente As Variant
Dim Botao As Variant

With MDIPrincipal
    
    For Each Botao In .TlbPrincipal.Buttons
        Botao.Enabled = Controle
    Next
               
End With

End Function

Public Sub Sair()
    
    'Dim Resp As String
    'Resp = MsgBox("Tem certeza que deseja finalizar o sistema ?", 36)
        
    'If Resp <> vbNo Then
        End
    'End If
    
    'obs.: abilitar linhas para mostrar mensagem
    
End Sub

Public Function Alinhar(Valor As String, Casas As Currency, Dec As Integer) As String
Dim VlrTmp As String
If IsNull(Valor) Then
    Valor = 0
End If
If Valor = "" Then
    Valor = 0
End If
If IsNull(Dec) Then
    Dec = 0
End If
If Dec = 2 Then
        VlrTmp = CStr(Format(CCur(Valor), "#,##0.00;-#,##0.00"))
ElseIf Dec = 3 Then
        VlrTmp = CStr(Format(CCur(Valor), "#,##0.000;-#,##0.000"))
    Else
        VlrTmp = CStr(Valor)
End If
If Dec <> 0 And (Casas - Len(Trim(VlrTmp)) & Trim(VlrTmp)) > 0 Then
    Alinhar = Space(Casas - Len(Trim(VlrTmp))) & Trim(VlrTmp)
Else
    Alinhar = VlrTmp
End If
'If Dec = 0 Then

End Function

Public Function Centralizar(Valor As String, Casas As Currency) As String
Dim VlrTmp As String
Dim Espacos As Currency

VlrTmp = Valor

Espacos = CInt((Casas - Len(Valor)) / 2)

If Espacos < 0 Then Espacos = 0

Centralizar = Space(Espacos) & VlrTmp

End Function

Public Sub ImprimeLinha(PulosLinha As Integer, PulosColuna As Integer, Valor As String, Tamanho As Currency, Condensado As Boolean, Dec As Integer)
Dim X
If IsNumeric(Valor) Then Valor = Alinhar(Trim(Valor), Tamanho, Dec)

If Valor = "" Then
    Valor = Space(Tamanho)
Else
    Valor = Left(Valor, Tamanho)
End If

If PulosLinha = 0 Then
    If Condensado = True Then
        Arq.Write Cond & Space(PulosColuna) & Mid(Valor, 1, Tamanho) & Space(Tamanho - Len(Valor)) & Nor
        'Arq.Write Cond & imprimeX(PulosColuna) & Mid(Valor, 1, Tamanho) & imprimeY(Tamanho - Len(Valor)) & Nor
    Else
        Arq.Write Space(PulosColuna) & Mid(Valor, 1, Tamanho) & Space(Tamanho - Len(Mid(Valor, 1, Tamanho))) & Nor
        'Arq.Write imprimeX(PulosColuna) & Mid(Valor, 1, Tamanho) & imprimeY(Tamanho - Len(Mid(Valor, 1, Tamanho))) & Nor
    End If
Else
    Arq.WriteBlankLines PulosLinha
End If

End Sub

Public Function imprimeX(Qtd As Integer) As String
Dim X As Currency
    For X = 1 To Qtd
        imprimeX = imprimeX & "A"
    Next
End Function

Public Function imprimeY(Qtd As Integer) As String
Dim X As Currency
    For X = 1 To Qtd
        imprimeY = imprimeY & "B"
    Next
End Function

Public Sub Iniciar(Objeto As Boolean)
Dim VarImpr1 As String
Dim VarImpr2 As String
Dim VarImpr3 As String

If Dir("C:\local.ini") <> "" Then
    If LerINI("IMPRESSORA", "Impr", "c:\local.ini") <> "" Then
        Porta = LerINI("IMPRESSORA", "Impr", "c:\local.ini")
    End If
Else
    If Porta = "" Then
        Set Configuracao = Banco.OpenRecordset("CFG")
        Porta = Configuracao("Porta")
'        If LerINI("IMPRESSORA", "Impr1", App.Path & "\config.ini") <> "" Then
'            VarImpr1 = LerINI("IMPRESSORA", "Impr1", App.Path & "\config.ini")
'        End If
'        If LerINI("IMPRESSORA", "Impr2", App.Path & "\config.ini") <> "" Then
'            VarImpr2 = LerINI("IMPRESSORA", "Impr2", App.Path & "\config.ini")
'        End If
'        If LerINI("IMPRESSORA", "Impr3", App.Path & "\config.ini") <> "" Then
'            VarImpr3 = LerINI("IMPRESSORA", "Impr3", App.Path & "\config.ini")
'        End If
    End If
End If

If Objeto Then
    Set Fs2 = CreateObject("Scripting.FileSystemObject")
    Set Arq = Fs2.CreateTextFile(Porta, True)
Else
    Open Porta For Output As #1
    Ant = 0
End If

End Sub

Public Sub Fim(Objeto As Boolean)

If Objeto Then
    Arq.Close
Else
    Close #1
End If

End Sub

Public Function RetornaMes(Mes As Integer) As String
Dim Nome
Nome = Array(" ", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", _
        "Junho", "Julho", "Agosto", "Setembro", "Outubro", _
        "Novembro", "Dezembro")
RetornaMes = Nome(Mes)

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

Public Sub Imprimir()
Dim X As Integer

MDIPrincipal.RptRel.RetrieveDataFiles

MDIPrincipal.RptRel.PrintReport

For X = 0 To 100
    MDIPrincipal.RptRel.Formulas(X) = ""
Next

End Sub

Public Sub BaixarComposicao(Baixa As Boolean, Operacao As Byte, Produto As Long, Qtd As Currency, LojaDesc As String)
Dim Criterio As String
Dim Prd  As Recordset
Dim Loja As Recordset
Dim Comp As Recordset
Dim SQL As String

Set Comp = Banco.OpenRecordset("Select * from Comp WHERE Produto = " & Produto & " Order By  Produto")

Comp.Requery
If Comp.RecordCount > 0 Then
    Comp.MoveLast
    Comp.MoveFirst
End If

While Not Comp.Eof
    If Operacao = 1 Or Operacao = 2 Or Operacao = 3 Or Operacao = 5 Or Operacao = 6 Or Operacao = 8 Or Operacao = 11 Or Operacao = 13 Then
        If Baixa = True Then
            Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda FROM Prd WHERE CodPrd = " & Comp!PrdComp)
            Prd.Requery
            If Prd.RecordCount > 0 Then
                Prd.Edit
                Prd!estoque = Prd!estoque - (Comp!Qtd * Qtd)
                Prd!DtVenda = Date
                Prd.Update
            End If
                        
            Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & Comp!PrdComp & " and DescricaoLoja = '" & LojaDesc & "'")
            Loja.Requery
            If Loja.RecordCount > 0 Then
                Loja.Edit
                Loja!estoque = Loja!estoque - (Comp!Qtd * Qtd)
                Loja.Update
            End If
        Else
            Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtCompra FROM Prd WHERE CodPrd = " & Comp!PrdComp)
            Prd.Requery
            If Prd.RecordCount > 0 Then
                Prd.Edit
                Prd!estoque = Prd!estoque + (Comp!Qtd * Qtd)
                Prd!DtCompra = Date
                Prd.Update
            End If
            
            Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & Comp!PrdComp & " and DescricaoLoja = '" & LojaDesc & "'")
            Loja.Requery
            If Loja.RecordCount > 0 Then
                Loja.Edit
                Loja!estoque = Loja!estoque + (Comp!Qtd * Qtd)
                Loja.Update
            End If
        End If
    End If
    Comp.MoveNext
Wend

End Sub

Public Sub BaixaEstoque(VarPedido As Double, Baixa As Boolean, Tipo As Byte)
Dim Prd As Recordset
Dim Loja As Recordset
Dim Pedido As Recordset

Dim VarQtd As Currency
Dim VarPrd As Long
Dim VarLoja As String
Dim VarBaixa As Boolean
Dim VarTemp As String
Dim Criterio  As String

Dim SQL As String

Set Pedido = Banco.OpenRecordset("SELECT ITP.*  FROM ITP Where Pedido = " & VarPedido & " Order by CodItp ")
Pedido.Requery
Pedido.MoveLast
Pedido.MoveFirst

'Pedido.Requery

'    Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, BaixarComp FROM Prd WHERE CodPrd = " & Pedido("Produto"))
    Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd Order By CodPrd ")
    Prd.Requery
    Prd.MoveLast
    Prd.MoveFirst
    
Set Loja = Banco.OpenRecordset("SELECT * FROM Loj  Order by Produto,  DescricaoLoja ")
Loja.Requery
Loja.MoveLast
Loja.MoveFirst

On Error GoTo ErroEstoque

If Pedido.RecordCount > 0 Then

    If Dir(Caminho & "\" & Trim(VarPedido) & ".txt") <> "" Then

            
        Open Caminho & "\" & Trim(VarPedido) & ".txt" For Input As #6
        Do
            Line Input #6, VarTemp
            VarQtd = Mid(VarTemp, 1, 10) / 100
            VarPrd = Mid(VarTemp, 11, 8)
            VarLoja = Mid(VarTemp, 19, 20)
            VarBaixa = Mid(VarTemp, 39, 1)
            If VarBaixa = True Then
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, BaixarComp FROM Prd WHERE CodPrd = " & VarPrd)
'                Prd.Requery
                Criterio = "CodPrd = " & VarPrd
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    Prd!estoque = Prd!estoque - VarQtd
                    Prd!DtVenda = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If

'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & VarPrd & " and DescricaoLoja = '" & VarLoja & "'")
'                Loja.Requery
                Criterio = "Produto = " & VarPrd & " and DescricaoLoja = '" & Trim(VarLoja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque - VarQtd
                    Loja.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Loja - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                
                End If
            Else
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd WHERE CodPrd = " & VarPrd)
'                Prd.Requery
'                If Prd.RecordCount > 0 Then
                Criterio = "CodPrd = " & VarPrd
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    Prd!estoque = Prd!estoque + VarQtd
                    Prd!DtCompra = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
                            
'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & VarPrd & " and DescricaoLoja = '" & VarLoja & "'")
'                Loja.Requery
'                If Loja.RecordCount > 0 Then
                Criterio = "Produto = " & VarPrd & " and DescricaoLoja = '" & Trim(VarLoja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque + VarQtd
                    Loja.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Loja - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1

                End If
            End If
'            BaixarEstoque VarQtd, VarPrd, VarLoja, VarBaixa
            
        Loop Until Eof(6)
        Close #6
        Kill App.Path & "\" & Trim(VarPedido) & ".txt"
            
    End If


    Pedido.MoveFirst
    While Not Pedido.Eof
        If Pedido!Entregue = True Then
            If Baixa = True Then
                Criterio = "CodPrd = " & Pedido("Produto")
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    Prd!estoque = Prd!estoque - Pedido!Qtd
                    Prd!DtVenda = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Baixa; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido!Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
                            
'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & Pedido("Produto") & " and DescricaoLoja = '" & Pedido!Loja & "'")
'                Loja.Requery
'                If Loja.RecordCount > 0 Then
                Criterio = "Produto = " & Pedido("Produto") & " and DescricaoLoja = '" & Trim(Pedido!Loja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque - Pedido!Qtd
                    Loja.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
'                    Print #1, "Loja; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido.Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Print #1, "Loja - Vendas - Baixa; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido!Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
            Else
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd WHERE CodPrd = " & Pedido("Produto"))
'                Prd.Requery
'                If Prd.RecordCount > 0 Then
                Criterio = "CodPrd = " & Pedido("Produto")
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    If Tipo = 0 And Prd!estoque > 0 Then
                        If Prd!CustoMedio <> 0 Then
                            Prd!CustoMedio = (((Prd!CustoMedio * Prd!estoque) + (Pedido!CustoDia * Pedido!Qtd)) / (Prd!estoque + Pedido!Qtd))
                        Else
                            Prd!CustoMedio = (((Prd!Custo * Prd!estoque) + (Pedido!CustoDia * Pedido!Qtd)) / (Prd!estoque + Pedido!Qtd))
                        End If
                    End If
                    Prd!estoque = Prd!estoque + Pedido!Qtd
                    Prd!DtCompra = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Baixa; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido!Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
                            
'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & Pedido("Produto") & " and DescricaoLoja = '" & Pedido!Loja & "'")
'                Loja.Requery
'                If Loja.RecordCount > 0 Then
                Criterio = "Produto = " & Pedido("Produto") & " and DescricaoLoja = '" & Trim(Pedido!Loja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque + Pedido!Qtd
                    Loja.Update
                Else
                    If Tipo = 0 Then
                        Loja.AddNew
                        Loja.Fields("Produto") = Pedido("Produto")
                        Loja.Fields("DescricaoLoja") = Pedido("Loja")
                        Loja.Fields("Estoque") = Pedido("Qtd")
                        Loja.Update
                    End If
                    Open Caminho & "\Estoque.txt" For Append As #1
'                    Print #1, "Loja; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido.Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Print #1, "Loja - Vendas - Baixa; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido!Loja & "; " & Date & "; " & Time & "; " & Maquina
                    
                    Close #1

                End If
            End If
            If Prd!BaixarComp = True Then
                BaixarComposicao Baixa, Tipo, Prd("CodPrd"), Pedido!Qtd, Pedido!Loja
            End If
        End If
        Pedido.MoveNext
    Wend
Else
    MsgBox "Arquivo de estoque vazio", vbInformation, App.Title
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, "Erro Estoque modulo Principal " & VarPedido & "; " & Pedido.RecordCount & ";" & Date & "; " & Time & "; " & Maquina
    Close #1
End If

Exit Sub

ErroEstoque:

    MsgBox Err.Description & Err.Number, vbCritical, Caminho
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, Err.Description & Err.Number, vbCritical, Caminho
    Close #1

End Sub


Public Sub BaixaEstoqueTeste(VarPedido As Double, Baixa As Boolean, Tipo As Byte)
Dim VarEstoque As Recordset
Dim Prd As Recordset
Dim Loja As Recordset
Dim Pedido As Recordset

Dim VarQtd As Currency
Dim VarPrd As Long
Dim VarLoja As String
Dim VarBaixa As Boolean
Dim VarTemp As String
Dim Criterio  As String

Dim SQL As String

Set VarEstoque = Banco.OpenRecordset("SELECT ITP.Pedido, ITP.Produto, ITP.Qtd, PRD.Estoque, LOJ.Estoque, PRD.DtVenda, PRD.DtCompra ,LOJ.DescricaoLoja, Itp.Entregue FROM (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) INNER JOIN LOJ ON PRD.CodPrd = LOJ.Produto WHERE Pedido = " & VarPedido & " Order by CodItp ")
VarEstoque.Requery
VarEstoque.MoveLast
VarEstoque.MoveFirst

Set Pedido = Banco.OpenRecordset("SELECT ITP.*  FROM ITP Where Pedido = " & VarPedido & " Order by CodItp ")
Pedido.Requery
'Pedido.Requery

'    Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, BaixarComp FROM Prd WHERE CodPrd = " & Pedido("Produto"))
    Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd Order By CodPrd ")
    Prd.Requery
    Prd.MoveLast
    Prd.MoveFirst
    
Set Loja = Banco.OpenRecordset("SELECT * FROM Loj  Order by Produto,  DescricaoLoja ")
Loja.Requery
Loja.MoveLast
Loja.MoveFirst



'On Error GoTo ErroEstoque

If Pedido.RecordCount > 0 Then

    If Dir(Caminho & "\" & Trim(VarPedido) & ".txt") <> "" Then

            
        Open Caminho & "\" & Trim(VarPedido) & ".txt" For Input As #6
        Do
            Line Input #6, VarTemp
            VarQtd = Mid(VarTemp, 1, 10) / 100
            VarPrd = Mid(VarTemp, 11, 8)
            VarLoja = Mid(VarTemp, 19, 20)
            VarBaixa = Mid(VarTemp, 39, 1)
            If VarBaixa = True Then
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda, BaixarComp FROM Prd WHERE CodPrd = " & VarPrd)
'                Prd.Requery
                Criterio = "CodPrd = " & VarPrd
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    Prd!estoque = Prd!estoque - VarQtd
                    Prd!DtVenda = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If

'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & VarPrd & " and DescricaoLoja = '" & VarLoja & "'")
'                Loja.Requery
                Criterio = "Produto = " & VarPrd & " and DescricaoLoja = '" & Trim(VarLoja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque - VarQtd
                    Loja.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Loja - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                
                End If
            Else
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd WHERE CodPrd = " & VarPrd)
'                Prd.Requery
'                If Prd.RecordCount > 0 Then
                Criterio = "CodPrd = " & VarPrd
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    Prd!estoque = Prd!estoque + VarQtd
                    Prd!DtCompra = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
                            
'                Set Loja = Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & VarPrd & " and DescricaoLoja = '" & VarLoja & "'")
'                Loja.Requery
'                If Loja.RecordCount > 0 Then
                Criterio = "Produto = " & VarPrd & " and DescricaoLoja = '" & Trim(VarLoja) & "'"
                Loja.FindFirst Criterio
                If Not Loja.NoMatch Then
                    Loja.Edit
                    Loja!estoque = Loja!estoque + VarQtd
                    Loja.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Loja - Vendas - Volta; " & VarPrd & "; " & VarQtd & "; " & VarLoja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1

                End If
            End If
'            BaixarEstoque VarQtd, VarPrd, VarLoja, VarBaixa
            
        Loop Until Eof(6)
        Close #6
        Kill App.Path & "\" & Trim(VarPedido) & ".txt"
            
    End If


    VarEstoque.MoveFirst
    While Not VarEstoque.Eof
        If VarEstoque!Entregue = True Then
            If Baixa = True Then
'                    VarEstoque.Edit
                    VarEstoque("Prd.Estoque") = VarEstoque("Prd.Estoque") - VarEstoque("Qtd")
                    VarEstoque("Loja.Estoque") = VarEstoque("Loja.Estoque") - VarEstoque("Itp.Qtd")
                    VarEstoque("DtVenda") = Date
                    VarEstoque.Update
            Else
'                Set Prd = Banco.OpenRecordset("Select Estoque, CodPrd, DtCompra, CustoMedio, Custo, BaixarComp FROM Prd WHERE CodPrd = " & Pedido("Produto"))
'                Prd.Requery
'                If Prd.RecordCount > 0 Then
                
                VarEstoque.Edit
                VarEstoque("Prd.Estoque") = VarEstoque("Prd.Estoque") + VarEstoque("Itp.Qtd")
                VarEstoque("Loja.Estoque") = VarEstoque("Loja.Estoque") + VarEstoque("Itp.Qtd")
                VarEstoque("DtVenda") = Date
                VarEstoque.Update
                
                Criterio = "CodPrd = " & Pedido("Produto")
                Prd.FindFirst Criterio
                If Not Prd.NoMatch Then
                    Prd.Edit
                    If Tipo = 0 And Prd!estoque > 0 Then
                        If Prd!CustoMedio <> 0 Then
                            Prd!CustoMedio = (((Prd!CustoMedio * Prd!estoque) + (Pedido!CustoDia * Pedido!Qtd)) / (Prd!estoque + Pedido!Qtd))
                        Else
                            Prd!CustoMedio = (((Prd!Custo * Prd!estoque) + (Pedido!CustoDia * Pedido!Qtd)) / (Prd!estoque + Pedido!Qtd))
                        End If
                    End If
                    
'                    Prd!Estoque = Prd!Estoque + Pedido!Qtd
'                    Prd!DtCompra = Date
                    Prd.Update
                Else
                    Open Caminho & "\Estoque.txt" For Append As #1
                    Print #1, "Produto - Vendas - Baixa; " & Pedido!Produto & "; " & Pedido!Qtd & "; " & Pedido!Loja & "; " & Date & "; " & Time & "; " & Maquina
                    Close #1
                End If
                            
            End If
            If Prd!BaixarComp = True Then
                BaixarComposicao Baixa, Tipo, Prd("CodPrd"), Pedido!Qtd, Pedido!Loja
            End If
        End If
        VarEstoque.MoveNext
    Wend
Else
    MsgBox "Arquivo de estoque vazio", vbInformation, App.Title
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, "Erro Estoque modulo Principal " & VarPedido & "; " & Pedido.RecordCount & ";" & Date & "; " & Time & "; " & Maquina
    Close #1
End If

Exit Sub

ErroEstoque:

    MsgBox Err.Description & Err.Number, vbCritical, Caminho
    Open Caminho & "\Estoque.txt" For Append As #1
    Print #1, Err.Description & Err.Number, vbCritical, Caminho
    Close #1

End Sub


Function LIMPACampo(Campo As String) As String
Dim X As Integer
Dim CampoLimpo As String

For X = 1 To Len(Campo)
    If Mid$(Campo, X, 1) >= Chr(32) And Mid$(Campo, X, 1) <= Chr(122) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
    End If
Next
LIMPACampo = psCompString(3, CampoLimpo, Len(Campo))
End Function

Function LIMPACampo2(Campo As String) As String
Dim X As Integer
Dim CampoLimpo As String

For X = 1 To Len(Campo)
    If Mid$(Campo, X, 1) <> "." And Mid$(Campo, X, 1) <> "/" And Mid$(Campo, X, 1) <> "-" And Mid$(Campo, X, 1) <> Chr(32) And Mid$(Campo, X, 1) <> "(" And Mid$(Campo, X, 1) <> ")" Then
            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
    End If
Next
LIMPACampo2 = psCompString(3, CampoLimpo, Len(Campo))
End Function

Public Function PassaExtenso(Valor As String) As String
  
  ' Passa um número para a DLL e
  ' recebe-o de volta por extenso

  On Error GoTo Passa_Err

  Dim Retorno As String, X%
  Retorno = Space$(512)
  X% = Extenso(Valor, Retorno)
    Retorno = Replace(Retorno, "à", "a")
    Retorno = Replace(Retorno, "á", "a")
    Retorno = Replace(Retorno, "ã", "a")
    Retorno = Replace(Retorno, "ê", "e")
    Retorno = Replace(Retorno, "é", "e")
    Retorno = Replace(Retorno, "í", "i")
    Retorno = Replace(Retorno, "ó", "o")
    Retorno = Replace(Retorno, "õ", "o")
    Retorno = Replace(Retorno, "ô", "o")
    Retorno = Replace(Retorno, "ú", "u")
    Retorno = Replace(Retorno, "ñ", "n")
    Retorno = Replace(Retorno, "ç", "c")

    Retorno = Replace(Retorno, "À", "A")
    Retorno = Replace(Retorno, "Á", "A")
    Retorno = Replace(Retorno, "Ã", "A")
    Retorno = Replace(Retorno, "Ê", "E")
    Retorno = Replace(Retorno, "É", "E")
    Retorno = Replace(Retorno, "Í", "I")
    Retorno = Replace(Retorno, "Ó", "O")
    Retorno = Replace(Retorno, "Õ", "O")
    Retorno = Replace(Retorno, "Ô", "O")
    Retorno = Replace(Retorno, "Ú", "U")
    Retorno = Replace(Retorno, "Ñ", "N")
    Retorno = Replace(Retorno, "Ç", "C")
    Retorno = Replace(Retorno, "ü", "u")
    Retorno = Replace(Retorno, "Ü", "U")
    
    PassaExtenso = Trim$(Retorno)

Passa_Fim:
  Exit Function
Passa_Err:
  MsgBox Error$(Err)
  Resume Passa_Fim
End Function


Public Sub CentralizarForm(FormX As Form, Tipo As Currency)
    
    If Tipo = 1 Then
        Screen.MousePointer = 11
        FormX.Top = ((Screen.Height) / 3 - FormX.Height / 2) + MDIPrincipal.TlbPrincipal.Height + 120
        FormX.Left = Screen.Width / 2 - FormX.Width / 2
        Screen.MousePointer = 0
    ElseIf Tipo = 2 Then
        Screen.MousePointer = 11
        'FormX.Top = ((Screen.Height) / 3 - FormX.Height / 2)
        FormX.Top = 100
        FormX.Left = Screen.Width / 2 - FormX.Width / 2
        Screen.MousePointer = 0
    End If
End Sub

Public Function DataExtenso(Data As Date) As String

Dim Mes As String

If Month(Date) = 1 Then
 Mes = "Janeiro"
ElseIf Month(Data) = 2 Then
 Mes = "Fevereiro"
ElseIf Month(Data) = 3 Then
 Mes = "Março"
ElseIf Month(Data) = 4 Then
 Mes = "Abril"
ElseIf Month(Data) = 5 Then
 Mes = "Maio"
ElseIf Month(Data) = 6 Then
 Mes = "Junho"
ElseIf Month(Data) = 7 Then
 Mes = "Julho"
ElseIf Month(Data) = 8 Then
 Mes = "Agosto"
ElseIf Month(Data) = 9 Then
 Mes = "Setembro"
ElseIf Month(Data) = 10 Then
 Mes = "Outubro"
ElseIf Month(Data) = 11 Then
 Mes = "Novembro"
ElseIf Month(Data) = 12 Then
 Mes = "Dezembro"
End If

    DataExtenso = Day(Data) & " de " & Mes & " de " & Year(Data)
    
End Function

Public Function LerINI(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim ret As String
 ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", ret, Len(ret), Arquivo)
 ret = Left$(ret, retlen)
 LerINI = ret
End Function

Public Sub EscreveINI(Secao As String, Entrada As String, Texto As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
  'texto= valor que vem depois do igual
  WritePrivateProfileString Secao, Entrada, Texto, Arquivo
End Sub

Public Sub SelecionaCampo()
On Error GoTo Linha
    With Screen.ActiveControl
        'If .Name = "TxtDados" Or .Name = "MebDados" Then
            .SelStart = 0
            .SelLength = Len(Screen.ActiveControl.Text) + 20
        'End If
    End With
Linha:
End Sub

Public Function ConsisteEstoque(Pedido As Double)
Dim VarBcoVda As Recordset

Set VarBcoVda = Banco.OpenRecordset("SELECT ITP.Pedido, Sum(PRD.Estoque) AS SomaDeEstoque FROM PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto GROUP BY ITP.Pedido HAVING ITP.Pedido= " & Pedido)
VarBcoVda.Requery

If VarBcoVda.RecordCount > 0 Then
    ConsisteEstoque = VarBcoVda!somadeestoque
End If

End Function

Public Function TraduzMes(Mes As Integer) As String
Select Case Mes
    Case 1: TraduzMes = "Janeiro"
    Case 2: TraduzMes = "Fevereiro"
    Case 3: TraduzMes = "Março"
    Case 4: TraduzMes = "Abril"
    Case 5: TraduzMes = "Maio"
    Case 6: TraduzMes = "Junho"
    Case 7: TraduzMes = "Julho"
    Case 8: TraduzMes = "Agosto"
    Case 9: TraduzMes = "Setembro"
    Case 10: TraduzMes = "Outubro"
    Case 11: TraduzMes = "Novembro"
    Case 12: TraduzMes = "Dezembro"
End Select
End Function

Public Function CarregaTipo()
Dim Temp As String
With Formulário

    If Dir(App.Path & "\Tipo.Txt") <> "" Then
        On Error GoTo Erro
        .CmbTipo.Clear
        Open App.Path & "\Tipo.Txt" For Input As #1
            Do
                Input #1, Temp
                .CmbTipo.AddItem Temp
            Loop Until Temp = ""
        Close #1
    End If

End With

Exit Function

Erro:
If Err.Number = 62 Then
    Close #1
End If

End Function

Public Sub ImprimiPromiMat(CodPed As Double)
Dim VarDados As Recordset

Set VarDados = Banco.OpenRecordset("")

'                If Parcelas.RecordCount > 0 Then
'                    FormPed.SecPromissoria.Mostrar = True
'                    FormPed.NroPedido.Caption = Format(Datdados(5).Recordset("CodPed"), "000000")
'                    FormPed.ExtensoDia.Caption = Day(Parcelas("DtVenc")) & " DIAS DO MÊS DE " & UCase(RetornaMes(Month(Parcelas("DtVenc")))) & " DO ANO DE " & Year(Parcelas("DtVenc"))
'                    FormPed.QuantiaExtenso.Caption = UCase(PassaExtenso(Datdados(5).Recordset("Total"))) & " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
'                    FormPed.ValorPromissoria.Caption = Format(Datdados(5).Recordset("Total"), "#,#0.00")
'                    FormPed.DiaVenc.Caption = Day(Parcelas("DtVenc"))
'                    FormPed.MesVenc.Caption = UCase(RetornaMes(Month(Parcelas("DtVenc"))))
'                    FormPed.AnoVenc.Caption = Year(Parcelas("DtVenc"))
'                    FormPed.Emitente.Caption = varPesquisa(1).Fields("Nome")
'                    FormPed.Emitente2.Caption = varPesquisa(1).Fields("Nome")
'                    If Not IsNull(varPesquisa(1).Fields("Cpf")) Then
'                        FormPed.EmitenteCpf.Caption = varPesquisa(1).Fields("CPF")
'                    End If
'                    FormPed.Pagavel.Caption = CidadeEmpresa
'                    FormPed.EmitenteEndereco.Caption = varPesquisa(1).Fields("Endereco") & " - " & varPesquisa(1).Fields("Nro")
'                    FormPed.NomeEmpresa.Caption = NomeInscr
'                    FormPed.EmpresaCGCCPF.Caption = Cnpj
'                    FormPed.RpDia.Caption = Day(Date)
'                    FormPed.RpMes.Caption = UCase(RetornaMes(Month(Date)))
'                    FormPed.RpAno.Caption = Year(Date)
'                Else
'                    MsgBox "Não foi gerado financeiro para este pedido", vbInformation, App.Title
'                End If
'            End If


End Sub

Function LimpaNro(Numero) As String
Dim X As Integer
Dim Resp As String
Resp = ""
For X = 1 To Len(Numero)
    If Mid$(Numero, X, 1) <> "," And Mid$(Numero, X, 1) <> "." Then
        Resp = Resp + Mid$(Numero, X, 1)
    End If
Next
LimpaNro = Resp
End Function

Function AtivaFundo(Ativar)

'MDIPrincipal.Toolbar2.Visible = Ativar
'MDIPrincipal.Picture1.Visible = Ativar
'MDIPrincipal.Picture2.Visible = Ativar


End Function


Public Function valorextenso(nvalor)
'Valida Argumento
If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then
Exit Function
End If

'Variáveis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), atexto(4) As String

'Matrizes de extensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nvalor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
  cParte = aGrupo(nContador)
  nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
  If nTamanho = 3 Then
    If Right$(cParte, 2) <> "00" Then
      atexto(nContador) = atexto(nContador) + aCentena(Left(cParte, 1)) + "e "
      nTamanho = 2
    Else
      atexto(nContador) = atexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", _
      aCentena(Left(cParte, 1)))
    End If
  End If
  If nTamanho = 2 Then
    If Val(Right(cParte, 2)) < 20 Then
      atexto(nContador) = atexto(nContador) + aUnid(Right(cParte, 2))
    Else
      atexto(nContador) = atexto(nContador) + aDezena(Mid(cParte, 2, 1))
      If Right$(cParte, 1) <> "0" Then
        atexto(nContador) = atexto(nContador) + "e "
        nTamanho = 1
      End If
    End If
  End If
  If nTamanho = 1 Then
    atexto(nContador) = atexto(nContador) + aUnid(Right(cParte, 1))
  End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
  cFinal = atexto(4) + IIf(Val(aGrupo(4)) = 1, "", "")
Else
  cFinal = ""
  cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, atexto(1) + IIf(Val(aGrupo(1)) > 1, _
  "milhões ", "milhão "), "")
  If Val(aGrupo(2) + aGrupo(3)) = 0 Then
    cFinal = cFinal + "de "
  Else
    cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, atexto(2) + "mil ", "")
  End If
  cFinal = cFinal + atexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "", _
  "")
  cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + atexto(4) + IIf(Val(aGrupo(4)) = 1, _
 "", ""), "")
End If
valorextenso = UCase$(cFinal)

End Function



Function AbreAcbr(Comando As String)
Dim X As Double

If Dir(PortaAcbr & "\Sai.txt") <> "" Then

    Kill PortaAcbr & "\Sai.txt"
End If

Open App.Path & "\Ent.Txt" For Output As #1
Print #1, Comando
Close #1

 FileCopy App.Path & "\Ent.Txt", PortaAcbr & "\Tmp.txt"
 Kill App.Path & "\Ent.Txt"
 Name PortaAcbr & "\Tmp.Txt" As PortaAcbr & "\Ent.txt"
 'Kill PortaAcbr & "\Tmp.Txt"

      
End Function

Function TrataAcbr(Operacao As String, Numero As String) As String
Dim VarAcbr As String
Dim VarTemp As String
Dim X As Double
Dim Var
Dim CodigoRet As String
Dim VarErro As String
On Error Resume Next


'DoEvents

X = 0
If Dir(PortaAcbr & "\Sai.txt") = "" Then

    While FileLen(PortaAcbr & "\Sai.txt") = 0

    Wend

End If

X = 0

    For X = 1 To 700000


    Next

X = 0
    For X = 1 To 700000
'    FrmFrente.Text2 = X

    Next


X = 0

If Dir(PortaAcbr & "\Sai.txt") = "" Then
    For X = 1 To 700000


    Next
Else
    For X = 1 To 700000

    Next
End If
                                                                                                                     
Open PortaAcbr & "\Sai.txt" For Input As #5

    Input #5, VarTemp
    CodigoRet = VarTemp
'    If Not Eof(5) Then
'        Input #5, varTemp
        If Left(CodigoRet, 2) = "OK" Then
            Select Case Operacao
              Case "Nota"
                   VarAcbr = LerINI("Envio", "cStat", PortaAcbr & "\Sai.txt")
                   If VarAcbr = "100" Then
                        VarChaveNfe = LerINI("Nfe" & Trim(Numero), "chDFe", PortaAcbr & "\Sai.txt")
'                        VarChaveNfe = Mid(varTemp, Len(CodigoRet) - 51, 44)
                   ElseIf VarAcbr = "104" Then
                            If LerINI("Nfe" & Trim(Numero), "cStat", PortaAcbr & "\Sai.txt") = "100" Then
                                VarChaveNfe = LerINI("Nfe" & Trim(Numero), "chDFe", PortaAcbr & "\Sai.txt")
                                VarAcbr = "100"
                            Else
                                MsgBox LerINI("Nfe" & Trim(Numero), "XMotivo", PortaAcbr & "\Sai.txt")
                            End If
                       Else
                            MsgBox LerINI("Envio", "XMotivo", PortaAcbr & "\Sai.txt")
                   End If
               Case "Visualiza"
                     VarChaveNfe = Mid(VarTemp, Len(CodigoRet) - 51, 44)
               Case "Cancela"
                       VarAcbr = LerINI("Cancelamento", "cStat", PortaAcbr & "\Sai.txt")
                       MsgBox LerINI("Cancelamento", "XMotivo", PortaAcbr & "\Sai.txt")
               Case "Carta"
                        VarAcbr = LerINI("Evento001", "cStat", PortaAcbr & "\Sai.txt")
                        MsgBox LerINI("Evento001", "XMotivo", PortaAcbr & "\Sai.txt")
               Case "Consulta"
                        VarAcbr = LerINI("Consulta", "cStat", PortaAcbr & "\Sai.txt")
                        MsgBox LerINI("Consulta", "XMotivo", PortaAcbr & "\Sai.txt")
               Case "Certificado"
                        FrmNota.Label41.Caption = "Vencimento do Certificado: " & Right(VarTemp, 10)
            End Select

        End If
                
Close #5

'If Left(CodigoRet, 8) <> "OK: Lote" Then
'    VarChaveNfe = Mid(varTemp, Len(CodigoRet) - 51, 44)
'Else
    If Left(CodigoRet, 2) = "OK" Then
        TrataAcbr = VarAcbr
    Else
    '    If VarAcbr <> "," Then
            MsgBox CodigoRet & " - " & VarErro, vbInformation, App.Title
            TrataAcbr = 0
    '    End If
    End If
'End If

' RetornoSat CodigoRet
If Dir(PortaAcbr & "\Sai.txt") <> "" Then

    Kill PortaAcbr & "\Sai.txt"
End If

'For X = 1 To 100000

'Next

End Function


Function RetornoSat(Codigo As String)
Dim Result As String
'ResultSat = Codigo
Result = ""

Select Case Codigo
    Case 4000:      Result = "Ativado corretamente SAT Ativado com Sucesso."
    Case 4001:      Result = "Erro na criação do certificado processo de ativação foi interrompido."
    Case 4002:      Result = "SEFAZ não reconhece este SAT (CNPJ inválido) Verificar junto a SEFAZ o CNPJ cadastrado."
    Case 4003:      Result = "SAT já ativado SAT disponível para uso."
    Case 4004:      Result = "SAT com uso cessado SAT bloqueado por cessação de uso."
    Case 4005:      Result = "Erro de comunicação com a SEFAZ Tentar novamente."
    Case 4006:      Result = "CSR ICP-BRASIL criado com sucesso Processo de criação do CSR para certificação ICP-BRASIL com sucesso"
    Case 4007:      Result = "Erro na criação do CSR ICP-BRASIL Processo de criação do CSR para certificação ICP-BRASIL com erro"
    Case 4098:      Result = "SAT em processamento. Tente novamente."
    Case 4099:      Result = "Erro desconhecido na ativação Informar ao administrador."
    Case 5000:      Result = "Certificado transmitido com Sucesso "
    Case 5001:      Result = "Código de ativação inválido."
    Case 5002:      Result = "Erro de comunicação com a SEFAZ. Tentar novamente."
    Case 5003:      Result = "Certificado Inválido "
    Case 5098:      Result = "SAT em processamento."
    Case 5099:      Result = "Erro desconhecido Informar o administrador."
    Case 6000:      Result = "Emitido com sucesso + conteúdo notas. Retorno CF-e-SAT ao AC para contingência."
    Case 6001:      Result = "Código de ativação inválido."
    Case 6002:      Result = "SAT ainda não ativado. Efetuar ativação."
    Case 6003:      Result = "SAT não vinculado ao AC Efetuar vinculação"
    Case 6004:      Result = "Vinculação do AC não confere Efetuar vinculação"
    Case 6005:      Result = "Tamanho do CF-e-SAT superior a 1.500KB"
    Case 6006:      Result = "SAT bloqueado pelo contribuinte"
    Case 6007:      Result = "SAT bloqueado pela SEFAZ"
    Case 6008:      Result = "SAT bloqueado por falta de comunicação"
    Case 6009:      Result = "SAT bloqueado, código de ativação incorreto"
    Case 6010:      Result = "Erro de validação do conteúdo."
    Case 6098:      Result = "SAT em processamento."
    Case 6099:      Result = "Erro desconhecido na emissão. Informar o administrador."
    Case 7000:      Result = "Cupom cancelado com sucesso + conteúdo CF-eSAT cancelado."
    Case 7001:      Result = "Código ativação inválido Verificar o código e tentar mais uma vez."
    Case 7002:      Result = "Cupom inválido Informar o administrador."
    Case 7003:      Result = "SAT bloqueado pelo contribuinte"
    Case 7004:      Result = "SAT bloqueado pela SEFAZ"
    Case 7005:      Result = "SAT bloqueado por falta de comunicação"
    Case 7006:      Result = "SAT bloqueado, código de ativação incorreto"
    Case 7007:      Result = "Erro de validação do conteúdo"
    Case 7098:      Result = "SAT em processamento."
    Case 7099:      Result = "Erro desconhecido no cancelamento."
    Case 8000:      Result = "SAT em operação. Verifica se o SAT está ativo."
    Case 8098:      Result = "SAT em processamento."
    Case 8099:      Result = "Erro desconhecido. Informar o administrador."
    Case 9000:      Result = "Emitido com sucesso Gera e envia um cupom de teste para SEFAZ, para verificar a comunicação."
    Case 9001:      Result = "código ativação inválido Verificar o código e tentar mais uma vez."
    Case 9002:      Result = "SAT ainda não ativado. Efetuar ativação "
    Case 9098:      Result = "SAT em processamento."
    Case 9099:      Result = "Erro desconhecido Informar o "
    Case 10000:     Result = "Resposta com Sucesso. Informações de status do SAT."
    Case 10001:     Result = "Código de ativação inválido"
    Case 10098:     Result = "SAT em processamento."
    Case 10099:     Result = "Erro desconhecido Informar o administrador."
    Case 11000:     Result = "Emitido com sucesso Retorna o conteúdo do CF-ao AC."
    Case 11001:     Result = "código ativação inválido Verificar o código e tentar mais uma vez."
    Case 11002:     Result = "SAT ainda não ativado. Efetuar ativação."
    Case 11003:     Result = "Sessão não existe. AC deve executar a sessão novamente."
    Case 11098:     Result = "SAT em processamento."
    Case 11099:     Result = "Erro desconhecido. Informar o administrador."
    Case 12000:     Result = "Rede Configurada com Sucesso"
    Case 12001:     Result = "código ativação inválido Verificar o código e tentar mais uma vez."
    Case 12002:     Result = "Dados fora do padrão a ser informado Corrigir dados"
    Case 12098:     Result = "SAT em processamento."
    Case 12099:     Result = "Erro desconhecido Informar o administrador."
    Case 13000:     Result = "Assinatura do AC"
    Case 13001:     Result = "código ativação inválido Verificar o código e tentar mais uma vez."
    Case 13002:     Result = "Erro de comunicação com a SEFAZ"
    Case 13003:     Result = "Assinatura fora do padrão informado Corrigir dados"
    Case 13004:     Result = "CNPJ da Software House + CNPJ do emitente assinado no campo “signAC” difere do informado no campo “CNPJvalue” Corrigir dados"
    Case 13098:     Result = "SAT em processamento."
    Case 13099:     Result = "Erro desconhecido Informar o administrador."
    Case 14000:     Result = "Software Atualizado com Sucesso "
    Case 14001:     Result = "Código de ativação inválido."
    Case 14002:     Result = "Atualização em Andamento"
    Case 14003:     Result = "Erro na atualização Não foi possível Atualizar o SAT."
    Case 14004:     Result = "Arquivo de atualização inválido"
    Case 14098:     Result = "SAT em processamento."
    Case 14099:     Result = "Erro desconhecido Informar o administrador."
    Case 15000:     Result = "Transferência completa Arquivos de Logs extraídos"
    Case 15001:     Result = "Código de ativação inválido."
    Case 15002:     Result = "Transferência em andamento"
    Case 15098:     Result = "SAT em processamento."
    Case 15099:     Result = "Erro desconhecido Informar o administrador."
    Case 16000:     Result = "Equipamento SAT bloqueado com sucesso."
    Case 16001:     Result = "Código de ativação inválido."
    Case 16002:     Result = "Equipamento SAT já está bloqueado."
    Case 16003:     Result = "Erro de comunicação com a SEFAZ"
    Case 16004:     Result = "Não existe parametrização de bloqueio disponível."
    Case 16098:     Result = "SAT em processamento."
    Case 16099:     Result = "Erro desconhecido Informar o administrador."
    Case 17000:     Result = "Equipamento SAT desbloqueado com sucesso."
    Case 17001:     Result = "Código de ativação inválido."
    Case 17002:     Result = "SAT bloqueado pelo contribuinte. Verifique configurações na SEFAZ"
    Case 17003:     Result = "SAT bloqueado pela SEFAZ"
    Case 17004:     Result = "Erro de comunicação com a SEFAZ"
    Case 17098:     Result = "SAT em processamento."
    Case 17099:     Result = "Erro desconhecido Informar o administrador."
    Case 18000:     Result = "Código de ativação alterado com sucesso."
    Case 18001:     Result = "Código de ativação inválido."
    Case 18002:     Result = "Código de ativação de emergência Incorreto."
    Case 18098:     Result = "SAT em processamento."
    Case 18099: Result = "Erro desconhecido Informar o administrador."

End Select

If Result <> "" Then
    MsgBox Result
End If

End Function

Public Function LerINI2(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim ret As String
 ret = String$(400, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", ret, Len(ret), Arquivo)
 ret = Left$(ret, retlen)
 LerINI2 = ret
End Function

Public Sub PesquisaGeral(Tabela As String, SQL As String, Campo As String)
    Formulário.FramePesquisaG.Visible = True

    Formulário.DatPesquisaG.DatabaseName = Tabela
    Formulário.DatPesquisaG.RecordSource = SQL
'    Formulário.DatPesquisaG.DataField = "Cidade"
'    Formulário.DatPesquisaG.ListField = Campo
    Formulário.DatPesquisaG.Refresh
    Formulário.DblPesquisaG.Refresh
'    Formulário.DblPesquisaG.SetFocus

End Sub

Public Sub CopiarTabelaExcel(Rs As Recordset, ws As Worksheet, StartingCell As ExlCell)

Dim Vetor() As Variant
Dim Row As Long, Col As Long
Dim fd As Field

Rs.MoveLast
ReDim Vetor(Rs.RecordCount + 1, Rs.Fields.Count)

Col = 0
For Each fd In Rs.Fields
  Vetor(0, Col) = fd.Name
  Col = Col + 1
Next

Rs.MoveFirst
For Row = 1 To Rs.RecordCount
   For Col = 0 To Rs.Fields.Count - 1
       Vetor(Row, Col) = Rs.Fields(Col).value
       If IsNull(Vetor(Row, Col)) Then Vetor(Row, Col) = ""
   Next
   Rs.MoveNext
Next
ws.Range(ws.Cells(StartingCell.Row, StartingCell.Col), ws.Cells(StartingCell.Row + Rs.RecordCount + 1, _
StartingCell.Col + Rs.Fields.Count)).value = Vetor

End Sub


Function Truncar(Valor As Currency, Decimais As Integer) As String
Dim X As Integer
Dim Tam As Integer
Valor = Format(Valor, "#,##0.0000")
For X = 1 To Len(Valor)
    If Mid$(Valor, X, 1) = "," Then
        Tam = X
    End If
Next

Truncar = Left(Valor, Tam + Decimais)

End Function

