Attribute VB_Name = "ModuloSRCaixa"
Public Consulta As New ADODB.Recordset 'ultilizada para consultas no módulo

Public Sub InsereBanco(InsTabela As String, InsCampos As String, InsValores As String)
'On Error GoTo Linha
Dim Sql As String

Sql = " INSERT INTO " & InsTabela & "(" & InsCampos & ") "
Sql = Sql & " VALUES (" & InsValores & ")"
BancoSRCaixa.Execute Sql

Linha:
    If Err.Number = "-2147217900" Then
        MsgBox "Já existe um registro com as informações informadas!", 16
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description
    End If
End Sub

Public Function Status(Tipo As Currency)
If Formulário.Name <> "FrmAbreForm" Then
    If Tipo = -1 Then
        Formulário.txtStatus.Text = ""
    ElseIf Tipo = 0 Then
        Formulário.BtPesquisa.Enabled = False
        Formulário.txtStatus.Text = "[F3-Novo Cadastro]"
    ElseIf Tipo = 1 Then
        Formulário.BtPesquisa.Enabled = True
        Formulário.txtStatus.Text = "[Enter-Pesquisa]-[F3-Novo Cadastro]"
    ElseIf Tipo = 2 Then
        Formulário.BtPesquisa.Enabled = True
        Formulário.txtStatus.Text = "[Enter-Pesquisa]-[F3-Novo Cadastro]-[F4-Editar]-[F7-Excluir]"
    ElseIf Tipo = 3 Then
        Formulário.BtPesquisa.Enabled = False
        Formulário.txtStatus.Text = "[F5-Salvar]-[F6-Cancelar]"
    ElseIf Tipo = 4 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa]-[F5-Salvar]-[F6-Cancelar]"
    ElseIf Tipo = 5 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa por Titular]-[F3-Pesquisa por Denominacao]"
    ElseIf Tipo = 6 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa]"
    ElseIf Tipo = 7 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa] - [ESC - SAIR]"
    ElseIf Tipo = 8 Then
        Formulário.txtStatus.Text = "[ESC - SAIR]"
    ElseIf Tipo = 9 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa por Titular]-[F3-Pesquisa por Denominacao]-[ESC-SAIR]"
    ElseIf Tipo = 10 Then
        Formulário.txtStatus.Text = "[F2-Pesquisa por Cliente/Titular]-[F3-Pesquisa por Banco]"
    End If
End If
End Function

Public Function Converte(exprData As String) As Date
Dim MontaData As String
    If Not IsDate(exprData) Then
        MontaData = Mid(exprData, 1, 5) & "/20" & Mid(exprData, 7, 2)
    End If
    If IsDate(MontaData) Then
        Converte = MontaData
    ElseIf IsDate(exprData) Then
        Converte = exprData
    Else
        Converte = Date
    End If
    
'    If Year(Converte) <> AnoGlobal And AnoGlobal <> "" Then
'        If MsgBox("Você esta fazendo um lançamento de ano diferente do ultimo lançamento, deseja continuar?", 36) = vbNo Then
'            Converte = Day(Converte) & "/" & Month(Converte) & "/" & AnoGlobal
'        End If
'        AnoGlobal = Year(Converte)
'    Else
'        AnoGlobal = Year(Converte)
'    End If

End Function

Public Function FuncaoTeclas(Teclas As Integer, Tipo As String)
On Error Resume Next
DoEvents

Dim x As Currency

If Tipo = 1 Then 'keydown
    
    If Teclas = vbKeyReturn Then 'vbkeyreturn = enter e 40 = flecha para baixo
        SendKeys "{Tab}" 'pressiona tab
    'ElseIf Teclas = 38 Then 'tecla flecha para cima
    '    SendKeys "+{tab}" 'pressiona shift + tab
    ElseIf Teclas = 114 And MDIPrincipal.TlbPrincipal.Buttons.Item(1).Visible = True Then  'tecla F3
        'Formulário.TabCad.Tab = 1
        Manutencoes 1 'novo
    ElseIf Teclas = 115 And MDIPrincipal.TlbPrincipal.Buttons.Item(2).Visible = True Then  'tecla F4
        'Formulário.TabCad.Tab = 1
        Manutencoes 2 'editar
    ElseIf Teclas = 116 And MDIPrincipal.TlbPrincipal.Buttons.Item(14).Visible = True Then  'tecla F5
        'Formulário.TabCad.Tab = 1
        Confirmacoes 14 'salvar
    ElseIf Teclas = 117 And MDIPrincipal.TlbPrincipal.Buttons.Item(15).Visible = True Then  'tecla F6
        'Formulário.TabCad.Tab = 1
        Confirmacoes 15 'cancelar
    ElseIf Teclas = 118 And MDIPrincipal.TlbPrincipal.Buttons.Item(3).Visible = True Then  'tecla F7
        'Formulário.TabCad.Tab = 1
        Manutencoes 3 'excluir
    ElseIf Teclas = 27 Then 'tecla esc
        If MDIPrincipal.TlbPrincipal.Buttons.Item(7).Visible = True Then
            Unload Formulário 'fecha formulario
        Else
            If MsgBox("Para sair da tela, primeiro você precisa salvar o registro" & Chr(10) & Chr(10) & "Deseja salva-lo agora ???", 36) = vbYes Then
                Confirmacoes 14
            Else
                Confirmacoes 15
            End If
        End If
    End If
    
ElseIf Tipo = 2 Then 'keyup"

    If Teclas = 13 Or Teclas = 9 Then
        SelecionaCampo
    End If
    
ElseIf Tipo = 3 Then 'keypress

    If Teclas = 5 And Formulário.TabCad.TabEnabled(0) = True Then  'tecla ctrl + e
        Formulário.TabCad.Tab = 0
        Formulário.TxtPesquisa.SetFocus
    ElseIf Teclas = 1 And Formulário.TabCad.TabEnabled(1) = True Then  'tecla ctrl + a
        Formulário.TabCad.Tab = 1
    ElseIf Teclas = 12 And Formulário.Registros.EditMode = dbEditNone Then 'tecla ctrl + l
        'FormLocalizar.Show
    End If
End If

End Function

Function Busca(varPesquisa As String, Campo As Variant, Tabela As String) As Boolean
Dim Valor As String

Busca = True

On Error GoTo LinhaErro
    If IsNumeric(Campo.Text) And Campo.Text <> "" Then
        If Consulta.State = 1 Then Consulta.Close
        
        If Campo.Text = "" Or Not IsNumeric(Campo.Text) Then
            Valor = 0
        Else
            If Campo.Tag = "SPlano" Then
                Valor = "'" & Formulário.Lbl(32).Caption & "§" & Campo.Text & "'"
            Else
                Valor = Campo.Text
            End If
        End If
        
        
        If Tabela = "CTR" And Campo = "1" Then
            Busca = False
        ElseIf Tabela = "CTA" And (Campo = "1000" Or Campo = "2000") Then
            Busca = False
        End If
        
        If Busca = False Then
            Consulta.Open "Select * From " & Tabela & " Where " & varPesquisa & " Like 0", BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
        Else
            Consulta.Open "Select * From " & Tabela & " Where " & varPesquisa & " Like " & Valor, BancoSRCaixa, adOpenDynamic, adLockOptimistic, adCmdText
        End If
            
             If Consulta.Bof Then
                Busca = False
                If Campo.Text <> 0 Then
                    MsgBox "Ocorrência não encontrada!", vbExclamation
                    Campo.SetFocus
                    Campo.SelLength = 5000
                End If
            Else
                Busca = True
            End If
    Else
        Busca = False
    End If
LinhaErro:

End Function

Public Function SeCampoNulo(varRecordset As Variant, Campo As String) As String
Dim Conv As Integer

    If IsNumeric(Campo) Then
        If varRecordset(CInt(Campo)) = 0 Then
            SeCampoNulo = ""
        ElseIf IsNull(varRecordset(CInt(Campo))) = True Then
            SeCampoNulo = ""
        Else
            SeCampoNulo = varRecordset(CInt(Campo))
        End If
    Else
        If varRecordset(Campo) = 0 Then
            SeCampoNulo = ""
        ElseIf IsNull(varRecordset(Campo)) = True Then
            SeCampoNulo = ""
        Else
            SeCampoNulo = Trim(varRecordset(Campo))
        End If
    End If
End Function

Public Sub PreencheCombo(Sql As String, Objeto As Variant)
    Set Consulta = New ADODB.Recordset
    Consulta.Open Sql, StringConexao, adOpenKeyset, adLockReadOnly
    
    Objeto.Clear
    
    Do While Not Consulta.Eof
        Objeto.AddItem Consulta(0)
        Consulta.MoveNext
    Loop
End Sub

Public Function ApenasNumeros(varString As String) As String
Dim x As Integer

    For x = 1 To Len(varString)
        If IsNumeric(Mid(varString, x, 1)) Then
            ApenasNumeros = ApenasNumeros & Mid(varString, x, 1)
        End If
    Next

End Function

Public Sub LogErros()
Dim StrErro As String

    StrErro = Now & " Ocorreu o erro: " & Err.Number & " - " & Err.Description & " - Formulário: " & Formulário.Name
    
    If Err.Number <> 0 Then
        MsgBox "Ocorreu o seguinte erro: " & Chr(10) & Err.Number & "-" & Err.Description, 16
        Call EscreveINI("Erros", StrErro, 0, App.Path & "\Log.ini")
        Err.Clear
    End If

End Sub
