VERSION 5.00
Begin VB.Form FrmFim 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controles"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox OpcLista 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      ItemData        =   "FrmFim.frx":0000
      Left            =   270
      List            =   "FrmFim.frx":001F
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   4650
   End
End
Attribute VB_Name = "FrmFim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
SendKeys "{TAB}"
OpcLista.ListIndex = 0
OpcLista.SetFocus
End Sub

Private Sub Form_Load()
Flag = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Flag = 0
End Sub

Private Sub OpcLista_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Retorno As Integer
Dim Status As String
Dim Entradas As Currency
Dim Saidas As Currency
Dim VdaPrazo  As Currency
Dim TotVda As Recordset
Dim TotCxa As Recordset

If KeyCode = 27 Then
    Unload Me
    Exit Sub
End If

If KeyCode = 13 Then
    Select Case OpcLista.ListIndex
        Case 0
            Select Case Impressora
                Case 1
                    Retorno = AcionarGaveta() ' "0", "32", "38")
                    MensErro (Retorno)
                Case 2
                    Status = CTV3.ABRIRGAVETA(Mensagem)
                    Mens.Caption = Status
                Case 99
                    If Left(PortaNF, 3) <> "COM" Then
                        If Impr40 = "Sim" Then
                            Iniciar (False)
                                Print #1, "" & Nor & Chr(27) & Chr(111) & Chr(27) & Chr(118)
                            Fim (False)
                        End If
                    End If
            End Select
        Case 1
            FrmFrente.soma = 1
            If MsgBox("Confirma leitura X ?", vbYesNo) = vbYes Then
                Select Case Impressora
                    Case 1
                        Retorno = LEITURAX("0")
                        MensErro (Retorno)
                    Case 2
                        Status = CTV3.LEITURAX("0", " ", Mensagem)
                        Mens.Caption = Status
                    Case 5
                        Retorno = Bematech_FI_LeituraX()
                        FrmFrente.Mens.Caption = Retorno
                    
                    Case 6
                        Retorno = Daruma_FI_LeituraX()
                        FrmFrente.Mens.Caption = Retorno
                    
                    Case 7
                        Retorno = Elgin.Elgin_LeituraX()
                        TrataRetorno2 Retorno
                                            
                    Case 8
                        DLLG2_ExecutaComando Handle, "EmiteLeituraX"
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                                                
                    Case 99
                        Entradas = 0
                        Saidas = 0
                        On Error GoTo ManutençãoErro
                        Set TotVda = FrmFrente.Banco.OpenRecordset("SELECT Sum(TOT.Total) AS SomaDeTotal, TOT.Pgto FROM TOT WHERE TOT.Data = #" & Format(Date, "mm/dd/yyyy") & "# AND TOT.Maquina = " & """" & LimpaCampo(Left(FrmFrente.Maquina, 15)) & """" & " AND TOT.Cancelado = False  GROUP BY TOT.Pgto ORDER BY TOT.Pgto")
                        TotVda.MoveFirst
                        
                        Set TotCxa = FrmFrente.Banco.OpenRecordset("SELECT * FROM CXA WHERE Data = #" & Format(Date, "mm/dd/yyyy") & "# AND Maquina = " & """" & LimpaCampo(Left(FrmFrente.Maquina, 15)) & """" & " AND Tipo = 'M' ORDER BY  D_C")
                        If TotCxa.RecordCount <> 0 Then
                            TotCxa.MoveFirst
                        End If
                        
                        Iniciar (False)
                        If Left(PortaNF, 3) <> "COM" Then

                            Print #1, ""
                            Print #1, Centralizar("LEITURA X", 47)
                            Print #1, "==============================================="
                            Print #1, "Data: " & Date & "   Hora: " & Time

                            Print #1, "-----------------------------------------------"
                            Print #1, Centralizar("VENDAS", 48)
                            Print #1, "-----------------------------------------------"
                            While Not TotVda.EOF
                                Select Case TotVda("Pgto")
                                    Case 1
                                        Print #1, "Dinheiro ....................(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                    Case 2
                                        Print #1, "Cheque ......................(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                    Case 3
                                        Print #1, "Cheque Pré ..................(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                    Case 4
                                        Print #1, "Cartão ......................(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                    Case 5
                                        Print #1, "Ticket ......................(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                    Case 6
                                        Print #1, "Venda a Prazo ...............(+) "; Alinhar(TotVda("SomaDeTotal"), 12, 2)
                                        VdaPrazo = VdaPrazo + TotVda("SomaDeTotal")
                                End Select
                                If TotVda("Pgto") < 6 Then
                                    Entradas = Entradas + TotVda("SomaDeTotal")
                                End If
                                TotVda.MoveNext
                            Wend
                            Print #1, "-----------------------------------------------"
                            Print #1, Centralizar("MOVIMENTACAO CAIXA", 47)
                            Print #1, "-----------------------------------------------"
                            Print #1, ""
                            
                            If TotCxa.RecordCount <> 0 Then
                                While Not TotCxa.EOF
                                    If TotCxa("D_C") = "C" Then
                                        Print #1, Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))); "(+)"; Alinhar(TotCxa("valor"), 12, 2)
                                        Entradas = Entradas + TotCxa("Valor")
                                    ElseIf TotCxa("D_C") = "D" Then
                                        Print #1, Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))); "(-)"; Alinhar(TotCxa("valor"), 12, 2)
                                        Saidas = Saidas + TotCxa("Valor")
                                    End If
                                    TotCxa.MoveNext
                                Wend
                            End If
                            Print #1, "-----------------------------------------------"
                            Print #1, Centralizar("TOTALIZACAO", 47)
                            Print #1, "-----------------------------------------------"
                            Print #1, ""
                            Print #1, "Total do Caixa...............(=) "; Alinhar(CCur(Entradas) + CCur(VdaPrazo), 12, 2)
                            Print #1, "Saidas do Caixa..............(-) "; Alinhar(CCur(Saidas), 12, 2)
                            Print #1, "Vendas a Prazo...............(-) "; Alinhar(CCur(VdaPrazo), 12, 2)
                            Print #1, "Total Liquido................(=) "; Alinhar(CCur(Entradas) - CCur(Saidas), 12, 2)
                            Print #1, "" & Nor
                            Print #1, ""
                        
                        Else
                        
                            Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX(Centralizar("LEITURA X", 47) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("===============================================" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("Data: " & Date & "   Hora: " & Time & Chr(10), 2, 0, 0, 0, 0)
                            
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX(Centralizar("VENDAS", 48) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            
                            While Not TotVda.EOF
                                Select Case TotVda("Pgto")
                                    Case 1
                                        Comando = FormataTX("Dinheiro ....................(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Case 2
                                        Comando = FormataTX("Cheque ......................(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Case 3
                                        Comando = FormataTX("Cheque Pré ..................(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Case 4
                                        Comando = FormataTX("Cartão ......................(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Case 5
                                        Comando = FormataTX("Ticket ......................(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                    Case 6
                                        Comando = FormataTX("Venda a Prazo ...............(+) " & Alinhar(TotVda("SomaDeTotal"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                        VdaPrazo = VdaPrazo + TotVda("SomaDeTotal")
                                End Select
                                If TotVda("Pgto") < 6 Then
                                    Entradas = Entradas + TotVda("SomaDeTotal")
                                End If
                                TotVda.MoveNext
                            Wend
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX(Centralizar("MOVIMENTACAO CAIXA", 47) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                           
                            If TotCxa.RecordCount <> 0 Then
                                While Not TotCxa.EOF
                                    If TotCxa("D_C") = "C" Then
                                        Comando = FormataTX(Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))) & "(+)" & Alinhar(TotCxa("valor"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                        Entradas = Entradas + TotCxa("Valor")
                                    ElseIf TotCxa("D_C") = "D" Then
                                        Comando = FormataTX(Left(TotCxa("Historico"), 30) + Space(30 - Len(Left(TotCxa("Historico"), 30))) & "(-)" & Alinhar(TotCxa("valor"), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                                        Saidas = Saidas + TotCxa("Valor")
                                    End If
                                    TotCxa.MoveNext
                                Wend
                            End If
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX(Centralizar("TOTALIZACAO", 47) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("-----------------------------------------------" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                            
                            Comando = FormataTX("Total do Caixa...............(=) " & Alinhar(CCur(Entradas) + CCur(VdaPrazo), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("Saidas do Caixa..............(-) " & Alinhar(CCur(Saidas), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("Vendas a Prazo...............(-) " & Alinhar(CCur(VdaPrazo), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("Total Liquido................(=) " & Alinhar(CCur(Entradas) - CCur(Saidas), 12, 2) & Chr(10), 2, 0, 0, 0, 0)
                            
                            Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                            Comando = FormataTX("" & Chr(10), 2, 0, 0, 0, 0)
                                                    
                        End If
                        Fim (False)
                End Select
            End If
            FrmFrente.soma = 0
        Case 2
            FrmFrente.soma = 1
            MsgBox "A T E N Ç Ã O . . .  Está operação fecha o caixa por hoje !", vbExclamation, App.Title
            If MsgBox("Confirma redução Z ?", vbYesNo + vbDefaultButton2) = vbYes Then
                Select Case Impressora
                    Case 1
                        Retorno = REDUCAOZ("0")
                        MensErro (Retorno)
                    Case 2
                        Status = CTV3.FecharEcf("0", Chr(13), Mensagem)
                        Mens.Caption = Status
                    Case 5
                        Retorno = Bematech_FI_ReducaoZ("", "")
                        FrmFrente.Mens.Caption = Retorno
                    Case 6
                        Retorno = Daruma_FI_ReducaoZ("", "")
                        FrmFrente.Mens.Caption = Retorno
                        
                    Case 7
                        Retorno = Elgin.Elgin_ReducaoZ("", "")
                        TrataRetorno2 Retorno
                        
                    Case 8
                        
                        DLLG2_LimpaParams Handle
                        DLLG2_DefineTimeout Handle, 30
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                                                                        
                        DLLG2_LimpaParams Handle
                        DLLG2_ExecutaComando Handle, "EmiteReducaoZ"
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                        
                        DLLG2_LimpaParams Handle
                        DLLG2_DefineTimeout Handle, 10
                        coderro = DLLG2_ObtemNomeErro(Handle, "coderro", 0)
                        circunst = DLLG2_ObtemCircunstancia(Handle, "circunst", 0)
                        If coderro <> "OK" Then
                            MsgBox "Nome do Erro: " & coderro & "    " & "Circunstancia: " & circunst
                            
                        End If
                        
                        
                End Select
            End If
            FrmFrente.soma = 0
        Case 3
            FormLancaCta.Show vbModal
        Case 4
            FormBaixarDupli.Show vbModal
        Case 5
            If ConsultaCh <> "calc" Then
                Shell ConsultaCh, vbHide
            End If
        Case 6
            FormFchCxanovo.Show vbModal
        Case 7
            FormCaixa.Show vbModal
        Case 8
            FrmFrente.soma = 1
            Select Case Impressora
                Case 1
                    CLOSECIF
                Case 2
                    Status = CTV3.FechaComunicacao(Mensagem)
                    FrmFrente.Mens.Caption = Status
                Case 5
                    Retorno = Bematech_FI_FechaPortaSerial()
                    FrmFrente.Mens.Caption = Retorno
                Case 6
                    Retorno = Daruma_FI_FechaPortaSerial()
                    FrmFrente.Mens.Caption = Retorno
                
                Case 8
                    DLLG2_EncerraDriver (Handle)
                                        
            End Select
            FrmFrente.soma = 0
            FrmFrente.Banco.Close
            Set Banco = Nothing
            Unload Me
            Unload FrmFrente
            FrmPath.Show
    End Select

End If

Exit Sub
ManutençãoErro:
    MsgBox Err.Description & Err.Number

End Sub

