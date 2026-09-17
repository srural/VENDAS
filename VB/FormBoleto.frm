VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormBoleto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   43
      Top             =   5940
      Width           =   14370
   End
   Begin VB.CommandButton CmbBoleto 
      Caption         =   "Boleto Banco"
      Height          =   375
      Left            =   7110
      TabIndex        =   42
      Top             =   3330
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12435
      TabIndex        =   38
      Top             =   0
      Width           =   12435
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão de Boleto"
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
         TabIndex        =   39
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.TextBox TxtEnt 
      Height          =   255
      Left            =   900
      TabIndex        =   37
      Top             =   4860
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   315
      ScaleHeight     =   885
      ScaleWidth      =   8670
      TabIndex        =   36
      Top             =   3870
      Width           =   8700
      Begin VB.TextBox TxtInf1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   285
         Width           =   8655
      End
      Begin VB.TextBox TxtInf2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   0
         Width           =   8700
      End
      Begin VB.TextBox TxtInf3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   570
         Width           =   8700
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Impressão Boleto"
      Height          =   690
      Left            =   3870
      TabIndex        =   34
      Top             =   4905
      Width           =   3840
      Begin VB.CommandButton BtIMprimir 
         Caption         =   "&Boleto Impresso"
         Height          =   330
         Left            =   2070
         TabIndex        =   12
         Top             =   225
         Width           =   1635
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "&A4"
         Height          =   240
         Index           =   1
         Left            =   1215
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "&Fomulário"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.CommandButton BtDupli 
      Caption         =   "&Duplicata"
      Height          =   480
      Left            =   7815
      TabIndex        =   13
      Top             =   4995
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imformações para recebimento"
      Height          =   2355
      Left            =   5310
      TabIndex        =   29
      Top             =   765
      Width           =   3705
      Begin VB.CheckBox ChkImprimir 
         Caption         =   "&Imprimir Informações"
         Height          =   375
         Left            =   1530
         TabIndex        =   8
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox ChkAceite 
         Caption         =   "&Aceite"
         Height          =   285
         Left            =   315
         TabIndex        =   7
         Top             =   1845
         Width           =   825
      End
      Begin VB.TextBox TxtPrazo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "3"
         Top             =   630
         Width           =   1140
      End
      Begin VB.TextBox TxtMulta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   4
         Text            =   "2"
         Top             =   630
         Width           =   1140
      End
      Begin VB.TextBox TxtMora 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   225
         TabIndex        =   5
         Text            =   "10"
         Top             =   1350
         Width           =   1140
      End
      Begin VB.TextBox TxtProtesto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1770
         TabIndex        =   6
         Text            =   "0"
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Prazo &Multa"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   405
         Width           =   840
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Multa (%)"
         Height          =   195
         Index           =   4
         Left            =   1725
         TabIndex        =   32
         Top             =   405
         Width           =   645
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Mo&ra Mês (%)"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   31
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Protestar Após"
         Height          =   195
         Index           =   8
         Left            =   1770
         TabIndex        =   30
         Top             =   1125
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2355
      Left            =   315
      TabIndex        =   18
      Top             =   765
      Width           =   4785
      Begin VB.TextBox TxtPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   28
         Top             =   465
         Width           =   1275
      End
      Begin VB.TextBox TxtPcl 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   27
         Top             =   1755
         Width           =   1275
      End
      Begin VB.TextBox TxtDoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   26
         Top             =   1110
         Width           =   1275
      End
      Begin MSMask.MaskEdBox MebEmiss 
         Height          =   285
         Left            =   135
         TabIndex        =   20
         Top             =   465
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pedido"
         Height          =   240
         Index           =   0
         Left            =   3465
         TabIndex        =   0
         Top             =   585
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Documento"
         Height          =   240
         Index           =   1
         Left            =   3465
         TabIndex        =   1
         Top             =   1155
         Width           =   1230
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Parcela"
         Height          =   240
         Index           =   2
         Left            =   3465
         TabIndex        =   2
         Top             =   1785
         Width           =   1005
      End
      Begin MSMask.MaskEdBox MebVenc 
         Height          =   285
         Left            =   135
         TabIndex        =   22
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MebValor 
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Top             =   1755
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "Va&lor"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   25
         Top             =   1515
         Width           =   360
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Vencimento"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   23
         Top             =   870
         Width           =   840
      End
      Begin VB.Label LblDados 
         AutoSize        =   -1  'True
         Caption         =   "&Emissão"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   225
         Width           =   585
      End
   End
   Begin MSMask.MaskEdBox MebDtSaida 
      DataField       =   "Custo"
      Height          =   315
      Left            =   1980
      TabIndex        =   15
      Tag             =   "S"
      Top             =   5220
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebTotal 
      DataField       =   "Custo"
      Height          =   315
      Left            =   360
      TabIndex        =   14
      Tag             =   "S"
      Top             =   5220
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      Format          =   "#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MebPcl 
      DataField       =   "Custo"
      Height          =   315
      Left            =   1170
      TabIndex        =   40
      Tag             =   "S"
      Top             =   3285
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Parcelas"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   41
      Top             =   3330
      Width           =   615
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Total"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   35
      Top             =   4965
      Width           =   360
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Processamento"
      Height          =   195
      Index           =   6
      Left            =   1980
      TabIndex        =   19
      Top             =   4965
      Width           =   1485
   End
End
Attribute VB_Name = "FormBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtDupli_Click()

Dim TabSql As Recordset
Dim X As Integer
Dim y As Integer
Dim Copias As Integer
Dim Conf As Recordset
Dim Multa As Currency
Dim Juros As Currency
Dim VarVenc As String
Dim DtProtesto As String

If Opt(2).value = True Then
            Set TabSql = Banco.OpenRecordset("SELECT Ent.CodEntidade, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF,ENT.InscrEst, ENT.CGC FROM  ENT  WHERE  CodEntidade = " & TxtEnt)
        TabSql.Requery
        
        If TabSql.RecordCount <= 0 Then
            MsgBox "Informações inconsistentes ", vbInformation, App.Title
            Exit Sub
        End If
        TabSql.MoveFirst
        
        Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 15 Order by Codigo")
        Conf.Requery
        If Conf.RecordCount <= 0 Then
            MsgBox "Boleto não configurado !", vbInformation, App.Path
            Exit Sub
        End If
        Conf.MoveFirst
        
        Copias = MebPcl
        
        Variavel(1) = "" 'emissao
        Variavel(2) = "" 'cgc
        Variavel(3) = "" 'Inscricao
        Variavel(4) = "" 'Extenso
        Variavel(5) = 0  'Total da Nota
        Variavel(6) = TxtInf1
        Variavel(7) = "" 'Pulo
        Variavel(8) = ""  ' Vencimento
        Variavel(9) = "" ' Numero da Nota
        Variavel(10) = ""
        
        If TabSql.RecordCount <> 0 Then
               Iniciar True 'inicia o registro onde sera gravado as informacoes
               
        '       Do While Not TabSql.Eof
                     For y = 1 To Copias
                         If TabSql!Classificacao = 2 Then
                            If TabSql!CGC <> "" Then
                                Variavel(2) = Format(TabSql!CGC, "@@.@@@.@@@/@@@@-@@")
                            Else
                                Variavel(2) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                            End If
                         ElseIf TabSql!Classificacao = 1 Then
                            Variavel(2) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                         End If
                         
                         If TabSql!Classificacao = 2 Then
                            If TabSql!InscrEst <> "" Then
                                Variavel(3) = TabSql!InscrEst
                            End If
                         ElseIf TabSql!Classificacao = 1 Then
        '                    If TabSql!Rg <> "" Then
        '                        Variavel(3) = TabSql!Rg
        '                    End If
                         End If
                            FormBaixarDupli.DatDados(0).Recordset.MoveFirst
                            Do Until FormBaixarDupli.DatDados(0).Recordset.Eof
                                If FormBaixarDupli.DatDados(0).Recordset("PgtTemp") = -1 Then
'                                    FormBaixarDupli.DatDados(0).Recordset.Edit
'                                    FormBaixarDupli.DatDados(0).Recordset.Fields("Tipo_Dup") = "DP"
                                     Variavel(10) = Variavel(10) & FormBaixarDupli.DatDados(0).Recordset.Fields("Pedido") & "/"
'                                    FormBaixarDupli.DatDados(0).Recordset.Update
                                    
                                End If
                                FormBaixarDupli.DatDados(0).Recordset.MoveNext
                            Loop
                         
                                                         
                         Variavel(1) = Day(Date) & "     " & RetornaMes(Month(Date)) & "       " & Year(Date)
                         Variavel(4) = "(" & Format(PassaExtenso(MebValor / Copias), ">") & ")"
                         Variavel(5) = MebValor / Copias
                         If y = 1 Then
                            Variavel(8) = MebVenc
                         Else
                            Variavel(8) = CDate(Variavel(8)) + 30
                         End If
        '                 If TabSql!NroNt <> 0 Then
        '                    Variavel(8) = TabSql!NroNt & "/" & TabSql("NroPcl")
        '                    Variavel(9) = TabSql!NroNt
        '                 Else
        '                    If Opt(1).value = True Then
        '                        Variavel(8) = TabSql!Documento & "/" & TabSql("NroPcl")
        '                        Variavel(9) = ""
        '                    Else
        '                        Variavel(8) = TabSql("Dup.Pedido") & "/" & TabSql("NroPcl")
        '                        Variavel(9) = ""
        '                    End If
        '                 End If
                         
                         ImprimeLinha 0, 0, Chr(27) & Chr(48), Len(Chr(27) & Chr(48)), False, False
            
                         LinhaAnt = 0
                         Conf.MoveFirst
                         While Not Conf.Eof
                                If Conf!Linha <> LinhaAnt Then
                                    For X = 1 To Conf!Linha - LinhaAnt
                                        If y = 3 Then
                                            If Conf!NomeCampo <> "%7" Then
                                                  ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                            End If
                                        Else
                                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                        End If
                                    Next
                                End If
                                LinhaAnt = Conf!Linha
                                If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                       If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                           ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                           If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           Else
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           End If
                                       End If
                                Else
                                       If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                            ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                            If Mid(Conf("Formato"), 1, 1) = "+" Then
                                               ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                            Else
                                               If IsNull(TabSql(Conf!NomeCampo)) Then
                                                  ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               Else
                                                  ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               End If
                                           End If
                                       End If
                                End If
                                Conf.MoveNext
                         Wend
'                         If Copias = 3 Then
'                               If y = 1 Then
'                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
'                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
'                                End If
'                         End If
        
                     Next
                     TabSql.MoveNext
'                     If Copias = 3 Then
'                        ImprimeLinha 0, 0, Chr(12), Len(Chr(12)), False, False ' Salta Pagina
'                     End If
                     
        '       Loop
               ImprimeLinha 0, 0, Chr(27) & Chr(50), Len(Chr(27) & Chr(50)), False, False
               
        '       For x = 1 To Conf!Linha - LinhaAnt
        '            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
        '       Next
                Fim True
                If MsgBox("Boleto impresso com sucesso ?", vbYesNo, App.Title) = vbYes Then
                    
                    For y = 1 To Copias
                        FormBaixarDupli.DatDados(0).Recordset.AddNew
                        FormBaixarDupli.DatDados(0).Recordset.Fields("Entidade") = TxtEnt
                        FormBaixarDupli.DatDados(0).Recordset.Fields("Tipo_Dup") = "DP"
                        FormBaixarDupli.DatDados(0).Recordset.Fields("DtEmiss") = Format(MebEmiss, "dd/mm/yyyy")
                        If y = 1 Then
                            FormBaixarDupli.DatDados(0).Recordset.Fields("DtVenc") = Format(MebVenc, "dd/mm/yyyy")
                        Else
                            MebVenc = CDate(MebVenc) + 30
                            FormBaixarDupli.DatDados(0).Recordset.Fields("DtVenc") = Format(MebVenc, "dd/mm/yyyy")
                        End If
                        FormBaixarDupli.DatDados(0).Recordset.Fields("Valor") = MebValor / Copias
                        FormBaixarDupli.DatDados(0).Recordset.Fields("Pedido") = 0
                        FormBaixarDupli.DatDados(0).Recordset.Fields("NroPcl") = 1
                        FormBaixarDupli.DatDados(0).Recordset.Update
                    Next
                        
                        FormBaixarDupli.DatDados(0).Recordset.MoveFirst
                        Do Until FormBaixarDupli.DatDados(0).Recordset.Eof
                            If FormBaixarDupli.DatDados(0).Recordset("PgtTemp") = -1 Then
                                FormBaixarDupli.DatDados(0).Recordset.Delete
    '                            FormBaixarDupli.DatDados(0).Recordset.Fields("Tipo_Dup") = "DP"
    '                            FormBaixarDupli.DatDados(0).Recordset.Fields("DtVenc") = Format(MebVenc, "dd/mm/yyyy")
    '                            FormBaixarDupli.DatDados(0).Recordset.Update
                                
                            End If
                            FormBaixarDupli.DatDados(0).Recordset.MoveNext
                        Loop
                    Unload Me
                End If
               
        End If
Else

        If Opt(0).value = True Then
            Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, DUP.Pedido,Ent.CodEntidade, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF,ENT.InscrEst, ENT.CGC, DUP.*, DUP.Cond FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE (((DUP.Pedido)=  " & TxtPedido & " ) AND ((DUP.Cond)=False))")  ' (((DUP.PgtTemp)= True) AND ((DUP.Cond)=False))")
            
'            Set TabSql = Banco.OpenRecordset("SELECT ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF,ENT.InscrEst, ENT.CGC FROM  ENT  WHERE  CodEntidade = " & TxtEnt)
        ElseIf Opt(1).value = True Then
                Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, Ent.CodEntidade, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC,ENT.InscrEst, DUP.*, DUP.Cond FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE (((DUP.Cond)=False) AND ((DUP.Documento)= " & """" & TxtDoc & """" & "))")
            Else
                Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, DUP.Pedido, Ent.CodEntidade, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC, Ent.InscrEst, DUP.*, DUP.Cond FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE DUP.CodDup = " & TxtPedido & " AND  DUP.Cond = False  ")
        End If
        
        TabSql.Requery
        
        If TabSql.RecordCount <= 0 Then
            MsgBox "Informações inconsistentes ", vbInformation, App.Title
            Exit Sub
        End If
        TabSql.MoveFirst
        
        Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 15 Order by Codigo")
        Conf.Requery
        If Conf.RecordCount <= 0 Then
            MsgBox "Boleto não configurado !", vbInformation, App.Path
            Exit Sub
        End If
        Conf.MoveFirst
        
        If OptTipo(0).value = True Then
            Copias = 1
        Else
            Copias = 1
        End If
        Variavel(1) = "" 'emissao
        Variavel(2) = "" 'cgc
        Variavel(3) = "" 'Inscricao
        Variavel(4) = "" 'Extenso
        Variavel(5) = 0  'Total da Nota
        Variavel(6) = TxtInf1
        Variavel(7) = "" 'Pulo
        Variavel(8) = "" ' vencimento
        Variavel(9) = "" ' Numero da Nota
        Variavel(10) = ""
        
        If TabSql.RecordCount <> 0 Then
               Iniciar True 'inicia o registro onde sera gravado as informacoes
               TabSql.MoveFirst
               Do While Not TabSql.Eof
        '             For y = 1 To Copias
                         If TabSql!Classificacao = 2 Then
                            If TabSql!CGC <> "" Then
                                Variavel(2) = Format(TabSql!CGC, "@@.@@@.@@@/@@@@-@@")
                            Else
                                Variavel(2) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                            End If
                         ElseIf TabSql!Classificacao = 1 Then
                            Variavel(2) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                         End If
                         
                         If TabSql!Classificacao = 2 Then
                            If TabSql!InscrEst <> "" Then
                                Variavel(3) = TabSql!InscrEst
                            End If
                         ElseIf TabSql!Classificacao = 1 Then
        '                    If TabSql!Rg <> "" Then
        '                        Variavel(3) = TabSql!Rg
        '                    End If
                         End If
                         
                         Variavel(10) = TabSql("dup.Pedido") & "/" & TabSql("NroPcl")
                                
                         Variavel(1) = Day(Date) & "     " & RetornaMes(Month(Date)) & "       " & Year(Date)
                         Variavel(4) = "(" & Format(PassaExtenso(MebValor), ">") & ")"
                         Variavel(5) = MebValor
                         Variavel(8) = TabSql!DtVenc
        '                 If TabSql!NroNt <> 0 Then
        '                    Variavel(8) = TabSql!NroNt & "/" & TabSql("NroPcl")
        '                    Variavel(9) = TabSql!NroNt
        '                 Else
        '                    If Opt(1).value = True Then
        '                        Variavel(8) = TabSql!Documento & "/" & TabSql("NroPcl")
        '                        Variavel(9) = ""
        '                    Else
        '                        Variavel(8) = TabSql("Dup.Pedido") & "/" & TabSql("NroPcl")
        '                        Variavel(9) = ""
        '                    End If
        '                 End If
                         
                         ImprimeLinha 0, 0, Chr(27) & Chr(48), Len(Chr(27) & Chr(48)), False, False
            
                         LinhaAnt = 0
                         Conf.MoveFirst
                         While Not Conf.Eof
                                If Conf!Linha <> LinhaAnt Then
                                    For X = 1 To Conf!Linha - LinhaAnt
                                        If y = 3 Then
                                            If Conf!NomeCampo <> "%7" Then
                                                  ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                            End If
                                        Else
                                            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                        End If
                                    Next
                                End If
                                LinhaAnt = Conf!Linha
                                If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                       If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                           ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                           If Mid(Conf("Formato"), 1, 1) = "+" Then
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           Else
                                                 ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           End If
                                       End If
                                Else
                                       If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                            ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                       Else
                                            If Mid(Conf("Formato"), 1, 1) = "+" Then
                                               ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                            Else
                                               If IsNull(TabSql(Conf!NomeCampo)) Then
                                                  ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               Else
                                                  ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                               End If
                                           End If
                                       End If
                                End If
                                Conf.MoveNext
                         Wend
'                         If Copias = 3 Then
'                               If y = 1 Then
'                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
'                                       ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
        '                               ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
'                                End If
'                         End If
        
        '             Next
                     TabSql.MoveNext
'                     If Copias = 3 Then
'                        ImprimeLinha 0, 0, Chr(12), Len(Chr(12)), False, False ' Salta Pagina
'                     End If
                     
               Loop
               ImprimeLinha 0, 0, Chr(27) & Chr(50), Len(Chr(27) & Chr(50)), False, False
               
        '       For x = 1 To Conf!Linha - LinhaAnt
        '            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
        '       Next
                Fim True
                If MsgBox("Boleto impresso com sucesso ?", vbYesNo, App.Title) = vbYes Then
                    TabSql.MoveFirst
                    Do Until TabSql.Eof
'                        If FormBaixarDupli.DatDados(0).Recordset("PgtTemp") = -1 Then
                            TabSql.Edit
                            TabSql.Fields("Tipo_Dup") = "DP"
'                            FormBaixarDupli.DatDados(0).Recordset.Fields("DtVenc") = Format(MebVenc, "mm/dd/yyyy")
                            TabSql.Update
                            
 '                       End If
                        TabSql.MoveNext
                    Loop
                    Unload Me

                End If
        End If
End If

End Sub

Private Sub BtImprimir_Click()
Dim TabSql As Recordset
Dim X As Integer
Dim y As Integer
Dim Copias As Integer
Dim Conf As Recordset
Dim Multa As Currency
Dim Juros As Currency
Dim VarVenc As String
Dim DtProtesto As String

If Opt(0).value = True Then
    Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, DUP.Pedido, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC, DUP.*, DUP.Cond FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE (((DUP.Pedido)=  " & TxtPedido & " ) AND ((DUP.Cond)=False))")
ElseIf Opt(1).value = True Then
'        Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC, DUP.*, DUP.Cond FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE (((DUP.Cond)=False) AND ((DUP.Documento)= " & """" & TxtDoc & """" & "))")
        Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC, DUP.*, DUP.Cond FROM DUP INNER JOIN ENT ON DUP.Entidade = ENT.CodEntidade  WHERE (((DUP.Cond)=False) AND ((DUP.Documento)= " & """" & TxtDoc & """" & "))" & " AND Dup.Entidade = " & TxtEnt)
    
    Else
        Set TabSql = Banco.OpenRecordset("SELECT DUP.CodDup, PED.NroNt, DUP.Pedido, ENT.Classificacao, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Cep, ENT.Uf, ENT.CPF, ENT.CGC, DUP.*, DUP.Cond, DUP.NroPcl FROM DUP INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON DUP.Pedido = PED.CodPed WHERE DUP.CodDup = " & TxtPedido & " AND  DUP.Cond = False  ")
End If

TabSql.Requery
If TabSql.RecordCount <= 0 Then
    MsgBox "Informações inconsistentes ", vbInformation, App.Title
    Exit Sub
End If
TabSql.MoveFirst

If Opt(0).value = True Then

    Set Conf = Banco.OpenRecordset("Select * From Cnt Where TipoDados = 10 Order by Codigo")
    Conf.Requery
    If Conf.RecordCount <= 0 Then
        MsgBox "Boleto não configurado !", vbInformation, App.Path
        Exit Sub
    End If
    Conf.MoveFirst
    
    If OptTipo(0).value = True Then
        Copias = 1
    Else
        Copias = 3
    End If
    Variavel(1) = "" 'Aceite
    Variavel(2) = "" 'Instruções 1
    Variavel(3) = "" 'Instruções 2
    Variavel(4) = "" 'Protesto
    Variavel(5) = "" 'Obs
    Variavel(6) = "" 'Cgc/Cpf
    Variavel(7) = "" 'Pulo
    Variavel(8) = "" ' Data Processamento
    
    If TabSql.RecordCount <> 0 Then
           Iniciar True 'inicia o registro onde sera gravado as informacoes
           
           Do While Not TabSql.Eof
                 For y = 1 To Copias
                     If ChkAceite.value = 0 Then
                        Variavel(1) = "N"
                     Else
                        Variavel(1) = "S"
                     End If
                     Variavel(8) = Date
                     Juros = TabSql!Valor * ((TxtMora / 100) / 30)
                     Multa = TabSql!Valor * (TxtMulta / 100)
                     VarVenc = TabSql!DtVenc + CInt(TxtPrazo)
                     DtProtesto = CDate(VarVenc) + CInt(TxtProtesto)
                     If ChkImprimir.value = 1 Then
                        If Juros <> 0 Then
                           Variavel(2) = "ATE " & TabSql!DtVenc & " COBRAR " & Format(TabSql!Valor, "#,##0.00") & " APOS " & TabSql!DtVenc & " ATE " & VarVenc
                           Variavel(3) = "COBRAR " & Format(TabSql!Valor, "#,##0.00") & " ACRESCIDO DE MULTA DIARIA DE " & Format(Multa, "#,##0.00") & " E MORA DIARIA DE " & Format(Juros, "#,##0.000")
                        End If
                        If TxtProtesto <> 0 Then
                           Variavel(4) = "** PROTESTAR APOS O DIA  " & DtProtesto
                        End If
                     Else
                        Variavel(2) = TxtInf1
                        Variavel(3) = TxtInf2
                        Variavel(4) = TxtInf3
                     End If
                     If TabSql!Classificacao = 2 Then
                        If TabSql!CGC <> "" Then
                            Variavel(6) = Format(TabSql!CGC, "@@.@@@.@@@/@@@@-@@")
                        Else
                            Variavel(6) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                        End If
                     ElseIf TabSql!Classificacao = 1 Then
                        Variavel(6) = Format(TabSql!Cpf, "@@@.@@@.@@@-@@")
                     End If
                     
                     ImprimeLinha 0, 0, Chr(27) & Chr(48), Len(Chr(27) & Chr(48)), False, False
        
                     LinhaAnt = 0
                     Conf.MoveFirst
                     While Not Conf.Eof
                            If Conf!Linha <> LinhaAnt Then
                                For X = 1 To Conf!Linha - LinhaAnt
                                    If y = 3 Then
                                        If Conf!NomeCampo <> "%7" Then
                                              ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                        End If
                                    Else
                                        ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                    End If
                                Next
                            End If
                            LinhaAnt = Conf!Linha
                            If Mid(Conf!NomeCampo, 1, 1) = "%" Then
                                   If Conf("Formato") <> "" And Mid(Conf("Formato"), 1, 1) <> "+" Then
                                       ImprimeLinha 0, Conf!Coluna, Format(Variavel(Mid(Conf!NomeCampo, 2, Len(Conf!TamanhoCampo))), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                   Else
                                       If Mid(Conf("Formato"), 1, 1) = "+" Then
                                             ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo))))) & Mid(Conf("Formato"), 2, Len(Conf("Formato")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                       Else
                                             ImprimeLinha 0, Conf!Coluna, CStr(Variavel(CCur((Mid(Conf!NomeCampo, 2, Len(Conf!NomeCampo)))))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                       End If
                                   End If
                            Else
                                   If Conf("Formato") <> "" And Mid(Conf("NomeCampo"), 1, 1) <> "+" Then
                                        ImprimeLinha 0, Conf!Coluna, Format(TabSql(Conf!NomeCampo), CStr(Conf!Formato)), Conf!TamanhoCampo, Conf!Condensado, Conf!Alinhar
                                   Else
                                        If Mid(Conf("Formato"), 1, 1) = "+" Then
                                           ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo) & Mid(Conf("Formato"), 2, Len(Conf("TamanhoCampo")))), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                        Else
                                           If IsNull(TabSql(Conf!NomeCampo)) Then
                                              ImprimeLinha 0, Conf!Coluna, Space(Conf("TamanhoCampo")), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           Else
                                              ImprimeLinha 0, Conf!Coluna, CStr(TabSql(Conf!NomeCampo)), Conf("TamanhoCampo"), Conf("Condensado"), Conf("Alinhar")
                                           End If
                                       End If
                                   End If
                            End If
                            Conf.MoveNext
                     Wend
                     If Copias = 3 Then
                           If y = 1 Then
                                   ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                                   ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
    '                               ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
                            End If
                     End If
    
                 Next
                 TabSql.MoveNext
                 If Copias = 3 Then
                    ImprimeLinha 0, 0, Chr(12), Len(Chr(12)), False, False ' Salta Pagina
                 End If
                 
           Loop
           ImprimeLinha 0, 0, Chr(27) & Chr(50), Len(Chr(27) & Chr(50)), False, False
           
    '       For x = 1 To Conf!Linha - LinhaAnt
    '            ImprimeLinha 1, 0, "", 1, False, False 'imprime uma linha em branco, toda vez que é efetuado o for next da linha
    '       Next
            Fim True
            If MsgBox("Boleto impresso com sucesso ?", vbYesNo, App.Title) = vbYes Then
                Unload Me
            End If
           
    End If
Else



End If

End Sub


Private Sub BtSair_Click()
Unload Me
End Sub

Private Sub ChkImprimir_Click()
Atualiza
End Sub

Private Sub Form_Activate()
Atualiza

MeLeft = (MDIPrincipal.Width - Me.Width) / 2
Me.Top = (MDIPrincipal.Height - Me.Height) / 6
MebDtSaida = Date
MebPcl = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim TipoBol As String

TipoBol = LerINI("BOLETO", "Tipo", App.Path & "\config.ini")

If TipoBol = "A4" Then
    OptTipo(1).value = True
Else
    OptTipo(0).value = True
End If
End Sub

Private Sub TxtPrazo_GotFocus()
If TxtPrazo = "" Then TxtPrazo = 0
End Sub

Private Sub TxtMulta_GotFocus()
If TxtMulta = "" Then TxtMulta = 0
End Sub

Private Sub TxtMora_GotFocus()
If TxtMora = "" Then TxtMora = 0
End Sub


Private Sub TxtProtesto_GotFocus()
If TxtProtesto = "" Then TxtProtesto = 0
End Sub


Private Sub TxtPrazo_LostFocus()
Atualiza
End Sub

Private Sub TxtMulta_LostFocus()
Atualiza
End Sub

Private Sub TxtMora_LostFocus()
Atualiza
End Sub

Private Sub TxtProtesto_LostFocus()
Atualiza
End Sub

Private Sub Atualiza()

If TxtPrazo = "" Then
    Exit Sub
End If

If TxtMulta = "" Then
    Exit Sub
End If

If TxtMora = "" Then
    Exit Sub
End If

If TxtProtesto = "" Then
    Exit Sub
End If

If ChkImprimir.value = 1 Then
    If MebValor = "" Then
        MebValor = 0
    End If
    Juros = Format(MebValor * ((TxtMora / 100) / 30), "#,##0.00")
    Multa = Format(MebValor * (TxtMulta / 100), "#,##0.00")
    VarVenc = CDate(MebVenc) + CInt(TxtPrazo)
    DtProtesto = CDate(VarVenc) + CInt(TxtProtesto)
    If Juros <> 0 Then
'       TxtInf1 = "ATE " & MebVenc & " COBRAR " & Format(MebValor, "#,##0.00") & " APOS " & MebVenc & " ATE " & VarVenc
'       TxtInf2 = "COBRAR " & Format(MebValor, "#,##0.00") & " ACRESCIDO DE MULTA DIARIA DE " & Format(Multa, "#,##0.00") & " E MORA DIARIA DE " & Format(Juros, "#,##0.000")
        TxtInf1 = " APOS O VENCIMENTO SERA COBRADO JUROS DE  " & Format(TxtMora / 30, "#,##0.00") & " % OU  R$ " & Format(Juros, "#,##0.00") & " AO DIA"
        TxtInf2 = ""
        TxtInf3 = ""
    End If
    If TxtProtesto <> 0 Then
       TxtInf3 = "** PROTESTAR APOS O DIA  " & DtProtesto
    Else
       TxtInf3 = ""
    End If
Else
    TxtInf1 = "DUPLICATA DESCONTADA NO BANCO: COBRANCA ESCRITURAL"
    TxtInf2 = "NAO CONCEDER DESCONTO"
    TxtInf3 = "PROTESTAR APOS O QUINTO DIA UTIL DO VENCIMENTO"
End If

End Sub
