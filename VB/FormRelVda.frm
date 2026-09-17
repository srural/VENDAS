VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormRelVda 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8415
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9690
   Begin VB.CheckBox ChkPositivo 
      Caption         =   "Negativo/Positivo"
      Height          =   240
      Left            =   225
      TabIndex        =   42
      Top             =   7740
      Width           =   1725
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -270
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   39
      Top             =   8100
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
      TabIndex        =   37
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatórios de Pedidos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   465
         Left            =   90
         TabIndex        =   38
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data"
      Height          =   1185
      Left            =   5670
      TabIndex        =   25
      Top             =   675
      Width           =   3615
      Begin VB.OptionButton OptData 
         Caption         =   "&En&trega"
         Height          =   285
         Index           =   2
         Left            =   2430
         TabIndex        =   32
         Top             =   810
         Width           =   1005
      End
      Begin VB.OptionButton OptData 
         Caption         =   "&Fatura"
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   31
         Top             =   810
         Width           =   870
      End
      Begin VB.OptionButton OptData 
         Caption         =   "&Emissão"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   30
         Top             =   810
         Value           =   -1  'True
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker DtData1 
         Height          =   330
         Left            =   135
         TabIndex        =   26
         Top             =   405
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   91750401
         CurrentDate     =   36641
      End
      Begin MSComCtl2.DTPicker DtData2 
         Height          =   330
         Left            =   1845
         TabIndex        =   27
         Top             =   405
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   91750401
         CurrentDate     =   36641
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inicial"
         Height          =   195
         Left            =   135
         TabIndex        =   29
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Final"
         Height          =   195
         Left            =   1845
         TabIndex        =   28
         Top             =   180
         Width           =   330
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenado por"
      Height          =   885
      Left            =   5670
      TabIndex        =   15
      Top             =   5385
      Width           =   3615
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Cidade"
         Height          =   375
         Index           =   4
         Left            =   1260
         TabIndex        =   22
         Top             =   180
         Width           =   1185
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Valor"
         Height          =   375
         Index           =   5
         Left            =   2595
         TabIndex        =   21
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Grupo"
         Height          =   375
         Index           =   3
         Left            =   2595
         TabIndex        =   19
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Entidade"
         Height          =   375
         Index           =   2
         Left            =   1260
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Produto"
         Height          =   375
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Top             =   480
         Width           =   1050
      End
      Begin VB.OptionButton OptOrdem 
         Caption         =   "Pedido"
         Height          =   375
         Index           =   0
         Left            =   225
         TabIndex        =   16
         Top             =   180
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.CommandButton CmdConfImpr 
      Caption         =   "&Configura"
      Height          =   315
      Left            =   5700
      TabIndex        =   14
      Top             =   7695
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos com"
      Height          =   1185
      Left            =   5670
      TabIndex        =   11
      Top             =   6345
      Width           =   3615
      Begin VB.CheckBox ChkVenc 
         Caption         =   "A Entregar"
         Height          =   240
         Left            =   1935
         TabIndex        =   41
         Top             =   855
         Width           =   1500
      End
      Begin VB.CheckBox PedCom 
         Caption         =   "Cupom / Sat"
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   810
         Width           =   1380
      End
      Begin VB.CheckBox PedCom 
         Caption         =   "Pedidos"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   510
         Width           =   1155
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1935
         TabIndex        =   24
         Top             =   495
         Width           =   1455
      End
      Begin VB.CheckBox PedCom 
         Caption         =   "Ped. Referência"
         Height          =   285
         Index           =   1
         Left            =   1935
         TabIndex        =   13
         Top             =   180
         Width           =   1485
      End
      Begin VB.CheckBox PedCom 
         Caption         =   "Nota Fiscal / Sat"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gerais"
      Height          =   3495
      Left            =   5670
      TabIndex        =   1
      Top             =   1875
      Width           =   3615
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Familia"
         Height          =   375
         Index           =   9
         Left            =   270
         TabIndex        =   40
         Top             =   3060
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Complemento"
         Height          =   375
         Index           =   8
         Left            =   270
         TabIndex        =   35
         Top             =   2750
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Tributação"
         Height          =   375
         Index           =   7
         Left            =   270
         TabIndex        =   34
         Top             =   2440
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Fornecedor"
         Height          =   375
         Index           =   6
         Left            =   270
         TabIndex        =   23
         Top             =   1510
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Cidade"
         Height          =   375
         Index           =   5
         Left            =   270
         TabIndex        =   20
         Top             =   2130
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Vendedor(a)"
         Height          =   375
         Index           =   4
         Left            =   270
         TabIndex        =   6
         Top             =   1820
         Width           =   2310
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Ordenado por Pedido"
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Produto"
         Height          =   375
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   580
         Width           =   2040
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Identificando a Entidade"
         Height          =   375
         Index           =   2
         Left            =   270
         TabIndex        =   3
         Top             =   890
         Width           =   2355
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Agrupado por Grupo"
         Height          =   375
         Index           =   3
         Left            =   270
         TabIndex        =   2
         Top             =   1200
         Width           =   1905
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Top             =   7695
      Width           =   1620
   End
   Begin VB.ListBox ListaOper 
      Appearance      =   0  'Flat
      Height          =   3345
      ItemData        =   "FormRelVda.frx":0000
      Left            =   210
      List            =   "FormRelVda.frx":0037
      TabIndex        =   7
      Top             =   990
      Width           =   5190
   End
   Begin VB.ListBox ListaSel 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   195
      TabIndex        =   8
      Top             =   4815
      Width           =   5220
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleção"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   4545
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operação"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   720
      Width           =   705
   End
End
Attribute VB_Name = "FormRelVda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Selecao As Recordset
Public VarBco As Recordset
Public VarTotal As Currency
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdConfImpr_Click()
MDIPrincipal.ComRel.ShowPrinter
End Sub

Private Sub CmdImprimir_Click()
Dim Ano1, Mes1, Dia1, Ano2, Mes2, Dia2 As Integer
Dim data1 As String
Dim Data2 As String
Dim Periodo As String
Dim SQL As String
Dim CamposSql As String
Dim CamposGroup As String
Dim Ordenado As String
Dim varFormRelatorio As Form
Dim Formula As String
Dim Titulo As String
Dim VarDesc As Recordset
Dim VarData As String

'On Error GoTo ErroRelVda

Formula = ""
data1 = Format(DtData1, "mm/dd/yyyy")
Data2 = Format(DtData2, "mm/dd/yyyy")

If OptData(0).value = True Then
        VarData = "DataEmiss"
    ElseIf OptData(1).value = True Then
        VarData = "DtFat"
    Else
        VarData = "DtSaida"
End If


SQL = "UPDATE PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto SET ITP.CustoDia = [prd].[custo] WHERE (((ITP.CustoDia)=0))"
Banco.Execute SQL

Periodo = "Período de " & DtData1 & " a " & DtData2
'MDIPrincipal.RptRel.Formulas(0) = "Formulas = '" & Periodo & "' "

If OPt(0).value = True Then
        Set varFormRelatorio = FrmRelConf
    ElseIf OPt(1).value = True Then
            Set varFormRelatorio = FrmRelVda
        ElseIf OPt(2).value = True Then
                If ListaSel.ListIndex = 1 Then
                    Set varFormRelatorio = FormRelVda
                ElseIf ListaSel.ListIndex = 2 Then
                    Set varFormRelatorio = FrmRelCliGru
                Else
                    Set varFormRelatorio = FrmRelVdaCli
                End If
            ElseIf OPt(3).value = True Then
                    If ListaSel.ListIndex = 0 Then
                        Set varFormRelatorio = FrmRelTotGru
                    Else
                        Set varFormRelatorio = FrmRelVdaGru
                    End If
                ElseIf OPt(4).value = True Then
                        Set varFormRelatorio = FrmRelComiss
                    ElseIf OPt(5).value = True Then
                            If ListaSel.ListIndex = 0 Then
                                Set varFormRelatorio = FrmRelVdaCidade
                            Else
                                Set varFormRelatorio = FrmRelComiss01
                            End If
                        ElseIf OPt(6).value = True Then
                                If ListaSel.ListIndex = 0 Then
                                    Set varFormRelatorio = FrmRelVdaForPed
                                ElseIf ListaSel.ListIndex = 1 Then
                                        Set varFormRelatorio = FrmRelVdaFor
                                    ElseIf ListaSel.ListIndex = 2 Then
                                            Set varFormRelatorio = FrmRelVdaForVend
                                        ElseIf ListaSel.ListIndex = 3 Then
                                                Set varFormRelatorio = FrmRelVdaForCid
                                           Else
                                                Set varFormRelatorio = FrmRelVdaForGru
                                End If
                        ElseIf OPt(7).value = True Then
                            Set varFormRelatorio = FrmRelVdaTrib
                        ElseIf OPt(8).value = True Then
                            Set varFormRelatorio = FrmRelVdaCompl
                        Else
                            Set varFormRelatorio = FrmRelVdaFam
                        
                        
                        
End If

If ListaOper.ListIndex = 0 Then
    
    Set VarDesc = Banco.OpenRecordset("SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  (Operacao = 1 or Operacao = 2 or Operacao = 3)   and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#")
    VarDesc.Requery
'    SQL = "SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  (Operacao = 1 or Operacao = 2 or Operacao = 3)   and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#"
    If ListaSel.ListIndex = 0 Or ListaSel.Enabled = False Then
        Formula = VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
    Else
'        Set VarDesc = Banco.OpenRecordset("SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  Operacao = " & ListaOper.ListIndex & "  and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#")
        If OPt(2).value = True Then
              If ListaSel.ListIndex = 1 Then
                    Formula = VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
              ElseIf ListaSel.ListIndex = 2 Then
                    Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
              Else
                    Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) and Entidade = " & CLng(Left(ListaSel.Text, 5))
              End If
           ElseIf OPt(3).value = True Then
                  If ListaSel.ListIndex = 0 Then
                      Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
                  ElseIf ListaSel.ListIndex = 1 Then
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                      Else
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) and Grupo = " & CLng(Left(ListaSel.Text, 4))
                      End If
              ElseIf OPt(4).value = True Then
                    If ListaSel.ListIndex = 1 Then
                        Set varFormRelatorio = FrmRelVdaFunGru
                        Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                    ElseIf ListaSel.ListIndex = 2 Then
                            Set varFormRelatorio = FrmRelComiss01
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "#  and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                        ElseIf ListaSel.ListIndex = 3 Then
                            Set varFormRelatorio = FrmRelComiss02
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                        ElseIf ListaSel.ListIndex = 4 Then
                            Set varFormRelatorio = FrmRelVdaFunCid
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "#  and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                        Else
                            Set varFormRelatorio = FrmRelComiss01 ' FrmRelVdaFunGru
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) and Vendedor.CodEntidade = " & CLng(Left(ListaSel.Text, 6)) & " And Ped.Total <> 0"
                    End If
                  ElseIf OPt(5).value = True Then
                          If ListaSel.ListIndex = 0 Then
                               Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
                          Else
                                Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) " ' and Ent.Cidade =  " & "'" & ListaSel.Text & "'"
                          End If
                      Else
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
        End If
    End If
Else
    Set VarDesc = Banco.OpenRecordset("SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  Operacao = " & ListaOper.ListIndex - 1 & "  and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#")
    VarDesc.Requery

'    SQL = "SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  Operacao = " & ListaOper.ListIndex & "  and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#"
    If ListaSel.ListIndex = 0 Or ListaSel.Enabled = False Then
        Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
'        Set VarDesc = Banco.OpenRecordset("SELECT  Sum(PED.Desconto) AS SomaDeDesconto FROM PED WHERE  (Operacao = 1 or Operacao = 2 or Operacao = 3)   and  " & VarData & " >= #" & data1 & "# AND " & VarData & " <= #" & Data2 & "#")
    Else

        If OPt(2).value = True Then
              If ListaSel.ListIndex = 1 Then
                    Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
              ElseIf ListaSel.ListIndex = 2 Then
                    Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and (Operacao = 1 or Operacao = 2 or Operacao = 3) "
              Else
                    Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1 & " and Entidade = " & CLng(Left(ListaSel.Text, 5))
              End If
           ElseIf OPt(3).value = True Then
                  If ListaSel.ListIndex = 0 Then
                      Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao  = " & ListaOper.ListIndex - 1
                  ElseIf ListaSel.ListIndex = 1 Then
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
                      Else
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao  = " & ListaOper.ListIndex - 1 & " and Grupo = " & CLng(Left(ListaSel.Text, 4))
                      End If
              ElseIf OPt(4).value = True Then
                    If ListaSel.ListIndex = 1 Then
                        Set varFormRelatorio = FrmRelVdaFunGru
                        Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao =  " & ListaOper.ListIndex - 1
                    ElseIf ListaSel.ListIndex = 2 Then
                            Set varFormRelatorio = FrmRelComiss01
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
                        ElseIf ListaSel.ListIndex = 3 Then
                            Set varFormRelatorio = FrmRelComiss02
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
                        Else
                            Set varFormRelatorio = FrmRelComiss01  'FrmRelVdaFun
                            Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1 & " and Vendedor.Codentidade = " & CLng(Left(ListaSel.Text, 5))
                    End If
                  ElseIf OPt(5).value = True Then
                          If ListaSel.ListIndex = 0 Then
                               Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
                          Else
                                Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1 & " and Ent.Cidade =  " & "'" & ListaSel.Text & "'"
                          End If
                    Else
                          Formula = VarData & "  >= #" & data1 & "# AND  " & VarData & "  <= #" & Data2 & "# and Operacao = " & ListaOper.ListIndex - 1
        End If
    
    End If

End If
       
If PedCom(0).value = 1 Then
    Formula = Formula & " and NroNt > 0 "
End If

If ChkVenc.value = 1 Then
    Formula = Formula & " and   IsNull(Ped.DtSaida) "
    ' = " & """" & "#__/__/____#" & """"
End If

If PedCom(1).value = 1 Then
    Formula = Formula & " and NroPedido <> '' "
End If

If PedCom(3).value = 1 Then
    Formula = Formula & " and (Cfo = 'CUPOM' or Cfo = 'SAT') "
End If

    
If TxtRef <> "" Then
    Formula = Formula & " and NroPedido = '" & TxtRef & "' "
End If

If PedCom(2).value = 1 Then
    Formula = Formula & " and NroNt = 0 "
End If

    If varFormRelatorio.Name = "FrmRelConf" Then
'        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
'        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ENT.Nome, PED.Operacao "
        CamposSql = " SELECT  ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, PRD.Descrição_Produto,Itp.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp, Ped.Desconto, First(PED.Obs) AS PrimeiroDeObs FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY  ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto,PRD.Descrição_Produto, Itp.Complemento, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp, ENT.Nome, PED.Operacao, Ped.Desconto "
'        Ordenado = " Order by CodPed, CodItp "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVda" Then
        If ChkPositivo = 1 Then
            CamposSql = " SELECT ITP.Produto, PRD.Descrição_Produto, Prd.Estoque, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto , Sum([ITP].[QTD]*[ITP].[CUSTODIA]) AS Somadecusto, TABELA.MED_NEGPOS, Avg([ITP].[VENDADIA]) AS MediaVendaDia  FROM  (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) INNER JOIN TABELA ON PRD.CodAbc = TABELA.MED_ABC "
            CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto, Prd.Estoque, TABELA.MED_NEGPOS HAVING TABELA.MED_NEGPOS = " & """" & "SIM" & """"
    '        Ordenado = " Order by Produto "
            SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        Else
            CamposSql = " SELECT ITP.Produto, PRD.Descrição_Produto, Prd.Estoque, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto , Sum([ITP].[QTD]*[ITP].[CUSTODIA]) AS Somadecusto, Avg([ITP].[VENDADIA]) AS MediaVendaDia FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto, Prd.Estoque  "
    '        Ordenado = " Order by Produto "
            SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        End If
    ElseIf varFormRelatorio.Name = "FrmRelVdaCli" Then
       If OptOrdem(0).value = True Then
            CamposSql = " SELECT Ped.CodPed, ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, Itp.Loja, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ITP.CodItp FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY Ped.CodPed, ENT.CodEntidade, ENT.Nome, PED.NroPedido, PED.NroNt, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, Itp.Loja, PRD.Descrição_Produto, ITP.Qtd, ITP.ValorUnit, ITP.Valor, ENT.Nome, PED.Operacao, ITP.CodItp "
'        Ordenado = " Order by PRD.Descrição_Produto "
        Else
            CamposSql = " SELECT  ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, ITP.Loja, PRD.Descrição_Produto,First(Ped.CodPed) AS PrimeiroDeCodPed, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Sum(Ped.Desconto) as SomaDesconto  FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON Ped.CodPed =ITP.Pedido "
            CamposGroup = " GROUP BY  ENT.CodEntidade, ENT.Nome, ENT.Endereco, ENT.Nro, ENT.Cidade, ENT.Fone, ITP.Produto, ITP.Loja, PRD.Descrição_Produto "
        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        
        If Not IsNull(VarDesc!SomaDeDesconto) Then
            FrmRelVdaCli.VarTotDesc = VarDesc!SomaDeDesconto
        Else
            FrmRelVdaCli.VarTotDesc = 0
        End If
        
    ElseIf varFormRelatorio.Name = "FrmRelVdaGru" Then
        CamposSql = " SELECT PRD.Estoque, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM GRU INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON GRU.CodGru = PRD.Grupo "
        CamposGroup = " GROUP BY Prd.Estoque, GRU.CodGru, GRU.Descrição_Grupo, PRD.CodPrd, PRD.Descrição_Produto "
'        Ordenado = " Order by CodGru, Descrição_Produto "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelTotGru" Then
        CamposSql = " SELECT GRU.CodGru, GRU.Descrição_Grupo, Sum(Ped.Desconto) as SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, (somadevalor / somadeqtd) AS MédiaDeValor FROM PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY GRU.CodGru, GRU.Descrição_Grupo "
        If OptOrdem(5).value = True Then
            Ordenado = " ORDER BY Sum(ITP.VlrVdaDia) DESC "
        Else
            Ordenado = " Order by CodGru, Descrição_Grupo "
        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        
        Set VarBco = Banco.OpenRecordset(SQL)
        If VarBco.RecordCount > 0 Then
            VarBco.MoveFirst
            VarTotal = 0
            While Not VarBco.Eof
                VarTotal = VarTotal + CCur(VarBco("SomaDeVlrVdaDia"))
                VarBco.MoveNext
            Wend
        End If
        If Not IsNull(VarDesc!SomaDeDesconto) Then
            FrmRelTotGru.VarTotDesc = VarDesc!SomaDeDesconto
        Else
            FrmRelTotGru.VarTotDesc = 0
        End If
                
    ElseIf varFormRelatorio.Name = "FrmRelComiss" Then
        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, Sum(Ped.Desconto) as SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Vendedor) INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome "
'        If OptOrdem(4).value Then
'            Ordenado = " Order by Cidade "
'        Else
            Ordenado = " Order by Nome "
'        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        
'        If Not IsNull(VarDesc!SomaDeDesconto) Then
'            FrmRelComiss.VarTotDesc = VarDesc!SomaDeDesconto
'        End If
                
    ElseIf varFormRelatorio.Name = "FrmRelVdaFunGru" Then
'        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo, Sum(ITP.Qtd) AS SomaDeQtd, Avg(ITP.ValorUnit) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.Comissao) AS SomaDeComissao, Count(PED.CodPed) AS ContarDeCodPed,  sum(PED.Desconto) AS SomaDeDesconto FROM ENT INNER JOIN (GRU INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON GRU.CodGru = PRD.Grupo) ON ENT.CodEntidade = PED.Vendedor "
'        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo "
        
        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.Comissao) AS SomaDeComissao, Count(PED.CodPed) AS ContarDeCodPed,  PED.Desconto FROM ENT INNER JOIN (GRU INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON GRU.CodGru = PRD.Grupo) ON ENT.CodEntidade = PED.Vendedor "
        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo, PED.Desconto "
        Ordenado = " Order by Nome, CodGru  "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        If IsNull(VarDesc!SomaDeDesconto) Then
            FrmRelVdaFunGru.VarTotDesc = 0
        Else
            FrmRelVdaFunGru.VarTotDesc = VarDesc!SomaDeDesconto
        End If

    ElseIf varFormRelatorio.Name = "FrmRelVdaFunCid" Then
'        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo, Sum(ITP.Qtd) AS SomaDeQtd, Avg(ITP.ValorUnit) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.Comissao) AS SomaDeComissao, Count(PED.CodPed) AS ContarDeCodPed,  sum(PED.Desconto) AS SomaDeDesconto FROM ENT INNER JOIN (GRU INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON GRU.CodGru = PRD.Grupo) ON ENT.CodEntidade = PED.Vendedor "
'        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo "
        
        CamposSql = "Select Sum(PED.Desconto) AS SomaDeDesconto, Vendedores.ENT.Nome, ENT.Cidade, First(Vendedores.CodEntidade) AS PrimeiroDeCodEntidade, Count(PED.CodPed) AS ContarDeCodPed, Sum(PED.Comissao) AS SomaDeComissao, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, (somadevalor/somadeqtd) AS MédiaDeValor FROM ((PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN Vendedores ON PED.Vendedor = Vendedores.CodEntidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido  "
        CamposGroup = " GROUP BY Vendedores.ENT.Nome, ENT.Cidade "
        Ordenado = " Order by Vendedores.ENT.Nome "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        
'        SQL = "SELECT Sum(PED.Desconto) AS SomaDeDesconto, Vendedores.ENT.Nome, ENT.Cidade, First(Vendedores.CodEntidade) AS PrimeiroDeCodEntidade, Count(PED.CodPed) AS ContarDeCodPed, Sum(PED.Comissao) AS SomaDeComissao, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, (somadevalor/somadeqtd) AS MédiaDeValor FROM ((PED INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade) INNER JOIN Vendedores ON PED.Vendedor = Vendedores.CodEntidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)>=#10/1/2016#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3)) GROUP BY Vendedores.ENT.Nome, ENT.Cidade ORDER BY  Vendedores.ENT.Nome"

        
        If IsNull(VarDesc!SomaDeDesconto) Then
            FrmRelVdaFunGru.VarTotDesc = 0
        Else
            FrmRelVdaFunGru.VarTotDesc = VarDesc!SomaDeDesconto
        End If

    ElseIf varFormRelatorio.Name = "FrmRelComiss01" Then
        CamposSql = "SELECT Vendedor.CodEntidade, Vendedor.Nome, PED.CodPed, PED.DataEmiss, PED.DtFat, PED.DtSaida, ENT.CodEntidade, ENT.Nome, PED.Total, PED.Desconto,  ENT.Cidade, sum(CCur(  [ITP].[Comissao] - (([ITP].[Comissao] * ([PED].[Desconto] /[PED].[Total] )))))  AS Valor, Sum(ITP.Comissao) AS SomaDeComissao FROM (ENT INNER JOIN (ENT AS Vendedor INNER JOIN PED ON Vendedor.CodEntidade = PED.Vendedor) ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido "
'        CamposSql = "SELECT Vendedor.CodEntidade, Vendedor.Nome, PED.CodPed, PED.DataEmiss, PED.DtFat, PED.DtSaida, ENT.CodEntidade, ENT.Nome, PED.Total, PED.Desconto,  ENT.Cidade, sum(CCur(  [ITP].[Comissao] - (([ITP].[Comissao] * iif([PED].[Desconto] = 0,[PED].[TOTAL], ([PED].[Desconto] /[PED].[Total]) )))))  AS Valor, Sum(ITP.Comissao) AS SomaDeComissao FROM (ENT INNER JOIN (ENT AS Vendedor INNER JOIN PED ON Vendedor.CodEntidade = PED.Vendedor) ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY Vendedor.CodEntidade, Vendedor.Nome, PED.CodPed, PED.DataEmiss, PED.DtFat, PED.DtSaida, ENT.CodEntidade, ENT.Nome, PED.Total, PED.Desconto, Ent.Cidade  "
        
        If OptOrdem(4).value Then
            Ordenado = " Having Ent.Cidade =  " & "'" & ListaSel.Text & "'" & " Order by Ent.Cidade "
        Else
            Ordenado = " Order by Vendedor.Nome, Ent.Nome "
        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelComiss02" Then
        CamposSql = " SELECT ITP.Produto, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, Sum(ITP.Comissao) AS SomaDeComissao, ENT.CodEntidade, ENT.Nome FROM ENT INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON ENT.CodEntidade = PED.Vendedor "
        CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto, ENT.CodEntidade, ENT.Nome "
        Ordenado = " ORDER BY  ENT.CodEntidade , ITP.Produto "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaCidade" Then
'        CamposSql = " SELECT ENT.Cidade, ENT.Uf, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum(ITP.Qtd) AS SomaDeQtd, Count(PED.CodPed) AS ContarDeCodPed FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido "
        CamposSql = " SELECT First(ENT.Cidade) AS PrimeiroDeCidade, First(ENT.Uf) AS PrimeiroDeUf, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum(ITP.Qtd) AS SomaDeQtd, PED.CodPed FROM (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) INNER JOIN ITP ON PED.CodPed = ITP.Pedido "
        CamposGroup = " GROUP BY Ped.CodPed "
        If OptOrdem(5).value = True Then
            Ordenado = " ORDER BY Sum(ITP.VlrVdaDia) DESC "
        Else
            Ordenado = " ORDER BY First(ENT.Cidade), First(ENT.Uf) "
        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaFor" Then
        CamposSql = " SELECT ITP.Produto, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, ENT.CodEntidade, ENT.Nome FROM (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON Ped.Entidade = ENT.CodEntidade"
        CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto, ENT.CodEntidade, ENT.Nome "
        Ordenado = " ORDER BY  ENT.CodEntidade , ITP.Produto "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaForPed" Then
         CamposSql = " SELECT Sum(ITP.Valor) AS SomaDeValor, First(ENT.Nome) AS PrimeiroDeNome, PED.DataEmiss, PED.DtFat, PED.DtSaida, PED.CodPed, First(PED.Desconto) AS SomaDeDesconto FROM ENT INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON ENT.CodEntidade = PED.Entidade"
         CamposGroup = " GROUP BY ENT.CodEntidade, PED.DataEmiss, Ped.DtFat, Ped.DtSaida, PED.CodPed"
         Ordenado = " ORDER BY ENT.CodEntidade "
         SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaForVend" Then
         CamposSql = " SELECT Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, ENT.CodEntidade, ENT.Nome, ENT_1.CodEntidade, ENT_1.Nome FROM (PED INNER JOIN ((PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT AS ENT_1 ON PED.Vendedor = ENT_1.CodEntidade "
         CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, ENT_1.CodEntidade, ENT_1.Nome "
         Ordenado = " ORDER BY ENT.Nome, ENT_1.Nome "
         SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
         FrmRelVdaForVend.VarTotDesc = VarDesc!SomaDeDesconto
         
    ElseIf varFormRelatorio.Name = "FrmRelVdaForCid" Then
         CamposSql = " SELECT Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, ENT.CodEntidade, ENT.Nome, ENT_1.Cidade FROM (PED INNER JOIN ENT AS ENT_1 ON PED.Entidade = ENT_1.CodEntidade) INNER JOIN ((PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade) ON PED.CodPed = ITP.Pedido "
         CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, ENT_1.Cidade "
         Ordenado = " ORDER BY ENT.Nome, ENT_1.Cidade "
         SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaForGru" Then
         CamposSql = " SELECT Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, ENT.CodEntidade, ENT.Nome, PRD.Grupo, GRU.Descrição_Grupo, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia FROM (PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON PRD.Fabricante = ENT.CodEntidade  "
         CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, PRD.Grupo, GRU.Descrição_Grupo "
         Ordenado = " ORDER BY ENT.Nome,GRU.Descrição_Grupo  "
         SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FormRelVda" Then
        CamposSql = " SELECT ITP.Produto, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValorUnit, Sum(ITP.Valor) AS SomaDeValor, Avg(ITP.CustoDia) AS MédiaDeCusto, ENT.CodEntidade, ENT.Nome FROM (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) INNER JOIN ENT ON PED.Entidade = ENT.CodEntidade"
        CamposGroup = " GROUP BY ITP.Produto, PRD.Descrição_Produto, ENT.CodEntidade, ENT.Nome "
        Ordenado = " ORDER BY  ENT.CodEntidade , ITP.Produto "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        Set varFormRelatorio = FrmRelVdaFor
        
    ElseIf varFormRelatorio.Name = "FrmRelCliGru" Then
        CamposSql = " SELECT ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.Comissao) AS SomaDeComissao FROM ENT INNER JOIN (GRU INNER JOIN (PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido) ON GRU.CodGru = PRD.Grupo) ON ENT.CodEntidade = PED.Entidade "
        CamposGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, GRU.CodGru, GRU.Descrição_Grupo "
        Ordenado = " Order by Nome "
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
        If IsNull(VarDesc!SomaDeDesconto) Then
            FrmRelVdaFunGru.VarTotDesc = 0
        Else
            FrmRelVdaFunGru.VarTotDesc = VarDesc!SomaDeDesconto
        End If
    ElseIf varFormRelatorio.Name = "FrmRelVdaTrib" Then
            CamposSql = " SELECT Prd.Classe, PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, First(PRD.ClasseFiscal) AS PrimeiroDeClasseFiscal, PRD.SitTrib FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido  "
            CamposGroup = " GROUP BY Prd.Classe, PRD.CodPrd, PRD.Descrição_Produto, PRD.SitTrib "
            Ordenado = " ORDER BY PRD.SitTrib "
            SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaCompl" Then
            CamposSql = " SELECT PRD.CodPrd, PRD.Descrição_Produto, Sum(ITP.Qtd) AS SomaDeQtd, (somadevalor / somadeqtd) AS MédiaDeValor, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, PRD.Complemento  FROM PED INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON PED.CodPed = ITP.Pedido "
            CamposGroup = " GROUP BY PRD.CodPrd, PRD.Descrição_Produto, PRD.Complemento "
            Ordenado = " ORDER BY  PRD.Complemento, PRD.Descrição_Produto "
            SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    ElseIf varFormRelatorio.Name = "FrmRelVdaFam" Then
            CamposSql = " SELECT Sum(PED.Desconto) AS SomaDesconto, Sum(ITP.Qtd) AS SomaDeQtd, Sum(ITP.Valor) AS SomaDeValor, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia, Sum([Itp].[Qtd]*[Itp].[CustoDia]) AS Custo, Sum([Itp].[Valor]-([Itp].[Qtd]*[Itp].[CustoDia])) AS Lucro, (somadevalor/somadeqtd) AS MédiaDeValor, FAM.CodFam, FAM.Descrição_Familia FROM (PED INNER JOIN (GRU INNER JOIN (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) ON GRU.CodGru = PRD.Grupo) ON PED.CodPed = ITP.Pedido) INNER JOIN FAM ON GRU.Familia = FAM.CodFam "
            CamposGroup = " GROUP BY FAM.CodFam, FAM.Descrição_Familia "
            Ordenado = " ORDER BY FAM.Descrição_Familia "
            SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    End If
        
    If Ordenado = "" Then
        If OptOrdem(0).value = True Then
            If varFormRelatorio.Name = "FrmRelVdaCli" Then
                Ordenado = "" ' Order by PrimeiroDeCodPed"
            Else
                If OptData(2).value = True Then
                    Ordenado = " Order by Ped.DtSaida"
                Else
                    Ordenado = " Order by CodPed"
                End If
            End If
        ElseIf OptOrdem(1).value = True Then
            Ordenado = " Order by PRD.Descrição_Produto  "
        ElseIf OptOrdem(2).value = True Then
            Ordenado = " Order by Ent.Nome  "
        ElseIf OptOrdem(3).value = True Then
            Ordenado = " Order by CodGru, Descrição_Grupo "
        End If
        SQL = CamposSql & " Where " & Formula & CamposGroup & Ordenado
    
    End If
       
    Set varPesquisa(1) = Banco.OpenRecordset(SQL)
    Set varFormRelatorio.Relatorio.Recordset = varPesquisa(1)
    varFormRelatorio.rpCabTitulo.Caption = "Relatório de " & ListaOper.Text
    varFormRelatorio.rpCabPeriodo.Caption = Periodo
    varFormRelatorio.rpCabSubTitulo.Caption = NomeEmpresa
    If varFormRelatorio.Name = "FrmRelVdaCli" And OptOrdem(0).value = True Then
        FrmRelVdaCli.flag = "Entidade"
    Else
        FrmRelVdaCli.flag = ""
    End If
    varFormRelatorio.Config

Exit Sub

ErroRelVda:

        MsgBox Err.Description & Err.Number, vbCritical, Caminho

End Sub

Private Sub Form_Load()
Set Formulário = Me

Formulário.Left = (MDIPrincipal.Width - Formulário.Width) / 2
Formulário.Top = (MDIPrincipal.Height - Formulário.Height) / 6

DtData1 = Date
DtData2 = Date

ListaOper.Enabled = True
ListaSel.Enabled = False
ListaOper.ListIndex = 0
End Sub

Private Sub Opt_Click(Index As Integer)
Select Case Index
    Case 0
        ListaSel.Enabled = False
        OptOrdem(0).value = True
    Case 1
        ListaSel.Enabled = False
        OptOrdem(1).value = True
    Case 2
        OptOrdem(2).value = True
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Todos)")
        ListaSel.AddItem ("(Separado por Entidade)")
        ListaSel.AddItem ("(Separado por Grupo)")
        If ListaOper.ListIndex = 1 Then
            Set Selecao = Banco.OpenRecordset("select * from ent where tipo = 2 or tipo = 5 order by nome")
        Else
            Set Selecao = Banco.OpenRecordset("select * from ent where tipo = 1 or tipo >= 5 order by nome")
        End If
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodEntidade"), 5) & " - " & Selecao.Fields("Nome"))
                Selecao.MoveNext
            Wend
        End If
        ListaSel.ListIndex = 0
    Case 3
        OptOrdem(3).value = True
    
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Totalizado)")
        ListaSel.AddItem ("(Todos)")
        Set Selecao = Banco.OpenRecordset("select * from gru  order by Descrição_Grupo")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodGru"), 4) & " - " & Selecao("Descrição_Grupo"))
                Selecao.MoveNext
            Wend
         End If
         ListaSel.ListIndex = 0
    Case 4
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Todos)")
        ListaSel.AddItem ("(Separado por Grupo)")
        ListaSel.AddItem ("(Separado por Pedido)")
        ListaSel.AddItem ("(Separado por Produto)")
        ListaSel.AddItem ("(Separado por Cidade)")
        Set Selecao = Banco.OpenRecordset("select * from ent where tipo = 3 or tipo = 7 order by nome")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                ListaSel.AddItem (psCompString(1, Selecao("CodEntidade"), 5) & " - " & Selecao.Fields("Nome"))
                Selecao.MoveNext
            Wend
        End If
        ListaSel.ListIndex = 0
    Case 5
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Todas)")
        Set Selecao = Banco.OpenRecordset("select Cidade, uf from ent where tipo = 1 or tipo >= 5 GROUP BY ent.Cidade, ent.uf ORDER BY ent.Cidade, ent.uf")
        If Selecao.RecordCount <> 0 Then
            Selecao.MoveFirst
            While Not Selecao.Eof
                If Not IsNull(Selecao.Fields("Cidade")) Then
                    If Selecao.Fields("Cidade") <> "" Then
                        ListaSel.AddItem (Selecao.Fields("Cidade"))
                    End If
                End If
                Selecao.MoveNext
            Wend
        End If
        ListaSel.ListIndex = 0
        OptOrdem(4).value = True
    Case 6
        ListaSel.Enabled = True
        ListaSel.Clear
        ListaSel.AddItem ("(Separado por Pedido)")
        ListaSel.AddItem ("(Separado por Produto)")
        ListaSel.AddItem ("(Separado por Vendedor)")
        ListaSel.AddItem ("(Separado por Cidade)")
        ListaSel.AddItem ("(Separado por Grupo)")
        ListaSel.ListIndex = 0
End Select

End Sub
