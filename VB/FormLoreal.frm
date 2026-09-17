VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormLoreal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   2010
   ClientTop       =   1860
   ClientWidth     =   6945
   Icon            =   "FormLoreal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6945
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -4545
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   10
      Top             =   2115
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
      TabIndex        =   8
      Top             =   0
      Width           =   13155
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Boticario"
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
         TabIndex        =   9
         Top             =   45
         Width           =   8115
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodo"
      Height          =   945
      Left            =   360
      TabIndex        =   3
      Top             =   855
      Width           =   3720
      Begin MSMask.MaskEdBox MebDtI 
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   1755
         TabIndex        =   5
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   1710
         TabIndex        =   7
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Inicial"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.PictureBox PctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   0
      Left            =   150
      ScaleHeight     =   1230
      ScaleWidth      =   6480
      TabIndex        =   1
      Top             =   750
      Width           =   6480
      Begin VB.CommandButton CmbSugestao 
         Caption         =   "&Gera Arquivo"
         Height          =   375
         Left            =   4410
         TabIndex        =   2
         Top             =   225
         Width           =   1425
      End
   End
   Begin VB.Label LblDados 
      AutoSize        =   -1  'True
      Caption         =   "&Pesquisa"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   0
      Top             =   870
      Width           =   645
   End
End
Attribute VB_Name = "FormLoreal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Registros As Recordset  'Representa um conjunto de registros cujo os campos são provenientes de uma ou mais tabelas

Private Sub CmbSugestao_Click()
Dim VarNomeAnt As String
Dim VarLinha As String
Dim VarVend As Recordset
Dim VarPrd As Recordset
Dim VarMov As Recordset
Dim SQL As String

VarNomeAnt = ""

        If MsgBox("Confirma geração arquivo texto produto ? ", vbYesNo, App.Title) = vbYes Then
'             SQL = "Select CodPrd, Descrição_Produto, Complemento , Embalagem, Custo, Margem, Venda, Estoque, PesoBruto FROM Prd Order By Descrição_Produto "
             SQL = "SELECT PRD.CodPrd, PRD.Descrição_Produto, PRD.Embalagem, GRU.Descrição_Grupo, PRD.CodBar, PRD.Estoque FROM GRU INNER JOIN PRD ON GRU.CodGru = PRD.Grupo WHERE (((PRD.Fabricante)=6))"

             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabProduto.csv" For Output As #1
                Print #1, "Codigo;Descrição;Embalagem;EAN;Estoque"
                While Not VarPrd.Eof
                    DoEvents

'                    VarLinha = ";"
                    VarLinha = Trim(VarPrd!CodPrd) & ";"
                    VarLinha = VarLinha + Trim(VarPrd!Descrição_Produto) & ";"
                    
                    If Not IsNull(VarPrd!Embalagem) Then
                        If VarPrd!Embalagem <> "" Then
                            If VarPrd!Embalagem <> "0" Then
                                VarLinha = VarLinha + psCompString(3, Left(VarPrd!Embalagem, 2), 2) & ";"
                            Else
                                VarLinha = VarLinha + "UN" & ";"
                            End If
                        Else
                            VarLinha = VarLinha + "UN" & ";"
                        End If
                    Else
                        VarLinha = VarLinha + "UN" & ";"
                    End If
                    
                    VarLinha = VarLinha + Trim(VarPrd!CodBar) & ";"
                    VarLinha = VarLinha + Format(VarPrd!estoque, "0.00") & ";"
                    Print #1, VarLinha

                    VarPrd.MoveNext
                Wend
'                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
             
             SQL = "SELECT ENT.Tipo, ENT.Nome, ENT.Cidade, ENT.Uf, ENT.CGC, ENT.CPF, PED.Operacao, PED.DataEmiss, PRD.CodBar, PRD.Descrição_Produto, ITP.Qtd, ITP.Valor, PRD.Fabricante, ENT.PesAut, ENT.Ativo, ENT.PesAut FROM (PRD INNER JOIN ITP ON PRD.CodPrd = ITP.Produto) INNER JOIN (ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade) ON ITP.Pedido = PED.CodPed WHERE (((ENT.Tipo)=1) AND ((PED.Operacao)=1 Or (PED.Operacao)=2) AND ((PED.DataEmiss)>=#" & Format(MebDtI, "mm/dd/yyyy") & "# And (PED.DataEmiss)<=#" & Format(MebDtF, "mm/dd/yyyy") & "#) AND ((PRD.Fabricante)=6) AND ((ENT.Ativo)=True)) order by dataemiss "
             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabVendas.csv" For Output As #1
                Print #1, "Tipo;Data;Franqueado;Razão Social;Canal do Cliente;CNPJ Cliente;UF;Cidade;Cliente;EAN;Descrição;Quantidade;Valor Venda"
                While Not VarPrd.Eof
                    DoEvents
                    
                    VarLinha = "V;"
                    VarLinha = VarLinha + Format(VarPrd!DataEmiss, "dd/mm/yyyy") & ";"
                    VarLinha = VarLinha + "16.593.897/0001-44;"
                    VarLinha = VarLinha + "JAZIEL REPEKER PINHEIRO;"
                    If Not IsNull(Trim(VarPrd!PesAut)) Then
                        VarLinha = VarLinha + Trim(VarPrd!PesAut) & ";"
                    Else
                        VarLinha = VarLinha + " " & ";"
                    End If
                    
                    If Not IsNull(VarPrd!CGC) Then
                        VarLinha = VarLinha + Trim(VarPrd!CGC) & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Cpf) & ";"
                    End If
                    VarLinha = VarLinha + Trim(VarPrd!Uf) & ";"
                    VarLinha = VarLinha + Trim(VarPrd!Cidade) & ";"
                    VarLinha = VarLinha + Trim(VarPrd!Nome) & ";"
                    
                    VarLinha = VarLinha + Trim(VarPrd!CodBar) & ";"
                    VarLinha = VarLinha + Trim(VarPrd!Descrição_Produto) & ";"
                    
                    VarLinha = VarLinha + Trim(VarPrd!Qtd) & ";"
                    VarLinha = VarLinha + Trim(VarPrd!Valor) & ";"
                    
                    
                    Print #1, VarLinha
                    VarLinha = " "

                    VarPrd.MoveNext
                Wend
                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
        
        
        
        
        End If


End Sub
