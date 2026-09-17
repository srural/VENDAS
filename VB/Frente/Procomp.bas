Attribute VB_Name = "Procomp"
Option Explicit
'*****************************
' Definir funcoes para ECF32M
'*****************************

Declare Sub CLOSECIF Lib "ECF32M.DLL" Alias "CloseCif" ()
Declare Function OPENCIF Lib "ECF32M.DLL" Alias "OpenCif" () As Long


'**********************
' Funcoes da impressora
'**********************

Declare Function MODOCHEQUEVALIDACAO Lib "ECF32M.DLL" Alias "ModoChequeValidacao" (ByVal Tipo As Byte, ByVal load As Byte) As Long
Declare Function ImprimeCheque Lib "ECF32M.DLL" (ByVal l1 As Byte, ByVal c1 As Byte, ByVal l2 As Byte, ByVal c2 As Byte, ByVal l3 As Byte, ByVal c3 As Byte, ByVal l4 As Byte, ByVal l5 As Byte, ByVal c5 As Byte, ByVal l6 As Byte, ByVal l7 As Byte, ByVal c8 As Byte, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal SetAno As Byte, ByVal Data As String, ByVal Com1 As String, ByVal Com2 As String) As Long
Declare Function IMPRIMEVALIDACAO Lib "ECF32M.DLL" Alias "ImprimeValidacao" (ByVal leg As String, ByVal LinhaOp As String) As Long
Declare Function CANCELACHEQUEVALIDACAO Lib "ECF32M.DLL" Alias "CancelaChequeValidacao" () As Long

'*****************************
' Funcoes de operacoes fiscais
'*****************************

Declare Function TOTALIZARCUPOMPARCIAL Lib "ECF32M.DLL" Alias "TotalizarCupomParcial" () As Long
Declare Function TOTALIZARCUPOM Lib "ECF32M.DLL" Alias "TotalizarCupom" (ByVal oper As Byte, ByVal toper As Byte, ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function TotCupomSemDescAcres Lib "ECF32M.DLL" () As Long
Declare Function TotCupomAcresValor Lib "ECF32M.DLL" (ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function TotCupomAcresPorcentagem Lib "ECF32M.DLL" (ByVal porcentagem As String, ByVal legendaOp As String) As Long
Declare Function TotCupomDescValor Lib "ECF32M.DLL" (ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function TotCupomDescPorcentagem Lib "ECF32M.DLL" (ByVal porcentagem As String, ByVal legendaOp As String) As Long
Declare Function Pagamento Lib "ECF32M.DLL" (ByVal reg As String, ByVal vpgto As String, ByVal subtr As Byte) As Long
Declare Function ABRECUPOMFISCAL Lib "ECF32M.DLL" Alias "AbreCupomFiscal" () As Long
Declare Function AbreCupomFiscalCPF_CNPJ Lib "ECF32M.DLL" (ByVal CPF_CNPJ As String) As Long
Declare Function VendaItemStr Lib "ECF32M.DLL" (ByVal fmt As String, ByVal Qtd As String, ByVal punit As String, ByVal trib As String, ByVal tdesc As String, ByVal Valor As String, ByVal unid As String, ByVal cod As String, ByVal ex As String, ByVal descr As String, ByVal legendaOp As String) As Long
Declare Function CANCELAMENTOITEM Lib "ECF32M.DLL" Alias "CancelamentoItem" (ByVal numitem As String) As Long
Declare Function DescontoItem Lib "ECF32M.DLL" (ByVal toper As Byte, ByVal Valor As String, ByVal legop As String) As Long
Declare Function DescontoItemPorcentagem Lib "ECF32M.DLL" (ByVal porcentagem As String, ByVal legop As String) As Long
Declare Function DescontoItemValor Lib "ECF32M.DLL" (ByVal Valor As String, ByVal legop As String) As Long
Declare Function FECHACUPOMFISCAL Lib "ECF32M.DLL" Alias "FechaCupomFiscal" (ByVal tam_msg As String, ByVal msg As String) As Long
Declare Function CANCELACUPOMFISCAL Lib "ECF32M.DLL" Alias "CancelaCupomFiscal" () As Long
Declare Function LEITURAX Lib "ECF32M.DLL" Alias "LeituraX" (ByVal relGer As Byte) As Long
Declare Function REDUCAOZ Lib "ECF32M.DLL" Alias "ReducaoZ" (ByVal relGer As Byte) As Long
Declare Function LeituraXComRelGer Lib "ECF32M.DLL" () As Long
Declare Function ReducaoZComRelGer Lib "ECF32M.DLL" () As Long
Declare Function LEMEMFISCALDATA Lib "ECF32M.DLL" Alias "LeMemFiscalData" (ByVal datai As String, ByVal dataf As String, ByVal res As Byte) As Long
Declare Function LEMEMFISCALREDUCAO Lib "ECF32M.DLL" Alias "LeMemFiscalReducao" (ByVal redi As String, ByVal redf As String, ByVal res As Byte) As Long

'******************************************
' Funcoes de operacoes nao sujeitas ao ICMS
'******************************************

Declare Function ABRECUPOMVINCULADO Lib "ECF32M.DLL" Alias "AbreCupomVinculado" () As Long
Declare Function ABRECUPOMNAOVINCULADO Lib "ECF32M.DLL" Alias "AbreCupomNaoVinculado" () As Long
Declare Function EncerraCupomNaoFiscal Lib "ECF32M.DLL" () As Long
Declare Function CANCELACUPOMNAOFISCAL Lib "ECF32M.DLL" Alias "CancelaCupomNaoFiscal" () As Long
Declare Function OPERREGNAOVINCULADO Lib "ECF32M.DLL" Alias "OperRegNaoVinculado" (ByVal reg As String, ByVal Valor As String, ByVal oper As Byte, ByVal toper As Byte, ByVal valorop As String, ByVal legop As String) As Long
Declare Function AbrirCupom Lib "ECF32M.DLL" (ByVal reg As String, ByVal Valor As String, ByVal oper As String, ByVal toper As String, ByVal valorop As String, ByVal legop As String, ByVal buffRet As String) As Long
Declare Function IMPRIMELINHANAOFISCAL Lib "ECF32M.DLL" Alias "ImprimeLinhaNaoFiscal" (ByVal Par As Byte, ByVal str As String) As Long
Declare Function ImprimeLinhaNaoFiscalTexto Lib "ECF32M.DLL" (ByVal Par As String, ByVal str As String) As Long
Declare Function PROGRAMALEGENDA Lib "ECF32M.DLL" Alias "ProgramaLegenda" (ByVal reg As String, ByVal leg As String) As Long

'*****************
' Funcoes diversas
'*****************

Declare Function AcionarGaveta Lib "ECF32M.DLL" () As Long
Declare Function ProgramaHorarioVeraoStr Lib "ECF32M.DLL" (ByVal hv As String) As Long
Declare Function IMPRIMETOTALIZADORES Lib "ECF32M.DLL" Alias "ImprimeTotalizadores" (ByVal reg As String) As Long
Declare Function TransTabAliquotas Lib "ECF32M.DLL" () As Long
Declare Function TRANSTOTCONT Lib "ECF32M.DLL" Alias "TransTotCont" () As Long
Declare Function TRANSSTATUS Lib "ECF32M.DLL" Alias "TransStatus" (ByVal bittest As Long, ByVal BufStat As String) As Long
Declare Function TransDataHora Lib "ECF32M.DLL" () As Long
Declare Function EcfPar Lib "ECF32M.DLL" (ByVal Par As String) As Long
Declare Function ProgLinhaAdicional Lib "ECF32M.DLL" (ByVal reg As String) As Long
Declare Function AjusteHora Lib "ECF32M.DLL" (ByVal dir As Byte, ByVal hora As String) As Long
Declare Function EcfID Lib "ECF32M.DLL" () As Long
Declare Function EsperaResposta Lib "ECF32M.DLL" (ByVal buf_ret As String) As Long
    'Declaração das Constantes de erro
    Public Const CIF_OK = 0
    Public Const CIF_ERR = -1
    Public Const CIF_EMEXECUCAO = -3
    Public Const CIF_TIMEOUT = -4
    Public Const CIF_OVERFLOW = -5
    Public Const CIF_ERR_CONFIG = -7
    Public Const CIF_ERR_SERIAL = -8
    Public Const CIF_TAMPA_ABERTA = -11
    Public Const CIF_ERRO_MECANICO = -12
    Public Const CIF_IRRECUPERAVEL = -13
    Public Const CIF_TEMPERATURA = -14
    Public Const CIF_POUCO_PAPEL = -15
    Public Const CIF_CUPOMVENDA = -16
    Public Const CIF_VENDADEITEM = -17
    Public Const CIF_CANCELAITEM = -18
    Public Const CIF_CANCELACUPOM = -19
    Public Const CIF_ENCERRACUPOM = -20
    Public Const CIF_REDUCAOZ = -21
    Public Const CIF_LEITURAX = -22
    Public Const CIF_MEMFISCAL = -23
    Public Const CIF_TOTALIZACAO = -24
    Public Const CIF_PAGAMENTO = -25
    Public Const CIF_SEMRETORNO = -26





