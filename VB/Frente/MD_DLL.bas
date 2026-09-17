Attribute VB_Name = "MD_DLL"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetInputState Lib "user32" () As Long

Public Declare Function Extenso Lib "Extens32.dll" Alias "extenso" (ByVal Valor As String, ByVal Retorno As String) As Integer

Public Vista As Boolean
Global varPesquisa(1 To 2) As Recordset
'Global Mesa(1 To 500) As Integer

Global ResultSat As String
Global VarContMesa As String

Public VarOrcaPrazo As Boolean
Public VarNroCupom As String

Global VarCodDev As Double
Global VarPedDec As Double
Global CfgDesc As String
Public Impressora As Byte
Public LocalBanco As String
Public PortaF As String
Public PortaNF As String
Public Leitor As String
Public ConsultaCh As String
Public NatOper As Byte
Public NroLinha As String
Public Impr40 As String
Public LimiteDesc As Currency
Public EditarFrente As Boolean
Public Receber As Boolean
Public VarDiaBom As Integer
Public VarEmpresa As String
Public VarCGC As String
Public VarInscEst As String
Public VarEnd  As String
Public VarCidade As String
Public VarDtVenc As String
Public VarPcoPrazo  As Boolean
Public VarSenhaBco As String
Public VarBaixaAuto As String
Public VarMostraCusto As String
Public VarSenhaLibera As String
Public VarTelaSenha As String
Public FormSenha As Boolean
Public VarReajConta As Boolean
Public Gelado As Boolean
Public VarCopias As Integer
Public VarMercado As Boolean
Public VarImprEntrega As String
Public VarMediaImpostoNac As Currency
Public VarMediaImpostoEst As Currency
Public VarImprCozinha As String
Public VarImprBar As String
Public VarEnter As Boolean
Public VarComprovante As Boolean
Public VarCaixaSeparado As Boolean
Public VarPacote As Boolean
Public VarVia As Integer
Public VarRapido As Boolean
Public VarDesc5 As Currency
Public VarDesc15 As Currency
Public VarDescPromo As Currency

' Sat

Public VarCnpjDesenvolvedor As String
Public VarChave As String
Public VarChave2 As String
Public VarVersao As String
Public VarCaixaSat As String
Public VarImprSat As String
Public VarSerieNfce As String
Public VarNroNfce As Double

Public VarCnpjEmite As String
Public VarUfEmite As String
Public VarIbgeEmite As String
Public VarRazaoEmite As String
Public VarFantasiaEmite As String
Public VarEndEmite As String
Public VarNroEmite As String
Public VarBairroEmite As String
Public VarCidadeEmite As String
Public VarInscEmite As String
Public VarCepEmite As String
Public VarInscrMuni  As String

Public VarPis As String
Public VarCofins As String

Public VarAliqPis As String
Public VarAliqCofins As String

Public PortaAcbr As String


Global FlagFrame As Boolean
Global VarRegime As Byte
Global VarChaveNfe As String


Public VarMp4000 As Boolean
Public VarImpressoraNf As String
Public VarConveniencia As Boolean
Public VarDescFim As Boolean
Public VarLanchonete As Boolean
Public VarTamBarra As Integer
Public VarLinhaDet As Integer
Public VarComissao As Currency
Public VarOrcamento As Boolean
Public VarTelaPreco As Boolean
Public VarMostraEstoque As Boolean
Public VarCaixa As Boolean
Public VarCnpj As String
Public VarConvenio As String
Public VarObsTotal As Boolean
Public VarAniversariante As Boolean
Public VarColombo As Boolean

Public VarBanco As String
Public VarAgencia  As String
Public VarConta As String

Public VarDocumento As String

Public VarPrdDesc As Currency

Global Formulário As Form

Global Arq 'Impressao em Linhas
Global Fs2 'Impressao em Linhas

Global CTV3 As Object
Global Mensagem As String

Global Nor As String 'Impressao em Modo Normal
Global Cond As String 'Impressao em Modo Condicionado
Global Gde As String 'Impressao em Modo Grande
Global Negi As String 'Impressao em Modo Negrito
Global NegF As String 'Desliga Negrito
Global SaltoPag As String 'Salto da Página
Global Porta As String 'Porta onde se localiza a impressora
Global VarLocalEst As String ' Local de Baixa de Estoque
Global VarMeiaFolha As String ' Impressão em meia folha de formulario
Global VarComplemento As String 'Permitir complementar o nome do produto
Global VarBalanca As String
Global VarBloqueado As Boolean
Global VarCozinha As Boolean

Global TxtImposto As String


Public RetornoBal As Long
Public PesoBal As String * 5

' Declarações da Biblioteca Bematech
'Public Declare Function FechaPortabema Lib "MP2032.DLL" () As Integer
'Public Declare Function IniciaPorta Lib "MP2032.DLL" (ByVal Porta As String) As Integer
'Public Declare Function BematechTX Lib "MP2032.DLL" (ByVal Comando As String) As Integer
'Public Declare Function FormataTX Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer
'Public Declare Function ImprimeBarra Lib "MP2032.DLL" (ByVal Largura As Integer, ByVal Flag As Integer) As Integer
'Public Declare Function ComandoTX Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal Flag As Integer) As Integer
'Public Declare Function SetaDCB Lib "MP2032.DLL" (ByVal Porta As String, ByVal Sets As String, ByVal RTS As Integer, ByVal XON As Integer) As Integer


' Balanças Toledo
    
    Declare Function PegaPeso Lib "P05.DLL" (ByVal OpcaoEscrita As Long, ByVal Peso As String, ByVal Diretorio As String) As Long
    Declare Function AbrePorta Lib "P05.DLL" (ByVal Porta As Long, ByVal BaudRate As Long, ByVal DataBits As Long, ByVal Paridade As Long) As Long
    Declare Function FechaPortaBal Lib "P05.DLL" () As Long
'    Declare Function FechaPortaP05 Lib "P05.DLL" () As Long
    Declare Sub VersaoDLL Lib "P05.DLL" (ByVal Versao As String)
    Declare Function DeterminaUmStopBit Lib "P05.DLL" () As Long



' Declarações da Biblioteca Bematech


Public Declare Function Bematech_FI_VendeItemArredondamentoMFD Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal UnidadeMedida As String, ByVal QtdFracionaria As String, ByVal VlrUnitario As String, ByVal Acrescimo As String, ByVal Desconto As String, Arredonda As Boolean) As Integer


Public Declare Function Bematech_FI_NumeroSerie Lib "BemaFi32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BemaFi32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BemaFi32.dll" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BemaFi32.dll" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BemaFi32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BemaFi32.dll" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BemaFi32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BemaFi32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BemaFi32.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BemaFi32.dll" (ByVal indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BemaFi32.dll" (ByVal indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BemaFi32.dll" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BemaFi32.dll" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BemaFi32.dll" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BemaFi32.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BemaFi32.dll" (ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BemaFi32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BemaFi32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BemaFi32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BemaFi32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BemaFi32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BemaFi32.dll" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BemaFi32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BemaFi32.dll" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BemaFi32.dll" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BemaFi32.dll" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BemaFi32.dll" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BemaFi32.dll" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BemaFi32.dll" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BemaFi32.dll" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BemaFi32.dll" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BemaFi32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BemaFi32.dll" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BemaFi32.dll" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BemaFi32.dll" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BemaFi32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BemaFi32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BemaFi32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BemaFi32.dll" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BemaFi32.dll" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BemaFi32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BemaFi32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BemaFi32.dll" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BemaFi32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BemaFi32.dll" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BemaFi32.dll" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BemaFi32.dll" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BemaFi32.dll" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BemaFi32.dll" (ByVal AliquotasIss As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BemaFi32.dll" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BemaFi32.dll" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BemaFi32.dll" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BemaFi32.dll" (ByRef TipoImpressora As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BemaFi32.dll" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BemaFi32.dll" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BemaFi32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BemaFi32.dll" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BemaFi32.dll" (ByRef Linhas As Integer) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BemaFi32.dll" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BemaFi32.dll" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BemaFi32.dll" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BemaFi32.dll" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BemaFi32.dll" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BemaFi32.dll" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BemaFi32.dll" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BemaFi32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BemaFi32.dll" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BemaFi32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BemaFi32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BemaFi32.dll" (ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BemaFi32.dll" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal Hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Analitico Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Mestre Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ImprimeConfiguracoesImpressora Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AberturaDoDia Lib "BemaFi32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechamentoDoDia Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ValorFormaPagamento Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscal Lib "BemaFi32.dll" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer

Public Declare Function Bematech_FI_AbreRelatorioGerencialMFD Lib "BemaFi32.dll" (ByVal indice As String) As Integer
Public Declare Function Bematech_FI_UsaRelatorioGerencialMFD Lib "BemaFi32.dll" (ByVal texto As String) As Integer


' Bematech não fiscal
    Public Declare Function FechaPorta Lib "MP2032.DLL" () As Integer
    Public Declare Function IniciaPorta Lib "MP2032.DLL" (ByVal Porta As String) As Integer
    Public Declare Function BematechTX Lib "MP2032.DLL" (ByVal Comando As String) As Integer
    Public Declare Function FormataTX Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer
    Public Declare Function ImprimeBarra Lib "MP2032.DLL" (ByVal Largura As Integer, ByVal Flag As Integer) As Integer
    Public Declare Function ComandoTX Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal Flag As Integer) As Integer
    Public Declare Function SetaDCB Lib "MP2032.DLL" (ByVal Porta As String, ByVal Sets As String, ByVal RTS As Integer, ByVal XON As Integer) As Integer
    Public Declare Function ImprimeCodigoQRCODE Lib "MP2032.DLL" (ByVal errorCorrectionLevel As Integer, ByVal moduleSize As Integer, ByVal codeType As Integer, ByVal QRCodeVersion As Integer, ByVal encodingModes As Integer, ByVal codeQr As String) As Integer
    Public Declare Function AcionaGuilhotina Lib "MP2032.DLL" (ByVal Modo As Integer) As Integer
    Public Declare Function ConfiguraCodigoBarras Lib "MP2032.DLL" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer
    Public Declare Function ImprimeCodigoBarrasCODE128 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
    Public Declare Function ConfiguraModeloImpressora Lib "MP2032.DLL" (ByVal ModeloImpressora As Integer) As Integer
    Public Declare Function ImprimeCodigoBarrasCODABAR Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
    Public Declare Function PrinterReset Lib "MP2032.DLL" () As Integer

Public ValorUnit As String

'Declaração do retorno das funções
Public Retorno As Long
Public Flag As Byte

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


Public Function Numero(Valor As String, Decimais As Byte) As String
If Decimais = 2 Then
    Valor = Format(Valor, "0.00")
    Numero = Left(Valor, Len(Valor) - 3) & "." & Right(Valor, 2)
Else
    Valor = Format(Valor, "0.000")
    Numero = Left(Valor, Len(Valor) - 4) & "." & Right(Valor, 3)
End If

End Function

Public Function NameOfPC(MachineName As String) As Long
Dim NameSize As Long
Dim x As Long
MachineName = Space$(16)
NameSize = Len(MachineName)
x = GetComputerName(MachineName, NameSize)
End Function

Function LimpaCampo(Campo As String) As String
Dim x As Integer
Dim CampoLimpo As String

For x = 1 To Len(Campo)
    If Mid$(Campo, x, 1) >= Chr(32) And Mid$(Campo, x, 1) <= Chr(122) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, x, 1)
    End If
Next
LimpaCampo = psCompString(3, CampoLimpo, Len(Campo))
End Function

Function Truncar(Valor As String, Decimais As Integer) As String
Dim x As Integer
Dim Tam As Integer

For x = 1 To Len(Valor)
    If Mid$(Valor, x, 1) = "," Then
        Tam = x
    End If
Next

Truncar = Left(Valor, Tam + Decimais + 1)

End Function

Public Sub Iniciar(Objeto As Boolean)

If Objeto Then
    Set Fs2 = CreateObject("Scripting.FileSystemObject")
    Set Arq = Fs2.CreateTextFile(PortaNF, True)
Else
    If Left(PortaNF, 3) = "COM" Then
        abre = IniciaPorta(PortaNF)
    Else
        If PortaNF <> "TELA" Then
            Open PortaNF For Output As #1
        Else
            If VarImprEntrega <> "" Then
                Open VarImprEntrega For Output As #1
            End If
        End If
    End If
End If
End Sub
Public Function Alinhar(Valor As String, Casas As Currency, Dec As Integer) As String
Dim VlrTmp As String
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
If Casas < Len(Trim(VlrTmp)) Then
    Casas = Len(Trim(VlrTmp))
End If
If Dec <> 0 Then Alinhar = Space(Casas - Len(Trim(VlrTmp))) & Trim(VlrTmp)
If Dec = 0 Then Alinhar = VlrTmp

End Function

Public Function Centralizar(Valor As String, Casas As Currency) As String
Dim VlrTmp As String
Dim Espacos As Currency

VlrTmp = Valor

Espacos = CInt((Casas - Len(Valor)) / 2)

Centralizar = Space(IIf(Espacos < 0, 0, Espacos)) & VlrTmp

End Function

Public Sub ImprimeLinha(PulosLinha As Integer, PulosColuna As Integer, Valor As String, Tamanho As Currency, Condensado As Boolean, Dec As Integer)

If IsNumeric(Valor) Then Valor = Alinhar(Trim(Valor), Tamanho, Dec)

'If Tamanho < Len(Valor) Then
'    Tamanho = Len(Valor)
'End If

If PulosLinha = 0 Then
    If Condensado = True Then
        Arq.Write Cond & Space(PulosColuna) & Mid(Valor, 1, Tamanho) & Space(Tamanho - Len(Valor)) & Nor
    Else
        Arq.Write Space(PulosColuna) & Mid(Valor, 1, Tamanho) & Space(Tamanho - Len(Mid(Valor, 1, Tamanho))) & Nor
    End If
Else
    Arq.WriteBlankLines (PulosLinha)

End If

End Sub

Public Sub Fim(Objeto As Boolean)

If Left(PortaNF, 3) <> "COM" Then
    If Objeto Then
        Arq.Close
    Else
        Close #1
    End If
Else
    fecha = FechaPorta
End If

End Sub

Function AbilitaBotoes(Controle)

Dim Componente As Variant
Dim Botao As Variant

With MDIPrincipal
    
    For Each Botao In .TlbPrincipal.Buttons
            Botao.Enabled = Controle
    Next
               
End With

End Function

Public Sub Imprimir()
Dim x As Integer

MDIPrincipal.RptRel.RetrieveDataFiles

MDIPrincipal.RptRel.PrintReport

For x = 0 To 100
    MDIPrincipal.RptRel.Formulas(x) = ""
Next

End Sub

Function MensErro(Retorno) As Integer
Dim buf_ret As String * 100
Dim retorno2 As Integer
Dim x As Integer

retorno2 = EsperaResposta(buf_ret)

'While retorno2 = -26
'       retorno2 = EsperaResposta(buf_ret)
'Wend

If retorno2 <> 0 Then
    FrmFrente.Mens.Caption = retorno2
    If retorno2 <> 0 Then
'        MsgBox " STATUS " & retorno2
    End If
Else
    FrmFrente.Mens.Caption = ""
End If


Select Case retorno2
    Case 33
        MsgBox "Impressora fiscal com pouco papel", vbExclamation, App.Title
    Case 28
        MsgBox "Redução Z pendente", vbExclamation, App.Title
    Case 47
        MsgBox "Inserir bobina nova", vbExclamation, App.Title
    Case 29
        MsgBox "Travado por redução Z", vbExclamation, App.Title
    Case 18
        MsgBox "Aliquota inválida", vbExclamation, App.Title
    Case 15
        MsgBox "Situação tributária inválida", vbExclamation, App.Title
End Select

Select Case Retorno
    Case CIF_OK
    
    Case CIF_ERR
      MsgBox "Erro no Comando", Buttons:=vbOKOnly + vbInformation
    Case CIF_EMEXECUCAO
      MsgBox "Comando sendo executado", Buttons:=vbOKOnly + vbExclamation
    Case CIF_OVERFLOW
      MsgBox "Tamanho de mensagem enviada é maior que o buffer de recepção", Buttons:=vbOKOnly + vbExclamation
    Case CIF_TIMEOUT
      MsgBox "Timeout na execução do comando", Buttons:=vbOKOnly + vbExclamation
    Case CIF_TAMPA_ABERTA
      MsgBox "Tampa Aberta", Buttons:=vbOKOnly + vbExclamation
    Case CIF_ERRO_MECANICO
      MsgBox "Erro mecânico", Buttons:=vbOKOnly + vbExclamation
    Case CIF_IRRECUPERAVEL
      MsgBox "Erro irrecuperável", Buttons:=vbOKOnly + vbExclamation
    Case CIF_TEMPERATURA
      MsgBox "Temperatura da cabeça de impressão está alta", Buttons:=vbOKOnly + vbExclamation
    Case CIF_POUCO_PAPEL
      MsgBox "Pouco Papel", Buttons:=vbOKOnly + vbExclamation
    Case CIF_CUPOMVENDA
      MsgBox "Em início de cupom de venda", Buttons:=vbOKOnly + vbExclamation
    Case CIF_VENDADEITEM
      MsgBox "Em Venda de Item", Buttons:=vbOKOnly + vbExclamation
    Case CIF_CANCELAITEM
      MsgBox "Em cancelamento de item", Buttons:=vbOKOnly + vbExclamation
    Case CIF_CANCELACUPOM
      MsgBox "Em cancelamento de Cupom", Buttons:=vbOKOnly + vbExclamation
    Case CIF_ENCERRACUPOM
      MsgBox "Em fechamento de cupom", Buttons:=vbOKOnly + vbExclamation
    Case CIF_REDUCAOZ
      MsgBox "Em redução Z", Buttons:=vbOKOnly + vbExclamation
    Case CIF_LEITURAX
      MsgBox "EEm leitura X", Buttons:=vbOKOnly + vbExclamation
    Case CIF_MEMFISCAL
      MsgBox "Em leitura da memória Fiscal", Buttons:=vbOKOnly + vbExclamation
    Case CIF_TOTALIZACAO
      MsgBox "Em totalização", Buttons:=vbOKOnly + vbExclamation
    Case CIF_PAGAMENTO
      MsgBox "Em pagamento", Buttons:=vbOKOnly + vbExclamation
End Select

MensErro = retorno2
End Function


Function CodTrib(Codigo As Integer, Aliq As Integer) As String
Select Case Impressora
    Case 1
        If Codigo = 0 Or Aliq = 0 Then
             CodTrib = "T03"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "T" + psCompString(1, Trim(str(Aliq - 1)), 2)
                Case 2
                    CodTrib = "F00"
                Case 3
                    CodTrib = "I00"
                Case 4
                    CodTrib = "N00"
            End Select
        End If
        
    Case 2
        If Codigo = 0 Or Aliq = 0 Then
           CodTrib = "T"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "T"
                Case 2
                    CodTrib = "F"
                Case 3
                    CodTrib = "I"
                Case 4
                    CodTrib = "N"
            End Select
'           CodTrib = Left(CodTrib, 1)
        End If
    Case 3
        If Codigo = 0 Or Aliq = 0 Then
           CodTrib = "17"
        Else
            Select Case Codigo
                Case 2
                    CodTrib = "18"
                Case 3
                    CodTrib = "17"
                Case 4
                    CodTrib = "19"
                Case Else
                    CodTrib = psCompString(1, CStr(Aliq - 1), 2)
            End Select
        End If
    Case 4
        If Codigo = 0 Or Aliq = 0 Then
           CodTrib = "16"
        Else
            Select Case Codigo
                Case 2
                    CodTrib = "16"
                Case 3
                    CodTrib = "17"
                Case 4
                    CodTrib = "18"
                Case Else
                    CodTrib = psCompString(1, CStr(Aliq - 2), 2)
            End Select
        End If
    Case 5
        If Codigo = 0 Or Aliq = 0 Then
             CodTrib = "FF"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "" + psCompString(1, Trim(str(Aliq - 1)), 2)
                Case 2
                    CodTrib = "FF"
                Case 3
                    CodTrib = "II"
                Case 4
                    CodTrib = "NN"
            End Select
        End If
        
    Case 6
        If Codigo = 0 Or Aliq = 0 Then
             CodTrib = "FF"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "" + psCompString(1, Trim(str(Aliq - 1)), 2)
                Case 2
                    CodTrib = "FF"
                Case 3
                    CodTrib = "II"
                Case 4
                    CodTrib = "NN"
            End Select
        End If
                
    Case 7
        If Codigo = 0 Or FrmFrente.Registros.Fields("Aliquota") = 0 Then
             CodTrib = "FF"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "" + psCompString(1, Trim(str(FrmFrente.Registros.Fields("Aliquota") - 1)), 2)
                Case 2
                    CodTrib = "FF"
                Case 3
                    CodTrib = "II"
                Case 4
                    CodTrib = "NN"
            End Select
        End If
        
     
     Case 8
        If Codigo = 0 And FrmFrente.Registros.Fields("Aliquota") = 0 Then
             CodTrib = "-2"
        Else
            Select Case Codigo
                Case 1
                    CodTrib = "" + psCompString(1, Trim(str(FrmFrente.Registros.Fields("Aliquota") - 1)), 2)
                Case 2
                    CodTrib = "-2"
                Case 3
                    CodTrib = "-3"
                Case 4
                    CodTrib = "-4"
            End Select
        End If
        
End Select
End Function

Function VlrAliquota(Codigo As Integer, Aliq As Integer) As String
If Codigo = 0 Or Aliq = 0 Then
     VlrAliquota = "03"
Else
     VlrAliquota = psCompString(1, Trim(str(Aliq - 1)), 2)
End If
End Function

Function LimpaPrd() As String
Dim Produto As String
Dim PrdSujo As String
Dim x As Integer
Produto = ""
PrdSujo = Left(FrmFrente.DatCompra.Recordset("Descrição_Produto"), 38)
For x = 1 To Len(PrdSujo)
    If Mid$(PrdSujo, x, 1) >= Chr(32) And Mid$(PrdSujo, x, 1) <= Chr(122) Then
            Produto = Produto + Mid$(PrdSujo, x, 1)
    End If
Next
LimpaPrd = psCompString(3, Produto, 38)
End Function

Function LimpaNro(Numero) As String
Dim x As Integer
Dim Resp As String
Resp = ""
For x = 1 To Len(Numero)
    If Mid$(Numero, x, 1) <> "," And Mid$(Numero, x, 1) <> "." Then
        Resp = Resp + Mid$(Numero, x, 1)
    End If
Next
LimpaNro = Resp
End Function


Public Function LerINI(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim Ret As String
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 LerINI = Ret
End Function

Public Sub EscreveINI(Secao As String, Entrada As String, texto As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
  'texto= valor que vem depois do igual
  WritePrivateProfileString Secao, Entrada, texto, Arquivo
End Sub


Public Function PassaExtenso(Valor As String) As String
  
  ' Passa um número para a DLL e
  ' recebe-o de volta por extenso

  On Error GoTo Passa_Err

  Dim Retorno As String, x%
  Retorno = Space$(512)
  x% = Extenso(Valor, Retorno)
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

Public Function SaltoPagina(Contador As Integer)
Dim x As Integer
If Contador < 31 Then
    For x = Contador To 31
        Print #1, " "
    Next
End If
End Function

Public Sub BaixarEstoque(Qtd As Currency, CodPrd As Long, Loja, Baixa As Boolean)
Dim Criterio As String
Dim VarTemp As String

On Error GoTo ErroEstoque

With FrmFrente
    Criterio = "Prd.CodPrd = " & CodPrd
    .Registros.FindFirst Criterio
    If Not .Registros.NoMatch Then
            .Registros.Edit
            If Baixa = True Then
                .Registros.Fields("Estoque") = .Registros.Fields("Estoque") - CDec(Qtd)
                .Registros.Fields("DtVenda") = Date
            Else
                .Registros.Fields("Estoque") = .Registros.Fields("Estoque") + CDec(Qtd)
            End If
            .Registros.Update
 '           .DatCompra.Recordset.Edit
 '           .DatCompra.Recordset("Entregue") = 1
 '           .DatCompra.Recordset.Update
    Else
        Open LocalBanco & "\Estoque.txt" For Append As #1
        Print #1, "Produto - Frente - Baixa; " & CodPrd & "; " & Qtd & "; " & Loja & "; " & Date & "; " & Time & "; " & FrmFrente.Maquina
        Close #1
        
    End If

    Criterio = "Produto = " & CodPrd & " and DescricaoLoja = " & """" & Trim(Loja) & """"
    .Loja.FindFirst Criterio
    If Not .Loja.NoMatch Then
        .Loja.Edit
        If Baixa = True Then
            .Loja.Fields("Estoque") = .Loja.Fields("Estoque") - Qtd
        Else
            .Loja.Fields("Estoque") = .Loja.Fields("Estoque") + Qtd
        End If
        .Loja.Update
     Else
        Open LocalBanco & "\Estoque.txt" For Append As #1
        Print #1, "Loja - Frente - Baixa; " & CodPrd & "; " & Qtd & "; " & Loja & "; " & Date & "; " & Time & "; " & FrmFrente.Maquina
        Close #1
     
     End If
                               
     If .Registros("BaixarComp") = True Then
            BaixarComposicao CodPrd, Qtd, VarLocalEst, Baixa
     End If

End With

Exit Sub

ErroEstoque:

    MsgBox Err.Description & Err.Number, vbCritical, Caminho
    Open LocalBanco & "\Estoque.txt" For Append As #1
    Print #1, Err.Description & Err.Number, vbCritical, Caminho
    Close #1
    
End Sub

Public Sub BaixarComposicao(Produto As Long, Qtd As Currency, LojaDesc As String, Baixa As Boolean)
Dim Criterio As String
Dim Prd  As Recordset
Dim Loja As Recordset
Dim Comp As Recordset
Dim sql As String

Set Comp = FrmFrente.Banco.OpenRecordset("Select * from Comp WHERE Produto = " & Produto & " Order By  Produto")

Comp.Requery
If Comp.RecordCount > 0 Then
    Comp.MoveLast
    Comp.MoveFirst
End If

While Not Comp.EOF
    Set Prd = FrmFrente.Banco.OpenRecordset("Select Estoque, CodPrd, DtVenda FROM Prd WHERE CodPrd = " & Comp!PrdComp)
    Prd.Requery
    If Prd.RecordCount > 0 Then
        Prd.Edit
        If Baixa = True Then
            Prd!estoque = Prd!estoque - (Comp!Qtd * Qtd)
            Prd!DtVenda = Date
        Else
            Prd!estoque = Prd!estoque + (Comp!Qtd * Qtd)
        End If
        Prd.Update
    End If
                
    Set Loja = FrmFrente.Banco.OpenRecordset("SELECT * FROM Loj WHERE Produto =" & Comp!PrdComp & " and DescricaoLoja = '" & LojaDesc & "'")
    Loja.Requery
    If Loja.RecordCount > 0 Then
        Loja.Edit
        If Baixa = True Then
            Loja!estoque = Loja!estoque - (Comp!Qtd * Qtd)
        Else
            Loja!estoque = Loja!estoque + (Comp!Qtd * Qtd)
        End If
        Loja.Update
    End If
    Comp.MoveNext
Wend

End Sub
Function LimpaNumero(Campo As String) As String
Dim x As Integer
Dim NumeroLimpo As String

For x = 1 To Len(Campo)
    If IsNumeric(Mid$(Campo, x, 1)) Then
            NumeroLimpo = NumeroLimpo & Mid$(Campo, x, 1)
    End If
Next
LimpaNumero = NumeroLimpo
End Function

Function LIMPACampo2(Campo As String) As String
Dim x As Integer
Dim CampoLimpo As String

For x = 1 To Len(Campo)
    If Mid$(Campo, x, 1) <> "." And Mid$(Campo, x, 1) <> "/" And Mid$(Campo, x, 1) <> "-" And Mid$(Campo, x, 1) <> Chr(32) Then
            CampoLimpo = CampoLimpo + Mid$(Campo, x, 1)
    End If
Next
LIMPACampo2 = psCompString(3, CampoLimpo, Len(Campo))

End Function


Public Function Numero2(Valor As String) As String
Valor = Format(Valor, "0.00")
Numero2 = Left(Valor, Len(Valor) - 3) & "." & Right(Valor, 2)

End Function


Function AbreAcbr(Comando As String)
' Dim x As Double

' Open PortaAcbr & "\Ent.Txt" For Output As #1
' Print #1, Comando
' Close #1
      
 Dim x As Double

If dir(PortaAcbr & "\Sai.txt") <> "" Then

    Kill PortaAcbr & "\Sai.txt"
End If

Open App.Path & "\Ent.Txt" For Output As #1
Print #1, Comando
Close #1

 FileCopy App.Path & "\Ent.Txt", PortaAcbr & "\Tmp.txt"
 Kill App.Path & "\Ent.Txt"
 Name PortaAcbr & "\Tmp.Txt" As PortaAcbr & "\Ent.txt"
     
      
      
End Function

Function TrataAcbr() As String
Dim VarAcbr As String
Dim VarTemp As String
Dim x As Double
Dim Var
Dim CodigoRet As String
On Error Resume Next

CodigoRet = ""

x = 0

If dir(PortaAcbr & "\Sai.txt") = "" Then

    While FileLen(PortaAcbr & "\Sai.txt") = 0

    Wend

End If



While dir(PortaAcbr & "\Sai.txt") = ""
'    X = X + 1
'    FrmFrente.Text1 = X
'    DoEvents
Wend

x = 0

VarAcbr = ""
CodigoRet = ""
'MsgBox dir(PortaAcbr & "\sai.txt", vbNormal)
If dir(PortaAcbr & "\Sai.txt") = "" Then
    For x = 1 To 700000
'    FrmFrente.Text2 = X
'    DoEvents
    Next
Else
    For x = 1 To 100000
    
    Next
End If

If dir(PortaAcbr & "\Sai.txt") = "" Then
    For x = 1 To 700000
'    FrmFrente.Text2 = X
'    DoEvents
    Next
Else
    For x = 1 To 100000
    
    Next
End If

'Open PortaAcbr & "\Sai.txt" For Input As #6

'Close #6
                                                                                                                     
Open PortaAcbr & "\Sai.txt" For Input As #5

'Do
    Line Input #5, VarTemp
'    VarAcbr = VarTemp
'    If Not EOF(5) Then
'        Input #5, VarTemp

        If Left(VarTemp, 2) = "OK" Then
                   
             CodigoRet = LerINI("ENVIO", "CodigoDeRetorno", PortaAcbr & "\Sai.txt")
             VarImprSat = LerINI("ENVIO", "Arquivo", PortaAcbr & "\Sai.txt")
             If CodigoRet = "" Then
                 CodigoRet = LerINI("CANCELAMENTO", "CodigoDeRetorno", PortaAcbr & "\Sai.txt")
             End If
        End If
        

'        If Left(VarTemp, 2) <> "OK" Or Left(VarTemp, 2) <> "ER" Then
'            VarAcbr = VarAcbr & VarTemp
'        End If
'        Input #5, VarTemp
'        Input #5, VarTemp
'        If Left(VarTemp, 16) = "codigoDeRetorno=" Then
'
'           CodigoRet = Trim(Mid(VarTemp, 17, 20))
'         End If
'        Input #5, VarTemp
'        Input #5, VarTemp
'        If Left(VarTemp, 8) = "Arquivo=" Then
'            VarImprSat = Trim(Mid(VarTemp, 9, 200))
'        End If
'    End If
'Loop Until EOF(5)
Close #5

'If CodigoRet = "" Then
'    TrataAcbr
'Else
'    If VarAcbr <> "," Then
''        MsgBox VarAcbr, vbInformation, App.Title
'    End If
'End If
'MsgBox CodigoRet, vbInformation

    If CodigoRet <> "" Then
        RetornoSat CodigoRet
    Else
        TrataAcbr
    End If




'Kill PortaAcbr & "\Sai.txt"

'For x = 1 To 100000

'Next

End Function


Function RetornoSat(Codigo As String)
Dim Result As String
ResultSat = Codigo
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
    Case 18099:     Result = "Erro desconhecido Informar o administrador."
    Case Else
        TrataAcbr

End Select

If Result <> "" Then
    MsgBox Codigo & "-" & Result
End If

End Function

Public Function LerINI2(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim Ret As String
 Ret = String$(400, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 LerINI2 = Ret
End Function


Public Function Arredondamento(ByVal Numero, ByVal Decimais)
      Arredondamento = Int(Numero * 10 ^ Decimais + 1 / 2) / 10 ^ Decimais
End Function



Public Function SomaPedido(ByVal Pedido)
Dim sql As String
Dim SqlSoma As Recordset

Set SqlSoma = FrmFrente.Banco.OpenRecordset("SELECT ITP.Pedido, Sum(ITP.VlrVdaDia) AS SomaDeVlrVdaDia FROM ITP GROUP BY ITP.Pedido HAVING ITP.Pedido = " & Pedido)
SqlSoma.Requery

SomaPedido = FrmFrente.MebTotal  'SqlSoma!SomadeVlrVdaDia

End Function


Function TrataAcbr2(Operacao As String, Numero As String) As String
Dim VarAcbr As String
Dim VarTemp As String
Dim x As Double
Dim Var
Dim CodigoRet As String
Dim VarErro As String
On Error Resume Next


'DoEvents

x = 0
If dir(PortaAcbr & "\Sai.txt") = "" Then

    While FileLen(PortaAcbr & "\Sai.txt") = 0

    Wend

End If

x = 0

    For x = 1 To 700000


    Next

x = 0
    For x = 1 To 700000
'    FrmFrente.Text2 = X

    Next


x = 0

If dir(PortaAcbr & "\Sai.txt") = "" Then
    For x = 1 To 700000


    Next
Else
    For x = 1 To 700000

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
                        VarChaveNfe = LerINI("Nfe" & Trim(Numero), "chdFe", PortaAcbr & "\Sai.txt")
'                        VarChaveNfe = Mid(varTemp, Len(CodigoRet) - 51, 44)
                   ElseIf VarAcbr = "104" Then
                            If LerINI("Nfe" & Trim(Numero), "cStat", PortaAcbr & "\Sai.txt") = "100" Then
                                VarChaveNfe = LerINI("Nfe" & Trim(Numero), "chdFe", PortaAcbr & "\Sai.txt")
                                VarAcbr = "100"
                            Else
                                MsgBox LerINI("Nfe" & Trim(Numero), "XMotivo", PortaAcbr & "\Sai.txt")
                            End If
                       Else
                            MsgBox LerINI("ENVIO", "XMotivo", PortaAcbr & "\Sai.txt")
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
                        FrmTotal.Label9.Caption = "Vencimento do Certificado: " & Right(VarTemp, 10)
            End Select

        End If
                
Close #5

'If Left(CodigoRet, 8) <> "OK: Lote" Then
'    VarChaveNfe = Mid(varTemp, Len(CodigoRet) - 51, 44)
'Else
    If Left(CodigoRet, 2) = "OK" Then
        TrataAcbr2 = VarAcbr
    Else
    '    If VarAcbr <> "," Then
            MsgBox CodigoRet & " - " & VarErro, vbInformation, App.Title
            TrataAcbr2 = 0
    '    End If
    End If
'End If

' RetornoSat CodigoRet
If dir(PortaAcbr & "\Sai.txt") <> "" Then

    Kill PortaAcbr & "\Sai.txt"
End If

'For X = 1 To 100000

'Next

End Function

