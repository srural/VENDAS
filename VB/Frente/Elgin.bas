Attribute VB_Name = "Elgin"
'===============================================================================
'********************************************************************************
'
'                   DECLARAÇÃO DAS VARIÁVEIS GLOBAIS DO DemoFit32
'
'********************************************************************************
'===============================================================================}
'Variável global que aponta para um objeto da classe CElgin
'Global Elgin32 As CElgin32

'===============================================================================
'********************************************************************************
'
'                      DECLARAÇÃO DAS FUNÇÕES DA Elgin.DLL
'
'********************************************************************************
'===============================================================================}
' Específicas Elgin MFD
Declare Function Elgin_RetornoImpressora Lib "Elgin.dll" (ByRef I As Integer, ByVal ErrorMsg As String) As Integer
Declare Function Elgin_VendaBruta Lib "Elgin.dll" (ByVal VendaBruta As String) As Integer
Declare Function Elgin_VendaLiquida Lib "Elgin.dll" (ByVal VendaLiquida As String) As Integer
Declare Function Elgin_TotalDocTroco Lib "Elgin.dll" (ByVal TotalDocTroco As String) As Integer
Declare Function Elgin_TotalDiaTroco Lib "Elgin.dll" (ByVal TotalDiaTroco As String) As Integer


' Funções de Inicialização
Declare Function Elgin_AbrePortaSerial Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaPortaSerial Lib "Elgin.dll" () As Integer

Declare Function Elgin_AlteraSimboloMoeda Lib "Elgin.dll" (ByVal SimboloMoeda As String) As Integer
Declare Function Elgin_ProgramaAliquota Lib "Elgin.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Declare Function Elgin_ProgramaHorarioVerao Lib "Elgin.dll" () As Integer
Declare Function Elgin_NomeiaDepartamento Lib "Elgin.dll" (ByVal indice As Integer, ByVal Departamento As String) As Integer
Declare Function Elgin_NomeiaTotalizadorNaoSujeitoIcms Lib "Elgin.dll" (ByVal indice As Integer, ByVal Totalizador As String) As Integer
Declare Function Elgin_ProgramaTruncamento Lib "Elgin.dll" () As Integer
Declare Function Elgin_ProgramaArredondamento Lib "Elgin.dll" () As Integer
Declare Function Elgin_LinhasEntreCupons Lib "Elgin.dll" (ByVal Linhas As Integer) As Integer
Declare Function Elgin_EspacoEntreLinhas Lib "Elgin.dll" (ByVal Dots As Integer) As Integer
Declare Function Elgin_ForcaImpactoAgulhas Lib "Elgin.dll" (ByVal ForcaImpacto As Integer) As Integer

' Funções do Cupom Fiscal

Declare Function Elgin_AbreCupom Lib "Elgin.dll" (ByVal CGC_CPF As String) As Integer
Declare Function Elgin_VendeItem Lib "Elgin.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
'Declare Function  Elgin_VendeItemDepartamento  Lib "Elgin.dll" ( ByVal  Codigo As String, Descricao As String, Aliquota As String, ValorUnitario As String, Quantidade As String, Acrescimo As String, Desconto As String, IndiceDepartamento As String, UnidadeMedida As String) As Integer
Declare Function Elgin_CancelaItemAnterior Lib "Elgin.dll" () As Integer
Declare Function Elgin_CancelaItemGenerico Lib "Elgin.dll" (ByVal NumeroItem As String) As Integer
Declare Function Elgin_CancelaCupom Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaCupomResumido Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Declare Function Elgin_FechaCupom Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ValorPago As String, Mensagem As String) As Integer
Declare Function Elgin_ResetaImpressora Lib "Elgin.dll" () As Integer
Declare Function Elgin_IniciaFechamentoCupom Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Declare Function Elgin_EfetuaFormaPagamento Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Declare Function Elgin_EfetuaFormaPagamentoDescricaoForma Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoFormaPagto As String) As Integer
Declare Function Elgin_TerminaFechamentoCupom Lib "Elgin.dll" (ByVal Mensagem As String) As Integer
Declare Function Elgin_EstornoFormasPagamento Lib "Elgin.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Declare Function Elgin_UsaUnidadeMedida Lib "Elgin.dll" (ByVal UnidadeMedida As String) As Integer
Declare Function Elgin_AumentaDescricaoItem Lib "Elgin.dll" (ByVal Descricao As String) As Integer
'Declare Function  Elgin_EfetuaFormaPagamentoImpAntiga   Lib "Elgin.dll" (ByVal FormaPagamento As String,ByVal ValorFormaPagamento As String ) As Integer

'Funções dos Relatórios Fiscais



Declare Function Elgin_LeituraX Lib "Elgin.dll" () As Integer
Declare Function Elgin_ReducaoZ Lib "Elgin.dll" (ByVal Data As String, ByVal hora As String) As Integer
Declare Function Elgin_RelatorioGerencial Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_RelatorioGerencialTEF Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_FechaRelatorioGerencial Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaRelatorioXouZ Lib "Elgin.dll" () As Integer
Declare Function Elgin_LeituraMemoriaFiscalData Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalReducao Lib "Elgin.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalSerialData Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalSerialReducao Lib "Elgin.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer

' Funções das Operações Não Fiscais

Declare Function Elgin_RecebimentoNaoFiscal Lib "Elgin.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Declare Function Elgin_AbreComprovanteNaoFiscalVinculado Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Declare Function Elgin_UsaComprovanteNaoFiscalVinculado Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_UsaComprovanteNaoFiscalVinculadoTEF Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_FechaComprovanteNaoFiscalVinculado Lib "Elgin.dll" () As Integer
Declare Function Elgin_Sangria Lib "Elgin.dll" (ByVal Valor As String) As Integer
Declare Function Elgin_Suprimento Lib "Elgin.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer

' Funções de Informações da Impressora

Declare Function Elgin_NumeroSerie Lib "Elgin.dll" (ByVal NumeroSerie As String) As Integer
Declare Function Elgin_SubTotal Lib "Elgin.dll" (ByVal SubTotal As String) As Integer
Declare Function Elgin_NumeroCupom Lib "Elgin.dll" (ByVal NumeroCupom As String) As Integer
Declare Function Elgin_LeituraXSerial Lib "Elgin.dll" () As Integer
Declare Function Elgin_VersaoFirmware Lib "Elgin.dll" (ByVal VersaoFirmware As String) As Integer
Declare Function Elgin_CGC_IE Lib "Elgin.dll" (ByVal CGC As String, ByVal ie As String) As Integer
Declare Function Elgin_CNPJ_IE Lib "Elgin.dll" (ByVal CNPJ As String, ByVal ie As String) As Integer
Declare Function Elgin_GrandeTotal Lib "Elgin.dll" (ByVal GrandeTotal As String) As Integer
Declare Function Elgin_Cancelamentos Lib "Elgin.dll" (ByVal ValorCancelamentos As String) As Integer
Declare Function Elgin_Descontos Lib "Elgin.dll" (ByVal ValorDescontos As String) As Integer
Declare Function Elgin_NumeroOperacoesNaoFiscais Lib "Elgin.dll" (ByVal NumeroOperacoes As String) As Integer
Declare Function Elgin_NumeroCuponsCancelados Lib "Elgin.dll" (ByVal NumeroCancelamentos As String) As Integer
Declare Function Elgin_NumeroIntervencoes Lib "Elgin.dll" (ByVal NumeroIntervencoes As String) As Integer
Declare Function Elgin_NumeroReducoes Lib "Elgin.dll" (ByVal NumeroReducoes As String) As Integer
Declare Function Elgin_NumeroSubstituicoesProprietario Lib "Elgin.dll" (ByVal NumeroSubstituicoes As String) As Integer
Declare Function Elgin_UltimoItemVendido Lib "Elgin.dll" (ByVal NumeroItem As String) As Integer
Declare Function Elgin_ClicheProprietario Lib "Elgin.dll" (ByVal Cliche As String) As Integer
Declare Function Elgin_NumeroCaixa Lib "Elgin.dll" (ByVal NumeroCaixa As String) As Integer
Declare Function Elgin_NumeroLoja Lib "Elgin.dll" (ByVal NumeroLoja As String) As Integer
Declare Function Elgin_SimboloMoeda Lib "Elgin.dll" (ByVal SimboloMoeda As String) As Integer
Declare Function Elgin_MinutosLigada Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_MinutosImprimindo Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_VerificaModoOperacao Lib "Elgin.dll" (ByVal modo As String) As Integer
Declare Function Elgin_VerificaEpromConectada Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_FlagsFiscais Lib "Elgin.dll" (ByRef Flag As Integer) As Integer
Declare Function Elgin_FlagsFiscaisStr Lib "Elgin.dll" (ByVal FlagFiscal As String) As Integer
Declare Function Elgin_ValorPagoUltimoCupom Lib "Elgin.dll" (ByVal ValorCupom As String) As Integer
Declare Function Elgin_DataHoraImpressora Lib "Elgin.dll" (ByVal Data As String, hora As String) As Integer
Declare Function Elgin_ContadoresTotalizadoresNaoFiscais Lib "Elgin.dll" (ByVal Contadores As String) As Integer
Declare Function Elgin_VerificaTotalizadoresNaoFiscais Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_DataHoraReducao Lib "Elgin.dll" (ByVal Data As String, ByVal hora As String) As Integer
Declare Function Elgin_DataMovimento Lib "Elgin.dll" (ByVal Data As String) As Integer
Declare Function Elgin_VerificaTruncamento Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_Acrescimos Lib "Elgin.dll" (ByVal ValorAcrescimos As String) As Integer
Declare Function Elgin_ContadorBilhetePassagem Lib "Elgin.dll" (ByVal ContadorPassagem As String) As Integer
Declare Function Elgin_VerificaAliquotasIss Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaFormasPagamento Lib "Elgin.dll" (ByVal Formas As String) As Integer
Declare Function Elgin_VerificaRecebimentoNaoFiscal Lib "Elgin.dll" (ByVal Recebimentos As String) As Integer
Declare Function Elgin_VerificaDepartamentos Lib "Elgin.dll" (ByVal Departamentos As String) As Integer
Declare Function Elgin_VerificaTipoImpressora Lib "Elgin.dll" (ByRef TipoImpressora As Integer) As Integer
Declare Function Elgin_VerificaTipoImpressoraStr Lib "Elgin.dll" (ByRef TipoImpressora As String) As Integer
Declare Function Elgin_VerificaTotalizadoresParciais Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_RetornoAliquotas Lib "Elgin.dll" (ByVal Aliquotas As String) As Integer
Declare Function Elgin_VerificaEstadoImpressora Lib "Elgin.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Declare Function Elgin_VerificaEstadoImpressoraStr Lib "Elgin.dll" (ByRef ACK As String, ByRef ST1 As String, ByRef ST2 As String) As Integer
Declare Function Elgin_DadosUltimaReducao Lib "Elgin.dll" (ByVal DadosReducao As String) As Integer
Declare Function Elgin_MonitoramentoPapel Lib "Elgin.dll" (ByRef Linhas As Integer) As Integer
Declare Function Elgin_VerificaIndiceAliquotasIss Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_ValorFormaPagamento Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String) As Integer
Declare Function Elgin_ValorTotalizadorNaoFiscal Lib "Elgin.dll" (ByVal Totalizador As String, ByVal Valor As String) As Integer
Declare Function Elgin_ImprimeClicheMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_VerificaSensorPoucoPapelMFD Lib "Elgin.dll" (ByVal Flag As String) As Integer

'Funções de Autenticação e Gaveta de Dinheiro

Declare Function Elgin_Autenticacao Lib "Elgin.dll" () As Integer
Declare Function Elgin_ProgramaCaracterAutenticacao Lib "Elgin.dll" (ByVal Parametros As String) As Integer
Declare Function Elgin_AcionaGaveta Lib "Elgin.dll" () As Integer
Declare Function Elgin_VerificaEstadoGaveta Lib "Elgin.dll" (ByRef EstadoGaveta As Integer) As Integer
Declare Function Elgin_VerificaEstadoGavetaStr Lib "Elgin.dll" (ByVal EstadoGaveta As String) As Integer

'Funções para a Impressora Restaurante

'Declare Function ElginR_AbreCupomRestaurante Lib "Elgin.dll" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
'Declare Function ElginR_RegistraVenda Lib "Elgin.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
'Declare Function ElginR_CancelaVenda Lib "Elgin.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
'Declare Function ElginR_ConferenciaMesa Lib "Elgin.dll" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
'Declare Function ElginR_AbreConferenciaMesa Lib "Elgin.dll" (ByVal Mesa As String) As Integer
'Declare Function ElginR_FechaConferenciaMesa Lib "Elgin.dll" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
'Declare Function ElginR_TransferenciaMesa Lib "Elgin.dll" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
'Declare Function ElginR_ContaDividida Lib "Elgin.dll" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
'Declare Function ElginR_FechaCupomContaDividida Lib "Elgin.dll" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
'Declare Function ElginR_TransferenciaItem Lib "Elgin.dll" (ByVal MesaOrigem As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
'Declare Function ElginR_RelatorioMesasAbertas Lib "Elgin.dll" (ByVal TipoRelatorio As Integer) As Integer
'Declare Function  ElginR_ImprimeCardapio As Integer
'Declare Function  ElginR_RelatorioMesasAbertasSerial As Integer
'Declare Function  ElginR_CardapioPelaSerial As Integer
'Declare Function ElginR_RegistroVendaSerial Lib "Elgin.dll" (ByVal Mesa As String) As Integer
'Declare Function ElginR_VerificaMemoriaLivre Lib "Elgin.dll" (ByVal Bytes As String) As Integer
'Declare Function ElginR_FechaCupomRestaurante Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal ValorFormaPagto As String, ByVal Mensagem As String) As Integer
'Declare Function ElginR_FechaCupomResumidoRestaurante Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
'
' Função para a Impressora Bilhete de Passagem

Declare Function Elgin_AbreBilhetePassagem Lib "Elgin.dll" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal Embarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer

'Funções de Impressão de Cheques

Declare Function Elgin_ProgramaMoedaSingular Lib "Elgin.dll" (ByVal MoedaSingular As String) As Integer
Declare Function Elgin_ProgramaMoedaPlural Lib "Elgin.dll" (ByVal MoedaPlural As String) As Integer
Declare Function Elgin_CancelaImpressaoCheque Lib "Elgin.dll" () As Integer
Declare Function Elgin_VerificaStatusCheque Lib "Elgin.dll" (ByRef StatusCheque As Integer) As Integer
Declare Function Elgin_ImprimeCheque Lib "Elgin.dll" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Declare Function Elgin_IncluiCidadeFavorecido Lib "Elgin.dll" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Declare Function Elgin_ImprimeCopiaCheque Lib "Elgin.dll" () As Integer

'Outras Funções

'Declare Function  Elgin_AbrePortaSerial As Integer
'Declare Function  Elgin_FechaPortaSerial As Integer,   Name 'Elgin_FechaPortaSerial',
Declare Function Elgin_MapaResumo Lib "Elgin.dll" () As Integer
Declare Function Elgin_AberturaDoDia Lib "Elgin.dll" (ByVal ValorCompra As String, ByVal FormaPagamento As String) As Integer
Declare Function Elgin_FechamentoDoDia Lib "Elgin.dll" () As Integer
Declare Function Elgin_ImprimeConfiguracoesImpressora Lib "Elgin.dll" () As Integer
Declare Function Elgin_ImprimeDepartamentos Lib "Elgin.dll" () As Integer
Declare Function Elgin_RelatorioTipo60Analitico Lib "Elgin.dll" () As Integer
Declare Function Elgin_RelatorioTipo60Mestre Lib "Elgin.dll" () As Integer
Declare Function Elgin_VerificaImpressoraLigada Lib "Elgin.dll" () As Integer
Declare Function Elgin_DadosSintegra Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer

'Funções da Impressora Fiscal MFD

Declare Function Elgin_AbreCupomMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_CancelaCupomMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_ProgramaFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagto As String, ByVal OperacaoTef As String) As Integer
Declare Function Elgin_EfetuaFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal Parcelas As String, ByVal DescricaoFormaPagto As String) As Integer
Declare Function Elgin_CupomAdicionalMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_AcrescimoDescontoItemMFD Lib "Elgin.dll" (ByVal Item As String, ByVal AcrescimoDesconto As String, TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Declare Function Elgin_NomeiaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal indice As String, ByVal Descricao As String) As Integer
Declare Function Elgin_AutenticacaoMFD Lib "Elgin.dll" (ByVal Linhas As String, ByVal Texto As String) As Integer
Declare Function Elgin_AbreComprovanteNaoFiscalVinculadoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String, ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_ReimpressaoNaoFiscalVinculadoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_AbreRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_EfetuaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal IndiceTotalizador As String, ByVal ValorRecebimento As String) As Integer
Declare Function Elgin_IniciaFechamentoCupomMFD Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Declare Function Elgin_IniciaFechamentoRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Declare Function Elgin_FechaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal Mensagem As String) As Integer
Declare Function Elgin_CancelaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_AbreRelatorioGerencialMFD Lib "Elgin.dll" (ByVal indice As String) As Integer
Declare Function Elgin_UsaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_UsaRelatorioGerencialMFDTEF Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_SegundaViaNaoFiscalVinculadoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_EstornoNaoFiscalVinculadoMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_NumeroSerieMFD Lib "Elgin.dll" (ByVal NumeroSerie As String) As Integer
Declare Function Elgin_VersaoFirmwareMFD Lib "Elgin.dll" (ByVal VersaoFirmware As String) As Integer
Declare Function Elgin_CNPJMFD Lib "Elgin.dll" (ByVal CNPJ As String) As Integer
Declare Function Elgin_InscricaoEstadualMFD Lib "Elgin.dll" (ByVal InscricaoEstadual As String) As Integer
Declare Function Elgin_InscricaoMunicipalMFD Lib "Elgin.dll" (ByVal InscricaoMunicipal As String) As Integer
Declare Function Elgin_TempoOperacionalMFD Lib "Elgin.dll" (ByVal TempoOperacional As String) As Integer
Declare Function Elgin_MinutosEmitindoDocumentosFiscaisMFD Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_ContadoresTotalizadoresNaoFiscaisMFD Lib "Elgin.dll" (ByVal Contadores As String) As Integer
Declare Function Elgin_VerificaTotalizadoresNaoFiscaisMFD Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_VerificaFormasPagamentoMFD Lib "Elgin.dll" (ByVal FormasPagamento As String) As Integer
Declare Function Elgin_VerificaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal Recebimentos As String) As Integer
Declare Function Elgin_VerificaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Relatorios As String) As Integer
Declare Function Elgin_ContadorComprovantesCreditoMFD Lib "Elgin.dll" (ByVal Comprovantes As String) As Integer
Declare Function Elgin_ContadorOperacoesNaoFiscaisCanceladasMFD Lib "Elgin.dll" (ByVal OperacoesCanceladas As String) As Integer
Declare Function Elgin_ContadorRelatoriosGerenciaisMFD Lib "Elgin.dll" (ByVal Relatorios As String) As Integer
Declare Function Elgin_ContadorCupomFiscalMFD Lib "Elgin.dll" (ByVal CuponsEmitidos As String) As Integer
Declare Function Elgin_ContadorFitaDetalheMFD Lib "Elgin.dll" (ByVal ContadorFita As String) As Integer
Declare Function Elgin_ComprovantesNaoFiscaisNaoEmitidosMFD Lib "Elgin.dll" (ByVal Comprovantes As String) As Integer
Declare Function Elgin_NumeroSerieMemoriaMFD Lib "Elgin.dll" (ByVal NumeroSerieMFD As String) As Integer
Declare Function Elgin_MarcaModeloTipoImpressoraMFD Lib "Elgin.dll" (ByVal Marca As String, ByVal modelo As String, ByVal Tipo As String) As Integer
Declare Function Elgin_ReducoesRestantesMFD Lib "Elgin.dll" (ByVal Reducoes As String) As Integer
Declare Function Elgin_VerificaTotalizadoresParciaisMFD Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_DadosUltimaReducaoMFD Lib "Elgin.dll" (ByVal DadosReducao As String) As Integer
Declare Function Elgin_LeituraChequeMFD Lib "Elgin.dll" (ByVal CodigoCMC7 As String) As Integer
Declare Function Elgin_ImprimeChequeMFD Lib "Elgin.dll" (ByVal NumeroBanco, ByVal Valor, ByVal Favorecido, ByVal Cidade, ByVal Data, ByVal Mensagem, ByVal ImpressaoVerso, ByVal Linhas As String) As Integer
Declare Function Elgin_HabilitaDesabilitaRetornoEstendidoMFD Lib "Elgin.dll" (ByVal FlagRetorno As String) As Integer
Declare Function Elgin_RetornoImpressoraMFD Lib "Elgin.dll" (ByRef ACK As Integer, ByVal ST1 As Integer, ByVal ST2 As Integer, ByVal ST3 As Integer) As Integer
'Declare Function  Elgin_AbreBilhetePassagemMFD Lib "Elgin.dll" ( ByVal Embarque As String , ByVal Destino As String , ByVal Linha As String , ByVal Agencia As String , ByVal Data As String , ByVal Hora As String , ByVal Poltrona As String , ByVal Plataforma As String , ByVal TipoPassagem As String ) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoItemMFD Lib "Elgin.dll" (ByVal cFlag, ByVal cItem As String) As Integer
Declare Function Elgin_SubTotalizaCupomMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_SubTotalizaRecebimentoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_TotalLivreMFD Lib "Elgin.dll" (ByVal cMemoriaLivre As String) As Integer
Declare Function Elgin_TamanhoTotalMFD Lib "Elgin.dll" (ByVal cTamanhoMFD As String) As Integer
Declare Function Elgin_AcrescimoDescontoSubtotalRecebimentoMFD Lib "Elgin.dll" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Declare Function Elgin_AcrescimoDescontoSubtotalMFD Lib "Elgin.dll" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoSubtotalMFD Lib "Elgin.dll" (ByVal cFlag As String) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoSubtotalRecebimentoMFD Lib "Elgin.dll" (ByVal cFlag As String) As Integer
Declare Function Elgin_TotalizaCupomMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_TotalizaRecebimentoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_PercentualLivreMFD Lib "Elgin.dll" (ByVal cMemoriaLivre As String) As Integer
Declare Function Elgin_DataHoraUltimoDocumentoMFD Lib "Elgin.dll" (ByVal cDataHora As String) As Integer
Declare Function Elgin_MapaResumoMFD Lib "Elgin.dll" () As Integer

Declare Function Elgin_RelatorioTipo60AnaliticoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_ValorFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String) As Integer
Declare Function Elgin_ValorTotalizadorNaoFiscalMFD Lib "Elgin.dll" (ByVal Totalizador As String, ByVal Valor As String) As Integer
Declare Function Elgin_VerificaEstadoImpressoraMFD Lib "Elgin.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer, ByRef ST3 As Integer) As Integer
'{
'Declare Function  Elgin_RelatorioSintegraMFD   Lib "Elgin.dll" ( ByVal  iRelatorios  As Integer,
'                                           cArquivo     As String,
'                                           cMes         As String,
'                                           cAno         As String,
'                                           cRazaoSocial As String,
'                                           cEndereco    As String,
'                                           cNumero      As String,
'                                           cComplemento As String,
'                                           cBairro      As String,
'                                           cCidade      As String,
'                                           cCEP         As String,
'                                           cTelefone    As String,
'                                           cFax         As String,
'                                           cContato     As String ) As Integer,   Name 'Elgin_RelatorioSintegraMFD',
'}
Declare Function Elgin_DownloadMF Lib "Elgin.dll" (ByVal Arquivo As String) As Integer
Declare Function Elgin_DownloadMFD Lib "Elgin.dll" (ByVal Arquivo As String, TipoDownload As String, ByVal ParametroInicial As String, ByVal ParametroFinal As String, ByVal UsuarioECF As String) As Integer
Declare Function Elgin_RegistrosTipo60 Lib "Elgin.dll" () As Integer

'{Declare Function  Elgin_FormatoDadosMFD   Lib "Elgin.dll" ( ByVal  ArquivoOrigem    As String,
'                                      ArquivoDestino   As String,
'                                      TipoFormato      As String,
'                                      TipoDownload     As String,
'                                      ParametroInicial As String,
'                                      ParametroFinal   As String,
'                                      UsuarioECF       As String ) As Integer,   Name 'Elgin_FormatoDadosMFD',
'}
'// Funções disponíveis somente na impressora fiscal MP-2000 TH FI versão 01.00.02
'//Declare Function  Elgin_AtivaDesativaVendaUmaLinhaMFD   Lib "Elgin.dll" ( ByVal  iFlag As Integer ) As Integer
'//Declare Function  Elgin_AtivaDesativaAlinhamentoEsquerdaMFD   Lib "Elgin.dll" ( ByVal  iFlag As Integer ) As Integer
'//Declare Function  Elgin_AtivaDesativaCorteProximoMFD Lib "Elgin.dll" ( ByVal ) As Integer
'//Declare Function  Elgin_AtivaDesativaTratamentoONOFFLineMFD   Lib "Elgin.dll" ( ByVal  iFlag As Integer ) As Integer
'//Declare Function  Elgin_StatusEstendidoMFD   Lib "Elgin.dll" ( ByVal  byRef iStatus As Integer ) As Integer
'//Declare Function  Elgin_VerificaFlagCorteMFD   Lib "Elgin.dll" ( ByVal  byRef iStatus As Integer ) As Integer
'//Declare Function  Elgin_TempoRestanteComprovanteMFD   Lib "Elgin.dll" ( ByVal  cTempo As String ) As Integer
'//Declare Function  Elgin_UFProprietarioMFD   Lib "Elgin.dll" ( ByVal  cUF As String ) As Integer
'//Declare Function  Elgin_GrandeTotalUltimaReducaoMFD   Lib "Elgin.dll" ( ByVal  cGT As String ) As Integer
'//Declare Function  Elgin_DataMovimentoUltimaReducaoMFD   Lib "Elgin.dll" ( ByVal  cData As String ) As Integer
'//Declare Function  Elgin_SubTotalComprovanteNaoFiscalMFD   Lib "Elgin.dll" ( ByVal  cSubTotal As String ) As Integer
'//Declare Function  Elgin_InicioFimCOOsMFD   Lib "Elgin.dll" ( ByVal  cCOOIni, ByVal cCOOFim As String ) As Integer
'//Declare Function  Elgin_InicioFimGTsMFD   Lib "Elgin.dll" ( ByVal  cGTIni, ByVal cGTFim As String ) As Integer

' Funções para a impressão de Código de Barras
'{
'Declare Function  Elgin_TerminaFechamentoCupomCodigoBarrasMFD   Lib "Elgin.dll" ( ByVal  cMensagem As String,
'                                                            cTipoCodigo As String,
'                                                            cCodigo As String,
'                                                            iAltura As Integer,
'                                                            iLargura As Integer,
'                                                            iPosicaoCaracteres As Integer,
'                                                            iFonte As Integer,
'                                                            iMargem As Integer,
'                                                            iCorrecaoErros As Integer,
'                                                            iColunas As Integer ) As Integer
'}

'===============================================================================
'********************************************************************************
'
'                   DECLARAÇÃO DAS FUNÇÕES GLOBAIS DO DemoFit32
'
'********************************************************************************
'===============================================================================}
Public Function TrataRetorno2(ByVal iRetorno As Integer) As Boolean
    Dim strMsgErro As String
    Dim bRetorno As Boolean
    
    bRetorno = False
    If (iRetorno <> 1) Then
        Select Case iRetorno
            Case 0
                If (ObtemRetornoECF(strMsgErro)) Then
                    MsgBox strMsgErro, vbCritical, "Erro na comunicação."
                Else
                    MsgBox "Erro na comunicação.", vbCritical
                End If
            Case -2
                MsgBox "Parâmetro inválido na função.", vbCritical
            Case -4
                MsgBox "O arquivo de inicialização Elgin.ini não foi encontrado no diretório de sistema do Windows.", vbCritical
            Case -5
                MsgBox "Erro ao abrir a porta de comunicação.", vbCritical
            Case -27
                MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2).", vbCritical
            Case Else
                MsgBox "Ocorreu um erro desconhecido. Erro nº " & CStr(iRetorno), vbCritical
        End Select
    Else
'        MsgBox "Operação realizada com sucesso", vbOKOnly
        bRetorno = True
    End If
    'TODO: Obter retorno da impressora
    TrataRetorno2 = bRetorno
    FrmFrente.Mens.Caption = iRetorno
End Function

Public Function ObtemRetornoECF(ByRef strMensagemErro As String) As Boolean
    Dim iRetorno As Integer
    Dim iCodErro As Integer
    Dim strErroMsg As String
    Dim bSucesso As Boolean
    
    strErroMsg = Space(100)
        
    iRetorno = Elgin_RetornoImpressora(iCodErro, strErroMsg)
    
    strMensagemErro = "Erro nº: " & CStr(iCodErro) & " - " & strErroMsg
    
    If (iRetorno = 1) Then
        bSucesso = True
    Else
        bSucesso = False
    End If
    
    ObtemRetornoECF = bSucesso

End Function
