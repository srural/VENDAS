Attribute VB_Name = "Daruma"
'************************************************************************************
'
'                                       FS345
'
'************************************************************************************

Public Declare Function Daruma_DOS_EnviarCmdECF Lib "Daruma32.dll" (ByVal ssBuffer_CMD As String, ByVal iTamanho As Integer, ByVal iiAguardar As String, ByVal iiAcao As String, ByVal iiQuantidade As String) As Integer
Public Declare Function Daruma_DOS_LeBuffer Lib "Daruma32.dll" (ByVal sBuffer As String) As Integer
Public Declare Function Daruma_FI_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_NumeroSerie Lib "Daruma32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Daruma_FI_VerificaIndiceAliquotasIss Lib "Daruma32.dll" (ByVal AliquotaIss As String) As Integer
Public Declare Function Daruma_FI_SubTotal Lib "Daruma32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Daruma_FI_NumeroCupom Lib "Daruma32.dll" (ByVal NumeroCupom As String) As Integer
Public Declare Function Daruma_FI_ResetaImpressora Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AbrePortaSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_LeituraX Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_LeituraXSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AbreCupom Lib "Daruma32.dll" (ByVal CGC_CPF As String) As Integer
Public Declare Function Daruma_FI_VendeItem Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Daruma_FI_VendeItem1Lin13Dig Lib "Daruma32.dll" (ByVal st As String, ByVal cod As String, ByVal d_a As String, ByVal porc As String, ByVal preco As String, ByVal quant As String, ByVal desc As String) As Integer
Public Declare Function Daruma_FI_VendeItem1Lin6Dig Lib "Daruma32.dll" (ByVal st As String, ByVal cod As String, ByVal d_a As String, ByVal porc As String, ByVal preco As String, ByVal quant As String, ByVal desc As String) As Integer
Public Declare Function Daruma_FI_IdentificaConsumidor Lib "Daruma32.dll" (ByVal Mensagem As String, ByVal msg2 As String, ByVal msg3 As String) As Integer
Public Declare Function Daruma_FI_CancelaItemAnterior Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_CancelaItemGenerico Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FI_CancelaCupom Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FechaCupomResumido Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_ReducaoZ Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Daruma_FI_FechaCupom Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_VendeItemDepartamento Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Daruma_FI_AumentaDescricaoItem Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
Public Declare Function Daruma_FI_UsaUnidadeMedida Lib "Daruma32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Daruma_FI_AlteraSimboloMoeda Lib "Daruma32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Daruma_FI_ProgramaAliquota Lib "Daruma32.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Daruma_FI_ProgramaHorarioVerao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_NomeiaDepartamento Lib "Daruma32.dll" (ByVal indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Daruma_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "Daruma32.dll" (ByVal indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Daruma_FI_ProgramaArredondamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ProgramaTruncamento Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_LinhasEntreCupons Lib "Daruma32.dll" (ByVal Linhas As Integer) As Integer
Public Declare Function Daruma_FI_EspacoEntreLinhas Lib "Daruma32.dll" (ByVal Dots As Integer) As Integer
Public Declare Function Daruma_FI_RelatorioGerencial Lib "Daruma32.dll" (ByVal cTexto As String) As Integer
Public Declare Function Daruma_FI_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Daruma_FI_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal texto As String) As Integer
Public Declare Function Daruma_FI_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_Sangria Lib "Daruma32.dll" (ByVal Valor As String) As Integer
Public Declare Function Daruma_FI_Suprimento Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalData Lib "Daruma32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalReducao Lib "Daruma32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialData Lib "Daruma32.dll" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialReducao Lib "Daruma32.dll" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Daruma_FI_VersaoFirmware Lib "Daruma32.dll" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Daruma_FI_CGC_IE Lib "Daruma32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Daruma_FI_GrandeTotal Lib "Daruma32.dll" (ByVal GrandeTotal As String) As Integer
Public Declare Function Daruma_FI_Cancelamentos Lib "Daruma32.dll" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Daruma_FI_Descontos Lib "Daruma32.dll" (ByVal ValorDescontos As String) As Integer
Public Declare Function Daruma_FI_NumeroOperacoesNaoFiscais Lib "Daruma32.dll" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Daruma_FI_NumeroCuponsCancelados Lib "Daruma32.dll" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Daruma_FI_NumeroIntervencoes Lib "Daruma32.dll" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Daruma_FI_NumeroReducoes Lib "Daruma32.dll" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Daruma_FI_NumeroSubstituicoesProprietario Lib "Daruma32.dll" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Daruma_FI_UltimoItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Daruma_FI_ClicheProprietario Lib "Daruma32.dll" (ByVal Cliche As String) As Integer
Public Declare Function Daruma_FI_NumeroCaixa Lib "Daruma32.dll" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Daruma_FI_NumeroLoja Lib "Daruma32.dll" (ByVal NumeroLoja As String) As Integer
Public Declare Function Daruma_FI_SimboloMoeda Lib "Daruma32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Daruma_FI_MinutosLigada Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Daruma_FI_MinutosImprimindo Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
Public Declare Function Daruma_FI_VerificaModoOperacao Lib "Daruma32.dll" (ByVal Modo As String) As Integer
Public Declare Function Daruma_FI_VerificaEpromConectada Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_FlagsFiscais Lib "Daruma32.dll" (ByRef Flag As Integer) As Integer
Public Declare Function Daruma_FI_ValorPagoUltimoCupom Lib "Daruma32.dll" (ByVal ValorCupom As String) As Integer
Public Declare Function Daruma_FI_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Daruma_FI_ContadoresTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Contadores As String) As Integer
Public Declare Function Daruma_FI_VerificaTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
Public Declare Function Daruma_FI_DataHoraReducao Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Daruma_FI_DataMovimento Lib "Daruma32.dll" (ByVal Data As String) As Integer
Public Declare Function Daruma_FI_VerificaTruncamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_Acrescimos Lib "Daruma32.dll" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Daruma_FI_ContadorBilhetePassagem Lib "Daruma32.dll" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Daruma_FI_VerificaAliquotasIss Lib "Daruma32.dll" (ByVal AliquotasIss As String) As Integer
Public Declare Function Daruma_FI_VerificaFormasPagamento Lib "Daruma32.dll" (ByVal Formas As String) As Integer
Public Declare Function Daruma_FI_VerificaRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Recebimentos As String) As Integer
Public Declare Function Daruma_FI_VerificaDepartamentos Lib "Daruma32.dll" (ByVal Departamentos As String) As Integer
Public Declare Function Daruma_FI_VerificaTipoImpressora Lib "Daruma32.dll" (ByRef TipoImpressora As Integer) As Integer
Public Declare Function Daruma_FI_VerificaTotalizadoresParciais Lib "Daruma32.dll" (ByVal cTotalizadores As String) As Integer
Public Declare Function Daruma_FI_RetornoAliquotas Lib "Daruma32.dll" (ByVal cAliquotas As String) As Integer
Public Declare Function Daruma_FI_VerificaEstadoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Daruma_FI_DadosUltimaReducao Lib "Daruma32.dll" (ByVal DadosReducao As String) As Integer
Public Declare Function Daruma_FI_MonitoramentoPapel Lib "Daruma32.dll" (ByRef Linhas As Integer) As Integer
Public Declare Function Daruma_FI_Autenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ProgramaCaracterAutenticacao Lib "Daruma32.dll" (ByVal Parametros As String) As Integer
Public Declare Function Daruma_FI_AcionaGaveta Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaEstadoGaveta Lib "Daruma32.dll" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Daruma_FI_ProgramaMoedaSingular Lib "Daruma32.dll" (ByVal MoedaSingular As String) As Integer
Public Declare Function Daruma_FI_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal FormasPG As String) As Integer
Public Declare Function Daruma_FI_ProgramaMoedaPlural Lib "Daruma32.dll" (ByVal MoedaPlural As String) As Integer
Public Declare Function Daruma_FI_CancelaImpressaoCheque Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_VerificaStatusCheque Lib "Daruma32.dll" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Daruma_FI_ImprimeCheque Lib "Daruma32.dll" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_ImprimeCopiaCheque Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_IncluiCidadeFavorecido Lib "Daruma32.dll" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Daruma_FI_EstornoFormasPagamento Lib "Daruma32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Public Declare Function Daruma_FI_ForcaImpactoAgulhas Lib "Daruma32.dll" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Daruma_FI_FechaPortaSerial Lib "Daruma32.dll" () As Integer

Public Declare Function Daruma_FI_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Daruma_FI_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Daruma_FI_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem As String) As Integer
Public Declare Function Daruma_FI_AbreBilhetePassagem Lib "Daruma32.dll" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal Hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Daruma_FI_MapaResumo Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioTipo60Analitico Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RelatorioTipo60Mestre Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ImprimeConfiguracoesImpressora Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ImprimeDepartamentos Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AberturaDoDia Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Daruma_FI_FechamentoDoDia Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_ValorFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Daruma_FI_ValorTotalizadorNaoFiscal Lib "Daruma32.dll" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
Public Declare Function Daruma_FI_VerificaDocAutenticacao Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_AutenticacaoStr Lib "Daruma32.dll" (ByVal Txto As String) As Integer
Public Declare Function Daruma_FI_StatusRelatorioGerencial Lib "Daruma32.dll" (ByVal StatusRel As String) As Integer
Public Declare Function Daruma_FI_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StatusRel As String) As Integer
Public Declare Function Daruma_FI_StatusComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal StatusRel As String) As Integer
Public Declare Function Daruma_FI_ImprimeConfiguracoes Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_FechaRelatorioXouZ Lib "Daruma32.dll" () As Integer


'funções de Configuração
Public Declare Function Daruma_FI_CfgFechaAutomaticoCupom Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgRedZAutomatico Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgLeituraXAuto Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgCalcArredondamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgHorarioVerao Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgSensorAut Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgLegProdutos Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgEspacamentoCupons Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgHoraMinReducaoZ Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgLimiarNearEnd Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgPermMensPromCNF Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgImpEstGavVendas Lib "Daruma32.dll" (ByVal Flag As String) As Integer
Public Declare Function Daruma_FI_CfgCupomAdicional Lib "Daruma32.dll" (ByVal Flag As String) As Integer


'Funcoes de Registry
Public Declare Function Daruma_Registry_Porta Lib "Daruma32.dll" (ByVal cKeyPorta As String) As Integer
Public Declare Function Daruma_Registry_Path Lib "Daruma32.dll" (ByVal cKeyPath As String) As Integer
Public Declare Function Daruma_Registry_Status Lib "Daruma32.dll" (ByVal cKeyStatus As String) As Integer
Public Declare Function Daruma_Registry_StatusFuncao Lib "Daruma32.dll" (ByVal cKeyStatusFuncao As String) As Integer
Public Declare Function Daruma_Registry_Retorno Lib "Daruma32.dll" (ByVal cKeyRetorno As String) As Integer
Public Declare Function Daruma_Registry_ControlePorta Lib "Daruma32.dll" (ByVal cKeyControlePorta As String) As Integer
Public Declare Function Daruma_Registry_ModoGaveta Lib "Daruma32.dll" (ByVal cKeyModoGaveta As String) As Integer
Public Declare Function Daruma_Registry_ConfigRede Lib "Daruma32.dll" (ByVal cKeyConfigRede As String) As Integer
Public Declare Function Daruma_Registry_Log Lib "Daruma32.dll" (ByVal cKeyLog As String) As Integer
Public Declare Function Daruma_Registry_NomeLog Lib "Daruma32.dll" (ByVal cKeyNomeLog As String) As Integer
Public Declare Function Daruma_Registry_Emulador Lib "Daruma32.dll" (ByVal cKeyEmulador As String) As Integer
Public Declare Function Daruma_Registry_Separador Lib "Daruma32.dll" (ByVal cKeySeparador As String) As Integer
Public Declare Function Daruma_Registry_SeparaMsgPromo Lib "Daruma32.dll" (ByVal cKeySeparaMsgPromo As String) As Integer
Public Declare Function Daruma_Registry_VendeItemUmaLinha Lib "Daruma32.dll" (ByVal cKeyVendeItem1Lin13Dig As String) As Integer
Public Declare Function Daruma_Registry_XAutomatica Lib "Daruma32.dll" (ByVal cKeyXAutomatica As String) As Integer
Public Declare Function Daruma_Registry_ZAutomatica Lib "Daruma32.dll" (ByVal cKeyZAutomatica As String) As Integer
Public Declare Function Daruma_Registry_AlteraRegistry Lib "Daruma32.dll" (ByVal cKey As String, ByVal cValue As String) As Integer
Public Declare Function Daruma_Registry_ImprimeRegistry Lib "Daruma32.dll" (ByVal Produto As String) As Integer
Public Declare Function Daruma_Registry_RetornaValor Lib "Daruma32.dll" (ByVal Produto As String, ByVal cKey As String, ByVal cKeyValor As String) As Integer

Public Declare Function Daruma_FI_VerificaDescricaoFormasPagamento Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_VerificaFormasPagamentoEx Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaAcrescimoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCFCancelados Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFCancelados Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCLX Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFNV Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCNFV Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCRO Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCRZ Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCRZRestante Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaCancelamentoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaDescontoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaGNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaTempoImprimindo Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaTempoLigado Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaTotalPagamentos Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_RetornaTroco Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_ClicheProprietarioEx Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_COO Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FI_VerificaModeloECF Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FI_RetornaErroExtendido Lib "Daruma32.dll" (ByVal cParam As String) As Integer


'Funcoes de TEF
Public Declare Function Daruma_TEF_TravarTeclado Lib "Daruma32.dll" (ByVal Travar As String) As Integer
Public Declare Function Daruma_TEF_EsperarArquivo Lib "Daruma32.dll" (ByVal sArquivo As String, ByVal sTempo As String, ByVal sTravar As String) As Integer
Public Declare Function Daruma_TEF_ImprimirResposta Lib "Daruma32.dll" (ByVal sArquivo As String, ByVal cForma As String, ByVal sTravar As String) As Integer
Public Declare Function Daruma_TEF_SetFocus Lib "Daruma32.dll" (ByVal cTitleWindow As String) As Integer
Public Declare Function Daruma_TEF_FechaRelatorio Lib "Daruma32.dll" () As Integer
                                                                              
'************************************************************************************
'
'                                       FS315
'
'************************************************************************************
'Cupom

Public Declare Function Daruma_FIR_AbreCupomRestaurante Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_FechaCupomRestaurante Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_FechaCupomRestauranteResumido Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_IdentificaConsumidor Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ImprimePrimeiroCupomDividido Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RestanteCupomDividido Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_IniciaFechamentoCupomComServico Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer

'Mesa
Public Declare Function Daruma_FIR_CancelarVenda Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ConferenciaMesa Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_FechaConferenciaMesa Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RegistrarVenda Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RegistroVendaSerial Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RelatorioMesasAbertas Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_TranferirVenda Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_TranferirMesa Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer


'Relatórios Fiscais
  

Public Declare Function Daruma_FIR_ReducaoZ Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_LeituraX Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalData Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalReducao Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalSerialData Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaFiscalSerialReducao Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer

'Funções de Operações não fiscais e Vinculados

Public Declare Function Daruma_FIR_RelatorioGerencial Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer


'Funções de Gaveta de Dinheiro
Public Declare Function Daruma_FIR_AcionaGaveta Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_VerificaEstadoGaveta Lib "Daruma32.dll" (ByRef iParam As Integer) As Integer
'Funções de Autenticação

Public Declare Function Daruma_FIR_AutenticacaoStr Lib "Daruma32.dll" (ByVal cParam As String) As Integer

'Informação
Public Declare Function Daruma_FIR_VerificaEstadoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Daruma_FIR_StatusComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_StatusCupomFiscal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_StatusRelatorioGerencial Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornoAliquotas Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaTotalizadoresParciais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_SubTotal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroSerie Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VersaoFirmware Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_GrandeTotal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroCupom Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroOperacoesNaoFiscais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroCuponsCancelados Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroReducoes Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroIntervencoes Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ClicheProprietario Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroCaixa Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_TempoOperacional Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_MinutosEmititindoDocFiscal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_DataHoraImpressora Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ContadoresTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaAliquotasIss Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaFormasPagamentoEx Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaIndiceAliquotasIss Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ValorFormaPagamento Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ValorTotalizadorNaoFiscal Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_VerificaModeloECF Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_DescontosNaoFiscais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_AcrescimoNaoFiscais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CancelamentoNaoFiscais Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_TotalTroco Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ReducoesRestantes Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CooComecoDia Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NumeroCancelamentoNaoFiscal Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaErroExtendido Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaAcrescimoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCFCancelados Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFCancelados Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCLX Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFNV Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCNFV Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRO Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRZ Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCRZRestante Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaCancelamentoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaDescontoNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaGNF Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaTempoImprimindo Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaTempoLigado Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaTotalPagamentos Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornaTroco Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RetornoImpressora Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_MesasLivres Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_RegistroVendaLivre Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CardapioLivre Lib "Daruma32.dll" (ByVal cParam As String) As Integer

'Programação e Configuração do ECF

'Inicialização
Public Declare Function Daruma_FIR_ProgramaAliquota Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_NomeiaTotalizadorNaoSujeitoIcms Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ProgramaOperador Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ProgramaMsgTaxaServico Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_AdicionaProdutoCardapio Lib "Daruma32.dll" (ByVal cParam As String, ByVal cParam As String, ByVal cParam As String, ByVal cParam As String) As Integer

'Configuração
Public Declare Function Daruma_FIR_CfgHorarioVerao Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CfgLegProdutos Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CfgEspacamentoCupons Lib "Daruma32.dll" (ByVal cParam As String) As Integer
Public Declare Function Daruma_FIR_CfgLimiarNearEnd Lib "Daruma32.dll" (ByVal cParam As String) As Integer

'Outras e Utilidades
Public Declare Function Daruma_FIR_ZeraCardapio Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_ImprimeCardapio Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_CardapioSerial Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_LeituraMemoriaTrabalho Lib "Daruma32.dll" () As Integer
Public Declare Function Daruma_FIR_RelatorioMesasAbertasSerial Lib "Daruma32.dll" () As Integer
