Attribute VB_Name = "Declaracoes"

Public Declare Function Bematech_FI_NumeroSerie Lib "BemaFi32.dll" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BemaFi32.dll" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BemaFi32.dll" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BemaFi32.dll" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BemaFi32.dll" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BemaFi32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BemaFi32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BemaFi32.dll" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BemaFi32.dll" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BemaFi32.dll" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BemaFi32.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BemaFi32.dll" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BemaFi32.dll" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BemaFi32.dll" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BemaFi32.dll" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BemaFi32.dll" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BemaFi32.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BemaFi32.dll" (ByVal Texto As String) As Integer
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


'Funções para Impressora restaurante
Public Declare Function Bematech_FIR_RegistraVenda Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BemaFi32.dll" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BemaFi32.dll" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BemaFi32.dll" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BemaFi32.dll" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BemaFi32.dll" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BemaFi32.dll" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BemaFi32.dll" (ByVal MesaOrigem As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertas Lib "BemaFi32.dll" (ByVal TipoRelatorio As Integer) As Integer
Public Declare Function Bematech_FIR_ImprimeCardapio Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertasSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_CardapioPelaSerial Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FIR_RegistroVendaSerial Lib "BemaFi32.dll" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_VerificaMemoriaLivre Lib "BemaFi32.dll" (ByVal Bytes As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomRestaurante Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomResumidoRestaurante Lib "BemaFi32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer

' Funções para o TEF

Public Declare Function Bematech_FITEF_Status Lib "BemaFi32.dll" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCartao Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ConfirmaVenda Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Header As String) As Integer
Public Declare Function Bematech_FITEF_NaoConfirmaVendaImpressao Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_CancelaVendaCartao Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Nsu As String, ByVal NumeroCupom As String, ByVal Hora As String, ByVal Data As String, ByVal Rede As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeTEF Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal FormaPagamento As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeRelatorio Lib "BemaFi32.dll" () As Integer
Public Declare Function Bematech_FITEF_ADM Lib "BemaFi32.dll" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCompleta Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal FormaPagamento As String, ByVal Texto As String) As Integer
Public Declare Function Bematech_FITEF_ConfiguraDiretorioTef Lib "BemaFi32.dll" (ByVal PathArqReq As String, ByVal PathArqResp As String) As Integer
Public Declare Function Bematech_FITEF_VendaCheque Lib "BemaFi32.dll" (ByVal Identificacao As String, ByVal Valor As String) As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Retorno As Integer
Public Funcao As Integer
Public LocalRetorno As String

