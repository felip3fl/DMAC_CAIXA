Attribute VB_Name = "Variaveis"
'---------------------------------------------
 '               DECLARAÇÕES
'---------------------------------------------

Global ConexaoDLLAdo As New DMACD.conexaoADO

Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal Numerocupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal quantidade As String, ByVal CasasDecimais As Integer, ByVal valorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal valorUnitario As String, ByVal quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BEMAFI32.DLL" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BEMAFI32.DLL" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BEMAFI32.DLL" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BEMAFI32.DLL" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BEMAFI32.DLL" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal IndiceTotalizador As String, ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal valor As String, ByVal Numerocupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BEMAFI32.DLL" (ByVal valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BEMAFI32.DLL" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BEMAFI32.DLL" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BEMAFI32.DLL" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BEMAFI32.DLL" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BEMAFI32.DLL" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BEMAFI32.DLL" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BEMAFI32.DLL" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BEMAFI32.DLL" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BEMAFI32.DLL" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BEMAFI32.DLL" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BEMAFI32.DLL" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BEMAFI32.DLL" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BEMAFI32.DLL" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BEMAFI32.DLL" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BEMAFI32.DLL" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BEMAFI32.DLL" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BEMAFI32.DLL" (ByVal AliquotasIss As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BEMAFI32.DLL" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BEMAFI32.DLL" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BEMAFI32.DLL" (ByRef tipoImpressora As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BEMAFI32.DLL" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BEMAFI32.DLL" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BEMAFI32.DLL" (ByRef Linhas As Integer) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BEMAFI32.DLL" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BEMAFI32.DLL" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BEMAFI32.DLL" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BEMAFI32.DLL" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BEMAFI32.DLL" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BEMAFI32.DLL" (ByVal Banco As String, ByVal valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BEMAFI32.DLL" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BEMAFI32.DLL" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal valor As String) As Integer

Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BEMAFI32.DLL" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Analitico Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Mestre Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeConfiguracoesImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AberturaDoDia Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechamentoDoDia Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ValorFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscal Lib "BEMAFI32.DLL" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer


'Funções para Impressora restaurante
Public Declare Function Bematech_FIR_RegistraVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal quantidade As String, ByVal valorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal quantidade As String, ByVal valorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BEMAFI32.DLL" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal quantidade As String, ByVal valorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertas Lib "BEMAFI32.DLL" (ByVal TipoRelatorio As Integer) As Integer
Public Declare Function Bematech_FIR_ImprimeCardapio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertasSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_CardapioPelaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RegistroVendaSerial Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_VerificaMemoriaLivre Lib "BEMAFI32.DLL" (ByVal Bytes As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomResumidoRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal mensagem As String) As Integer

' Funções para o TEF

Public Declare Function Bematech_FITEF_Status Lib "BEMAFI32.DLL" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCartao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ConfirmaVenda Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Header As String) As Integer
Public Declare Function Bematech_FITEF_NaoConfirmaVendaImpressao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_CancelaVendaCartao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Nsu As String, ByVal Numerocupom As String, ByVal hora As String, ByVal Data As String, ByVal Rede As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeTEF Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal FormaPagamento As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeRelatorio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FITEF_ADM Lib "BEMAFI32.DLL" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCompleta Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal FormaPagamento As String, ByVal Texto As String) As Integer
Public Declare Function Bematech_FITEF_ConfiguraDiretorioTef Lib "BEMAFI32.DLL" (ByVal PathArqReq As String, ByVal PathArqResp As String) As Integer
Public Declare Function Bematech_FITEF_VendaCheque Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal valor As String) As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long




'---------------------------------------------

Option Explicit
Global wConta As Long
Global wContItem As Integer
Global wContaC As Long
Global wPedido As String
Global Glb_NfDevolucao As Boolean
Global Wav As String
Global wCFOItem As Double
Global wCFO1 As String * 6
Global wCFO2 As String * 6
Global wCarimbo1 As String * 132
Global wCarimbo2 As String * 132
Global wCarimbo3 As String * 132
Global wCarimbo4 As String * 132
Global wCarimbo5 As String * 132
Global wStr0, wStr1, wStr2, wStr3, wStr4, wStr5, wStr6, wStr7 As String
Global wStr8, wStr9, wStr10, wStr11, wStr12, wStr13, wStr15, wStr16, wStr17, wStr18, wStr19, wStr20, wStr21, wStr22 As String
Global wRecebeCarimboAnexo As String * 132
Global wEndImp As String
Global wAnexoIten As Integer
Global WAnexoAux As String * 20
Global wAnexo As String
Global wAnexo1 As String
Global wAnexo2 As String
Global wRomaneio As Boolean
Global wValorRomaneio As Double

Global adoCNMatriz As New ADODB.Connection
Global rdoCNLoja As New ADODB.Connection
Global rdoCNRetaguarda As New ADODB.Connection
Global rdoCNTEF As New ADODB.Connection
Global wConectouRetaguarda As Boolean
Global rdoCnLojaBach  As New ADODB.Connection
Global RsDados As New ADODB.Recordset
Global pegadados As New ADODB.Recordset
Global RsDadosLoja As New ADODB.Recordset
Global rsItensVenda As New ADODB.Recordset
Global rsNFE As New ADODB.Recordset

Global RsdadosItens As New ADODB.Recordset
Global rdoParametro As New ADODB.Recordset
Global rdoTrans As New ADODB.Recordset

Global rdocontrole As New ADODB.Recordset
Global rsTEF As New ADODB.Recordset
Global rspegaloja As New ADODB.Recordset
Global RsPegaValorCodigoZero As New ADODB.Recordset
Global RsDadosTef As New ADODB.Recordset
Global rdoContaItens As New ADODB.Recordset
Global ADOCancela As New ADODB.Recordset
Global ADOSituacao As New ADODB.Recordset
Global RsPegaGrupoMovCaixa As New ADODB.Recordset
Global rsOperacoes As New ADODB.Recordset

Global ISQL As New ADODB.Recordset

Global adoCNAccess As New ADODB.Connection
Global adoCNLoja As New ADODB.Connection
Global rdoConexaoINI As New ADODB.Recordset
Global rdoParametroINI As New ADODB.Recordset
Global rdoComplemento As New ADODB.Recordset
Global RsPegaItensEspeciais As New ADODB.Recordset
Global FechaRetaguarda As New ADODB.Recordset
Global rdoCNLojaINI As New ADODB.Connection
Global RsDadosINI As New ADODB.Recordset
Global rdoSerie As New ADODB.Recordset
Global PegaLoja As New ADODB.Recordset
Global PegaSerie As New ADODB.Recordset
Global RsDadosCapa As New ADODB.Recordset
Global rsVerItensPed As New ADODB.Recordset
Global rsVerProdPed As New ADODB.Recordset
Global rdoDadosProdu As New ADODB.Recordset
Global RsItensNF As New ADODB.Recordset
Global rsItemNota  As New ADODB.Recordset
Global RsCapaNF As New ADODB.Recordset
Global RsICMSIntER As New ADODB.Recordset
Global rdoVlNota As New ADODB.Recordset
Global rdoVlItem As New ADODB.Recordset
Global rdoConPag As New ADODB.Recordset
Global rdoModalidade As New ADODB.Recordset
Global rdoTransfNumerario As New ADODB.Recordset
Global TipoPedido As String
Global wValoTotalNotaAlternativa As Double
Global wSerie As String
Global wPagina As Integer

Global wPegaCliente As String
Global wPegaDesconto As Double
Global wPegaFrete As Double
Global wDocumento   As String
Global wPessoa As Double

Global wErroApresenta As Byte

Global wLoja As String * 5
Global wMensagemECF As String * 48
Global wAdicionaisECF As String * 48
Global wSenhaLiberacao As String * 6
Global rdoPegaCliente As New ADODB.Recordset
Global rdoRegiao As New ADODB.Recordset
Global RdoComponente As New ADODB.Recordset
Global rdoLoja As New ADODB.Recordset
Global rdoProduto As New ADODB.Recordset
Global rdoFormaPagamento As New ADODB.Recordset
Global rdoItensVenda As New ADODB.Recordset
Global rdoFechamentoGeral As New ADODB.Recordset
Global rdoCapa As New ADODB.Recordset
Global rdoDataFechamentoRetaguarda As New ADODB.Recordset
Global rdoDataFechamento As New ADODB.Recordset
Global RsSaldoCaixa As New ADODB.Recordset
Global RsControleCaixa As New ADODB.Recordset
Global RsCarimbo    As New ADODB.Recordset
Global rsComplementoVenda As New ADODB.Recordset
Global rsProdutoGarantiaEstendida As New ADODB.Recordset
Global GLB_Loja As String
Global GLB_NF As Long
Global GLB_Serie As String
Global GLB_SerieCF As String
Global GLB_USU_Nome As String
Global GLB_USU_Codigo As String
Global GLB_CTR_Protocolo As Long
Global GLB_ECF As String
Global GLB_ServidorTEF As String
Global GLB_BancoTEF As String
Global GLB_UsuarioTEF As String
Global GLB_SenhaTEF As String
Global GLB_Impressora00 As String
Global Glb_AlteraResolucao As Boolean
'Global GLB_EnderecoPortal As String
Global GLB_EnderecoPastaRESP As String
Global GLB_EnderecoPastaFIL As String
Global GLB_Caixa As String
Global GLB_Administrador As Boolean
Global GLB_ADMNome As String
Global GLB_ADMProtocolo As Integer
Global GLB_VerificaImpressoraFiscal As String
Global GLB_NomeServidor As String
Global GLB_NomeBanco As String
Global GLB_Usuario As String
Global GLB_Senha As String
Global Glb_BancoLocal As String
Global GLB_Banco As String
Global GLB_Servidorlocal As String
Global GLB_Servidor As String
Global Glb_ImpNotaFiscal As String
Global GLB_ConectouOK As Boolean
Global GLB_DataInicial As String
Global GLB_HoraInicial As String
Global GLB_DataFinal As String
Global GLB_Logo As String
Global pedido As Long
Global emitiNota As Boolean
Global cancelaNota As Boolean
Global cancelaNotaResultado As Boolean
Global NroPedido As Integer
Global wItensVenda As Long
Global wItens  As Integer
Global wTotalVenda As Double
Global wtotalitens As Long
Global wtotalGarantia As Double
Global NroItens As Long
Global FechouATelaFormapagamentonoX As Boolean
Global saldoAnterior As Long
Global lsDSN As String

Global wIE_icmsAplicado As Double
Global wIE_icmsFECPAplicado As Double
Global wIE_icmsFECPDiferencial As Double
Global wIE_icmsFECPPart As Double
Global wIE_icmsFECPUFDEST As Double
Global wIE_icmsFECPUFDESTTotal As Double
Global wIE_icmsFECPUFREMET As Double
Global wIE_icmsFECPUFREMETTotal As Double
Global wIE_icmsFECPAliqDest As Double
Global wIE_icmsFECPAliqInter As Double
Global wIE_Tributacao  As String * 3
Global wIE_Cfo  As Integer
Global wIE_BasedeReducao  As Double
Global wIE_icmsdestino   As Double
Global wST20 As String
Global wST60 As String

Global Wecf As Integer

Global NroNotaFiscal As Long
Global wNotaFiscalReemissao As Long
Global wSerieReemissao As String
Global wLinha As Integer
Global wlblloja As String
Global wUSU_Nome As String
Global wCTR_Protocolo As String
Global wST1 As Integer
Global wST2 As Integer
Global wValorRetorno As String
Global wPLISTA As Double
Global wICMS As Double
Global wCodBarra As String
Global wSecao As Integer
Global wVlTotItem As Double
Global wIcmPdv As Integer
Global ValDinheiro As Double
Global ValTroco As Double
Global ValCheque As Double
Global ValCartao As Double
Global TotPago As Double
Global wCodigo As String
Global wValoraPagarNORMAL As Double
Global wNomeservidor As String
Global wNomeBanco As String
Global I As Integer
Global wNomeservidorMatriz As String
Global wNomeBancoMatriz As String

Global Wusuario As String
Global wSenha As String
Global Faturada As Boolean
Global Financiada As Boolean

Global GLB_TotalICMSCalculado As Double
Global GLB_ValorCalculadoICMS As Double
Global GLB_BasedeCalculoICMS As Double
Global GLB_AliquotaAplicadaICMS As Double
Global GLB_AliquotaICMS As Double
Global GLB_BaseTotalICMS As Double
Global GLB_Tributacao As String * 3
Global GLB_CFOP As String
Global TipoProcesso As Integer
Global Servidor As String
Global Banco As String

Public Retorno As Integer
Public Funcao As Integer
Public LocalRetorno As String
Global wlblTotalvenda As Double
Global wImpressoraNota As String

Global wRazao As String
Global wNovaRazao As String
Global wCGC As String
Global WIest As String
Global Wendereco As String
Global wbairro As String
Global WMunicipio As String
Global westado As String
Global WCep As String
Global WFone As String
Global WQtdeCopiaNE As Integer
Global wDDDLoja As String
Global WFax As String
Global wtxtCGC_CPF As String
Global wCupomAberto As Boolean
 
Global wNumeroCupom As String * 6
Global wData As Date
Global wValorParcela As Double
Global wMid As Integer

Global wRestoItens As Integer
Global wTotalLinha As Integer
Global wContLinha As Integer
Global wContCarimbo As Integer
Global wTotalCarimbo As Integer
Global wQtdeItensNF As Integer
Global wPermitirVenda As Boolean
Global wQdteViasImpressao As Integer

Global adoConsulta As New ADODB.Recordset
Global RsMovimentoCaixa As New ADODB.Recordset

Global NomeImpressora As Printer

'''''Global wPagamentoECF As Integer

Global GLB_Pessoa As String
Global wFechamentoGeral As Boolean

Global tipoCupomEmite As String

Global tipoZero As Boolean

Global Nf_Dev As Double
Global Serie_Dev As String
Global NfDev_Dev As Double
Global SerieDev_Dev As String
Global DataDev_Dev As String
Global ValorNotaCredito_Dev As Double
Global NotaCredito_Dev As Double
Global ReImpressao_Dev As Boolean

'Global GLB_EnderecoPortal As String


