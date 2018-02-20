Attribute VB_Name = "Modulo_SiTef"

Public Declare Function ConfiguraIntSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByRef pProximoComando As Long, ByRef pTipoCampo As Long, ByRef pTamanhoMinimo As Integer, ByRef pTamanhoMaximo As Integer, ByVal pBuffer As String, ByVal TamMaxBuffer As Long, ByVal ContinuaNavegacao As Long) As Long

Global Resultado     As Long



Public Function retornoFuncoesConfiguracoes(codigo As String)

    Select Case codigo
        Case 0
            retornoFuncoesConfiguracoes = "Não ocorreu erro "
        Case 1
            retornoFuncoesConfiguracoes = "Endereço IP inválido ou não resolvido "
        Case 2
            retornoFuncoesConfiguracoes = "Código da loja inválido "
        Case 3
            retornoFuncoesConfiguracoes = "Código de terminal inválido "
        Case 6
            retornoFuncoesConfiguracoes = "Erro na inicialização do TcpretornoFuncoesConfiguracoes= /Ip "
        Case 7
            retornoFuncoesConfiguracoes = "Falta de memória "
        Case 8
            retornoFuncoesConfiguracoes = "Não encontrou a CliSiTef ou ela está com problemas "
        Case 9
            retornoFuncoesConfiguracoes = "Configuração de servidores SiTef foi excedida. "
        Case 10
            retornoFuncoesConfiguracoes = "Erro de acesso na pasta CliSiTef (possível falta de permissão para escrita) "
        Case 11
            retornoFuncoesConfiguracoes = "Dados inválidos passados pela automação. "
        Case 12
            retornoFuncoesConfiguracoes = "Modo seguro não ativo (possível falta de configuração no servidor SiTef do arquivo .cha). "
        Case 13
            retornoFuncoesConfiguracoes = "Caminho DLL inválido (o caminho completo das bibliotecas está muito grande)."
            
        Case Else
        retornoFuncoesConfiguracoes = "[ERRO] Código de retorno desconhecido"
        
    End Select

End Function

