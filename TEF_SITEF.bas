Attribute VB_Name = "Modulo_SiTef"

Public Declare Function ConfiguraIntSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByRef pProximoComando As Long, ByRef pTipoCampo As Long, ByRef pTamanhoMinimo As Integer, ByRef pTamanhoMaximo As Integer, ByVal pBuffer As String, ByVal TamMaxBuffer As Long, ByVal ContinuaNavegacao As Long) As Long

Global Resultado     As Long



Public Function retornoFuncoesConfiguracoes(codigo As String)

    Select Case codigo
        Case 0
            retornoFuncoesConfiguracoes = "N�o ocorreu erro "
        Case 1
            retornoFuncoesConfiguracoes = "Endere�o IP inv�lido ou n�o resolvido "
        Case 2
            retornoFuncoesConfiguracoes = "C�digo da loja inv�lido "
        Case 3
            retornoFuncoesConfiguracoes = "C�digo de terminal inv�lido "
        Case 6
            retornoFuncoesConfiguracoes = "Erro na inicializa��o do TcpretornoFuncoesConfiguracoes= /Ip "
        Case 7
            retornoFuncoesConfiguracoes = "Falta de mem�ria "
        Case 8
            retornoFuncoesConfiguracoes = "N�o encontrou a CliSiTef ou ela est� com problemas "
        Case 9
            retornoFuncoesConfiguracoes = "Configura��o de servidores SiTef foi excedida. "
        Case 10
            retornoFuncoesConfiguracoes = "Erro de acesso na pasta CliSiTef (poss�vel falta de permiss�o para escrita) "
        Case 11
            retornoFuncoesConfiguracoes = "Dados inv�lidos passados pela automa��o. "
        Case 12
            retornoFuncoesConfiguracoes = "Modo seguro n�o ativo (poss�vel falta de configura��o no servidor SiTef do arquivo .cha). "
        Case 13
            retornoFuncoesConfiguracoes = "Caminho DLL inv�lido (o caminho completo das bibliotecas est� muito grande)."
            
        Case Else
        retornoFuncoesConfiguracoes = "[ERRO] C�digo de retorno desconhecido"
        
    End Select

End Function

