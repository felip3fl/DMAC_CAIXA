Attribute VB_Name = "Modulo_SiTef"

Public Declare Function ConfiguraIntSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByRef pProximoComando As Long, ByRef pTipoCampo As Long, ByRef pTamanhoMinimo As Integer, ByRef pTamanhoMaximo As Integer, ByVal pBuffer As String, ByVal TamMaxBuffer As Long, ByVal ContinuaNavegacao As Long) As Long

Private Declare Function LeSimNaoPinPad Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Funcao As String) As Long
Private Declare Function EscreveMensagemPermanentePinPad Lib "C:\Users\felipelima\Desktop\DMAC_Caixa\CliSitef32I.dll" (ByVal Funcao As String) As Long

Global Resultado     As Long
Global ComprovantePagamento As String
Global GLB_TefHabilidado As Boolean




Public Function retornoFuncoesConfiguracoes(codigo As String)

    Select Case codigo
        Case 0
            retornoFuncoesConfiguracoes = "0 Não ocorreu erro "
        Case 1
            retornoFuncoesConfiguracoes = "1 Endereço IP inválido ou não resolvido "
        Case 2
            retornoFuncoesConfiguracoes = "2 Código da loja inválido "
        Case 3
            retornoFuncoesConfiguracoes = "3 Código de terminal inválido "
        Case 6
            retornoFuncoesConfiguracoes = "6 Erro na inicialização do TcpretornoFuncoesConfiguracoes= /Ip "
        Case 7
            retornoFuncoesConfiguracoes = "7 Falta de memória "
        Case 8
            retornoFuncoesConfiguracoes = "8 Não encontrou a CliSiTef ou ela está com problemas "
        Case 9
            retornoFuncoesConfiguracoes = "9 Configuração de servidores SiTef foi excedida. "
        Case 10
            retornoFuncoesConfiguracoes = "10 Erro de acesso na pasta CliSiTef (possível falta de permissão para escrita) "
        Case 11
            retornoFuncoesConfiguracoes = "11 Dados inválidos passados pela automação. "
        Case 12
            retornoFuncoesConfiguracoes = "12 Modo seguro não ativo (possível falta de configuração no servidor SiTef do arquivo .cha). "
        Case 13
            retornoFuncoesConfiguracoes = "13 Caminho DLL inválido (o caminho completo das bibliotecas está muito grande)."
            
        Case Else
        retornoFuncoesConfiguracoes = "[ERRO] Código " + codigo + " desconhecido"
        
    End Select

End Function

Public Function retornoFuncoesTEF(codigo As String)


    Select Case codigo
        Case "1"
            retornoFuncoesTEF = "0 Sucesso na execução da função. "
        Case "10000"
            retornoFuncoesTEF = "10000 Deve ser chamada a rotina de continuidade do processo. "
        Case Is > 1
            retornoFuncoesTEF = "outro valor positivo Negada pelo autorizador. "
        Case "-1"
            retornoFuncoesTEF = "-1 Módulo não inicializado. O PDV tentou chamar alguma rotina sem antes executar a função configura. "
        Case "-2"
            retornoFuncoesTEF = "-2 Operação cancelada pelo operador. -3 O parâmetro função / modalidade é inexistente/inválido. "
        Case "-4"
            retornoFuncoesTEF = "-4 Falta de memória no PDV. -5 Sem comunicação com o SiTef. -6 Operação cancelada pelo usuário (no pinpad). "
        Case "-7"
            retornoFuncoesTEF = "-7 Reservado -8 A CliSiTef não possui a implementação da função necessária, provavelmente está desatualizada (a CliSiTefI é mais recente). "
        Case "-9"
            retornoFuncoesTEF = "-9 A automação chamou a rotina ContinuaFuncaoSiTefInterativo sem antes iniciar uma função iterativa. "
        Case "-10"
            retornoFuncoesTEF = "-10 Algum parâmetro obrigatório não foi passado pela automação comercial. "
        Case "-12"
            retornoFuncoesTEF = "-12 Erro na execução da rotina iterativa. Provavelmente o processo iterativo anterior não foi executado até o final (enquanto o retorno for igual a 10000). "
        Case "-13"
            retornoFuncoesTEF = "-13 Documento fiscal não encontrado nos registros da CliSiTef. Retornado em funções de consulta tais como ObtemQuantidadeTransaçõesPendentes. "
        Case "-15"
            retornoFuncoesTEF = "-15 Operação cancelada pela automação comercial. "
        Case "-20"
            retornoFuncoesTEF = "-20 Parâmetro inválido passado para a função. "
        Case "-21"
            retornoFuncoesTEF = "-21 Utilizada uma palavra proibida, por exemplo SENHA, para coletar dados em aberto no pinpad. Por exemplo na função ObtemDadoPinpadDiretoEx. "
        Case "-25"
            retornoFuncoesTEF = "-25 Erro no Correspondente Bancário: Deve realizar sangria. "
        Case "-30"
            retornoFuncoesTEF = "-30 Erro de acesso ao arquivo. Certifique-se que o usuário que roda a aplicação tem direitos de leitura/escrita. "
        Case "-40"
            retornoFuncoesTEF = "-40 Transação negada pelo servidor SiTef. "
        Case "-41"
            retornoFuncoesTEF = "-41 Dados inválidos. "
        Case "-42"
            retornoFuncoesTEF = "-42 Reservado"
        Case "-43"
            retornoFuncoesTEF = "-43 Problema na execução de alguma das rotinas no pinpad. "
        Case "-50"
            retornoFuncoesTEF = "-50 Transação não segura. "
        Case "-100"
            retornoFuncoesTEF = "-100 Erro interno do módulo. outro valor negativo Erros detectados internamente pela rotina."
        Case Else
            retornoFuncoesTEF = "[ERRO] Código " + codigo + " desconhecido"
    End Select

End Function


Public Sub exibirMensagemTEF(mensagem As String)

On Error GoTo TrataErro

    If Not GLB_TefHabilidado Then Exit Sub

    Screen.MousePointer = 11

    EscreveMensagemPermanentePinPad (mensagem)
                                     
TrataErro:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox "Erro no TEF " & Err.Number & vbNewLine & Err.Description, vbCritical, "TEF"
    End If
End Sub

Public Sub exibirMensagemPedidoTEF(numeroPedido As String, parcelas As Byte)
    
    Dim msgParcela As String
    
    msgParcela = " parcela"
    
    If parcelas > 1 Then msgParcela = msgParcela + "s"
        
    exibirMensagemTEF ("Pedido " & Trim(numeroPedido) & vbNewLine & _
                   "" & parcelas & msgParcela)
                   

End Sub

Public Sub ImprimeComprovanteTEF(ByRef mensagemComprovanteTEF As String)
    
    If mensagemComprovanteTEF = "" Then Exit Sub
    
    Screen.MousePointer = 11
    
    impressoraRelatorio "[INICIO]"
    impressoraRelatorio mensagemComprovanteTEF
    impressoraRelatorio "[FIM]"
 
    Screen.MousePointer = 0
    
    mensagemComprovanteTEF = ""
    
End Sub

Public Sub exibirMensagemPadraoTEF()

    exibirMensagemTEF "  Conectado ao " & vbNewLine & "   DMAC CAIXA"

End Sub

