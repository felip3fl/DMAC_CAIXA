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
            retornoFuncoesConfiguracoes = "0 N�o ocorreu erro "
        Case 1
            retornoFuncoesConfiguracoes = "1 Endere�o IP inv�lido ou n�o resolvido "
        Case 2
            retornoFuncoesConfiguracoes = "2 C�digo da loja inv�lido "
        Case 3
            retornoFuncoesConfiguracoes = "3 C�digo de terminal inv�lido "
        Case 6
            retornoFuncoesConfiguracoes = "6 Erro na inicializa��o do TcpretornoFuncoesConfiguracoes= /Ip "
        Case 7
            retornoFuncoesConfiguracoes = "7 Falta de mem�ria "
        Case 8
            retornoFuncoesConfiguracoes = "8 N�o encontrou a CliSiTef ou ela est� com problemas "
        Case 9
            retornoFuncoesConfiguracoes = "9 Configura��o de servidores SiTef foi excedida. "
        Case 10
            retornoFuncoesConfiguracoes = "10 Erro de acesso na pasta CliSiTef (poss�vel falta de permiss�o para escrita) "
        Case 11
            retornoFuncoesConfiguracoes = "11 Dados inv�lidos passados pela automa��o. "
        Case 12
            retornoFuncoesConfiguracoes = "12 Modo seguro n�o ativo (poss�vel falta de configura��o no servidor SiTef do arquivo .cha). "
        Case 13
            retornoFuncoesConfiguracoes = "13 Caminho DLL inv�lido (o caminho completo das bibliotecas est� muito grande)."
            
        Case Else
        retornoFuncoesConfiguracoes = "[ERRO] C�digo " + codigo + " desconhecido"
        
    End Select

End Function

Public Function retornoFuncoesTEF(codigo As String)


    Select Case codigo
        Case "1"
            retornoFuncoesTEF = "0 Sucesso na execu��o da fun��o. "
        Case "10000"
            retornoFuncoesTEF = "10000 Deve ser chamada a rotina de continuidade do processo. "
        Case Is > 1
            retornoFuncoesTEF = "outro valor positivo Negada pelo autorizador. "
        Case "-1"
            retornoFuncoesTEF = "-1 M�dulo n�o inicializado. O PDV tentou chamar alguma rotina sem antes executar a fun��o configura. "
        Case "-2"
            retornoFuncoesTEF = "-2 Opera��o cancelada pelo operador. -3 O par�metro fun��o / modalidade � inexistente/inv�lido. "
        Case "-4"
            retornoFuncoesTEF = "-4 Falta de mem�ria no PDV. -5 Sem comunica��o com o SiTef. -6 Opera��o cancelada pelo usu�rio (no pinpad). "
        Case "-7"
            retornoFuncoesTEF = "-7 Reservado -8 A CliSiTef n�o possui a implementa��o da fun��o necess�ria, provavelmente est� desatualizada (a CliSiTefI � mais recente). "
        Case "-9"
            retornoFuncoesTEF = "-9 A automa��o chamou a rotina ContinuaFuncaoSiTefInterativo sem antes iniciar uma fun��o iterativa. "
        Case "-10"
            retornoFuncoesTEF = "-10 Algum par�metro obrigat�rio n�o foi passado pela automa��o comercial. "
        Case "-12"
            retornoFuncoesTEF = "-12 Erro na execu��o da rotina iterativa. Provavelmente o processo iterativo anterior n�o foi executado at� o final (enquanto o retorno for igual a 10000). "
        Case "-13"
            retornoFuncoesTEF = "-13 Documento fiscal n�o encontrado nos registros da CliSiTef. Retornado em fun��es de consulta tais como ObtemQuantidadeTransa��esPendentes. "
        Case "-15"
            retornoFuncoesTEF = "-15 Opera��o cancelada pela automa��o comercial. "
        Case "-20"
            retornoFuncoesTEF = "-20 Par�metro inv�lido passado para a fun��o. "
        Case "-21"
            retornoFuncoesTEF = "-21 Utilizada uma palavra proibida, por exemplo SENHA, para coletar dados em aberto no pinpad. Por exemplo na fun��o ObtemDadoPinpadDiretoEx. "
        Case "-25"
            retornoFuncoesTEF = "-25 Erro no Correspondente Banc�rio: Deve realizar sangria. "
        Case "-30"
            retornoFuncoesTEF = "-30 Erro de acesso ao arquivo. Certifique-se que o usu�rio que roda a aplica��o tem direitos de leitura/escrita. "
        Case "-40"
            retornoFuncoesTEF = "-40 Transa��o negada pelo servidor SiTef. "
        Case "-41"
            retornoFuncoesTEF = "-41 Dados inv�lidos. "
        Case "-42"
            retornoFuncoesTEF = "-42 Reservado"
        Case "-43"
            retornoFuncoesTEF = "-43 Problema na execu��o de alguma das rotinas no pinpad. "
        Case "-50"
            retornoFuncoesTEF = "-50 Transa��o n�o segura. "
        Case "-100"
            retornoFuncoesTEF = "-100 Erro interno do m�dulo. outro valor negativo Erros detectados internamente pela rotina."
        Case Else
            retornoFuncoesTEF = "[ERRO] C�digo " + codigo + " desconhecido"
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

