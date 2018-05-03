Attribute VB_Name = "Modulo_SiTef"
Option Explicit


Public Declare Function ConfiguraIntSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function ConfiguraIntSiTefInterativoEx Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer, ByVal pParamAdic As String) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)

Public Declare Function ObtemQuantidadeTransacoesPendentes Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal DataFiscal As String, ByVal cupomFiscal As String) As Long

                   
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" _
(ByRef pProximoComando As Long, _
ByRef pTipoCampo As Long, _
ByRef pTamanhoMinimo As Integer, _
ByRef pTamanhoMaximo As Integer, _
ByVal pBuffer As String, _
ByVal TamMaxBuffer As Long, _
ByVal ContinuaNavegacao As Long) As Long

Private Declare Function LeSimNaoPinPad Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As String) As Long
Private Declare Function EscreveMensagemPermanentePinPad Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As String) As Long

Global Resultado     As Long
Global comprovantePagamento As String
Global ComprovantePagamentoFila As String
Global GLB_TefHabilidado As Boolean
Global GLB_HabilidadoCieloTEF As Boolean
Private Const endereco = ""
Private filaCuponsTEF(1 To 20) As String
Private posicaoFila As Integer

Global operacaoTEFCompleta As Boolean

Public Const GLB_ENDERECOCOMPROVANTETEF = "C:\Sistemas\DMAC Caixa\Sitef\Comprovantes\"

Public Type notaFiscalTEF
    numero As String
    loja As String
    serie As String
    pedido As String
    cfop As String
    valor As String
    numeroTEF As String
    Parcelas As Integer
    dataEmissao As String
    codigoBandeiraCartao As String
    comprovantePagamento As String
    sequenciaMovimentoCaixa As String
End Type

'Public Sub criaLogTef(numeropedido As String, sequenciaTEF As String, via As String, mensagem As String)
'
'    Open GLB_ENDERECOCOMPROVANTETEF & numeropedido & "-" & sequenciaTEF & "-" & Nome & ".txt" For Output As #1
'
'        Print #1, mensagem
'
'    Close #1
'
'End Sub

'Private Sub imprimirComprovantesTEF(numeropedido As String)
'
'    Dim mensagemArquivoTXT As TextStream
'    Dim fso As New FileSystemObject
'    Dim informacaoArquivo As String
'
'    numeropedido = "2611"
'
'    Arquivo = Dir(GLB_ENDERECOCOMPROVANTETEF & numeropedido & "*" & ".txt", vbDirectory)
'
'    If Arquivo <> "" Then
'
'         Set mensagemArquivoTXT = fso.OpenTextFile(GLB_ENDERECOCOMPROVANTETEF & Arquivo)
'         informacaoArquivo = mensagemArquivoTXT.ReadAll
'         mensagemArquivoTXT.Close
'
'
'
'    End If
'
'End Sub

Public Sub CriaComprovanteTef(numeropedido As String, sequenciaTEF As String, via As String, comprovante As String)

    Open GLB_ENDERECOCOMPROVANTETEF & numeropedido & "-" & sequenciaTEF & "-" & via & ".txt" For Output As #1
        Print #1, comprovante
    Close #1
    
End Sub



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
            retornoFuncoesTEF = "[0] Sucesso na execu��o da fun��o. "
        Case "10000"
            retornoFuncoesTEF = "[10000] Deve ser chamada a rotina de continuidade do processo. "
        Case Is > 1
            retornoFuncoesTEF = "[1] outro valor positivo Negada pelo autorizador. "
        Case "-1"
            retornoFuncoesTEF = "[-1] M�dulo n�o inicializado. O PDV tentou chamar alguma rotina sem antes executar a fun��o configura. "
        Case "-2"
            retornoFuncoesTEF = "[-2] Opera��o cancelada pelo operador. "
        Case "-3"
            retornoFuncoesTEF = "[-3] O par�metro fun��o / modalidade � inexistente/inv�lido. "
        Case "-4"
            retornoFuncoesTEF = "[-4] Falta de mem�ria no PDV."
        Case "-5"
            retornoFuncoesTEF = "[-5] Sem comunica��o com o SiTef. "
        Case "-6"
            retornoFuncoesTEF = "[-6] Opera��o cancelada pelo usu�rio (no pinpad). "
        Case "-7"
            retornoFuncoesTEF = "[-7] Reservado"
        Case "-7"
            retornoFuncoesTEF = "[-8] A CliSiTef n�o possui a implementa��o da fun��o necess�ria, provavelmente est� desatualizada (a CliSiTefI � mais recente). "
        Case "-9"
            retornoFuncoesTEF = "[-9] A automa��o chamou a rotina ContinuaFuncaoSiTefInterativo sem antes iniciar uma fun��o iterativa. "
        Case "-10"
            retornoFuncoesTEF = "[-10] Algum par�metro obrigat�rio n�o foi passado pela automa��o comercial. "
        Case "-12"
            retornoFuncoesTEF = "[-12] Erro na execu��o da rotina iterativa. Provavelmente o processo iterativo anterior n�o foi executado at� o final (enquanto o retorno for igual a 10000). "
        Case "-13"
            retornoFuncoesTEF = "[-13] Documento fiscal n�o encontrado nos registros da CliSiTef. Retornado em fun��es de consulta tais como ObtemQuantidadeTransa��esPendentes. "
        Case "-15"
            retornoFuncoesTEF = "[-15] Opera��o cancelada pela automa��o comercial. "
        Case "-20"
            retornoFuncoesTEF = "[-20] Par�metro inv�lido passado para a fun��o. "
        Case "-21"
            retornoFuncoesTEF = "[-21] Utilizada uma palavra proibida, por exemplo SENHA, para coletar dados em aberto no pinpad. Por exemplo na fun��o ObtemDadoPinpadDiretoEx. "
        Case "-25"
            retornoFuncoesTEF = "[-25] Erro no Correspondente Banc�rio: Deve realizar sangria. "
        Case "-30"
            retornoFuncoesTEF = "[-30] Erro de acesso ao arquivo. Certifique-se que o usu�rio que roda a aplica��o tem direitos de leitura/escrita. "
        Case "-40"
            retornoFuncoesTEF = "[-40] Transa��o negada pelo servidor SiTef. "
        Case "-41"
            retornoFuncoesTEF = "[-41] Dados inv�lidos. "
        Case "-42"
            retornoFuncoesTEF = "[-42] Reservado"
        Case "-43"
            retornoFuncoesTEF = "[-43] Problema na execu��o de alguma das rotinas no pinpad. "
        Case "-50"
            retornoFuncoesTEF = "[-50] Transa��o n�o segura. "
        Case "-100"
            retornoFuncoesTEF = "[-100] Erro interno do m�dulo. outro valor negativo Erros detectados internamente pela rotina."
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

Public Sub exibirMensagemPedidoTEF(numeropedido As String, Parcelas As Byte)
    
    Dim msgParcela As String
    
    msgParcela = " parcela"
    
    If Parcelas > 1 Then msgParcela = msgParcela + "s"
        
    exibirMensagemTEF ("Pedido " & Trim(numeropedido) & vbNewLine & _
                   "" & Parcelas & msgParcela)
                   

End Sub

Public Sub ImprimeComprovanteTEF(ByRef pedido As String)
    
    Dim i As Integer
    Dim Arquivo As String
    
    Screen.MousePointer = 11
    
    exibirMensagemTEF " Imprimindo TEF" & vbNewLine & "   Aguarde..."
    
    Dim mensagemArquivoTXT As TextStream
    Dim fso As New FileSystemObject
    Dim informacaoArquivo As String


    Arquivo = Dir(GLB_ENDERECOCOMPROVANTETEF & pedido & "*" & ".txt", vbDirectory)
    
    Do While Arquivo <> ""
    
         Set mensagemArquivoTXT = fso.OpenTextFile(GLB_ENDERECOCOMPROVANTETEF & Arquivo)
         informacaoArquivo = mensagemArquivoTXT.ReadAll
         mensagemArquivoTXT.Close
    
        impressoraRelatorio "[INICIO]"
        impressoraRelatorio informacaoArquivo
        impressoraRelatorio "[FIM]"
        
    
         Arquivo = Dir
    
    Loop
 
    Screen.MousePointer = 0

    
End Sub

Public Sub adicionarFilaComprovanteTEF(mensagem As String)

    If posicaoFila <= 0 Then posicaoFila = 1

    filaCuponsTEF(posicaoFila) = mensagem

    posicaoFila = posicaoFila + 1
    
End Sub

Public Sub exibirMensagemPadraoTEF()

    exibirMensagemTEF "  Conectado ao " & vbNewLine & "   DMAC CAIXA"

End Sub

Public Function lerCamporResultadoTEF(informacoes As String, campo As String) As String
    
    If informacoes Like "*" & campo & ": *" Then
        Dim inicioCampo, fimCampo As Integer
    
        inicioCampo = (InStr(informacoes, campo & ": ")) + (Len(campo)) + 2
        fimCampo = (InStr(inicioCampo, informacoes, Chr(10))) - inicioCampo
    
        If inicioCampo + fimCampo <> 0 Then
            lerCamporResultadoTEF = Mid$(informacoes, inicioCampo, fimCampo)
        End If
    Else
        lerCamporResultadoTEF = ""
    End If
    
End Function


Public Function EfetuaOperacaoTEF(ByVal codigoOperacao As String, _
                                  ByRef nf As notaFiscalTEF, _
                                  ByRef bandeiraCartao As Label, _
                                  ByRef campoExibirMensagem As Label) _
                                  As Boolean
                                  

    Dim retorno        As Long
    Dim Buffer         As String * 20000
    Dim resposta       As String
    Dim ProximoComando As Long
    Dim TipoCampo      As Long
    Dim TamanhoMinimo  As Integer
    Dim tamanhoMaximo  As Integer
    Dim ContinuaNavegacao  As Long

    Dim logOperacoesTEF As String
    Dim dataOperacao As String
    Dim horaOperacaoFormatado As String
    
    Dim valores As String
    Dim tipoOperacao As String
    
    If codigoOperacao = "" Then
        MsgBox "N�o foi poss�vel realizar a opera��o. C�digo de Opera��o n�o informado", vbExclamation, "TEF"
        Exit Function
    End If
    
    horaOperacaoTEF = Time
    horaOperacaoFormatado = Format(horaOperacaoTEF, "HHMMSS")
    dataOperacao = Format(GLB_DataInicial, "YYYYMMDD")
    
    nf.valor = Format(nf.valor, "###,###,##0.00")
    nf.dataEmissao = Format(nf.dataEmissao, "DDMMYYYY")
    
    Screen.MousePointer = 11
    
    retorno = IniciaFuncaoSiTefInterativo(codigoOperacao, _
                                          nf.valor & Chr(0), _
                                          nf.pedido & Chr(0), _
                                          dataOperacao & Chr(0), _
                                          horaOperacaoFormatado & Chr(0), _
                                          Trim(GLB_USU_Nome) & Chr(0), _
                                          Chr(0))
                                        
    
    ProximoComando = 0
    TipoCampo = 0
    TamanhoMinimo = 0
    tamanhoMaximo = 0
    ContinuaNavegacao = 0
    Resultado = 0
    Buffer = String(20000, 0)
    campoExibirMensagem.Caption = ""
    
    Do
            DoEvents
            Resultado = 0
            
            If retornaOperacaoTEF Then
                retornaOperacaoTEF = False
                Resultado = 1
            End If
            
            If cancelarOperacaoTEF Then
                cancelarOperacaoTEF = False
                Resultado = 2
            End If
            
            If valores <> "" Then
                retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, _
                                                        TipoCampo, _
                                                        TamanhoMinimo, _
                                                        tamanhoMaximo, _
                                                        valores, _
                                                        Len(valores), _
                                                        Resultado)
            Else
                retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, _
                                                TipoCampo, _
                                                TamanhoMinimo, _
                                                tamanhoMaximo, _
                                                Buffer, _
                                                Len(Buffer), Resultado)
            End If
        

                                                
        valores = ""
                                                
        logOperacoesTEF = logOperacoesTEF & "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                              "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                              "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] " & left(Buffer, 200) & vbNewLine
                              
        Debug.Print "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                    "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                    "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] " & left(Buffer, 200) & vbNewLine
        
        If (retorno = 10000) Then
        
            If ProximoComando > 0 And ProximoComando < 4 Then
                campoExibirMensagem.Caption = UCase(Trim(Buffer))
                campoExibirMensagem.Refresh
            End If
        
            'If Not GLB_Administrador Then

                
                Select Case TipoCampo
                Case -1
                
                    Select Case ProximoComando
                    Case 21, 30
                        
                        If GLB_Administrador Then
                            valores = entradaDeValores("ProximoComando[" & ProximoComando & "]", Buffer, TamanhoMinimo, tamanhoMaximo, False)
                        Else
                            valores = "1"
                            If Buffer Like "1:A Vista*" Then
                                If nf.Parcelas > 1 Then valores = "2"
                            ElseIf Buffer Like "1:Magnetico/Chip;2:Digitado*" Then
                                valores = "1"
                            ElseIf Buffer Like "1:Rede;2:Cielo;3:Outros;*" Then
                                valores = "1"
                                If GLB_HabilidadoCieloTEF Then valores = "2"
                            Else
                                valores = entradaDeValores("ProximoComando = " & TipoCampo, Buffer, TamanhoMinimo, tamanhoMaximo, False)
                            End If
                        End If
                    Case 20
                        valores = entradaDeValores("TipoCampo = " & TipoCampo, "0:Sim;1:Nao;" & Buffer, TamanhoMinimo, tamanhoMaximo, False)
                    Case 22
                        
                         'Call frmExibicaoDeMensagem.exibirMensagem("Mensagem do TEF (ProximoComando[" & ProximoComando & "])", Trim(Buffer))
                         'frmExibicaoDeMensagem.Show 1
                         MsgBox Trim(Buffer), vbInformation, "Mensagem do TEF (ProximoComando[" & ProximoComando & "])"
                        
                    Case 132
                        bandeiraCartao.Caption = Mid(Buffer, 1, 5)
                        bandeiraCartao.Caption = obterTipoPagamentoCreditoTEF(bandeiraCartao.Caption)
                    End Select
                
                        
                Case 121 'Buffer cont�m a primeira via do comprovante de pagamento
                        adicionarFilaComprovanteTEF Buffer
                        nf.comprovantePagamento = Buffer
                        CriaComprovanteTef nf.pedido, nf.numeroTEF, "cliente", nf.comprovantePagamento
                Case 122
                        adicionarFilaComprovanteTEF Buffer
                        nf.comprovantePagamento = Buffer
                        CriaComprovanteTef nf.pedido, nf.numeroTEF, "operador", nf.comprovantePagamento
                Case 132
                        bandeiraCartao.Caption = Mid(Buffer, 1, 5)
                        bandeiraCartao.Caption = obterTipoPagamentoCreditoTEF(bandeiraCartao.Caption)
                Case 515
                        valores = nf.dataEmissao
                        If GLB_Administrador Then valores = entradaDeValores("TipoCampo = " & TipoCampo, "Data da transacao (DDMMAAAA)", TamanhoMinimo, tamanhoMaximo, False)
                Case 516
                        'valores = nf.numeroTEF 'numero tef
                        valores = formataCampoTEF(nf.numeroTEF)
                        If GLB_Administrador Then valores = entradaDeValores("TipoCampo = " & TipoCampo, "Forneca o numero do documento", TamanhoMinimo, tamanhoMaximo, False)
                Case 146
                        If ProximoComando = 34 Then
                            valores = nf.valor
                            If GLB_Administrador Then valores = entradaDeValores("TipoCampo = " & TipoCampo, "Forneca o valor da transacao a ser cancelada", TamanhoMinimo, tamanhoMaximo, False)
                        End If
                Case 512, 513
                        'valores = InputBox(Buffer)
                        valores = entradaDeValores("TipoCampo = " & TipoCampo, Buffer, TamanhoMinimo, tamanhoMaximo, False)
                Case 514
                        valores = entradaDeValores("TipoCampo = " & TipoCampo, "C�digo de ok,  do Cart�o", TamanhoMinimo, tamanhoMaximo, False)
                Case 500
                        valores = GLB_USU_Codigo
                        If GLB_Administrador Then
                            valores = entradaDeValores("TipoCampo = " & TipoCampo, "Forneca o codigo do supervisor", TamanhoMinimo, tamanhoMaximo, True)
                        End If
                Case 5013
                        If MsgBox("Deseja cancelar essa opera��o?", vbQuestion + vbYesNo, "TipoCampo = " & TipoCampo) = vbYes Then
                            valores = "-2"
                        End If
                Case 505
                        valores = Format(nf.Parcelas, "0")
                        If GLB_Administrador Then valores = entradaDeValores("TipoCampo = " & TipoCampo, Buffer, TamanhoMinimo, tamanhoMaximo, False)
                Case 952
                        nf.numeroTEF = Val(Mid(Buffer, 1, 10))
                        atualizaSequenciaTEF nf.sequenciaMovimentoCaixa, nf.numeroTEF
                Case 1190
                        valores = entradaDeValores("TipoCampo = " & TipoCampo, Buffer, TamanhoMinimo, tamanhoMaximo, False)
                    
                'Case Else
                        'If GLB_Administrador Then valores = InputBox(Trim(Buffer), "TipoCampo = " & TipoCampo)
                        
                        
                End Select
            
           ' Else
            
             '   Select Case ProximoComando
                'Case 21, 30
                    'valores = InputBox(Trim(Buffer), "ProximoComando = " & ProximoComando)
                
             '   Case 22, 4
             '       MsgBox Trim(Buffer), vbInformation, "ProximoComando = " & ProximoComando
              '  Case 132
              '          bandeiraCartao.Caption = Mid(Buffer, 1, 5)
             '           bandeiraCartao.Caption = obterTipoPagamentoCreditoTEF(bandeiraCartao.Caption)
               ' End Select
                
                'If TipoCampo > 1 And left(Buffer, 1) = " " Then
                  '  valores = InputBox(Trim(Buffer), "TipoCampo = " & TipoCampo)
                'End If
                
            'End If
            
            If valores <> "" Then
                logOperacoesTEF = logOperacoesTEF & "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                              "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                              "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] VALORES: " & left(valores, 200) & vbNewLine
                
                Debug.Print "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                            "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                            "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] VALORES: " & left(valores, 200) & vbNewLine
                
                If valores < "0" Then retorno = -2
                
            End If
        
        End If
    
    Loop Until Not (retorno = 10000)
    
    
    If (retorno = 0) Then
        campoExibirMensagem.Caption = "Opera��o completada com sucesso"
        campoExibirMensagem.Refresh
        'nf.numeroTEF = lerCamporResultadoTEF(nf.comprovantePagamento, "Host")
        EfetuaOperacaoTEF = True
        operacaoTEFCompleta = True
    Else
        MsgBox retornoFuncoesTEF(Str(retorno)), vbCritical, "Erro TEF"
        campoExibirMensagem.Caption = retornoFuncoesTEF(Str(retorno))
        campoExibirMensagem.Refresh
        operacaoTEFCompleta = False
    End If
                                     
                                     
    Screen.MousePointer = 0
    
    If codigoOperacao < 100 Then
        tipoOperacao = "Venda" & pedido
    Else
        tipoOperacao = "Cancelamento" & pedido
    End If
    
    'criaLogTef tipoOperacao, logOperacoesTEF
    
 
End Function

Private Function formataCampoTEF(valores As String)
    formataCampoTEF = Format(valores, "000000")
    If Not GLB_HabilidadoCieloTEF Then
        If Len(valores) < 9 Then
            formataCampoTEF = "999" + formataCampoTEF
        End If
    End If
End Function

Private Function validadorBancoDeDados(numeropedido As String, serie As String, _
                                  ByRef dataTransicao As String, ByRef horaTransicao As String) As Boolean

    Dim RsDados As New ADODB.Recordset
    Dim sql As String

    On Error GoTo TrataErro

    sql = "select top 1" & vbNewLine & _
          "MC_DataBaixaAVR horaOperacao, " & vbNewLine & _
          "MC_Data data " & vbNewLine & _
          "from MovimentoCaixa " & vbNewLine & _
          "where MC_Pedido = '" & numeropedido & "'" & vbNewLine & _
          "and MC_SequenciaTEF > 0" & vbNewLine & _
          "and MC_serie = '" & serie & "'"
    
    Screen.MousePointer = 11
    
    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Screen.MousePointer = 0
    
    If Not RsDados.EOF Then
        dataTransicao = RsDados("data")
        horaTransicao = RsDados("horaOperacao")
        validadorBancoDeDados = True
    End If
    
    
    RsDados.Close

TrataErro:

    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao finalizar transacao no TEF" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "TEF"
        validadorBancoDeDados = False
    End If
        
End Function


Public Sub finalizarTransacaoTEF(numeropedido As String, serie As String, validaAntesDeFinalizar As Boolean)

    Dim dataTransicao As String
    Dim horaTransicao As String
    Dim realizarFinalizacao As Boolean
    
    dataTransicao = Date
    horaTransicao = horaOperacaoTEF
    realizarFinalizacao = True
    
    If validaAntesDeFinalizar Then
        realizarFinalizacao = validadorBancoDeDados(numeropedido, serie, dataTransicao, horaTransicao)
    End If

    If realizarFinalizacao Then
        
        FinalizaTransacaoSiTefInterativo 1, _
                                     numeropedido, _
                                     Format(dataTransicao, "YYYYMMDD"), _
                                     Format(horaTransicao, "HHMMSS")

    End If

End Sub

Private Sub carregarDadosTEFBancoDeDados(ByRef labelMensagem As Label, _
                                         ByRef Ip As String, _
                                         ByRef IdTerminal As String, _
                                         ByRef IdLoja As String, _
                                         ByRef HabilitaTEF As Boolean, _
                                         ByRef CNPJDesenvolvedor As String)

    Dim RsDados As New ADODB.Recordset
    Dim sql As String

    On Error GoTo TrataErro

    sql = "select " & vbNewLine & _
          "LT_TefHabilidado Habilitado, " & vbNewLine & _
          "LT_IPSiTef IPSiTef, " & vbNewLine & _
          "LT_IdLoja IdLoja, " & vbNewLine & _
          "LT_IdTerminal IdTerminal, " & vbNewLine & _
          "LT_HabilitaCielo HabilitaCielo, " & vbNewLine & _
          "'60872124000199' CNPJDesenvolvedor, " & vbNewLine & _
          "LT_Reservado Reservado" & vbNewLine & _
          "from lojaTEF " & vbNewLine & _
          "where LT_LOJA = '" & wLoja & "'" & vbNewLine & _
          "and LT_NumeroCaixa = '" & GLB_Caixa & "'"
    
    Screen.MousePointer = 11
    
    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    Screen.MousePointer = 0
    
    If Not RsDados.EOF Then
        If RsDados("Habilitado") = "S" Then GLB_TefHabilidado = True
        Ip = RsDados("IPSiTef")
        
        IdLoja = RsDados("IdLoja")
        IdTerminal = RsDados("IdTerminal")
        CNPJDesenvolvedor = RsDados("CNPJDesenvolvedor")
        
        If RsDados("HabilitaCielo") = "S" Then GLB_HabilidadoCieloTEF = True
        
    End If

    RsDados.Close

TrataErro:

    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        labelMensagem.Caption = "Erro buscar dados do TEF no banco de dados" & vbNewLine & Err.Number & " - " & Err.Description
    End If

End Sub

Public Sub conectarTEF(ByRef labelMensagem As Label)

  Dim retorno As Long
  Dim Ip As String
  Dim IdTerminal As String
  Dim IdLoja As String
  Dim CNPJDesenvolvedor As String
  
    Dim data As String
  Dim cupomFiscal As String
  
  On Error GoTo TrataErro

  labelMensagem.Caption = ""
  Screen.MousePointer = 11

  carregarDadosTEFBancoDeDados labelMensagem, Ip, IdTerminal, IdLoja, GLB_TefHabilidado, CNPJDesenvolvedor

  If Not GLB_TefHabilidado Then
    Exit Sub
  End If
  
  frmControlaCaixa.cmdNroCaixa.Caption = frmControlaCaixa.cmdNroCaixa.Caption & " (TEF)"
  
  If right(IdTerminal, 3) >= 900 And right(IdTerminal, 3) <= 999 Then
     MsgBox "Aten��o: A automa��o comercial n�o deve utilizar a identifica��o de terminal na faixa entre 000900 a 000999 que � reservada para uso pelo SiTef: Fun��o ConfiguraIntSiTefInterativo (EndSiTef, IdLoja, IdTerminal, Reservado);", vbCritical, "Inicializa��o TEF"
  End If
        
  retorno = ConfiguraIntSiTefInterativoEx(Ip & Chr(0), IdLoja & Chr(0), IdTerminal & Chr(0), 0, "[ParmsClient=1=" & wCGC & ";2=" & CNPJDesenvolvedor & "]")

  If (retorno = 0) Then
    labelMensagem.Caption = "Conex�o com o sistema SITEF realizada com sucesso"
  Else
    labelMensagem.Caption = "TEF Erro: Retorno -> " & CStr(retorno)
    MsgBox "Erro no envio do CNPJ (TEF). Verifique se o CNPJ " & wCGC & " est� valido.", vbCritical, "Valida��o CNPJ TEF (" & CStr(retorno) & ")"
  End If
  
  exibirMensagemPadraoTEF
  
TrataErro:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        labelMensagem.Caption = "Erro no TEF " & Err.Number & vbNewLine & Err.Description
        frmControlaCaixa.cmdNroCaixa.ForeOver = vbRed
        frmControlaCaixa.cmdNroCaixa.ForeColor = vbRed
        frmControlaCaixa.cmdNroCaixa.Caption = Replace(frmControlaCaixa.cmdNroCaixa.Caption, "(TEF)", "(Erro no TEF)")
    End If
  
End Sub


Private Function obterTipoPagamentoCreditoTEF(codigoBandeiraTEF As String)
    Select Case codigoBandeiraTEF
    Case "00000"
        obterTipoPagamentoCreditoTEF = "D�bito Gen�rico"
    Case "00001"
        obterTipoPagamentoCreditoTEF = "VISA"
    Case "00002"
        obterTipoPagamentoCreditoTEF = "MASTERCARD"
    Case "00003"
        obterTipoPagamentoCreditoTEF = "Diners"
    Case "00004"
        obterTipoPagamentoCreditoTEF = "American Express"
    Case "00005"
        obterTipoPagamentoCreditoTEF = "Sollo"
    Case "00006"
        obterTipoPagamentoCreditoTEF = "Sidecard (Redecard)"
    Case "00007"
        obterTipoPagamentoCreditoTEF = "Private Label (Redecard)"
    Case "00008"
        obterTipoPagamentoCreditoTEF = "REDESHOP"
    Case "00009"
        obterTipoPagamentoCreditoTEF = "P�o de A��car"
    Case "00010"
        obterTipoPagamentoCreditoTEF = "Fininvest (Visanet)"
    Case "00011"
        obterTipoPagamentoCreditoTEF = "JCB"
    Case "00012"
        obterTipoPagamentoCreditoTEF = "HIPERCARD"
    Case "00013"
        obterTipoPagamentoCreditoTEF = "Aura"
    Case "00014"
        obterTipoPagamentoCreditoTEF = "Losango"
    Case "00015"
        obterTipoPagamentoCreditoTEF = "Sorocred"
    Case "10014"
        obterTipoPagamentoCreditoTEF = "Discovery"
    Case "20002" 'DEBITO
        obterTipoPagamentoCreditoTEF = "VISA ELEC."
    Case "20001" 'Maestro 'DEBITO
        obterTipoPagamentoCreditoTEF = "REDESHOP"
    Case "20013", "20032", "00031"
        obterTipoPagamentoCreditoTEF = "Elo"
    Case Else
        obterTipoPagamentoCreditoTEF = "CART�O DESCONHECIDO"
    End Select
End Function

Public Sub limparArquivosImpressaoTEF()
    deletaArquivo GLB_ENDERECOCOMPROVANTETEF & "*.txt"
End Sub

