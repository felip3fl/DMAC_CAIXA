Attribute VB_Name = "Modulo_SiTef"
Option Explicit


Public Declare Function ConfiguraIntSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
                   
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByRef pProximoComando As Long, ByRef pTipoCampo As Long, ByRef pTamanhoMinimo As Integer, ByRef pTamanhoMaximo As Integer, ByVal pBuffer As String, ByVal TamMaxBuffer As Long, ByVal ContinuaNavegacao As Long) As Long

Private Declare Function LeSimNaoPinPad Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As String) As Long
Private Declare Function EscreveMensagemPermanentePinPad Lib "C:\Sistemas\DMAC Caixa\Sitef\CliSitef32I.dll" (ByVal Funcao As String) As Long

Global Resultado     As Long
Global ComprovantePagamento As String
Global GLB_TefHabilidado As Boolean
Private Const endereco = ""



Public Sub criaLogTef(Nome As String, Mensagem As String)

    Open "C:\Sistemas\DMAC Caixa\Sitef\log\log" & Nome & ".txt" For Output As #1
            
        Print #1, Mensagem
    
    Close #1
    
End Sub


Public Function retornoFuncoesConfiguracoes(codigo As String)

    endereco = App.Path

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
            retornoFuncoesTEF = "1 outro valor positivo Negada pelo autorizador. "
        Case "-1"
            retornoFuncoesTEF = "-1 M�dulo n�o inicializado. O PDV tentou chamar alguma rotina sem antes executar a fun��o configura. "
        Case "-2"
            retornoFuncoesTEF = "-2 Opera��o cancelada pelo operador. "
        Case "-3"
            retornoFuncoesTEF = "-3 O par�metro fun��o / modalidade � inexistente/inv�lido. "
        Case "-4"
            retornoFuncoesTEF = "-4 Falta de mem�ria no PDV."
        Case "-5"
            retornoFuncoesTEF = "-5 Sem comunica��o com o SiTef. "
        Case "-6"
            retornoFuncoesTEF = "-6 Opera��o cancelada pelo usu�rio (no pinpad). "
        Case "-7"
            retornoFuncoesTEF = "-7 Reservado"
        Case "-7"
            retornoFuncoesTEF = "-8 A CliSiTef n�o possui a implementa��o da fun��o necess�ria, provavelmente est� desatualizada (a CliSiTefI � mais recente). "
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


Public Sub exibirMensagemTEF(Mensagem As String)

On Error GoTo TrataErro

    If Not GLB_TefHabilidado Then Exit Sub

    Screen.MousePointer = 11

    EscreveMensagemPermanentePinPad (Mensagem)
                                     
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
    
    exibirMensagemTEF " Imprimindo TEF" & vbNewLine & "   Aguarde..."
    
    impressoraRelatorio "[INICIO]"
    impressoraRelatorio mensagemComprovanteTEF
    impressoraRelatorio "[FIM]"
 
    Screen.MousePointer = 0
    
    mensagemComprovanteTEF = ""
    
    Esperar 1
    
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
                                  ByRef Nf As notaFiscal, _
                                  ByRef campoExibirMensagem As Label) _
                                  As Boolean
                                  

    Dim Retorno        As Long
    Dim Buffer         As String * 20000
    Dim BufferAux      As String * 20000
    Dim resposta       As String
    Dim ProximoComando As Long
    Dim TipoCampo      As Long
    Dim TamanhoMinimo  As Integer
    Dim TamanhoMaximo  As Integer
    Dim ContinuaNavegacao  As Long

    Dim logOperacoesTEF As String
    Dim horaOperacao As String
    Dim dataOperacao As String
    
    Dim valores As String
    Dim tipoOperacao As String
    
    If codigoOperacao = "" Then
        MsgBox "N�o foi poss�vel realizar a opera��o. C�digo de Opera��o n�o informado", vbExclamation, "TEF"
        Exit Function
    End If
    
    horaOperacao = Format(Time, "HHMMSS")
    dataOperacao = Format(GLB_DataInicial, "YYYYMMDD")
    
    Nf.valor = Format(Nf.valor, "###,###,##0.00")
    Nf.dataEmissao = Format(Nf.dataEmissao, "DDMMYYYY")
    
    Screen.MousePointer = 11
    
    Retorno = IniciaFuncaoSiTefInterativo(codigoOperacao, _
                                          Nf.valor & Chr(0), _
                                          Nf.pedido & Chr(0), _
                                          dataOperacao & Chr(0), _
                                          horaOperacao & Chr(0), _
                                          Trim(GLB_USU_Nome) & Chr(0), _
                                          Chr(0))
                                        
    
    ProximoComando = 0
    TipoCampo = 0
    TamanhoMinimo = 0
    TamanhoMaximo = 0
    ContinuaNavegacao = 0
    Resultado = 0
    Buffer = String(20000, 0)
    campoExibirMensagem.Caption = ""
    
    Do
        
            If valores <> "" Then
                Retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, _
                                                        TipoCampo, _
                                                        TamanhoMinimo, _
                                                        TamanhoMaximo, _
                                                        valores, _
                                                        Len(valores), _
                                                        Resultado)
            Else
                Retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, _
                                                TipoCampo, _
                                                TamanhoMinimo, _
                                                TamanhoMaximo, _
                                                Buffer, _
                                                Len(Buffer), _
                                                Resultado)
            End If
        

                                                
        valores = ""
                                                
        BufferAux = Buffer
        logOperacoesTEF = logOperacoesTEF & "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                              "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                              "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] " & left(Buffer, 200) & vbNewLine
        
        If (Retorno = 10000) Then
        
            If ProximoComando > 0 And ProximoComando < 4 Then
                campoExibirMensagem.Caption = UCase(Trim(Buffer))
                campoExibirMensagem.Refresh
            End If
        
            Select Case TipoCampo
            Case -1
            
                If ProximoComando = 21 Then
                
                    valores = "1"
                
                    If Buffer Like "1:A Vista;2:Parcelado pelo Estabelecimento;3:Parcelado pela Administradora;*" And Nf.parcelas > 1 Then
                        valores = Nf.parcelas
                    End If
                        
                    If Buffer Like "Forneca o numero de parcela*" Then
                        valores = Format(Nf.parcelas, "0")
                    End If
                    
                    'If Buffer Like "1:Magnetico/Chip;2:Digitado*" Then
                     '   valores = "1"
                        'valores = InputBox(Buffer)
                    'End If
                    
                ElseIf ProximoComando = 22 Then
                    'MsgBox Trim(Buffer)
                End If
            
                    
            Case 121, 122 'Buffer cont�m a primeira via do comprovante de pagamento
                    ComprovantePagamento = ComprovantePagamento & Buffer & vbNewLine
                    If Nf.numeroTEF = "" Then Nf.numeroTEF = lerCamporResultadoTEF(ComprovantePagamento, "Host")
            Case 515 'Data da transa��o a ser cancelada (DDMMAAAA) ou a ser re-impressa
                    valores = Nf.dataEmissao
            Case 516 'N�mero do documento a ser cancelado ou a ser re-impresso
                    valores = Nf.numeroTEF 'numero tef
            Case 146 'A rotina est� sendo chamada para ler o Valor a ser cancelado.
                    valores = Nf.valor
            Case 512, 513, 514
                    valores = InputBox(Buffer)
            Case 21
            
            Case 4
                    
                    
            End Select
            

            
            logOperacoesTEF = logOperacoesTEF & "[Coma:" & Space(4 - Len(Trim(ProximoComando))) & ProximoComando & "]" & _
                              "[Resu:" & Space(4 - Len(Trim(Resultado))) & Resultado & "]" & _
                              "[Tipo:" & Space(4 - Len(Trim(TipoCampo))) & TipoCampo & "] VALORES: " & left(valores, 200) & vbNewLine
        
        End If
    
    Loop Until Not (Retorno = 10000)
    
    If (Retorno = 0) Then
        campoExibirMensagem.Caption = "Opera��o completada com sucesso"
        EfetuaOperacaoTEF = True
    Else
        MsgBox "Erro " & "" & retornoFuncoesTEF(Str(Retorno)), vbCritical, "Erro TEF"
    End If
    
    FinalizaTransacaoSiTefInterativo 1, _
                                     Nf.pedido, _
                                     dataOperacao, _
                                     horaOperacao
                                     
                                     
    Screen.MousePointer = 0
    
    If codigoOperacao < 100 Then
        tipoOperacao = "Venda" & pedido
    Else
        tipoOperacao = "Cancelamento" & pedido
    End If
    
    criaLogTef tipoOperacao, logOperacoesTEF
    
 
End Function


Private Sub carregarDadosTEFBancoDeDados(ByRef labelMensagem As Label, ByRef Ip As String, ByRef IdTerminal As String, ByRef IdLoja As String, ByRef HabilitaTEF As Boolean)

    Dim RsDados As New ADODB.Recordset
    Dim sql As String

    On Error GoTo TrataErro

    sql = "select " & vbNewLine & _
          "LT_TefHabilidado Habilitado, " & vbNewLine & _
          "LT_IPSiTef IPSiTef, " & vbNewLine & _
          "LT_IdLoja IdLoja, " & vbNewLine & _
          "LT_IdTerminal IdTerminal, " & vbNewLine & _
          "LT_Reservado Reservado" & vbNewLine & _
          "from lojaTEF " & vbNewLine & _
          "where LT_LOJA = '" & wLoja & "'" & vbNewLine & _
          "and LT_NumeroCaixa = '" & GLB_Caixa & "'"

    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not RsDados.EOF Then
        If RsDados("Habilitado") = "S" Then GLB_TefHabilidado = True
        Ip = RsDados("IPSiTef")
        
        IdLoja = RsDados("IdLoja")
        IdTerminal = RsDados("IdTerminal")
    End If

    RsDados.Close

TrataErro:

    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        labelMensagem.Caption = "Erro buscar dados do TEF no banco de dados" & vbNewLine & Err.Number & " - " & Err.Description
        frmControlaCaixa.cmdNroCaixa.ForeOver = vbRed
        frmControlaCaixa.cmdNroCaixa.ForeColor = vbRed
        frmControlaCaixa.cmdNroCaixa.Caption = Replace(frmControlaCaixa.cmdNroCaixa.Caption, "(TEF)", "(Erro no TEF)")
    End If

End Sub

Public Sub conectarTEF(ByRef labelMensagem As Label)

  Dim Retorno As Long
  Dim Ip As String
  Dim IdTerminal As String
  Dim IdLoja As String
  
  On Error GoTo TrataErro

  labelMensagem.Caption = ""
  Screen.MousePointer = 11

  carregarDadosTEFBancoDeDados labelMensagem, Ip, IdTerminal, IdLoja, GLB_TefHabilidado

  If Not GLB_TefHabilidado Then
    Exit Sub
  End If
  
  If right(IdTerminal, 3) >= 900 And right(IdTerminal, 3) <= 999 Then
     MsgBox "Aten��o: A automa��o comercial n�o deve utilizar a identifica��o de terminal na faixa entre 000900 a 000999 que � reservada para uso pelo SiTef: Fun��o ConfiguraIntSiTefInterativo (EndSiTef, IdLoja, IdTerminal, Reservado);", vbCritical, "Inicializa��o TEF"
  End If
        
  Retorno = ConfiguraIntSiTefInterativo(Ip & Chr(0), IdLoja & Chr(0), IdTerminal & Chr(0), 0)

  If (Retorno = 0) Then
    labelMensagem.Caption = "Conex�o com o sistema SITEF realizada com sucesso"
  Else
    labelMensagem.Caption = "TEF Erro: Retorno -> " & CStr(Retorno)
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

