Attribute VB_Name = "modFuncoes"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declaraciones para 32 bits
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F

Dim lFlag As Boolean
Public iTransacao As Integer
Dim i As Integer
Dim response As Integer
Dim linhaArquivo As String
Public naoConfirmado As Boolean
'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    VerificaGerenciadorPadrao
'// Objetivo:
'//    Verificar se o Gerenciador Padr�o est� ativo
'// Par�metro:
'//    n�o h�
'// Retorno:
'//    True para Gerenciador Padr�o ATIVO
'//    False para Gerenciador Padr�o INATIVO
'//
'////////////////////////////////////////////////////////////////////////////////
Function VerificaGerenciadorPadrao() As Boolean
    Dim cConteudoArquivo As String
    Dim hora As Date
    Dim iTentativas As Integer
    
    hora = Date & " " & Time
    cConteudoArquivo = ""
    cConteudoArquivo = "000-000 = ATV" & vbCrLf & _
              "001-000 = " & hora & _
              "999-999 = 0"
    Call GravaArquivo_Binario(App.Path & "\INTPOS.001", cConteudoArquivo)
       
   ' Copia o arquivo para o diret�rio do Gerenciador Padr�o
    FileCopy App.Path & "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
    
    ' Apaga o arquivo local
    MataArquivo (App.Path & "\INTPOS.001")
   
    For iTentativas = 1 To 7 Step 1
        If Dir("C:\TEF_DIAL\RESP\ATIVO.001") = "" Or Dir("C:\TEF_DIAL\RESP\INTPOS.STS") = "" Then
            lFlag = True
            Sleep (1000)
            VerificaGerenciadorPadrao = True
            Exit Function
            
        End If
        If iTentativas = 7 Then
            lFlag = False
            VerificaGerenciadorPadrao = True
            Exit Function
        End If
    Next iTentativas

End Function
'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    RealizaTransacao
'// Objetivo:
'//    Realiza a transa��o TEF
'// Par�metros:
'//   TDateTime para identificar o n�mero da transa��o
'//   String para o N�mero do Cupom Fiscal (COO)
'//   String para a Valor da Forma de Pagamento
'//   Integer com o n�mero da transa��o
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function RealizaTransacao(hora As Date, cNumeroCupom As String, _
                           cValorPago As String, iConta As Integer) As Integer
    Dim cConteudoArquivo As String
    Dim cLinhaArquivo As String
    Dim cLinha As String
    Dim cCampoArquivo As String
    Dim iArquivo As Integer
    Dim arquivoIncorreto As Boolean
    Dim lFlag As Boolean
    Dim iTentativas As Integer
    Dim iVezes As Integer
    
    Dim bTransacao As Boolean
    Dim bFlagArq As Integer
    Dim lNumeroLinha As Long
    Dim iAux As Integer
   
    arquivoIncorreto = True
    
    '''''''''''''''CRIANDO A SOLICITA��O DA TRANSA��O TEF'''''''''''''''''
    ' Conte�do do arquivo INTPOS.001 para solicitar a transa��o TEF.
    cConteudoArquivo = ""
    cConteudoArquivo = "000-000 = CRT" & vbCrLf & _
                       "001-000 = " & Format(hora, "HhNnSs") & vbCrLf & _
                       "002-000 = " & cNumeroCupom & vbCrLf & _
                       "003-000 = " & cValorPago & vbCrLf & _
                       "999-999 = 0"
    Call GravaArquivo_Binario(App.Path & "\INTPOS.001", cConteudoArquivo)
    ' Copia o arquivo para o diret�rio do Gerenciador Padr�o
    FileCopy App.Path & "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
    ' Apaga o arquivo local
    MataArquivo (App.Path & "\INTPOS.001")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Se j� existe um IMPRIME[conta].TXT, deleta ele
    MataArquivo (App.Path & "\IMPRIME" & CStr(iConta) & ".TXT")
        
    RealizaTransacao = -2
    'Enquanto o gerenciador padr�o n�o tiver mandado resposta, fica em loop
    'Excedendo 7 segundos, sai da fun��o retornando 0
    While Dir("C:\TEF_DIAL\RESP\INTPOS.STS") = ""  ' Verifica o arquivo INTPOS.001 de resposta.
        DoEvents
        Sleep (1000)
        iTentativas = iTentativas + 1
        If iTentativas > 7 Then
            frmTEFVariosCartoes.lblMsg.Caption = "Gerenciador Padr�o n�o est� ativo!"
            RealizaTransacao = 0
            Exit Function
        End If
    Wend
    
    lNumeroLinha = 0
    cLinhaArquivo = ""
    cLinha = ""
    Do
        While Dir("C:\TEF_DIAL\RESP\INTPOS.001") = ""  ' Verifica o arquivo INTPOS.001 de resposta.
            DoEvents
        Wend
        
        'verifica se o arquivo � valido
        iArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\INTPOS.001" For Input As iArquivo
            
        While Not EOF(iArquivo)
            Line Input #iArquivo, cLinhaArquivo 'L� uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo

            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
            If (cCampoArquivo = "001") Then
                If Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10) = Format(hora, "HhNnSs") Then
                    arquivoIncorreto = False
                End If
            End If
        Wend
        Close iArquivo
        If arquivoIncorreto Then
            MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.001")
        End If
    
    Loop While arquivoIncorreto
    
    While (RealizaTransacao = -2) 'FOR1-IF1-WHILE1
        
        iArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\INTPOS.001" For Input As iArquivo
            
        While Not EOF(iArquivo) 'FOR1-IF1-WHILE1-IF1-DOWHILE1
            Line Input #iArquivo, cLinhaArquivo 'L� uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
            lNumeroLinha = lNumeroLinha + 1
            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)

            Select Case CInt(cCampoArquivo) 'FOR1-IF1-WHILE1-IF1-WHILE1-SELECT1
                Case 9: ' Verifica se a Transa��o foi Aprovada.
                    If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) = "0" Then
                        bTransacao = True
                        RealizaTransacao = 1
                    End If
                    If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> "0" Then
                        bTransacao = False
                        RealizaTransacao = -1
                    End If
                Case 28: ' Verifica se existem linhas para serem impressas.
                    If (CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> 0) And (bTransacao = True) Then
                        '� realizada uma c�pia tempor�ria do arquivo INTPOS.001 para cada transa��o efetuada.
                        'Caso a transa��o necessite ser cancelada, as informa��es estar�o neste arquivo.
                         ' Copia o arquivo para o diret�rio do Gerenciador Padr�o
                        'Se est� aberto, fecha para copiar
                        
                        
                        Close iArquivo 'fecha arquivo
                        FileCopy "C:\TEF_DIAL\RESP\INTPOS.001", "C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001"

                        RealizaTransacao = 1
                        iArquivo = FreeFile
                        Open "C:\TEF_DIAL\RESP\INTPOS.001" For Input As iArquivo
                        While bFlagArq = False
                            Line Input #iArquivo, cLinhaArquivo
                            If Mid(cLinhaArquivo, 1, 3) = 28 Then
                                bFlagArq = True
                            End If
                        Wend
                        For iVezes = 1 To CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) Step 1
                            Line Input #iArquivo, cLinhaArquivo 'L� uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                            If Mid(cLinhaArquivo, 1, 3) = "029" Then
                                cLinha = cLinha + Mid(cLinhaArquivo, 12, Len(cLinhaArquivo) - 12) + vbCrLf
                            End If
                        Next iVezes
                    End If

                Case 30: ' Verifica se o campo � o 030 para mostrar a mensagem para o operador
                    If cLinha <> "" Then
                        frmTEFVariosCartoes.lblMsg.Caption = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                    Else
                        MataArquivo ("C:\TEF_DIAL\REQ\INTPOS.001")
                        frmTEFVariosCartoes.lblMsg.Caption = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                        RealizaTransacao = -1
                    End If
                End Select 'FOR1-IF1-WHILE1-IF1-WHILE1-ENDSELECT1
        Wend
        
    Wend
        ' Cria o arquivo tempor�rio IMPRIME.TXT com a imagem do comprovante
        If (cLinha <> "") Then
            Close iArquivo
            Call GravaArquivo_Binario(App.Path & "\IMPRIME" & CStr(iConta) & ".TXT", cLinha)
        End If
        
        Sleep (1000)
        ' O arquivo INTPOS.STS n�o retornou em 7 segundos, ent�o o operador � informado.
        If (iTentativas = 7) Then
            If Dir("C:\TEF_DIAL\REQ\INTPOS.001") <> "" Then
                MataArquivo ("C:\TEF_DIAL\REQ\INTPOS.001")
                frmTEFVariosCartoes.lblMsg.Caption = "Gerenciador Padr�o n�o est� ativo!"
                RealizaTransacao = 0
                Exit Function
            End If
        End If
        If (RealizaTransacao = 0) Or (RealizaTransacao = -1) Then
            Close iArquivo
        Else
            RealizaTransacao = 1
            Call GravaArquivo_Binario(App.Path & "\PENDENTE.TXT", Trim(CStr(iConta)))
        End If
        
    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.001")

End Function

'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    ImprimeTransacao
'// Objetivo:
'//    Realiza a impress�o da Transa��o TEF
'// Par�metros:
'//   String para a Forma de Pagamento
'//   String para a Valor da Forma de Pagamento
'//   String para o N�mero do Cupom Fiscal (COO)
'//   TDateTime para identificar o n�mero da transa��o
'//   Integer com o n�mero da transa��o
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function ImprimeTransacao(ByVal cFormaPGTO As String, ByVal cValorPago As String, _
                          ByVal cCOO As String, ByVal hora As String, _
                          ByVal iConta As Integer, ByVal Gerencial As Boolean) As Boolean
    Dim cLinhaArquivo As String
    Dim cLinha  As String
    Dim cSaltaLinha As String
    Dim iArquivo As Integer
    Dim iVezes As Integer
    Dim iRetorno As Integer
    Dim tipoImpressora As Integer
    Dim via As Integer
    
'   Neste ponto � criado o arquivo TEF.TXT, indicando que h� uma opera��o de TEF sendo
'   realizada. Caso ocorra uma queda de energia, no momento da impress�o do TEF, e a
'   aplica��o for inicializada, ao identificar a exist�ncia deste arquivo, a transa��o do TEF
'   dever� ser concelada.
    
    Call GravaArquivo_Binario(App.Path & "\TEF.TXT", CStr(iTransacao))
    iRetorno = Bematech_FI_IniciaModoTEF()

    ImprimeTransacao = False
    If Trim(cCOO) = "" Then
        MsgBox "N�o foi poss�vel obter o n�mero do comprovante."
        Call Bematech_FI_FinalizaModoTEF
        If (ImprimeGerencial(iConta) = 1) Then
            ImprimeTransacao = True
            Exit Function
        Else
            Exit Function
        End If
    End If
    If Dir(App.Path + "\IMPRIME" & CStr(iConta) & ".TXT") <> "" Then
        DoEvents
        
        ' Fun��o para bloqueio do teclado e mouse
        iRetorno = Bematech_FI_IniciaModoTEF()
        iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
        
        If Not Gerencial Then
            iRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado(cFormaPGTO, cValorPago, cCOO)
            If Not VerificaRetornoFuncaoImpressora(iRetorno) Then
                Exit Function
            End If
        End If
        
        cLinha = ""
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '          IN�CIO DA LEITURA DE ARQUIVO PARA IMPRESS�O          '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For via = 1 To 2 Step 1
            iArquivo = FreeFile
            Open App.Path & "\IMPRIME" & CStr(iConta) & ".TXT" For Input As iArquivo
            
            While Not EOF(iArquivo)
            '''''''''''''L� uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                Line Input #iArquivo, cLinhaArquivo
                
                'A fun��o de impress�o n�o aceita strings vazias
                If cLinhaArquivo = "" Then
                    cLinhaArquivo = " "
                End If
                
                '''''''''''''Imprime o que foi lido
                If Gerencial Then
                    iRetorno = Bematech_FI_RelatorioGerencial(cLinhaArquivo & vbCrLf)
                Else
                    iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cLinhaArquivo & vbCrLf)
                End If
                   
                '''''''''''''Aqui � feito o tratamento de erro de comunica��o com a impressora
                '''''''''''''(desligamento da impressora durante a impress�o do comprovante).
                If Not (VerificaRetornoFuncaoImpressora(iRetorno)) Then
                    Close iArquivo
                    iRetorno = Bematech_FI_FinalizaModoTEF()
                    iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
                    ImprimeTransacao = False
                    Exit Function
                End If
            Wend
            
            
            
            '''''''''''''Aciona o corte de papel
            If via = 1 Then
                '''''''''''''Pula 7 linhas
                cSaltaLinha = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
                iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
                iRetorno = Bematech_FI_VerificaTipoImpressora(tipoImpressora)
                If ((tipoImpressora = "2") Or (tipoImpressora = "4") Or (tipoImpressora = "6") Or (tipoImpressora = "8")) Then
                    iRetorno = Bematech_FI_AcionaGuilhotinaMFD(0)
                End If
                '''''''''''''Exibe mensagem na tela
                With frmTEFVariosCartoes
                    .lblMsg.Caption = "Por favor, destaque a " & via & "� via."
                    .Show
                    .Refresh
                End With
                Sleep (3000)
            End If
            
            Close iArquivo
            With frmTEFVariosCartoes
                .lblMsg.Caption = ""
                .Show
                .Refresh
            End With
        Next via
        Close iArquivo
        iRetorno = Bematech_FI_FinalizaModoTEF()
        iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
        With frmTEFVariosCartoes
            .lblMsg.Caption = "Por favor, destaque a " & (via - 1) & "� via."
            .Show
            .Refresh
        End With
        Sleep (3000)
        With frmTEFVariosCartoes
            .lblMsg.Caption = ""
            .Show
            .Refresh
        End With
        ImprimeTransacao = True
    End If

    'Desbloqeia o teclado e o mouse
    iRetorno = Bematech_FI_FinalizaModoTEF()
End Function

'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    ConfirmaTransacao
'// Objetivo:
'//    Confirmar a Transa��o TEF
'// Par�metros:
'//   Integer com o n�mero da transa��o
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function ConfirmaTransacao(iConta As Integer) As Boolean

   Dim cLinhaArquivo As String
   Dim cConteudo As String
   Dim iArquivo As Integer
   Dim lFlag As Boolean
   Dim iVezes As Integer
   
   cLinhaArquivo = ""
   cConteudo = ""

    If Dir("C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001") <> "" Then
        If (iConta <> 0) Then
            iArquivo = FreeFile
            Open "C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001" For Binary As iArquivo
        Else
            iArquivo = FreeFile
            Open "C:\TEF_DIAL\RESP\INTPOS.001" For Binary As iArquivo
        End If
        While Not EOF(iArquivo)
            DoEvents
            On Error GoTo FimArquivo
            Line Input #iArquivo, cLinhaArquivo
            If (Mid(cLinhaArquivo, 1, 3) = "001") Or (Mid(cLinhaArquivo, 1, 3) = "002") Or (Mid(cLinhaArquivo, 1, 3) = "010") Or (Mid(cLinhaArquivo, 1, 3) = "012") Or (Mid(cLinhaArquivo, 1, 3) = "027") Then
                cConteudo = cConteudo & cLinhaArquivo & vbCrLf
            End If
            If (Mid(cLinhaArquivo, 1, 3) = "999") Then
                  cConteudo = cConteudo & cLinhaArquivo
            End If
FimArquivo: Wend
        Close iArquivo
        
        cConteudo = "000-000 = CNF" & vbCrLf & cConteudo
        Call GravaArquivo_Binario(App.Path & "\INTPOS.001", cConteudo)
        FileCopy App.Path & "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
        MataArquivo (App.Path & "\INTPOS.001")
        While Not Dir("C:\TEF_DIAL\RESP\INTPOS.STS") <> ""
            DoEvents
            Sleep (1000)
        Wend

        MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
    End If

    'Se o arquivo TEF.TXT, que identifica que houve uma transa��o impressa
    'existir, o mesmo ser� exlu�do.
    MataArquivo (App.Path & "\TEF.TXT")

End Function

'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    NaoConfirmaTransacao
'// Objetivo:
'//    N�o Confirmar a Transa��o TEF
'// Par�metros:
'//   Integer com o n�mero da transa��o
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function NaoConfirmaTransacao(ByVal iConta As Integer) As Boolean
    Dim cLinhaArquivo As String
    Dim cConteudo As String
    Dim cCampoArquivo As String
    Dim iArquivo As Integer
    Dim lFlag As Boolean
    Dim cValor As String
    Dim cNomeRede As String
    Dim cNSU As String
    Dim cIdent As String
    Dim cData As String
    Dim cHora As String
    Dim iVezes As Integer
    
    MataArquivo (App.Path & "\IMPRIME" + CStr(iConta) + ".TXT")
    cLinhaArquivo = ""
    cConteudo = ""
    
    'Se achou o INTPOS[conta].001 na pasta C:\TEF_DIAL\RESP
    If Dir("C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001") <> "" Then
        iArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001" For Input As iArquivo
        While Not EOF(iArquivo)
            DoEvents
            Line Input #iArquivo, cLinhaArquivo
            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
            Select Case CInt(cCampoArquivo)
                Case 1:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                Case 3:
                    cValor = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 10:
                      cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                      cNomeRede = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 12:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                    cNSU = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 27:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                Case 999:
                
                cConteudo = cConteudo & cLinhaArquivo
             End Select
        Wend
        Close iArquivo

        cConteudo = "000-000 = NCN" & vbCrLf & cConteudo
        iArquivo = FreeFile
        
        Open App.Path & "\INTPOS.001" For Output As iArquivo
        Print #iArquivo, cConteudo
        Close iArquivo
        
        FileCopy App.Path & "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
        MataArquivo (App.Path & "\INTPOS.001")
        
        While Dir("C:\TEF_DIAL\RESP\INTPOS.STS") = ""
            DoEvents
            Sleep (1000)
        Wend
    
        MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
    
        'Se o arquivo TEF.TXT, que identifica que houve uma transa��o impressa
        'existir, o mesmo ser� exlu�do.
        MataArquivo (App.Path & "\TEF.TXT")
        frmTEFVariosCartoes.MousePointer = vbDefault
        MsgBox "Cancelada a Transa��o" & vbCrLf & vbCrLf & "Rede: " & _
            cNomeRede & vbCrLf & "Doc N�: " & cNSU & vbCrLf & "Valor: " & _
            Format(CDbl(cValor) / 100, "#,##0.00"), vbOKOnly + vbInformation, _
            "Aten��o"
        frmTEFVariosCartoes.MousePointer = vbHourglass
        MataArquivo ("C:\TEF_DIAL\RESP\INTPOS" & CStr(iConta) & ".001")
        Call Bematech_FI_FechaRelatorioGerencial
        iConta = iConta - 1
        If iConta > 0 Then
            For iVezes = 1 To iConta Step 1
                If Dir("C:\TEF_DIAL\RESP\INTPOS" + CStr(iVezes) + ".001") <> "" Then
                    cLinhaArquivo = ""
                    cConteudo = ""
                    iArquivo = FreeFile
                    Open "C:\TEF_DIAL\RESP\INTPOS" & CStr(iVezes) & ".001" For Input As iArquivo
                        While Not EOF(iArquivo)
                            DoEvents
                            Line Input #iArquivo, cLinhaArquivo
                            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
                            Select Case CInt(cCampoArquivo)
                                Case 1:
                                    cIdent = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 3:
                                    cValor = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 10:
                                    cNomeRede = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 12:
                                    cNSU = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 22:
                                    cData = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 23:
                                    cHora = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                            End Select
                        Wend
                        Close iArquivo
                        MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
                        Call CancelaTransacaoTEF(cNSU, cValor, cNomeRede, cNSU, cData, cHora, iVezes)
                        ConfirmaTransacao (iVezes)
                        Call Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial (iVezes)
                        MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
                        ' Se o arquivo TEF.TXT, que identifica que houve uma transa��o impressa
                        ' existir, o mesmo ser� exclu�do.
                        MataArquivo (App.Path & "\TEF.TXT")
                End If
            Next iVezes
        End If
    
        If iConta > 0 Then
            For iVezes = 1 To iConta Step 1
                MataArquivo ("C:\TEF_DIAL\RESP\INTPOS" & CStr(iVezes) & ".001")
                MataArquivo ("C:\TEF_DIAL\RESP\CANCEL" & CStr(iVezes) & ".001")
                MataArquivo (App.Path & "\IMPRIME" & CStr(iConta) & ".TXT")
                naoConfirmado = True
           Next iVezes
        End If
    End If
    
End Function

'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    CancelaTransacaoTEF
'// Objetivo:
'//    Cancelar uma transa��o j� confirmada
'// Par�metros:
'//    String com o n�mero de identifica��o (NSU)
'//    String com o valor da transa��o
'//    String com o valor da transa��o
'//    String com o nome e bandeira (REDE)
'//    String com o n�mero do documento
'//    String com a data da transa��o no formato DDMMAAAA
'//    String com a hora da transa��o no formato HHSMMSS
'//    Integer com o n�mero da transa��o
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function CancelaTransacaoTEF(ByVal cNSU As String, ByVal cValor As String, ByVal cNomeRede As String, _
         ByVal cNumeroDOC As String, ByVal cData As String, ByVal cHora As String, ByVal iVezes As Integer) As Boolean
    Dim cConteudo As String
    Dim iArquivo As Integer
    Dim lFlag As Boolean
    
    cConteudo = ""
    cConteudo = "000-000 = CNC" & vbCrLf & _
                "001-000 = " & cNSU & vbCrLf & _
                "003-000 = " & cValor & vbCrLf & _
                "010-000 = " & cNomeRede & vbCrLf & _
                "012-000 = " & cNumeroDOC & vbCrLf & _
                "022-000 = " & cData & vbCrLf & _
                "023-000 = " & cHora & vbCrLf & _
                "999-999 = 0"
    iArquivo = FreeFile
    Open App.Path + "\INTPOS.001" For Output As iArquivo
   
    Print #iArquivo, cConteudo
    Close iArquivo
    FileCopy App.Path + "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
    MataArquivo (App.Path + "\INTPOS.001")

    While Dir("C:\TEF_DIAL\RESP\INTPOS.001") = ""
        Sleep (1000)
    Wend

    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
    FileCopy "C:\TEF_DIAL\RESP\INTPOS.001", "C:\TEF_DIAL\RESP\CANCEL" & CStr(iVezes) & ".001"
    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.001")

End Function

'////////////////////////////////////////////////////////////////////////////////
'// Fun��o:
'//    FuncaoAdministrativaTEF
'// Objetivo:
'//    Chamar o m�dulo administrativo da bandeira
'// Par�metro:
'//    String com o identificador
'// Retorno:
'//    1 para OK
'//    diferente de 1 para n�o OK
'////////////////////////////////////////////////////////////////////////////////
Function FuncaoAdministrativaTEF(ByVal hora As String) As Integer


    Dim iArquivo As Integer
    Dim lFlag As Boolean
    Dim cConteudoArquivo As String
    
    'Conte�do do arquivo INTPOS.001 para solicitar a transa��o TEF
    cConteudoArquivo = ""
    cConteudoArquivo = "000-000 = ADM" & vbCrLf & _
                       "001-000 = " & Format(hora, "HhNnSs") & vbCrLf & _
                       "999-999 = 0"
    Call GravaArquivo_Binario(App.Path + "\INTPOS.001", cConteudoArquivo)
    
    FileCopy App.Path & "\INTPOS.001", "C:\TEF_DIAL\REQ\INTPOS.001"
    MataArquivo (App.Path & "\INTPOS.001")

End Function
'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    ImprimeGerencial
'// Objetivo:
'//    Imprimir atrav�s do Relat�rio Gerencial a transa��o efetuada.
'// Par�metro:
'//    Integer com o n�mero da transa��o
'// Retorno:
'//    1 para OK
'//    diferente de 1 para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function ImprimeGerencial(ByVal iConta As Integer) As Integer
    Dim iArquivo As Integer
    Dim iTentativas As Integer
    Dim iVezes As Integer
    Dim iRetorno As Integer
    Dim via As Integer
    Dim tipoImpressora As Integer
    Dim bTransacao As Boolean
    Dim cArquivoTexto As String
    Dim cArquivoIntPos As String
    Dim cArquivoCancel As String
    Dim cCampoArquivo As String
    Dim cLinha As String
    Dim cSaltaLinha As String
    Dim cLinhaArquivo As String
    
    If iConta = 0 Then
        cArquivoTexto = "IMPRIME.TXT"
        cArquivoIntPos = "INTPOS.001"
    Else
        cArquivoTexto = "IMPRIME" & CStr(iConta) & ".TXT"
        cArquivoIntPos = "INTPOS" & CStr(iConta) & ".001"
        cArquivoCancel = "CANCEL" & CStr(iConta) & ".001"
    End If
    MataArquivo (App.Path & "\" & cArquivoTexto)
    
    If Dir("C:\TEF_DIAL\RESP\" & cArquivoCancel) <> "" Then
        cArquivoIntPos = "CANCEL" & CStr(iConta) & ".001"
    End If
    ImprimeGerencial = -2
    
    For iTentativas = 1 To 7 Step 1
        cLinhaArquivo = ""
        cLinha = ""
        While (ImprimeGerencial = -2)
            If Dir("C:\TEF_DIAL\RESP\" & cArquivoIntPos) <> "" Then
                iArquivo = FreeFile
                Open "C:\TEF_DIAL\RESP\" & cArquivoIntPos For Input As iArquivo
                While Not EOF(iArquivo)
                    Line Input #iArquivo, cLinhaArquivo
                    cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
                    Select Case CInt(cCampoArquivo)
                        Case 9: ' Verifica se a Transa��o foi Aprovada
                            If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) = "0" Then
                                bTransacao = True
                                ImprimeGerencial = 1
                            End If
                            If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> "0" Then
                                bTransacao = False
                                ImprimeGerencial = -1
                            End If

                        Case 28: 'Verifica se existem linhas para serem impressas
                            If (CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> 0) And (bTransacao = True) Then
                                ImprimeGerencial = 1
                                For iVezes = 1 To CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) Step 1
                                    Line Input #iArquivo, cLinhaArquivo
                                    If Mid(cLinhaArquivo, 1, 3) = "029" Then
                                        cLinha = cLinha & Mid(cLinhaArquivo, 12, Len(cLinhaArquivo) - 12) & vbCrLf
                                    End If
                                Next iVezes
                            End If

                        Case 30: 'Verifica se o campo � o 030 para mostrar a mensagem para o operador
                            If cLinha <> "" Then
                                frmTEFVariosCartoes.lblMsg.Caption = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                            Else
                                If Dir("C:\TEF_DIAL\REQ\INTPOS.001") <> "" Then
                                    MataArquivo ("C:\TEF_DIAL\REQ\INTPOS.001")
                                    frmTEFVariosCartoes.lblMsg.Caption = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                    ImprimeGerencial = -1
                                End If
                            End If
                    End Select
                Wend
            End If
        Wend

        'Cria o arquivo tempor�rio IMPRIME.TXT com a imagem do comprovante
        If (cLinha <> "") Then
            Close iArquivo
            Call GravaArquivo_Binario(App.Path & "\IMPRIME" & CStr(iConta) & ".TXT", cLinha)
            Exit For
        End If

        Sleep (1000)

        'O arquivo INTPOS.STS n�o retornou em 7 segundos, ent�o o operador � informado.
        If (iTentativas = 7) Then
        
            MataArquivo ("C:\TEF_DIAL\REQ\INTPOS.001")
            frmTEFVariosCartoes.lblMsg.Caption = "Gerenciador Padr�o n�o est� ativo!"
            ImprimeGerencial = 0
            Exit For
        End If
        If (ImprimeGerencial = 0) Or (ImprimeGerencial = -1) Then
            Close iArquivo
            Exit For
        End If
    Next iTentativas

    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.STS")
    MataArquivo ("C:\TEF_DIAL\RESP\INTPOS.001")

    If Dir(App.Path + "\IMPRIME" & CStr(iConta) & ".TXT") <> "" Then
        'Bloqueia o teclado e o mouse para a impress�o do TEF
        iRetorno = Bematech_FI_IniciaModoTEF()
        
        ''''''''IMPRESS�O DO RELAT�RIO GERENCIAL'''''''''''
        
        For via = 1 To 2 Step 1
            iArquivo = FreeFile
            Open App.Path & "\IMPRIME" & CStr(iConta) & ".TXT" For Input As iArquivo
            
            While Not EOF(iArquivo)
            '''''''''''''L� uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                Line Input #iArquivo, cLinhaArquivo
                'A fun��o de impress�o n�o aceita strings vazias
                If cLinhaArquivo = "" Then
                    cLinhaArquivo = " "
                End If
                
                '''''''''''''Imprime o que foi lido
                iRetorno = Bematech_FI_RelatorioGerencial(cLinhaArquivo & vbCrLf)
                   
                '''''''''''''Aqui � feito o tratamento de erro de comunica��o com a impressora
                '''''''''''''(desligamento da impressora durante a impress�o do comprovante).
                If Not (VerificaRetornoFuncaoImpressora(iRetorno)) Then
                    iRetorno = Bematech_FI_FinalizaModoTEF()
                    frmTEFVariosCartoes.MousePointer = vbDefault
                    If (MsgBox("A impressora n�o responde!" & vbCrLf & _
                        "Deseja imprimir novamente?", vbYesNo + vbQuestion, "Aten��o") = vbYes) Then
                        Close iArquivo
                        iRetorno = Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial (iConta)
                        Exit Function
                    Else
                        Close iArquivo
                        iRetorno = Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial = 0
                        Exit Function
                    End If
                End If
            Wend
            
            
            
            '''''''''''''Aciona o corte de papel
            If via = 1 Then
                '''''''''''''Pula 7 linhas
                cSaltaLinha = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
                iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
                iRetorno = Bematech_FI_VerificaTipoImpressora(tipoImpressora)
                If ((tipoImpressora = "2") Or (tipoImpressora = "4") Or (tipoImpressora = "6") Or (tipoImpressora = "8")) Then
                    iRetorno = Bematech_FI_AcionaGuilhotinaMFD(0)
                End If
                '''''''''''''Exibe mensagem na tela
                With frmTEFVariosCartoes
                    .lblMsg.Caption = "Por favor, destaque a " & via & "� via."
                    .Show
                    .Refresh
                End With
                Sleep (3000)
            End If
            
            Close iArquivo
            With frmTEFVariosCartoes
                .lblMsg.Caption = ""
                .Show
                .Refresh
            End With
        Next via
        Close iArquivo
        iRetorno = Bematech_FI_FechaRelatorioGerencial()
        VerificaRetornoFuncaoImpressora (iRetorno)
    End If

    'Desbloqeia o teclado e o mouse
    iRetorno = Bematech_FI_FinalizaModoTEF()
    MataArquivo (App.Path & "\IMPRIME" & CStr(iConta) & ".TXT")

End Function
'////////////////////////////////////////////////////////////////////////////////
'//
'// Fun��o:
'//    VerificaRetornoFuncaoImpressora
'// Objetivo:
'//    Verificar o retorno da impressora e da fun��o utilizada
'// Retorno:
'//    True para OK
'//    False para n�o OK
'//
'////////////////////////////////////////////////////////////////////////////////
Function VerificaRetornoFuncaoImpressora(ByVal iRetorno As Integer) As Boolean

   Dim cMSGErro As String
   Dim iACK As Integer
   Dim iST1 As Integer
   Dim iST2 As Integer
   
   iACK = 0: iST1 = 0: iST2 = 0
   
    cMSGErro = ""
    VerificaRetornoFuncaoImpressora = False
    Select Case iRetorno
        Case 0:
           cMSGErro = "Erro de Comunica��o !"
        Case -1:
            cMSGErro = "Erro de execu��o na Fun��o !"
        Case -2:
            cMSGErro = "Par�metro inv�lido na Fun��o !"
        Case -3:
            cMSGErro = "Al�quota n�o Programada !"
        Case -4:
            cMSGErro = "Arquivo BEMAFI32.INI n�o Encontrado !"
        Case -5:
            cMSGErro = "Erro ao abrir a Porta de Comunica��o !"
        Case -6:
            cMSGErro = "Impressora Desligada ou Cabo de Comunica��o Desconectado !"
        Case -7:
            cMSGErro = "C�digo do Banco n�o encontrado no arquivo BEMAFI32.INI !"
        Case -8:
            cMSGErro = "Erro ao criar ou gravar arquivo STATUS.TXT ou RETORNO.TXT !"
        Case -27:
            cMSGErro = "Status diferente de 6, 0, 0 !"
        Case -30:
            cMSGErro = "Fun��o incompat�vel com a impressora fiscal YANCO !"
    End Select

    If cMSGErro <> "" Then 'IF1
        Call Bematech_FI_FinalizaModoTEF
        VerificaRetornoFuncaoImpressora = False
    End If

    cMSGErro = ""
    If iRetorno = 1 Then 'IF2
      
        Call Bematech_FI_RetornoImpressora(iACK, iST1, iST2)
        If iACK = 21 Then 'IF2-1
            Call Bematech_FI_FinalizaModoTEF
            MsgBox "A Impressora retornou NAK !" & vbCrLf & _
                                       "Erro de Protocolo de Comunica��o !", vbOKOnly, _
                                       "Aten��o"
            VerificaRetornoFuncaoImpressora = False
        
        Else 'ELSEIF2-1
            If (iST1 <> 0) Or (iST2 <> 0) Then 'IF2-1-1
                  ' Analisa ST1
                If (iST1 >= 128) Then 'IF2-1-1-1
                    iST1 = iST1 - 128
                    cMSGErro = cMSGErro & "Fim de Papel" & vbCrLf
                End If 'ENDIF2-1-1-1
                If (iST1 >= 64) Then 'IF2-1-1-2
                    iST1 = iST1 - 64
                    cMSGErro = cMSGErro & "Pouco Papel" & vbCrLf
                    VerificaRetornoFuncaoImpressora = True
                    Exit Function
                End If 'ENDIF2-1-1-2
                If (iST1 >= 32) Then 'IF2-1-1-3
                    iST1 = iST1 - 32
                    cMSGErro = cMSGErro & "Erro no Rel�gio" & vbCrLf
                End If 'ENDIF2-1-1-3
                If (iST1 >= 16) Then 'IF2-1-1-4
                    iST1 = iST1 - 16
                    cMSGErro = cMSGErro & "Impressora em Erro" & vbCrLf
                End If 'ENDIF2-1-1-4
                If (iST1 >= 8) Then 'IF2-1-1-5
                    iST1 = iST1 - 8
                    cMSGErro = cMSGErro & "Primeiro Dado do Comando n�o foi ESC" & vbCrLf
                End If 'ENDIF2-1-1-5
                If iST1 >= 4 Then 'IF2-1-1-6
                    iST1 = iST1 - 4
                    cMSGErro = cMSGErro & "Comando Inexistente" & vbCrLf
                End If 'ENDIF2-1-1-6
                If iST1 >= 2 Then 'IF2-1-1-7
                    iST1 = iST1 - 2
                    cMSGErro = cMSGErro & "Cupom Fiscal Aberto" & vbCrLf
                End If 'ENDIF2-1-1-7
                If iST1 >= 1 Then 'IF2-1-1-8
                    iST1 = iST1 - 1
                    cMSGErro = cMSGErro & "N�mero de Par�metros Inv�lidos" & vbCrLf
                End If 'ENDIF2-1-1-8
                'Analisa ST2
                If iST2 >= 128 Then 'IF2-1-1-9
                    iST2 = iST2 - 128
                    cMSGErro = cMSGErro & "Tipo de Par�metro de Comando Inv�lido" & vbCrLf
                End If 'ENDIF2-1-1-9
                If iST2 >= 64 Then 'IF2-1-1-10
                    iST2 = iST2 - 64
                    cMSGErro = cMSGErro & "Mem�ria Fiscal Lotada" & vbCrLf
                End If 'ENDIF2-1-1-10
                If iST2 >= 32 Then 'IF2-1-1-11
                    iST2 = iST2 - 32
                    cMSGErro = cMSGErro & "Erro na CMOS" & vbCrLf
                End If 'ENDIF2-1-1-11
                If iST2 >= 16 Then 'IF2-1-1-12
                    iST2 = iST2 - 16
                    cMSGErro = cMSGErro & "Al�quota n�o Programada" & vbCrLf
                End If 'ENDIF2-1-1-12
                If iST2 >= 8 Then 'IF2-1-1-13
                    iST2 = iST2 - 8
                    cMSGErro = cMSGErro & "Capacidade de Al�quota Program�veis Lotada" & vbCrLf
                End If 'ENDIF2-1-1-13
                If iST2 >= 4 Then 'IF2-1-1-14
                     iST2 = iST2 - 4
                     cMSGErro = cMSGErro & "Cancelamento n�o permitido" & vbCrLf
                End If 'ENDIF2-1-1-14
                If iST2 >= 2 Then 'IF2-1-1-15
                    iST2 = iST2 - 2
                    cMSGErro = cMSGErro & "CGC/IE do Propriet�rio n�o Programados" & vbCrLf
                End If 'ENDIF2-1-1-15
                If iST2 >= 1 Then 'IF2-1-1-16
                    iST2 = iST2 - 1
                    cMSGErro = cMSGErro & "Comando n�o executado" & vbCrLf
                End If 'ENDIF2-1-1-16
                If (cMSGErro <> "") Then 'IF2-1-1-17
                    Call Bematech_FI_FinalizaModoTEF
                    MsgBox cMSGErro, vbOKOnly + vbExclamation, "Aten��o"
                    If VerificaRetornoFuncaoImpressora = True Then
                        VerificaRetornoFuncaoImpressora = False
                    End If
                End If 'ENDIF2-1-1-17
            Else
                VerificaRetornoFuncaoImpressora = True
            End If 'ENDIF2-1-1
        End If 'ENDIF2-1
    End If 'ENDIF2

End Function
Public Sub CarregarFormasPagamento()
    Dim formasPagto As New Collection
    Dim formasdePagamento As String
        
    Dim i As Integer
    Dim j As Integer
    Dim tamanho As Integer
    Dim Item As Variant

    ' Verifica se existe o arquivo TEF.TXT, indicando que houve uma queda de
    ' energia e que existe uma transa��o pendente.
    formasdePagamento = Space(3016)
    response = Bematech_FI_VerificaFormasPagamento(formasdePagamento)
    j = 3016
    Set formasPagto = Nothing
    tamanho = 16
    For i = 1 To j Step 58
        formasPagto.Add (Mid(formasdePagamento, i, tamanho))
    Next i
    For Each Item In formasPagto
        If Trim(Item) <> "" Then
            frmTEFVariosCartoes.cboFormaPagto.AddItem (Trim(Item))
        End If
    Next Item

End Sub
Public Sub CancelarTransacoesPendentes()
    Dim iArquivo As Integer
    iArquivo = FreeFile
    Open App.Path + "\TEF.TXT" For Input As iArquivo
    'L� o conte�do do arquivo
    If Not EOF(iArquivo) Then
        Line Input #iArquivo, linhaArquivo
    End If
    Close iArquivo
    
    'Se leu algo do arquivo ent�o...
    If linhaArquivo <> "" Then
        For i = 0 To Len(linhaArquivo) Step 1
            'Se o que leu for num�rico...
            If IsNumeric(Mid(linhaArquivo, i + 1, 1)) Then
                'o auxiliar cLinha1 recebe o conte�do num�rico de cLinha
                Call NaoConfirmaTransacao(CInt(Mid(linhaArquivo, i + 1, 1)))
            End If
        Next i
    End If
End Sub
Public Sub MataArquivo(ByVal caminho As String)
    If Dir(caminho) <> "" Then
            Kill caminho
    End If
End Sub
Public Sub GravaArquivo_Binario(ByVal caminho As String, ByVal dados As String)
    Dim iArquivo As Integer
    
    iArquivo = FreeFile
    Open caminho For Binary As iArquivo
        ' Escreve no arquivo
        Put iArquivo, , dados
        ' Fecha o arquivo
    Close iArquivo
End Sub
Public Sub GravaArquivo_Random(ByVal caminho As String, ByVal dados As String)
    Dim iArquivo As Integer
    
    iArquivo = FreeFile
    Open caminho For Random As iArquivo
        ' Escreve no arquivo
        Put #iArquivo, , dados
        ' Fecha o arquivo
    Close iArquivo

End Sub


Public Sub GravaArquivo_Output(ByVal caminho As String, ByVal dados As String)
    Dim iArquivo As Integer
    
    iArquivo = FreeFile
    Open caminho For Output As iArquivo
        ' Escreve no arquivo
        Print #iArquivo, , dados
        ' Fecha o arquivo
    Close iArquivo

End Sub
