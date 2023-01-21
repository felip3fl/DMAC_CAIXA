Attribute VB_Name = "Modulo_NFE_GeraTexto"
'Global GLB_ConectouOK As Boolean
'Global rdoCNLoja As New ADODB.Connection
'Global RsDados As New ADODB.Recordset
Dim RsMovimentoCaixa As New ADODB.Recordset
Dim RsDadosUP As New ADODB.Recordset
Dim sql As String
Dim SQLItens As String
Dim SQLPROD As String
Dim SQLCLI As String
'---------------------------------
 Dim wAliqICMS As String
 Dim wValorUnitario As String
 Dim wValorUnitarioInteiro As Long
 Dim wQuantidade As String
 Dim wQuantidadeInteiro As Long
 Dim wUnidade As String
 Dim wCodigoProduto As String
 Dim wDescricaoProduto As String
 Dim wDesconto As String
 Dim wDescontoInteiro As Long
 Dim wInformacaoAdicional As String
 Dim wFormaPgtoCodigo As String
 Dim wFormaPgtoValor As String
 Dim wFormaPgtoValorInteiro As Long
 Dim wFormaPgtoDescricao As String
 Dim wCPFCNPJCLiente As String
 Dim wNomeCLiente As String
 'Dim wLoja As String
 Dim wNF  As String
 Dim wSerie As String
'---------------------------------
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'-----------------------------------------------------------------------------------------
Type notaFiscal
    numero As String
    loja As String
    eSerie As String
    'serie As String
    cnpj As String
    chave As String
    pedido As String
    cfop As String
End Type


Public Function GeraArqECFTXT(ByVal NroPedido As Long)

sql = " "
sql = "Select I.NUMEROPED,I.REFERENCIA,I.QTDE,I.VLUNIT,I.ICMS,C.DESCONTO,CE_razao," _
    & "CE_cgc,PR_Descricao,PR_Unidade,PR_SubstituicaoTributaria,C.LojaOrigem,C.NF,C.Serie,C.cpfnfp " _
    & "From NFItens as I ,NFCAPA as C,ProdutoLoja,fin_cliente " _
    & "Where I.Referencia = PR_Referencia and I.NUMEROPED = C.NUMEROPED " _
    & " and I.NUMEROPED = " & NroPedido & " and C.Cliente = ce_CodigoCliente"
     RsDados.CursorLocation = adUseClient
     RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '
     If RsDados.EOF Then
        RsDados.Close
        Exit Function
     End If
     
     If Not RsDados.EOF Then
        txtECF = FreeFile
        wNomeArquivo = right("00000000" & NroPedido, 8) & ".Ped"
       
        Open "C:\ECF\" & wNomeArquivo For Output Access Write As #txtECF
       
        linhaArquivo = "00 "
        Print #txtECF, linhaArquivo
        wCPFCNPJCLiente = right("00000000000000" & RsDados("cpfnfp"), 14)
        wNomeCLiente = RsDados("ce_razao")
        wInformacaoAdicional = " "
        'wDescontoInteiro = (RsDados("DESCONTO") * 100)
        wDescontoInteiro = Format(RsDados("DESCONTO") * 100)
        
        wLoja = RsDados("LojaOrigem")
        wNF = RsDados("NF")
        wSerie = RsDados("Serie")
        wDesconto = right("00000000000000" & wDescontoInteiro, 14)
        
        Do While Not RsDados.EOF
           If RsDados("PR_SubstituicaoTributaria") = "S" Then
              wAliqICMS = "FF"
           Else
              wAliqICMS = Format(RsDados("ICMS"), "00")
           End If
              wValorUnitarioInteiro = Format(RsDados("VLUNIT") * 100)
           
           wQuantidadeInteiro = (RsDados("QTDE") * 1000)
           wUnidade = RsDados("PR_Unidade")
           wCodigoProduto = RsDados("Referencia")
           wDescricaoProduto = RsDados("PR_Descricao")
           wAliqICMS = right("00" & wAliqICMS, 2)
           wValorUnitario = right("000000000" & wValorUnitarioInteiro, 9)
           wQuantidade = right("0000000" & wQuantidadeInteiro, 7)
           wCodigoProduto = right("00000000000000" & wCodigoProduto, 14)
          
           linhaArquivo = "63 " & wAliqICMS & wValorUnitario & wQuantidade & wUnidade & wCodigoProduto _
                                & wDescricaoProduto
           Print #txtECF, linhaArquivo
           RsDados.MoveNext
        Loop
        
        linhaArquivo = "32 " & wDesconto
        Print #txtECF, linhaArquivo
        sql = " "
        sql = "Select MO_Descricao,MO_OrdemApresentacao,Movimentocaixa.* " _
            & " From Movimentocaixa,Modalidade Where mo_grupo = mc_grupo and mc_documento= '" _
            & LTrim(RTrim(wNF)) & "' and mc_Serie = '" & LTrim(RTrim(frmFormaPagamento.txtSerie.text)) & "' and MC_pedido = '" & NroPedido _
            & "' and MC_loja = '" & LTrim(RTrim(wLoja)) & "' and mc_grupo between '10101' and '10304'" _
            & " order by mc_documento"
            RsMovimentoCaixa.CursorLocation = adUseClient
            RsMovimentoCaixa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '
        If RsMovimentoCaixa.EOF Then
           RsMovimentoCaixa.Close
           RsDados.Close
           Printer.EndDoc
           Close #txtECF
           Exit Function
        End If
        If Not RsMovimentoCaixa.EOF Then
           Do While Not RsMovimentoCaixa.EOF
              wFormaPgtoCodigo = RsMovimentoCaixa("MO_OrdemApresentacao")
              wFormaPgtoDescricao = RsMovimentoCaixa("MO_Descricao")
              wFormaPgtoValorInteiro = (RsMovimentoCaixa("MC_Valor") * 100)
              wFormaPgtoValor = wFormaPgtoValorInteiro
              wFormaPgtoCodigo = right("00" & wFormaPgtoCodigo, 2)
              wFormaPgtoValor = right("00000000000000" & wFormaPgtoValor, 14)
              linhaArquivo = "72 " & wFormaPgtoCodigo & wFormaPgtoValor & wFormaPgtoDescricao
              Print #txtECF, linhaArquivo
              RsMovimentoCaixa.MoveNext
           Loop
        End If

        
        
       ' Call FormadePagamento
        wCPFCNPJCLiente = right("00000000000000" & Trim(wCPFCNPJCLiente), 14)
        linhaArquivo = "44 " & wCPFCNPJCLiente
        Print #txtECF, linhaArquivo
        linhaArquivo = "45 " & wNomeCLiente
        Print #txtECF, linhaArquivo
       
        Printer.EndDoc
        Close #txtECF
     
     End If
        RsDados.Close
        RsMovimentoCaixa.Close
End Function
           
Private Sub VerificarexistenciaTabela()


Set RsDados = rdoCNLoja.OpenSchema(adSchemaTables, Array(Empty, Empty, GLB_Tabela, "TABELA"))

If Not RsDados.EOF Then
   GLB_ExisteTabela = 1
Else
   GLB_ExisteTabela = 0
End If

RsDados.Close

End Sub

Public Function montaTXTSAT(pedido As String) As String

    Dim ado_estrutura As New ADODB.Recordset
    
    sql = "exec SP_VDA_Cria_Cupom '" & pedido & "'"
    rdoCNLoja.Execute sql
    
    sql = "select snf_Descricao as descricao, snf_Sinal as sinal, snf_Dados as dados " & _
          "from sat_nf " & _
          "where snf_pedido = '" & pedido & "' " & _
          "order by snf_Sequencia"
    
    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    Do While Not ado_estrutura.EOF
        montaTXTSAT = montaTXTSAT & ado_estrutura("descricao") & ado_estrutura("sinal") & " " & ado_estrutura("dados")
        montaTXTSAT = montaTXTSAT & vbNewLine
        
        ado_estrutura.MoveNext
    Loop
        
    ado_estrutura.Close
    
End Function

Public Function montaTXT(Nf As notaFiscal) As String
    Dim ado_estrutura As New ADODB.Recordset

    sql = "select nfl_descricao, nfl_dados " & _
          "from NFE_NFLojas " & _
          "where nfl_loja = '" & Nf.loja & "' and nfl_nroNFE = '" & Nf.numero & "'" & _
          "order by NFL_sequencia, nfl_NROnfe, nfl_dados desc"
    
    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    Do While Not ado_estrutura.EOF
        If left(ado_estrutura("nfl_descricao"), 1) = "[" Or left(ado_estrutura("nfl_descricao"), 2) = "--" Then
            montaTXT = montaTXT & vbNewLine & vbNewLine & ado_estrutura("nfl_descricao")
        Else
            montaTXT = montaTXT & vbNewLine & ado_estrutura("nfl_descricao") & "= " & Trim(ado_estrutura("nfl_dados"))
        End If
        
        ado_estrutura.MoveNext
    Loop
        
    ado_estrutura.Close
End Function

Public Sub mensagemErroDesconhecido(numeroErro As ErrObject, nomeFormulario As String)
    MsgBox "Ocorreu um erro desconhecido durante a execução" & vbNewLine & _
    "Código: " & numeroErro.Number & vbNewLine & "Descrição: " & numeroErro.Description, vbCritical, nomeFormulario
    End
End Sub

Public Function criaTXTSAT(tipoTXT As String, Nf As notaFiscal)
    Dim corpoMensagem As String
    
On Error GoTo TrataErro
    
    corpoMensagem = montaTXTSAT(Nf.pedido)
    Open GLB_EnderecoPastaFIL & _
    tipoTXT & (Format(Nf.pedido, "000000000")) & "#" & Nf.cnpj & ".txt" For Output As #1
         Print #1, corpoMensagem
    Close #1
    
    Exit Function
TrataErro:
    Select Case Err.Number
    Case Else
        mensagemErroDesconhecido Err, "Criação de arquivo"
    End Select
End Function

Public Function criaTXT(tipoTXT As String, Nf As notaFiscal)
    Dim corpoMensagem As String
    
On Error GoTo TrataErro
    
    corpoMensagem = montaTXT(Nf)
    Open GLB_EnderecoPastaFIL & _
    tipoTXT & (Format(Nf.numero, "000000000")) & "#" & Nf.cnpj & ".txt" For Output As #1
         Print #1, Mid(corpoMensagem, 4, Len(corpoMensagem))
    Close #1
    
    Exit Function
TrataErro:
    Select Case Err.Number
    Case Else
        mensagemErroDesconhecido Err, "Criação de arquivo"
    End Select
End Function
