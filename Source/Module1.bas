Attribute VB_Name = "Module1"
'==== M a n u t en c a o


   Global wRazao As String
   Global Wendereco As String
   Global wbairro      As String
   Global WCGC      As String
   Global WIest      As String
   Global WMunicipio      As String
   Global westado      As String
   Global WCep     As String
   Global WFone      As String
   Global wDDDLoja      As String
   Global WFax      As String
   Global wLoja      As String
   Global wNovaRazao      As String
   Global wPessoa As Double
   Global wChaveICMSItem As Double
   Global wCarimbo1 As String * 132
   Global wCarimbo2 As String * 132
   Global wCarimbo3 As String * 132
   Global wCarimbo5 As String * 132
   Global wRecebeCarimboAnexo As String * 132
   Global wAnexoIten As Integer
   Global wAnexo1 As String
   Global wAnexo2 As String
   Global GLB_AliquotaICMS As Double
   Global GLB_IE_BasedeReducao As Double
   Global wAliqICMSInterEstadual As Double
   Global wUFCliente As String
   Global wValorComplementoAlfa As String
   Global wTipodeComplemento As Integer
   Global wValorComplementoDate As String
   Global wValorComplementoNumerico As Double
   Global Wentrada As Double
   Global Wcondicao As String * 30
   Global wStr0, wStr1, wStr2, wStr3, wStr4, wStr5, wStr6, wStr7 As String
   Global wStr8, wStr9, wStr10, wStr11, wStr12, wStr13, wStr15, wStr16, wStr17, wStr18, wStr19, wStr20, wStr21 As String
   Global WNatureza As String
   Global wReferenciaEspecial As String
   Global wPegaCarimbo1 As String
   Global wPegaCarimbo2 As String
   Global wPegaCarimbo3 As String
   Global wPgentra As Double
   Global wPegaCliente As String
   Global wPegaDesconto As String
   Global wPegaCarimbo5 As String
   Global wPegaDataSaida As String
   Global wPegaFrete As Double
   Global wBaseIcms As Double
   Global wVLRICMS As Double
   Global wPegaDescricaoAlternativa As String
   Global Glb_ImpNotaFiscal As String
   Global wPegaloja As String
   Global wPegaVendedorVenda As String
   Global wLojaVenda As String
   Global wTipoNota As String
   Global wDatapag As Date
   Global wCodigoOperacao As Integer
   Global wPedCli As String
   Global wVLUNIT2 As Double
   Global GLB_BasedeCalculoICMS As Double
   Global GLB_BaseTotalICMS As Double
   Global GLB_TotalICMSCalculado As Double
   Global wTotalVenda As Double
   Global wTotalnota As Double
   Global AuxChaveICMSItem As Double
   Global wPagina As Integer
   
   Global GridPrecoUnitario As String * 12
   Global GridPrecoUnitariodouble As Double
   Global GridValorTotalItem As String * 12
   
'==== M a n u t en c a o
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Ret As String
Public index As Integer
Public indexs As String


Public Sub WriteINI(FileName As String, Section As String, key As String, Text As String)
WritePrivateProfileString Section, key, Text, FileName
End Sub

Public Function ReadINI(FileName As String, Section As String, key As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Section, key, "", Ret, Len(Ret), FileName)
    If RetLen = 0 Then
        Exit Function
    End If
    Ret = Left$(Ret, RetLen)
    ReadINI = Ret
End Function
Public Function LimpaGrid(ByRef GradeUsu)
    GradeUsu.Rows = GradeUsu.FixedRows + 1
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows
End Function
Public Function ConectaBancoLoja()
  If ConectaOdbcBalcao(rdoCNLoja, Usuario, Senha) = False Then
        MsgBox "Não foi possivel conectar-se ao banco de dados do Balcão", vbCritical, "Aviso"
        Exit Function
  Else
        MsgBox "Conexão estabelecida com sucesso", vbInformation
  End If
End Function

Public Function Numeros(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim Charlido As String * 1
    Dim Retorno As String
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        Charlido = Mid(Texto, Char, 1)
        If IsNumeric(Charlido) Then
            Retorno = Retorno & Charlido
        End If
    Next Char
    
    Texto = Retorno
    
    Numeros = Texto

End Function



Public Function DadosLoja()

    SQL = ""
    SQL = "Select CTS_Loja,Lojas.* from lojas,Controlesistema where lo_loja=CTS_Loja"

    RsDados.CursorLocation = adUseClient
    RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not RsDados.EOF Then

       wRazao = Trim(RsDados("lo_Razao"))
       Wendereco = RsDados("lo_ENDERECO")
       wbairro = RsDados("lo_bairro")
       WCGC = RsDados("lo_CGC")
       WIest = RsDados("lo_INSCRICAOESTADUAL")
       WMunicipio = RsDados("lo_MUNICIPIO")
       westado = RsDados("lo_UF")
       WCep = RsDados("lo_CEP")
       WFone = RsDados("lo_TELEFONE")
       wDDDLoja = RsDados("LO_DDD")
       WFax = RsDados("lo_Fax")
       wLoja = RsDados("CTS_Loja")
       wNovaRazao = IIf(IsNull(RsDados("lo_NovaRazao")), "0", RsDados("lo_NovaRazao"))
    
    End If
    
    RsDados.Close

End Function

Function Cabecalho(ByVal TipoNota As String)
    Dim wCgcCliente As String
    Dim impri As Long
    'Dim rdoConPag As rdoResultset
    impri = Printer.Orientation
    'Printer.PrintQuality = vbPRPQDraft
    wDatapag = 0
    wcondpag = 0
    wCodigoOperacao = 0
    wPgentra = 0
    wPedCli = 0
    
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    
            
    Wcondicao = "            "
    Wav = "          "
    
   
    wCarimbo4 = ""
    
       '===  Pegando Data Pagamento
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 18) = True Then
       
        If wTipodeComplemento = 1 Then
           wDatapag = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wDatapag = Trim(wValorComplementoNumerico)
            Else
               wDatapag = Trim(wValorComplementoAlfa)
        End If
       
       'wDatapag = Trim(wValorComplementoAlfa))
    End If
    
       '===  Pegando Condicao de pagamento
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 4) = True Then
       'Dim wDatapag As String
       If wTipodeComplemento = 1 Then
           wcondpag = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wcondpag = Trim(wValorComplementoNumerico)
       Else
               wcondpag = Trim(wValorComplementoAlfa)
       End If
       
       'wCondpag = Trim(wValorComplementoAlfa)
      If Trim(wValorComplementoAlfa) = "85" Then
          wCarimbo4 = Format(wDatapag, "dd/mm/yyyy")
          
      Else
          SQL = ""
          SQL = "Select CP_Condicao from Condicaopagamento " _
              & "where CP_Codigo=" & wcondpag & ""
           rdoComplemento.CursorLocation = adUseClient
           rdoComplemento.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
          
           If Not rdoComplemento.EOF Then
              wCarimbo4 = rdoComplemento("CP_Condicao")
           End If
           rdoComplemento.Close
      End If
    End If
    
       '===  Pegando Loja Venda
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 9) = True Then
       If wTipodeComplemento = 1 Then
           wPegaloja = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wPegaloja = Trim(wValorComplementoNumerico)
       Else
               wPegaloja = Trim(wValorComplementoAlfa)
       End If
       
      ' wPegaloja = Trim(wValorComplementoAlfa)
    End If
    
       '===  Pegando Vendedor Venda
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 10) = True Then
       If wTipodeComplemento = 1 Then
           wPegaVendedorVenda = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wPegaVendedorVenda = Trim(wValorComplementoNumerico)
       Else
               wPegaVendedorVenda = Trim(wValorComplementoAlfa)
       End If
       
      ' wPegaVendedorVenda = Trim(wValorComplementoAlfa)
    End If
    
      
    wLojaVenda = "            "
    wVendedorLojaVenda = "            "
    wLojaVenda = IIf(wPegaloja = "", Trim(RsDados("ITV_Loja")), wPegaloja)
    wVendedorLojaVenda = IIf(wPegaVendedorVenda = "", 0, wPegaVendedorVenda)
    Wentrada = 0
    Wcondicao = "            "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = "               "
    If Val(wcondpag) = 1 Then
       Wcondicao = "Avista"
    ElseIf Val(wcondpag) = 3 Then
       Wcondicao = "Financiada"
    ElseIf Val(wcondpag) > 3 Then
       
       Wcondicao = wCarimbo4
    End If
    

        '===  Pegando Codigo Operacao
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 2) = True Then
        
        If wTipodeComplemento = 1 Then
           wCodigoOperacao = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wCodigoOperacao = Trim(wValorComplementoNumerico)
        Else
               wCodigoOperacao = Trim(wValorComplementoAlfa)
        End If
        
        'wCodigoOperacao = Trim(wValorComplementoAlfa)
    End If
    
 
        WNatureza = "VENDA"
  
    If Trim(wLojaVenda) <> "" Then
       If Trim(wLojaVenda) > 0 Then
          If Trim(wLojaVenda) <> Trim(RsDados("ITV_Loja")) Then
              wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
          Else
              wStr6 = ""
          End If
       Else
         wStr6 = ""
       End If
    End If
    
    wStr17 = "Pedido        : " & RsDados("ITV_NUMEROPEDIDO")
    wStr18 = "Vendedor      : " & RsDados("ITV_VENDEDOR")
    
        '===  Pegando Carimbo3
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 13) = True Then
       If wTipodeComplemento = 1 Then
           wPegaCarimbo3 = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wPegaCarimbo3 = Trim(wValorComplementoNumerico)
        Else
               wPegaCarimbo3 = Trim(wValorComplementoAlfa)
        End If
        
      '  wPegaCarimbo3 = Trim(wValorComplementoAlfa)
    End If
    
    If Trim(Wcondicao) <> "" Then
        wStr19 = "Cond Pagto : " & Trim(Wcondicao)
    ElseIf Trim(wPegaCarimbo3) <> "" Then
        wStr19 = "Transporte    : " & Left(Format(Trim(wPegaCarimbo3)) & Space(10), 10)
    Else
        Wcondicao = "            "
    End If
    
        '===  Pegando Pagamento Entrada
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 17) = True Then
        If wTipodeComplemento = 1 Then
           wPgentra = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wPgentra = Trim(wValorComplementoNumerico)
        Else
               wPgentra = Trim(wValorComplementoAlfa)
        End If
        'wPgentra = wValorComplementoNumerico
    End If

    If wPgentra <> 0 Then
       Wentrada = Format(wPgentra, "#####0.00")
       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
    End If
    
      '===  Pegando Pedido Cliente
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 7) = True Then
        
        If wTipodeComplemento = 1 Then
           wPedCli = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wPedCli = Trim(wValorComplementoNumerico)
        Else
               wPedCli = Trim(wValorComplementoAlfa)
        End If
        'wPedCli = wValorComplemento
    End If
    If wPedCli <> "" Then
       If (IIf(IsNull(wPedCli), 0, wPedCli)) <> 0 Then
           wStr7 = "Ped. Cliente    : " & Trim(wPedCli)
       End If
    End If
    
  
    'Printer.FontSize = 8
    If wPagina = 1 Then
       
        
        WCGC = Right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "###,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        WCGC = Right(String(18, "0") & WCGC, 18)
    End If
 '****************************************************************************
    wStr0 = Space(105) & wPagina & "/" & RsDados("ITV_PAGINANF") 'Inicio Impressão
    Printer.Print wStr0
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 6
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 6
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 6#
    
    If wNovaRazao <> "0" Then
        wStr1 = Space(64) & wNovaRazao
        Printer.Print wStr1
        Printer.Print ""
    Else
        Printer.Print ""
    End If
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
   
        wStr1 = Space(2) & Left(Format(wStr17) & Space(34), 34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & Left(Format(RsDados("ITV_Notafiscal"), "######"), 7)
    
    Printer.Print wStr1
    wStr2 = Space(2) & Left(Format(wStr18) & Space(34), 34) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
    Printer.Print wStr2
   
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(34), 34) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "00000-000"), 9)
   '
    Printer.Print wStr3
   
        'wStr4 = Space(2) & Left(Format(wStr20) & Space(40), 40) & Space(46) & Left(Trim(Format(WCGC, "###,###,###")), 19)
        wStr4 = Space(2) & Left(Format(wStr20) & Space(40), 40) & Space(46) & Left(Trim(WCGC), 19)
   
    Printer.Print wStr4
    Printer.Print ""
    If Trim(Wav) <> "" Then
            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(27) & Left$(wCodigoOperacao, 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = Space(31) & Left(Trim(WNatureza) & Space(26), 26) & Left$(wCodigoOperacao, 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
   '
    Printer.Print wStr5
    
    Printer.Print ""
    Printer.Print ""
   
   
       '===  Pegando Clente
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 6) = True Then
        If wTipodeComplemento = 1 Then
           wPegaCliente = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wPegaCliente = Trim(wValorComplementoNumerico)
        Else
               wPegaCliente = Trim(wValorComplementoAlfa)
        End If
       ' wPegaCliente = wValorComplemento
    End If
    
    
          SQL = ""
          SQL = "Select * from cliente " _
              & "where CE_CodigoCliente=" & wPegaCliente & ""
          
           rdoPegaCliente.CursorLocation = adUseClient
           rdoPegaCliente.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
        If rdoPegaCliente("CE_Tipopessoa") = "F" Or rdoPegaCliente("CE_Tipopessoa") = "U" Then
           wCgcCliente = Right(String(11, "0") & Trim(rdoPegaCliente("CE_CGC")), 11)
           wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 2), "000,000,000") & "-" & Mid(wCgcCliente, 10, 2)
           
        Else
           wCgcCliente = Right(String(14, "0") & Trim(rdoPegaCliente("CE_CGC")), 14)
           wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
           wCgcCliente = Right(String(18, "0") & Trim(wCgcCliente), 18)
        End If
        wStr6 = Left(Trim(wStr6) & Space(31), 31) & Left$(Format(Trim(rdoPegaCliente("CE_CodigoCliente"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(rdoPegaCliente("CE_Razao")), ">") & Space(45), 45) & Left$(Trim(wCgcCliente) & Space(24), 24) & Space(1) & Left$(Format(RsDados("ITV_Data"), "dd/mm/yy") & Space(12), 12)
   
    
    Printer.Print wStr6
    
       '===  Pegando Emite Data de Saida
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 22) = True Then
        If wTipodeComplemento = 1 Then
           wPegaDataSaida = Format(wValorComplementoDate, "yyyy/mm/dd")
        ElseIf wTipodeComplemento = 2 Then
               wPegaDataSaida = Trim(wValorComplementoNumerico)
        Else
               wPegaDataSaida = Trim(wValorComplementoAlfa)
        End If
        
        
    End If
    
  
       wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(rdoPegaCliente("CE_Endereco")), ">") & Space(42), 42) & Left$(Format(Trim(rdoPegaCliente("CE_Bairro")), ">") & Space(21), 21) & Right$(Space(11) & Format(rdoPegaCliente("CE_CEP"), "00000-000"), 11) & Space(7) & Left$(Format(RsDados("ITV_Data"), "dd/mm/yy"), 12)
   
    
    Printer.Print ""
    Printer.Print wStr7
        
        wStr8 = Space(31) & Left$(Format(Trim(rdoPegaCliente("CE_Municipio")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(Format(rdoPegaCliente("CE_Telefone"), "####-####"))) & Space(15), 15) & Left$(Trim(rdoPegaCliente("CE_Estado")), 2) & Space(5) & Left$(Trim(Format(rdoPegaCliente("CE_InscricaoEstadual"), "###,###,###,###")), 15)
    
    Printer.Print ""
    Printer.Print wStr8
    
    Printer.Print ""
    Printer.Print ""
    rdoPegaCliente.Close


           
End Function



Public Function EmiteNotafiscal(ByVal Nota As Double, ByVal Serie As String)
    WNatureza = ""
    wReferenciaEspecial = ""
    wPegaCarimbo1 = ""
    wPegaCarimbo2 = ""
    wPegaCarimbo3 = ""
    wPgentra = 0
    wPegaCliente = ""
    wPegaDesconto = ""
    wPegaCarimbo5 = ""
    wPegaDataSaida = ""
    wPegaFrete = 0
    Glb_ImpNotaFiscal = ""
    wPegaloja = ""
    wPegaVendedorVenda = ""
    wLojaVenda = ""
    wTipoNota = ""
    wDatapag = 0
    wCodigoOperacao = 0
    wPedCli = ""
    wPessoa = 0
    wPegaDescricaoAlternativa = ""
    wTotalVenda = 0
    wTotalnota = 0
    wAnexo = ""
    wAnexo1 = ""
    wAnexo2 = ""
    wCarimbo2 = ""
    wCarimbo5 = ""
    
Dim wControlaQuebraDaPagina As Integer
wControlaQuebraDaPagina = 0
Glb_ImpNotaFiscal = "NOTA FISCAL"
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(Glb_ImpNotaFiscal) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next

    WNF = Nota
    wSerie = Serie
    
    wNotaTransferencia = False
    wPagina = 1
           
    Call DadosLoja
        
    SQL = "Select * From ItensVenda " _
             & "Where ITV_notafiscal = " & Nota & " and ITV_Serie in ('" & Serie & "') " _
             & "  and ITV_Loja ='" & Trim(wLoja) & "' " _
             & " order by ITV_Item"
             RsDados.CursorLocation = adUseClient
             RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    
    
    If Not RsDados.EOF Then
           
          '===  Pegando Tipo Nota
       If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 1) = True Then
          If wTipodeComplemento = 1 Then
             wTipoNota = Format(wValorComplementoDate, "yyyy/mm/dd")
          ElseIf wTipodeComplemento = 2 Then
                 wTipoNota = Trim(wValorComplementoNumerico)
          Else
                 wTipoNota = Trim(wValorComplementoAlfa)
          End If
          '   wTipoNota = Trim(wValorComplementoNumerico)
        ''  End If
          
       End If
       
       If wTipoNota <> "" Then
          Cabecalho Trim(wTipoNota)
       End If
            
      SQL = "Select ProdutoLoja.PRL_referencia,ProdutoLoja.PRL_descricao,ProdutoLoja.PRL_CodigoReducaoICMS, " _
          & "ProdutoLoja.PRL_classefiscal,ProdutoLoja.PRL_unidade,ProdutoLoja.PRL_substituicaotributaria, " _
          & "ProdutoLoja.PRL_icmssaida, " _
          & " ITV_Notafiscal, " _
          & " ITV_Loja, " _
          & " ITV_CodigoProduto, " _
          & " ITV_Quantidade, " _
          & " ITV_PrecoUnitario, " _
          & " ITV_PrecoAlternativo, " _
          & " ITV_DescricaoAlternativa, " _
          & " (ITV_Quantidade * ITV_PrecoUnitario) as wVltotitem, " _
          & " ITV_Situacao, " _
          & " ITV_Tributacao, " _
          & " ITV_detalheImpressao,ITV_item," _
          & " ITV_ReferenciaAlternativa " _
          & "from ProdutoLoja,ItensVenda " _
          & "where ProdutoLoja.PRL_referencia=ITV_CodigoProduto " _
          & "and ITV_Notafiscal = " & Nota & " and ITV_Serie='" & Serie & "' order by ITV_item"

             RsdadosItens.CursorLocation = adUseClient
             RsdadosItens.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

                 
           '===  Pegando Cliente
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 6) = True Then
          If wTipodeComplemento = 1 Then
             wPegaCliente = Format(wValorComplementoDate, "yyyy/mm/dd")
          ElseIf wTipodeComplemento = 2 Then
                 wPegaCliente = Trim(wValorComplementoNumerico)
          Else
                 wPegaCliente = Trim(wValorComplementoAlfa)
          End If
     '     wPegaCliente = Trim(wValorComplementoNumerico)
     ' End If
       
      ' wPegaCliente = Trim(wValorComplementoAlfa)
    End If
    
      
       If wPegaCliente <> "" Then
         wChaveICMSItem = 0
         wPessoa = 0

         SQL = "Select * from Cliente Where CE_CodigoCliente = " & wPegaCliente
                rdocontrole.CursorLocation = adUseClient
                rdocontrole.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
                If rdocontrole("CE_Tipopessoa") = "J" Then
                   wPessoa = 1
                Else
                   wPessoa = 2
                End If
                wUFCliente = rdocontrole("CE_Estado")
                
         SQL = ""
         SQL = "Select UF_Regiao From Estados Where UF_Estado = '" & rdocontrole("CE_Estado") & "'"
                rdoRegiao.CursorLocation = adUseClient
                rdoRegiao.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   
                AuxChaveICMSItem = rdoRegiao("UF_Regiao")
                AuxChaveICMSItem = AuxChaveICMSItem & wPessoa
                
                rdoRegiao.Close
                rdocontrole.Close
       End If
       
      If Not RsdadosItens.EOF Then
         wConta = 0
         Do While Not RsdadosItens.EOF
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = RsdadosItens("PRL_Referencia")
                                 
                   wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("ITV_DescricaoAlternativa")), RsdadosItens("PRL_Descricao"), RsdadosItens("ITV_DescricaoAlternativa"))
                   If wPegaDescricaoAlternativa = "" Then
                        wPegaDescricaoAlternativa = "0"
                   End If
                   If wPegaDescricaoAlternativa <> "0" Then
                         wDescricao = wPegaDescricaoAlternativa
                   Else
                         wDescricao = Trim(RsdadosItens("PRL_descricao"))
                   End If
                   
               
                   If RsdadosItens("PRL_IcmsSaida") = 0 And RsdadosItens("PRL_substituicaotributaria") = "N" Then
                        wCarimbo5 = "S"
                    Else
                        If Trim(wCarimbo5) = "" Then
                           wCarimbo5 = ""
                        End If
                    End If
                   
                   
                  ' If wUFCliente = "SP" Then
                  '     wAliqICMSInterEstadual = RsdadosItens("PRL_ICMSSaida")
                  ' Else
                       
                                       
                       If RsdadosItens("PRL_substituicaotributaria") = "S" Then
                          wSubstituicaoTributaria = 1
                          wCarimbo2 = "S"
                       Else
                          If Trim(wCarimbo2) = "" Then
                             wCarimbo2 = ""
                          End If
                          wSubstituicaoTributaria = 0
                       End If
                      
                       wChaveICMSItem = AuxChaveICMSItem & RsdadosItens("PRL_icmssaida") & RsdadosItens("PRL_codigoreducaoicms") & wSubstituicaoTributaria
                       If AcharICMSInterEstadual(RsdadosItens("PRL_Referencia"), wChaveICMSItem) = True Then
                          'Exit Function
                       End If
                       
                       If GLB_AliquotaICMS <> 0 Then
                          wAliqICMSInterEstadual = GLB_AliquotaICMS
                       Else
                          wAliqICMSInterEstadual = RsdadosItens("PRL_ICMSSaida")
                       End If
                       '''====================================================================
                       
                       wAnexoIten = RsdadosItens("PRL_CodigoReducaoICMS")
                        
                        If wAnexoIten <> 0 Then
                            If wAnexoIten = 1 Then
                                wAnexo1 = RsdadosItens("ITV_Item") & "," & wAnexo1
                            ElseIf wAnexoIten = 2 Then
                                wAnexo2 = RsdadosItens("ITV_Item") & "," & wAnexo2
                            End If
                        End If
                       
                       '''====================================================================
                       
                  
                   'End If
                   
                    '   If RsdadosItens("PRL_IcmsSaida") = 0 And RsdadosItens("PRL_substituicaotributaria") = "N" Then
                    '      wCarimbo5 = "S"
                    '   Else
                    '      wCarimbo5 = ""
                    '   End If
                       'wIE_BasedeReducao
                       'wVLUNIT2 = ITV_Precounitario - ITV_desconto
                       'GLB_BasedeCalculoICMS = Format(RsItensNF("VLUNIT2") - ((RsItensNF("VLUNIT2") * GLB_IE_BasedeReducao) / 100), "0.00")
                       'GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("PRL_referencia") & Space(7), 7) _
                         & Space(1) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
                         & Space(16) & Left$(Format(Trim(RsdadosItens("PRL_classefiscal")), ">") _
                         & Space(12), 12) & Left$(Trim(RsdadosItens("ITV_Tributacao")) & Space(3), 3) _
                         & "" & Space(3) & Left$(Trim(RsdadosItens("PRL_unidade")) & Space(2), 2) _
                         & Right$(Space(6) & Format(RsdadosItens("ITV_Quantidade"), "##0"), 6) _
                         & Right$(Space(12) & Format(RsdadosItens("ITV_PrecoUnitario"), "#####0.00"), 14) _
                         & Right$(Space(15) & Format((RsdadosItens("ITV_PrecoUnitario") * RsdadosItens("ITV_Quantidade")), "#####0.00"), 15) & Space(1) _
                         & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
                  
            
                      Printer.Print wStr16
                      wTotalVenda = (wTotalVenda + (Format((RsdadosItens("ITV_PrecoUnitario") * RsdadosItens("ITV_Quantidade")), "###,##0.00")))
                      If RsdadosItens("ITV_DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("ITV_DetalheImpressao") = "C" Then
                             Do While wConta < 28
                                wConta = wConta + 1
                                Printer.Print ""
                             Loop
                             RsdadosItens.MoveNext
                         
                         wStr13 = Space(78) & "CX 0" & frmControlaCaixa.lblNroCaixa & Space(3) & "Lj " & RsDados("ITV_Loja") & Space(3) & Right$(Space(7) & Format(RsDados("ITV_NotaFiscal"), "###,###"), 7)
                         Printer.Print wStr13
                                                  
                         wConta = 0
                         wPagina = wPagina + 1
                         
 '------------------------------------------------------------------------------
                 'Acerto emissao de nota com mais de um formulario
                       ' Printer.EndDoc
                        
                         'Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         
                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If
'----------------------------------------------------------------------------------
                         Cabecalho Trim(wTipoNota)
                      ElseIf RsdadosItens("ITV_DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                         Call FinalizaNota
                      Else
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      End If
            Loop
            RsdadosItens.Close
         Else
            'Close #Notafiscal
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
  
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If
    RsDados.Close
  
End Function

Private Sub FinalizaNota()
'RsdadosItens.
If wNotaTransferencia = False Then
         If wReferenciaEspecial <> "" Then
             SQL = ""
             SQL = "Select * from CarimbosEspeciais " _
                & "where CE_Referencia='" & wReferenciaEspecial & "'"
                RsPegaItensEspeciais.CursorLocation = adUseClient
                RsPegaItensEspeciais.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
             If Not RsPegaItensEspeciais.EOF Then
                i = 0
        
                If RsPegaItensEspeciais("CE_Linha1") <> "" Then
                    wConta = wConta + 7
                    'Print #Notafiscal, ""
                    If Trim(RsPegaItensEspeciais("CE_Linha5")) = "" Then
                        Printer.Print Space(7) & "______________________________________________________________"
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Printer.Print Space(9) & "___________________________________     ____/____/______   "
                        Printer.Print Space(9) & "            Assinatura                        Data         "
                        'Print #Notafiscal, Space(15) & "____________________________________________________________"
                    Else
                        Printer.Print Space(7) & "______________________________________________________________"
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Printer.Print Space(8) & Right(RsPegaItensEspeciais("CE_Linha5"), 60)
                        Printer.Print Space(9) & "___________________________________     ____/____/______   "
                        Printer.Print Space(9) & "            Assinatura                        Data         "
                        'Print #Notafiscal, Space(15) & "____________________________________________________________"
                    End If


'                    Print #Notafiscal, Space(15) & "_____________________________________________________________"
'                    Print #Notafiscal, Tab(15); "|"; Tab(16); RsPegaItensEspeciais("CE_Linha2"); Tab(76); "|"
'                    Print #Notafiscal, Tab(15); "|"; Tab(16); RsPegaItensEspeciais("CE_Linha3"); Tab(76); "|"
'                    Print #Notafiscal, Tab(15); "|"; Tab(16); RsPegaItensEspeciais("CE_Linha4"); Tab(76); "|"
'                    Print #Notafiscal, Tab(15); "|"; Tab(17); "___________________________________     ____/____/______   |"
'                    Print #Notafiscal, Tab(15); "|"; Tab(17); "            Assinatura                        Data         |"
'                    Print #Notafiscal, Space(14) & "|____________________________________________________________|"
                End If
             End If
           RsPegaItensEspeciais.Close
        End If
End If

'''================================================================================

SQL = "Select CA_Descricao,CA_CodigoCarimbo from CarimboNotaFiscal "
       RsCarimbo.CursorLocation = adUseClient
       RsCarimbo.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        '    Set RsCarimbo = rdoCNLoja.OpenResultset(SQL)
                         
                If Not RsCarimbo.EOF Then
                       Do While Not RsCarimbo.EOF
                           If RsCarimbo("CA_CodigoCarimbo") = 1 Then
                               If wAnexo1 <> "" Or wAnexo2 <> "" Then
                                   wPegaCarimboNF = RsCarimbo("CA_Descricao")
                                   wRecebeCarimboAnexo = Mid(wPegaCarimboNF, 1, 79) & wAnexo1 & Mid(wPegaCarimboNF, 80, Len(wPegaCarimboNF)) & wAnexo2
                               Else
                                   wRecebeCarimboAnexo = ""
                               End If
                           ElseIf RsCarimbo("CA_CodigoCarimbo") = 2 Then
                               If Trim(wCarimbo2) = "S" Then
                                   wCarimbo2 = RsCarimbo("CA_Descricao")
                               Else
                                   wCarimbo2 = ""
                               End If
                           ElseIf RsCarimbo("CA_CodigoCarimbo") = 5 Then
                               'If RsCapaNF("TipoNota") <> "E" Then
                                   If Trim(wCarimbo5) <> "" Then
                                       wCarimbo5 = RsCarimbo("CA_Descricao") & " " & wCarimbo5
                                   Else
                                       wCarimbo5 = ""
                                   End If
                               'End If
                           End If
                           
                           RsCarimbo.MoveNext
                       Loop
                       
                End If
        RsCarimbo.Close
'''================================================================================



     Do While wConta < 7
        wConta = wConta + 1
        Printer.Print ""
     Loop
 
                   
           '===  Pegando Desconto
     If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 15) = True Then
        
          If wTipodeComplemento = 1 Then
             wPegaDesconto = Format(wValorComplementoDate, "yyyy/mm/dd")
          ElseIf wTipodeComplemento = 2 Then
                 wPegaDesconto = Trim(wValorComplementoNumerico)
                 'wPegaDesconto = wValorComplementoAlfa
          Else
                 wPegaDesconto = wValorComplementoAlfa
          End If
        
        'wPegaDesconto = Trim(wValorComplementoNumerico)
     End If
   
    
     If Trim(wRecebeCarimboAnexo) <> "" And wPegaDesconto <> "" Then
        Printer.Print Space(1) & Left(wRecebeCarimboAnexo & Space(106), 106) & Left("Desc." & Space(7), 7) & Left(Format(wPegaDesconto, "0.00") & Space(10), 10)
     ElseIf Trim(wRecebeCarimboAnexo) <> "" Then
        Printer.Print Space(1) & wRecebeCarimboAnexo
     ElseIf wPegaDesconto <> "" Then
        Printer.Print Space(91) & "Desconto" & Space(13) & Format(wPegaDesconto, "0.00")
     Else
        Printer.Print ""
     End If
     
     If wCarimbo2 <> "" Then
        Printer.Print Space(4) & wCarimbo2
        wConta = wConta + 1
     End If
     
     wConta = wConta + 1
     
     If (IIf(IsNull(wCarimbo5), "", wCarimbo5)) <> "" Then
        Printer.Print Space(4) & wCarimbo5
     Else
        Printer.Print ""
     End If
        
     Do While wConta < 14
        wConta = wConta + 1
        Printer.Print ""
     Loop
        
            '===  Pegando Frete
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 16) = True Then
       If wTipodeComplemento = 1 Then
           wPegaFrete = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wPegaFrete = Trim(wValorComplementoNumerico)
       Else
               wPegaFrete = Trim(wValorComplementoAlfa)
       End If
       
      ' wPegaloja = Trim(wValorComplementoAlfa)
    End If
 
             '===  Pegando Base do ICMS
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 21) = True Then
       If wTipodeComplemento = 1 Then
           wBaseIcms = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wBaseIcms = Trim(wValorComplementoNumerico)
       Else
               wBaseIcms = Trim(wValorComplementoAlfa)
       End If
       
      ' wPegaloja = Trim(wValorComplementoAlfa)
    End If
    
              '===  Pegando Valor do ICMS
    If PegarValorComplemento(Val(frmFormaPagamento.txtPedido.Text), 22) = True Then
       If wTipodeComplemento = 1 Then
           wVLRICMS = Format(wValorComplementoDate, "yyyy/mm/dd")
       ElseIf wTipodeComplemento = 2 Then
               wVLRICMS = Trim(wValorComplementoNumerico)
       Else
               wVLRICMS = Trim(wValorComplementoAlfa)
       End If
       
      ' wPegaloja = Trim(wValorComplementoAlfa)
    End If
        
        If Trim(wPegaDesconto) = "" Then
           wPegaDesconto = 0
        End If
        
               
        wTotalnota = wTotalVenda - wPegaDesconto
        wTotalnota = wTotalnota + wPegaFrete
        
        wStr9 = Right$(Space(9) & Format(wBaseIcms, "######0.00"), 9) & Right$(Space(9) & Format(wVLRICMS, "######0.00"), 9) & Space(35) & Right$(Space(10) & Format(wTotalnota, "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = Right(Space(9) & Format(Space(9) & wPegaFrete, "######0.00"), 9) & Space(44) & Right(Space(10) & Format(wTotalnota, "######0.00"), 10)
        Printer.Print wStr10
   
     
     wStr11 = Space(2) & "                          "
     Printer.Print wStr11
     wStr12 = Space(2) & "                                                     "
     Printer.Print wStr12
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(78) & "CX 0" & wNumeroCaixa & Space(3) & "Lj " & RsDados("ITV_Loja") & Space(4) & Right$(Space(7) & Format(RsDados("ITV_Notafiscal"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     
        
     Printer.EndDoc
     
End Sub
Function PegarValorComplemento(ByVal NumeroPedido As String, ByVal SequenciaComplemento As String) As Boolean

 wValorComplementoAlfa = ""
 wValorComplementoNumerico = 0
 wValorComplementoDate = ""
  
 
 SQL = "Select COV_ValorComplemento from ComplementoVenda Where COV_numeroPedido = " _
      & NumeroPedido & " and COV_CodigoComplemento = 1 and COV_SequenciaComplemento = " & SequenciaComplemento
        rdoComplemento.CursorLocation = adUseClient
        rdoComplemento.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
        PegarValorComplemento = False
             
        If Not rdoComplemento.EOF Then
          
          If IsNumeric(rdoComplemento("COV_ValorComplemento")) Then
                 wValorComplementoNumerico = Val(rdoComplemento("COV_ValorComplemento"))
                 wTipodeComplemento = 2
          ElseIf IsDate(rdoComplemento("COV_ValorComplemento")) Then
                 wValorComplementoDate = Format(rdoComplemento("COV_ValorComplemento"), "yyyy/mm/dd")
                 wTipodeComplemento = 1
          Else
                 wValorComplementoAlfa = Trim(rdoComplemento("COV_ValorComplemento"))
                 wTipodeComplemento = 3
          End If
           
           
        PegarValorComplemento = True
        End If
     
     
        rdoComplemento.Close
        
End Function
'Private Sub PesquisaDadosdoCliente()


 ' SQL = "Select COV_ValorComplemento from ComplementoVenda Where COV_numeroPedido = " _
'          & RsdadosItens("ITV_Numeropedido") & " and COV_CodigoComplemento = 1 and COV_SequenciaComplemento = 6"
 '           rdoComplemento.CursorLocation = adUseClient
 '           rdoComplemento.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 
 Function AcharICMSInterEstadual(ByVal Referencia As String, ByVal ChaveIcms As Double) As Boolean
        
    SQL = "SELECT * from IcmsInterEstadual where IE_Codigo = " & ChaveIcms
    RsICMSInter.CursorLocation = adUseClient
    RsICMSInter.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If RsICMSInter.EOF Then
        AcharICMSInterEstadual = False
        MsgBox "ICMS inter estadual da referencia " & Referencia & " não encontrado" & Chr(10) & "A nota não pode ser impressa", vbCritical, "Aviso"
        Exit Function
    Else
        AcharICMSInterEstadual = True
        GLB_AliquotaICMS = RsICMSInter("IE_IcmsDestino")
        GLB_IE_BasedeReducao = RsICMSInter("IE_BasedeReducao")
    End If
    RsICMSInter.Close
        
End Function


Public Function ExtraiSeqNotaControle() As Double
     Dim WnovaSeqNota As Long
     
     SQL = ""
     SQL = "Select CT_SeqNota + 1 as NumNota from controle"
     RsDados.CursorLocation = adUseClient
     RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If Not RsDados.EOF Then
        WNF = RsDados("NumNota")
        ExtraiSeqNotaControle = RsDados("NumNota")
        SQL = "update controle set CT_SeqNota= " & RsDados("NumNota") & ""
        rdoCNLoja.Execute (SQL)
     End If
     RsDados.Close

End Function




