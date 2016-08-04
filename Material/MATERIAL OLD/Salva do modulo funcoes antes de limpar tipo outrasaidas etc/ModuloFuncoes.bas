Attribute VB_Name = "ModuloFuncoes"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Fim
Dim lsDSN As String


Public Function ConectaOdbcLocal(ByRef RdoVar, ByVal Banco As String, ByVal Servidor As String, ByVal Usuario As String, ByVal Senha As String) As Boolean
        
        On Error GoTo ConexaoErro
    
        With RdoVar
            .Connect = "Dsn=" & Trim(Servidor) & ";" _
                    & "Server=" & Trim(Servidor) & ";" _
                    & "DataBase=" & Trim(Banco) & ";" _
                    & "MaxBufferSize=512;" _
                    & "PageTimeout=5;" _
                    & "UID=" & Usuario & ";" _
                    & "PWD=" & Senha & ";"
    
            .LoginTimeout = 10
            .CursorDriver = rdUseClientBatch
            .EstablishConnection rdDriverNoPrompt
        End With
    
        ConectaOdbcLocal = True
        Exit Function
    
ConexaoErro:

    ConectaOdbcLocal = False

End Function

Sub ConectaODBC()
lsDSN = "Driver={Microsoft Access Driver (*.mdb)};" & _
        "Dbq=c:\bdini.mdb;" & _
        "Uid=Admin; Pwd=astap36"

rdoCNLoja.Open lsDSN

SQL = "Select * from ConectaserverLocal"

RsDados.CursorLocation = adUseClient
RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic


With RsDados
 If .BOF And .EOF Then
   MsgBox "Não há dados para exibir ! "
Else
   wNomeservidor = .Fields("NomeServidor")
   wNomeBanco = .Fields("NomeBanco")
   Wusuario = .Fields("Usuario")
   wSenha = .Fields("Senha")
   wNumeroCaixa = .Fields("NumeroCaixa")
   glb_ECF = .Fields("Numero_ECF")
RsDados.Close
rdoCNLoja.Close
End If
End With

'  =========  Conexao  ADO com SQL Server 2000 ========

rdoCNLoja.Provider = "SQLOLEDB"
rdoCNLoja.Properties("Data Source").Value = wNomeservidor
rdoCNLoja.Properties("Initial Catalog").Value = wNomeBanco
rdoCNLoja.Properties("User ID").Value = Wusuario
rdoCNLoja.Properties("Password").Value = wSenha

rdoCNLoja.Open

GLB_ConectouOK = True
Exit Sub
ConexaoErro:
MsgBox "Erro ao abrir banco de localizacao! "

GLB_ConectouOK = False
   
        Exit Sub
            

'ConexaoErro:

    

End Sub

Function PegaSerieNota() As String
    'Dim rdoSerie As rdoResultset
    
    SQL = ""
    SQL = "Select CT_SerieNota from Controle"
    'Set rdoSerie = rdoCNLoja.OpenResultset(SQL)
    
    rdoSerie.CursorLocation = adUseClient
    rdoSerie.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rdoSerie.EOF Then
        PegaSerieNota = rdoSerie("CT_SerieNota")
    End If
    rdoSerie.Close

End Function

Public Function LerDirINI(ByVal Parametro As String, ByVal Chave As String) As String

    Dim Buffer As String * 255
    Dim Tamanho As Long
    Dim Retorno As String

    Tamanho = GetPrivateProfileString(Parametro, Chave, "", Buffer, 255, "TraderCaixa.ini")

    If Tamanho <> 0 Then
        Retorno = Left(Buffer, Tamanho)
    Else
        Retorno = ""
    End If
    
    LerDirINI = Retorno

End Function

Public Function Numeros(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim CharLido As String * 1
    Dim Retorno As String
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        CharLido = Mid(Texto, Char, 1)
        If IsNumeric(CharLido) Then
            Retorno = Retorno & CharLido
        End If
    Next Char
    
    Texto = Retorno
    
    Numeros = Texto

End Function

Public Function ExtraiSeqNotaControle() As Double
     Dim WnovaSeqNota As Long
     
     SQL = ""
     SQL = "Select CT_SeqNota + 1 as NumNota from controle"
     Set RsDados = rdoCNLoja.OpenResultset(SQL)
     
     If Not RsDados.EOF Then
        WNF = RsDados("NumNota")
        ExtraiSeqNotaControle = RsDados("NumNota")
        SQL = "update controle set CT_SeqNota= " & RsDados("NumNota") & ""
        rdoCNLoja.Execute (SQL)
     End If

End Function

Public Function EmiteNotafiscal(ByVal Nota As Double, ByVal Serie As String)
Dim wControlaQuebraDaPagina As Integer
wControlaQuebraDaPagina = 0
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(Glb_ImpNotaFiscal) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next

    WNF = Nota
    Wserie = Serie
    
    wNotaTransferencia = False
    wPagina = 1
    
    Call DadosLoja
            
    SQL = ""
    SQL = "Select NFCAPA.FreteCobr,NFCAPA.Carimbo5,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV,NFCAPA.Carimbo3,NFCAPA.Carimbo2,NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF, " _
        & "NFCAPA.CLIENTE,NFCAPA.FONECLI,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
        & "NFCAPA.UFCLIENTE,NFCAPA.TOTALNOTAALTERNATIVA,NFCAPA.VALORTOTALCODIGOZERO,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NfItens.TipoNota,NfCapa.EmiteDataSaida " _
        & "From NFCAPA,NFITENS " _
        & "Where NfCapa.nf= " & Nota & " and NfCapa.Serie in ('" & Serie & "') " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "' " _
        & "and NfItens.LojaOrigem=NfCapa.LojaOrigem " _
        & "and NfItens.Serie=NfCapa.Serie " _
        & "and NfItens.Nf=NfCapa.NF"
        
    Set RsDados = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsDados.EOF Then
           
      Cabecalho RsDados("TipoNota")
      
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota,NfItens.Tributacao, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.icmpdv,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & Nota & " and Serie='" & Serie & "' order by nfitens.item"

      Set RsdadosItens = rdoCNLoja.OpenResultset(SQL)

      If Not RsdadosItens.EOF Then
         wConta = 0
         Do While Not RsdadosItens.EOF
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = RsdadosItens("PR_Referencia")
            If Wsm = True Then
                 wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), RsdadosItens("PR_Descricao"), RsdadosItens("DescricaoAlternativa"))
                 
                ' If RsDados("UFCliente") = "SP" Then
                '    wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
                ' Else
                '    wAliqICMSInterEstadual = GLB_AliquotaICMS
                ' End If
                
                'wAliqICMSInterEstadual = RsdadosItens("icms")
                 
                 ' Adilson --> Dentro de São Paulo pega do cadastro de produto
                 '             Fora   de São Paulo pega da rotina AcharIcmsInterEstadual
                 
                 If RsDados("UFCliente") = "SP" Then
                    wAliqICMSInterEstadual = RsdadosItens("icms")
                 Else
                    wAliqICMSInterEstadual = RsdadosItens("icmpdv")
                 End If
                 
                 
                 
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("ReferenciaAlternativa") & Space(7), 7) _
                          & Space(1) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
                          & Space(16) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                          & Space(12), 12) & Left$(Trim(RsdadosItens("Tributacao")) & Space(3), 3) _
                          & "" & Space(3) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                          & Right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) _
                          & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "#####0.00"), 14) _
                          & Right$(Space(15) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "#####0.00"), 15) & Space(1) _
                          & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
                          
                          
            
            Else
                     
                   wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), RsdadosItens("PR_Descricao"), RsdadosItens("DescricaoAlternativa"))
                   If wPegaDescricaoAlternativa = "" Then
                        wPegaDescricaoAlternativa = "0"
                   End If
                   If wPegaDescricaoAlternativa <> "0" Then
                         wDescricao = wPegaDescricaoAlternativa
                   Else
                         wDescricao = Trim(RsdadosItens("pr_descricao"))
                   End If
                   
                   'If RsDados("UFCliente") = "SP" Then
                   '    wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
                   'Else
                   '    wAliqICMSInterEstadual = GLB_AliquotaICMS
                   'End If
                   
                   ' Adilson --> Dentro de São Paulo pega do cadastro de produto
                 '             Fora   de São Paulo pega da rotina AcharIcmsInterEstadual
                 
                 If RsDados("UFCliente") = "SP" Then
                    wAliqICMSInterEstadual = RsdadosItens("icms")
                 Else
                    wAliqICMSInterEstadual = RsdadosItens("icmpdv")
                 End If
                 
                 
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("pr_referencia") & Space(7), 7) _
                         & Space(1) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
                         & Space(16) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                         & Space(12), 12) & Left$(Trim(RsdadosItens("Tributacao")) & Space(3), 3) _
                         & "" & Space(3) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                         & Right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) _
                         & Right$(Space(12) & Format(RsdadosItens("vlunit"), "#####0.00"), 14) _
                         & Right$(Space(15) & Format(RsdadosItens("VlTotItem"), "#####0.00"), 15) & Space(1) _
                         & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)

                                  
            End If
                      
                      'On Error Resume Next
                      Printer.Print wStr16
                      'If Err.Number = 52 Then
                        'Close #Notafiscal
                        'Print #Notafiscal, wStr16
                      'End If
                        
                      
                      If RsdadosItens("DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                         Do While wConta < 28
                            wConta = wConta + 1
                            Printer.Print ""
                         Loop
                         RsdadosItens.MoveNext
                         'wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "'"',##'"), 7)
                         wStr13 = Space(78) & "CX 0" & GLB_NumeroCaixa & Space(3) & "Lj " & RsDados("LojaOrigem") & Space(3) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                         Printer.Print wStr13
                         'Printer.Print ""
                         'Printer.Print ""
                         'Printer.Print Chr(18)  'Finaliza Impressão
                         'Close #Notafiscal
                         
                         wConta = 0
                         wPagina = wPagina + 1
                         'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'                         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
 '------------------------------------------------------------------------------
                 'Acerto emissao de nota com mais de um formulario
                       ' Printer.EndDoc
                        
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
                         Cabecalho RsdadosItens("TipoNota")
                      ElseIf RsdadosItens("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                         Call FinalizaNota
                      Else
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      End If
            Loop
         Else
            'Close #Notafiscal
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
        
         'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If




'    Printer.ScaleMode = vbMillimeters
'    Printer.ForeColor = "0"
'    Printer.FontSize = 8
'    Printer.FontName = "draft 10cpi"
'    Printer.FontSize = 8
'    Printer.FontBold = False
'    Printer.DrawWidth = 3
'    Screen.MousePointer = 11
'    wlin = 99
'
'    WNatureza = "VENDAS"
'
'    Call DadosLoja
'
'    SQL = ""
'    SQL = "Select NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
'        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
'        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
'        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
'        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
'        & "NFCAPA.UFCLIENTE,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
'        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
'        & "From NFCAPA INNER JOIN NFITENS " _
'        & "on (NfCapa.nf=Nfitens.nf) " _
'        & "Where NfCapa.nf= " & WNF & " " _
'        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
'
'    Set RsDados = rdocnloja.OpenResultset (SQL)
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'    If Not RsDados.EOF Then
'        If RsDados("CondPag") = 85 Then
'            wCarimbo4 = RsDados("DataPag")
'        Else
'            wCarimbo4 = IIf(IsNull(RsDados("Carimbo4")), "", RsDados("Carimbo4"))
'        End If
'        tmporient = Printer.Orientation
'        wConta = 0
'        wChave = 0
'        wReduz = 0
'        wStr15 = ""
'        wStr17 = ""
'        wStr18 = ""
'        wStr19 = ""
'        wStr20 = ""
'
'        If Val(RsDados("CONDPAG")) = 1 Then
'           Wcondicao = "Avista"
'        ElseIf Val(RsDados("CONDPAG")) = 3 Then
'           Wcondicao = "Financiada"
'        ElseIf Val(RsDados("CONDPAG")) > 3 Then
'           Wcondicao = "Faturada " & wCarimbo4
'        End If
'
'        wStr17 = "Pedido        : " & RsDados("NUMEROPED")
'        wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
'        wStr19 = "Cond. Pagto   : " & Trim(Wcondicao)
'
'        If RsDados("Pgentra") <> 0 Then
'           Wentrada = Format(RsDados("Pgentra"), "'"#####0.00")
'           wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
'        End If
'
'        wStr1 = Space(2) & Left$(Format(wStr17) & Space(50), 50) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(2) & "X" & Space(31) & Left$(Format(RsDados("nf"), "'"',##'"), 7)
'        wStr2 = Space(2) & Left$(Format(wStr18) & Space(50), 50) & Left$(Format(Trim(WMunicipio), ">") & Space(15), 15) & Space(29) & Left$(Trim(westado), 2)
'        wStr3 = Space(2) & Left$(Format(wStr19) & Space(50), 50) & "(011)" & Left$(Trim(Format(WFone, "'"#-###'")), 9) & "/(011)" & Left$(Format(WFone, "'"#-###'"), 9) & Space(11) & Left$(Format((WCep), "'"##-##'"), 9)
'        wStr4 = Space(2) & Left$(Format(wStr19) & Space(100), 100) & Left$(Trim(Format(WCGC, "'"',###,##'")), 10) & "/" & Format(Mid((WCGC), 11, 5), "'"#-#'")
'        wStr5 = Space(44) & Trim(WNatureza) & Space(24) & Left$(RsDados("CFOAUX"), 10) & Space(27) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'        wStr6 = Space(44) & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(17) & Left$(Trim(Format(RsDados("CGCCLI"), "'"',###,##'")), 10) & "/" & Right$(Format(RsDados("CGCCLI"), "'"#-#'"), 7) & Space(5) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr7 = Space(44) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(16) & Left$(RsDados("CEPCLI"), 11) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr8 = Space(44) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(43) & Left$(Trim(RsDados("UFCLIENTE")), 9) & Space(14) & Left$(Trim(Format(RsDados("INSCRICLI"), "'"',###,###,##'")), 15)
'
'
''        wStr6 = Space(40) & Left$(Format(Trim(rdorsExtra2("em_descricao")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(rdorsExtra2("lo_cgc"), "'"',###,##'")), 10) & "/" & Right$(Format(rdorsExtra2("lo_cgc"), "'"#-#'"), 7) & Space(5) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr7 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(15), 15) & Space(32) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr8 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(43) & Left$(Trim(rdorsExtra2("lo_uf")), 9) & Space(14) & Left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "'"',###,###,##'")), 15)
'
'        wStr9 = Space(4) & Right$(Space(12) & Format(RsDados("BaseICMS"), "'"#####0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "'"#####0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("VlrMercadoria"), "'"#####0.00"), 12)
'        wStr10 = Space(67) & Right(Space(12) & Format(RsDados("VlrMercadoria"), "'"#####0.00"), 12)
'        wStr11 = Space(2) & "                          "
'        wStr12 = Space(2) & "                                                     "
'        wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(13) & Right$(Space(7) & Format(RsDados("Nf"), "'"',##'"), 7)
'
'        Printer.ScaleMode = vbMillimeters
'        Printer.ForeColor = "0"
'        Printer.FontSize = 8
'        Printer.FontName = "draft 10cpi"
'        Printer.FontSize = 8
'        Printer.FontBold = False
'        Printer.DrawWidth = 3
'        wpagina = 1
'
'        Call Cabechalho
'
'          SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
'              & "produto.pr_classefiscal,produto.pr_unidade, " _
'              & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
'              & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao " _
'              & "from produto,nfitens " _
'              & "where produto.pr_referencia=nfitens.referencia " _
'              & "and nfitens.nf = " & WNF & ""
'
'          Set RsdadosItens = rdocnloja.OpenResultset (SQL)
'
'          If Not RsdadosItens.EOF Then
'             Do While Not RsdadosItens.EOF
'
'                      wStr16 = ""
'                      wStr16 = Space(6) & Left$(RsdadosItens("pr_referencia") & Space(8), 8) _
'                             & Space(2) & Left$(Format(Trim(RsdadosItens("pr_descricao")), ">") & Space(38), 38) _
'                             & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
'                             & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
'                             & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
'                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "'"##0"), 6) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("vlunit"), "'"#####0.00"), 12) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "'"#####0.00"), 15) & Space(2) _
'                             & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "'0"), 2)
'
'                             Printer.Print wStr16
'
'                   If RsdadosItens("DetalheImpressao") = "D" Then
'                             wConta = wConta + 1
'                             RsdadosItens.MoveNext
'                   ElseIf RsdadosItens("DetalheImpressao") = "C" Then
'                             wConta = 0
'                             RsdadosItens.MoveNext
'                             Printer.NewPage
'                             wpagina = wpagina + 1
'                             Call Cabechalho
'                   ElseIf RsdadosItens("DetalheImpressao") = "T" Then
'                             wConta = wConta + 1
'                             RsdadosItens.MoveNext
'                             Call FinalizaNota
'                   Else
'                             wConta = wConta + 1
'                             RsdadosItens.MoveNext
'                   End If
'             Loop
'          End If
'   Else
'           MsgBox "Impossivel imprimir nota fiscal", vbCritical, "Error"
'           Call Finaliza
'   End If
'
'
'    flg = 0
'    wlin = 99
'    Screen.MousePointer = 0

'    Printer.EndDoc
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
Public Sub ImprimeRomaneio()
   
   For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "CODIGO ZERO" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
   
    'Printer.Print
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 6.5
    Printer.FontName = "draft 10cpi"
    Printer.FontSize = 6.5
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Screen.MousePointer = 11
   
   ValorlItem = 0
   Valordesconto = 0
   SubTotal = 0

   SQL = ("Select * from Lojas Where LO_Loja='" & Trim(wlblloja & "'"))
   Set RsDados = rdoCNLoja.OpenResultset(SQL)
   
   Printer.Print Tab(10); RsDados("LO_Razao")
   Printer.Print ; "CNPJ: " & RsDados("LO_CGC") & " I.E.: " & RsDados("LO_InscricaoEstadual")
   Printer.Print ; RsDados("LO_Endereco")
   Printer.Print ; "TELEFONE: "; RsDados("LO_Telefone")
   Printer.Print ; Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(16) & Format(NroNotaFiscal, "###000")
   Printer.Print "========================================"
   Printer.Print "DESCRICAO DO PRODUTO                    "
   Printer.Print "CODIGO  PRODUTO  QTDxUNIT.   VALOR TOTAL"
   Printer.Print "________________________________________"
   RsDados.Close
   
  SQL = "Select * From Nfcapa Where  NF = " & NroNotaFiscal & " and Serie='00'"
             
              Set RsDadosCapa = rdoCNLoja.OpenResultset(SQL)
             
             If Not RsDadosCapa.EOF Then
                wPegaDesconto = RsDadosCapa("Desconto")
               ' wPegaFrete = RsDadosCapa("ValFrete")
             End If
   
   SQL = "Select * from Nfitens " _
       & "Where  NF = " & NroNotaFiscal & " and Serie='00'"
       
       Set RsDados = rdoCNLoja.OpenResultset(SQL)
       
       If Not RsDados.EOF Then
          Do While Not RsDados.EOF
             SQL = "Select PR_Descricao from Produto Where PR_Referencia ='" & RsDados("Referencia") & "'"
             Set rdoProduto = rdoCNLoja.OpenResultset(SQL)
             ValorlItem = (RsDados("vlunit") * RsDados("Qtde"))
             SubTotal = (SubTotal + ValorlItem)
            ' ValorDesconto = RsDados("ITV_Desconto")
             Printer.Print rdoProduto("PR_Descricao")
             Printer.Print RsDados("referencia") _
             & Right(Space(4) & Format(RsDados("Qtde"), "###0"), 4) & "x" _
             & Format(RsDados("vlunit"), "###,###,###.00") & Space(5) _
             & Right(Space(10) & Format(ValorlItem, "###,###,###.00"), 14)
             rdoProduto.Close
             RsDados.MoveNext
          Loop
       End If
       RsDados.Close
       TotalVenda = (SubTotal - Valordesconto)
       Printer.Print ""
       Printer.Print "SUB TOTAL " & Right(Space(10) & Format(SubTotal, "###,###,##0.00"), 14)
       Printer.Print ""
       Printer.Print "DESCONTO  " & Right(Space(10) & Format(Valordesconto, "###,###,##0.00"), 14)
       Printer.Print ""
       Printer.Print "TOTAL     " & Right(Space(10) & Format(TotalVenda, "###,###,##0.00"), 14)
       Printer.Print ""
       Printer.Print "________________________________________"
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       
       Printer.EndDoc

End Sub
Function EncerraVenda(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double) As Boolean
    Dim SerieProd As String
            
        wVerificaTM = False
        wQuantdadeTotalItem = 0
        wAnexo = ""
        wAnexo1 = ""
        wAnexo2 = ""
        wQuantItensCapaNF = 0
        wCFO2 = " "
        wCFO1 = " "
        wChaveICMS = 0
        GLB_TotalICMSCalculado = 0
        GLB_ValorCalculadoICMS = 0
        GLB_BasedeCalculoICMS = 0
        GLB_AliquotaAplicadaICMS = 0
        GLB_AliquotaICMS = 0
        GLB_BaseTotalICMS = 0
        GLB_Tributacao = 0
        wCFOItem = 0
        wUltimoItem = 0
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wCarimbo5 = ""
        wCarimbo2 = ""
        EncerraVenda = True
        SerieProd = ""
        wRecebeCarimboAnexo = ""
        

'  --------------------------------- CALCULO DO ICMS ------------------------------------------------------------------------
'
        If ConsistenciaNota(NumeroDocumento, SerieDocumento) = False Then
            EncerraVenda = False
            Exit Function
        End If
        SQL = "Select nfcapa.*, Estados.* from nfcapa, Estados " _
              & "where nfcapa.numeroped = " & NumeroDocumento & "" _
              & "And nfcapa.ufCliente = Estados.UF_Estado"
              Set RsCapaNF = rdoCNLoja.OpenResultset(SQL)
        
        If Not RsCapaNF.EOF Then
               wConfereCodigoZero = IIf(IsNull(RsCapaNF("ValorTotalCodigoZero")), 0, RsCapaNF("ValorTotalCodigoZero"))
               If Trim(RsCapaNF("TipoNota")) = "E" Or Trim(RsCapaNF("TipoNota")) = "T" Or Trim(RsCapaNF("TipoNota")) = "S" Then
                    wPessoa = RsCapaNF("PessoaCli")
               End If
               'If RsCapaNF("TM") <> 1 Then
                    wECFNF = 2
                    wNumeroECF = 2
                    wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
                    If RsCapaNF("UFCliente") = "SP" Then
                        If wPessoa = 2 Then
                           If WEmiteCupom = True Then
                              wECFNF = 1
                              wNumeroECF = glb_ECF
                           Else
                              wNumeroECF = glb_ECF
                              wECFNF = 2
                           End If
                        End If
                    End If
               'End If
               If RsCapaNF("Serie") <> "S1" And RsCapaNF("Serie") <> "D1" Then
                    Wserie = ""
                    WNF = ""
               Else
                    Wserie = IIf(IsNull(RsCapaNF("Serie")), "", RsCapaNF("Serie"))
                    'wECFNF = 2
                    WNF = IIf(IsNull(RsCapaNF("NF")), "", RsCapaNF("NF"))
                    'WEmiteCupom = False
               End If
               If RsCapaNF("CondPag") >= 3 Or RsCapaNF("Cliente") <> "999999" Then
                    wECFNF = 2
               End If
        Else
            MsgBox "Nota não encontrada", vbInformation, "Atenção"
            Exit Function
        End If
                    
        SQL = "Select produto.*, nfitens.* from produto,nfitens " _
              & "where nfitens.numeroped = " & NumeroDocumento & "" _
              & " and pr_referencia = nfitens.referencia order by NfItens.Item"
              Set RsItensNF = rdoCNLoja.OpenResultset(SQL)
          
          If Not RsItensNF.EOF Then
             Do While Not RsItensNF.EOF
               'If RsCapaNF("TM") <> 1 Then
                     wChaveICMSItem = wChaveICMS
                  If Trim(wCarimbo5) = "" Then
                    If RsItensNF("PR_IcmsSaida") = 0 And RsItensNF("PR_substituicaotributaria") = "N" Then
                        wCarimbo5 = "S"
                    Else
                        wCarimbo5 = ""
                    End If
                  End If
                  
                  ' If Trim(wCarimbo2) = "" Then
                 '   If RsItensNF("PR_substituicaotributaria") = "S" Then
                 '       wSubstituicaoTributaria = 1
                 '       wCarimbo2 = "S"
                 '   Else
                 '       wCarimbo2 = ""
                 '       wSubstituicaoTributaria = 0
                 '   End If
                 ' End If
                   
                    If RsItensNF("PR_substituicaotributaria") = "S" Then
                        wSubstituicaoTributaria = 1
                        wCarimbo2 = "S"
                    Else
                       ' wCarimbo2 = ""
                        wSubstituicaoTributaria = 0
                    End If
                 ' If Trim(wCarimbo2) = "" Then
                 ' End If
                 
                    wChaveICMSItem = wChaveICMSItem & RsItensNF("pr_icmssaida") & RsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                    If AcharICMSInterEstadual(RsItensNF("PR_Referencia"), wChaveICMSItem) = False Then
                        EncerraVenda = False
                        Exit Function
                    End If
                    
                    If wOutraSaida = True Then
                        If frmFechaPedido.fraOutraSaida.Visible = True And RsItensNF("TipoNota") = "PD" And frmFechaPedido.chkCalcImposto.Value = False Then
                            wCarimbo5 = "NAO INCIDENCIA DO ICMS CONF. ART 7º DO INCISO X DO RICMS"
                        
                            If RsCapaNF("TipoNota") <> "E" And RsCapaNF("TipoNota") <> "T" And RsCapaNF("TipoNota") <> "S" Then
                                If wCFOItem = 5102 Or wCFOItem = 6102 Then
                                    wCFO1 = wCFOItem
                                ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                                    wCFO2 = wCFOItem
                                    If Trim(wCFO2) = 6405 Then
                                       wCFO2 = 6404
                                    End If
                                End If
                            ElseIf RsCapaNF("TipoNota") = "T" Then
                                   If RsItensNF("PR_substituicaotributaria") = "S" Then
                                      wCFO2 = 5409
                                   Else
                                      wCFO1 = 5152
                                   End If
                                  ' wCFO1 = 5152
                                  ' If RsItensNF("PR_substituicaotributaria") = "S" Then
                                  '    If wCFO2 = 5409 Then
                                  '    Else
                                  '       wCFO2 = 5409
                                  '    End If
                                  ' Else
                                  ' wCFO2 = " "
                                  ' End If
                            Else
                                    If RsCapaNF("TipoNota") <> "S" Then
                                        If RsCapaNF("UFCliente") = "SP" Then
                                            wCFO1 = 1202 'Devolucao dentro estado
                                        Else
                                            wCFO1 = 2202 'Devolucao p/ fora do estado
                                        End If
                                    End If
                            End If
                            If Trim(wCFO1) = "" And Trim(wCFO2) = "" And RsCapaNF("TipoNota") <> "S" Then
                                wCFO1 = wCFOItem
                            ElseIf RsCapaNF("TipoNota") = "S" Then
                                wCFO1 = frmFechaPedido.mskCFOP.Text
                            End If
                        End If
                    Else
                        GLB_AliquotaAplicadaICMS = RsICMSInter("IE_icmsAplicado")
                        
                       ' If RsCapaNF("UFCliente") <> "SP" Then
                           'If GLB_AliquotaICMS > 0 Then
                           'Else
                       '      GLB_AliquotaICMS = RsICMSInter("IE_IcmsDestino")
                          ' End If
                       ' Else
                       '    GLB_AliquotaICMS = RsItensNF("ICMS")
                       ' End If
                        
                        GLB_Tributacao = RsICMSInter("IE_Tributacao")
                        wCFOItem = RsICMSInter("IE_Cfo")
                        
                        wAnexoIten = RsItensNF("PR_CodigoReducaoICMS")
                        
                        If wAnexoIten <> 0 Then
                            If wAnexoIten = 1 Then
                                wAnexo1 = RsItensNF("Item") & "," & wAnexo1
                            ElseIf wAnexoIten = 2 Then
                                wAnexo2 = RsItensNF("Item") & "," & wAnexo2
                            End If
                        End If
                            
                        If wConfereCodigoZero > 0 Then
                            GLB_ValorCalculadoICMS = Format((((RsItensNF("ValorMercadoriaAlternativa") - (RsItensNF("VlTotItem") - RsItensNF("VlUnit2"))) * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                            GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
                            If GLB_TotalICMSCalculado > 0 Then
                                If RsICMSInter("IE_BasedeReducao") = 0 Then
                                    If GLB_AliquotaAplicadaICMS = 0 Then
                                        GLB_BasedeCalculoICMS = 0
                                    Else
                                        GLB_BasedeCalculoICMS = RsItensNF("ValorMercadoriaAlternativa") - (RsItensNF("VlTotItem") - RsItensNF("VlUnit2"))
                                    End If
                                Else
                                    GLB_BasedeCalculoICMS = Format((RsItensNF("ValorMercadoriaAlternativa") - (RsItensNF("VlTotItem") - RsItensNF("VlUnit2"))) - (((RsItensNF("ValorMercadoriaAlternativa") - (RsItensNF("VlTotItem") - RsItensNF("VlUnit2"))) * RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
                                End If
                                GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
                            End If
                        Else
                            GLB_ValorCalculadoICMS = Format(((RsItensNF("VLUNIT2") * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                            GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
                            If GLB_TotalICMSCalculado > 0 Then
                                If RsICMSInter("IE_BasedeReducao") = 0 Then
                                    If GLB_AliquotaAplicadaICMS = 0 Then
                                        GLB_BasedeCalculoICMS = 0
                                    Else
                                        GLB_BasedeCalculoICMS = RsItensNF("VLUNIT2")
                                    End If
                                Else
                                    GLB_BasedeCalculoICMS = Format(RsItensNF("VLUNIT2") - ((RsItensNF("VLUNIT2") * RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
                                End If
                                GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
                            End If
                        End If
                        WAnexoAux = ""
                        If RsItensNF("pr_codigoreducaoicms") <> 0 Then
                            WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "0")
                        End If
                        
                        If RsCapaNF("TipoNota") <> "E" And RsCapaNF("TipoNota") <> "T" And RsCapaNF("TipoNota") <> "S" Then
                            If wCFOItem = 5102 Or wCFOItem = 6102 Then
                                wCFO1 = wCFOItem
                            ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                                wCFO2 = wCFOItem
                                If Trim(wCFO2) = 6405 Then
                                   wCFO2 = 6404
                                End If
                            End If
                        ElseIf RsCapaNF("TipoNota") = "T" Then
                               
                               If RsItensNF("PR_substituicaotributaria") = "S" Then
                                  wCFO2 = 5409
                               Else
                                  wCFO1 = 5152
                               End If
                        Else
    '                        If wCFOItem = 5102 Or wCFOItem = 6102 Then
    '                            wCFO1 = wCFOItem
                                If RsCapaNF("TipoNota") <> "S" Then
                                    If RsCapaNF("UFCliente") = "SP" Then
                                        wCFO1 = 1202 'Devolucao dentro estado
                                    Else
                                        wCFO1 = 2202 'Devolucao p/ fora do estado
                                    End If
                                End If
    '                        ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
    '                            wCFO2 = wCFOItem
    '                            If wCFO2 = 5405 Then
    '                                wCFO2 = 1202 'Devolucao dentro estado
    '                            Else
    '                                wCFO2 = 2202 'Devolucao p/ fora do estado
    '                            End If
    '                        End If
                        End If
                        If Trim(wCFO1) = "" And Trim(wCFO2) = "" And RsCapaNF("TipoNota") <> "S" Then
                            wCFO1 = wCFOItem
                        ElseIf RsCapaNF("TipoNota") = "S" Then
                            wCFO1 = frmFechaPedido.mskCFOP.Text
                        End If
                    End If
               'Else
               '     wVerificaTM = True
               'End If
' -------------------------------------- CALCULO DA COMISSÃO DE VENDA ---------------------------------------------------
'
                wComissaoVenda = (RsItensNF("VLUNIT2") * RsItensNF("pr_percentualcomissao") / 100)
                If RsItensNF("TipoNota") = "E" Then
                    wComissaoVenda = wComissaoVenda * -1
                End If
                        
'
' -------------------------------------- CALCULO DA MARGEM DE VENDA ---------------------------------------------------
'
                wSomaVenda = wSomaVenda + RsItensNF("VLUNIT2")
                wSomaMargem = wSomaMargem + (RsItensNF("VLUNIT2") - (RsItensNF("pr_precocusto1") * RsItensNF("qtde")))


'
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'
            '    If RsCapaNF("TM") <> 1 Then
                    wQuantItensCapaNF = RsCapaNF("QtdItem")
                    wQuantItensNF = RsItensNF("Item")
                    wQuantdadeTotalItem = wQuantdadeTotalItem + 1
                    If RsItensNF("TipoNota") <> "E" Then
                        wQuant = (wQuantItensNF Mod 8)
                        If wQuant <> 0 Then
                            wDetalheImpressao = "D"
                        Else
                            wDetalheImpressao = "C"
                            If wQuantItensCapaNF > wQuantItensNF Then
                                wUltimoItem = wUltimoItem + 1
                            End If
                        End If
                        
                        If wQuantItensCapaNF = wQuantItensNF Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        End If
                    Else
                        'Quebra Nota Devolução
                        wQuant = (wQuantItensNF Mod 6)
                        If wQuant <> 0 Then
                            wDetalheImpressao = "D"
                        Else
                            wDetalheImpressao = "C"
                            wUltimoItem = wUltimoItem + 1
                        End If
                                        
                        If wQuantItensCapaNF = wQuantItensNF Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        End If
                    End If
                    If RsItensNF("TipoNota") <> "E" Then
                        If (IIf(IsNull(RsItensNF("SerieProd1")), "", RsItensNF("SerieProd1"))) <> "" Then
                            If (IIf(IsNull(RsItensNF("SerieProd2")), "", RsItensNF("SerieProd2"))) <> "" Then
                                If RsItensNF("SerieProd2") <> "0" Then
                                    SerieProd = Trim(SerieProd) & Trim(RsItensNF("Item")) & "-" & Trim(RsItensNF("SerieProd1")) & "/" & Trim(RsItensNF("SerieProd2")) & ","
                                End If
                            Else
                                If RsItensNF("SerieProd1") <> "0" Then
                                    SerieProd = Trim(SerieProd) & Trim(RsItensNF("Item")) & "-" & Trim(RsItensNF("SerieProd1")) & ","
                                End If
                            End If
                        End If
                    Else
                        wCarimbo5 = Trim(IIf(IsNull(RsCapaNF("Carimbo5")), "", RsCapaNF("Carimbo5")))
                    End If
                    
                    SQL = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & ", " _
                    & "Comissao = " & ConverteVirgula(Format(wComissaoVenda, "0.00")) & ", " _
                    & "IcmPDV = " & ConverteVirgula(Format(RsICMSInter("IE_icmsdestino"), "0.00")) & ", " _
                    & "Tributacao = '" & RsICMSInter("IE_Tributacao") & "', " _
                    & "Bcomis = " & ConverteVirgula(RsItensNF("PR_PercentualComissao")) & " , " _
                    & "DetalheImpressao = '" & wDetalheImpressao & "' " _
                    & " where nfitens.numeroped = " & NumeroDocumento & "" _
                    & " and Referencia = '" & RsItensNF("PR_Referencia") & "' and Item=" & RsItensNF("Item") & ""
                    rdoCNLoja.Execute (SQL)
                    
                    
                    'wUltimoItem = RsItensNF("Item")
                'End If
                    
                    SQL = ""
                    SQL = "Update nfitens set Comissao = " & ConverteVirgulaNegativa(Format(wComissaoVenda, "0.00")) & " " _
                        & " where nfitens.numeroped = " & NumeroDocumento & "" _
                        & " and Referencia = '" & RsItensNF("PR_Referencia") & "'"
                        rdoCNLoja.Execute (SQL)
                
                RsItensNF.MoveNext
             Loop
        End If
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
'
        'If RsCapaNF("TM") <> 1 Then
            If Trim(wCarimbo5) <> "S" And Trim(SerieProd) <> "" And RsCapaNF("TipoNota") <> "E" Then
                SerieProd = "SERIE(S) ITEM(S):" & SerieProd
            End If
            SQL = "Select CA_Descricao,CA_CodigoCarimbo from CarimboNotaFiscal "
            Set RsCarimbo = rdoCNLoja.OpenResultset(SQL)
                
            If Trim(wCarimbo5) <> "NAO INCIDENCIA DO ICMS CONF. ART 7º DO INCISO X DO RICMS" Then
                
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
                               If RsCapaNF("TipoNota") <> "E" Then
                                   If Trim(wCarimbo5) <> "" Then
                                       'Segundo Ailton não é mais necessário emitir este carimbo
                                       'wCarimbo5 = RsCarimbo("CA_Descricao") & " " & wCarimbo5
                                       wCarimbo5 = ""
                                   Else
                                       wCarimbo5 = ""
                                   End If
                               End If
                           End If
                           
                           RsCarimbo.MoveNext
                       Loop
                End If
            End If
            'wUltimoItem = ((wUltimoItem / 10) + 0.9)
             
             SQL = "UPDATE nfcapa set baseicms = " & ConverteVirgula(GLB_BaseTotalICMS) & ", " _
                & "Vlricms = " & ConverteVirgula(GLB_TotalICMSCalculado) & ", " _
                & "cfoaux = '" & Trim(wCFO1 & wCFO2) & "', " _
                & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", " _
                & "Pessoacli = " & wPessoa & ", " _
                & "ECFNF = " & glb_ECF & ", " _
                & "Regiaocli = " & RsCapaNF("UF_Regiao") & ", " _
                & "Carimbo1 = '" & Trim(wRecebeCarimboAnexo) & "', " _
                & "Carimbo5 = '" & Trim(wCarimbo5) & Trim(SerieProd) & "', " _
                & "Carimbo2= '" & Trim(wCarimbo2) & "' " _
                & "where nfcapa.numeroped = " & NumeroDocumento & ""
                rdoCNLoja.Execute (SQL)
                            

                'GravaSequenciaLeitura 5, NumeroDocumento, SerieDocumento
        'End If
' -------------------------------------- ATUALIZA MARGEM DE VENDA ---------------------------------------------------
'
                SQL = "UPDATE vende set VE_totalvenda = VE_totalVenda + " & ConverteVirgula(wSomaVenda) & ", " _
                & "VE_MargemVenda = VE_MargemVenda + " & ConverteVirgula(wSomaMargem) & " " _
                & "where VE_Codigo = " & RsCapaNF("Vendedor") & " "
                rdoCNLoja.Execute (SQL)


'
' -------------------------------------- ATUALIZA CONTORLE DE OPERAÇÂO ---------------------------------------------------
'

                SQL = "UPDATE CTcaixa set Ct_operacoes = Ct_operacoes + 1 " _
                & "where ct_situacao = 'A' "
                rdoCNLoja.Execute (SQL)
          
     
            
    
End Function
Sub DadosECF()


          Screen.MousePointer = 11

        
        SQL = "Select * from ControleSistema "
        
        rdocontrole.CursorLocation = adUseClient
        rdocontrole.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        
        NroNotaFiscal = rdocontrole("CTS_Numero00") + 1
        rdocontrole.Close
         
        
        
        rdoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        
        SQL = "Update ControleSistema set CTS_Numero00 =" & NroNotaFiscal
        
        rdoCNLoja.Execute SQL
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
     
            
         rdoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          
          SQL = "Update nfcapa set NF = " & NroNotaFiscal _
              & ", Serie = '00' Where NumeroPed = " & frmCaixaNF.txtPedido.Text _
              & " and tiponota = 'PA'"
          
          rdoCNLoja.Execute SQL
          Screen.MousePointer = vbNormal
          rdoCNLoja.CommitTrans
          
          rdoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          
          SQL = "Update nfitens set NF = " & NroNotaFiscal _
              & ", Serie = '00' Where NumeroPed = " & frmCaixaNF.txtPedido.Text _
              & " and tiponota = 'PA'"
          
          rdoCNLoja.Execute SQL
          Screen.MousePointer = vbNormal
          rdoCNLoja.CommitTrans
End Sub



Function AcharICMSInterEstadual(ByVal Referencia As String, ByVal ChaveIcms As Double) As Boolean
    
    SQL = "SELECT * from IcmsInterEstadual where IE_Codigo = " & ChaveIcms
    Set RsICMSInter = rdoCNLoja.OpenResultset(SQL)
    
    If RsICMSInter.EOF Then
        AcharICMSInterEstadual = False
        MsgBox "ICMS inter estadual da referencia " & Referencia & " não encontrado" & Chr(10) & "A nota não pode ser impressa", vbCritical, "Aviso"
        Exit Function
    Else
        AcharICMSInterEstadual = True
    End If
        
End Function


Function ConsistenciaNota(ByVal Pedido As Double, ByVal Serie As String) As Boolean
    Dim rsItemNota As rdoResultset
    
    SQL = ""
    SQL = "Select count(NfItens.Referencia) as QuantRef, NfCapa.QtdItem from NfCapa,NfItens " _
        & "where NfCapa.NumeroPed=" & Pedido & " " _
        & "and NfItens.NumeroPed=NfCapa.NumeroPed " _
        & "Group by NfCapa.QtdItem " _
        & "having Count(NfItens.Referencia) = NfCapa.QtdItem"
    Set rsItemNota = rdoCNLoja.OpenResultset(SQL)
    If Not rsItemNota.EOF Then
        ConsistenciaNota = True
    Else
        MsgBox "A nota não pode ser impressa porque exite um erro com a quantidade de itens ", vbCritical, "Atenção"
        ConsistenciaNota = False
    End If

End Function



Sub Esperar(ByVal Tempo As Integer)
    
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
        DoEvents
    Loop

End Sub


Function FU_ValidaCPF(CPF As String) As Integer
'
    Dim soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    
    'Valida argumento
    If Len(CPF) <> 11 Then
        FU_ValidaCPF = False
        Exit Function
    End If

        
    
    soma = 0
    For i = 1 To 9
        soma = soma + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
        
    soma = 0
    For i = 1 To 10
        soma = soma + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
    
    FU_ValidaCPF = True

End Function

Function FU_ValidaCGC(CGC As String) As Integer
        Dim Retorno, a, j, i, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For i = 1 To 7
               a = Val(Mid(CGC, i, 1))
               If (i Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next i
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              FU_ValidaCGC = True
           Else
              FU_ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              i = 0
              d1 = 0
              d2 = 0
              j = 5
              For i = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              i = 0
              j = 6
              For i = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d2 = IIf(a > 1, 11 - a, 0)
              If (d1 = Val(Mid(CGC, 13, 1)) And d2 = Val(Mid(CGC, 14, 1))) Then
                 FU_ValidaCGC = True
              Else
                 FU_ValidaCGC = False
              End If
           Else
              FU_ValidaCGC = False
           End If
        End If
End Function





