Attribute VB_Name = "ModuloFuncoes"
'Global lsDSN As String
Dim wControlaQuebraDaPagina As Integer
Dim wContaItem As Integer
Dim rsNFELoja As New ADODB.Recordset
Dim rsNFECapa As New ADODB.Recordset
Dim RSControleImpostos As New ADODB.Recordset
Dim Sql As String
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Fim

Dim wTNDinheiro As Double
Dim wTNCheque As Double
Dim wTNVisa As Double
Dim wTNRedecard As Double
Dim wTNBNDES As Double
Dim wTNAmex As Double
Dim wTNHiperCard As Double
Dim wTNVisaEletron As Double
Dim wTNRedeShop As Double
Dim wTNDeposito As Double
Dim wTNNotaCredito As Double
Dim wTNConducao As Double
Dim wTNDespLoja As Double
Dim wTNOutros As Double
Dim wTNTotal As Double

Dim wPICMSFECP As Double
Dim wVICMSFECP As Double
Dim wTotalVICMSFECP As Double




Sub ConectaODBC()

On Error GoTo ConexaoErro:

If rdoCNLoja.State = 1 Then
    rdoCNLoja.Close
End If


rdoCNLoja.Provider = "SQLOLEDB"
rdoCNLoja.Properties("Data Source").Value = GLB_Servidorlocal
rdoCNLoja.Properties("Initial Catalog").Value = Glb_BancoLocal
'rdoCNLoja.Properties("User ID").Value = GLB_Usuario
'rdoCNLoja.Properties("Password").Value = GLB_Senha


If ConexaoDLLAdo.abrirConexaoADO(rdoCNLoja, GLB_Servidorlocal, Glb_BancoLocal) Then
    GLB_ConectouOK = True
    Exit Sub
End If

'rdoCNLoja.Open

ConexaoErro:
MsgBox "Erro ao abrir banco de localizacao! "

GLB_ConectouOK = False
  
Exit Sub

End Sub

Sub ConectaODBCMatriz()

'On Error GoTo ConexaoErro:

'rdoCNRetaguarda.Provider = "SQLOLEDB"
'rdoCNRetaguarda.Properties("Data Source").Value = GLB_Servidor
'rdoCNRetaguarda.Properties("Initial Catalog").Value = GLB_Banco
'rdoCNRetaguarda.Properties("User ID").Value = GLB_Usuario
'rdoCNRetaguarda.Properties("Password").Value = GLB_Senha

'rdoCNRetaguarda.Open

'wConectouRetaguarda = True
'Exit Sub
'ConexaoErro:
'wConectouRetaguarda = False
'Exit Sub

On Error GoTo ConexaoErro:

If rdoCNRetaguarda.State <> 1 Then

    If Not ConexaoDLLAdo.abrirConexaoADO(rdoCNRetaguarda, GLB_Servidor, GLB_Banco) Then
        Exit Sub
    End If
    
End If

wConectouRetaguarda = True

Exit Sub
ConexaoErro:
wConectouRetaguarda = False

End Sub

Sub ConectaODBCTEF()

On Error GoTo ConexaoErro:

If rdoCNTEF.State = 1 Then
    rdoCNTEF.Close
End If

rdoCNTEF.Provider = "SQLOLEDB"
rdoCNTEF.Properties("Data Source").Value = GLB_ServidorTEF
rdoCNTEF.Properties("Initial Catalog").Value = GLB_BancoTEF
rdoCNTEF.Properties("User ID").Value = GLB_UsuarioTEF
rdoCNTEF.Properties("Password").Value = GLB_SenhaTEF

rdoCNTEF.Open

GLB_ConectouOK = True
Exit Sub
ConexaoErro:
MsgBox "Erro ao abrir banco de localizacao! "

GLB_ConectouOK = False
  
Exit Sub

End Sub
Function Cliptografia(ByRef ValorClipt As String)

    Dim Ret As String
    Dim CharLido As String
    Dim Maximo As Long
    Dim I As Long

    
    Ret = ""
    Maximo = Len(ValorClipt)
    
    For I = 1 To Maximo
        CharLido = UCase(Mid(ValorClipt, I, 1))
        If CharLido = "A" Then
            CharLido = "E"
        ElseIf CharLido = "B" Then
            CharLido = "F"
        ElseIf CharLido = "C" Then
            CharLido = "G"
        ElseIf CharLido = "D" Then
            CharLido = "H"
        ElseIf CharLido = "E" Then
            CharLido = "I"
        ElseIf CharLido = "F" Then
            CharLido = "J"
        ElseIf CharLido = "G" Then
            CharLido = "L"
        ElseIf CharLido = "H" Then
            CharLido = "M"
        ElseIf CharLido = "I" Then
            CharLido = "N"
        ElseIf CharLido = "J" Then
            CharLido = "O"
        ElseIf CharLido = "L" Then
            CharLido = "P"
        ElseIf CharLido = "M" Then
            CharLido = "Q"
        ElseIf CharLido = "N" Then
            CharLido = "R"
        ElseIf CharLido = "O" Then
            CharLido = "S"
        ElseIf CharLido = "P" Then
            CharLido = "T"
        ElseIf CharLido = "Q" Then
            CharLido = "U"
        ElseIf CharLido = "R" Then
            CharLido = "V"
        ElseIf CharLido = "S" Then
            CharLido = "X"
        ElseIf CharLido = "T" Then
            CharLido = "Z"
        ElseIf CharLido = "U" Then
            CharLido = "K"
        ElseIf CharLido = "V" Then
            CharLido = "W"
        ElseIf CharLido = "X" Then
            CharLido = "Y"
        ElseIf CharLido = "Z" Then
            CharLido = "A"
        ElseIf CharLido = "W" Then
            CharLido = "B"
        ElseIf CharLido = "K" Then
            CharLido = "C"
        ElseIf CharLido = "Y" Then
            CharLido = "D"
        ElseIf CharLido = "1" Then
            CharLido = "6"
        ElseIf CharLido = "2" Then
            CharLido = "5"
        ElseIf CharLido = "3" Then
            CharLido = "7"
        ElseIf CharLido = "4" Then
            CharLido = "8"
        ElseIf CharLido = "5" Then
            CharLido = "9"
        ElseIf CharLido = "6" Then
            CharLido = "0"
        ElseIf CharLido = "7" Then
            CharLido = "1"
        ElseIf CharLido = "8" Then
            CharLido = "3"
        ElseIf CharLido = "9" Then
            CharLido = "2"
        ElseIf CharLido = "0" Then
            CharLido = "4"
        End If
        Ret = Ret & CharLido
    Next
    Cliptografia = Ret

End Function





Function ConsistePedido(ByRef pedido As Double)

    
    Sql = ""
    Sql = "Select Sum(VlUnit * qtde) as TotItem " _
        & "From NfItens " _
        & "where NumeroPed = " & pedido & " "
       
        rdoVlItem.CursorLocation = adUseClient
        rdoVlItem.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Sql = ""
    Sql = "Select TotalNota,FreteCobr,desconto " _
        & "From NfCapa " _
        & "where NumeroPed = " & pedido & " "
       
        rdoVlNota.CursorLocation = adUseClient
        rdoVlNota.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
   
    
    If (Not rdoVlNota.EOF) And (Not rdoVlItem.EOF) Then
        If Format(rdoVlItem("TotItem") + rdoVlNota("FreteCobr"), "#,##0.00") <> Format(rdoVlNota("TotalNota"), "#,##0.00") Then
            Sql = ""
            Sql = "Update NfCapa Set TotalNota = " & ConverteVirgula(Format(rdoVlItem("TotItem") + rdoVlNota("FreteCobr") - rdoVlNota("Desconto"), "#,##0.00")) & ", VlrMercadoria = " & ConverteVirgula(Format(rdoVlItem("TotItem"), "#,##0.00")) & " " _
                & "Where NfCapa.NumeroPed = " & pedido & " "
            
            rdoCNLoja.Execute (Sql)
        End If
    End If
    
    rdoVlNota.Close
    rdoVlItem.Close
    
End Function

Function CriaSAT(NroNotaFiscal As Long, pedido As Long)
                   
     emitiNota = True
     frmEmissaoNFe.Show vbModal
                      
End Function

Function CriaNFE(NroNotaFiscal As Long, pedido As Long)
                   
     Dim Carimbo As String
                   
     EnviaEmail pedido
                   
     emitiNota = True
     frmEmissaoNFe.Show vbModal
                                      
End Function

Sub GeraArqTXTok()
crlf = Chr(13) & Chr(10)

wdelimitador = "|"

'Call Ler_NF
GLB_Contingencia = "N"

If GLB_Contingencia = "N" Then


    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_IDE (Sem Contingência)
    '----------------------------------------------------------------------------------------------------------------------

''''''''''''''''''''''

       GLB_Loja = Trim(wLoja)
       GLB_NF = NroNotaFiscal
       GLB_Serie = "NE"
       Sql = "Select ce_tipopessoa from fin_cliente,nfcapa where lojaorigem = '" & wLoja & "' " & _
             "and NF = " & NroNotaFiscal & " and Serie = 'NE' and cliente = ce_codigoCliente"
       rsNFE.CursorLocation = adUseClient
       rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 '      GLB_Pessoa = rsComplementoVenda("ce_tipopessoa")
       rsNFE.Close
       
       txtNFe = FreeFile
       wExtensaoDoArquivo = "TXT"
       wNomeArquivo = "NFE" & LTrim(RTrim(GLB_Loja)) & LTrim(RTrim(GLB_NF)) & ".txt"
       
       Open Trim(PegaCaminhoNFe) & "\nfestart\txt\FIL_" & LTrim(RTrim(GLB_Loja)) & "\txt\" & wNomeArquivo For Output Access Write As #txtNFe
       
      Sql = "Select  * from NFE_IDE" _
              & " Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" & GLB_Serie & "'"
      rsNFE.CursorLocation = adUseClient
      rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '
       
       
       linhaArquivo = "NOTA FISCAL|1"
       Print #txtNFe, linhaArquivo
       linhaArquivo = "A|2.00|NFe|"
       Print #txtNFe, linhaArquivo
       linhaArquivo = "B|" & LTrim(RTrim(rsNFE("cUF"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cNF"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("natop"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("indpag"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("mod"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("Serie"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("nnf"))) & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("demi"), "yyyy-mm-dd") & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("dsaient"), "yyyy-mm-dd") & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("hsaient"), "hh:mm:ss") & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpnf"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cmunfg"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpimp"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpemis"))) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpamb"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("finnfe"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("procemi"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("verproc"))) & LTrim(RTrim(wdelimitador))
                       
       Print #txtNFe, linhaArquivo
       rsNFE.Close
   
   
   
Else
      
      
    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_IDE (Com Contingência)
    '----------------------------------------------------------------------------------------------------------------------
  
    wTipoNf = "5"
  
    Sql = ""
    Sql = "Select Top 1 * from NFE_IDE" _
          & " where Situacao = 'A' "
    
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    If rsNFE.EOF Then
       rsNFE.Close
       Exit Sub
    End If
     
    If Not rsNFE.EOF Then
       GLB_Loja = rsNFE("eloja")
       GLB_NF = rsNFE("enf")
       GLB_Serie = rsNFE("eserie")
       Sql = "Select ce_tipopessoa from fin_cliente,nfcapa where lojaorigem = '" & wLoja & "' " & _
             "and NF = " & NroNotaFiscal & " and Serie = 'NE' and cliente = ce_codigoCliente"
       rsNFE.CursorLocation = adUseClient
       rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       GLB_Pessoa = rsComplementoVenda("ce_tipopessoa")
       rsNFE.Close
       txtNFe = FreeFile
       wExtensaoDoArquivo = "TXT"
       wNomeArquivo = "NFE" & LTrim(RTrim(GLB_Loja)) & LTrim(RTrim(GLB_NF)) & "." & wExtensaoDoArquivo
       
       Open Trim(PegaCaminhoNFe) & " \nfestart\txt\FIL_" & LTrim(RTrim(GLB_Loja)) & "\txt\" & wNomeArquivo For Output Access Write As #txtNFe
       
       linhaArquivo = "NOTA FISCAL|1"
       Print #txtNFe, linhaArquivo
       linhaArquivo = "A|2.00|NFe|"
       Print #txtNFe, linhaArquivo
       linhaArquivo = "B|" & LTrim(RTrim(rsNFE("cUF"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cNF"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("natop"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("indpag"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("mod"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("Serie"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("nnf"))) & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("demi"), "yyyy-mm-dd") & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("dsaient"), "yyyy-mm-dd") & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("hsaient"), "hh:mm:ss") & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpnf"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cmunfg"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpimp"))) & LTrim(RTrim(wdelimitador)) _
                       & wTipoNf & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("tpamb"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("finnfe"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("procemi"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("verproc"))) & LTrim(RTrim(wdelimitador)) _
                       & Format(rsNFE("dhCont"), "yyyy-mm-dd hh:mm:ss") & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xJust"))) & LTrim(RTrim(wdelimitador))
                       
       Print #txtNFe, linhaArquivo
       rsNFE.Close
   End If

End If
   
   
   
   
'----------------------------------------------------------------------------------------------------------------------
' NFE_EMIT
'----------------------------------------------------------------------------------------------------------------------

Sql = ""
Sql = "Select  * from NFE_EMIT" _
    & " Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" & GLB_Serie & "'"
      rsNFE.CursorLocation = adUseClient
      rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '

     
       linhaArquivo = "C|" & LTrim(RTrim(rsNFE("xnome"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xfant"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("ie"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("CRT"))) & LTrim(RTrim(wdelimitador))
                       
       Print #txtNFe, linhaArquivo
       
       linhaArquivo = "C02|" & LTrim(RTrim(rsNFE("cnpj")))
       Print #txtNFe, linhaArquivo
       
 linhaArquivo = "C05|" & LTrim(RTrim(rsNFE("xlgr"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("nro"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xcpl"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xbairro"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cmun"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xmun"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("uf"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cep"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cpais"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xpais"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("fone"))) & LTrim(RTrim(wdelimitador))
                       
       Print #txtNFe, linhaArquivo
       rsNFE.Close
       
    
       
'----------------------------------------------------------------------------------------------------------------------
' NFE_DEST
'----------------------------------------------------------------------------------------------------------------------

Sql = ""
Sql = "Select  * from NFE_DEST" _
    & " Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" & GLB_Serie & "'"
      rsNFE.CursorLocation = adUseClient
      rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
      
      If GLB_Pessoa <> "F" And GLB_Pessoa <> "U" Then
         GLB_IE = LTrim(RTrim(rsNFE("ie")))
      Else
         GLB_IE = "ISENTO"
      End If
  
       linhaArquivo = "E|" & LTrim(RTrim(rsNFE("xnome"))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(GLB_IE)) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(rsNFE("isuf"))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(rsNFE("email"))) & LTrim(RTrim(wdelimitador))
                           
       Print #txtNFe, linhaArquivo
                
       If GLB_Pessoa <> "F" And GLB_Pessoa <> "U" Then
          linhaArquivo = "E02|" & LTrim(RTrim(rsNFE("cnpj")))
          Print #txtNFe, linhaArquivo
       Else
          linhaArquivo = "E03|" & LTrim(RTrim(rsNFE("cpf")))
          Print #txtNFe, linhaArquivo
       End If
 '      rsNFE.Close

       
       linhaArquivo = "E05|" & LTrim(RTrim(rsNFE("xlgr"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("nro"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xcpl"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xbairro"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cmun"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xmun"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("uf"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cep"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("cpais"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("xpais"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("fone"))) & LTrim(RTrim(wdelimitador))
                       
                       
       Print #txtNFe, linhaArquivo
       
       
       rsNFE.Close
       
'----------------------------------------------------------------------------------------------------------------------
' NFE_PROD
'----------------------------------------------------------------------------------------------------------------------
 Sql = ""
 Sql = "Select * from NFE_PROD,ProdutoLoja " _
       & "Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" _
       & GLB_Serie & "' and pr_referencia=i_cprod "
    
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
    If Not rsNFE.EOF Then
       Do While Not rsNFE.EOF
          If rsNFE("I_vdesc") = 0 Then
             wString = " "
          Else
             wString = LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vdesc"), "0.0000"))))
          End If
          
          linhaArquivo = "H|" & LTrim(RTrim(rsNFE("H_nitem"))) & LTrim(RTrim(wdelimitador))
          Print #txtNFe, linhaArquivo

          linhaArquivo = "I|" & LTrim(RTrim(rsNFE("I_cprod"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_cean"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_xprod"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_ncm"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_extipi"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_cfop"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_ucom"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_qcom"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vuncom"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vprod"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_ceantrib"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_utrib"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_qtrib"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vuntrib"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vFrete"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vSeg"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & wString & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("I_vOutro"), "0.0000")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(rsNFE("I_indTot"))) & LTrim(RTrim(wdelimitador))
                       
           Print #txtNFe, linhaArquivo
'----------------------------------------------------------------------------------------------------------------------
           
           linhaArquivo = "M"
           
           Print #txtNFe, linhaArquivo
           
'----------------------------------------------------------------------------------------------------------------------
           linhaArquivo = "N"
           
           Print #txtNFe, linhaArquivo
           
'----------------------------------------------------------------------------------------------------------------------
        If rsNFE("pr_substituicaotributaria") = "N" And rsNFE("pr_codigoreducaoicms") = 0 Then
           GLB_SitTributaria = "00"
           GLB_NSITRIB = "N02"
           linhaArquivo = LTrim(RTrim(GLB_NSITRIB)) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(rsNFE("N_origicms"))) _
                      & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(GLB_SitTributaria)) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(rsNFE("N_modbcicms"))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vbcicms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(rsNFE("N_picms"))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vicms"), "0.00")))) & LTrim(RTrim(wdelimitador))
          Print #txtNFe, linhaArquivo
        End If
        If rsNFE("pr_substituicaotributaria") = "N" And rsNFE("pr_codigoreducaoicms") > 0 Then
             GLB_SitTributaria = "20"
             GLB_NSITRIB = "N04"
            linhaArquivo = LTrim(RTrim(GLB_NSITRIB)) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(rsNFE("N_origicms"))) _
                      & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(GLB_SitTributaria)) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(rsNFE("N_modbcicms"))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_predbcicms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vbcicms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_picms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                      & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vicms"), "0.00")))) & LTrim(RTrim(wdelimitador))
            Print #txtNFe, linhaArquivo
        End If
          If rsNFE("pr_substituicaotributaria") = "S" Then
             GLB_SitTributaria = "60"
             GLB_NSITRIB = "N08"
             linhaArquivo = LTrim(RTrim(GLB_NSITRIB)) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(rsNFE("N_origicms"))) _
                      & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(GLB_SitTributaria)) & LTrim(RTrim(wdelimitador)) _
                         & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vbcst"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                        & LTrim(RTrim(ConverteVirgula(Format(rsNFE("N_vicmsst"), "0.00")))) & LTrim(RTrim(wdelimitador))
            Print #txtNFe, linhaArquivo
          
   
          End If

 '----------------------------------------------------------------------------------------------------------------------
                       
       linhaArquivo = "Q"
       Print #txtNFe, linhaArquivo
       
       linhaArquivo = "Q02|" & LTrim(RTrim(rsNFE("Q_cstpis"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("Q_vbcpis"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("Q_ppis"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("Q_vpis"), "0.00")))) & LTrim(RTrim(wdelimitador))
           Print #txtNFe, linhaArquivo
'----------------------------------------------------------------------------------------------------------------------
                    
   linhaArquivo = "S"
   Print #txtNFe, linhaArquivo
   
   linhaArquivo = "S02|" & LTrim(RTrim(rsNFE("S_cstcofins"))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("S_vbccofins"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("S_pcofins"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                       & LTrim(RTrim(ConverteVirgula(Format(rsNFE("S_vcofins"), "0.00")))) & LTrim(RTrim(wdelimitador))
       Print #txtNFe, linhaArquivo
'----------------------------------------------------------------------------------------------------------------------
               
               rsNFE.MoveNext
           Loop
       
           rsNFE.Close
           
        End If
          
    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_TOTAL
    '----------------------------------------------------------------------------------------------------------------------
     Sql = ""
     Sql = "Select * from NFE_TOTAL Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" _
            & GLB_Serie & "'"

        rsNFE.CursorLocation = adUseClient
        rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '

        If Not rsNFE.EOF Then
           Do While Not rsNFE.EOF
        linhaArquivo = "W"
        Print #txtNFe, linhaArquivo

              linhaArquivo = "W02|" & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vbcicms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vicms"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vbcst"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vst"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vprod"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vfrete"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vseg"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vdesc"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vii"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vipi"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vpis"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vcofins"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("voutro"), "0.00")))) & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vnf"), "0.00"))))

               Print #txtNFe, linhaArquivo
               rsNFE.MoveNext
           Loop
       

           

        End If
rsNFE.Close
          
    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_TRANSP
    '----------------------------------------------------------------------------------------------------------------------
     Sql = ""
     Sql = "Select * from NFE_TRANSP Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" _
            & GLB_Serie & "'"

       rsNFE.CursorLocation = adUseClient
       rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '

        If Not rsNFE.EOF Then
           Do While Not rsNFE.EOF
           
              linhaArquivo = "X|" & LTrim(RTrim(rsNFE("modfrete")))
              Print #txtNFe, linhaArquivo
              
              linhaArquivo = "X26|" & LTrim(RTrim(rsNFE("qvol"))) & LTrim(RTrim(wdelimitador)) & _
                             LTrim(RTrim(rsNFE("esq"))) & LTrim(RTrim(wdelimitador)) & LTrim(RTrim(wdelimitador)) & _
                             LTrim(RTrim(wdelimitador)) & LTrim(RTrim(rsNFE("pesoL"))) & LTrim(RTrim(wdelimitador)) & _
                             LTrim(RTrim(rsNFE("pesoB"))) & LTrim(RTrim(wdelimitador))
              Print #txtNFe, linhaArquivo
              
               rsNFE.MoveNext
           Loop
       ' & LTrim(RTrim(GLB_SitTributaria)) &
 
           rsNFE.Close
           
        End If

    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_COBR
    '----------------------------------------------------------------------------------------------------------------------
       Sql = ""
       Sql = "Select * from NFE_COBR Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" _
            & GLB_Serie & "'"

        rsNFE.CursorLocation = adUseClient
        rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '
        
        If Not rsNFE.EOF Then
           Do While Not rsNFE.EOF
              
              linhaArquivo = "Y"
              Print #txtNFe, linhaArquivo
              
              linhaArquivo = "Y07|" & LTrim(RTrim(rsNFE("ndup"))) & LTrim(RTrim(wdelimitador)) _
                           & Format(rsNFE("dvend"), "yyyy-mm-dd") & LTrim(RTrim(wdelimitador)) _
                           & LTrim(RTrim(ConverteVirgula(Format(rsNFE("vdup"), "#####0.00"))))
                           
               Print #txtNFe, linhaArquivo
               rsNFE.MoveNext
           Loop
        End If
       rsNFE.Close
    '----------------------------------------------------------------------------------------------------------------------
    ' NFE_INFADIC
    '----------------------------------------------------------------------------------------------------------------------
       Sql = ""
       Sql = "Select * from NFE_INFADIC Where eloja ='" & GLB_Loja & "' and enf = " & GLB_NF & " and eserie ='" _
            & GLB_Serie & "'"

        rsNFE.CursorLocation = adUseClient
        rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic '

        If Not rsNFE.EOF Then
           Do While Not rsNFE.EOF
           
              linhaArquivo = "Z|" & LTrim(RTrim(rsNFE("infadfisco"))) & LTrim(RTrim(wdelimitador)) _
                             & LTrim(RTrim(rsNFE("infcpl")))
              Print #txtNFe, linhaArquivo
              
              linhaArquivo = "Z04|" & LTrim(RTrim(rsNFE("xCampoCont"))) & LTrim(RTrim(wdelimitador)) _
                             & LTrim(RTrim(rsNFE("xTextoCont")))
              Print #txtNFe, linhaArquivo
              
               rsNFE.MoveNext
           Loop

        End If
     linhaArquivo = "FIM" & LTrim(RTrim(wdelimitador))
     Print #txtNFe, linhaArquivo

     rsNFE.Close
     Close #txtNFe
     
     Sql = ""
     Sql = "Update NFE_IDE Set Situacao = 'P' Where eloja='" & LTrim(RTrim(GLB_Loja)) & "' and enf = " & GLB_NF _
         & " and eserie = '" & GLB_Serie & "'"
       rsNFE.CursorLocation = adUseClient
       rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     
'     Print #txtNFE, LTrim(RTrim(wdelimitador))

End Sub

Private Sub ImprimeCarimbo()
                 
                       Sql = ""
'                       SQL = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
'                             "CNF_Nf = " & rsNFE("nf") & " and CNF_Serie = '" & rsNFE("Serie") & "' and CNF_Loja = '" & rsNFE("Lojaorigem") & "'" & _
'                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       Sql = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
                             "CNF_Numeroped = '" & rsNFE("numeroped") & "'" & _
                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       
                       RsPegaItensEspeciais.CursorLocation = adUseClient
                       RsPegaItensEspeciais.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

                       If Not RsPegaItensEspeciais.EOF Then
                            Printer.Print ""
                            Do While Not RsPegaItensEspeciais.EOF
                                If Trim(RsPegaItensEspeciais("CNF_tipocarimbo")) = "Z" Then
                                 wStr16 = right$(Space(116) & Trim(RsPegaItensEspeciais("CNF_Carimbo")), 116)
                                Else
                                 wStr16 = Space(5) & left$(RsPegaItensEspeciais("CNF_Carimbo") & Space(116), 116)
                                End If
                                 Printer.Print wStr16

                                 If RsPegaItensEspeciais("CNF_DetalheImpressao") = "D" Then
                                     wConta = wConta + 1
'                                     RsPegaItensEspeciais.MoveNext
                                 ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "C" Then
                                     
                                     Do While wConta < 34
                                       wConta = wConta + 1
                                       Printer.Print ""
                                     Loop

                                     wConta = 0

                         
                                     wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                                     If wControlaQuebraDaPagina = 3 Then
                                        Printer.Print ""
                                        wControlaQuebraDaPagina = 0
                                     End If

                                     Cabecalho rsNFE("tiponota")
                                     Printer.Print ""
                                ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "T" Then
                                       wConta = wConta + 1
                                       Printer.Print ""
                                       Call FinalizaNota(wPedido)
                                Else
                                       wConta = wConta + 1
                                End If
                                RsPegaItensEspeciais.MoveNext
                            Loop

                             RsPegaItensEspeciais.Close
'                             Call FinalizaNota(wPedido)
                             Exit Sub
                         Else
                             RsPegaItensEspeciais.Close
                             Call FinalizaNota(wPedido)
                         End If

End Sub

Function PegaLojaControle() As String

    Sql = ""
    Sql = "Select CTS_Loja from ControleSistema"
       
        rspegaloja.CursorLocation = adUseClient
        rspegaloja.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rspegaloja.EOF Then
        PegaLojaControle = Trim(rspegaloja("CTS_Loja"))
    End If
   rspegaloja.Close
End Function
Function PegaCaminhoNFe() As String

    Sql = ""
    Sql = "Select CTS_CaminhoNFe from ControleSistema"
       
        rspegaloja.CursorLocation = adUseClient
        rspegaloja.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rspegaloja.EOF Then
        PegaCaminhoNFe = Trim(rspegaloja("CTS_CaminhoNFe"))
    End If
   rspegaloja.Close
End Function

Function PegaSerieNota() As String

       Sql = ""
       Sql = "Select CS_Serie AS serie from ControleSerie where CS_NroCaixa = '" & GLB_Caixa & "'"
     
       rdoSerie.CursorLocation = adUseClient
       rdoSerie.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rdoSerie.EOF Then
           PegaSerieNota = RTrim(rdoSerie("serie"))
       End If
       rdoSerie.Close

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
     
     Sql = ""
     Sql = "Select CTS_NumeroNF + 1 as NumNota from ControleSistema"
   
     rsNFE.CursorLocation = adUseClient
     rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If Not rsNFE.EOF Then
        
        ExtraiSeqNotaControle = rsNFE("NumNota")
        Sql = "update ControleSistema set CTS_NumeroNF= " & rsNFE("NumNota") & ""
        rdoCNLoja.Execute (Sql)
     End If
     rsNFE.Close
End Function

Public Function ExtraiSeqNEControle() As Double
     Dim WnovaSeqNota As Long
     
     Sql = ""
     Sql = "Select CTS_NumeroNE + 1 as NumNota from ControleSistema"
   
     rsNFE.CursorLocation = adUseClient
     rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If Not rsNFE.EOF Then
        
        ExtraiSeqNEControle = rsNFE("NumNota")
        Sql = "update ControleSistema set CTS_NumeroNE= " & rsNFE("NumNota") & ""
        rdoCNLoja.Execute (Sql)
     End If
     rsNFE.Close
End Function


Public Function ExtraiSeq00Controle() As Double
     Dim WnovaSeqNota As Long
     
     Sql = ""
     Sql = "Select CTS_Numero00 + 1 as NumNota from ControleSistema"
   
     rsNFE.CursorLocation = adUseClient
     rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     If Not rsNFE.EOF Then
        
        ExtraiSeq00Controle = rsNFE("NumNota")
        Sql = "update ControleSistema set CTS_Numero00= " & rsNFE("NumNota") & ""
        rdoCNLoja.Execute (Sql)
     End If
     rsNFE.Close
End Function

Public Function EmiteNotafiscal(ByVal nota As Double, ByVal Serie As String)
Dim wControlaQuebraDaPagina As Integer
wControlaQuebraDaPagina = 0

    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(Glb_ImpNotaFiscal) Then
           ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
      
    wSerie = Serie
    wNotaTransferencia = False
    wPagina = 0
    
    Call DadosLoja
            
    Sql = ""
    Sql = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.QTDITEM, " & _
          "NFCAPA.AV,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.VOLUME,NFCAPA.PESOBR, " & _
          "NFCAPA.CLIENTE,fin_CLIENTE.CE_Telefone,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA,fin_cliente.ce_numero, " & _
          "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,Nfcapa.nf,NfCapa.Desconto,NFCAPA.CODOPER,NFCAPA.TOTALNOTA, " & _
          "NFCAPA.VlrMercadoria,Nfcapa.lojaOrigem,NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,fin_Cliente.ce_razao, " & _
          "fin_Cliente.ce_cgc,NFCAPA.CONDPAG,fin_cliente.ce_Endereco,fin_Cliente.ce_municipio,fin_Cliente.ce_Bairro,fin_Cliente.CE_Cep," & _
          "fin_Cliente.ce_inscricaoEstadual,NfCapa.CondPag,NfCapa.DataPag,fin_Cliente.ce_Estado,NFCAPA.TOTALNOTAALTERNATIVA, " & _
          "NFCAPA.VALORTOTALCODIGOZERO, NfItens.Referencia , NfItens.QTDE, NfItens.VLUNIT, NfItens.VLTOTITEM, " & _
          "NfItens.ICMS, NfItens.TipoNota, NFCAPA.EmiteDataSaida, fin_cliente.ce_TipoPessoa,NFCAPA.CPFNFP " & _
          "From NFCAPA,NFITENS,fin_Cliente Where NfCapa.nf= " & nota & " and NfCapa.Serie in ('" & Serie & "') and " & _
          "NfCapa.lojaorigem='" & Trim(wLoja) & "' and NfItens.LojaOrigem=NfCapa.LojaOrigem  and " & _
          "NfItens.Serie = NfCapa.Serie And NfItens.nf = NfCapa.nf And fin_Cliente.ce_CodigoCliente = NfCapa.Cliente"
   
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If Not rsNFE.EOF Then
      Cabecalho "V"
      
      Sql = "Select produtoloja.pr_referencia,produtoloja.pr_descricao, " _
          & "produtoloja.pr_classefiscal,produtoloja.pr_unidade,produtoloja.pr_st, " _
          & "produtoloja.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota," _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao,nfitens.CSTICMS," _
          & "nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produtoloja,nfitens " _
          & "where produtoloja.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & nota & " and Serie='" & Serie & "' order by nfitens.item"
     
      rsItensVenda.CursorLocation = adUseClient
      rsItensVenda.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

      If Not rsItensVenda.EOF Then
         wConta = 0
         wContItem = 0
         Printer.Print ""
         Do While Not rsItensVenda.EOF
            wContItem = wContItem + 1
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = rsItensVenda("PR_Referencia")
            If Wsm = True Then
                 wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
                        
                   wStr16 = ""
                   wStr16 = left$(rsItensVenda("ReferenciaAlternativa") & Space(7), 7) _
                         & Space(2) & left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(55), 55) _
                         & left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
                         & Space(11), 11) & left$("0" + Trim(rsItensVenda("CSTICMS")) & Space(5), 5) _
                         & left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
                         & right$(Space(8) & Format(rsItensVenda("QTDE"), "##0"), 8) _
                         & right$(Space(13) & Format(rsItensVenda("vlunit"), "#####0.00"), 13) _
                         & right$(Space(13) & Format(rsItensVenda("VlTotItem"), "#####0.00"), 13) _
                         & right$(Space(4) & Format(rsItensVenda("pr_icmssaida"), "#0"), 4)
            Else
                     
                   wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
                   If wPegaDescricaoAlternativa = "" Then
                        wPegaDescricaoAlternativa = "0"
                   End If
                   If wPegaDescricaoAlternativa <> "0" Then
                         wDescricao = wPegaDescricaoAlternativa
                   Else
                         wDescricao = Trim(rsItensVenda("pr_descricao"))
                   End If
                   
                   
                   wStr16 = ""
                   wStr16 = left$(rsItensVenda("pr_referencia") & Space(7), 7) _
                         & Space(2) & left$(Format(Trim(wDescricao), ">") & Space(55), 55) _
                         & left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
                         & Space(11), 11) & left$(Trim("0" + Format(rsItensVenda("CSTICMS"), "00")) & Space(5), 5) _
                         & left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
                         & right$(Space(8) & Format(rsItensVenda("QTDE"), "##0"), 8) _
                         & right$(Space(13) & Format(rsItensVenda("vlunit"), "#####0.00"), 13) _
                         & right$(Space(13) & Format(rsItensVenda("VlTotItem"), "#####0.00"), 13) _
                         & right$(Space(4) & Format(rsItensVenda("pr_icmssaida"), "#0"), 4)
                                  
            End If

                   Printer.Print wStr16
                      
                      If rsItensVenda("DetalheImpressao") = "D" Then
                         wConta = wConta + 1

                      ElseIf rsItensVenda("DetalheImpressao") = "C" Then
                            
                        Do While wConta < 34
                            wConta = wConta + 1
                            Printer.Print ""
                        Loop
                         
                         wConta = 0

                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If

                         Cabecalho rsItensVenda("TipoNota")
                         Printer.Print ""
                         
                       If wContItem = rsNFE("QTDITEM") Then
                          Call ImprimeCarimbo
                       End If

                      ElseIf rsItensVenda("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         Call ImprimeCarimbo

                      Else
                         wConta = wConta + 1
                      End If
                       rsItensVenda.MoveNext
            Loop
         Else
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
        Exit Function
    End If
rsItensVenda.Close
rsNFE.Close


End Function


Private Sub FinalizaNota(wPedido As String)
     If wNotaTransferencia = False Then
   
        Do While wConta < 13
        wConta = wConta + 1
        Printer.Print ""
        Loop
       
     End If

     If Wsm = True Then
        wStr9 = right$(Space(9) & Format(rsNFE("BaseICMS"), "######0.00"), 9) & right$(Space(25) & Format(rsNFE("VLRICMS"), "######0.00"), 9) & Space(34) & right$(Space(10) & Format(rsNFE("TotalNotaAlternativa"), "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = right(Space(9) & Format(Space(9) & rsNFE("FreteCobr"), "######0.00"), 9) & Space(46) & right(Space(10) & Format(rsNFE("TotalNotaAlternativa"), "######0.00"), 10)
        Printer.Print wStr10
     Else
        wStr9 = right$(Space(9) & Format(rsNFE("BaseICMS"), "######0.00"), 9) & right$(Space(25) & Format(rsNFE("VLRICMS"), "######0.00"), 12) & Space(34) & right$(Space(10) & Format(rsNFE("VlrMercadoria"), "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = right(Space(9) & Format(Space(9) & rsNFE("FreteCobr"), "######0.00"), 9) & Space(46) & right(Space(10) & Format(rsNFE("TotalNota"), "######0.00"), 10)
        Printer.Print wStr10
     End If
     
     wStr11 = Space(2) & "                          "
     Printer.Print wStr11
     wStr12 = Space(2) & "                                                     "
     Printer.Print wStr12
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = right$(Space(5) & Format(rsNFE("Volume"), "######0.00"), 5) & Space(5) & "Volume(s)" & Space(25) & right$(Space(7) & Format(rsNFE("PesoBR"), "######0.00"), 7) & Space(5) & right$(Space(7) & Format(rsNFE("PesoBR"), "######0.00"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(105) & right$(Space(7) & Format(rsNFE("Nf"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
'     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.EndDoc
     

End Sub
    

Function Cabecalho(ByVal tiponota As String)
        
    Dim wCgcCliente As String
    Dim impri As Long
    Dim Linha(15) As String
    Dim ContLinha As Integer
    Dim ContParcela As Integer
    
    impri = Printer.Orientation
    wPagina = wPagina + 1
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    Linha(1) = "          "
    Linha(2) = "          "
    Linha(3) = "          "
    Linha(4) = "          "
    Linha(5) = "          "
    Linha(6) = "          "
    Linha(7) = "          "
    Linha(8) = "          "
    Linha(9) = "          "
    Linha(10) = "          "
    Linha(11) = "          "
    Linha(12) = "          "
    Linha(13) = "          "
    Linha(14) = "          "
    Linha(15) = "          "
    ContLinha = 1
    
    wCondicao = "            "
    Wav = "          "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = ""
    wentrada = "        "
    
    wLojaVenda = IIf(IsNull(rsNFE("LojaVenda")), rsNFE("LojaOrigem"), rsNFE("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(rsNFE("VendedorLojaVenda")), 0, rsNFE("VendedorLojaVenda"))


    If UCase(tiponota) = "T" Then
        WNatureza = "TRANSFERENCIA"
    ElseIf UCase(tiponota) = "V" Then
        WNatureza = "VENDA"
    ElseIf UCase(tiponota) = "E" Then
        WNatureza = "DEVOLUCAO"
    ElseIf UCase(tiponota) = "S" And GLB_CFOP = "5949" Or GLB_CFOP = "6949" Then
        WNatureza = "OUTRAS OPER Ñ ESPEC."
    End If
    
    
    If Trim(wLojaVenda) > 0 Then
        If Trim(wLojaVenda) <> Trim(rsNFE("LojaOrigem")) Then
            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
        Else
            wStr6 = ""
        End If
    Else
        wStr6 = ""
    End If
    If Trim(rsNFE("AV")) > 1 Then
        If Mid(wCondicao, 1, 9) = "Faturada " Then
            Wav = "AV            : " & Trim(rsNFE("AV"))
        End If
    End If
    
    If Trim(WNatureza) = "TRANSFERENCIA" Then
        wCondicao = "            "
    ElseIf Trim(WNatureza) = "DEVOLUCAO" Then
        wCondicao = "            "
    End If
    
    Linha(ContLinha) = "PEDIDO " & rsNFE("NUMEROPED") & "  VEN " & rsNFE("VENDEDOR")
    ContLinha = ContLinha + 1
             
    Sql = "select mo_descricao,mc_valor,mo_grupo from movimentocaixa,modalidade " & _
          "where mc_grupo = mo_grupo and mc_documento = " & rsNFE("nf") & " and mc_Serie ='" & rsNFE("serie") & _
          "' and mc_loja = '" & Trim(rsNFE("lojaorigem")) & "' and mc_grupo like '10%'"

    rdoModalidade.CursorLocation = adUseClient
    rdoModalidade.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
    If Not rdoModalidade.EOF Then
        Do While Not rdoModalidade.EOF
         
          If rdoModalidade("mo_grupo") = 10501 Then
            
               Sql = "Select cp_condicao,cp_intervaloParcelas,cp_parcelas from CondicaoPagamento " _
                    & "where  CP_Codigo =" & rsNFE("CondPag")

               rdoConPag.CursorLocation = adUseClient
               rdoConPag.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
               wValorParcela = Format((rsNFE("totalnota") - rsNFE("pgentra")) / rdoConPag("cp_parcelas"), "###,##0.00")
               ContParcela = 1
               wMid = 1
               Linha(ContLinha) = "Faturada " & rdoConPag("cp_parcelas") & " Parc    " & wValorParcela
               ContLinha = ContLinha + 1
               
               Do While Len(rdoConPag("cp_intervaloParcelas")) > wMid
               
                 If rdoConPag("cp_Parcelas") = 1 Then
                     Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                     wMid = wMid + 3
                 ElseIf rdoConPag("cp_Parcelas") Mod 2 = 0 Then
                       Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                       wMid = wMid + 6
                 Else
                       If Len(rdoConPag("cp_intervaloParcelas")) - 3 > wMid Then
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                           wMid = wMid + 6
                       Else
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                           wMid = wMid + 3
                       End If
                 End If
                 ContLinha = ContLinha + 1
              Loop
              rdoConPag.Close
           Else
              Linha(ContLinha) = rdoModalidade("mo_descricao") & ":   " & Format(rdoModalidade("mc_valor"), "0.00")
              ContLinha = ContLinha + 1
           End If
           rdoModalidade.MoveNext
        Loop
    End If
rdoModalidade.Close

    If rsNFE("Pgentra") <> 0 Then
       wentrada = Format(rsNFE("Pgentra"), "#####0.00")
       Linha(ContLinha) = "Entrada : " & Format(wentrada, "0.00")
       ContLinha = ContLinha + 1
    End If
    If (IIf(IsNull(rsNFE("PedCli")), 0, rsNFE("PedCli"))) <> 0 Then
       Linha(ContLinha) = "Ped. Cliente    : " & Trim(rsNFE("PedCli"))
       ContLinha = ContLinha + 1
    End If
    
    'SQL = "Select CT_pis,CT_cofins from Controle"
    'RSControleImpostos.CursorLocation = adUseClient
    'RSControleImpostos.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    'Linha(ContLinha) = "Impostos:"
    'ContLinha = ContLinha + 1
    'Linha(ContLinha) = "COFINS " & Format(((rsNFE("totalnota") * RSControleImpostos("CT_cofins")) / 100), "0.00") _
    '                 & " / PIS " & Format(((rsNFE("totalnota") * RSControleImpostos("CT_pis")) / 100), "0.00")
    'RSControleImpostos.Close
     
    If wPagina = 1 Then
        wCGC = right(String(14, "0") & wCGC, 14)
        wCGC = Format(Mid(wCGC, 1, Len(wCGC) - 6), "###,###,###") & "/" & Mid(wCGC, Len(wCGC) - 5, Len(wCGC) - 10) & "-" & Mid(wCGC, 13, Len(wCGC))
        wCGC = right(String(18, "0") & wCGC, 18)
    End If
    wStr0 = Space(110) & wPagina & "/" & rsNFE("PAGINANF")  'Inicio Impressão
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
    

    If Glb_NfDevolucao = True Then
        WNatureza = "DEVOLUCAO"
        wStr1 = (left$(Linha(1) & Space(27), 27)) & Space(10) & left(Format(Trim(UCase(Wendereco)), "<") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(25) & right(Format(rsNFE("nf"), "######"), 7)
    Else
        wStr1 = (left$(Linha(1) & Space(27), 27)) & Space(10) & left(Format(Trim(UCase(Wendereco)), "<") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(25) & right(Format(rsNFE("nf"), "######"), 7)
    End If
    Printer.Print UCase(wStr1)
    wStr2 = (left$(Linha(2) & Space(27), 27)) & Space(10) & left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & left$(Trim(westado), 2)
    Printer.Print UCase(wStr2)
    wStr3 = (left$(Linha(3) & Space(27), 27)) & Space(10) & "(" & wDDDLoja & ")" & left$(Trim(Format(WFone, "####-####")), 9) & "/(" & wDDDLoja & ")" & left$(Format(WFax, "####-####"), 9) & Space(5) & left$(Format((WCep), "00000-000"), 9)
    Printer.Print UCase(wStr3)
    If wSerie = "CT" Then
        wStr4 = (left$(Linha(4) & Space(27), 27))
    Else
        wStr4 = (left$(Linha(4) & Space(27), 27)) & Space(60) & left(Trim(Format(wCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
     wStr4 = (left$(Linha(5) & Space(27), 27))
    
     Printer.Print UCase(wStr4)
    
    If wSerie = "CT" Then
        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = (left$(Linha(6) & Space(27), 27)) & Format(Trim(WNatureza), ">") & Space(27) & left$(rsNFE("codOper"), 10)
        End If
    Else

        If Trim(Wav) <> "" Then
            wStr5 = (left$(Linha(6) & Space(32), 32)) & left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(25) & left$(rsNFE("codOper"), 10) & Space(25) & left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = (left$(Linha(6) & Space(32), 32)) & left(Trim(WNatureza) & Space(25), 25) & left$(rsNFE("codOper"), 10) & Space(28) & left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
    End If
    Printer.Print wStr5
    wStr5 = (left$(Linha(7) & Space(27), 27))
    Printer.Print wStr5

    
    If Trim(rsNFE("ce_TipoPessoa")) = "J" Or Trim(rsNFE("ce_TipoPessoa")) = "O" Then
        wCgcCliente = right(String(14, "0") & Trim(rsNFE("ce_cgc")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = right(String(18, "0") & Trim(wCgcCliente), 18)
    Else
        wCgcCliente = right(String(11, "0") & Trim(rsNFE("ce_cgc")), 11)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 2), "###,###,###") & "-" & Mid(wCgcCliente, 10, Len(wCgcCliente))
        wCgcCliente = right(String(14, "0") & Trim(wCgcCliente), 14)
    End If
    
    If Trim(rsNFE("cliente")) = "999999" Then
      If Len(Trim(rsNFE("CPFNFP"))) = 14 Then
        wCgcCliente = right(String(14, "0") & Trim(rsNFE("CPFNFP")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = right(String(18, "0") & Trim(wCgcCliente), 18)
      Else
        wCgcCliente = right(String(11, "0") & Trim(rsNFE("CPFNFP")), 11)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 2), "###,###,###") & "-" & Mid(wCgcCliente, 10, Len(wCgcCliente))
        wCgcCliente = right(String(14, "0") & Trim(wCgcCliente), 14)
      End If
    
    End If
    
    Printer.Print ""
    If wSerie = "CT" Then
        wStr6 = (left$(Linha(8) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(rsNFE("ce_razao")), ">") & Space(45), 45) & left$(Format(rsNFE("Dataemi"), "yyyy/mm/dd"), 12)
    Else
        wStr6 = (left$(Linha(8) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(rsNFE("ce_razao")), ">") & Space(45), 45) & left$(Trim(wCgcCliente) & Space(24), 24) & left$(Format(rsNFE("Dataemi"), "dd/mm/yy") & Space(12), 12)
    End If
    
    Printer.Print UCase(wStr6)
    
    wStr6 = (left$(Linha(9) & Space(27), 27))
    Printer.Print UCase(wStr6)
    
    If rsNFE("EmiteDataSaida") = "S" Then
        If wSerie = "CT" Then
            wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("ce_endereco") & ", " & rsNFE("ce_numero")), ">") & Space(42), 42) & Space(14) & left$(Format(rsNFE("Dataemi"), "yyyy/mm/dd"), 12)
        Else
            wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("ce_endereco")), ">") & Space(42), 42) & left$(Format(Trim(rsNFE("ce_bairro")), ">") & Space(21), 21) & right$(Space(11) & rsNFE("ce_cep"), 11) & Space(8) & left$(Format(rsNFE("Dataemi"), "dd/mm/yy"), 12)
        End If
    Else: wquant = (wQuantItensNF Mod 8)

        If wSerie = "CT" Then
            wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("ce_endereco") & ", " & rsNFE("ce_numero")), ">") & Space(42), 42) & Space(14) & left$(Format(rsNFE("Dataemi"), "yyyy/mm/dd"), 12)
        Else
            wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("ce_endereco") & ", " & rsNFE("ce_numero")), ">") & Space(42), 42) & left$(Format(Trim(rsNFE("ce_bairro")), ">") & Space(18), 18) & right$(Space(12) & Format(rsNFE("ce_cep"), "#####-###"), 12) '& Space(7) & Left$(Format(rsNFE("Dataemi"), "dd/mm/yy"), 12)
        End If
    End If
    Printer.Print UCase(wStr7)
    wStr7 = (left$(Linha(11) & Space(27), 27))
    Printer.Print UCase(wStr7)
    If wSerie = "CT" Then
        wStr8 = (left$(Linha(12) & Space(27), 27))
    Else
        wStr8 = (left$(Linha(12) & Space(27), 27)) & Space(5) & left$(Format(Trim(rsNFE("ce_municipio")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(rsNFE("ce_telefone"))) & Space(15), 15) & left$(Trim(rsNFE("ce_estado")), 2) & Space(5) & left$(Trim(Format(rsNFE("ce_inscricaoEstadual"), "###,###,###,###")), 15)
    End If

    Printer.Print UCase(wStr8)
    
    Printer.Print ""


    If rdoConPag.State = 1 Then
        rdoConPag.Close
    End If

End Function



Public Function DadosLoja()

    'SQL = ""
    Sql = "Select CTS_Loja,CTS_SenhaLiberacao,CTS_LogoPedido,Loja.* from loja,Controlesistema " & _
          "where lo_loja=CTS_Loja"

    rsNFELoja.CursorLocation = adUseClient
    rsNFELoja.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rsNFELoja.EOF Then

       wRazao = Trim(rsNFELoja("lo_Razao"))
       Wendereco = UCase(rsNFELoja("lo_ENDERECO") & ", " & rsNFELoja("lo_numero"))
       wbairro = rsNFELoja("lo_bairro")
       wCGC = rsNFELoja("lo_CGC")
       WIest = rsNFELoja("lo_INSCRICAOESTADUAL")
       WMunicipio = rsNFELoja("lo_MUNICIPIO")
       westado = rsNFELoja("lo_UF")
       WCep = rsNFELoja("lo_CEP")
       WFone = rsNFELoja("lo_TELEFONE")
       wDDDLoja = rsNFELoja("LO_DDD")
       WFax = rsNFELoja("lo_Fax")
       wLoja = rsNFELoja("CTS_Loja")
       wSenhaLiberacao = rsNFELoja("CTS_SenhaLiberacao")
       GLB_Loja = rsNFELoja("CTS_Loja")
       wNovaRazao = IIf(IsNull(rsNFELoja("lo_Razao")), "0", rsNFELoja("lo_Razao"))
       GLB_Logo = RTrim(rsNFELoja("CTS_LogoPedido"))
       'wMensagemECF = rsNFELoja("CTS_MensagemECF")
    End If
    
    rsNFELoja.Close

End Function
Public Sub ImprimeRomaneio()
Dim wNomeVendedor As String

'Open GLB_Impressora00 For Output As #1
 
    Screen.MousePointer = 11
   
    ValorlItem = 0
    ValorDesconto = 0
    SubTotal = 0

    Sql = ("Select * from Loja Where LO_Loja='" & Trim(GLB_Loja & "'"))

    rsNFE.CursorLocation = adUseClient
    rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    impressoraRelatorio "[INICIO]"
    
   impressoraRelatorio UCase(rsNFE("LO_Razao"))
   impressoraRelatorio UCase(rsNFE("LO_Endereco")) & ", " & rsNFE("LO_numero")
   impressoraRelatorio "TELEFONE: " & rsNFE("LO_Telefone")
   impressoraRelatorio Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(16) & Format(NroNotaFiscal, "###000")
   impressoraRelatorio "========================================"
   impressoraRelatorio "DESCRICAO DO PRODUTO                    "
   impressoraRelatorio "CODIGO  PRODUTO  QTDxUNIT.   VALOR TOTAL"
   impressoraRelatorio "________________________________________"
   rsNFE.Close
   
  Sql = "Select Nfcapa.*,ve_nome From Nfcapa,vende Where  NF = " & NroNotaFiscal & " and Serie='00' and vendedor = ve_codigo"
             
    
              rsNFECapa.CursorLocation = adUseClient
              rsNFECapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             
             If Not rsNFECapa.EOF Then
                wPegaDesconto = rsNFECapa("Desconto")
                wPegaFrete = rsNFECapa("FreteCobr")
              
             End If
             
             
   
   Sql = "Select * from Nfitens " _
       & "Where  NF = " & NroNotaFiscal & " and Serie='00'"
       
      
       rsNFE.CursorLocation = adUseClient
       rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       
       If Not rsNFE.EOF Then
          Do While Not rsNFE.EOF
             Sql = "Select PR_Descricao from Produtoloja Where PR_Referencia ='" & rsNFE("Referencia") & "'"
             rdoProduto.CursorLocation = adUseClient
             rdoProduto.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             
             ValorlItem = (rsNFE("vlunit") * rsNFE("Qtde"))
             SubTotal = (SubTotal + ValorlItem)
           
             impressoraRelatorio Trim(rdoProduto("PR_Descricao"))
             impressoraRelatorio rsNFE("referencia") _
             & Space(3) & right(Space(4) & Format(rsNFE("Qtde"), "###0"), 4) & "x" _
             & Format(rsNFE("vlunit"), "###,###,###.00") & Space(5) _
             & right(Space(10) & Format(ValorlItem, "###,###,###.00"), 14)
             rdoProduto.Close
             rsNFE.MoveNext
          Loop
       End If
       
       rsNFE.Close
       
       If rsNFECapa("vendedor") = 725 Then
         wNomeVendedor = "Caixa"
       Else
         wNomeVendedor = rsNFECapa("ve_nome")
       End If
         
       TotalVenda = (SubTotal - ValorDesconto)
       impressoraRelatorio " "
'       Print #1, "SUB TOTAL " & Space(16) & Right(Space(10) & Format(rsNFECapa("vlrMercadoria"), "###,###,##0.00"), 14)
'       Print #1, ""
'       Print #1, "DESCONTO  " & Space(16) & Right(Space(10) & Format(rsNFECapa("desconto"), "###,###,##0.00"), 14)
'       Print #1, " "
'       Print #1, "FRETE     " & Space(16) & Right(Space(10) & Format(rsNFECapa("fretecobr"), "###,###,##0.00"), 14)
'       Print #1, " "
       impressoraRelatorio "TOTAL     " & Space(16) & right(Space(10) & Format(rsNFECapa("totalnota"), "###,###,##0.00"), 14)
       impressoraRelatorio ""
       impressoraRelatorio "________________________________________"
       impressoraRelatorio "ROMANEIO DE CONFERENCIA"
       impressoraRelatorio "PEDIDO: " & rsNFECapa("numeroped") & _
                 Space(6) & right(Space(20) & rsNFECapa("vendedor"), 20)
       impressoraRelatorio "========================================"
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
       impressoraRelatorio " "
     rsNFECapa.Close
     impressoraRelatorio "[FIM]"
      Screen.MousePointer = 0

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
        wTotalVICMSFECP = 0
        wUltimoItem = 1
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wCarimbo5 = ""
        wCarimbo2 = ""
        wST20 = "N"
        wST60 = "N"
        EncerraVenda = True
        SerieProd = ""
        wRecebeCarimboAnexo = ""
        wQuantItensNF = 0
        
        wIE_icmsFECPUFREMET = 0
        wIE_icmsFECPUFDEST = 0
        wIE_icmsFECPUFREMETTotal = 0
        wIE_icmsFECPUFDESTTotal = 0
        wIE_icmsFECPAliqDest = 0
        wIE_icmsFECPAliqInter = 0
        wIE_icmsFECPPart = 0
        wVICMSFECP = 0
        
        If ConsistenciaNota(NumeroDocumento, SerieDocumento) = False Then
            EncerraVenda = False
            Exit Function
        End If


If RsCapaNF.State = 1 Then
  RsCapaNF.Close
End If

Sql = "Select nfcapa.*, fin_Estado.*,fin_Cliente.* from nfcapa, fin_Estado, fin_cliente where nfcapa.numeroped = " & _
       NumeroDocumento & " and nfcapa.cliente = fin_cliente.ce_codigocliente " & _
      "And fin_cliente.ce_estado = fin_Estado.UF_Estado"
             
             RsCapaNF.CursorLocation = adUseClient
             RsCapaNF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If Not RsCapaNF.EOF Then
           If RsCapaNF("ce_Tipopessoa") = "F" Or RsCapaNF("ce_Tipopessoa") = "U" Then
              wPessoa = 2
           ElseIf RsCapaNF("ce_Tipopessoa") = "O" Then
              wPessoa = 1
           Else
              wPessoa = 1
           End If
           
           'NOVO 2016
           'ICMS FECP
               
           wPICMSFECP = 0
           'IMPOSTO 2016/07
           'If wPessoa = 2 And RsCapaNF("ce_estado") <> "SP" Then
           If RsCapaNF("ce_estado") <> "SP" Then
                 
                Sql = "Select UF_FECP AS FECP, " & vbNewLine _
                      & "UF_ICMSInterEstadual as ICMSInterEstadual, " & vbNewLine _
                      & "UF_ICMSInterno as  ICMSInterno, " & vbNewLine _
                      & "UF_ICMSInterEstadual as  ICMSInterEstadual, " & vbNewLine _
                      & "UF_ICMSDifal as  ICMSDifal, " & vbNewLine _
                      & "UF_Participacao AS Participacao " & vbNewLine _
                      & "from fin_estado " & vbNewLine _
                      & "where UF_Estado = '" & RsCapaNF("ce_estado") & "'"
                RsItensNF.CursorLocation = adUseClient
                RsItensNF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
                    wPICMSFECP = RsItensNF("FECP")
                    wIE_icmsFECPAplicado = RsItensNF("ICMSInterEstadual")
                    wIE_icmsFECPDiferencial = RsItensNF("ICMSDifal")
                    wIE_icmsFECPPart = RsItensNF("Participacao")
                    wIE_icmsFECPAliqDest = RsItensNF("ICMSInterno")
                    wIE_icmsFECPAliqInter = RsItensNF("ICMSInterEstadual")
                
                RsItensNF.Close
                 
           End If
           
           'FIM 2016
           
           wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
           If RsCapaNF("Serie") <> "S1" And RsCapaNF("Serie") <> "D1" Then
              wSerie = ""
           Else
              wSerie = IIf(IsNull(RsCapaNF("Serie")), "", RsCapaNF("Serie"))
           End If
           
        Else
            MsgBox "Nota não encontrada", vbInformation, "Atenção"
            Exit Function
        End If
                 
                    
        Sql = "Select produtoloja.*, nfitens.* from produtoloja,nfitens " _
              & "where nfitens.numeroped = " & NumeroDocumento & "" _
              & " and pr_referencia = nfitens.referencia order by NfItens.Item"
          
              RsItensNF.CursorLocation = adUseClient
              RsItensNF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

          
          
          wLoja = RsCapaNF("lojaorigem")
          wTM = RsCapaNF("TM")

'aqui
          
          If Not RsItensNF.EOF Then
             Do While Not RsItensNF.EOF
                    
                  wChaveICMSItem = wChaveICMS
                  If Trim(wCarimbo5) = "" Then
                    If RsItensNF("PR_substituicaotributaria") = "N" _
                       And RsItensNF("PR_codigoreducaoicms") > 0 Then
                        wST20 = "S"
                    End If
                                    
                    If RsItensNF("PR_substituicaotributaria") = "S" Then
                        wSubstituicaoTributaria = 1
                        wST60 = "S"
                        wChaveICMSItem = wChaveICMSItem & "000" & wSubstituicaoTributaria
                    Else
                        wSubstituicaoTributaria = 0
                        wChaveICMSItem = wChaveICMSItem & Format(RsItensNF("pr_icmssaida"), "####00") & RsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                    End If
                                     
                    
                    If AcharICMSInterEstadual(RsItensNF("PR_Referencia"), wChaveICMSItem) = False Then
                          
                          If AcharICMSInterEstadual(RsItensNF("PR_Referencia"), Mid(Trim(wChaveICMSItem), 1, 2) & "1200") = False Then
                                EncerraVenda = False
                                RsItensNF.Close
                                Exit Function
                          End If
                    End If
                    
                    
                                      
                        wCFOItem = wIE_Cfo
                        GLB_AliquotaAplicadaICMS = wIE_icmsAplicado
                        GLB_Tributacao = wIE_Tributacao
                        GLB_CFOP = wIE_Cfo
                        wAnexoIten = RsItensNF("PR_CodigoReducaoICMS")
                        
                        If wAnexoIten <> 0 Then
                            If wAnexoIten = 1 Then
                                wAnexo1 = RsItensNF("Item") & "," & wAnexo1
                            ElseIf wAnexoIten = 2 Then
                                wAnexo2 = RsItensNF("Item") & "," & wAnexo2
                            End If
                        End If
                        
                        
                            GLB_ValorCalculadoICMS = Format((((RsItensNF("vltotitem") - RsItensNF("desconto")) * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                            GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
                            If GLB_TotalICMSCalculado > 0 Then
                                If wIE_BasedeReducao = 0 Then
                                    If GLB_AliquotaAplicadaICMS = 0 Then
                                        GLB_BasedeCalculoICMS = 0
                                    Else
                                        GLB_BasedeCalculoICMS = (RsItensNF("vltotitem") - RsItensNF("desconto"))
                                    End If
                                Else
                                    GLB_BasedeCalculoICMS = Format((RsItensNF("vltotitem") - RsItensNF("desconto")) - _
                                    (((RsItensNF("vltotitem") - RsItensNF("desconto")) * wIE_BasedeReducao) / 100), "0.00")
                                End If
                                GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
                            End If
                       
                            WAnexoAux = ""
                            If RsItensNF("pr_codigoreducaoicms") <> 0 Then
                               WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "0")
                            End If
                        
                            If wCFOItem = 5102 Or wCFOItem = 6102 Then
                                wCFO1 = wCFOItem
                            ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                                wCFO2 = wCFOItem
                                If Trim(wCFO2) = 6405 Then
                                   wCFO2 = 6404
                                End If
                                If Trim(wCFOItem) = 6405 Then
                                   wCFOItem = 6404
                                End If
                                If Trim(GLB_CFOP) = 6405 Then
                                    GLB_CFOP = 6404
                                End If
                            End If
                        
                        If Trim(wCFO1) = "" And Trim(wCFO2) = "" And RsCapaNF("TipoNota") <> "S" Then
                            wCFO1 = wCFOItem
                     
                        End If
                        
                              
                        If wPICMSFECP > 0 Then
                        
                              wVICMSFECP = Format((((RsItensNF("vltotitem") - RsItensNF("desconto")) * wPICMSFECP) / 100), "0.00")
                              wTotalVICMSFECP = wTotalVICMSFECP + wVICMSFECP
                              'GLB_ValorCalculadoICMS
                              wIE_icmsFECPUFDEST = (RsItensNF("vltotitem") * wIE_icmsFECPDiferencial) / 100
                              wIE_icmsFECPUFREMET = (wIE_icmsFECPUFDEST * (100 - wIE_icmsFECPPart)) / 100
                              wIE_icmsFECPUFDEST = (wIE_icmsFECPUFDEST * wIE_icmsFECPPart) / 100
                              
                              wIE_icmsFECPUFREMETTotal = wIE_icmsFECPUFREMET + wIE_icmsFECPUFREMETTotal
                              wIE_icmsFECPUFDESTTotal = wIE_icmsFECPUFDEST + wIE_icmsFECPUFDESTTotal
                              
                                    
                        End If
                   
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'aqui
                    wQuantItensCapaNF = RsCapaNF("QtdItem")
                    wQuantItensNF = wQuantItensNF + 1
                    wQuantdadeTotalItem = wQuantdadeTotalItem + 1
                    wquant = (wQuantItensNF Mod 12)
                      
                        If wquant <> 0 Then

                             If wQuantItensCapaNF = wQuantItensNF Then
                               If wquant = 11 Then
                                   wDetalheImpressao = "C"
                               Else
                                   wDetalheImpressao = "T"
                               End If
                             ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                               If wquant = 11 Then
                                   wDetalheImpressao = "C"
                               Else
                                   wDetalheImpressao = "T"
                               End If
                             Else
                                 wDetalheImpressao = "D"
                             End If

                        Else
                            
                            wDetalheImpressao = "C"
                            wUltimoItem = wUltimoItem + 1
                        End If
     
                    If wRomaneio = True Then
                       GLB_BasedeCalculoICMS = 0
                       GLB_ValorCalculadoICMS = 0
                    End If

                    rdoCNLoja.BeginTrans
                                        
                    'GLB_DataInicial
                                        
                    Sql = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & " ," _
                    & "valorICMSFECP = " & ConverteVirgula(wVICMSFECP) & " ," _
                    & "aliqICMSFECP = " & ConverteVirgula(wPICMSFECP) & " ," _
                    & "valICMSRemet = " & ConverteVirgula(wIE_icmsFECPUFREMET) & " ," _
                    & "valICMSDest = " & ConverteVirgula(wIE_icmsFECPUFDEST) & " ," _
                    & "aliqICMSInter = " & ConverteVirgula(wIE_icmsFECPAliqInter) & " ," _
                    & "aliqICMSDest = " & ConverteVirgula(wIE_icmsFECPAliqDest) & " ," _
                    & "ICMSInterpart = " & ConverteVirgula(wIE_icmsFECPPart) & " ," _
                    & "DetalheImpressao = '" & wDetalheImpressao & "', CSTICMS = " & Format(GLB_Tributacao, "00") & ", " _
                    & "CFOP = " & GLB_CFOP & ", ICMSAplicado = " & ConverteVirgula(wIE_icmsdestino) _
                    & " where nfitens.numeroped = " & NumeroDocumento _
                    & " and Referencia = '" & RsItensNF("PR_Referencia") & "' and Item=" & RsItensNF("Item") & ""
                    rdoCNLoja.Execute (Sql)
                
                    If Err.Number = 0 Then
                        rdoCNLoja.CommitTrans
                    Else
                        rdoCNLoja.RollbackTrans
                    End If
                    
                RsItensNF.MoveNext
                End If
             Loop
     ' End If 'estava com '
        If wRomaneio = True Then
           wRomaneio = False
        End If
        
        RsItensNF.Close
 
   
' -------------------------------------- INSERIR CARIMBOS --------------------------------------------------

           


            rdoCNLoja.BeginTrans

            Sql = ""
            Sql = "update CarimboNotafiscal set CNF_Serie = '" & RsCapaNF("Serie") & "', CNF_NF = " & RsCapaNF("nf") & _
                  " , CNF_Situacaoprocesso = 'A' , CNF_DataProcesso = '" & Format(Date, "yyyy/mm/dd") & "' " & _
                  " where CNF_NumeroPed = " & NumeroDocumento
              rdoCNLoja.Execute (Sql)
  
            If Err.Number = 0 Then
                 rdoCNLoja.CommitTrans
            Else
                 rdoCNLoja.RollbackTrans
            End If


            'NOVO 2016
            If wIE_icmsFECPUFREMETTotal > 0 Then
            
               rdoCNLoja.BeginTrans
               
               wSequenciaS = wSequenciaS + 1
                            
               Sql = ""
                Sql = ""
                Sql = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja," & vbNewLine & _
                      "CNF_Serie,CNF_NF," & vbNewLine & _
                      "CNF_Sequencia,CNF_Carimbo," & vbNewLine & _
                      "CNF_TipoCarimbo,CNF_DetalheImpressao," & vbNewLine & _
                      "CNF_Data,CNF_SituacaoProcesso," & vbNewLine & _
                      "CNF_DataProcesso) " & _
                      "Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "'," & vbNewLine & _
                      "'" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      "," & wSequenciaS & " ," & "'FCP Total: " & Format(wVICMSFECP, "#####0.00") & " - ICMS Difer. origem: " & Format(wIE_icmsFECPUFREMETTotal, "#####0.00") & "   ICMS Difer. destino: " & Format(wIE_icmsFECPUFDESTTotal, "#####0.00") & "    Total: " & Format(wIE_icmsFECPUFREMETTotal + wIE_icmsFECPUFDESTTotal, "#####0.00") & "'" & " , " & vbNewLine & _
                      "'S',' '," & vbNewLine & _
                      "'" & Format(Date, "yyyy/mm/dd") & "'," & vbNewLine & _
                      "'A','" & Format(Date, "yyyy/mm/dd") & "')"
                
                rdoCNLoja.Execute (Sql)

                If Err.Number = 0 Then
                  rdoCNLoja.CommitTrans
                Else
                 rdoCNLoja.RollbackTrans
                End If
            End If

            If RsCapaNF("desconto") > 0 Then
            
               rdoCNLoja.BeginTrans
                            
               Sql = ""
               Sql = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                      " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      ",1,' DESCONTO:    " & Format(RsCapaNF("desconto"), "#####0.00") & _
                      "' , 'Z',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (Sql)

                If Err.Number = 0 Then
                  rdoCNLoja.CommitTrans
                Else
                 rdoCNLoja.RollbackTrans
                End If
            End If
            
            If wST20 = "S" Then
            
                wSequenciaS = wSequenciaS + 1
                  
                Sql = ""
                Sql = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999991'"
                RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                  
                 rdoCNLoja.BeginTrans
                  
                Sql = ""
                Sql = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                  " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                  "," & wSequenciaS & " ,'" & RsCarimbo("CE_linha12") & "' , 'S',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (Sql)
                RsCarimbo.Close
            
                If Err.Number = 0 Then
                    rdoCNLoja.CommitTrans
                Else
                    rdoCNLoja.RollbackTrans
                End If
                
            End If
            
            If wST60 = "S" Then
            
                wSequenciaS = wSequenciaS + 1
                            
                Sql = ""
                Sql = "Select CE_linha12 from CarimbosEspeciais where ce_Referencia = '9999992' "
                RsCarimbo.CursorLocation = adUseClient
                RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
                rdoCNLoja.BeginTrans
            
                Sql = ""
                Sql = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                      " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      "," & wSequenciaS & " ,'" & RsCarimbo("CE_linha12") & "' , 'S',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (Sql)
            
                If Err.Number = 0 Then
                    rdoCNLoja.CommitTrans
                Else
                    rdoCNLoja.RollbackTrans
                End If
            
                RsCarimbo.Close
                
                End If
                    
           Else
               MsgBox "Não foi possível acessar os carimbos fiscais", vbCritical, "AVISO"
               RsItensNF.Close
               RsCapaNF.Close
               Exit Function

          End If
          
          If tipoCupomEmite Like "CE*" Then
                
                Sql = ""
                Sql = "Select top 1 ve_codigo as codigoVendedor, ve_nome as nome from nfcapa, vende where vendedor = ve_codigo and numeroped = " & NumeroDocumento
                RsCarimbo.CursorLocation = adUseClient
                RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
                rdoCNLoja.BeginTrans
            
                Sql = ""
                Sql = " Insert into CarimboNotafiscal (CNF_NumeroPed,CNF_Loja,CNF_Serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_DetalheImpressao,CNF_Data,CNF_SituacaoProcesso,CNF_DataProcesso) " & _
                      " Values(" & RsCapaNF("NumeroPed") & ",'" & Trim(wLoja) & "','" & RsCapaNF("Serie") & "'," & RsCapaNF("nf") & _
                      "," & 0 & " ,'" & "Pedido: " & RsCapaNF("NumeroPed") & ", Vendedor: " & RsCarimbo("codigoVendedor") & " - " & RsCarimbo("nome") & "" & "' , 'S',' ','" & Format(Date, "yyyy/mm/dd") & "','A','" & Format(Date, "yyyy/mm/dd") & "')"
                rdoCNLoja.Execute (Sql)
            
                If Err.Number = 0 Then
                    rdoCNLoja.CommitTrans
                Else
                    rdoCNLoja.RollbackTrans
                End If
            
                RsCarimbo.Close
          
          End If
          
            Sql = ""
            Sql = "select count(*) as somacarimbo from carimbonotafiscal where cnf_Loja = '" & RsCapaNF("Lojaorigem") & "' and CNF_NF = " & RsCapaNF("nf") & _
                         " and cnf_serie = '" & RsCapaNF("serie") & "' " & _
                         " and CNF_NumeroPed = " & RsCapaNF("numeroped") & " "
            RsCarimbo.CursorLocation = adUseClient
            RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            If Not RsCarimbo.EOF Then
               wTotalCarimbo = RsCarimbo("somacarimbo")
            End If
            RsCarimbo.Close
            
          
            Sql = ""
            Sql = "select * from carimbonotafiscal where cnf_Loja = '" & RsCapaNF("Lojaorigem") & "' and CNF_NF = " & RsCapaNF("nf") & _
                  " and cnf_serie = '" & RsCapaNF("serie") & "' " & _
                  " and CNF_NumeroPed = " & RsCapaNF("numeroped") & _
                  " order by cnf_tipocarimbo desc, cnf_sequencia asc"
            RsCarimbo.CursorLocation = adUseClient
            RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

            If Not RsCarimbo.EOF Then
                
                wRestoItens = ((RsCapaNF("QTDItem")) Mod 12)
                wTotalLinha = ((wTotalCarimbo + wRestoItens) + 1)
                wContCarimbo = 0
                  
                If wRestoItens <> 0 Then
                   wContLinha = (wRestoItens + 1)
                End If
                                            
                Do While Not RsCarimbo.EOF
                   wContLinha = wContLinha + 1
                   wContCarimbo = wContCarimbo + 1
                   If (wContLinha Mod 12) <> 0 Then
                       wDetalheImpressao = "D"
                   Else
                       wDetalheImpressao = "C"
                       wUltimoItem = wUltimoItem + 1
                   End If
                
                   
                   If wTotalLinha = wContLinha Then
                       wDetalheImpressao = "T"
                   End If
                   
                   If wTotalCarimbo = wContCarimbo Then
                      wDetalheImpressao = "T"
                   End If

                   rdoCNLoja.BeginTrans
                   
                   Sql = ""
                   Sql = "update CarimboNotafiscal set CNF_DetalheImpressao = '" & wDetalheImpressao & "', cnf_data = '" & Format(RsCapaNF("dataemi"), "yyyy/mm/dd") & "'" & _
                         " where cnf_Loja = '" & RsCarimbo("cnf_Loja") & "' and cnf_nf = " & RsCarimbo("cnf_nf") & _
                         " and cnf_serie = '" & RsCarimbo("cnf_serie") & "' and cnf_tipocarimbo = '" & RsCarimbo("cnf_tipocarimbo") & "' " & _
                         " and cnf_sequencia = '" & RsCarimbo("cnf_sequencia") & "' " & _
                         " and CNF_NumeroPed = " & RsCapaNF("numeroped") & " "
                         rdoCNLoja.Execute (Sql)
                         
                   If Err.Number = 0 Then
                        rdoCNLoja.CommitTrans
                   Else
                        rdoCNLoja.RollbackTrans
                   End If
                   
                   RsCarimbo.MoveNext
                Loop
             End If
             RsCarimbo.Close
             
'-------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
             
             Sql = "update nfitens set BaseICMS = 0 where BaseICMS is null and numeroped = " & NumeroDocumento
             rdoCNLoja.Execute (Sql)
             
             Sql = "Select sum(BASEICMS) as BaseICMS from nfitens where numeroped = " & NumeroDocumento
             RsCarimbo.CursorLocation = adUseClient
             RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
             Sql = "select top 1 CFOP from nfitens where numeroped = " & NumeroDocumento
             RsItensNF.CursorLocation = adUseClient
             RsItensNF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             rdoCNLoja.BeginTrans

'''             sql = "UPDATE nfcapa set " _
'''                & "Paginanf = " & ConverteVirgula(wUltimoItem) & ",BaseICMS = " _
'''                & ConverteVirgula(RsCarimbo("BaseICMS")) & ", " _
'''                & "ECF  = " & GLB_ECF & ",TipoNota = 'V' , CodOper = " & RsItensNF("CFOP") & ", " _
'''                & "CFOAUX = " & RsItensNF("CFOP") & " " _
'''                & "where nfcapa.numeroped = " & NumeroDocumento & ""
'''                rdoCNLoja.Execute (sql)
                
             Sql = "UPDATE nfcapa set " _
                & "Paginanf = " & ConverteVirgula(wUltimoItem) & ",BaseICMS = " _
                & ConverteVirgula(RsCarimbo("BaseICMS")) & ", " _
                & "valorICMSFECP = '" & ConverteVirgula(wTotalVICMSFECP) & "', " _
                & "valICMSRemet = '" & ConverteVirgula(wIE_icmsFECPUFREMETTotal) & "', " _
                & "valICMSDest = '" & ConverteVirgula(wIE_icmsFECPUFDESTTotal) & "', " _
                & "ECF  = " & GLB_ECF & ", CodOper = " & RsItensNF("CFOP") & " " _
                & "where nfcapa.numeroped = " & NumeroDocumento & ""
                rdoCNLoja.Execute (Sql)

                
             If Err.Number = 0 Then
                rdoCNLoja.CommitTrans
             Else
                rdoCNLoja.RollbackTrans
             End If
             
             RsItensNF.Close
             RsCarimbo.Close
                
'--------------------------------------  ATUALIZA ESTOQUE LOJA ----------------------------------------------------
             
            'rdoCNLoja.BeginTrans
             
            'sql = ""
            'sql = "UPDATE EstoqueLoja Set EL_Estoque = (EL_Estoque - QTDE) FROM NFItens, EstoqueLoja " _
                 & "Where EL_Referencia = Referencia and NumeroPed = " & NumeroDocumento
             
            'rdoCNLoja.Execute sql
            
            'If Err.Number = 0 Then
               'rdoCNLoja.CommitTrans
            'Else
               'rdoCNLoja.RollbackTrans
            'End If
        
Exit Function
ErroEncerraTransferencia:
     MsgBox Err.Number & " - " & Err.Description & vbLf & _
           "Não foi possível encerrar a nota fiscal de venda.", vbCritical, "AVISO"

    RsItensNF.Close
    RsCapaNF.Close

    Exit Function


    RsItensNF.Close
    RsCapaNF.Close
'End Sub
            
    
End Function


Sub PegaNumeroRomaneio()

        Screen.MousePointer = 11
        
        Sql = "Select * from ControleSistema "
        
        rdocontrole.CursorLocation = adUseClient
        rdocontrole.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
        NroNotaFiscal = rdocontrole("CTS_Numero00") + 1
        rdocontrole.Close
                 
        rdoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        
        Sql = "Update ControleSistema set CTS_Numero00 =" & NroNotaFiscal
        
        rdoCNLoja.Execute Sql
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
     
            
         rdoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          
          Sql = "Update nfcapa set NF = " & NroNotaFiscal _
              & ", Serie = '00' Where NumeroPed = " & pedido _
              & " and tiponota = 'PA'"
          
          rdoCNLoja.Execute Sql
          Screen.MousePointer = vbNormal
          rdoCNLoja.CommitTrans
          
          rdoCNLoja.BeginTrans
          Screen.MousePointer = vbHourglass
          
          Sql = "Update nfitens set NF = " & NroNotaFiscal _
              & ", Serie = '00' Where NumeroPed = " & pedido _
              & " and tiponota = 'PA'"
          
          
          rdoCNLoja.Execute Sql
          Screen.MousePointer = vbNormal
          rdoCNLoja.CommitTrans
End Sub


Function AcharICMSInterEstadual(ByVal Referencia As String, ByVal ChaveIcms As Double) As Boolean
    
    wIE_icmsAplicado = 0
    wIE_Tributacao = 0
    wIE_Cfo = 0
    wIE_BasedeReducao = 0
    wIE_icmsdestino = 0
    
    Sql = "SELECT * from IcmsInterEstadual where IE_Codigo = " & ChaveIcms
       
    RsICMSIntER.CursorLocation = adUseClient
    RsICMSIntER.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
           
    If RsICMSIntER.EOF Then
        AcharICMSInterEstadual = False
        RsICMSIntER.Close
        Exit Function
    Else
        AcharICMSInterEstadual = True
    End If
    
    wIE_icmsAplicado = RsICMSIntER("IE_icmsAplicado")
    wIE_Tributacao = RsICMSIntER("IE_CST")
    wIE_Cfo = RsICMSIntER("IE_Cfop")
    wIE_BasedeReducao = RsICMSIntER("IE_BasedeReducao")
    wIE_icmsdestino = RsICMSIntER("IE_icmsdestino")
    RsICMSIntER.Close
    
    If wIE_icmsFECPAplicado > 0 And wIE_icmsAplicado <> 0 Then
        wIE_icmsAplicado = wIE_icmsFECPAplicado
    End If
    
End Function


Function ConsistenciaNota(ByVal pedido As Double, ByVal Serie As String) As Boolean
    
    
    Sql = ""
    Sql = "Select count(NfItens.Referencia) as QuantRef, NfCapa.QtdItem from NfCapa,NfItens " _
        & "where NfCapa.NumeroPed=" & pedido & " " _
        & "and NfItens.NumeroPed=NfCapa.NumeroPed " _
        & "Group by NfCapa.QtdItem " _
        & "having Count(NfItens.Referencia) = NfCapa.QtdItem"
    
    rsItemNota.CursorLocation = adUseClient
    rsItemNota.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not rsItemNota.EOF Then
        ConsistenciaNota = True
    Else
        MsgBox "A nota não pode ser impressa porque exite um erro com a quantidade de itens ", vbCritical, "Atenção"
        ConsistenciaNota = False
    End If
    rsItemNota.Close
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
    Dim I As Integer
    
    'Valida argumento
    If Len(CPF) <> 11 Then
        FU_ValidaCPF = False
        Exit Function
    End If

        
    
    soma = 0
    For I = 1 To 9
        soma = soma + Val(Mid$(CPF, I, 1)) * (11 - I)
    Next I
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
        
    soma = 0
    For I = 1 To 10
        soma = soma + Val(Mid$(CPF, I, 1)) * (12 - I)
    Next I
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        FU_ValidaCPF = False
        Exit Function
    End If
    
    FU_ValidaCPF = True

End Function

Function FU_ValidaCGC(CGC As String) As Integer
        Dim Retorno, a, j, I, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For I = 1 To 7
               a = Val(Mid(CGC, I, 1))
               If (I Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next I
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              FU_ValidaCGC = True
           Else
              FU_ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              I = 0
              d1 = 0
              d2 = 0
              j = 5
              For I = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              I = 0
              j = 6
              For I = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
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


Public Function VerificaRetornoImpressora(Label As String, RetornoFuncao As String, TituloJanela As String)
    
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim RetornaMensagem As Integer
    Dim StringRetorno As String
    Dim ValorRetorno As String
    Dim RetornoStatus As Integer
    Dim Mensagem As String
    
    wVerificaImpressoraFiscal = False
    
    If Retorno = 0 Then
        MsgBox "Erro de comunicação com a impressora.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    
    ElseIf Retorno = 1 Then
        RetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ValorRetorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
        
        If Label <> "" And RetornoFuncao <> "" Then
            RetornaMensagem = 1
        End If
        
        If ACK = 21 Then
            
            Exit Function
        End If
        
        If (ST1 <> 0 Or ST2 <> 0) Then
                If (ST1 >= 128) Then
                    StringRetorno = "Fim de Papel" & vbCr
                    ST1 = ST1 - 128
                End If
                
                If (ST1 >= 64) Then
                    StringRetorno = StringRetorno & "Pouco Papel" & vbCr
                    ST1 = ST1 - 64
                End If
                
                If (ST1 >= 32) Then
                    StringRetorno = StringRetorno & "Erro no relógio" & vbCr
                    ST1 = ST1 - 32
                End If
                
                If (ST1 >= 16) Then
                    StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                    ST1 = ST1 - 16
                End If
                    
                If (ST1 >= 8) Then
                    StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
                    ST1 = ST1 - 8
                End If
                
                If (ST1 >= 4) Then
                    StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                    ST1 = ST1 - 4
                End If
                    
                If (ST1 >= 2) Then
                    StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                    ST1 = ST1 - 2
                End If
                
                If (ST1 >= 1) Then
                    StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
                    ST1 = ST1 - 1
                End If
                    
                If (ST2 >= 128) Then
                    StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
                    ST2 = ST2 - 128
                End If
                
                If (ST2 >= 64) Then
                    StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
                    ST2 = ST2 - 64
                End If
                
                If (ST2 >= 32) Then
                    StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                    ST2 = ST2 - 32
                End If
                
                If (ST2 >= 16) Then
                    StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
                    ST2 = ST2 - 16
                End If
                    
                If (ST2 >= 8) Then
                    StringRetorno = StringRetorno & "Capacidade de alíquota programáveis lotada" & vbCr
                    ST2 = ST2 - 8
                End If
                
                If (ST2 >= 4) Then
                    StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
                    ST2 = ST2 - 4
                End If
                    
                If (ST2 >= 2) Then
                    StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
                    ST2 = ST2 - 2
                End If
                
                If (ST2 >= 1) Then
                   
                    ST2 = ST2 - 1
                End If
                
                If RetornaMensagem Then
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                           vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                           Label & RetornoFuncao
                Else
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                       vbCr & vbLf & StringRetorno
                End If
                wValorRetorno = ValorRetorno
                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
                Exit Function
        End If 'fim do ST1 <> 0 and ST2 <> 0
        
        If RetornaMensagem Then
            Mensagem = Label & RetornoFuncao
        End If
        
        If Mensagem <> "" Then
            MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
        End If
        Exit Function
    ElseIf Retorno = -1 Then
        MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    ElseIf Retorno = -2 Then
        MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -3 Then
        MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
        Retorno = Bematech_FI_ResetaImpressora()
        Exit Function
    ElseIf Retorno = -6 Then
        MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -7 Then
        MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    ElseIf Retorno = -8 Then
        MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    End If
    wVerificaImpressoraFiscal = True
   
End Function



Function ConverteVirgula(ByVal numero As String) As String

    Dim Ret As String
    Dim CharLido As String
    Dim Maximo As Long
    Dim I As Long
    
    Ret = ""
    numero = IIf(IsNull(numero), 0, numero)
    Maximo = Len(numero)
    
    For I = 1 To Maximo
        CharLido = Mid(numero, I, 1)
        If IsNumeric(CharLido) Then
            Ret = Ret & CharLido
        ElseIf CharLido = "," And InStr(Ret, ".") = 0 Then
            Ret = Ret & "."
        End If
    Next
    
    ConverteVirgula = Ret

End Function

Function ReplaceVirgula(ByVal numero As String) As String

    Dim Ret As String
    Dim CharLido As String
    Dim Maximo As Long
    Dim I As Long
    
    Ret = "0"
    numero = IIf(IsNull(numero), 0, numero)
    Maximo = Len(numero)
    
    For I = 1 To Maximo
        CharLido = Mid(numero, I, 1)
        
        
        If IsNumeric(CharLido) Then
            Ret = Ret & CharLido
        ElseIf CharLido = "," And InStr(Ret, ".") = 0 Then
            Ret = Ret & ""
        End If
    Next
    
    ReplaceVirgula = Ret

End Function


Sub Main()
 
 
'On Error GoTo ConexaoErro
 
wErroApresenta = 0
Call verificaAppExecucao
 
If adoCNAccess.State = 1 Then
     adoCNAccess.Close
End If
 
 
lsDSN = "Driver={Microsoft Access Driver (*.mdb)};" & _
          "Dbq=c:\sistemas\DMACini.mdb;" & _
          "Uid=Admin; Pwd=astap36"
  adoCNAccess.Open lsDSN
  Sql = "Select * from ConexaoTEF"
  rdoConexaoINI.CursorLocation = adUseClient
  rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
 
        If Not rdoConexaoINI.EOF Then
           
           GLB_ServidorTEF = Trim(rdoConexaoINI("TEF_Servidor"))
           GLB_BancoTEF = Trim(rdoConexaoINI("TEF_Banco"))
         End If
  rdoConexaoINI.Close
   
  Sql = "Select count(*) as QtdeDeLojasINI from ConexaoSistema"
   
  rdoConexaoINI.CursorLocation = adUseClient
  rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
 
        If Not rdoConexaoINI.EOF Then
            
          Sql = "Select * from ParametroSistema"
          rdoParametroINI.CursorLocation = adUseClient
          rdoParametroINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
           
          If Not rdoParametroINI.EOF Then
             GLB_ECF = Trim(rdoParametroINI("CXA_ECF"))
             GLB_Caixa = Trim(rdoParametroINI("CXA_NumeroCaixa"))
             Glb_ImpNotaFiscal = Trim(rdoParametroINI("GLB_ImpNotaFiscal"))
             GLB_Impressora00 = Trim(rdoParametroINI("GLB_imp00"))
             Glb_AlteraResolucao = rdoParametroINI("GLB_AlteraResolucao")
             
             rdoParametroINI.Close
           Else
             MsgBox "Problemas no banco de dados de inicializacao", vbCritical, "Aviso"
             rdoParametroINI.Close
             rdoConexaoINI.Close
             End
             Exit Sub
           End If
           
           
           If rdoConexaoINI("QtdeDeLojasINI") = 1 Then
           
              rdoConexaoINI.Close
              
              Sql = "Select * from ConexaoSistema"
                     rdoConexaoINI.CursorLocation = adUseClient
                     rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
                     
                     If Not rdoConexaoINI.EOF Then
                        GLB_Servidor = Trim(rdoConexaoINI("GLB_ServidorRetaguarda"))
                        GLB_Loja = Trim(rdoConexaoINI("GLB_Loja"))
                        GLB_Banco = Trim(rdoConexaoINI("GLB_BancoRetaguarda"))
                        GLB_Servidorlocal = Trim(rdoConexaoINI("GLB_ServidorLocal"))
                        Glb_BancoLocal = Trim(rdoConexaoINI("GLB_BancoLocal"))
                        'GLB_EnderecoPortal = Trim(rdoConexaoINI("GLB_Portal"))
                        GLB_EnderecoPastaRESP = Trim(rdoConexaoINI("GLB_EnderecoResp"))
                        GLB_EnderecoPastaFIL = Trim(rdoConexaoINI("GLB_EnderecoFil"))
                        
                        rdoConexaoINI.Close
                        
                     Else
                        MsgBox "Problemas no banco de dados de inicializacao", vbCritical, "Aviso"
                        rdoConexaoINI.Close
                        End
                        Exit Sub
                     End If
                     
           Else
              rdoConexaoINI.Close
          
              frmInicio.Show
              Exit Sub
           End If
        Else
           MsgBox "Banco de dados de inicializacao Vazio", vbCritical, "Aviso"
           End
           Exit Sub
        End If
 
ConectaODBC
    
Continua:
   
 
    If GLB_ConectouOK = True Then
       Call DadosLoja
        Sql = "Select * from ControleCaixa Where CTR_Supervisor = 99 and CTR_SituacaoCaixa='F' " _
           & "and CTR_datainicial >= '" & Format(Date, "yyyy/mm/dd") & "'"
            rsNFE.CursorLocation = adUseClient
            rsNFE.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            If Not rsNFE.EOF Then
               MsgBox "Fechamento Geral de hoje já foi efetuado. Não é possivel abrir o Caixa."
               wPermitirVenda = False
               rsNFE.Close
               'Unload Me
               Exit Sub
            End If



            rsNFE.Close
'       SQL = "Select ControleCaixa.*,USU_Codigo,USU_Nome from ControleCaixa,UsuarioCaixa" _
'            & " Where CTR_Supervisor <> 99 and CTR_Operador = USU_Codigo and CTR_SituacaoCaixa='A' and CTR_NumeroCaixa = " & GLB_Caixa
'
'
'            rsNFE.CursorLocation = adUseClient
'            rsNFE.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'            If rsNFE.EOF = False Then
'               GLB_USU_Nome = rsNFE("USU_Nome")
'               GLB_USU_Codigo = rsNFE("USU_Codigo")
'               GLB_CTR_Protocolo = rsNFE("CTR_Protocolo")
'               GLB_DataInicial = Format(rsNFE("CTR_DataInicial"), "YYYY/MM/DD")
'            End If
'
'            If rsNFE.EOF = False Then
'              If rsNFE("CTR_Situacaocaixa") = "A" And rsNFE("ctr_datainicial") < Date Then
'                 MsgBox "Data do caixa incorreta.Favor efetuar o Fechamento", vbCritical, "Atenção"
'                 wPermitirVenda = False
'
'                 frmControlaCaixa.cmdBotoesParte2(2).Visible = False
'              Else
'                 wPermitirVenda = True
'              End If
'
'            End If




            
            If carregaControleCaixa = False Then
               frmLoginCaixa.Show
            Else
               frmBandeja.Show
            End If
    Else
        MsgBox "Erro ao conectar-se ao Banco de Dados", vbCritical, "Atenção"
        Exit Sub
    End If
    
 
 
' Exit Sub
'ConexaoErro:
'MsgBox "Erro ao abrir banco de Dados da Loja! "
'End
   
 
   
  End Sub

  


Public Sub CancelaCupomFiscal()
    
    Retorno = Bematech_FI_CancelaCupom()
    Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
    

End Sub



Sub AtualizaNumeroCupom()

    Sql = ""
    Sql = "Update controleEcf set ct_ultimocupom= CT_UltimoCupom + 1 " _
        & "where CT_Ecf=" & Val(GLB_ECF) & ""
           rdoCNLoja.Execute (Sql)

End Sub


Sub AjustaTela(ByRef Formulario As Form)

  Formulario.top = frmControlaCaixa.webPadraoTamanho.top
  Formulario.left = frmControlaCaixa.webPadraoTamanho.left
  Formulario.Width = frmControlaCaixa.webPadraoTamanho.Width
  Formulario.Height = frmControlaCaixa.webPadraoTamanho.Height

End Sub

Sub GravaValorCarrinho(ByRef Formulario As Form, ByRef TotalItem As String)

  If TotalItem <> "" Then

'    frmControlaCaixa.webInternet1.Picture = LoadPicture(endIMG("topoCarinho1024768hd"))
    frmControlaCaixa.cmdTotalItens.Caption = Formulario.lblTotalItens.Caption
    frmControlaCaixa.cmdTotalVenda.Caption = Formulario.lblTotalvenda.Caption
    frmControlaCaixa.cmdTotalPedidoGE.Caption = Formulario.lblTotalGarantia.Caption
    If frmControlaCaixa.cmdTotalPedidoGE.Caption = "+ G.E " & "0,00" Then
        frmControlaCaixa.cmdTotalPedidoGE.Visible = False
    Else
        frmControlaCaixa.cmdTotalPedidoGE.Visible = True
    End If
            
    'Do While Len(frmControlaCaixa.cmdTotalVenda.Caption) <= 12
       frmControlaCaixa.cmdTotalVenda.Caption = frmControlaCaixa.cmdTotalVenda.Caption
    'Loop
    
    'Do While Len(frmControlaCaixa.cmdTotalItens.Caption) <= 5
       frmControlaCaixa.cmdTotalItens.Caption = frmControlaCaixa.cmdTotalItens.Caption
    'Loop
    
  Else
      
'    frmControlaCaixa.webInternet1.Picture = LoadPicture(endIMG("topo1024768hd"))
'    frmControlaCaixa.webInternet1.Play
    frmControlaCaixa.cmdTotalItens.Caption = ""
    frmControlaCaixa.cmdTotalVenda.Caption = ""
    frmControlaCaixa.cmdTotalPedidoGE.Caption = ""
    

  End If
       

End Sub


Function TiraZero(ByVal numero As String) As String


 If Val(numero) = 0 Then
 TiraZero = ""
Else
 TiraZero = numero
End If

End Function


Public Function Replace(Source As String, Find As String, ReplaceStr As String, _
    Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
    Dim findLen As Long
    Dim replaceLen As Long
    Dim Index As Long
    Dim counter As Long
    
    findLen = Len(Find)
    replaceLen = Len(ReplaceStr)
    If findLen = 0 Then Err.Raise 5
    
    If Start < 1 Then Start = 1
    Index = Start
    
    Replace = Source
    
    Do
        Index = InStr(Index, Replace, Find, Compare)
        If Index = 0 Or Count = 0 Then Exit Do
        If findLen = replaceLen Then
            Mid$(Replace, Index, findLen) = ReplaceStr
        Else
            Replace = left$(Replace, Index - 1) & ReplaceStr & Mid$(Replace, _
            Index + findLen)
        End If
        Index = Index + replaceLen
        counter = counter + 1
    Loop Until counter = Count
    
    If Start > 1 Then Replace = Mid$(Replace, Start)
    
End Function

Function ConcatenaCarimboNF(NroPedido As Long)

   Dim Carimbo As String
   Carimbo = ""
   Sql = ""
   Sql = "select cnf_Carimbo from carimbonotafiscal where cnf_numeroped = " & NroPedido & _
         " order by cnf_tipocarimbo desc, cnf_sequencia asc"
   RsCarimbo.CursorLocation = adUseClient
   RsCarimbo.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

   Do While Not RsCarimbo.EOF
   
       Carimbo = Carimbo + ", " + RTrim(LTrim(RsCarimbo("cnf_Carimbo")))
       RsCarimbo.MoveNext
   
   Loop
   RsCarimbo.Close
   ConcatenaCarimboNF = Carimbo
   
End Function

Function BuscaQtdeViaImpressaoMovimento()
   
    Sql = "Select CTS_QtdeViaMovimento from controlesistema"
    rdocontrole.CursorLocation = adUseClient
    rdocontrole.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    wQdteViasImpressao = IIf(IsNull(rdocontrole("CTS_QtdeViaMovimento")), 1, rdocontrole("CTS_QtdeViaMovimento"))
    rdocontrole.Close
    
End Function


Function BuscaQtdeViaImpressaoRomaneio()

    Sql = "Select CTS_QtdeViaRomaneio from controlesistema"
    rdocontrole.CursorLocation = adUseClient
    rdocontrole.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    wQdteViasImpressao = IIf(IsNull(rdocontrole("CTS_QtdeViaMovimento")), 1, rdocontrole("CTS_QtdeViaMovimento"))
    rdocontrole.Close
    
End Function

Function InserirPreVenda(NroPedido As Long)
  Dim wCGC As String
  Dim wDescontoTEF As Integer
       
  GLB_CodigoNFTef = ""
  wDescontoTEF = 0
  
  Sql = "Select (totalnota + desconto) as totalnota,cpfnfp,desconto from nfcapa " _
      & "where numeroped = " & NroPedido
  rsTEF.CursorLocation = adUseClient
  rsTEF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If Not rsTEF.EOF Then

      Sql = "select count(*) as QtdePag from movimentocaixa where mc_grupo like '10%' and mc_pedido = " & NroPedido
      rdocontrole.CursorLocation = adUseClient
      rdocontrole.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

      If rdocontrole("QtdePag") > 1 Then
         wPagamentoECF = 0
      End If
      
      
      If ConverteVirgula(rsTEF("Desconto")) > 0 Then
         wDescontoTEF = 1
      End If
      
      wCGC = Val(rsTEF("cpfnfp"))
     
      Sql = ""
      Sql = "Exec InserirPreVenda '" & NroPedido & "','" & Format(Date, "yyyy/mm/dd") & "', " _
            & ConverteVirgula(Format(rsTEF("Totalnota"), "0.00")) & "," & wDescontoTEF & "," & wDescontoTEF & "," _
            & ConverteVirgula(rsTEF("Desconto")) & ",0,0,0,'" _
            & wCGC & "'," & wPagamentoECF & ",0,0"

      rdoCNTEF.Execute Sql
      

''    @Referencia AS VARCHAR(6),
''    @DataHoraReferencia AS DATETIME,
''    @ValorTotal AS FLOAT,
''    @DescontoOuAcrescimo AS SMALLINT,   --1
''    @PorcentagemOuAbsoluto AS SMALLINT, --1
''    @ValorDescontoAbsoluto AS FLOAT,    --- valor desconto
''    @ValorDescontoPorcentagem AS FLOAT,
''    @ValorAcrescimoAbsoluto AS FLOAT,
''    @ValorAcrescimoPorcentagem AS FLOAT,
''    @CPFCNPJ AS VARCHAR(14),
''    @CodigoFormaPagamento AS SMALLINT = 0, --NOVO
''    @Sequencial AS INT OUT,
''    @Retorno AS INT OUT

  Else
     
     MsgBox "Pedido inexistente"
  
  End If
  
  rsTEF.Close
  rdocontrole.Close
End Function

Function InserirItemPreVenda(NroPedido As Long)
  Dim wSequencia As Integer
  Dim wicmstef As Double
  Dim situacaotributaria As String

  wSequencia = 0
  
  Sql = ""
  Sql = "Select sequencial,valortotalitens from prevenda where referencia = " & NroPedido
  rsTEF.CursorLocation = adUseClient
  rsTEF.Open Sql, rdoCNTEF, adOpenForwardOnly, adLockPessimistic
  wSequencia = rsTEF("sequencial")
  rsTEF.Close
  
  Sql = ""
  Sql = "Select referencia,qtde,pr_descricao,pr_icmpdv,pr_icmspdvsaidaiva,pr_st,PR_SubstituicaoTributaria, " & _
        "vlunit as preco " & _
        "from nfitens,produtoloja where pr_referencia = referencia " & _
        "and numeroped = " & NroPedido
  
  rsTEF.CursorLocation = adUseClient
  rsTEF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

 If Not rsTEF.EOF Then
     Do While Not rsTEF.EOF
       
       If rsTEF("pr_ST") = 60 Then
           situacaotributaria = "F"
           wicmstef = 0
       Else
           situacaotributaria = "T"
           wicmstef = rsTEF("pr_icmpdv")
       End If

       If rsTEF("PR_SubstituicaoTributaria") = "N" And wicmstef = 0 Then
           situacaotributaria = "F"
       End If

       Sql = ""
       Sql = "Exec AtualizarProdutoServico '0000000" & Trim(rsTEF("referencia")) & "','" & Trim(rsTEF("pr_descricao")) & "'," & _
             ConverteVirgula(Format(rsTEF("preco"), "0.00")) & ",0,'PC',0," & rsTEF("qtde") & "," & ConverteVirgula(wicmstef) & ",'" & situacaotributaria & "',0"
       rdoCNTEF.Execute Sql

       Sql = ""
       Sql = "Exec InserirItemPreVenda " & wSequencia & ",'0000000" & Trim(rsTEF("referencia")) & "'," & rsTEF("qtde") & ",0,0,0,0,0,0,0"
       rdoCNTEF.Execute Sql
       rsTEF.MoveNext
     Loop
  Else
     MsgBox "Não há itens para este pedido"
  End If
  rsTEF.Close

  
End Function

Function InserirPagamentoPreVenda(NroPedido As Long)
Dim wSequencia As Integer
Dim wSeqPagamento As Integer
  
  Sql = "select count(*) as QTDEPag from movimentocaixa " _
           & "Where mc_grupo like '10%' and mc_pedido = " & NroPedido
      rsTEF.CursorLocation = adUseClient
      rsTEF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  
  If rsTEF("QTDEPag") = 1 Then
      rsTEF.Close
      Exit Function
  End If
  rsTEF.Close
  
  wSequencia = 0
  wSeqPagamento = 0
  
  Sql = ""
  Sql = "Select sequencial from prevenda where referencia = " & NroPedido
  rsTEF.CursorLocation = adUseClient
  rsTEF.Open Sql, rdoCNTEF, adOpenForwardOnly, adLockPessimistic
  
  wSequencia = rsTEF("sequencial")
  rsTEF.Close
  

       Sql = "select fpt_CodigoTEF,MC_Valor from formapagamentotef,modalidade,movimentocaixa " _
           & "Where RTrim(LTrim(fpt_condicao)) = RTrim(LTrim(mo_descricao)) " _
           & "and mo_grupo like '10%' and mo_grupo =  mc_grupo and mc_pedido = " & NroPedido
      rsTEF.CursorLocation = adUseClient
      rsTEF.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

      If Not rsTEF.EOF Then
         Do While Not rsTEF.EOF
            wSeqPagamento = wSeqPagamento + 1
         
            Sql = "Exec InserirFormaPagamentoPreVenda " & rsTEF("FPT_CodigoTEF") & "," & wSequencia & "," _
                & wSeqPagamento & "," _
                & ConverteVirgula(rsTEF("MC_Valor")) & ",0"
            rdoCNTEF.Execute Sql
            rsTEF.MoveNext
          Loop
      Else
         MsgBox "Pagamento não encontrado - TEF"
         rsTEF.Close
         Exit Function
      End If
  
  rsTEF.Close

End Function



Function BuscaNroCupomFiscal(NroPedido As Long)
Dim wCont As Integer
Screen.MousePointer = 11
  wCont = 0
  frmFormaPagamento.Enabled = False
  Do While wCont < 10
     
     Sql = "select NumeroCupomReferencia from LogOperacao where ReferenciaPreVenda = " & NroPedido
     rsTEF.CursorLocation = adUseClient
     rsTEF.Open Sql, rdoCNTEF, adOpenForwardOnly, adLockPessimistic
  
     If Not rsTEF.EOF Then

        Sql = "update nfcapa set serie = '" & GLB_SerieCF & "', nf = " & rsTEF("NumeroCupomReferencia") & " where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql
  
        Sql = "update nfitens set serie = '" & GLB_SerieCF & "',nf = " & rsTEF("NumeroCupomReferencia") & " where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql
  
        Sql = "update movimentocaixa set mc_serie = '" & GLB_SerieCF & "',mc_documento = " & rsTEF("NumeroCupomReferencia") & " where mc_pedido = " & NroPedido
        rdoCNLoja.Execute Sql
        
        rsTEF.Close
        frmFormaPagamento.Enabled = True
        Screen.MousePointer = 0
        Exit Function
     End If
     
     rsTEF.Close
     Esperar 3
     wCont = wCont + 1

     If wCont < 10 Then

        
       If MsgBox("Número TEF não encontrado. Tentar novamente? ", vbQuestion + vbYesNo, "Atenção") = vbNo Then
         'frmFormaPagamento.Enabled = True
         Screen.MousePointer = 11

         
        Sql = "update nfcapa set serie = null, tiponota = 'PA' , nf = null where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql
  
        Sql = "update nfitens set serie = null, tiponota = 'PA' , nf = null where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql
  
        Sql = "delete movimentocaixa where mc_pedido = " & NroPedido
        rdoCNLoja.Execute Sql
        
        Sql = "delete carimbonotafiscal where cnf_numeroped =" & NroPedido
        
        frmFormaPagamento.txtTipoNota = "PA"
        Screen.MousePointer = 0
        Exit Function
           
''          If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.Text) = "FRMCAIXATEF" Then
''            frmCaixaTEF.txtCodigoProduto = ""
''            frmCaixaTEF.txtCGC_CPF.Text = ""
''            LimpaGrid frmCaixaTEF.grdItens
''            frmCaixaTEF.grdItens.Rows = 1
''            wItens = 0
''            frmCaixaTEF.lblTotalvenda.Caption = ""
''            frmCaixaTEF.lblTotalItens.Caption = ""
''            Call GravaValorCarrinho(frmCaixaTEF, frmCaixaTEF.lblTotalItens.Caption)
''            frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = ""
''            frmCaixaTEF.cmdTotalVenda.Caption = ""
''            frmCaixaTEF.cmdItens.Caption = ""
''            frmCaixaTEF.lblDescricaoProduto.Caption = ""
''            frmCaixaTEF.fraProduto.Visible = False
''            frmCaixaTEF.fraNFP.Visible = True
''          End If
''
''          If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.Text) = "FRMCAIXATEFPEDIDO" Then
''            LimpaGrid frmCaixaTEFPedido.grdItens
''            frmCaixaTEFPedido.grdItens.Rows = 1
''            frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = ""
''            frmCaixaTEFPedido.lblTotalvenda.Caption = ""
''            frmCaixaTEFPedido.lblTotalItens.Caption = ""
''            Call GravaValorCarrinho(frmCaixaTEFPedido, frmCaixaTEFPedido.lblTotalItens.Caption)
''            frmCaixaTEFPedido.fraNFP.Visible = False
''            frmCaixaTEFPedido.txtPedido.Text = ""
''            frmCaixaTEFPedido.txtCGC_CPF.Text = ""
''            frmCaixaTEFPedido.fraPedido.Visible = True
''          End If
''
''          Call ZeraVariaveis
''            frmFormaPagamento.fraRecebimento.Visible = False
''            frmFormaPagamento.lblTotalPedido.Visible = False
''            frmFormaPagamento.lblValorTotalPedido.Visible = False
''            frmFormaPagamento.lblTootip.Text = ""
''            frmFormaPagamento.chbOkPag.Enabled = False
''            Unload Me
''
''            Unload frmCaixaTEF
''            Unload frmCaixaTEFPedido
''

''         Exit Function
         
        End If
       End If

     wCont = 0
     
  Loop
  Screen.MousePointer = 0
End Function

Public Sub VerificaSeEmiteCupom()

    Dim Sql As String

    Sql = "Select upper(CS_SerieCF) as serie from Controleserie"
    
    rdoSerie.CursorLocation = adUseClient
    rdoSerie.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rdoSerie.EOF Then
        tipoCupomEmite = rdoSerie("serie")
        If tipoCupomEmite Like "CF*" Then
            tipoCupomEmite = "CF"
        ElseIf tipoCupomEmite Like "CE*" Then
            tipoCupomEmite = "CE"
        Else
            tipoCupomEmite = ""
        End If
    Else
        tipoCupomEmite = ""
    End If
    rdoSerie.Close
    
    '

End Sub



Function BuscaCodigoPagamentoTEF(FormaTEF As String) As Integer

       Sql = ""
       Sql = "Select FPT_CodigoTEF from FormaPagamentoTEF where FPT_Condicao = '" & Trim(UCase(FormaTEF)) & "'"
       
       rdoSerie.CursorLocation = adUseClient
       rdoSerie.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rdoSerie.EOF Then
         If frmFormaPagamento.chbPOS.Value = 1 And rdoSerie("FPT_CodigoTEF") = 3 Then
           BuscaCodigoPagamentoTEF = 5
         ElseIf frmFormaPagamento.chbPOS.Value = 1 And rdoSerie("FPT_CodigoTEF") = 4 Then
           BuscaCodigoPagamentoTEF = 6
         Else
           BuscaCodigoPagamentoTEF = rdoSerie("FPT_CodigoTEF")
         End If
       Else
           BuscaCodigoPagamentoTEF = 1
       End If
       rdoSerie.Close

End Function


Function ConverteNFparaCF(NroPedido As Long)
''        Dim NroNF As Integer
        Sql = "update nfcapa set tiponota = 'PA' where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql

        Sql = "update nfitens set tiponota = 'PA' where numeroped = " & NroPedido
        rdoCNLoja.Execute Sql

        Sql = "delete movimentocaixa where mc_pedido = " & NroPedido
        rdoCNLoja.Execute Sql
          
        Sql = "delete CarimboNotaFiscal where cnf_numeroped = " & NroPedido
        rdoCNLoja.Execute Sql
        
        Sql = "Delete itemprevenda"
        rdoCNTEF.Execute Sql
        
        Sql = "Delete prevendaxformapagamento"
        rdoCNTEF.Execute Sql

        Sql = "Delete prevenda"
        rdoCNTEF.Execute Sql
        
End Function

Public Sub limpaGrid(ByRef GradeUsu)
    GradeUsu.Rows = GradeUsu.FixedRows
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows
End Sub

Public Sub ZeraVariaveis()
ValorPagamentoCartao = 0
ValDinheiro = 0
ValTroco = 0
ValCheque = 0
ValCartaoAmex = 0
ValCartaoBNDES = 0
ValCartaoMastercard = 0
ValCartaoVisa = 0
ValTEFVisaElectron = 0
valValoraPagar = 0
ValTEFRedeShop = 0
ValTEFHiperCard = 0
TotPago = 0
Modalidade = 0
'wTEFRedeShop = 0
'wTEFHiperCard = 0
ValNotaCredito = 0
frmFormaPagamento.chbValorPago.Caption = 0
frmFormaPagamento.chbValorPago.Caption = Format(frmFormaPagamento.chbValorPago.Caption, "##,###0.00")
frmFormaPagamento.chbValoraPagar.Caption = Format(frmFormaPagamento.chbValorPago.Caption, "##,###0.00")
frmFormaPagamento.chbValorFalta.Caption = Format(frmFormaPagamento.chbValoraPagar.Caption, "##,###0.00")
frmFormaPagamento.txtValorModalidade.text = ""


 wCodigoModalidadeDINHEIRO = ""
 WCodigoModalidadeAMEX = ""
 WCodigoModalidadeCHEQUE = ""
 wCodigoModalidadeBNDES = ""
 wCodigoModalidadeMASTERCARD = ""
 wCodigoModalidadeNOTACREDITO = ""
 wCodigoModalidadeFINANCIADO = ""
 wCodigoModalidadeFATURADO = ""
 wTEFVisaElectron = ""
 wTEFRedeShop = ""
 wTEFHiperCard = ""
 WCodigoModalidadeVISA = ""
End Sub

Public Sub verificaGarantiaEstendida(ByRef pedido As String)
 '  Private Sub pedidoComGarantia(NumeroPedido As String)
    Screen.MousePointer = 11
    Sql = "select count(*) garantiaEstendida " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where numeroPed = " & pedido & " and garantiaEstendida = 'S' and tipoNota = 'V' "
    
    rsProdutoGarantiaEstendida.CursorLocation = adUseClient
    rsProdutoGarantiaEstendida.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    'Set rsProdutoGarantiaEstendida = rdoCNLoja.OpenResultset(SQL)
        Screen.MousePointer = 0
        If rsProdutoGarantiaEstendida("garantiaEstendida") > 0 Then
            frmBilheteGarantia.Show vbModal
        End If
    rsProdutoGarantiaEstendida.Close
    Screen.MousePointer = 0
'End Sub
End Sub

Public Function ReplaceString(Texto As String, caracter As String, caracterParaSubstituir As String) As String
    
    Do While Texto Like "*" & caracter & "*"
        Texto = left$(Texto, (InStr(Texto, caracter) - 1)) _
        & caracterParaSubstituir _
        & right$(Texto, (Len(Texto) - (InStr(Texto, caracter))))
    Loop
    
    ReplaceString = Texto
    
End Function

Function VerificaSeEmiteCodigoZero() As String

       Sql = ""
       Sql = "Select CTS_EmiteCodigoZero from ControleSistema"
     
       rdoSerie.CursorLocation = adUseClient
       rdoSerie.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rdoSerie.EOF Then
           VerificaSeEmiteCodigoZero = rdoSerie("CTS_EmiteCodigoZero")
       End If
       rdoSerie.Close

End Function


Public Function EmiteNotafiscalTransferencia(ByVal nota As Double, ByVal Serie As String)
Dim wControlaQuebraDaPagina As Integer
wControlaQuebraDaPagina = 0

    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(Glb_ImpNotaFiscal) Then
           ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
      
    wSerie = Serie
    wNotaTransferencia = False
    wPagina = 0

    Call DadosLoja
            
    Sql = "select qtditem from nfcapa Where NumeroPed = " & frmControlaCaixa.txtPedido.text
    rsComplementoVenda.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            
    Sql = ""
   
    Sql = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV," _
        & "NFCAPA.Nf,NFCAPA.BASEICMS, NFCAPA.Serie, NFCAPA.PAGINANF, " _
        & "NFCAPA.volume,NFCAPA.PESOBR, NFCAPA.PESOLQ,  " _
        & "NFCAPA.CLIENTE,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.VLRMERCADORIA,Nfcapa.nf,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.lojaOrigem,NFCapa.PgEntra," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,LOJA.*,NFCAPA.CONDPAG, " _
        & "NfCapa.DataPag,NFCAPA.TOTALNOTAALTERNATIVA,NFCAPA.VALORTOTALCODIGOZERO," _
        & "NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NfItens.TipoNota,NfCapa.EmiteDataSaida " _
        & "From NFCAPA,NFITENS,LOJA " _
        & "Where NfCapa.nf= " & nota & " and NfCapa.Serie in ('" & Serie & "') " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "' " _
        & "and NfItens.LojaOrigem=NfCapa.LojaOrigem " _
        & "and NfItens.Serie=NfCapa.Serie " _
        & "and NfItens.Nf=NfCapa.NF " _
        & "and ltrim(rtrim(convert(char(5),NFCAPA.lojat))) = LOJA.LO_LOJA"


    RsDados.CursorLocation = adUseClient
    RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If Not RsDados.EOF Then
      Cabecalhotransferencia "T"
      
      Sql = "Select produtoloja.pr_referencia,produtoloja.pr_descricao, " _
          & "produtoloja.pr_classefiscal,produtoloja.pr_unidade,produtoloja.pr_st, " _
          & "produtoloja.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota," _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao,nfitens.CSTICMS," _
          & "nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produtoloja,nfitens " _
          & "where produtoloja.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & nota & " and Serie='" & Serie & "' order by nfitens.item"
     
      rsItensVenda.CursorLocation = adUseClient
      rsItensVenda.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

      If Not rsItensVenda.EOF Then
         wConta = 0
         wContItem = 0
         Printer.Print ""
         Do While Not rsItensVenda.EOF
            wContItem = wContItem + 1
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = rsItensVenda("PR_Referencia")
           
                     
           wPegaDescricaoAlternativa = IIf(IsNull(rsItensVenda("DescricaoAlternativa")), rsItensVenda("PR_Descricao"), rsItensVenda("DescricaoAlternativa"))
           If wPegaDescricaoAlternativa = "" Then
               wPegaDescricaoAlternativa = "0"
           End If
           If wPegaDescricaoAlternativa <> "0" Then
               wDescricao = wPegaDescricaoAlternativa
           Else
               wDescricao = Trim(rsItensVenda("pr_descricao"))
           End If
                    
                    
                    wStr16 = ""
                    wStr16 = left$(rsItensVenda("pr_referencia") & Space(7), 7) _
                          & Space(2) & left$(Format(Trim(wDescricao), ">") & Space(55), 55) _
                          & left$(Format(Trim(rsItensVenda("pr_classefiscal")), ">") _
                          & Space(11), 11) & left$(Trim("0" + Format(rsItensVenda("CSTICMS"), "00")) & Space(5), 5) _
                          & left$(Trim(rsItensVenda("pr_unidade")) & Space(2), 2) _
                          & right$(Space(8) & Format(rsItensVenda("QTDE"), "##0"), 8) _
                          & right$(Space(13) & Format(rsItensVenda("vlunit"), "#####0.00"), 13) _
                          & right$(Space(13) & Format(rsItensVenda("VlTotItem"), "#####0.00"), 13) _
                          & right$(Space(4) & Format(rsItensVenda("pr_icmssaida"), "#0"), 4)
                                  

                   Printer.Print wStr16
                      
                      If rsItensVenda("DetalheImpressao") = "D" Then
                         wConta = wConta + 1

                      ElseIf rsItensVenda("DetalheImpressao") = "C" Then
                            
                        Do While wConta < 34
                            wConta = wConta + 1
                            Printer.Print ""
                        Loop
                         
                         wConta = 0

                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If

                         Cabecalhotransferencia rsItensVenda("TipoNota")
                         Printer.Print ""
                         
                         
                       If wContItem = rsComplementoVenda("QTDITEM") Then
                          Call ImprimeCarimboTransferencia
                          Exit Function
                       End If

                      ElseIf rsItensVenda("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         Call ImprimeCarimboTransferencia
                         Exit Function
                      Else
                         wConta = wConta + 1
                      End If
                       rsItensVenda.MoveNext
            Loop
         Else
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
        'Exit Function
    End If
rsItensVenda.Close
RsDados.Close
rsComplementoVenda.Close
Printer.EndDoc

End Function


Function Cabecalhotransferencia(ByVal tiponota As String)
        
    Dim wCgcCliente As String
    Dim impri As Long
    Dim Linha(15) As String
    Dim ContLinha As Integer
    Dim ContParcela As Integer
    
    impri = Printer.Orientation
    wPagina = wPagina + 1
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    Linha(1) = "          "
    Linha(2) = "          "
    Linha(3) = "          "
    Linha(4) = "          "
    Linha(5) = "          "
    Linha(6) = "          "
    Linha(7) = "          "
    Linha(8) = "          "
    Linha(9) = "          "
    Linha(10) = "          "
    Linha(11) = "          "
    Linha(12) = "          "
    Linha(13) = "          "
    Linha(14) = "          "
    Linha(15) = "          "
    ContLinha = 1
    
    wCondicao = "            "
    Wav = "          "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = ""
    wentrada = "        "
    
    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))

    WNatureza = "TRANSFERENCIA"

    If Trim(wLojaVenda) > 0 Then
        If Trim(wLojaVenda) <> Trim(RsDados("LojaOrigem")) Then
            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
        Else
            wStr6 = ""
        End If
    Else
        wStr6 = ""
    End If
    If Trim(RsDados("AV")) > 1 Then
        If Mid(wCondicao, 1, 9) = "Faturada " Then
            Wav = "AV            : " & Trim(RsDados("AV"))
        End If
    End If
    
     wCondicao = "            "

    
    Linha(ContLinha) = "Pedido " & RsDados("NUMEROPED") & "  Ven " & RsDados("VENDEDOR")
    ContLinha = ContLinha + 1
             
    Sql = "select mo_descricao,mc_valor,mo_grupo from movimentocaixa,modalidade " & _
          "where mc_grupo = mo_grupo and mc_documento = " & RsDados("nf") & " and mc_Serie ='" & RsDados("serie") & _
          "' and mc_loja = '" & Trim(RsDados("lojaorigem")) & "' and mc_grupo like '10%'"

    rdoModalidade.CursorLocation = adUseClient
    rdoModalidade.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
    If Not rdoModalidade.EOF Then
        Do While Not rdoModalidade.EOF
         
          If rdoModalidade("mo_grupo") = 10501 Then
            
               Sql = "Select cp_condicao,cp_intervaloParcelas,cp_parcelas from CondicaoPagamento " _
                    & "where  CP_Codigo =" & RsDados("CondPag")

               rdoConPag.CursorLocation = adUseClient
               rdoConPag.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
               wValorParcela = Format((RsDados("totalnota") - RsDados("pgentra")) / rdoConPag("cp_parcelas"), "###,##0.00")
               ContParcela = 1
               wMid = 1
               Linha(ContLinha) = "Faturada " & rdoConPag("cp_parcelas") & " Parc    " & wValorParcela
               ContLinha = ContLinha + 1
               
               Do While Len(rdoConPag("cp_intervaloParcelas")) > wMid
               
                 If rdoConPag("cp_Parcelas") = 1 Then
                     Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                     wMid = wMid + 3
                 ElseIf rdoConPag("cp_Parcelas") Mod 2 = 0 Then
                       Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                       wMid = wMid + 6
                 Else
                       If Len(rdoConPag("cp_intervaloParcelas")) - 3 > wMid Then
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy") _
                           + "     " + Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid + 3, 3), "dd/mm/yyyy")
                           wMid = wMid + 6
                       Else
                           Linha(ContLinha) = Format(Date + Mid(rdoConPag("cp_intervaloParcelas"), wMid, 3), "dd/mm/yyyy")
                           wMid = wMid + 3
                       End If
                 End If
                 ContLinha = ContLinha + 1
              Loop
              rdoConPag.Close
           Else
              Linha(ContLinha) = rdoModalidade("mo_descricao") & ":   " & Format(rdoModalidade("mc_valor"), "0.00")
              ContLinha = ContLinha + 1
           End If
           rdoModalidade.MoveNext
        Loop
    End If
rdoModalidade.Close

    If RsDados("Pgentra") <> 0 Then
       wentrada = Format(RsDados("Pgentra"), "#####0.00")
       Linha(ContLinha) = "Entrada : " & Format(wentrada, "0.00")
       ContLinha = ContLinha + 1
    End If
    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
       Linha(ContLinha) = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
       ContLinha = ContLinha + 1
    End If
   
    If wPagina = 1 Then
        wCGC = right(String(14, "0") & wCGC, 14)
        wCGC = Format(Mid(wCGC, 1, Len(wCGC) - 6), "###,###,###") & "/" & Mid(wCGC, Len(wCGC) - 5, Len(wCGC) - 10) & "-" & Mid(wCGC, 13, Len(wCGC))
        wCGC = right(String(18, "0") & wCGC, 18)
    End If
  '  wStr0 = Space(110) & wPagina & "/" & RsDados("PAGINANF")  'Inicio Impressão
    wStr0 = Space(110) & "1" & "/" & "1"  'Inicio Impressão

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
    wStr1 = (left$(Linha(1) & Space(27), 27)) & Space(10) & left(Format(Trim(UCase(Wendereco)), "<") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(25) & right(Format(RsDados("nf"), "######"), 7)
    Printer.Print UCase(wStr1)
    wStr2 = (left$(Linha(2) & Space(27), 27)) & Space(10) & left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & left$(Trim(westado), 2)
    Printer.Print UCase(wStr2)
    wStr3 = (left$(Linha(3) & Space(27), 27)) & Space(10) & "(" & wDDDLoja & ")" & left$(Trim(Format(WFone, "####-####")), 9) & "/(" & wDDDLoja & ")" & left$(Format(WFax, "####-####"), 9) & Space(5) & left$(Format((WCep), "00000-000"), 9)
    Printer.Print UCase(wStr3)
    If wSerie = "CT" Then
        wStr4 = (left$(Linha(4) & Space(27), 27))
    Else
        wStr4 = (left$(Linha(4) & Space(27), 27)) & Space(60) & left(Trim(Format(wCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
     wStr4 = (left$(Linha(5) & Space(27), 27))
    
     Printer.Print UCase(wStr4)
     wStr5 = (left$(Linha(6) & Space(32), 32)) & left(Trim(WNatureza) & Space(25), 25) & left$(RsDados("codOper"), 10) & Space(28) & left$(Trim(Format((WIest), "###,###,###,###")), 15)

    Printer.Print wStr5
    wStr5 = (left$(Linha(7) & Space(27), 27))
    Printer.Print wStr5

        wCgcCliente = right(String(14, "0") & Trim(RsDados("LO_cgc")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = right(String(18, "0") & Trim(wCgcCliente), 18)

    
    Printer.Print ""

    wStr6 = (left$(Linha(8) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(RsDados("lo_razao")), ">") & Space(45), 45) & left$(Trim(wCgcCliente) & Space(24), 24) & left$(Format(RsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
    Printer.Print UCase(wStr6)
    
    wStr6 = (left$(Linha(9) & Space(27), 27))
    Printer.Print UCase(wStr6)
    wStr7 = (left$(Linha(10) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("lo_endereco") & ", " & RsDados("lo_numero")), ">") & Space(42), 42) & left$(Format(Trim(RsDados("lo_bairro")), ">") & Space(18), 18) & right$(Space(12) & Format(RsDados("lo_cep"), "#####-###"), 12) '& Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)


    Printer.Print UCase(wStr7)
    wStr7 = (left$(Linha(11) & Space(27), 27))
    Printer.Print UCase(wStr7)

    wStr8 = (left$(Linha(12) & Space(27), 27)) & Space(5) & left$(Format(Trim(RsDados("lo_municipio")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(RsDados("lo_telefone"))) & Space(15), 15) & left$(Trim(RsDados("lo_UF")), 2) & Space(5) & left$(Trim(Format(RsDados("lo_inscricaoEstadual"), "###,###,###,###")), 15)
    Printer.Print UCase(wStr8)
    Printer.Print ""

    If rdoConPag.State = 1 Then
        rdoConPag.Close
    End If

End Function


Private Sub ImprimeCarimboTransferencia()
                 
                       Sql = ""
'                       SQL = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
'                             "CNF_Nf = " & rsNFE("nf") & " and CNF_Serie = '" & rsNFE("Serie") & "' and CNF_Loja = '" & rsNFE("Lojaorigem") & "'" & _
'                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       Sql = "Select CNF_Carimbo,CNF_DetalheImpressao,CNF_TipoCarimbo from CarimboNotaFiscal where " & _
                             "CNF_Numeroped = '" & Trim(frmControlaCaixa.txtPedido.text) & "'" & _
                             "order by cnf_tipocarimbo desc, cnf_sequencia asc"
                       
                       RsPegaItensEspeciais.CursorLocation = adUseClient
                       RsPegaItensEspeciais.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

                       If Not RsPegaItensEspeciais.EOF Then
                            Printer.Print ""
                            Do While Not RsPegaItensEspeciais.EOF
                                If Trim(RsPegaItensEspeciais("CNF_tipocarimbo")) = "Z" Then
                                 wStr16 = right$(Space(116) & Trim(RsPegaItensEspeciais("CNF_Carimbo")), 116)
                                Else
                                 wStr16 = Space(5) & left$(RsPegaItensEspeciais("CNF_Carimbo") & Space(116), 116)
                                End If
                                 Printer.Print wStr16

                                 If RsPegaItensEspeciais("CNF_DetalheImpressao") = "D" Then
                                     wConta = wConta + 1
'                                     RsPegaItensEspeciais.MoveNext
                                 ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "C" Then
                                     
                                     Do While wConta < 34
                                       wConta = wConta + 1
                                       Printer.Print ""
                                     Loop

                                     wConta = 0

                         
                                     wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                                     If wControlaQuebraDaPagina = 3 Then
                                        Printer.Print ""
                                        wControlaQuebraDaPagina = 0
                                     End If

                                     Cabecalho rsNFE("tiponota")
                                     Printer.Print ""
                                ElseIf RsPegaItensEspeciais("CNF_DetalheImpressao") = "T" Then
                                       wConta = wConta + 1
                                       Printer.Print ""
                                       Call FinalizaNotaTransferencia(frmControlaCaixa.txtPedido)
                                Else
                                       wConta = wConta + 1
                                End If
                                RsPegaItensEspeciais.MoveNext
                            Loop

                             RsPegaItensEspeciais.Close
'                             Call FinalizaNota(wPedido)
                             Exit Sub
                         Else
                             RsPegaItensEspeciais.Close
                             Printer.Print ""
                             Printer.Print ""
                             Call FinalizaNotaTransferencia(frmControlaCaixa.txtPedido)
                         End If

End Sub



Private Sub FinalizaNotaTransferencia(wPedido As String)
     If wNotaTransferencia = False Then
   
        Do While wConta < 13
        wConta = wConta + 1
        Printer.Print ""
        Loop
       
     End If


        wStr9 = right$(Space(9) & Format(RsDados("BaseICMS"), "######0.00"), 9) & right$(Space(25) & Format(RsDados("VLRICMS"), "######0.00"), 12) & Space(34) & right$(Space(10) & Format(RsDados("VlrMercadoria"), "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = right(Space(9) & Format(Space(9) & RsDados("FreteCobr"), "######0.00"), 9) & Space(46) & right(Space(10) & Format(RsDados("TotalNota"), "######0.00"), 10)
        Printer.Print wStr10

     
     wStr11 = Space(2) & "                          "
     Printer.Print wStr11
     wStr12 = Space(2) & "                                                     "
     Printer.Print wStr12
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = right$(Space(5) & Format(RsDados("Volume"), "######0.00"), 5) & Space(5) & "Volume(s)" & Space(25) & right$(Space(7) & Format(RsDados("PesoBR"), "######0.00"), 7) & Space(5) & right$(Space(7) & Format(RsDados("PesoBR"), "######0.00"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(105) & right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
'     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.EndDoc
     

End Sub

Public Sub verificaAppExecucao()
    If App.PrevInstance Then
       MsgBox App.EXEName + " Já está executando", vbCritical
       End
       
    End If
End Sub

Public Function carregaProdutoGarantia(pedido As String) As Boolean

    Dim rsProdGarantiaEstendida As New ADODB.Recordset
    
    aceitaGarantia = False
    
    Sql = "select count(*) itensGarantia " & _
    "from produtoLoja as p, nfitens as i, nfcapa as c " & _
    "where i.numeroPed = " & pedido & " and  " & _
    "p.pr_referencia = i.referencia and " & _
    "p.pr_garantiaEstendida = 'S' and i.numeroPed = c.numeroPed and " & _
    "c.vendedor not in (999,888,777) and c.garantiaEstendida = 'N' and " & _
    "c.condpag  < 3"
    
    rsProdGarantiaEstendida.CursorLocation = adUseClient
    rsProdGarantiaEstendida.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Val(rsProdGarantiaEstendida("itensGarantia")) > 0 Then
            frmGarantiaEstendida.wNumeroPedido = pedido
            'frmGarantiaEstendida.ZOrder 0
            'frmFormaPagamento.Visible = False
            frmGarantiaEstendida.Show 1
            'frmFormaPagamento.Visible = True
            
        End If
    rsProdGarantiaEstendida.Close
    
End Function


Public Sub ImprimeTransferencia00(ByVal nota As Double)
'Dim wNomeVendedor As String
'-==========================Emerson

    Dim ValorlItem As Double
    Dim ValorDesconto As Double
    Dim SubTotal As Double
    Dim TotalVenda As Double

    Dim nomeEmpresa As String * 48
    Dim cnpj As String * 48
    Dim Data As String * 48
    Dim Endereco As String * 48
    Dim Telefone As String * 48
    Dim pedido As String * 48
    'MsgBox GLB_Impressora00
    'Open GLB_Impressora00 For Output As #1
 
    Screen.MousePointer = 11
   
    ValorlItem = 0
    ValorDesconto = 0
    SubTotal = 0

    Sql = "Select Lojaorigem,cliente,totalnota,lojat from nfcapa Where numeroped = " & frmControlaCaixa.txtPedido.text
    
    RsDadosCapa.CursorLocation = adUseClient
    RsDadosCapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    Sql = "Select * from Loja Where LO_Loja= '" & Trim(RsDadosCapa("Lojaorigem")) & "'"

    RsDados.CursorLocation = adUseClient
    RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    impressoraRelatorio "[INICIO]"

    impressoraRelatorio "                                                "
    impressoraRelatorio "                                                "

   nomeEmpresa = RsDados("LO_Razao")
   impressoraRelatorio nomeEmpresa
   
   cnpj = "CNPJ: " & RsDados("LO_CGC") & " I.E.: " & RsDados("LO_InscricaoEstadual")
   impressoraRelatorio cnpj
   
   Endereco = UCase(RsDados("LO_Endereco")) & ", " & RsDados("LO_numero")
   impressoraRelatorio Endereco
   
   Telefone = "TELEFONE: " & RsDados("LO_Telefone")
   impressoraRelatorio Telefone
   
   Data = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(23) & NroNotaFiscal
   impressoraRelatorio Data
   
   impressoraRelatorio "                                                "
   
   impressoraRelatorio "================================================"

   'impressoraRelatorio  & Space(21)
   impressoraRelatorio centralizarTexto("TRANSFERENCIA PARA LOJA " & Trim(RsDadosCapa("lojat")), 48)
   impressoraRelatorio "________________________________________________"
   impressoraRelatorio "DESCRICAO DO PRODUTO                            "



   impressoraRelatorio "CODIGO  PRODUTO  QTDxUNIT.   VALOR TOTAL        "
   impressoraRelatorio "________________________________________________"

   RsDados.Close

   
   Sql = "Select * from Nfitens " _
       & "Where numeroped = " & frmControlaCaixa.txtPedido.text
       
      
       RsDados.CursorLocation = adUseClient
       RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       
       If Not RsDados.EOF Then
          Do While Not RsDados.EOF
             Sql = "Select PR_Descricao from Produtoloja Where PR_Referencia ='" & RsDados("Referencia") & "'"
             rdoProduto.CursorLocation = adUseClient
             rdoProduto.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             
             ValorlItem = (RsDados("vlunit") * RsDados("Qtde"))
             SubTotal = (SubTotal + ValorlItem)
           
'             impressoraRelatorio right(Space(10) & rdoProduto("PR_Descricao"), 30) & Space(18)

             impressoraRelatorio left(Trim(rdoProduto("PR_Descricao")) & Space(32), 48)
             impressoraRelatorio Trim(RsDados("referencia")) _
             & Space(3) & right(Space(4) & Format(RsDados("Qtde"), "###0"), 4) & "x" _
             & left(Format(RsDados("vlunit"), "###,###,###.00") & Space(20), 11) _
             & left(right(Space(30) & Format(ValorlItem, "###,###,###.00"), 10) & Space(30), 22)
             impressoraRelatorio "                                                "
             rdoProduto.Close
             RsDados.MoveNext
          Loop
       End If
       
       RsDados.Close
       
         
'       totalvenda = (SubTotal - ValorDesconto)
     
       impressoraRelatorio "                                                "

'       impressoraRelatorio   "SUB TOTAL " & Space(16) & Right(Space(10) & Format(RsDadosCapa("vlrMercadoria"), "###,###,##0.00"), 14)
'       impressoraRelatorio   ""

'       impressoraRelatorio   "DESCONTO  " & Space(16) & Right(Space(10) & Format(RsDadosCapa("desconto"), "###,###,##0.00"), 14)
'       impressoraRelatorio   " "

'       impressoraRelatorio   "FRETE     " & Space(16) & Right(Space(10) & Format(RsDadosCapa("fretecobr"), "###,###,##0.00"), 14)
'       impressoraRelatorio   " "

       impressoraRelatorio "TOTAL     " & Space(20) & left(Format(RsDadosCapa("totalnota"), "###,###,##0.00") & Space(15), 18)
       impressoraRelatorio "________________________________________________"

    
       pedido = "Pedido: " & Trim(frmControlaCaixa.txtPedido.text)
       impressoraRelatorio pedido
       impressoraRelatorio "================================================"
       impressoraRelatorio "                                                "
       impressoraRelatorio "                                                "
       impressoraRelatorio "                                                "


     RsDadosCapa.Close
      impressoraRelatorio "[FIM]"
   


      Screen.MousePointer = 0

End Sub


Public Function CriaMovimentoCaixa(ByVal Nf As Double, ByVal Serie As String, ByVal TotalNota As Double, ByVal loja As String, ByVal Grupo As Double, ByVal NroProtocolo As Integer, ByVal nroCaixa As Integer, ByVal NroPedido As Double)
    
    Sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo, MC_NroCaixa, MC_DataProcesso, MC_Pedido) values(" & GLB_ECF & ",'0','" & Trim(loja) & "', " _
        & " '" & Format(Date, "yyyy/mm/dd") & "'," & Grupo & ", " & Nf & ",'" & Serie & "', " _
        & "" & ConverteVirgula(Format(TotalNota, "##,###0.00")) & ", " _
        & "0,0,0,0,0,9,'A'," & NroProtocolo & "," & nroCaixa & ",'" & Format(Date, "yyyy/mm/dd") & "'," & NroPedido & ")"
        adoCNLoja.Execute (Sql)

End Function

Public Function campoCaixa(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoPadrao(KeyAscii) _
    Or KeyAscii = 77 Or KeyAscii = 68 _
    Or KeyAscii = 66 Or KeyAscii = 98 _
    Or KeyAscii = 109 Or KeyAscii = 100 Then
        campoCaixa = KeyAscii
    End If
End Function

Public Function campoNumerico(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNumerico = KeyAscii
    End If
End Function

Public Function digitoNumerico(KeyAscii As Integer) As Boolean
    digitoNumerico = False
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        digitoNumerico = True
    End If
End Function

Public Function digitoNumericoComVirgula(KeyAscii As Integer) As Boolean
    digitoNumericoComVirgula = False
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 44 Then
        digitoNumericoComVirgula = True
    End If
End Function

Public Function digitoVirgulaPonto(KeyAscii As Integer) As Boolean
    digitoVirgulaPonto = False
    If KeyAscii = 44 Or KeyAscii = 46 Then
        digitoVirgulaPonto = True
    End If
End Function

Public Function digitoPadrao(KeyAscii As Integer) As Boolean
    digitoPadrao = False
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 27 Then
        digitoPadrao = True
    End If
End Function

Public Sub CarregaMovimento(grid, protocolo As String)

     Dim wSubTotal As Double
     Dim wSubTotal_S As Double
     Dim wTotalNf As Double
     Dim wTotalFatFin As Double
     Dim wTotalReforco As Double

  wTNNotaCredito = 0
  
  'lblCabec.Caption = lblCabec
  grid.Rows = 1
  grid.Rows = 1
  grid.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '1
  grid.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                      '2
  grid.AddItem "Cartões >>"                                                                            '3
  grid.AddItem "  Visa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50301"                      '4
  grid.AddItem "  MasterCard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50302"                '5
  grid.AddItem "  Amex" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50303"                      '6
  grid.AddItem "  BNDES" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50304"                     '7
  grid.AddItem "  Rede Shop" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50203"                 '8
  grid.AddItem "  Visa Elec." & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50203"                '9
  grid.AddItem "Nota de Credito" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50701"             '10
  grid.AddItem "Hypercard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50205"                   '11
  grid.AddItem ""                                                                                      '12
  grid.AddItem "*** TOTAL CAIXA" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70204"             '13
  grid.AddItem ""                                                                                      '14
  grid.AddItem "Faturado" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                       '15
  grid.AddItem "Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                     '16
  grid.AddItem "Reforco de Caixa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50801"            '17
  grid.AddItem ""                                                                                      '18
  grid.AddItem "*** TOTAL GERAL" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                '19
  grid.AddItem ""                                                                                      '20
  grid.AddItem "Entrada Faturada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50502"            '21
  grid.AddItem "Entrada Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50601"          '22
  grid.AddItem "Garantia Estendida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                             '23
  grid.AddItem ""                                                                                      '24
  grid.AddItem "*** Movimento NF"                                                                      '25
  grid.AddItem ""                                                                                      '26
  grid.AddItem "CE" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '27
  grid.AddItem "NE" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '28
'  grid.AddItem "D1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '29 ''''''''''
'  grid.AddItem "S1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '30 ''''''''''
  grid.AddItem ""                                                                                      '31 '29
  grid.AddItem "*** TOTAL NF" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '32 '30
  grid.AddItem ""                                                                                      '33 '31
  grid.AddItem "Transferencia Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                            '34 '32
  grid.AddItem "Remessa Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                  '35 '33
  grid.AddItem "Devolucao Entrada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                              '36 '34
  grid.AddItem ""                                                                                      '37 '35
  grid.AddItem "CE Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '38 '36
  grid.AddItem "NE Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '39 '37
'  grid.AddItem "D1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                  '40 ''''''''''''''
'  grid.AddItem "S1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                  '41 ''''''''''''''
  grid.AddItem ""                                                                                      '42 '38
  grid.AddItem "** Saldo Anterior**"                                                                   '43 '39
  grid.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '44 '40
  grid.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                      '45 '41
  grid.AddItem "Total do Saldo" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "00000"              '46 '42
  grid.AddItem ""
  wTotalSaldo = 0
  wTotalSaldo_S = 0

 wTotalNf = 0
 wTotalFatFin = 0
 grid.Row = 1

 Sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_Protocolo in (" & protocolo _
       & ") and  MC_Serie <> '00' and (MC_Grupo like '10%' or MC_Grupo like '11%'" _
       & " or MC_Grupo like '50%' or MC_Grupo like '20%') AND MC_TipoNota in ('V','T','E','S') group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
 If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "10101" Then
           grid.TextMatrix(1, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10201" Then
           grid.TextMatrix(2, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "10301" Then
           grid.TextMatrix(4, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10302" Then
           grid.TextMatrix(5, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10303" Then
           grid.TextMatrix(6, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10304" Then
           grid.TextMatrix(7, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10203" Then
           grid.TextMatrix(8, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10206" Then
           grid.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10205" Then
           grid.TextMatrix(11, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10701" Then
           grid.TextMatrix(10, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           'wTNNotaCredito = wTNNotaCredito + rdoFormaPagamento("TotalModalidade")
        'ElseIf rdoFormaPagamento("MC_Grupo") = "10204" Then 'AVR
           'grid.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11004" Then
           grid.TextMatrix(21, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10601" Then
           grid.TextMatrix(16, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10801" Then
           grid.TextMatrix(17, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "11008" Then 'AVR
           'grid.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11006" Then 'SALDO ANTERIOR
           grid.TextMatrix(40, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior = (wtotalSaldoAnterior + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11007" Then 'SALDO ANTERIOR
           grid.TextMatrix(41, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior = (wtotalSaldoAnterior + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10501" Then
           grid.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11005" Then
           grid.TextMatrix(22, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11009" Then
           grid.TextMatrix(23, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wtotalGarantia = (wtotalGarantia + rdoFormaPagamento("TotalModalidade"))
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           'wSubTotal_S = (wSubTotal_S - rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20101" Then
           grid.TextMatrix(27, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(27, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalReforco = rdoFormaPagamento("TotalModalidade")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
           grid.TextMatrix(28, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(28, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
'           grid.TextMatrix(29, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grid.TextMatrix(29, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
'           grid.TextMatrix(30, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grid.TextMatrix(30, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "20111" Then
        '   grid.TextMatrix(33, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
        '   grid.TextMatrix(33, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        '   wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20109" Then
           grid.TextMatrix(32, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(32, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20110" Then
           grid.TextMatrix(33, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(33, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20201" Then
           grid.TextMatrix(34, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(34, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        
        ElseIf rdoFormaPagamento("MC_Grupo") = "50101" Then
           grid.TextMatrix(1, 2) = CDbl(Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")) + CDbl(grid.TextMatrix(1, 2))
           'grid.TextMatrix(1, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50201" Then
           grid.TextMatrix(2, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "50301" Then
           grid.TextMatrix(4, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50302" Then
           grid.TextMatrix(5, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50303" Then
           grid.TextMatrix(6, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50304" Then
           grid.TextMatrix(7, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50203" Then
           grid.TextMatrix(8, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50206" Then
           grid.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50205" Then
           grid.TextMatrix(11, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50701" Then
           grid.TextMatrix(10, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "50204" Then 'AVR
           'grid.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50502" Then
           grid.TextMatrix(21, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50602" Then
           grid.TextMatrix(22, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50009" Then
           grid.TextMatrix(1, 2) = CDbl(Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")) + CDbl(grid.TextMatrix(1, 2))
           grid.TextMatrix(23, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
           
        ElseIf rdoFormaPagamento("MC_Grupo") = "50801" Then
           grid.TextMatrix(17, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade")) ''RETIRADO DIA 02/11
           
        'ElseIf rdoFormaPagamento("MC_Grupo") = "50804" Then
           'grid.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50806" Then
           grid.TextMatrix(40, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior_S = (wtotalSaldoAnterior_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50807" Then
           grid.TextMatrix(41, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior_S = (wtotalSaldoAnterior_S + rdoFormaPagamento("TotalModalidade"))
        End If
       rdoFormaPagamento.MoveNext
     Loop
     
     grid.TextMatrix(1, 1) = Format((grid.TextMatrix(1, 1)), "###,###,###,##0.00")
     grid.TextMatrix(1, 2) = Format((grid.TextMatrix(1, 2)), "###,###,###,##0.00")
     
     grid.TextMatrix(1, 3) = Format((grid.TextMatrix(1, 1) - grid.TextMatrix(1, 2)), "###,###,###,##0.00")
     grid.TextMatrix(2, 3) = Format((grid.TextMatrix(2, 1) - grid.TextMatrix(2, 2)), "###,###,###,##0.00")
     grid.TextMatrix(4, 3) = Format((grid.TextMatrix(4, 1) - grid.TextMatrix(4, 2)), "###,###,###,##0.00")
     grid.TextMatrix(5, 3) = Format((grid.TextMatrix(5, 1) - grid.TextMatrix(5, 2)), "###,###,###,##0.00")
     grid.TextMatrix(6, 3) = Format((grid.TextMatrix(6, 1) - grid.TextMatrix(6, 2)), "###,###,###,##0.00")
     grid.TextMatrix(7, 3) = Format((grid.TextMatrix(7, 1) - grid.TextMatrix(7, 2)), "###,###,###,##0.00")
     grid.TextMatrix(8, 3) = Format((grid.TextMatrix(8, 1) - grid.TextMatrix(8, 2)), "###,###,###,##0.00")
     grid.TextMatrix(9, 3) = Format((grid.TextMatrix(9, 1) - grid.TextMatrix(9, 2)), "###,###,###,##0.00")
     grid.TextMatrix(10, 3) = Format((grid.TextMatrix(10, 1) - grid.TextMatrix(10, 2)), "###,###,###,##0.00")
     grid.TextMatrix(11, 3) = Format((grid.TextMatrix(11, 1) - grid.TextMatrix(11, 2)), "###,###,###,##0.00")
     'grid.TextMatrix(9, 3) = Format((grid.TextMatrix(9, 1) - grid.TextMatrix(9, 2)), "###,###,###,##0.00")
     grid.TextMatrix(21, 3) = Format((grid.TextMatrix(21, 1) - grid.TextMatrix(21, 2)), "###,###,###,##0.00")
     grid.TextMatrix(22, 3) = Format((grid.TextMatrix(22, 1) - grid.TextMatrix(22, 2)), "###,###,###,##0.00")
     grid.TextMatrix(23, 3) = Format((grid.TextMatrix(23, 1) - grid.TextMatrix(23, 2)), "###,###,###,##0.00")
     grid.TextMatrix(17, 3) = Format((grid.TextMatrix(17, 1) - grid.TextMatrix(17, 2)), "###,###,###,##0.00")
     grid.TextMatrix(40, 3) = Format((grid.TextMatrix(40, 1) - grid.TextMatrix(40, 2)), "###,###,###,##0.00")
     grid.TextMatrix(41, 3) = Format((grid.TextMatrix(41, 1) - grid.TextMatrix(41, 2)), "###,###,###,##0.00")
     
     'grid.TextMatrix(18, 1) = Format(wTotalSaldo, "###,###,###,##0.00")
     'grid.TextMatrix(18, 2) = Format(wTotalSaldo_S, "###,###,###,##0.00")
     'grid.TextMatrix(18, 3) = Format((wTotalSaldo - wTotalSaldo_S), "###,###,###,##0.00")


    'wSubTotal = grid.TextMatrix(11, 1)
     grid.TextMatrix(42, 1) = Format(wTotalSaldo, "###,###,###,##0.00")
     grid.TextMatrix(42, 2) = Format(wTotalSaldo_S, "###,###,###,##0.00")
     grid.TextMatrix(42, 3) = Format((wTotalSaldo - wTotalSaldo_S), "###,###,###,##0.00")

     grid.TextMatrix(13, 1) = Format((wSubTotal), "###,###,###,##0.00")
     grid.TextMatrix(13, 2) = Format((wSubTotal_S), "###,###,###,##0.00")
     grid.TextMatrix(13, 3) = Format(((wSubTotal) - (wSubTotal_S)), "###,###,###,##0.00")
     grid.TextMatrix(30, 1) = Format(wTotalNf, "###,###,###,##0.00")
     grid.TextMatrix(19, 1) = Format((wTotalFatFin + wSubTotal) - wtotalGarantia, "###,###,###,##0.00")
     
  End If
  rdoFormaPagamento.Close
  
 Sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_Protocolo in (" & protocolo _
       & ") and  MC_Serie <> '00' and MC_tiponota = 'C' and MC_Grupo like '20%' group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
  If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "20101" Then
           grid.TextMatrix(36, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(36, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
           grid.TextMatrix(37, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grid.TextMatrix(37, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
'           grid.TextMatrix(40, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00") '?
'           grid.TextMatrix(40, 2) = Format(rdoFormaPagamento("Quantidade"), "0") '?
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
'           grid.TextMatrix(41, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00") '?
'           grid.TextMatrix(41, 2) = Format(rdoFormaPagamento("Quantidade"), "0") '?
        'ElseIf rdoFormaPagamento("MC_Grupo") = "20111" Then
        '   grid.TextMatrix(44, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
        '   grid.TextMatrix(44, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        End If
        rdoFormaPagamento.MoveNext
     Loop
  End If
  
  rdoFormaPagamento.Close
  wSubTotal = 0
  wSubTotal_S = 0
  wtotalGarantia = 0

  Call CarregaMovimentoZERO(grid, protocolo)

End Sub

Private Sub CarregaMovimentoZERO(grid, protocolo As String)
    
    Dim Sql As String

    Sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
          & " Where MC_Protocolo in (" & protocolo _
          & ") and  MC_Serie = '00' and (MC_Grupo like '20105') AND MC_TipoNota in ('V','T','E','S') group by mc_grupo")
          
    rdoFormaPagamento.CursorLocation = adUseClient
    rdoFormaPagamento.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rdoFormaPagamento.EOF Then
    
        grid.TextMatrix(29, 0) = "00 *"
        grid.TextMatrix(29, 1) = rdoFormaPagamento("TotalModalidade")
        grid.TextMatrix(29, 2) = rdoFormaPagamento("Quantidade")
        grid.TextMatrix(29, 3) = "0"
    
    End If
    
    rdoFormaPagamento.Close
    
     Sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
          & " Where MC_Protocolo in (" & protocolo _
          & ") and  MC_Serie = '00' and (MC_Grupo like '20105') AND MC_TipoNota in ('C') group by mc_grupo")
          
    rdoFormaPagamento.CursorLocation = adUseClient
    rdoFormaPagamento.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rdoFormaPagamento.EOF Then
    
        grid.TextMatrix(38, 0) = "00 Cancelada *"
        grid.TextMatrix(38, 1) = rdoFormaPagamento("TotalModalidade")
        grid.TextMatrix(38, 2) = rdoFormaPagamento("Quantidade")
        grid.TextMatrix(38, 3) = "0"
    
    End If
    
    rdoFormaPagamento.Close
    
End Sub

Public Function carregaControleCaixa() As Boolean

    Dim Sql As String
    carregaControleCaixa = False

    Sql = "Select ControleCaixa.*,USU_Codigo,USU_Nome from ControleCaixa,UsuarioCaixa" _
     & " Where CTR_Supervisor <> 99 and CTR_Operador = USU_Codigo and CTR_SituacaoCaixa='A' and CTR_NumeroCaixa = " & GLB_Caixa
    
     RsDados.CursorLocation = adUseClient
     RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     If RsDados.EOF = False Then
        GLB_USU_Nome = RsDados("USU_Nome")
        GLB_USU_Codigo = RsDados("USU_Codigo")
        GLB_CTR_Protocolo = RsDados("CTR_Protocolo")
        GLB_DataInicial = Format(RsDados("CTR_DataInicial"), "YYYY/MM/DD")
        GLB_HoraInicial = Format(RsDados("CTR_DataInicial"), "HH:MM:SS")
        carregaControleCaixa = True
     End If
     
     If RsDados.EOF = False Then
       If RsDados("CTR_Situacaocaixa") = "A" And Format(RsDados("ctr_datainicial"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
          MsgBox "Data do caixa incorreta.Favor efetuar o Fechamento", vbCritical, "Atenção"
          wPermitirVenda = False
     Else
          wPermitirVenda = True
      End If
     End If
     
     RsDados.Close
     
End Function




 Public Sub NOVO_ImprimeMovimento(grid, Titulo As String, operador As String, _
                                  nroCaixa As String, dataInicio As String, _
                                  horaInicio As String, dataFim As String, _
                                  horaFim As String, protocolo As String)

    Dim wTotalTransferencia As Double
    Dim wTotalSaldo As Double
    Dim wSaldoAnterior As Double
    Dim wMovimentoPeriodo As Double
    Dim wSaldoFinalDinheiro As Double
    Dim wSaldoFinalCheque As Double
     
     Dim tamanhoPadraoLinhas As Integer
     
     tamanhoPadraoLinhas = 48
    
     wSaldoFinalDinheiro = 0
     wSaldoFinalCheque = 0
     wSaldoFinalAVR = 0

    Screen.MousePointer = 11
    impressoraRelatorio "[INICIO]"
     
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio centralizarTexto(Titulo, tamanhoPadraoLinhas)
    impressoraRelatorio " DATA " & left(dataInicio, 10) & "                        Loja " & Format(GLB_Loja, "000")
    impressoraRelatorio "________________________________________________"
 
     For Idx = 1 To grid.Rows - 6 Step 1
         If (Idx = 25) Or (Idx = 26) Or (Idx = 27) Or (Idx = 28) Or (Idx = 30) Or (Idx = 31) Or (Idx = 32) Or _
         (Idx = 35) Or (Idx = 36) Or (Idx = 37) Then
                          
         impressoraRelatorio left(grid.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grid.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & _
               right(Space(5) & grid.TextMatrix(Idx, 2), 5)
         Else
         
          impressoraRelatorio left(grid.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grid.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & "     "
         End If
     Next Idx
     
     wMovimentoPeriodo = Format(grid.TextMatrix(13, 1), "###,###,##0.00")
     wSaldoAnterior = Format(grid.TextMatrix(41, 1), "###,###,##0.00")
     wSaldoFinalDinheiro = Format(grid.TextMatrix(1, 3), "###,###,##0.00")
     wSaldoFinalDinheiro = (wSaldoFinalDinheiro + Format(grid.TextMatrix(40, 3), "###,###,##0.00"))
     wSaldoFinalCheque = Format(grid.TextMatrix(2, 3))
     wSaldoFinalCheque = (wSaldoFinalCheque + Format(grid.TextMatrix(41, 3), "###,###,##0.00"))
     wTotalTransferencia = CDbl(grid.TextMatrix(13, 2)) + CDbl(grid.TextMatrix(42, 2))
     wTotalTransferencia = wTotalTransferencia + CDbl(grid.TextMatrix(17, 2))
     wTotalTransferencia = wTotalTransferencia - wtotalGarantia
    
     
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio "               MOVIMENTO DE CAIXA               "
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio "SALDO ANTERIOR >> " & right(Space(30) & Format("", ""), 30)
    impressoraRelatorio "  DINHEIRO        " & right(Space(30) & Format(grid.TextMatrix(40, 1), "###,###,##0.00"), 30)
    impressoraRelatorio "  CHEQUE          " & right(Space(30) & Format(grid.TextMatrix(41, 1), "###,###,##0.00"), 30)
    impressoraRelatorio "  TOTAL           " & right(Space(30) & Format(wTotalSaldo, "###,###,##0.00"), 30)
    impressoraRelatorio "MOVIMENTO PERIODO " & right(Space(30) & Format(wMovimentoPeriodo, "###,###,##0.00"), 30)
    impressoraRelatorio "REFORCO           " & right(Space(30) & Format(grid.TextMatrix(17, 1), "###,###,##0.00"), 30)
    impressoraRelatorio "TRANSFERENCIA NUM." & right(Space(30) & Format(wTotalTransferencia, "###,###,##0.00"), 30)
    impressoraRelatorio "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30)
    
     
    impressoraRelatorio "SALDO DO CAIXA >>                               "
    impressoraRelatorio "  Dinheiro        " & right(Space(30) & Format(wSaldoFinalDinheiro, "###,###,##0.00"), 30)
    impressoraRelatorio "  Cheque          " & right(Space(30) & Format(wSaldoFinalCheque, "###,###,##0.00"), 30)
    impressoraRelatorio "  SALDO FINAL     " & right(Space(30) & Format((wSaldoFinalDinheiro + wSaldoFinalCheque), "###,###,##0.00"), 30)
    
     
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio left("Caixa Nro.   " & nroCaixa & Space(48), 48)
    impressoraRelatorio left("Operador     " & operador & Space(48), 48)
    impressoraRelatorio left("Data Inicial " & left(dataInicio, 10) & " " & left(horaInicio, 5) & Space(48), 48)
    impressoraRelatorio left("Data Final   " & left(dataFim, 10) & " " & left(horaFim, 5) & Space(48), 48)
    impressoraRelatorio left("Protocolo    " & protocolo & Space(48), 48)
    impressoraRelatorio left("Versão Caixa " & App.Major & "." & App.Minor & "." & App.Revision & Space(48), 48)
               
    imprimeCampoGerenteOperador
    
    impressoraRelatorio "[FIM]"

    
    Screen.MousePointer = 0
    
End Sub

Public Sub imprimeCampoGerenteOperador()
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio ".                                               "
    impressoraRelatorio " ______________________   ______________________"
    impressoraRelatorio "        OPERADOR                SUPERVISOR      "
    impressoraRelatorio "                                                "
    impressoraRelatorio "                                                "
End Sub


Public Sub NOVO_ImprimeTransfNumerario(grid, Titulo As String, operador As String, _
                                       nroCaixa As String, dataInicio As String, _
                                       horaInicio As String, dataFim As String, _
                                       horaFim As String, protocolo As String)
     
     Dim tamanhoPadraoLinhas As Integer
     tamanhoPadraoLinhas = 48
     
    Screen.MousePointer = 11
    impressoraRelatorio "[INICIO]"
    
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio centralizarTexto(Titulo, tamanhoPadraoLinhas)
    impressoraRelatorio " DATA " & left(dataInicio, 10) & "                        Loja " & Format(GLB_Loja, "000")
    impressoraRelatorio "________________________________________________"
    
    
    impressoraRelatorio "DINHEIRO /P TESOU." & right(Space(30) & Format(wTNDinheiro, "###,###,##0.00"), 30)
    impressoraRelatorio "CHEQUE /P TESOU.  " & right(Space(30) & Format(wTNCheque, "###,###,##0.00"), 30)
    impressoraRelatorio "VISA              " & right(Space(30) & Format(wTNVisa, "###,###,##0.00"), 30)
    impressoraRelatorio "MASTERCARD        " & right(Space(30) & Format(wTNRedecard, "###,###,##0.00"), 30)
    impressoraRelatorio "AMEX              " & right(Space(30) & Format(wTNAmex, "###,###,##0.00"), 30)
    impressoraRelatorio "BNDS              " & right(Space(30) & Format(wTNBNDES, "###,###,##0.00"), 30)
    impressoraRelatorio "REDE SHOP         " & right(Space(30) & Format(wTNRedeShop, "###,###,##0.00"), 30)
    impressoraRelatorio "VISA ELEC.        " & right(Space(30) & Format(wTNVisaEletron, "###,###,##0.00"), 30)
    
    impressoraRelatorio "HIPERCARD         " & right(Space(30) & Format(wTNHiperCard, "###,###,##0.00"), 30)
    impressoraRelatorio "DEPOSITO          " & right(Space(30) & Format(wTNDeposito, "###,###,##0.00"), 30)
    impressoraRelatorio "NOTA CREDITO      " & right(Space(30) & Format(wTNNotaCredito, "###,###,##0.00"), 30)
    impressoraRelatorio "OUTRAS DESPESAS   " & right(Space(30) & Format(wTNConducao, "###,###,##0.00"), 30)
    impressoraRelatorio "DESPESA LOJA      " & right(Space(30) & Format(wTNDespLoja, "###,###,##0.00"), 30)
    impressoraRelatorio "                                                "
    impressoraRelatorio "TOTAL             " & right(Space(30) & Format(wTNTotal, "###,###,##0.00"), 30)
    
    impressoraRelatorio "                  " & right(Space(30) & "", 30)
    impressoraRelatorio "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30)
    impressoraRelatorio "ENTRADA FINANCIADA" & right(Space(30) & Format(wTNFinanciado, "###,###,##0.00"), 30)
    impressoraRelatorio "ENTRADA FATURADA  " & right(Space(30) & Format(wTNFaturado, "###,###,##0.00"), 30)
    
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio left("Caixa Nro.   " & nroCaixa & Space(48), 48)
    impressoraRelatorio left("Operador     " & operador & Space(48), 48)
    impressoraRelatorio left("Data Inicial " & left(dataInicio, 10) & " " & left(horaInicio, 5) & Space(48), 48)
    impressoraRelatorio left("Data Final   " & left(dataFim, 10) & " " & left(horaFim, 5) & Space(48), 48)
    impressoraRelatorio left("Protocolo    " & protocolo & Space(48), 48)
    impressoraRelatorio left("Versão Caixa " & App.Major & "." & App.Minor & "." & App.Revision & Space(48), 48)
    
    imprimeCampoGerenteOperador
    
    impressoraRelatorio "[FIM]"
    
    Screen.MousePointer = 0
     
End Sub

 Private Function centralizarTexto(text As String, tamanhoCampo As Integer)
    Dim inicioTexto As Integer
    inicioTexto = (tamanhoCampo / 2) - (Len(text) / 2)
    centralizarTexto = Space(inicioTexto) & text
    centralizarTexto = centralizarTexto & Space(tamanhoCampo - Len(centralizarTexto))
 End Function

Public Sub impressoraRelatorio(Texto As String)
    
    If GLB_Impressora00 = "CUPOM" _
    Or GLB_Impressora00 = "CF" _
    Or GLB_Impressora00 = "FISCAL" _
    Or GLB_Impressora00 = "CUPOM FISCAL" Then
    
        If Texto = "[INICIO]" Then
            Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
        ElseIf Texto = "[FIM]" Then
            Retorno = Bematech_FI_FechaRelatorioGerencial()
        Else
            Retorno = Bematech_FI_UsaRelatorioGerencialMFD(Texto)
        End If
    
       
    Else


        If Texto = "[INICIO]" Then
        
            'Printer.PaintPicture frmSangria.imgLogo.Picture, 1150, 0, 2000, 2000
            'For i = 0 To 15
                'Printer.Print " "
            'Next i
            'Printer.Print "---------------------"
            'Printer.EndDoc
            
        '            Debug.Print ""
        
        ElseIf Texto = "[FIM]" Then
            'Print #1, "                    CORTE AQUI                  "
            'Print #1, " -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  - "
            Printer.Print "."
            Printer.Print "."
            Printer.Print left("Data e Hora impressão: " & Format(Date, "DD/MM/YYYY") & " - " & Format(Time, "HH:MM:SS") & Space(48), 48)
            'Print #1, " "
            'Print #1, " "
            'Print #1, " "
            'Print #1, " "
            'Print #1, " "
            'Print #1, " "
            Printer.EndDoc
            'Close #1
            
        Else
            Printer.Print Texto
        End If

    End If
End Sub

Public Sub CarregaValoresTransfNumerario(protocolo As Long)

    wTNDinheiro = 0
    wTNCheque = 0
    wTNVisa = 0
    wTNRedecard = 0
    wTNAmex = 0
    wTNHiperCard = 0
    wTNBNDES = 0
    wTNVisaEletron = 0
    wTNRedeShop = 0
    wTNDeposito = 0
    wTNNotaCredito = 0
    wTNConducao = 0
    wTNDespLoja = 0
    wTNOutros = 0
    wTNTotal = 0
   
   Sql = ("SELECT MC_GrupoAuxiliar,MO_Descricao,SUM(MC_Valor) as Valor FROM MOVIMENTOCAIXA,MODALIDADE WHERE Mo_GRUPO=MC_GrupoAuxiliar" _
        & " AND MC_PROTOCOLO in (" & protocolo & ") AND MC_GRUPOAUXILIAR LIKE '30%'" _
        & "GROUP BY MC_GrupoAuxiliar,MO_DESCRICAO order by MC_GrupoAuxiliar")
    
       rdoTransfNumerario.CursorLocation = adUseClient
       rdoTransfNumerario.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
  If Not rdoTransfNumerario.EOF Then
     Do While Not rdoTransfNumerario.EOF
        If rdoTransfNumerario("MC_GrupoAuxiliar") = 30101 Then
           wTNDeposito = wTNDeposito + rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30201 Then
           wTNDeposito = wTNDeposito + rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30106 Then
           wTNDinheiro = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30107 Then
           wTNCheque = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30203 Then
           wTNRedeShop = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30206 Then
           wTNVisaEletron = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30301 Then
           wTNVisa = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30302 Then
           wTNRedecard = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30303 Then
           wTNAmex = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30304 Then
           wTNBNDES = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30103 Then
           wTNConducao = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30104 Then
           wTNDespLoja = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30205 Then
           wTNHiperCard = rdoTransfNumerario("Valor")
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30009 Then
            wtotalGarantia = rdoTransfNumerario("Valor")
           'wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30701" Then
           wTNNotaCredito = (wTNNotaCredito + rdoTransfNumerario("Valor"))
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30502" Then
           wTNFaturado = (wTNFaturado + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30602" Then
           wTNFinanciado = (wTNFinanciado + rdoTransfNumerario("Valor"))
        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30108" Then
           wTNOutros = (wTNOutros + rdoTransfNumerario("Valor"))
           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
        End If
           
       rdoTransfNumerario.MoveNext
       
     Loop
 End If
  rdoTransfNumerario.Close
End Sub

Public Function endIMG(nomeBotao As String) As String
    
    Dim Arquivo As String
    Dim enderecoArquivo As String
    
    enderecoArquivo = "c:\sistemas\dmac caixa\imagens\lojas\" & GLB_Logo & "_" & nomeBotao
    Arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If Arquivo = Empty Then
        enderecoArquivo = "c:\sistemas\dmac caixa\imagens\" & nomeBotao
    End If
    
    endIMG = enderecoArquivo
    
End Function

Public Sub notificacaoEmail(Mensagem As String)
    
    ConectaODBCMatriz
    
    Dim teste As String
    Mensagem = Replace(Mensagem, "'", "''")
    ')
    
    Sql = "insert into alerta_movimento_email" & vbNewLine & _
          "(AME_Numero,AME_NroCaixa, AME_Mensagem, AME_Usuario, AME_DataHora, AME_Situacao, AME_Loja)" & vbNewLine & _
          "values (" & GLB_ADMProtocolo & ", " & GLB_Caixa & ", '" & Mensagem & "', '" & GLB_ADMNome & "'," & vbNewLine & _
          "'" & Format(Date, "YYYY/MM/DD") & " " & Time & "', '" & "A" & "', '" & GLB_Loja & "')" & vbNewLine & _
          ""
    rdoCNRetaguarda.Execute (Sql)
End Sub


Public Function validaNumeroCupom() As Boolean

    Dim rdoValida As New ADODB.Recordset
    Dim Sql As String

    Screen.MousePointer = vbHourglass

    validaNumeroCupom = False

    
    Sql = "Select count(nf) as qtdeNF from nfcapa " _
    & "where NF = " & wNumeroCupom & " AND serie = '" & GLB_SerieCF & "'"
    
    rdoValida.CursorLocation = adUseClient
    rdoValida.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If rdoValida("qtdeNF") < 1 Then
            validaNumeroCupom = True
        End If
    
    rdoValida.Close
    
    Screen.MousePointer = vbNormal

End Function


Public Sub defineImpressora()
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_Impressora00) Then
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    Printer.FontName = "Lucida Console"
    Printer.FontSize = 7
    
End Sub


Public Function EnviaEmail(ByVal numped)
    
    Dim Sql As String
    Dim rsEmail As New ADODB.Recordset
    
    Dim Nf As String
    Dim Serie As String
    Dim condpag As String
    Dim loja As String
    
    On Error GoTo TrataErro

    
    Sql = "select lojaOrigem, nf, serie, condpag from nfcapa where numeroped = " & numped & " and condpag >= 3"
    
    rsEmail.CursorLocation = adUseClient
    rsEmail.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsEmail.EOF Then
        Nf = Trim(rsEmail("nf"))
        Serie = Trim(rsEmail("serie"))
        loja = Trim(rsEmail("lojaOrigem"))
        Sql = "Exec SP_ALERTA_FATURADA '" & loja & "', " & Nf & ",'" & Serie & "'"
        rdoCNLoja.Execute (Sql)
    End If

Exit Function

TrataErro:
    MsgBox "Erro ao criar e-mail do faturado (E-mail Alerta)" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical

End Function




Public Function Devolucao(pedido As String)
 Dim adoDevolucao As New ADODB.Recordset
 Dim SqlDev As String
 
 
 
 ReImpressao_Dev = False
 SqlDev = " Select NF,SERIE,NfDevolucao,SerieDevolucao,DATAEMI,TOTALNOTA,NotaCredito,tiponota from  nfcapa where serie='NE' and tiponota='E' and numeroped=" & pedido
     adoDevolucao.CursorLocation = adUseClient
    adoDevolucao.Open SqlDev, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not adoDevolucao.EOF Then
                If Trim(adoDevolucao("tiponota")) = "E" Then
                         Nf_Dev = adoDevolucao("NF")
                         Serie_Dev = adoDevolucao("SERIE")
                         NfDev_Dev = adoDevolucao("NfDevolucao")
                         SerieDev_Dev = adoDevolucao("SerieDevolucao")
                         DataDev_Dev = Format(adoDevolucao("DATAEMI"), "DD/MM/YYYY")
                         ValorNotaCredito_Dev = Format(adoDevolucao("TOTALNOTA"), "###,###,###0.00")
                         NotaCredito_Dev = adoDevolucao("NotaCredito")
                         ReImpressao_Dev = True
                        
                End If
                
    End If
    adoDevolucao.Close
End Function

Public Function CriaNotaCredito1(ByVal Nf As Double, ByVal Serie As String, ByVal NfDev As Double, ByVal SerieDev As String, ByVal DataDev As String, ByVal ValorNotaCredito As Double, ByVal NotaCredito As Double, ByVal ReImpressao As Boolean)
    Dim rsDadosNfCapa As New ADODB.Recordset
    Dim rsVerLoja As New ADODB.Recordset
    Dim rsDataEmiDevol As New ADODB.Recordset
    Dim Linha1 As String
    Dim wTotalNota As Double
    Dim wValorExtenso As String
    Dim wDataEmiDevolucao As Date
    ConectaODBCMatriz
   
    For Each NomeImpressora In Printers
        If Trim(UCase(NomeImpressora.DeviceName)) = UCase(Glb_ImpNotaFiscal) Then
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    Sql = "Select * from NfCapa, fin_cliente " _
        & "where Nf=" & Nf & " and cliente = ce_codigocliente " _
        & "and Serie='" & Serie & "'"
    rsDadosNfCapa.CursorLocation = adUseClient
    rsDadosNfCapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rsDadosNfCapa.EOF Then
        If ReImpressao = True Then
            Sql = "Select DataEmi From NfCapa Where NF = '" & NfDev & "' and " _
                & "Serie = '" & SerieDev & "' and Lojaorigem = '" & Trim(GLB_Loja) & "'"
                
             rsDataEmiDevol.CursorLocation = adUseClient
             rsDataEmiDevol.Open Sql, rdoCNRetaguarda, adOpenForwardOnly, adLockPessimistic
            
            If rsDataEmiDevol.EOF Then
                rsDataEmiDevol.Close
                
                MsgBox "Irei conectar na Retaguarda para localizar a nota." & Chr(10) & "Pois a mesma não foi encontrado no BANCO LOCAL", vbInformation + vbOKOnly
                
                If GLB_ConectouOK = True Then
                    Sql = ""
                    Sql = "Select DataEmi From NfCapa Where NF = " & rsDadosNfCapa("NfDevolucao") & " and " _
                        & "Serie = '" & rsDadosNfCapa("SerieDevolucao") & "' and Lojaorigem = '" & Trim(GLB_Loja) & "'"
    
                    rsDataEmiDevol.CursorLocation = adUseClient
                    rsDataEmiDevol.Open Sql, rdoCNRetaguarda, adOpenForwardOnly, adLockPessimistic
             
                    If rsDataEmiDevol.EOF Then
                        wDataEmiDevolucao = Date
                    Else
                        wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                    End If
                    rsDataEmiDevol.Close

                End If
            Else
                wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                rsDataEmiDevol.Close
            End If
        Else
                wDataEmiDevolucao = DataDev
        End If
            
        Sql = ""
        Sql = "Select CTS_Loja,LO_Razao,CTS_numeroNCredito,LO_Razao,Loja.* from ControleSistema,Loja " _
            & "where LO_Loja=CTS_Loja"
        rsVerLoja.CursorLocation = adUseClient
        rsVerLoja.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       
        If Not rsVerLoja.EOF Then
            If Serie = "SM" Then
                wTotalNota = rsDadosNfCapa("TotalNotaAlternativa")
            Else
                wTotalNota = ValorNotaCredito
            End If

            For I = 1 To 4

                Printer.ScaleMode = vbMillimeters
                Printer.FontName = "Romam"
                Printer.FontSize = 9
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                
                If Len(Trim(rsVerLoja("LO_Razao"))) < 1 Then
                    If ReImpressao = False Then
                        Printer.Print Space(2) & rsVerLoja("CTS_Razao")
                    Else
                        Printer.Print Space(2) & rsVerLoja("CTS_Razao") & Space(84) & "RE-IMPRESSAO"
                    End If
                Else
                    If ReImpressao = False Then
                        Printer.Print Space(2) & Trim(Mid(rsVerLoja("LO_Razao"), 1, Len(rsVerLoja("LO_Razao"))))
                    Else
                        Printer.Print Space(2) & Trim(Mid(rsVerLoja("LO_Razao"), 20, Len(rsVerLoja("LO_Razao")))) & Space(84) & "RE-IMPRESSAO"
                    End If
                End If
                
                Printer.Print Space(2) & left(rsVerLoja("LO_Endereco") & Space(30), 30) _
                    & "    -    " & rsVerLoja("LO_Cep") & "   -   " & rsVerLoja("LO_Municipio") _
                    & right(Space(72) & "NOTA DE CREDITO", 72)
                Printer.Print Space(2) & "FONE : " & "(" & right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" _
                        & left(rsVerLoja("LO_Telefone") & Space(10), 10) & " -  " _
                        & "FAX : " & "(" & right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" & left(rsVerLoja("LO_Telefone") & Space(10), 10)
                Printer.Print Space(2) & "C.G.C : " & left(rsVerLoja("LO_CGC") & Space(25), 25) & "INSCR.EST. : " & rsVerLoja("LO_InscricaoEstadual")
                Printer.Print Space(140) & "NUM.  " & right(String(9, "0") & NotaCredito, 9) & right(Space(10) & I & "a.VIA", 10)
                Printer.Print Space(2) & "A"
                Printer.Print Space(2) & rsDadosNfCapa("ce_razao")
                Printer.Print Space(2) & left(rsDadosNfCapa("ce_endereco") & Space(130), 130) & left("DATA : " & rsDadosNfCapa("DataEmi") & Space(18), 18)
                Printer.Print Space(2) & rsDadosNfCapa("ce_Municipio") & "  -   " & rsDadosNfCapa("ce_estado")
                Printer.Print Space(2) & "FONE : " & rsDadosNfCapa("ce_telefone")
                Printer.Print Space(2) & "EFETUAMOS NESTA DATA EM SUA CONTA CORRENTE O SEGUINTE LANÇAMENTO:"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print Space(40) & "HISTORICO" & Space(40) & "| DEBITO" & Space(30) & "| CREDITO"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(2) & "PELO RECEBIMENTO DA MERCADORIA EM DEVOLUÇÃO"
                Printer.Print Space(2) & "CONFORME NF " & Nf & " SERIE " & Serie & " DE " & rsDadosNfCapa("DataEmi")
                Printer.Print Space(2) & "NO VALOR DE R$          " & Format(wTotalNota, "###,###,###0.00")
                Printer.Print
                Printer.Print Space(2) & "REFERENTE NF DE VENDA " & NfDev & " - " & SerieDev & " DE " & wDataEmiDevolucao
                Printer.Print Space(2) & "DA LOJA " & rsDadosNfCapa("LojaOrigem") & Space(140) & Format(wTotalNota, "###,###,###0.00")
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(142) & "ATENCIOSAMENTE"
                Printer.Print
                Printer.Print Space(120) & "_______________________________________"
                If Len(Trim(rsVerLoja("LO_Razao"))) < 1 Then
                    Printer.Print Space(120) & rsVerLoja("CTS_Razao")
                Else
                    Printer.Print Space(120) + Space((39 - Len(Trim(rsVerLoja("LO_Razao"))))) & Trim(Mid(rsVerLoja("LO_Razao"), 1, Len(rsVerLoja("LO_Razao"))))
                End If
                Printer.Print
                Printer.Print
                If I = 2 Then
                    Printer.NewPage
                End If
            Next I
            Printer.EndDoc
        End If
    End If


End Function


Public Function obterReferenciaPorItem(numeroPed As String, Item As String) As String
    Dim Sql As String
    Dim rsNFECapa As New ADODB.Recordset
    
    Sql = "select REFERENCIA" & vbNewLine & _
          "from nfitens " & vbNewLine & _
          "where NUMEROPED = '" & numeroPed & "' " & vbNewLine & _
          "AND item = '" & Item & "'"
    
    rsNFECapa.CursorLocation = adUseClient
    rsNFECapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If rsNFECapa.EOF = False Then
            obterReferenciaPorItem = rsNFECapa("REFERENCIA")
        Else
            obterReferenciaPorItem = "[REFERENCIA NÃO ENCONTRADA]"
        End If
    
    rsNFECapa.Close
    
End Function

Public Sub criaDuplicataBanco()
    On Error GoTo TrataErro

    Sql = "exec Sp_Cria_Duplicatas '" & wLoja & "','" & Format(Date, "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "'"
    rdoCNLoja.Execute (Sql)
    
Exit Sub

TrataErro:
    MsgBox "Erro ao criar as duplicatas no banco de dados" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical
    
End Sub
