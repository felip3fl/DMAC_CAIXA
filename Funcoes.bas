Attribute VB_Name = "Funcoes"


'Ler os Valores dos parâmetros nas seções do arquivo ini
Function LeParametrosIni(Secao As String, Label As String) As String
  
   Const TamanhoParametro = 80
   Dim ParametroIni As String * TamanhoParametro
   Dim RetornoFuncao
   Dim ArquivoIni As String
   Dim Contador As Integer
   ParametroIni = ""
     
   RetornoFuncao = GetSystemDirectory(ParametroIni, TamanhoParametro)
   ArquivoIni = Left(ParametroIni, RetornoFuncao) + "\BemaFI32.ini"
   ParametroIni = ""
   RetornoFuncao = GetPrivateProfileString(Secao, Label, "-2", ParametroIni, TamanhoParametro, ArquivoIni)
   RetornoFuncao = Mid(ParametroIni, 1, 2)
   If Val(RetornoFuncao) <> -2 Then
       Contador = 1
       Do
           Tst = Mid(ParametroIni, Contador, 1)
           If Asc(Tst) <> 0 Then
               Contador = Contador + 1
           End If
       Loop While ((Asc(Tst) <> 0) And (Contador < Len(ParametroIni)))
       RetornoFuncao = Mid(ParametroIni, 1, Contador)
   End If
   LeParametrosIni = RetornoFuncao
End Function


Public Sub CentralizaJanela(Form As Form)
    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.Width - Form.Width) / 2
End Sub

Public Function AnalisaFlagsFiscais(FlagFiscal As Integer) As String
    Dim StringRetorno As String
    
    If (FlagFiscal >= 128) Then
        StringRetorno = "Memória fiscal lotada" & vbCr
        FlagFiscal = FlagFiscal - 128
    End If
    
    If (FlagFiscal >= 32) Then
        StringRetorno = StringRetorno & "Permite o cancelamento do cupom" & vbCr
        FlagFiscal = FlagFiscal - 32
    End If
    
    If (FlagFiscal >= 8) Then
        StringRetorno = StringRetorno & "Já houve redução 'Z' no dia" & vbCr
        FlagFiscal = FlagFiscal - 8
    End If
    
    If (FlagFiscal >= 4) Then
        StringRetorno = StringRetorno & "Horário de verão selecionado" & vbCr
        FlagFiscal = FlagFiscal - 4
    End If
        
    If (FlagFiscal >= 2) Then
        StringRetorno = StringRetorno & "Fechamento de formas de pagamento iniciado" & vbCr
        FlagFiscal = FlagFiscal - 2
    End If
    
    If (FlagFiscal >= 1) Then
        StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
        FlagFiscal = FlagFiscal - 1
    End If

    AnalisaFlagsFiscais = StringRetorno

End Function


Public Function AnalisaStatusCheque(StatusCheque As Integer) As String
    Dim StringRetorno As String
    
    If (StatusCheque = 1) Then
        StringRetorno = "Impressora ok." & vbCr
    
    ElseIf (StatusCheque = 2) Then
        StringRetorno = "Cheque em impressão." & vbCr
    
    ElseIf (StatusCheque = 3) Then
        StringRetorno = "Cheque posicionado." & vbCr

    ElseIf (StatusCheque = 4) Then
        StringRetorno = "Aguardando o posicionamento do cheque." & vbCr
    
    End If
    
    AnalisaStatusCheque = StringRetorno

End Function

Public Sub DestacaTexto(Objeto As TextBox)
    Objeto.SelStart = 0
    Objeto.SelLength = Len(Objeto.Text)
End Sub
