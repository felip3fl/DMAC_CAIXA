Attribute VB_Name = "ModCaixa"
Option Explicit
Public DBLoja As Database
Public DBFBanco As Database
Global FechouATelaFormapagamentonoX As Boolean
Global NroNotaFiscal As Long
Public PegaDadosCaixa As rdoResultset
Public VerificaCaixa As rdoResultset
Public RdoGravaDados As rdoResultset
Public NFCapaDBF As rdoResultset
Public NFItemDBF As rdoResultset
Public RsPegaNumNote As rdoResultset
Public rsPegaLoja As rdoResultset
Public RsVerificaPedido As rdoResultset
Public ISQL As rdoResultset
Public RSTipoControle As rdoResultset
Public RsPegaItensPedi As Recordset
Public RsPegaItensEspeciais As rdoResultset
Public FechaRetaguarda As rdoResultset
Dim rdoVersao As rdoResultset
Global rsLoja As rdoResultset
Global rdoParametro As New ADODB.Recordset
Global rdocontrole As rdoResultset
Global lsDSN As String

Global GLB_USU_Nome As String
Global GLB_USU_Codigo As String
Global GLB_CTR_Protocolo As Integer

Global wNumeroCaixa As Integer
Global wNomeservidor As String
Global wNomeBanco As String
Global Wusuario As String
Global wSenha As String

Global wNotaFiscalReemissao As Long
Global wSerieReemissao As String
Global wPegaCliente As String
Global wPegaDesconto As String
Global wPegaFrete As Double

Global Const AmareloGrid = &HC0FFFF
Global Const VerdeGrid = &HC0FFC
Global Const AzulGrid = &HFFFFC0

Global GLB_Cotacao As String
Global GLB_TrocaVersao As String
Global GLB_Versao As String
Global GLB_VersaoNova As String
Global TipoPedido As String
Global NomeImpressora As Printer
Global Wimpressora As String
Global ContaImpressora As Integer
Global wSair As Boolean
Global GLB_Banco As String
Global LiberaDesc As Boolean

Global GLB_Usuario As String
Global GLB_Senha As String
Global GLB_Servidor As String
Global GLB_NumeroCaixa As String
Global wlblloja As String

Global dbMovDia As Database
'Global RsDados As rdoResultset
Global RsAtivaLojaOnLine As rdoResultset
Global RsDadosDbf As rdoResultset
Global RsCapaNF As rdoResultset
Global RsItensNF As rdoResultset
Global RsICMSInter As rdoResultset
Global RsUsuario As rdoResultset
Global RsSelecionaMovCaixa As rdoResultset
Global RsSelecionaMovBanco As rdoResultset
Global RsSelecionaMovEstoque As rdoResultset
Global RsSelecionaDivEstoque As rdoResultset
Global RsCarimbo As rdoResultset
Global RsPegaControleMigracao As rdoResultset
Global RsDescProduto As rdoResultset
Global RsPegaDescricaoAlternativa As rdoResultset
Global RsNumeroECF As rdoResultset

Global PedidoSM As Boolean
Global FecParcial As Boolean
Global FecTotal As Boolean
Global Faturada As Boolean
Global Financiada As Boolean
Global WEmiteCupom As Boolean
Global Wcheque As Boolean
Global WAbrirCaixa As Boolean
Global wVerificaImpressoraFiscal As Boolean
Global wNotaDoDia As Boolean
Global wPegaDescricaoAlternativa As String
Global SQL As String
Global wPegaImpressora As String
Global wPegaLojaControle As String
Global wPegaCarimboNF As String
Global wDetalheImpressao As String
Global FundoMdi_Versao As String
Global WNF As String
Global wLoja As String * 5
Global wRazao As String
Global wNovaRazao As String
Global WCGC As String
Global WIest As String
Global Wendereco As String
Global wbairro As String
Global WMunicipio As String
Global westado As String
Global WCep As String
Global WFone As String
Global wDDDLoja As String
Global WFax As String
Global wNumeroCupom As String * 6
Global Wserie As String
Global WENDCLI As String
Global WMUNCLI As String
Global WUF As String
Global WCGCCLI As String
Global WIECLI As String
Global wCFO1 As String * 6
Global wCFO2 As String * 6
Global WNomeCliente As String
Global Wtipovenda As String
Global wCondPag As String
Global Wav As String
Global WdataPag As String
Global wSituacao As String
Global WSTATUS As String
Global Wtipo As String
Global WOutraLoja As String
Global WNOMCLI As String
Global WENDENTCLI As String
Global WREGIAO As String
Global Wlojat As String
Global WTipoP As String
Global WTipoNota As String
Global WdataPg As String
Global WdataEnt As String
Global Wdata As String
Global wAnexo As String
Global wAnexo1 As String
Global wAnexo2 As String
Global wRecebeCarimboAnexo As String * 132
Global Wteste As String
Global wReferencia As String
Global wPepaUsuario As String
Global WSERIE1 As String
Global WSERIE2 As String
Global WSERIE3 As String
Global WSUBTRIBUT As String
Global wCodBarra As String
Global glb_AbilitaCaixa As String * 1
Global glb_ECF As String
Global WAnexoAux As String * 20
Global WCFOAux As String * 25
Global WcaminhoTextos As String
Global WcaminhoTextosAtu As String
Global wCaminhoAtualizacao As String
Global wMigracao As String
Global wAtualizaCentral As String
Global wPegaUsuario As String
Global WbancoAccess As String
Global WbancoDbf As String
'Global Wusuario As String
Global WNatureza As String
Global WVendedor As String
Global wCarimbo4 As String * 132
Global wStr0, wStr1, wStr2, wStr3, wStr4, wStr5, wStr6, wStr7 As String
Global wStr8, wStr9, wStr10, wStr11, wStr12, wStr13, wStr15, wStr16, wStr17, wStr18, wStr19, wStr20, wStr21 As String
Global Wcondicao As String * 30
Global wVerificaLojaOnLine As String * 1
Global arquivo As String
Global buffer As String
Global wDescricao As String
Global wLinhaCarimbo As String
Global wReferenciaEspecial As String
Global wSerieProd1 As String
Global wSerieProd2 As String
Global wCarimbo5 As String * 132

Global wAliqICMSInterEstadual As Double

Global wPedidoCliente  As Double
Global wDescEspecial As Double

Global Wentrada As Double
Global WcontaArq As Integer
Global wNumeroECF As Integer

Global WCliente As Long
Global WnumeroPed As Long
Global WnumeroNotaDbf As Long
Global WnumeroPedidoDbf As Long
Global Woperacao As Long
Global Wpedcli As Long
Global wQtde As Long
Global wlin As Long
Global tmporient As Long
Global wConta As Long
Global wChave As Long
Global wReduz As Long
Global wCodIPI As Long
Global wCodTri As Long
Global i As Long
Global flg As Long

Global wConfereCodigoZero As Double
Global WbaseIcm As Double
Global PORICM As Double
Global VLRICM As Double
Global GLB_TotalICMSCalculado As Double
Global GLB_ValorCalculadoICMS As Double
Global GLB_BasedeCalculoICMS As Double
Global GLB_AliquotaAplicadaICMS As Double
Global GLB_AliquotaICMS As Double
Global GLB_BaseTotalICMS As Double
Global GLB_Tributacao As String * 3
Global wCFOItem As Double
Global wQuantItensCapaNF As Double
Global wQuantItensNF As Double
Global wQuant As Double
Global wComissaoVenda As Double
Global wSomaVenda As Double
Global wSomaMargem As Double
Global wPessoa As Double
Global WTotPedido As Double
Global Wdescontop As Double
Global wSubTotal As Double
Global WPGENTRA As Double
Global wValFrete As Double
Global WFRETECOBR As Double
Global wChaveICMS As Double
Global wChaveICMSItem As Double
Global WVALFRETECB As Double
Global wVlUnit As Double
Global wVlUnit2 As Double
Global wVlTotItem As Double
Global WVLUNITAL As Double
Global WVLTOTITEMAL As Double
Global WREFALTERNA As Double
Global wCancelaVenda As Double
Global WDESCRAT As Double
Global WENTRAT As Double
Global WVLIPI As Double
Global wPLISTA As Double
Global WCMR As Double
Global wICMS As Double
Global WBCOMIS As Double
Global WCSPROD As Double
Global WVBUNIT As Double
Global WPERDESC As Double
Global wPegaSequenciaCO As Double
Global wTotalPed As Double

Global ValorlItem As Double
Global Valordesconto As Double
Global TotalVenda As Double

Global wUltimoItem As Integer

Global WCOMISSAO As Integer
Global WCODOPER As Integer
Global wQtdItem As Integer
Global WTm As Integer
Global WPesoBr As Integer
Global WPesoLq As Integer
Global WOUTROVEND As Integer
Global WDDD As Integer
Global wAnexoIten As Integer
Global WTP As Integer
Global WTRIBUTO As Integer
Global WCONTROLE As Integer
Global wItem As Integer
Global wLinha As Integer
Global wSecao As Integer
Global wIcmPdv As Integer
Global WParcelas As Integer
Global Wecf As Integer
Global wSubstituicaoTributaria As Integer
Global wECFNF As Integer
Global wPagina As Integer
Global wQuantdadeTotalItem As Integer
Global wTipoMovimentacao As Integer
Global wNumPed As String

Global Wimprimecheque As Boolean
Global WTotalCheque As Double
Global WDESCRIPAG As String
Global AbreMigracao As String
Global AbreAtualizacaoOnLine As String
Global Wdate As Date

Global wReemissaoNotaFiscal As Boolean
Global wVerificaTM As Boolean


Global WcontaTempo As Long
Global Wconectou As Boolean
Global Conexao As New rdoConnection
Global ConexaoBach As New rdoConnection
Global RdoDados As rdoResultset
Global Servidor As String
Global WBANCO As String
Global HandlerWindow As Long
Global rtn
Global NomeArquivo As String
Global Temporario As String
Global NotaFiscal As Long
Global Wsm As Boolean
Global Pedido As Double
Global Nota As Double

Global wValorTotalCodigoZero As Double
Global Total As Double
Global SubTotal As Double
Global TotNota As Double
Global wValorICMSAlternativa As Double
Global wBaseICMSAlternativa As Double
Global wValorTotalMercadoriaAlternativa As Double
Global wValorTotalItemAlternativa As Double
Global wValoTotalNotaAlternativa As Double
Global wTotalNotaAlternativa As Double
Global DataEmi As String
Global CODVEND As Double
Global VLRMERC As Double
Global Desconto As Double
Global tipovenda As String
Global CondPagto As Double
Global av As Double
Global Cliente As String
Global NatOper As Double
Global DataPag As String
Global PgEntra As Double
Global lojat As String
Global TOTITENS As Double
Global PEDCLI As Double
Global PesoBr As Double
Global PesoLq As Double
Global OUTRALOJA As String
Global ValFrete As Double
Global FreteCobr As Double
Global OUTROVEND As String
Global notafis As Double
Global BASEICM As Double
Global Hora As String
Global TOTIPI As Double
Global nomecli As String
Global endcli As String
Global muncli As String
Global cgccli As String
Global fonecli As String
Global Pessoa As String
Global ufcli As String
Global cepcli As String
Global bairrocli As String
Global TipoNota As String
Global ECF As Double
Global numerosf As Double
Global Processado As String
Global X As Integer
Global wNumSeqNF As Double
Global wValorMercadoriaAlternativa As Double
Global wPrecoUnitarioAlternativa As Double
Global Referencia As String
Global Quant As Double
Global unidade As String
Global PrecoUni As String
Global valormerc As Double
Global aliqipi As Double
Global plista As Double
Global icms As Double
Global Comissao As Double
Global bcomis As Double
Global Linha As Double
Global Secao As Double
Global csprod As String
Global vlripi As Double
Global Item As Double
Global PedidoItem As Double
Global tipomov As Double
Global sitnf As String
Global Status As String
Global wCarimbo1 As String * 132
Global wCarimbo2 As String * 132
Global wCarimbo3 As String * 132
Global wVendedorLojaVenda As String
Global wLojaVenda As String

Global WnumeroPedido As Integer
Global wTotalNotaTransferencia As Double

Global wGravaCheque As Boolean
Global WNfTransferencia As String

Global wNotaDevolucao As Boolean
Global wReemissao As Boolean
Global wNfCapa As Boolean
Global wNFitens As Boolean
Global wNotaTransferencia As Boolean
Global wNumeroCaixaUso As Integer
Global Glb_SeriePed As String * 2
Global Glb_Nf As Double
Global Glb_NfDevolucao As Boolean
Global Glb_SenhaEstoque As String
Global Glb_UsuarioEstoque As String
Global glb_LiberaSenha As Integer
Dim RsTravaCaixa As rdoResultset
Global Glb_BancoLocal As String
Global Glb_ServidorLocal As String
Global rdoCNLoja As New rdoConnection
Global rdoCnLojaBach As New rdoConnection

'Global rdoCNLojaINI As New rdoConnection
'Global RsDadosINI As New rdoConnection


Global rdoCNLojaINI As New ADODB.Connection
Global RsDadosINI As New ADODB.Recordset

Global Glb_ImpNotaFiscal As String
Global GLB_ImpCotacao As String
Global VerDataEstoque As Boolean

Global CodigoZero As Boolean
Global DescEspecial As Boolean

Global ConexaoCD As New rdoConnection
Global PegaDiretorio As String

'========================================
Global wAtualizaVersao As Boolean
Global wArquivoSaida As String
Global STRI As String
Global Versao_Atual As String
Global DirApp As String
Global rsChecaVersaonoFTP As rdoResultset
Global wValorRetorno As String

Global wAbrirTela As String
'========================================

Global wOutraSaida As Boolean
Global wClienteNFPaulista As String
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Sub Main()
    Dim VersaoEXE As String
    
    Call AbilitarCaixa

    'Call Shell("C:\TrocaVersao.bat")
    'Call Shell("C:\pkunzip c:\versao~1\bc2000.zip")
    'FileCopy "C:\bc2000.exe", "C:\Versao\bc2000.exe"
    
    
    
    'If ComparaDataVersao(GLB_Versao, GLB_VersaoNova) = True Then
    '    Call Shell(GLB_TrocaVersao, vbNormalFocus)
    '    End
    'End If
    
    
    If ConectaOdbcLocal(rdoCNLoja, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = False Then
        MsgBox "Não foi possivel conectar-se ao banco de dados", vbCritical, "Aviso"
        Exit Sub
    Else
        ConectaOdbcLocal rdoCnLojaBach, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)
    End If
    'Set DBLoja = Workspaces(0).OpenDatabase(WbancoAccess)
    
    'Set DBFBanco = Workspaces(0).OpenDatabase(WbancoDbf, False, False, "DBase IV")
    
    '
    '------------------------------Verifica situação do sistema-----------------------
    '
    'SQL = ""
    'SQL = "Select CT_SituacaoCaixa from ControleECF where CT_SituacaoCaixa='T' and CT_ECF=" & Val(glb_ECF) & ""
        'Set RsTravaCaixa = rdocnloja.OpenResultset (SQL)
    'If Not RsTravaCaixa.EOF Then
        'MsgBox "O sistema está travado por falta de conexão com a central, " _
            & "tire nota manual e entre em contato com o Fernando Alfano para destravar o sistema", vbCritical, "Aviso"
        'frmLogin.Show
        'Exit Sub
    'End If
    '**********************************************************************************
    
    
    If VerificaExeRodando("Migracao.exe") = False Then
        If wMigracao <> "" Then
           'SQL = ""
           'SQL = "Select CT_ControleMigracao from Controle where CT_ControleMigracao = 'F' "
           '    Set RsPegaControleMigracao = rdocnloja.OpenResultset (SQL)
        
           'If Not RsPegaControleMigracao.EOF Then
               AbreMigracao = Shell(wMigracao, 1) 'Abre Migracao.exe
               SQL = ""
               SQL = "Update Controle set CT_ControleMigracao = 'A'"
               rdoCNLoja.Execute (SQL)
           'End If
        End If
    End If
    
    If VerificaExeRodando("AtualizaCentral.exe") = False Then
        If wAtualizaCentral <> "" Then
            'SQL = ""
            'SQL = "Select CT_OnLine from Controle where CT_AtualizacaoCentral = 'F' "
            '    Set RsAtivaLojaOnLine = rdocnloja.OpenResultset (SQL)
            'If Not RsAtivaLojaOnLine.EOF Then
                AbreAtualizacaoOnLine = Shell(wAtualizaCentral, 1) 'Abre AtualizaCental.exe
                SQL = ""
                SQL = "Update Controle set CT_AtualizacaoCentral = 'A' "
                    rdoCNLoja.Execute (SQL)
            'End If
        End If
    End If
    
    SQL = ""
    SQL = "Select CT_VersaoBalcao2000 From Controle"
    Set rdoVersao = rdoCNLoja.OpenResultset(SQL)
        
   ' VersaoEXE = App.Major & "." & App.Minor & "." & App.Revision
    
   ' If IIf(IsNull(rdoVersao("CT_VersaoBalcao2000")), "0.0.0", rdoVersao("CT_VersaoBalcao2000")) < Trim(VersaoEXE) Then
   '     SQL = ""
   '     SQL = "Update Controle Set CT_VersaoBalcao2000 = '" & Trim(VersaoEXE) & "'"
   '     rdoCnLoja.Execute (SQL)
   ' ElseIf rdoVersao("CT_VersaoBalcao2000") > Trim(VersaoEXE) Then
   '     Set mdiBalcao.Picture = LoadPicture(App.Path & "/fundomdi_versao.bmp")
   ' End If
    
    
    If GLB_NumeroCaixa = 1 Then
        SQL = ""
        SQL = "Update Controle set CT_Balcao = 'A'"
            rdoCNLoja.Execute (SQL)
    End If
    mdiBalcao.Show

End Sub

Public Function ValidaAbertura()

    SQL = "Select max(Ct_Data) as DataMov,max(Ct_Sequencia) as Seq from CTCaixa " _
        & "where CT_NumeroECF = " & Val(glb_ECF) & ""
        Set VerificaCaixa = rdoCNLoja.OpenResultset(SQL)

    If (Not VerificaCaixa.EOF) And (VerificaCaixa("Seq") > 0) Then
       SQL = "Select * from CTCaixa " _
            & "where ct_data= '" & Format(VerificaCaixa("datamov"), "mm/dd/yyyy") & "' " _
            & "and Ct_Sequencia = " & VerificaCaixa("Seq") & " " _
            & "and CT_NumeroEcf=" & Val(glb_ECF) & "   "
       Set PegaDadosCaixa = rdoCNLoja.OpenResultset(SQL)
       
       If Not PegaDadosCaixa.EOF Then
          If PegaDadosCaixa("ct_Data") = Date And PegaDadosCaixa("ct_Situacao") = "A" Then
             If FecTotal = False Then
                 'frmCaixa.Show
                If wAbrirTela = 1 Then
                   frmCaixaTEFPedido.Show
                ElseIf wAbrirTela = 2 Then
                       frmCaixaNF.Show
                ElseIf wAbrirTela = 3 Then
                       frmCaixaTEF.Show
                End If
             End If
          ElseIf PegaDadosCaixa("Ct_Data") <> Date And PegaDadosCaixa("Ct_Situacao") = "A" Then
             MsgBox "Data do caixa incorreta.Favor efetuar o Fechamento Geral.", vbCritical, "Atenção"
             FecTotal = True
             ValidaFechamento
             Exit Function
          ElseIf PegaDadosCaixa("Ct_Situacao") = "P" Then
                frmAbrirFecharCaixa.cmdCaixaAberto.Visible = True
                frmAbrirFecharCaixa.cmdCaixaFechado.Visible = False
                frmAbrirFecharCaixa.Caption = "Abertura do Caixa"
                frmAbrirFecharCaixa.Show
          ElseIf PegaDadosCaixa("Ct_Situacao") = "T" Then
             If PegaDadosCaixa("Ct_Data") = Date Then
                MsgBox "Não é possível efetuar abertura, fechamento geral já foi concluído.", vbCritical, "Atenção"
                Exit Function
             ElseIf PegaDadosCaixa("Ct_Data") <> Date Then
                FecTotal = True
                frmAbrirFecharCaixa.cmdCaixaAberto.Visible = True
                frmAbrirFecharCaixa.cmdCaixaFechado.Visible = False
                frmAbrirFecharCaixa.Caption = "Abertura do Caixa"
                frmAbrirFecharCaixa.Show
             End If
          End If
       End If
    Else
       FecTotal = True
       frmAbrirFecharCaixa.cmdCaixaAberto.Visible = True
       frmAbrirFecharCaixa.cmdCaixaFechado.Visible = False
       frmAbrirFecharCaixa.Caption = "Abertura do Caixa"
       frmAbrirFecharCaixa.Show
    End If


End Function

Public Function ValidaFechamento()
    Set VerificaCaixa = rdoCNLoja.OpenResultset("Select max(Ct_Data) as datamov,max(Ct_Sequencia) as Seq from CTCaixa where CT_NumeroECF = " & glb_ECF & "")
    
    If Not VerificaCaixa.EOF Then
       Set PegaDadosCaixa = rdoCNLoja.OpenResultset("Select * from CTCaixa where ct_data= '" & Format(VerificaCaixa("datamov"), "mm/dd/yyyy") & "' and Ct_Sequencia = " & VerificaCaixa("Seq") & "  and CT_NumeroECF = " & glb_ECF & "")
       
       If Not PegaDadosCaixa.EOF Then
          If FecTotal = True Then
             If PegaDadosCaixa("Ct_Situacao") = "T" Then
                MsgBox "Fechamento geral já foi efetuado.", vbInformation, "Informação"
                Exit Function
             ElseIf PegaDadosCaixa("Ct_Situacao") = "A" Then
                If MsgBox("Deseja realmente fazer o fechamento geral do caixa ?", vbQuestion + vbYesNo, "Fechamento Geral") = vbYes Then
                   frmAbrirFecharCaixa.cmdCaixaFechado.Visible = True
                   frmAbrirFecharCaixa.cmdCaixaAberto.Visible = False
                   frmAbrirFecharCaixa.Caption = "Fechamento do Caixa"
                   frmAbrirFecharCaixa.Show
                Else
                   Exit Function
                End If
             ElseIf PegaDadosCaixa("Ct_Situacao") = "P" Then
                MsgBox "Não existe caixa aberto para validar o fechamento geral.", vbInformation, "Informação"
                Exit Function
             End If
          ElseIf FecParcial = True Then
             If PegaDadosCaixa("Ct_Situacao") = "P" Or PegaDadosCaixa("Ct_Situacao") = "T" Then
                MsgBox "Não há caixa aberto.", vbInformation, "Informação"
                Exit Function
             Else
                frmAbrirFecharCaixa.cmdCaixaFechado.Visible = True
                frmAbrirFecharCaixa.cmdCaixaAberto.Visible = False
                frmAbrirFecharCaixa.Caption = "Fechamento do Caixa"
                frmAbrirFecharCaixa.Show
             End If
          End If
       End If
    End If
    
'    VerificaCaixa.Close
'    PegaDadosCaixa.Close
End Function


Public Function AtualizanfItem()


   If frmCaixa.txtPedido.Text <> "" Then
       SQL = "Update NfItens set nf = " & WNF & ",Serie= '" & Wserie & "', " _
           & "TipoNota='V', DataEmi = '" & Format(Date, "MM/DD/YYYY") & "' Where numeroped = " & frmCaixa.txtPedido.Text & ""
       rdoCNLoja.Execute (SQL)
   Else
       SQL = "Update NfItens set nf = " & WNF & ",Serie= '" & Wserie & "'," _
           & "TipoNota='V',DataEmi = '" & Format(Date, "MM/DD/YYYY") & "' Where numeroped = " & WnumeroPed & ""
       rdoCNLoja.Execute (SQL)
   End If
       
End Function


Public Function Atualizanfcapa()

    If frmCaixa.txtPedido.Text <> "" Then
       SQL = "Update NfCapa set nf = " & WNF & ",Serie= '" & Wserie & "', " _
           & "TipoNota='V',hora=  getdate() , DataEmi = '" & Format(Date, "MM/DD/YYYY") & "' Where numeroped = " & frmCaixa.txtPedido.Text & ""
       rdoCNLoja.Execute (SQL)
    Else
       SQL = "Update NfCapa set nf = " & WNF & ",Serie= '" & Wserie & "', " _
           & "TipoNota='V',hora= getdate() ,DataEmi = '" & Format(Date, "MM/DD/YYYY") & "' Where numeroped = " & WnumeroPed & ""
       rdoCNLoja.Execute (SQL)
    End If



'    If Wserie = "SM" And frmCaixa.cmdCodigoZero.Caption = "CO" Then
'        SQL = "Update EvDesDBF set NotaFis = " & WNF & " where NumPed = " & frmCaixa.txtPedido.Text & " "
'            rdocnloja.Execute (SQL)
'    End If
End Function

Public Function EmiteCupom()

    SQL = ""
    SQL = "Select ct_ecf,ct_loja from controle"
    
    Set RsDados = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsDados.EOF Then

       wLoja = Trim(RsDados("Ct_LOJA"))
       
       If Trim(RsDados("Ct_ECF")) = "S" Then
       
 '*** Desabilitado 12/2009  WEmiteCupom = True
       End If
          WEmiteCupom = False
    End If



End Function


Public Function DadosLoja()

    SQL = ""
    SQL = "Select CT_Loja,CT_Razao,CT_NovaRazao,Lojas.* from lojas,Controle where lo_loja=CT_Loja"

    Set RsDados = rdoCNLoja.OpenResultset(SQL)

    If Not RsDados.EOF Then

       wRazao = RsDados("CT_Razao")
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
       wLoja = RsDados("CT_Loja")
       wNovaRazao = IIf(IsNull(RsDados("CT_NovaRazao")), "0", RsDados("CT_NovaRazao"))
    
    End If


End Function


Public Function ExtraiNumeroCupom()

      wSair = False
      WNF = 0
      wNumeroCupom = 0
      For i = 0 To 5
   '*** Desabilitado 12/2009     Retorno = Bematech_FI_NumeroCupom(wNumeroCupom)
   '*** Desabilitado 12/2009     If Retorno <> 1 Or wNumeroCupom > 0 Then
   '*** Desabilitado 12/2009         Exit For
   '*** Desabilitado 12/2009     End If
      Next i
   '*** Desabilitado 12/2009   Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      WNF = wNumeroCupom & Val(glb_ECF)
      
     If TipoPedido <> "CriaPedido" Then
        If wNumeroCupom > 0 Then
           SQL = ""
           SQL = "Update controleECF set ct_ultimocupom= " & WNF & " " _
                & "where CT_Ecf=" & Val(glb_ECF) & ""
           rdoCNLoja.Execute (SQL)
        Else
           wECFNF = 2
           MsgBox "O sistema vai emitir nota fiscal", vbExclamation, "Aviso"
           SQL = "Update ControleECF set CT_SituacaoCupomFiscal='F' where CT_Ecf=" & Val(glb_ECF) & ""
           rdoCNLoja.Execute (SQL)
           WNF = 0
           wSair = True
        End If
      Else
        If Retorno = 1 Then
            wSair = True
        End If
      End If

End Function

Public Function ExtraiSeqPedido() As Double
   Dim rdoPedido As rdoResultset
   
   SQL = ""
    SQL = "Select CT_NumPed from Controle"
        Set rdoPedido = rdoCNLoja.OpenResultset(SQL)
    If Not rdoPedido.EOF Then
        ExtraiSeqPedido = rdoPedido("CT_NumPed")
        SQL = ""
        SQL = "Update Controle set CT_NumPed=CT_NumPed + 1"
            rdoCNLoja.Execute (SQL)
    End If
   
End Function

Public Function ExtraiSeqNota() As Double

    SQL = ""
    SQL = "Select Ct_SeqNota from controle"
    Set RsDados = rdoCNLoja.OpenResultset(SQL)

    If Not RsDados.EOF Then
       WNF = RsDados("Ct_SeqNota") + 1
       SQL = "Update controle set Ct_SeqNota=" & WNF & " "
       ExtraiSeqNota = WNF
       rdoCNLoja.Execute (SQL)
    Else
       MsgBox "Erro ao atualizar sequencia de nota"
    End If

End Function


Public Function GravaNFCapa() As Boolean
       
    If Trim(wVendedorLojaVenda) = "" Or Trim(wVendedorLojaVenda) = "0" Then
         wVendedorLojaVenda = WVendedor
    End If
    If Trim(wLojaVenda) = "" Then
         wLojaVenda = wLoja
    End If
     If UCase(Glb_SeriePed) <> "S1" And UCase(Glb_SeriePed) <> "D1" Then
          If Wserie <> PegaSerieNota And Wserie <> "CT" Then
              Wserie = ""
              Glb_Nf = 0
          End If
     Else
          Wserie = UCase(Glb_SeriePed)
     End If
    On Error Resume Next
    BeginTrans
    
    SQL = ""
    SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
         & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
         & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
         & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
         & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF)" _
         & "Values (" & WnumeroPed & ", '" & Format(Wdata, "mm/dd/yyyy") & "'," & WVendedor & ", " _
         & "" & ConverteVirgula(WTotPedido) & "," & ConverteVirgula(WTotPedido) & "," & ConverteVirgula(Format(Wdescontop, "0.00")) & ", " _
         & "" & ConverteVirgula(wSubTotal) & ",'" & Trim(wLoja) & "'," & wQtdItem & ", " _
         & "'" & WTipoNota & "','" & wCondPag & "'," & Wav & "," & WCliente & ", " _
         & "" & WCODOPER & ",'" & Format(Date, "mm/dd/yyyy") & "'," & ConverteVirgula(WPGENTRA) & ", " _
         & "'" & Wlojat & "'," & WPesoBr & "," & WPesoLq & ", " _
         & "" & ConverteVirgula(WFRETECOBR) & "," & ConverteVirgula(WFRETECOBR) & ",'" & WOutraLoja & "'," & WOUTROVEND & ", " _
         & "'" & PegaSerieNota & "','" & WUF & "','" & WNOMCLI & "','" & WENDCLI & "','" & WCGCCLI & "','" & WMUNCLI & "', " _
         & "" & wPessoa & ",'" & WFone & "',0,'" & WIest & "','" & wbairro & "'," _
         & "'" & WCep & "','" & WDESCRIPAG & "','A'," & ConverteVirgula(Format(wValorTotalCodigoZero, "0.00")) & "," & ConverteVirgula(Format(wTotalNotaAlternativa, "0.00")) & "," & ConverteVirgula(Format(wTotalNotaAlternativa, "0.00")) & ",'" & wCarimbo3 & "','" & WCFOAux & "','" & wLojaVenda & "','" & wVendedorLojaVenda & "', " & wPedidoCliente & "," & Val(glb_ECF) & "," & Glb_Nf & ")"
    rdoCnLojaBach.Execute (SQL)
    If Err.Number = 0 Then
         CommitTrans
         GravaNFCapa = True
         wNfCapa = True
    Else
         Rollback
    End If
    
     
            
End Function


Public Sub GravaNfItens()
    On Error Resume Next
    'WNF = Glb_Nf
    BeginTrans
    SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
        & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
        & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
        & "Values (" & WnumeroPed & ", '" & Format(Wdata, "mm/dd/yyyy") & "','" & Trim(wReferencia) & "', " _
        & "" & wQtde & ", " & ConverteVirgula(wVlUnit) & ", " & ConverteVirgula(wVlUnit2) & ", " _
        & "" & ConverteVirgula(wVlTotItem) & "," & ConverteVirgula(WDESCRAT) & "," _
        & "" & wItem & "," & wLinha & "," & wSecao & "," & WCSPROD & ", " _
        & "" & ConverteVirgula(wPLISTA) & "," & WTRIBUTO & "," & ConverteVirgula(wIcmPdv) & ", " _
        & "'" & wCodBarra & "'," & WNF & ", '" & Wserie & "'," & WCliente & ", " _
        & "'" & WTipoNota & "'," & WVendedor & ",'" & wLoja & "'," & wTipoMovimentacao & ",'A'," & ConverteVirgula(wValorMercadoriaAlternativa) & "," & ConverteVirgula(wValorTotalItemAlternativa) & ",'" & WREFALTERNA & "','" & wPegaDescricaoAlternativa & "','" & wSerieProd1 & "' , '" & wSerieProd2 & "')"
     
     rdoCnLojaBach.Execute (SQL)
     If Err.Number = 0 Then
        CommitTrans
        wNFitens = True
     Else
        Rollback
     End If
    
End Sub

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

Public Function EmiteNotafiscalOutrasSaidas(ByVal Nota As Double, ByVal Serie As String, ByVal NatOper As String)
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
           
      CabecalhoOutrasSaidas NatOper
      
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota,NfItens.Tributacao, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
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
                 
                 If RsDados("UFCliente") = "SP" Then
                    wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
                 Else
                    wAliqICMSInterEstadual = GLB_AliquotaICMS
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
                   
                   If RsDados("UFCliente") = "SP" Then
                       wAliqICMSInterEstadual = RsdadosItens("PR_ICMSSaida")
                   Else
                       wAliqICMSInterEstadual = GLB_AliquotaICMS
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
                      
                      Printer.Print wStr16
                      
                      If RsdadosItens("DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                         Do While wConta < 28
                            wConta = wConta + 1
                            Printer.Print ""
                         Loop
                         RsdadosItens.MoveNext
                         wStr13 = Space(78) & "CX 0" & GLB_NumeroCaixa & Space(3) & "Lj " & RsDados("LojaOrigem") & Space(3) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                         Printer.Print wStr13
                         
                         wConta = 0
                         wPagina = wPagina + 1
                         Printer.EndDoc
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
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
        
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If
    
End Function



Private Sub FinalizaNota()
       If wNotaTransferencia = False Then
         If wReferenciaEspecial <> "" Then
             SQL = ""
             SQL = "Select * from CarimbosEspeciais " _
                & "where CE_Referencia='" & wReferenciaEspecial & "'"
                Set RsPegaItensEspeciais = rdoCNLoja.OpenResultset(SQL)
                
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
        End If
     End If
     Do While wConta < 7
        wConta = wConta + 1
        Printer.Print ""
     Loop

     If RsDados("Carimbo1") <> "" And RsDados("Desconto") <> 0 And Wsm = True Then
        'Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(120), 120)
        Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(106), 106) & Left("Desc." & Space(7), 7) & Left(Format((RsDados("Desconto") + RsDados("ValorTotalCodigoZero")), "0.00") & Space(10), 10)
     ElseIf RsDados("Carimbo1") <> "" And RsDados("Desconto") <> 0 Then
        Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(106), 106) & Left("Desc." & Space(7), 7) & Left(Format(RsDados("Desconto"), "0.00") & Space(10), 10)
     ElseIf RsDados("Carimbo1") <> "" And Wsm = True Then
        Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(106), 106) & Left("Desc." & Space(7), 7) & Left(Format((RsDados("Desconto") + RsDados("ValorTotalCodigoZero")), "0.00") & Space(10), 10)
     ElseIf RsDados("Carimbo1") <> "" Then
        Printer.Print Space(1) & RsDados("Carimbo1")
     ElseIf RsDados("Desconto") <> 0 And Wsm = False Then
        Printer.Print Space(91) & "Desconto" & Space(13) & Format(RsDados("Desconto"), "0.00")
     ElseIf Wsm = True Or Wserie = "SM" Then
        Printer.Print Space(91) & "Desconto" & Space(13) & Format((RsDados("Desconto") + RsDados("ValorTotalCodigoZero")), "0.00")
     Else
        Printer.Print ""
     End If
     If RsDados("Carimbo2") <> "" Then
        Printer.Print Space(4) & RsDados("Carimbo2")
        wConta = wConta + 1
     End If
     
     wConta = wConta + 1
     
     If (IIf(IsNull(RsDados("Carimbo5")), "", RsDados("Carimbo5"))) <> "" Then
        Printer.Print Space(4) & RsDados("Carimbo5")
     Else
        Printer.Print ""
     End If
        
     Do While wConta < 14
        wConta = wConta + 1
        Printer.Print ""
     Loop

     If Wsm = True Then
        'Printer.Print ""
        'Printer.Print ""
     Else
        'Print #Notafiscal, ""
        'If RsDados("Desconto") <> 0 Then
            'Print #Notafiscal, Space(114) & "Desconto" & Space(13) & Format(RsDados("Desconto"), "0.00")
            'Print #Notafiscal, ""
        'Else
            'Print #Notafiscal, ""
            'Print #Notafiscal, ""
        'End If
    End If
     If Wsm = True Then
        wStr9 = Right$(Space(9) & Format(RsDados("BaseICMS"), "######0.00"), 9) & Right$(Space(9) & Format(RsDados("VLRICMS"), "######0.00"), 9) & Space(35) & Right$(Space(10) & Format(RsDados("TotalNotaAlternativa"), "######0.00"), 10)
        'wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "'"#####0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "'"#####0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "'"#####0.00"), 12)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = Right(Space(9) & Format(Space(9) & RsDados("FreteCobr"), "######0.00"), 9) & Space(44) & Right(Space(10) & Format(RsDados("TotalNotaAlternativa"), "######0.00"), 10)
        'wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "'"#####0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "'"#####0.00"), 12)
        Printer.Print wStr10
     Else
        wStr9 = Right$(Space(9) & Format(RsDados("BaseICMS"), "######0.00"), 9) & Right$(Space(9) & Format(RsDados("VLRICMS"), "######0.00"), 9) & Space(35) & Right$(Space(10) & Format(RsDados("VlrMercadoria"), "######0.00"), 10)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = Right(Space(9) & Format(Space(9) & RsDados("FreteCobr"), "######0.00"), 9) & Space(44) & Right(Space(10) & Format(RsDados("TotalNota"), "######0.00"), 10)
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
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(78) & "CX 0" & GLB_NumeroCaixa & Space(3) & "Lj " & RsDados("LojaOrigem") & Space(4) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     'Print #Notafiscal, ""
     'printer.Print  Chr(18) 'Finaliza Impressão
     'Print #Notafiscal, Chr(27)
     
      
     'Close #Notafiscal
     'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'     FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
     Printer.EndDoc
     'wTotalNotaTransferencia = RsDados("VlrMercadoria")
     'If wReemissao = False Then
        'SQL = "Select * from CtCaixa order by CT_Data desc"
        '   Set rsPegaLoja = rdoCnLoja.OpenResultset(SQL)
        'If Not rsPegaLoja.EOF Then
        '   If WNatureza = "TRANSFERENCIAS" Or WNatureza = "TRANSFERENCIA" Then
        '       SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & rsPegaLoja("ct_operador") & "','" & rsPegaLoja("ct_loja") & "', " _
                   & " '" & Format(rsPegaLoja("ct_data"), "mm/dd/yyyy") & "', " & 20109 & "," & WNfTransferencia & ",'" & Wserie & "', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
         '          rdoCnLoja.Execute (SQL)
           'ElseIf WNatureza = "DEVOLUCAO" Then
               'SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & RsPegaLoja("ct_operador") & "','" & RsPegaLoja("ct_loja") & "', " _
                   & " '" & Format(RsPegaLoja("ct_data"), "mm/dd/yyyy") & "', " & 20201 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "'"',###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   'rdocnloja.Execute (SQL)
           'End If
        'End If
    'End If
       
End Sub


Private Sub Finaliza()

    flg = 0
    wlin = 99
    Screen.MousePointer = 0

End Sub


Public Function ExtraiLoja()

    SQL = "Select Ct_loja From controle"
    Set RsDados = rdoCNLoja.OpenResultset(SQL)
    If Not RsDados.EOF Then
       wLoja = RsDados("Ct_Loja")
    Else
       MsgBox "Problemas no controle"
       Exit Function
    End If

End Function
    
    
Public Function ExtraiDataMovimento()
    
    SQL = "Select max(Ct_Data) as WdataMax From Ctcaixa"
    Set RsDados = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsDados.EOF Then
       Wdata = RsDados("WdataMax")
       
       'Wdata = Mid(Isql("WdataMax"), 1, 2) & Mid(Isql("WdataMax"), 4, 2) & Mid(Isql("WdataMax"), 7, 2)
       'Wdate = Format(Wdata, "dd,mm,yyyy")
    Else
       MsgBox "Problemas no Ctcaixa"
       Exit Function
    End If
    
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

          
Public Function ExtraiSeqPedidoDbf()

     Dim WNovoSeqPed As Long
     
     If frmCaixa.txtPedido.Text = "" Then
     
        WnumeroPedidoDbf = 0
        WNovoSeqPed = 0
        WnumeroPed = 0
        
        Set DBFBanco = Workspaces(0).OpenDatabase(WbancoDbf, False, False, "DBase IV")

        Set RsDadosDbf = DBFBanco.OpenRecordset("Select * from controle.dbf ")
        
        WnumeroPedidoDbf = RsDadosDbf("NumPed") + 1
        WNovoSeqPed = WnumeroPedidoDbf
        WnumeroPed = WnumeroPedidoDbf
        
        BeginTrans
        
        SQL = "update controle.dbf set NumPed= " & WNovoSeqPed & ""
        DBFBanco.Execute (SQL)
        
        CommitTrans
     
        DBFBanco.Close
        
     End If

End Function



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


Function GravaSequenciaLeitura(ByVal CodigoTabela As Double, SequenciaGravacao As Double, ByVal SerieDocumento As String)
  
'  SQL = "Insert into SequenciaGravacao(SL_CodigoTabela, SL_sequenciaGravacao, SL_Serie, SL_Situacao,SL_Data) " _
'        & "Values (" & CodigoTabela & ", " & SequenciaGravacao & ", " _
'        & "'" & SerieDocumento & "', 'A','" & Format(Date, "DD/MM/YYYY") & "') "
'  rdocnloja.Execute (SQL)

End Function


Sub SelecionaMovimentoCaixa()
    'SQL = ""
    'SQL = "Select MC_Sequencia,MC_Remessa from MovimentoCaixa " _
        & "where MC_Remessa = 9 "
    'Set RsSelecionaMovCaixa = rdocnloja.OpenResultset (SQL)
    
    'If Not RsSelecionaMovCaixa.EOF Then
        'Do While Not RsSelecionaMovCaixa.EOF
            'GravaSequenciaLeitura 4, RsSelecionaMovCaixa("MC_Sequencia"), 0
            'RsSelecionaMovCaixa.MoveNext
        'Loop
    'End If
            
    'SQL = ""
    'SQL = "Update MovimentoCaixa set MC_Remessa = 0 " _
         & "Where MC_Remessa = " & 9 & " and MC_Grupo <> " & 10101 & " and MC_Grupo <> " & 30101
    'rdocnloja.Execute (SQL)
    
    'SQL = "Update MovimentoCaixa set MC_Remessa = 1 " _
         & "Where MC_Remessa = " & 9 & " and MC_Grupo = " & 10101 & " or MC_Grupo = " & 30101
    'rdocnloja.Execute (SQL)
End Sub

Sub SelecionaMovimentoBancario()
    SQL = ""
    SQL = "Select MB_Sequencia,MB_TipoMovimentacao from MovimentoBancario " _
        & "Where MB_TipoMovimentacao = 9 "
        Set RsSelecionaMovBanco = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsSelecionaMovBanco.EOF Then
        Do While Not RsSelecionaMovBanco.EOF
            GravaSequenciaLeitura 3, RsSelecionaMovBanco("MB_Sequencia"), 0
            RsSelecionaMovBanco.MoveNext
        Loop
    End If

    SQL = "Update MovimentoBancario set MB_TipoMovimentacao = 0 " _
         & "Where MB_TipoMovimentacao = " & 9
    rdoCNLoja.Execute (SQL)

End Sub

    
Sub SelecionaMovimentoEstoque()

    SQL = ""
    SQL = "Select ME_Sequencia,ME_Situacao from MovimentacaoEstoque " _
        & "Where ME_Situacao = '9' "
    Set RsSelecionaMovEstoque = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsSelecionaMovEstoque.EOF Then
        Do While Not RsSelecionaMovEstoque.EOF
            GravaSequenciaLeitura 2, RsSelecionaMovEstoque("ME_Sequencia"), 0
            RsSelecionaMovEstoque.MoveNext
        Loop
    End If
    
    SQL = "Update MovimentacaoEstoque set ME_Situacao = 0 " _
         & "Where ME_Situacao = '" & 9 & "' "
    rdoCNLoja.Execute (SQL)
    
End Sub



Sub SelecionaDivergenciaEstoque()
    SQL = ""
    SQL = "Select DE_Sequencia from DivergenciaEstoque order by DE_sequencia desc"
    Set RsSelecionaDivEstoque = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsSelecionaDivEstoque.EOF Then
        GravaSequenciaLeitura 1, RsSelecionaDivEstoque("DE_Sequencia"), 0
    End If
End Sub


Function EncerraVendaMigracao(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double) As Boolean

Dim wQuantdadeTotalItem As Integer
Dim wAnexo As String
Dim wAnexo1 As String
Dim wAnexo2 As String
Dim wQuantItensCapaNF As Integer
Dim wCFO2 As String
Dim wCFO1 As String
Dim wChaveICMS As String
Dim GLB_TotalICMSCalculado As Double
Dim GLB_ValorCalculadoICMS As Double
Dim GLB_BasedeCalculoICMS As Double
Dim GLB_AliquotaAplicadaICMS As Double
Dim GLB_AliquotaICMS As Double
Dim GLB_BaseTotalICMS As Double
Dim wCFOItem As Double
Dim wComissaoVenda As Double
Dim wSomaVenda As Double
Dim wSomaMargem As Double
Dim wQuantItensNF As Integer
Dim wDetalheImpressao As String
Dim RsCarimbo As rdoResultset
Dim wPegaCarimboNF As String
Dim wRecebeCarimboAnexo As String
Dim wConfereCodigoZero As String
Dim wECFNF As Double
Dim wPessoa As Double

Dim wSubstituicaoTributaria As Double
Dim wAnexoIten As String
Dim WAnexoAux As String

        
        EncerraVendaMigracao = True
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
        wCFOItem = 0
        wUltimoItem = 0
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wConfereCodigoZero = 0
        wECFNF = 0
        wChaveICMSItem = 0
        wSubstituicaoTributaria = 0
        wAnexoIten = ""
        WAnexoAux = ""
        wPessoa = 1
        
'
'  --------------------------------- CALCULO DO ICMS ------------------------------------------------------------------------
'
        SQL = "Select nfcapa.*, Estados.* from nfcapa, Estados " _
              & "where nfcapa.numeroped = " & NumeroDocumento & "" _
              & "And nfcapa.ufCliente = Estados.UF_Estado"
              Set RsCapaNF = rdoCNLoja.OpenResultset(SQL)
        
            
        If Not RsCapaNF.EOF Then
            wECFNF = glb_ECF
            wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
            If RsCapaNF("UFCliente") = "SP" Then
                If wPessoa = 2 Then
                    wECFNF = glb_ECF
                End If
            End If
        Else
            MsgBox "Nota não encontrada", vbInformation, "Atenção"
            Exit Function
        End If
        
        
        
        
          SQL = "Select produto.*, nfitens.* from produto,nfitens " _
              & "where nfitens.numeroped = " & NumeroDocumento & "" _
              & "and pr_referencia = nfitens.referencia order by Nfitens.Item"
              Set RsItensNF = rdoCNLoja.OpenResultset(SQL)
          
          If Not RsItensNF.EOF Then
             Do While Not RsItensNF.EOF
               
               
               'If RsCapaNf("TM") <> 1 Then
                     wChaveICMSItem = wChaveICMS
                    If RsItensNF("PR_substituicaotributaria") = "S" Then
                        wSubstituicaoTributaria = 1
                    Else
                        wSubstituicaoTributaria = 0
                    End If
                    'If RsCapaNF("UFCliente") <> "SP" Then
                        wChaveICMSItem = wChaveICMSItem & RsItensNF("pr_icmssaida") & RsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                        If AcharICMSInterEstadual(RsItensNF("PR_Referencia"), wChaveICMSItem) = False Then
                            EncerraVendaMigracao = False
                            Exit Function
                        End If
                        GLB_AliquotaAplicadaICMS = RsICMSInter("IE_icmsAplicado")
                        GLB_AliquotaICMS = RsICMSInter("IE_IcmsDestino")
                        'If RsCapaNF("TipoNota") = "T" Then
                            'wCFOItem = "522"
                        'Else
                            'wCFOItem = RsICMSInter("IE_Cfo")
                        'End If
                        
                        wAnexoIten = RsItensNF("PR_CodigoReducaoICMS")
                        If wAnexoIten <> 0 Then
                            If wAnexoIten = 1 Then
                                wAnexo1 = RsItensNF("Item") & "," & wAnexo1
                            ElseIf wAnexoIten = 2 Then
                                wAnexo2 = RsItensNF("Item") & "," & wAnexo2
                            End If
                        End If
                    
                    If wConfereCodigoZero > 0 Then
                        GLB_ValorCalculadoICMS = Format(((RsItensNF("ValorMercadoriaAlternativa") * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                        GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
                        If GLB_TotalICMSCalculado > 0 Then
                            If RsICMSInter("IE_BasedeReducao") = 0 Then
                                GLB_BasedeCalculoICMS = RsItensNF("ValorMercadoriaAlternativa")
                            Else
                                GLB_BasedeCalculoICMS = Format(RsItensNF("ValorMercadoriaAlternativa") - ((RsItensNF("ValorMercadoriaAlternativa") * RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
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




'                    If wConfereCodigoZero > 0 Then
'                        GLB_ValorCalculadoICMS = Format(((RsItensNF("ValorMercadoriaAlternativa") * GLB_AliquotaAplicadaICMS) / 100), "0.00")
'                        GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
'                        If GLB_TotalICMSCalculado > 0 Then
'                            If RsICMSInter("IE_BasedeReducao") = 0 Then
'                                GLB_BasedeCalculoICMS = RsItensNF("ValorMercadoriaAlternativa")
'                            Else
'                                GLB_BasedeCalculoICMS = Format(RsItensNF("ValorMercadoriaAlternativa") - ((RsItensNF("ValorMercadoriaAlternativa") * RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
'                            End If
'                            GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
'                        End If
'                    Else
'                        GLB_ValorCalculadoICMS = Format(((RsItensNF("VLUNIT2") * GLB_AliquotaAplicadaICMS) / 100), "0.00")
'                        GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
'                        If GLB_TotalICMSCalculado > 0 Then
'                            If RsICMSInter("IE_BasedeReducao") = 0 Then
'                                GLB_BasedeCalculoICMS = RsItensNF("VLUNIT2")
'                            Else
'                                GLB_BasedeCalculoICMS = Format(RsItensNF("VLUNIT2") - ((RsItensNF("VLUNIT2") * RsICMSInter("IE_BasedeReducao")) / 100), "0.00")
'                            End If
'                            GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
'                        End If
'                    End If
                    'WAnexoAux = ""
                    'If RsItensNF("pr_codigoreducaoicms") <> 0 Then
                    '    WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "'"0")
                    'End If
                    
                    'If wCFOItem = 5102 Or wCFOItem = 6102 Then
                    '    wCFO1 = wCFOItem
                    'ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                    '    wCFO2 = wCFOItem
                    'End If
               'Else
                    'wVerificaTM = True
               'End If
' -------------------------------------- CALCULO DA COMISSÃO DE VENDA ---------------------------------------------------
'
                wComissaoVenda = (RsItensNF("VLUNIT2") * RsItensNF("pr_percentualcomissao") / 100)

                        
'
' -------------------------------------- CALCULO DA MARGEM DE VENDA ---------------------------------------------------
'
                wSomaVenda = wSomaVenda + RsItensNF("VLUNIT2")
                wSomaMargem = wSomaMargem + (RsItensNF("VLUNIT2") - (RsItensNF("pr_precocusto1") * RsItensNF("qtde")))


'
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'
        
'        SQL = "Select nfcapa.*, Estados.* from nfcapa, Estados " _
'            & "where nfcapa.numeroped = " & NumeroDocumento & " " _
'            & "And nfcapa.serie = 'SN' " _
'            & "And nfcapa.ufCliente = Estados.UF_Estado"
'        Set RsCapaNf = rdocnloja.OpenResultset (SQL)
        
        
'        SQL = ""
'        SQL = "Select produto.*, nfitens.* from produto,nfitens " _
'              & "where nfitens.numeroped = " & NumeroDocumento & "" _
'              & "and nfitens.serie = 'SN' " _
'              & "and pr_referencia = nfitens.referencia "
'        Set RsItensNf = rdocnloja.OpenResultset (SQL)
'
'        Do While Not RsItensNf.EOF
        
        
           wQuantItensCapaNF = RsCapaNF("QtdItem")
           wQuantItensNF = RsItensNF("Item")
           wQuantdadeTotalItem = wQuantdadeTotalItem + 1
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
           
           
        SQL = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
            & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & ", " _
            & "Comissao = " & ConverteVirgula(Format(wComissaoVenda, "0.00")) & ", " _
            & "Icms = " & RsICMSInter("IE_icmsdestino") & ", " _
            & "Bcomis = " & ConverteVirgula(RsItensNF("PR_PercentualComissao")) & " , " _
            & "DetalheImpressao = '" & wDetalheImpressao & "' " _
            & " where nfitens.numeroped = " & NumeroDocumento & "" _
            & " and Referencia = '" & RsItensNF("PR_Referencia") & "'"
            rdoCNLoja.Execute (SQL)
                    
                    
            'wUltimoItem = RsItensNF("Item")
           
           
'        SQL = "UPDATE nfitens set DetalheImpressao = '" & wDetalheImpressao & "' " _
'            & " where nfitens.numeroped = " & NumeroDocumento & "" _
'            & " and Referencia = '" & RsItensNf("PR_Referencia") & "'"
'            rdocnloja.Execute (SQL)
            
        
           
           'If RsCapaNf("CODOPER") <> "522" Then
            
'                wComissaoVenda = 0
'                wComissaoVenda = (RsItensNf("VLUNIT2") * RsItensNf("pr_percentualcomissao") / 100)
'
'                wSomaVenda = wSomaVenda + RsItensNf("VLUNIT2")
'                wSomaMargem = wSomaMargem + (RsItensNf("VLUNIT2") - (RsItensNf("pr_precocusto1") * RsItensNf("qtde")))
'
'                SQL = ""
'                SQL = "Update nfitens set Comissao = " & ConverteVirgula(Format(wComissaoVenda, "'0.00")) & " " _
'                    & " where nfitens.numeroped = " & NumeroDocumento & "" _
'                    & " and Referencia = '" & RsItensNf("PR_Referencia") & "'"
'                rdocnloja.Execute (SQL)
        
' -------------------------------------- ATUALIZA MARGEM DE VENDA ---------------------------------------------------

                SQL = "UPDATE vende set VE_totalvenda = VE_TotalVenda + " & ConverteVirgula(wSomaVenda) & ", " _
                    & "VE_MargemVenda = VE_MargemVenda + " & ConverteVirgula(wSomaMargem) & " " _
                    & "where VE_Codigo = " & RsCapaNF("Vendedor") & " "
                rdoCNLoja.Execute (SQL)
           
           'else
           'End If
            
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
'
                SQL = "Select CA_Descricao from CarimboNotaFiscal where CA_CodigoCarimbo = 1"
                Set RsCarimbo = rdoCNLoja.OpenResultset(SQL)
                
                If Not RsCarimbo.EOF Then
                    If wAnexo1 <> "" Or wAnexo2 <> "" Then
                        wPegaCarimboNF = RsCarimbo("CA_Descricao")
                        wRecebeCarimboAnexo = Mid(wPegaCarimboNF, 1, 79) & wAnexo1 & Mid(wPegaCarimboNF, 80, Len(wPegaCarimboNF)) & wAnexo2
                    Else
                        wRecebeCarimboAnexo = ""
                    End If
                End If
                
             
                'wUltimoItem = ((wUltimoItem / 10) + 0.9)
             
                SQL = "UPDATE nfcapa set baseicms = " & ConverteVirgula(GLB_BaseTotalICMS) & ", " _
                    & "Vlricms = " & ConverteVirgula(GLB_TotalICMSCalculado) & ", " _
                    & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", " _
                    & "Pessoacli = " & wPessoa & ", " _
                    & "ECFNF = " & wECFNF & ", " _
                    & "Regiaocli = " & RsCapaNF("UF_Regiao") & ", " _
                    & "Carimbo1 = '" & wRecebeCarimboAnexo & "' " _
                    & "where nfcapa.numeroped = " & NumeroDocumento & ""
                    rdoCNLoja.Execute (SQL)
                
                
'                SQL = "UPDATE nfcapa set " _
'                    & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", " _
'                    & "Carimbo1 = '" & wRecebeCarimboAnexo & "' " _
'                    & "where nfcapa.numeroped = " & NumeroDocumento & ""
'                rdocnloja.Execute (SQL)
           
           'End If
        

' -------------------------------------- ATUALIZA CONTROLE DE OPERAÇÂO ---------------------------------------------------
         

           SQL = "UPDATE CTcaixa set Ct_operacoes = Ct_operacoes + 1 " _
               & "where ct_situacao = 'A' "
           rdoCNLoja.Execute (SQL)
          
           RsItensNF.MoveNext
      Loop
    End If

End Function


Public Sub LeituraZ()

'*** Desabilitado 12/2009    Retorno = Bematech_FI_ReducaoZ("", "")
'*** Desabilitado 12/2009    Call VerificaRetornoImpressora("", "", "Redução Z")
'*** Desabilitado 12/2009    If Retorno = 1 Then
'*** Desabilitado 12/2009        Call AtualizaNumeroCupom
'*** Desabilitado 12/2009    End If
    
End Sub




Public Sub EmiteCodigoZero()
    Dim wEndCliente As String
    Dim wCgcCliente As String
    Dim RsVende As rdoResultset
    Dim RsSerieProduto As rdoResultset
    
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
    SQL = ""
    SQL = "Select NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.CLIENTE,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI," _
        & "NFCAPA.UFCLIENTE,NFCapa.Vendedor,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT2,NFITENS.VLUNIT,NFITENS.DescricaoAlternativa," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NFITENS.SERIEPROD1,NFITENS.SERIEPROD2 " _
        & "From NFCAPA INNER JOIN NFITENS " _
        & "on (NfCapa.nf=Nfitens.nf)  " _
        & "Where NfCapa.nf= " & WNF & " and NfCapa.Serie = '" & Wserie & "' and NfItens.Serie=NfCapa.Serie " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
    Set RsDados = rdoCNLoja.OpenResultset(SQL)


    If Not RsDados.EOF Then
        SQL = ""
        SQL = "Select * From Vende Where VE_Codigo = " & RsDados("Vendedor") & ""
        Set RsVende = rdoCNLoja.OpenResultset(SQL)
        If Not RsVende.EOF Then
            WVendedor = RsDados("Vendedor") & " - " & RsVende("VE_Nome")
        Else
            WVendedor = RsDados("Vendedor")
        End If
        wTotalPed = Format(RsDados("TotalNota"), "0.00")
        wDesconto = 0
        wDesconto = RsDados("Desconto")
        wCodigo = RsDados("Cliente")
        WNomeCliente = RsDados("NomCli")
        wEndCliente = RsDados("EndCli")
        wCgcCliente = RsDados("CGCCli")
        Printer.Print
        Call CabecalhoCodigoZero
        Do While Not RsDados.EOF
            wDescricao = ""
            wPegaDescricaoAlternativa = "0"
            wPegaDescricaoAlternativa = IIf(IsNull(RsDados("DescricaoAlternativa")), "0", RsDados("DescricaoAlternativa"))
            
            SQL = ""
            SQL = "Select PR_Descricao from Produto where PR_Referencia = '" & RsDados("Referencia") & "'"
                Set RsDescProduto = rdoCNLoja.OpenResultset(SQL)
            If Not RsDescProduto.EOF Then
                If wPegaDescricaoAlternativa <> "0" Then
                    wDescricao = wPegaDescricaoAlternativa
                Else
                    wDescricao = RsDescProduto("PR_Descricao")
                End If
                Printer.Print "" & RsDados("Referencia"); Tab(10); " " & wDescricao; ""
                Printer.Print "      " & RsDados("QTDE") & "x" & Format(RsDados("VLUNIT"), "0.00"); Tab(39); " " & Format(RsDados("Qtde") * RsDados("VLUNIT"), "0.00")
            End If
            RsDados.MoveNext
        Loop
        Printer.Print "  Desconto    R$                       " & Format(wDesconto, "0.00")
        Printer.Print "  TOTAL       R$                       " & Format(wTotalPed, "0.00")
'        Printer.Print "  Valor Recebido R$                    " & Format(wTotalPed, "0.00")
'        Printer.Print "  Troco   R$                           0,00"
        Printer.Print "________________________________________________"
        Printer.Print
        Printer.Print "Cliente : " & wCodigo & "  " & WNomeCliente
        Printer.Print "Endereço : " & wEndCliente
        Printer.Print "CNPJ : " & wCgcCliente
        Printer.Print
        Printer.Print "* Nº Serie do(s) Produto(s) *"
        Printer.Print "CODIGO    SERIE "
        SQL = ""
        SQL = "Select Referencia, SerieProd1, SerieProd2 From Nfitens " _
            & "Where NF = " & WNF & " and Serie = '" & Wserie & "' and LojaOrigem = '" & Trim(wLoja) & "'"
        Set RsSerieProduto = rdoCNLoja.OpenResultset(SQL)
        Do While Not RsSerieProduto.EOF
            If Not IsNull(RsSerieProduto("SerieProd1")) Or RsSerieProduto("SerieProd1") <> "0" Then
                Printer.Print RsSerieProduto("Referencia") & " - " & RsSerieProduto("SerieProd1")
            End If
            If Not IsNull(RsSerieProduto("SerieProd2")) Or RsSerieProduto("SerieProd2") <> "0" Then
                Printer.Print RsSerieProduto("Referencia") & " - " & RsSerieProduto("SerieProd2")
            End If
            RsSerieProduto.MoveNext
        Loop
        RsSerieProduto.Close
        Printer.Print
        Printer.Print "Vendedor  " & WVendedor
        Printer.Print
        'If Trim(wLoja) <> "800" Then
        '    Printer.Print "DE MEO a mais de 106 anos vendendo qualidade"
        'Else
        '    Printer.Print "DM Motores sempre o melhor preço "
        'End If
        Printer.Print "------------------------------------------------"
        Printer.Print "         COTACAO - SEM VALOR FISCAL"
        Printer.Print "------------------------------------------------"
        
        
        Printer.EndDoc
    End If
    Screen.MousePointer = 0


End Sub



Sub CabecalhoCodigoZero()
    SQL = ""
    SQL = "Select Lojas.*,CT_Loja,CT_SeqC0,CT_Razao from Lojas,Controle where LO_Loja = CT_Loja"
        Set rsPegaLoja = rdoCNLoja.OpenResultset(SQL)
    If Not rsPegaLoja.EOF Then
        
        
        Printer.ScaleMode = vbMillimeters
        Printer.ForeColor = "0"
        Printer.FontSize = 6.5
        Printer.FontName = "draft 10cpi"
        Printer.FontSize = 6.5
        Printer.FontBold = False
        Printer.DrawWidth = 3
        
        Printer.Print "      " & rsPegaLoja("CT_Razao")
        Printer.Print "      CNPJ " & rsPegaLoja("LO_CGC") & "  IE " & rsPegaLoja("LO_InscricaoEstadual")
        Printer.Print "      " & rsPegaLoja("LO_Endereco")
        Printer.Print "      Telefone : (" & rsPegaLoja("LO_DDD") & ")" & rsPegaLoja("LO_Telefone")
        Printer.Print "      " & Format(Date, "DD/MM/YYYY") & "   " & Format(Time, "hh:mm") & "       NUMERO: " & wPegaSequenciaCO
        Printer.Print
        Printer.Print "================================================"
        Printer.Print " CÓDIGO                           DESCRIÇÃO"
        Printer.Print "   QTDxUNITARIO                       VALOR(R$)"
        Printer.Print "________________________________________________"
    End If
        
        
End Sub

Public Sub ExtraiSequenciaCodigoZero()
    SQL = ""
    SQL = "Select Lojas.*,CT_Loja,CT_SeqC0 from Lojas,Controle where LO_Loja = CT_Loja"
        Set rsPegaLoja = rdoCNLoja.OpenResultset(SQL)
    If Not rsPegaLoja.EOF Then
        wPegaSequenciaCO = Val(rsPegaLoja("CT_SeqC0") + 1)
        SQL = ""
        SQL = "Update Controle set CT_SeqC0 = " & wPegaSequenciaCO
            rdoCNLoja.Execute (SQL)
    End If
End Sub


Public Sub EmiteNotaFiscalSM()
'    For Each NomeImpressora In Printers
'        If Trim(NomeImpressora.DeviceName) = "NOTA FISCAL" Then
'            ' Seta impressora no sistema
'            Set Printer = NomeImpressora
'            Exit For
'        End If
'    Next
'
'
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
'        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NfCapa.TotalNotaAlternativa,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
'        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.DataPag,NfCapa.CondPag," _
'        & "NFCAPA.UFCLIENTE,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
'        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
'        & "From NFCAPA INNER JOIN NFITENS " _
'        & "on (NfCapa.nf=Nfitens.nf) " _
'        & "Where NfCapa.nf= " & WNF & " " _
'        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
'
'    Set RsDados = rdoCnLoja.OpenResultset(SQL)
'
'    If Not RsDados.EOF Then
'        If RsDados("CondPag") = 85 Then
'            wCarimbo4 = RsDados("DataPag")
'        Else
'            wCarimbo4 = RsDados("Carimbo4")
'        End If
'
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
'           wStr20 = "Entrada       : " & Wentrada
'        End If
'
'        wStr1 = Space(2) & Left$(Format(wStr17) & Space(50), 50) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(2) & "X" & Space(31) & Left$(Format(RsDados("nf"), "'"',##'"), 7)
'        wStr2 = Space(2) & Left$(Format(wStr18) & Space(50), 50) & Left$(Format(Trim(WMunicipio), ">") & Space(15), 15) & Space(29) & Left$(Trim(westado), 2)
'        wStr3 = Space(2) & Left$(Format(wStr19) & Space(50), 50) & "(011)" & Left$(Trim(Format(WFone, "'"'-###'")), 8) & "/(011)" & Left$(Format(WFone, "'"'-###'"), 8) & Space(11) & Left$(Format((WCep), "'"##-##'"), 8)
'        wStr4 = Space(2) & Left$(Format(wStr20) & Space(100), 100) & Left$(Trim(Format(WCGC, "'"',###,##'")), 10) & "/" & Format(Mid((WCGC), 11, 5), "'"#-#'")
'        wStr5 = Space(40) & Trim(WNatureza) & Space(15) & Left$(RsDados("CFOAUX"), 10) & Space(40) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'        wStr6 = Space(40) & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(RsDados("CGCCLI"), "'"',###,##'")), 10) & "/" & Right$(Format(RsDados("CGCCLI"), "'"#-#'"), 7) & Space(5) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr7 = Space(40) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(12) & Left$(RsDados("CEPCLI") & Space(16), 16) & Space(4) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr8 = Space(40) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(43) & Left$(Trim(RsDados("UFCLIENTE")), 9) & Space(14) & Left$(Trim(Format(RsDados("INSCRICLI"), "'"',###,###,##'")), 15)
'
'
''        wStr6 = Space(40) & Left$(Format(Trim(rdorsExtra2("em_descricao")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(rdorsExtra2("lo_cgc"), "'"',###,##'")), 10) & "/" & Right$(Format(rdorsExtra2("lo_cgc"), "'"#-#'"), 7) & Space(5) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr7 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(15), 15) & Space(32) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr8 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(43) & Left$(Trim(rdorsExtra2("lo_uf")), 9) & Space(14) & Left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "'"',###,###,##'")), 15)
'
'        wStr9 = Space(4) & Right$(Space(12) & Format(RsDados("BaseICMS"), "'"#####0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "'"#####0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "'"#####0.00"), 12)
'        wStr10 = Space(67) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "'"#####0.00"), 12)
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
'        Call Cabecalho
'
'          SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
'              & "produto.pr_classefiscal,produto.pr_unidade, " _
'              & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.ReferenciaAlternativa, " _
'              & "nfitens.vlunit,nfitens.vltotitem,NfItens.ValorMercadoriaAlternativa,NfItens.PrecoUnitAlternativa,nfitens.icms,nfitens.detalheImpressao " _
'              & "from produto,nfitens " _
'              & "where produto.pr_referencia=nfitens.referencia " _
'              & "and nfitens.nf = " & WNF & ""
'
'          Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
'
'          If Not RsdadosItens.EOF Then
'             Do While Not RsdadosItens.EOF
'                      wPegaDescricaoAlternativa = IIf(IsNull(RsDados("Referencia")), "0", RsDados("Referencia"))
'                    SQL = ""
'                    SQL = "Select Desc from EvDesDBF where NotaFis = " & WNF & " " _
'                        & "And Ref = '" & wPegaDescricaoAlternativa & "'"
'                        Set RsPegaDescricaoAlternativa = rdoCnLoja.OpenResultset(SQL)
'                    wPegaDescricaoAlternativa = ""
'                    If Not RsPegaDescricaoAlternativa.EOF Then
'                        wPegaDescricaoAlternativa = IIf(IsNull(RsPegaDescricaoAlternativa("Desc")), 0, RsPegaDescricaoAlternativa("Desc"))
'                    End If
'
'                      wStr16 = ""
'                      wStr16 = Space(6) & Left$(RsdadosItens("ReferenciaAlternativa") & Space(8), 8) _
'                             & Space(2) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
'                             & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
'                             & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
'                             & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
'                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "'"##0"), 6) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "'"#####0.00"), 12) & Space(2) _
'                             & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "'"#####0.00"), 15) & Space(2) _
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
'                             Call Cabecalho
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

End Sub


Public Sub Discador()


    Wconectou = False
    
    On Error Resume Next
   

    Set RdoDados = Conexao.OpenResultset("Select LO_Loja from Loja where Lo_loja='315'", Options:=rdExecDirect)
    
    
    If Err.Number = 40071 Then
       rtn = Shell("rundll32.exe rnaui.dll,RnaDial " & "VicNet", 1)
       'HandlerWindow = FindWindow("'32770", "Conectar a")
       SendKeys "{ENTER}", True
       'SendKeys "{ENTER}", True
       Wconectou = False
       Esperar 30
    Else
       Wconectou = True
    End If
    
    
   
   
   Err.Clear
   rdoErrors.Clear
   
   i = 0
   
   For i = 1 To 10

       If Wconectou = False Then
          
          ConectaODBC Conexao, "sa", ""
          If Wconectou = True Then
             Exit For
          End If
       
       Else
          Exit For
       End If

   Next
   
         If Wconectou = False Then
            MsgBox i & "Conexão Falhou"
         End If


End Sub


Function ConectaODBC(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
    
        
        'If i = 1 Then
           DescricaoOperacao "Conectando ao banco central"
           On Error GoTo ConexaoErro
        
        'End If
    
        With RdoVar
            Servidor = GLB_Servidor
            WBANCO = GLB_Banco
    
            .Connect = "Driver={SQL Server};" _
                    & "Server=" & Trim(Servidor) & ";" _
                    & "DataBase=" & Trim(WBANCO) & ";" _
                    & "MaxBufferSize=512;" _
                    & "PageTimeout=5;" _
                    & "UID=" & Usuario & ";" _
                    & "PWD=" & Senha & ";"
    
            .LoginTimeout = 10
            .CursorDriver = rdUseClientBatch
            .EstablishConnection rdDriverNoPrompt
        End With
    
        ConectaODBC = True
        Wconectou = True
        DescricaoOperacao "Pronto"
        Exit Function
            

ConexaoErro:

    ConectaODBC = False
    Wconectou = False
    DescricaoOperacao "Pronto"

End Function


Function ConectaOdbcLocal(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
    
        
        'If i = 1 Then
           
           On Error GoTo ConexaoErro
        
        'End If
    
        With RdoVar
            Servidor = Glb_ServidorLocal
            WBANCO = Glb_BancoLocal
    
            .Connect = "Dsn=" & Trim(Servidor) & ";" _
                    & "Server=" & Trim(Servidor) & ";" _
                    & "DataBase=" & Trim(WBANCO) & ";" _
                    & "MaxBufferSize=512;" _
                    & "PageTimeout=5;" _
                    & "UID=" & Usuario & ";" _
                    & "PWD=" & Senha & ";"
    
            .LoginTimeout = 10
            .CursorDriver = rdUseClientBatch
            .EstablishConnection rdDriverNoPrompt
        End With
    
        ConectaOdbcLocal = True
        Wconectou = True
        Exit Function
    
ConexaoErro:

    ConectaOdbcLocal = False
    Wconectou = False

End Function




Public Function CriaArquivoNF()
    wNotaTransferencia = False
    wPagina = 1
    WNatureza = "VENDAS"
    Temporario = "C:\NOTASVB\"
            
    Call DadosLoja
            
    SQL = ""
    SQL = "Select NFCAPA.FreteCobr,NFCAPA.Carimbo5,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV,NFCAPA.Carimbo3,NFCAPA.Carimbo2,NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF, " _
        & "NFCAPA.CLIENTE,NFCAPA.FONECLI,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
        & "NFCAPA.UFCLIENTE,NFCAPA.TOTALNOTAALTERNATIVA,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
        & "From NFCAPA INNER JOIN NFITENS " _
        & "on (NfCapa.nf=Nfitens.nf) " _
        & "Where NfCapa.nf= " & WNF & " and NfCapa.Serie='" & Wserie & "' and NfItens.Serie=NfCapa.Serie " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
        
    Set RsDados = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsDados.EOF Then
      
      If Glb_NfDevolucao = True And RsDados("Serie") = "SM" Then
            Wsm = True
      End If
      
      Call CabecalhoArq
            
      If Err.Number <> 0 Then
         Exit Function
      End If
      
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,NfItens.IcmPdv,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & WNF & " and NfItens.Serie = '" & Wserie & "' order by nfitens.item"

      Set RsdadosItens = rdoCNLoja.OpenResultset(SQL)

      If Not RsdadosItens.EOF Then
         wConta = 0
         Do While Not RsdadosItens.EOF
            wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = RsdadosItens("PR_Referencia")
            If Wsm = True Then
                 wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), "0", RsdadosItens("DescricaoAlternativa"))
                   
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("ReferenciaAlternativa") & Space(8), 8) _
                          & Space(2) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
                          & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                          & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
                          & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                          & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) & Space(2) _
                          & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "0.00"), 12) & Space(1) _
                          & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "0.00"), 15) & Space(1) _
                          & Right$(Space(2) & Format(RsdadosItens("IcmPdv"), "0"), 2)
            
            Else
                     
                   wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), "0", RsdadosItens("DescricaoAlternativa"))
                   If Trim(wPegaDescricaoAlternativa) = "" Then
                        wPegaDescricaoAlternativa = "0"
                   End If
                   If wPegaDescricaoAlternativa <> "0" Then
                         wDescricao = wPegaDescricaoAlternativa
                   Else
                         wDescricao = Trim(RsdadosItens("pr_descricao"))
                   End If
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("pr_referencia") & Space(8), 8) _
                         & Space(2) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
                         & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                         & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
                         & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                         & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) & Space(2) _
                         & Right$(Space(12) & Format(RsdadosItens("vlunit"), "0.00"), 12) & Space(1) _
                         & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "0.00"), 15) & Space(1) _
                         & Right$(Space(2) & Format(RsdadosItens("IcmPdv"), "0"), 2)

                                  
            End If
                      
                      On Error Resume Next
                      Print #NotaFiscal, wStr16
                      'If Err.Number = 52 Then
                        'Close #Notafiscal
                        'Print #Notafiscal, wStr16
                      'End If
                        
                      
                      If RsdadosItens("DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                         Do While wConta < 21
                            wConta = wConta + 1
                            Print #NotaFiscal, ""
                         Loop
                         RsdadosItens.MoveNext
                         wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "0"), 7)
                         Print #NotaFiscal, wStr13
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, Chr(18) 'Finaliza Impressão
                         Close #NotaFiscal
                         wConta = 0
                         wPagina = wPagina + 1
                         FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'                         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
                         Call CabecalhoArq
                      ElseIf RsdadosItens("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                         Call FinalizaArqNf
                      Else
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      End If
                      
            Loop
         Else
            Close #NotaFiscal
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
        
         'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
         'FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If
End Function
              
Private Sub CabecalhoArq()
        
        Dim wCgcCliente As String
        
        NomeArquivo = "nf" & Trim(RsDados("NF")) & wPagina & ".txt"
        
        NotaFiscal = FreeFile
        On Error Resume Next
        Open Temporario & NomeArquivo For Output Access Write As #NotaFiscal
        If Err.Number = 55 Then
            MsgBox "Nota não pode ser impressa," & Chr(10) & "Existe um erro com a quantidades de itens desta nota", vbCritical, "Aviso"
            Exit Sub
        End If
        Wcondicao = "            "
        Wav = "          "
        If RsDados("CondPag") = 85 Then
            wCarimbo4 = Format(RsDados("DataPag"), "mm/dd/yyyy")
        Else
            wCarimbo4 = IIf(IsNull(RsDados("Carimbo4")), "", RsDados("Carimbo4"))
        
        End If
        wLojaVenda = "            "
        wVendedorLojaVenda = "            "
        wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
        wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))
        Wentrada = 0
        Wcondicao = "            "
        wStr20 = ""
        wStr19 = "               "
        wStr7 = "               "
        If Val(RsDados("CONDPAG")) = 1 Then
           Wcondicao = "Avista"
        ElseIf Val(RsDados("CONDPAG")) = 3 Then
           Wcondicao = "Financiada"
        ElseIf Val(RsDados("CONDPAG")) > 3 Then
           Wcondicao = wCarimbo4
        End If
        
        
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
            If Mid(Wcondicao, 1, 9) = "Faturada " Then
                Wav = "AV            : " & Trim(RsDados("AV"))
            End If
        End If
        If Glb_NfDevolucao = True Then
            WNatureza = "DEVOLUCAO"
        End If
        
        If Trim(WNatureza) = "TRANSFERENCIAS" Then
            Wcondicao = "            "
        ElseIf Trim(WNatureza) = "DEVOLUCAO" Then
            Wcondicao = "            "
        End If
        
        wStr17 = "Pedido        : " & RsDados("NUMEROPED")
        wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
        If Trim(Wcondicao) <> "" Then
            wStr19 = "Cond. Pagto   : " & Trim(Wcondicao)
        ElseIf Trim(RsDados("Carimbo3")) <> "" Then
            wStr19 = Trim(RsDados("Carimbo3"))
        Else
            Wcondicao = "            "
        End If

        If RsDados("Pgentra") <> 0 Then
           Wentrada = Format(RsDados("Pgentra"), "0.00")
           wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
        End If
        If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
            wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
        End If
        
        WCGC = Right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "###,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        
        wStr1 = Space(120) & wPagina & "/" & RsDados("PAGINANF") 'Inicio Impressão
        Print #NotaFiscal, Chr(15) & wStr1
        'Print #Notafiscal, wStr1
        If Trim(WNatureza) = "DEVOLUCAO" Then
            wStr1 = Space(2) & Left$(Format(wStr17) & Space(40), 40) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(16) & "X" & Space(16) & Left$(Format(RsDados("nf"), "###,###"), 7)
        Else
            wStr1 = Space(2) & Left$(Format(wStr17) & Space(40), 40) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(5) & "X" & Space(30) & Left$(Format(RsDados("nf"), "###,###"), 7)
        End If
        Print #NotaFiscal, wStr1
        wStr2 = Space(2) & Left$(Format(wStr18) & Space(40), 40) & Left$(Format(Trim(WMunicipio), ">") & Space(15), 15) & Space(29) & Left$(Trim(westado), 2)
        Print #NotaFiscal, wStr2
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-###")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFone, "###-###"), 9) & Space(11) & Left$(Format((WCep), "###-###"), 9)
        Print #NotaFiscal, wStr3
        wStr4 = Space(2) & Left$(Format(wStr20) & Space(100), 100) & Left$(Trim(Format(WCGC, "###,###,###")), 19) '& Format(Mid((WCGC), 11, 5), "###-###")
        Print #NotaFiscal, wStr4
        'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
        If Trim(WNatureza) = "TRANSFERENCIAS" Then
            wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        ElseIf Trim(Wav) <> "" Then
            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
        Print #NotaFiscal, wStr5
        'Print #Notafiscal, ""
        Print #NotaFiscal, ""
        If Mid(RsDados("CLIENTE"), 1, 5) <> "99999" Then
            wCgcCliente = Right(String(14, "0") & Trim(RsDados("CGCCLI")), 14)
            wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        Else
            wCgcCliente = "00.000.000/0000-00"
        End If
        If wStr6 <> "" Then
            wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Trim(wCgcCliente), 19) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr6 = Space(34) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(11) & Left$(Trim(wCgcCliente), 19) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        End If
        Print #NotaFiscal, wStr6
        Print #NotaFiscal, ""
        wStr7 = Space(2) & Left(wStr7 & Space(32), 32) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(19) & Left$(RsDados("CEPCLI"), 11) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Print #NotaFiscal, wStr7
        wStr8 = Space(34) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Space(8) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
        Print #NotaFiscal, wStr8
        
        Print #NotaFiscal, ""
        Print #NotaFiscal, ""
           
End Sub
              
  
Private Sub FinalizaArqNf()
     If wNotaTransferencia = False Then
         If wReferenciaEspecial <> "" Then
             SQL = ""
             SQL = "Select * from CarimbosEspeciais " _
                & "where CE_Referencia='" & wReferenciaEspecial & "'"
                Set RsPegaItensEspeciais = rdoCNLoja.OpenResultset(SQL)
                
             If Not RsPegaItensEspeciais.EOF Then
                i = 0
        
                If RsPegaItensEspeciais("CE_Linha1") <> "" Then
                    wConta = wConta + 7
                    'Print #Notafiscal, ""
                    If Trim(RsPegaItensEspeciais("CE_Linha5")) = "" Then
                        Print #NotaFiscal, Space(15) & "______________________________________________________________"
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Print #NotaFiscal, Space(17) & "___________________________________     ____/____/______   "
                        Print #NotaFiscal, Space(17) & "            Assinatura                        Data         "
                        'Print #Notafiscal, Space(15) & "____________________________________________________________"
                    Else
                        Print #NotaFiscal, Space(15) & "______________________________________________________________"
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Print #NotaFiscal, Space(16) & Right(RsPegaItensEspeciais("CE_Linha5"), 60)
                        Print #NotaFiscal, Space(17) & "___________________________________     ____/____/______   "
                        Print #NotaFiscal, Space(17) & "            Assinatura                        Data         "
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
        End If
     End If
     Do While wConta < 7
        wConta = wConta + 1
        Print #NotaFiscal, ""
     Loop

     If RsDados("Carimbo1") <> "" And RsDados("Desconto") <> 0 And Wsm = True Then
        Print #NotaFiscal, Space(1) & Left(RsDados("Carimbo1") & Space(120), 120)
     ElseIf RsDados("Carimbo1") <> "" And RsDados("Desconto") <> 0 Then
        Print #NotaFiscal, Space(1) & Left(RsDados("Carimbo1") & Space(119), 119) & Left("Desconto" & Space(12), 12) & Left(Format(RsDados("Desconto"), "0.00") & Space(10), 10)
     ElseIf RsDados("Carimbo1") <> "" Then
        Print #NotaFiscal, Space(1) & RsDados("Carimbo1")
     ElseIf RsDados("Desconto") <> 0 And Wsm = False Then
        Print #NotaFiscal, Space(111) & "Desconto" & Space(13) & Format(RsDados("Desconto"), "0.00")
     Else
        Print #NotaFiscal, ""
     End If
     If RsDados("Carimbo2") <> "" Then
        Print #NotaFiscal, Space(4) & RsDados("Carimbo2")
     End If
     
     wConta = wConta + 1
     
     If (IIf(IsNull(RsDados("Carimbo5")), "", RsDados("Carimbo5"))) <> "" Then
        Print #NotaFiscal, Space(4) & RsDados("Carimbo5")
        wConta = wConta + 1
     Else
        Print #NotaFiscal, ""
        wConta = wConta + 1
     End If
        
     Do While wConta < 10
        wConta = wConta + 1
        Print #NotaFiscal, ""
     Loop

     If Wsm = True Then
        Print #NotaFiscal, ""
        Print #NotaFiscal, ""
     Else
        'Print #Notafiscal, ""
        'If RsDados("Desconto") <> 0 Then
            'Print #Notafiscal, Space(114) & "Desconto" & Space(13) & Format(RsDados("Desconto"), "0.00")
            'Print #Notafiscal, ""
        'Else
            'Print #Notafiscal, ""
            'Print #Notafiscal, ""
        'End If
    End If
     If Wsm = True Then
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "######0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "######0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "######0.00"), 12)
        Print #NotaFiscal, wStr9
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "######0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "######0.00"), 12)
        Print #NotaFiscal, wStr10
     Else
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "######0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "######0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("VlrMercadoria"), "######0.00"), 12)
        Print #NotaFiscal, wStr9
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "######0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("VlrMercadoria"), "######0.00"), 12)
        Print #NotaFiscal, wStr10
     End If
     
     wStr11 = Space(2) & "                          "
     Print #NotaFiscal, wStr11
     wStr12 = Space(2) & "                                                     "
     Print #NotaFiscal, wStr12
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     wStr13 = Space(87) & "CX 0" & GLB_NumeroCaixa & "  Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
     Print #NotaFiscal, wStr13
     Print #NotaFiscal, ""
     Print #NotaFiscal, ""
     'Print #Notafiscal, ""
     Print #NotaFiscal, Chr(18) 'Finaliza Impressão
     Print #NotaFiscal, Chr(27)
     
      
     Close #NotaFiscal
     FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
     'FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
     wTotalNotaTransferencia = RsDados("VlrMercadoria")
     If wReemissao = False Then
        SQL = "Select * from CtCaixa order by CT_Data desc"
           Set rsPegaLoja = rdoCNLoja.OpenResultset(SQL)
        If Not rsPegaLoja.EOF Then
           If WNatureza = "TRANSFERENCIAS" Then
               SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(" & Val(glb_ECF) & ",'" & rsPegaLoja("ct_operador") & "','" & rsPegaLoja("ct_loja") & "', " _
                   & " '" & Format(rsPegaLoja("ct_data"), "mm/dd/yyyy") & "', " & 20109 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   rdoCNLoja.Execute (SQL)
           'ElseIf WNatureza = "DEVOLUCAO" Then
               'SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & RsPegaLoja("ct_operador") & "','" & RsPegaLoja("ct_loja") & "', " _
                   & " '" & Format(RsPegaLoja("ct_data"), "mm/dd/yyyy") & "', " & 20201 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "'"',###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   'rdocnloja.Execute (SQL)
           End If
        End If
    End If
End Sub
           

Function Cabecalho(ByVal TipoNota As String)
    Dim wCgcCliente As String
    Dim impri As Long
    Dim rdoConPag As rdoResultset
    impri = Printer.Orientation
    'Printer.PrintQuality = vbPRPQDraft
    
    
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
    If RsDados("CondPag") = 85 Then
        wCarimbo4 = Format(RsDados("DataPag"), "dd/mm/yyyy")
    Else
        SQL = ""
        SQL = "Select CP_Descricao from CondicaoPagto " _
            & "where CP_VendaCompra='V' and CP_CodigoCondicao=" & RsDados("CondPag") & ""
        Set rdoConPag = rdoCNLoja.OpenResultset(SQL)
        If Not rdoConPag.EOF Then
            wCarimbo4 = rdoConPag("CP_Descricao")
        End If
    End If
    wLojaVenda = "            "
    wVendedorLojaVenda = "            "
    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))
    Wentrada = 0
    Wcondicao = "            "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = "               "
    If Val(RsDados("CONDPAG")) = 1 Then
       Wcondicao = "Avista"
    ElseIf Val(RsDados("CONDPAG")) = 3 Then
       Wcondicao = "Financiada"
    ElseIf Val(RsDados("CONDPAG")) > 3 Then
       
       Wcondicao = wCarimbo4
    End If
    
    If UCase(TipoNota) = "T" Then
        WNatureza = "TRANSFERENCIA"
    ElseIf UCase(TipoNota) = "V" Then
        WNatureza = "VENDA"
    ElseIf UCase(TipoNota) = "E" Then
        WNatureza = "DEVOLUCAO"
    ElseIf UCase(TipoNota) = "S" And (RsDados("CFOAUX") = "5949" Or RsDados("CFOAUX") = "6949") Then
        WNatureza = "OUTRAS OPER Ñ ESPEC."
    End If
    
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
        If Mid(Wcondicao, 1, 9) = "Faturada " Then
            Wav = "AV            : " & Trim(RsDados("AV"))
        End If
    End If
    
    If Trim(WNatureza) = "TRANSFERENCIA" Then
        Wcondicao = "            "
    ElseIf Trim(WNatureza) = "DEVOLUCAO" Then
        Wcondicao = "            "
    End If
    
    wStr17 = "Pedido        : " & RsDados("NUMEROPED")
    wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
    If Trim(Wcondicao) <> "" Then
        wStr19 = "Cond Pagto : " & Trim(Wcondicao)
    ElseIf Trim(RsDados("Carimbo3")) <> "" Then
        wStr19 = "Transporte    : " & Left(Format(Trim(RsDados("Carimbo3"))) & Space(10), 10)
    Else
        Wcondicao = "            "
    End If

    If RsDados("Pgentra") <> 0 Then
       Wentrada = Format(RsDados("Pgentra"), "#####0.00")
       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
    End If
    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
        wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
    End If
    
    
    
    'wLinha2 = Space(2) & Left(rsDadosCliente("CLI_RazaoSocial") & Space(100), 100) _
            & Left(rsDadosCliente("CLI_Cnpj") & Space(18), 18) _
            & Left(Format(Date, "dd/mm/yyyy") & Space(10), 10)
    
    'Printer.FontSize = 8
    If wPagina = 1 Then
        WCGC = Right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "###,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        WCGC = Right(String(18, "0") & WCGC, 18)
    End If
    wStr0 = Space(105) & wPagina & "/" & RsDados("PAGINANF") 'Inicio Impressão
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
    
    'Print #Notafiscal, wStr1
    If Glb_NfDevolucao = True Then
        WNatureza = "DEVOLUCAO"
        wStr1 = Space(2) & Left(Format(wStr17) & Space(34), 34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(15) & "X" & Space(16) & Left(Format(RsDados("nf"), "######"), 7)
    Else
        wStr1 = Space(2) & Left(Format(wStr17) & Space(34), 34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & Left(Format(RsDados("nf"), "######"), 7)
    End If
    Printer.Print wStr1
    wStr2 = Space(2) & Left(Format(wStr18) & Space(34), 34) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
    Printer.Print wStr2
    If Wserie = "CT" Then
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(34), 34) & Space(29) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-##'"), 9)
    Else
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(34), 34) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-###"), 9)
    End If
    Printer.Print wStr3
    If Wserie = "CT" Then
        wStr4 = ""
    Else
        wStr4 = Space(2) & Left(Format(wStr20) & Space(40), 40) & Space(46) & Left(Trim(Format(WCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
    Printer.Print ""
    'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
    If Wserie = "CT" Then
        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(18) & Left$(RsDados("CFOAUX"), 10) '& Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
        End If
    Else
        'If Trim(WNatureza) = "TRANSFERENCIA" Then
        '    wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
        If Trim(Wav) <> "" Then
            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(27) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = Space(31) & Left(Trim(WNatureza) & Space(26), 26) & Left$(RsDados("CFOAUX"), 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
    End If
    Printer.Print wStr5
    'Print #Notafiscal, ""
    Printer.Print ""
    Printer.Print ""
   ' If Mid(RsDados("CLIENTE"), 1, 5) <> "99999" Then
        wCgcCliente = Right(String(14, "0") & Trim(RsDados("CGCCLI")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = Right(String(18, "0") & Trim(wCgcCliente), 18)
   ' Else
   '     wCgcCliente = "00.000.000/0000-00"
   ' End If
    If Wserie = "CT" Then
        If wStr6 <> "" Then
            wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr6 = Space(36) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(45), 45) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        End If
    Else
        wStr6 = Left(Trim(wStr6) & Space(31), 31) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(45), 45) & Left$(Trim(wCgcCliente) & Space(24), 24) & Space(1) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
    End If
    
    Printer.Print wStr6
    If RsDados("EmiteDataSaida") = "S" Then
        If Wserie = "CT" Then
            wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Space(14) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(21), 21) & Right$(Space(11) & RsDados("CEPCLI"), 11) & Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
        End If
    Else
        If Wserie = "CT" Then
            wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) '& Space(14) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(21), 21) & Right$(Space(11) & RsDados("CEPCLI"), 11) '& Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
        End If
    End If
    
    Printer.Print ""
    Printer.Print wStr7
    If Wserie = "CT" Then
        wStr8 = ""
    Else
        wStr8 = Space(31) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
    End If
    Printer.Print ""
    Printer.Print wStr8
    
    Printer.Print ""
    Printer.Print ""

'              Printer.Print
'              Printer.Print Space(120) & wpagina & "/" & RsDados("PAGINANF")
'              Printer.Print wStr1
'              Printer.Print wStr2
'              Printer.Print wStr3
'              Printer.Print wStr4
'              Printer.Print
'              Printer.Print wStr5
'              Printer.Print
'              Printer.CurrentY = Printer.CurrentY + 2
'              Printer.Print wStr6
'              Printer.Print
'              Printer.CurrentY = Printer.CurrentY - 2
'              Printer.Print wStr7
'              Printer.Print
'              Printer.Print wStr8
'              Printer.Print
'              Printer.Print
           
End Function

Function CabecalhoOutrasSaidas(ByVal NatOper As String)
    Dim wCgcCliente As String
    Dim impri As Long
    Dim rdoConPag As rdoResultset
    impri = Printer.Orientation
    
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
    wLojaVenda = "            "
    wVendedorLojaVenda = "            "
    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))
    Wentrada = 0
    Wcondicao = "            "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = "               "
    
    WNatureza = NatOper
    
    wStr6 = ""
    
    Wcondicao = "            "
    
    wStr17 = "Pedido        : " & RsDados("NUMEROPED")
    wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
    wStr19 = ""

    If wPagina = 1 Then
        WCGC = Right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "###,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        WCGC = Right(String(18, "0") & WCGC, 18)
    End If
    wStr0 = Space(105) & wPagina & "/" & RsDados("PAGINANF") 'Inicio Impressão
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
    
    If Glb_NfDevolucao = True Then
        WNatureza = "DEVOLUCAO"
        wStr1 = Space(2) & Left(Format(wStr17) & Space(34), 34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(15) & "X" & Space(16) & Left(Format(RsDados("nf"), "######"), 7)
    Else
        wStr1 = Space(2) & Left(Format(wStr17) & Space(34), 34) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & Left(Format(RsDados("nf"), "######"), 7)
    End If
    Printer.Print wStr1
    wStr2 = Space(2) & Left(Format(wStr18) & Space(34), 34) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
    Printer.Print wStr2
    If Wserie = "CT" Then
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(34), 34) & Space(29) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-##'"), 9)
    Else
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(34), 34) & "(" & wDDDLoja & ")" & Left$(Trim(Format(WFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & Left$(Format(WFax, "###-####"), 9) & Space(5) & Left$(Format((WCep), "####-###"), 9)
    End If
    Printer.Print wStr3
    If Wserie = "CT" Then
        wStr4 = ""
    Else
        wStr4 = Space(2) & Left(Format(wStr20) & Space(40), 40) & Space(46) & Left(Trim(Format(WCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
    Printer.Print ""
    If Wserie = "CT" Then
        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(18) & Left$(RsDados("CFOAUX"), 10) '& Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
        End If
    Else
        If Trim(Wav) <> "" Then
            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(27) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = Space(31) & Left(Trim(WNatureza) & Space(26), 26) & Left$(RsDados("CFOAUX"), 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
    End If
    Printer.Print wStr5
    Printer.Print ""
    Printer.Print ""
    If Mid(RsDados("CLIENTE"), 1, 5) <> "99999" Then
        wCgcCliente = Right(String(14, "0") & Trim(RsDados("CGCCLI")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = Right(String(18, "0") & Trim(wCgcCliente), 18)
    Else
        wCgcCliente = "00.000.000/0000-00"
    End If
    If Wserie = "CT" Then
        If wStr6 <> "" Then
            wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr6 = Space(36) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(45), 45) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        End If
    Else
        wStr6 = Left(Trim(wStr6) & Space(31), 31) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(45), 45) & Left$(Trim(wCgcCliente) & Space(24), 24) & Space(1) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
    End If
    
    Printer.Print wStr6
    wStr7 = Space(2) & Left(wStr7 & Space(29), 29) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(21), 21) & Right$(Space(11) & RsDados("CEPCLI"), 11) & Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
    
    Printer.Print ""
    Printer.Print wStr7
    If Wserie = "CT" Then
        wStr8 = ""
    Else
        wStr8 = Space(31) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
    End If
    Printer.Print ""
    Printer.Print wStr8
    
    Printer.Print ""
    Printer.Print ""
           
End Function

Function CabecalhoNovo()

'     Wcondicao = "            "
'    Wav = "          "
'    If RsDados("CondPag") = 85 Then
'        wCarimbo4 = Format(RsDados("DataPag"), "mm/dd/yyyy")
'    Else
'        wCarimbo4 = IIf(IsNull(RsDados("Carimbo4")), "", RsDados("Carimbo4"))
'
'    End If
'    wLojaVenda = "            "
'    wVendedorLojaVenda = "            "
'    wLojaVenda = IIf(IsNull(RsDados("LojaVenda")), RsDados("LojaOrigem"), RsDados("LojaVenda"))
'    wVendedorLojaVenda = IIf(IsNull(RsDados("VendedorLojaVenda")), 0, RsDados("VendedorLojaVenda"))
'    Wentrada = 0
'    Wcondicao = "            "
'    wStr20 = ""
'    wStr19 = "               "
'    wStr7 = "               "
'    If Val(RsDados("CONDPAG")) = 1 Then
'       Wcondicao = "Avista"
'    ElseIf Val(RsDados("CONDPAG")) = 3 Then
'       Wcondicao = "Financiada"
'    ElseIf Val(RsDados("CONDPAG")) > 3 Then
'       Wcondicao = wCarimbo4
'    End If
'
'
'    If Trim(wLojaVenda) > 0 Then
'        If Trim(wLojaVenda) <> Trim(RsDados("LojaOrigem")) Then
'            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
'        Else
'            wStr6 = ""
'        End If
'    Else
'        wStr6 = ""
'    End If
'    If Trim(RsDados("AV")) > 1 Then
'        If Mid(Wcondicao, 1, 9) = "Faturada " Then
'            Wav = "AV            : " & Trim(RsDados("AV"))
'        End If
'    End If
'
'    If Trim(WNatureza) = "TRANSFERENCIAS" Then
'        Wcondicao = "            "
'    ElseIf Trim(WNatureza) = "DEVOLUCAO" Then
'        Wcondicao = "            "
'    End If
'
'    wStr17 = "Pedido        : " & RsDados("NUMEROPED")
'    wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
'    If Trim(Wcondicao) <> "" Then
'        wStr19 = "Cond. Pagto   : " & Trim(Wcondicao)
'    ElseIf Trim(RsDados("Carimbo3")) <> "" Then
'        wStr19 = Trim(RsDados("Carimbo3"))
'    Else
'        Wcondicao = "            "
'    End If
'
'    If RsDados("Pgentra") <> 0 Then
'       Wentrada = Format(RsDados("Pgentra"), "'"#####0.00")
'       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
'    End If
'    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
'        wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
'    End If
'
'
'    'wLinha2 = Space(2) & Left(rsDadosCliente("CLI_RazaoSocial") & Space(100), 100) _
'            & Left(rsDadosCliente("CLI_Cnpj") & Space(18), 18) _
'            & Left(Format(Date, "dd/mm/yyyy") & Space(10), 10)
'
'
'    Printer.Print Space(20) & wpagina & "/" & RsDados("PAGINANF")  'Inicio Impressão
'    'Printer.Print wStr1
'    'Print #Notafiscal, wStr1
'
'
'    wStr1 = Space(2) & Left(Format(wStr17) & Space(40), 40) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(15), 15) & Space(7) & "X" & Space(20) & Left(Format(RsDados("nf"), "'"',##'"), 7)
'    Printer.Print wStr1
'    wStr2 = Space(2) & Left(Format(wStr18) & Space(40), 40) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
'    Printer.Print wStr2
'    wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & "(011)" & Left$(Trim(Format(WFone, "'"#-###'")), 9) & "/(011)" & Left$(Format(WFone, "'"#-###'"), 9) & Space(5) & Left$(Format((WCep), "'"##-##'"), 9)
'    Printer.Print wStr3
'    Printer.Print ""
'    wStr4 = Space(2) & Left(Format(wStr20) & Space(20), 20) & Space(20) & Left(Trim(Format(WCGC, "'"',###,##'")), 19) '& Format(Mid((WCGC), 11, 5), "'"#-#'")
'    Printer.Print wStr4
'    Printer.Print ""
'    'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'    If Trim(WNatureza) = "TRANSFERENCIAS" Then
'        wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'    ElseIf Trim(Wav) <> "" Then
'        wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'    Else
'        wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(28) & Left$(Trim(Format((WIest), "'"',###,###,##'")), 15)
'    End If
'    Printer.Print wStr5
'    'Print #Notafiscal, ""
'    Printer.Print ""
'    If wStr6 <> "" Then
'        wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Trim(RsDados("CGCCLI")), 19) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'    Else
'        wStr6 = Space(34) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(11) & Left$(Trim(RsDados("CGCCLI")), 19) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'    End If
'    Printer.Print wStr6
'    Printer.Print ""
'    wStr7 = Space(2) & Left(wStr7 & Space(32), 32) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(19) & Left$(RsDados("CEPCLI"), 11) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'    Printer.Print wStr7
'    wStr8 = Space(34) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Space(8) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "'"',###,###,##'")), 15)
'    Printer.Print wStr8
'
'    Printer.Print ""
'    Printer.Print ""
'
'
'
''              Printer.Print
''              Printer.Print Space(120) & wpagina & "/" & RsDados("PAGINANF")
''              Printer.Print wStr1
''              Printer.Print wStr2
''              Printer.Print wStr3
''              Printer.Print wStr4
''              Printer.Print
''              Printer.Print wStr5
''              Printer.Print
''              Printer.CurrentY = Printer.CurrentY + 2
''              Printer.Print wStr6
''              Printer.Print
''              Printer.CurrentY = Printer.CurrentY - 2
''              Printer.Print wStr7
''              Printer.Print
''              Printer.Print wStr8
''              Printer.Print
''              Printer.Print

End Function


Function EmiteNotaTransferencia(ByVal Nf As Double, Serie As Double)
        
    
   
    
End Function


Sub GravaNFCapaDBFAccess()
    Dim wCfoAuxDev As String
    
    wTotalNotaTransferencia = 0
    wValorTotalCodigoZero = 0
    wValorTotalCodigoZero = 0
    WTipoNota = IIf(IsNull(NFCapaDBF("TIPOVENDA")), "", NFCapaDBF("TIPOVENDA"))
    
    If Not IsNull(NFCapaDBF("totnota")) Then
        If IsNull(NFCapaDBF("desconto")) Then
            Total = NFCapaDBF("totnota")
            wTotalNotaTransferencia = NFCapaDBF("totnota")
        Else
            SubTotal = (NFCapaDBF("totnota") + NFCapaDBF("desconto"))
            TotNota = NFCapaDBF("totnota")
            wTotalNotaTransferencia = NFCapaDBF("totnota")
        End If
    Else
        SubTotal = 0
        TotNota = 0
    End If
    
    wValorICMSAlternativa = IIf(IsNull(NFCapaDBF("VlrICMAlt")), 0, NFCapaDBF("VlrICMAlt"))
    wBaseICMSAlternativa = IIf(IsNull(NFCapaDBF("BaseIcmAlt")), 0, NFCapaDBF("BaseIcmAlt"))
    wValorTotalMercadoriaAlternativa = IIf(IsNull(NFCapaDBF("VlrMercAlt")), 0, NFCapaDBF("VlrMercAlt"))
    wTotalNotaAlternativa = IIf(IsNull(NFCapaDBF("TotNotaAlt")), 0, NFCapaDBF("TotNotaAlt"))
    wCarimbo3 = IIf(IsNull(NFCapaDBF("Observacao")), "", NFCapaDBF("Observacao"))
    'DataEmi = IIf(IsNull(NFCapaDBF("DATAEMI")), 0, NFCapaDBF("DATAEMI"))
    wCarimbo1 = IIf(IsNull(NFCapaDBF("Carimbo1")), "", NFCapaDBF("Carimbo1"))
    wCarimbo2 = IIf(IsNull(NFCapaDBF("Carimbo2")), "", NFCapaDBF("Carimbo2"))
    DataEmi = Format(Date, "dd/mm/yyyy")
    CODVEND = IIf(IsNull(NFCapaDBF("CODVEND")), 0, NFCapaDBF("CODVEND"))
    VLRMERC = IIf(IsNull(NFCapaDBF("VLRMERC")), 0, NFCapaDBF("VLRMERC"))
    Desconto = IIf(IsNull(NFCapaDBF("Desconto")), 0, NFCapaDBF("Desconto"))
    wLoja = IIf(IsNull(NFCapaDBF("Loja")), "", NFCapaDBF("Loja"))
    tipovenda = IIf(IsNull(NFCapaDBF("tipovenda")), "", NFCapaDBF("tipovenda"))
    CondPagto = IIf(IsNull(NFCapaDBF("condpagto")), 0, NFCapaDBF("condpagto"))
    av = IIf(IsNull(NFCapaDBF("AV")), 0, NFCapaDBF("AV"))
    Cliente = IIf(IsNull(NFCapaDBF("CLIENTE")), 0, NFCapaDBF("CLIENTE"))
    NatOper = IIf(IsNull(NFCapaDBF("NATOPER")), 0, NFCapaDBF("NATOPER"))
    DataPag = IIf(IsNull(NFCapaDBF("DATAPAG")), 0, NFCapaDBF("DATAPAG"))
    PgEntra = IIf(IsNull(NFCapaDBF("PGENTRA")), 0, NFCapaDBF("PGENTRA"))
    lojat = IIf(IsNull(NFCapaDBF("LOJAT")), "", NFCapaDBF("LOJAT"))
    TOTITENS = IIf(IsNull(NFCapaDBF("TOTITENS")), 0, NFCapaDBF("TOTITENS"))
    PEDCLI = IIf(IsNull(NFCapaDBF("PEDCLI")), 0, NFCapaDBF("PEDCLI"))
    PesoBr = IIf(IsNull(NFCapaDBF("PESOBR")), 0, NFCapaDBF("PESOBR"))
    PesoLq = IIf(IsNull(NFCapaDBF("PESOLQ")), 0, NFCapaDBF("PESOLQ"))
    OUTRALOJA = IIf(IsNull(NFCapaDBF("OUTRALOJA")), "", NFCapaDBF("OUTRALOJA"))
    ValFrete = IIf(IsNull(NFCapaDBF("VALFRETE")), 0, NFCapaDBF("VALFRETE"))
    FreteCobr = IIf(IsNull(NFCapaDBF("FRETECOBR")), 0, NFCapaDBF("FRETECOBR"))
    OUTROVEND = IIf(IsNull(NFCapaDBF("OUTROVEND")), 0, NFCapaDBF("OUTROVEND"))
    notafis = IIf(IsNull(NFCapaDBF("NOTAFIS")), 0, NFCapaDBF("NOTAFIS"))
    BASEICM = IIf(IsNull(NFCapaDBF("BASEICM")), 0, NFCapaDBF("BASEICM"))
    VLRICM = IIf(IsNull(NFCapaDBF("VLRICM")), 0, NFCapaDBF("VLRICM"))
    Wserie = IIf(IsNull(NFCapaDBF("SERIE")), "", NFCapaDBF("SERIE"))
    Hora = Time
    TOTIPI = IIf(IsNull(NFCapaDBF("TOTIPI")), 0, NFCapaDBF("TOTIPI"))
    Pedido = IIf(IsNull(NFCapaDBF("pedido")), 0, NFCapaDBF("pedido"))
    nomecli = IIf(IsNull(NFCapaDBF("nomcli")), 0, NFCapaDBF("nomcli"))
    endcli = IIf(IsNull(NFCapaDBF("endcli")), 0, NFCapaDBF("endcli"))
    muncli = IIf(IsNull(NFCapaDBF("muncli")), 0, NFCapaDBF("muncli"))
    cgccli = IIf(IsNull(NFCapaDBF("cgccli")), 0, NFCapaDBF("cgccli"))
    fonecli = IIf(IsNull(NFCapaDBF("fone")), 0, NFCapaDBF("fone"))
    Pessoa = IIf(IsNull(NFCapaDBF("pessoa")), 0, NFCapaDBF("pessoa"))
    ufcli = IIf(IsNull(NFCapaDBF("uf")), "SP", NFCapaDBF("uf"))
    cepcli = IIf(IsNull(NFCapaDBF("cep")), "", NFCapaDBF("cep"))
    bairrocli = IIf(IsNull(NFCapaDBF("bairro")), "", NFCapaDBF("bairro"))
    
    
    
    wValorTotalCodigoZero = 0
    If wValorTotalMercadoriaAlternativa > 0 Then
        wValorTotalCodigoZero = Val(TotNota - wValorTotalMercadoriaAlternativa)
    End If
    WTipoNota = IIf(IsNull(NFCapaDBF("TIPOVENDA")), "", NFCapaDBF("TIPOVENDA"))
    
    If NFCapaDBF("serie") = "RS" Then
        TipoNota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "RC" Then
        TipoNota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "RA" Then
        TipoNota = "RA"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "R2" Then
        TipoNota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    Else
        
        TipoNota = NFCapaDBF("TipoVenda")
        If TipoNota = "E" Then
            wNotaDevolucao = True
        Else
            wNotaDevolucao = False
        End If
    End If
    
    If IsNull(NFCapaDBF("datapag")) Then
        DataPag = "00:00:00"
    Else
        DataPag = NFCapaDBF("datapag")
    End If
    If IsNull(NFCapaDBF("condpagto")) Then
        CondPagto = 0
    Else
        CondPagto = NFCapaDBF("condpagto")
    End If
    If IsNull(NFCapaDBF("codvend")) Then
        CODVEND = "0"
    Else
        CODVEND = NFCapaDBF("codvend")
    End If
    If IsNull(NFCapaDBF("lojat")) Then
        lojat = "0"
    Else
        lojat = NFCapaDBF("lojat")
    End If
    If IsNull(NFCapaDBF("outraloja")) Then
        OUTRALOJA = "0"
    Else
        OUTRALOJA = NFCapaDBF("outraloja")
    End If
    If IsNull(NFCapaDBF("outrovend")) Then
        OUTROVEND = 0
    Else
        OUTROVEND = NFCapaDBF("outrovend")
    End If
    If IsNull(NFCapaDBF("ECF")) Then
        ECF = 0
    Else
        ECF = NFCapaDBF("ecf")
    End If
    If IsNull(NFCapaDBF("NUMEROSF")) Then
        numerosf = 0
    Else
        numerosf = NFCapaDBF("numerosf")
    End If
    If IsNull(NFCapaDBF("AV")) Then
        av = 0
    Else
        av = NFCapaDBF("av")
    End If
    If IsNull(NFCapaDBF("PorIcm")) Then
        PORICM = 0
    Else
        PORICM = NFCapaDBF("PorIcm")
    End If
    BeginTrans
    If Trim(NatOper) = 132 Then
        wCfoAuxDev = 1202
    ElseIf Trim(NatOper) = 232 Then
        wCfoAuxDev = 2202
    End If
       
    SQL = "Insert into nfcapa " & _
          "(numeroped,dataemi,vendedor,vlrmercadoria,desconto, subtotal,lojaorigem,tiponota,condpag,av,cliente, " & _
          "codoper,CfoAux,datapag,pgentra,lojat,qtditem,pedcli,tm, " & _
          "pesobr,pesolq,valfrete,fretecobr,outraloja,outrovend,nf,totalnota, " & _
          "baseicms,aliqicms,vlricms,serie,hora,totalipi,nomcli,fonecli,cgccli,endcli,ufcliente,municipiocli,pessoacli,ecf,numerosf, " & _
          "ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa, " & _
          "SituacaoEnvio,cepcli,bairrocli,EcfNF,Carimbo1,Carimbo2,Carimbo3) " & _
          "Values " & _
          "(" & Pedido & ", '" & Format(DataEmi, "MM/DD/YYYY") & "', " & CODVEND & ", " & ConverteVirgula(VLRMERC) & " , " & _
          "" & ConverteVirgula(Desconto) & " ," & ConverteVirgula(SubTotal) & " ,'" & wLoja & "', '" & TipoNota & "', " & _
          "" & CondPagto & " , " & av & " , " & Cliente & " , " & NatOper & ", " & wCfoAuxDev & " , " & _
          "'" & DataPag & "' , " & ConverteVirgula(PgEntra) & " , '" & lojat & "' , " & TOTITENS & " , " & _
          "" & PEDCLI & " , 1," & ConverteVirgula(PesoBr) & " , " & ConverteVirgula(PesoLq) & " , " & ConverteVirgula(ValFrete) & " , " & _
          "" & ConverteVirgula(FreteCobr) & " ,  '" & OUTRALOJA & "' , " & OUTROVEND & " , " & notafis & " , " & _
          "" & ConverteVirgula(TotNota) & ", " & ConverteVirgula(BASEICM) & ", " & ConverteVirgula(PORICM) & " ," & ConverteVirgula(VLRICM) & " , " & _
          "'" & Wserie & "' ,'" & Format(Hora, "hh:mm") & "'," & ConverteVirgula(TOTIPI) & ", '" & nomecli & "','" & fonecli & "','" & cgccli & "', '" & endcli & "','" & ufcli & "', " & _
          "'" & muncli & "'," & Pessoa & "," & Val(glb_ECF) & ", " & numerosf & ", " & _
          "" & ConverteVirgula(wValorTotalCodigoZero) & "," & ConverteVirgula(wTotalNotaAlternativa) & ", " & ConverteVirgula(wValorTotalMercadoriaAlternativa) & ",'A','" & cepcli & "','" & muncli & "','" & Val(glb_ECF) & "','" & wCarimbo1 & "','" & wCarimbo2 & "','" & wCarimbo3 & "') "
    
    rdoCNLoja.Execute (SQL)
    
    CommitTrans
   

End Sub

Sub GravaNFItemDBFAccess()
    
    
    
    If Not NFItemDBF.EOF Then
        Do While Not NFItemDBF.EOF
            
            If Nota = NFItemDBF("NOTAFIS") Then
                
                wValorMercadoriaAlternativa = IIf(IsNull(NFItemDBF("ValorMercA")), 0, NFItemDBF("ValorMercA"))
                wPrecoUnitarioAlternativa = IIf(IsNull(NFItemDBF("PrecoUniA")), 0, NFItemDBF("PrecoUniA"))
                wReferenciaAlternativa = IIf(IsNull(NFItemDBF("RefalternA")), 0, NFItemDBF("RefalternA"))
                Referencia = IIf(IsNull(NFItemDBF("referencia")), "", NFItemDBF("referencia"))
                Quant = IIf(IsNull(NFItemDBF("quant")), 0, NFItemDBF("quant"))
                unidade = IIf(IsNull(NFItemDBF("unidade")), "", NFItemDBF("unidade"))
                PrecoUni = IIf(IsNull(NFItemDBF("precouni")), 0, NFItemDBF("precouni"))
                valormerc = IIf(IsNull(NFItemDBF("valormerc")), 0, NFItemDBF("valormerc"))
                notafis = IIf(IsNull(NFItemDBF("notafis")), 0, NFItemDBF("notafis"))
                Wserie = IIf(IsNull(NFItemDBF("serie")), "", NFItemDBF("serie"))
                wLoja = IIf(IsNull(NFItemDBF("loja")), "", NFItemDBF("loja"))
                Cliente = IIf(IsNull(NFItemDBF("cliente")), 0, NFItemDBF("cliente"))
                aliqipi = IIf(IsNull(NFItemDBF("aliqipi")), 0, NFItemDBF("aliqipi"))
                plista = IIf(IsNull(NFItemDBF("plista")), 0, NFItemDBF("plista"))
                icms = IIf(IsNull(NFItemDBF("icms")), 0, NFItemDBF("icms"))
                Desconto = IIf(IsNull(NFItemDBF("desconto")), 0, NFItemDBF("desconto"))
                Comissao = IIf(IsNull(NFItemDBF("Comissao")), 0, NFItemDBF("Comissao"))
                bcomis = IIf(IsNull(NFItemDBF("bcomis")), 0, NFItemDBF("bcomis"))
                Linha = IIf(IsNull(NFItemDBF("linha")), 0, NFItemDBF("linha"))
                Secao = IIf(IsNull(NFItemDBF("Secao")), 0, NFItemDBF("Secao"))
                csprod = IIf(IsNull(NFItemDBF("csprod")), 0, NFItemDBF("csprod"))
                vlripi = IIf(IsNull(NFItemDBF("vlripi")), 0, NFItemDBF("vlripi"))
                Item = IIf(IsNull(NFItemDBF("Item")), 0, NFItemDBF("Item"))
                'DataEmi = IIf(IsNull(NFItemDBF("DataEmi")), 0, NFItemDBF("DataEmi"))
                DataEmi = Format(Date, "dd/mm/yyyy")
                CODVEND = IIf(IsNull(NFItemDBF("codvend")), 0, NFItemDBF("codvend"))
                vlripi = IIf(IsNull(NFItemDBF("vlripi")), 0, NFItemDBF("vlripi"))
                PedidoItem = IIf(IsNull(NFCapaDBF("pedido")), 0, NFCapaDBF("pedido"))
                tipomov = IIf(IsNull(NFItemDBF("tipomov")), 0, NFItemDBF("tipomov"))
                sitnf = IIf(IsNull(NFItemDBF("sitnf")), 0, NFItemDBF("sitnf"))
                Status = IIf(IsNull(NFItemDBF("status")), 0, NFItemDBF("status"))
                wUltimoItem = Item
                wVlUnit2 = Format(valormerc - Desconto, "0.00")
                
                
                If NFItemDBF("serie") = "RS" Then
                    TipoNota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "RC" Then
                    TipoNota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "RA" Then
                    TipoNota = "RA"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "R2" Then
                    TipoNota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                Else
                    If tipomov = 12 Then
                        TipoNota = "T"
                    ElseIf tipomov = 23 Then
                        TipoNota = "E"
                    End If
                End If
                
                
                BeginTrans
                
                SQL = "Insert INTO nfitens " & _
                      "(numeroped,dataemi,referencia,qtde,vlunit,vlunit2,vltotitem," & _
                      "item,vlipi,desconto,plista,comissao,icms,bcomis,csprod,linha,secao," & _
                      "nf,serie,lojaorigem,cliente,vendedor,aliqipi,tiponota,tipomovimentacao, " & _
                      "ValorMercadoriaAlternativa,PrecoUnitAlternativa,ReferenciaAlternativa,SituacaoEnvio) " & _
                      "Values " & _
                      "(" & PedidoItem & ",'" & Format(DataEmi, "MM/DD/YYYY") & "','" & Referencia & "'," & Quant & "," & _
                      "" & ConverteVirgula(PrecoUni) & "," & ConverteVirgula(wVlUnit2) & "," & ConverteVirgula(valormerc) & "," & Item & "," & ConverteVirgula(vlripi) & "," & ConverteVirgula(Desconto) & "," & _
                      "" & ConverteVirgula(plista) & "," & Comissao & "," & icms & "," & bcomis & "," & csprod & "," & _
                      "" & Linha & "," & Secao & "," & notafis & ",'" & Wserie & "','" & wLoja & "'," & _
                      "" & Cliente & "," & CODVEND & "," & aliqipi & ",'" & TipoNota & "'," & tipomov & ", " & _
                      "" & ConverteVirgula(wValorMercadoriaAlternativa) & "," & ConverteVirgula(wPrecoUnitarioAlternativa) & "," & ConverteVirgula(wReferenciaAlternativa) & ",'A') "
                
                rdoCNLoja.Execute (SQL)
                CommitTrans
                BeginTrans
        
                DBFBanco.Execute "Update nfitem.dbf set " & _
                    "sitnf='P' WHERE notafis=" & Nota & " " & _
                    "and referencia='" & Referencia & "'", dbFailOnError
        
                CommitTrans
                
            End If
            NFItemDBF.MoveNext
        Loop
            
    End If
        BeginTrans
    
        DBFBanco.Execute "Update nfcapa.dbf set " & _
            "sitnf='P' WHERE notafis=" & Nota & "", dbFailOnError
    
        CommitTrans

   
    
End Sub


Public Function CriaArquivoNFTransferencia()

'    Dim WZERO As Double
'
'
'    wpagina = 1
'    wNotaTransferencia = True
'    If WTipoNota = "T" Then
'       WNatureza = "TRANSFERENCIAS"
'    Else
'       WNatureza = "DEVOLUCAO"
'    End If
'
'    Temporario = "C:\NOTASVB\"
'
'    Call DadosLoja
'
'    SQL = ""
'    SQL = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.Carimbo5,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV,NFCAPA.Carimbo3,NFCAPA.Carimbo2,NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.LOJAT, " _
'        & "NFCAPA.CLIENTE,NFCAPA.FONECLI,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
'        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
'        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
'        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
'        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
'        & "NFCAPA.UFCLIENTE,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
'        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
'        & "From NFCAPA INNER JOIN NFITENS " _
'        & "on (NfCapa.nf=Nfitens.nf) " _
'        & "Where NfCapa.nf= " & WNfTransferencia & " " _
'        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
'
'    Set RsDados = rdoCnLoja.OpenResultset(SQL)
'
'    If Not RsDados.EOF Then
'
'      Call CabecalhoArq
'
'      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
'          & "produto.pr_classefiscal,produto.pr_unidade, " _
'          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
'          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao " _
'          & "from produto,nfitens " _
'          & "where produto.pr_referencia=nfitens.referencia " _
'          & "and nfitens.nf = " & WNfTransferencia & " order by nfitens.item"
'
'      Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
'
'      If Not RsdadosItens.EOF Then
'         wConta = 0
'         Do While Not RsdadosItens.EOF
'
'
'               If Wsm = True Then
'                    wPegaDescricaoAlternativa = IIf(IsNull(RsDados("Referencia")), "0", RsDados("Referencia"))
'                      wStr16 = ""
'                      wStr16 = Left$(RsdadosItens("ReferenciaAlternativa") & Space(8), 8) _
'                             & Space(2) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
'                             & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
'                             & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
'                             & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
'                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "'"##0"), 6) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "'"#####0.00"), 12) & Space(1) _
'                             & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "'"#####0.00"), 15) & Space(1) _
'                             & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "'0"), 2)
'
'               Else
'
'                      WZERO = 0
'                      wStr16 = ""
'                      wStr16 = Left$(RsdadosItens("pr_referencia") & Space(8), 8) _
'                            & Space(2) & Left$(Format(Trim(RsdadosItens("pr_descricao")), ">") & Space(38), 38) _
'                            & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
'                            & Space(10), 10) & Space(2) & Left$(Trim(WZERO), 1) & Left$(Trim(WZERO), 1) _
'                            & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
'                            & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "'"##0"), 6) & Space(2) _
'                            & Right$(Space(12) & Format(RsdadosItens("vlunit"), "'"#####0.00"), 12) & Space(1) _
'                            & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "'"#####0.00"), 15) & Space(1) _
'                            & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "'0"), 2)
'
'
'               End If
'
'                      Print #NotaFiscal, wStr16
'
'                      If RsdadosItens("DetalheImpressao") = "D" Then
'                         wConta = wConta + 1
'                         RsdadosItens.MoveNext
'                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
'                         Do While wConta < 21
'                            wConta = wConta + 1
'                            Print #NotaFiscal, ""
'                         Loop
'                         RsdadosItens.MoveNext
'                         wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "'"',##'"), 7)
'                         Print #NotaFiscal, wStr13
'                         Print #NotaFiscal, ""
'                         Print #NotaFiscal, ""
'                         Print #NotaFiscal, Chr(18) 'Finaliza Impressão
'                         Close #NotaFiscal
'                         wConta = 0
'                         wpagina = wpagina + 1
'                         FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
''                         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
'                         Call CabecalhoArq
'                      ElseIf RsdadosItens("DetalheImpressao") = "T" Then
'                         wConta = wConta + 1
'                         RsdadosItens.MoveNext
'                         Call FinalizaArqNf
'                      Else
'                         wConta = wConta + 1
'                         RsdadosItens.MoveNext
'                      End If
'
'            Loop
'         Else
'            Close #NotaFiscal
'            MsgBox "Produto não encontrado", vbInformation, "Aviso"
'         End If
'
'         'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'         'FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
'         'FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\Notas" & NomeArquivo
'
'
'    End If
End Function

Function QuebraNotaDevolucao(ByVal wNumeroNota As Double)
    wQuantdadeTotalItem = 0
    wUltimoItem = 0
    SQL = ""
    SQL = "Select NfCapa.QtdItem,NfItens.Item from NfCapa,NfItens " _
        & "where Nfcapa.Nf=" & wNumeroNota & " " _
        & "and NfItens.Nf=NfCapa.NF order by NfItens.Item "
        Set RsCapaNF = rdoCNLoja.OpenResultset(SQL)
    
    If Not RsCapaNF.EOF Then
        Do While Not RsCapaNF.EOF
            wQuantItensCapaNF = RsCapaNF("QtdItem")
            wQuantItensNF = RsCapaNF("Item")
            'wUltimoItem = RsCapaNF("Item")
            wQuantdadeTotalItem = wQuantdadeTotalItem + 1
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
    
            SQL = ""
            SQL = "Update NfItens set DetalheImpressao='" & wDetalheImpressao & "' where Nf=" & wNumeroNota & " and Item=" & wQuantItensNF & ""
                rdoCNLoja.Execute (SQL)
        
            RsCapaNF.MoveNext
        Loop
    
    SQL = ""
    SQL = "Update NfCapa set PaginaNF=" & wUltimoItem & " "
        rdoCNLoja.Execute (SQL)
    
    End If
End Function


Sub BuscaTransferencia()


Dim Wtexto As String
Dim WAtualizados As String
Dim Wmaximo As Integer
Dim i As Integer
Dim Conta As Integer
Dim matArquivos() As String
Dim WARQUIVO As String

    Wtexto = WcaminhoTextos
    WAtualizados = WcaminhoTextosAtu
    
    WARQUIVO = Dir(Wtexto)
    
        Do While WARQUIVO <> ""
            If Mid(WARQUIVO, 1, 2) = "tr" Or Mid(WARQUIVO, 1, 2) = "ct" Then
                If Mid(WARQUIVO, 1, 2) = "ct" Then
                    Wserie = "CT"
                Else
                    Wserie = ""
                End If
                wNumPed = Mid(WARQUIVO, 3, Len(WARQUIVO) - 6)
                wNfCapa = False
                wNFitens = False
                arquivo = FreeFile
                SQL = ""
                SQL = "Select NumeroPed from NfCapa where NumeroPed=" & Mid(WARQUIVO, 3, Len(WARQUIVO) - 6) & " "
                    Set RsVerificaPedido = rdoCNLoja.OpenResultset(SQL)
                If RsVerificaPedido.EOF Then
                    Open Wtexto & WARQUIVO For Input Access Read As #arquivo
        
                    Do While Not EOF(arquivo)
                        Line Input #arquivo, buffer
                        If Mid(buffer, 1, 3) = "000" Then
                            Call AtualizaCapaTransf
                        ElseIf Mid(buffer, 1, 3) <> "PRO" Then
                            Call AtualizaItensTransf
                        End If
                    Loop
                    If wNfCapa = False Or wNFitens = False Then
                        PegaItensPedTransf False, ""
                    End If
                    wReemissao = False
                    NotaTransferencia Mid(WARQUIVO, 3, Len(WARQUIVO) - 6)
                End If
                Close #arquivo
                FileCopy Wtexto & WARQUIVO, WAtualizados & WARQUIVO
                Kill Wtexto & WARQUIVO
        
            End If
            WARQUIVO = Dir()
        Loop
End Sub


Sub AtualizaCapaTransf()

    SQL = "select  CT_Loja from Controle"
        Set rsPegaLoja = rdoCNLoja.OpenResultset(SQL)
    If Not rsPegaLoja.EOF Then
        wLoja = rsPegaLoja("CT_Loja")
    End If
        
        
    wVendedorLojaVenda = 0
    wLojaVenda = ""
    WTotPedido = 0
    wSubTotal = 0
    wTotalNotaAlternativa = 0
    wValorTotalCodigoZero = 0
    WnumeroPed = Mid(buffer, 4, 8)
    WCliente = Mid(buffer, 12, 6)
    WNomeCliente = Mid(buffer, 39, 18)
    WVendedor = Mid(buffer, 207, 3)
    WCOMISSAO = "7"
    wTotalNotaAlternativa = 0
    wValorTotalCodigoZero = 0
    wCarimbo3 = ""
    wPedidoCliente = 0
    Wdata = Format(Date, "dd/mm/yyyy")
    If Trim(Mid(buffer, 249, 14)) <> "" Then
       WTotPedido = ConverteVirgula2(Mid(buffer, 249, 14))
    End If
    If Trim(Mid(buffer, 235, 14)) <> "" Then
       wSubTotal = ConverteVirgula2(Mid(buffer, 235, 14))
    End If
    
    Wlojat = ConverteVirgula2(Mid(buffer, 331, 5))
    If Trim(Mid(buffer, 336, 15)) <> "" Then
        wCarimbo3 = Mid(buffer, 342, 15)
    End If
    Wdescontop = Format(wSubTotal - WTotPedido, "0.00")
    'Wlojat = "999"
    
    
    
    If Mid(buffer, 267, 3) <> "" Then
       WCODOPER = Mid(buffer, 267, 3)
       WCFOAux = WCODOPER
       If WCFOAux = 522 Then
            WCFOAux = 5152
       End If
       WPGENTRA = ConverteVirgula2(Mid(buffer, 271, 14))
       'WCONDPAG = Mid(BUFFER, 286, 2)
       wCondPag = 0
       WDESCRIPAG = Mid(buffer, 289, 13)
    End If
    
    wQtdItem = Mid(buffer, 224, 3)
    
    Wtipovenda = "T"
    Wav = 1
    wSituacao = "F"
    WSTATUS = "F"
    WTm = 1
    WPesoBr = 0
    WPesoLq = 0
    wValFrete = 0
    WFRETECOBR = 0
    Wtipo = "T"
    WOutraLoja = ""
    WTipoNota = "T"
    WTipoP = "F"
    If Wserie = "" Then
       Wserie = "SN"
    End If
    
    WNOMCLI = Mid(buffer, 18, 39)
    WENDCLI = Mid(buffer, 57, 39)
    WENDENTCLI = Mid(buffer, 57, 39)
    wbairro = Mid(buffer, 96, 16)
    WMUNCLI = Mid(buffer, 112, 21)
    WUF = Mid(buffer, 133, 2)
    WREGIAO = Mid(buffer, 96, 16)
    WCep = Mid(buffer, 135, 9)
    WIest = Mid(buffer, 157, 16)
    WCGCCLI = Mid(buffer, 174, 18)
    WDDD = 0
    WFone = Mid(buffer, 192, 9)
    If IsDate(Mid(buffer, 304, 14)) = True Then
        WdataPag = Mid(buffer, 305, 14)
    Else
        WdataPag = "00:00:00"
    End If
    wPessoa = 1
          
    If WOutraLoja = "" Then
        WOutraLoja = 0
    End If
    Call GravaNFCapa
End Sub


Sub AtualizaItensTransf()

    WNF = 0
    wVlUnit = 0
    wVlUnit2 = 0
    wVlTotItem = 0
    WVLUNITAL = 0
    WVLTOTITEMAL = 0
    Wteste = ""
    WREFALTERNA = 0
    wPegaDescricaoAlternativa = "0"
    wValorMercadoriaAlternativa = 0
    wValorTotalItemAlternativa = 0
    wTipoMovimentacao = 12
    
    wReferencia = Mid(buffer, 4, 13)
    wQtde = Val(Mid(buffer, 49, 6))
    WTP = 1
    If Trim(Mid(buffer, 57, 14)) <> "" Then
       wVlUnit = ConverteVirgula2(Mid(buffer, 57, 14))
       wVlUnit2 = ConverteVirgula2(Mid(buffer, 85, 14))
       wVlTotItem = wVlUnit * wQtde
    End If
    
    WDESCRAT = ConverteVirgula2(Mid(buffer, 99, 14))
    wSituacao = "F"
    WSTATUS = "F"
    wItem = Val(Mid(buffer, 1, 3))
    WSERIE1 = "123"
    WSERIE2 = "456"
    WSERIE3 = "789"
    Wserie = "SN"
    WENTRAT = 0
    WVLIPI = 0
    Wdescontop = 0
    WCOMISSAO = 0
    WCMR = 0
    wICMS = 0
    WBCOMIS = 0
    WVBUNIT = 0
    WPERDESC = 0


    SQL = "Select * from produto where Pr_Referencia= '" & Trim(wReferencia) & "'"
    Set ISQL = rdoCNLoja.OpenResultset(SQL)
    
    If Not ISQL.EOF Then
       wLinha = ISQL("Pr_LINHA")
       wSecao = ISQL("Pr_SECAO")
       WSUBTRIBUT = ISQL("PR_SubstituicaoTributaria")
       wPLISTA = ISQL("pr_precovenda1")
       WTRIBUTO = ISQL("pr_icmssaida")
       wIcmPdv = ISQL("pr_icmssaida")
       wCodBarra = ISQL("pr_codigobarra")
    Else
       If MsgBox("Referencia " & Trim(wReferencia) & "  do pedido de transferencia " & Val(wNumPed) & " não encontrada, " & Chr(10) & " Deseja cadastrar essa referencia agora", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            GravaProduto wReferencia
       Else
            MsgBox "A Transferencia só podera ser emitida quando esta referencia estiver cadastrada", vbCritical, "Atenção"
       End If
       PegaItensPedTransf True, Trim(wReferencia)
       Exit Sub
    End If
 
    Call GravaNfItens


End Sub

Function PegaItensPedTransf(ByVal IncluirReferencia As Boolean, ByVal Referencia As String)
    
    If IncluirReferencia = True Then
        SQL = ""
        SQL = "Select * from Pedi " _
        & "where NUMERO=" & wNumPed & " and Referencia = '" & Referencia & "'"
    Else
        SQL = ""
        SQL = "Select * from Pedi " _
            & "where NUMERO=" & wNumPed & ""
    End If
        Set RsPegaItensPedi = DBFBanco.OpenRecordset(SQL)
        
    If Not RsPegaItensPedi.EOF Then
        Do While Not RsPegaItensPedi.EOF
            'SELECT [NUMERO], [DATA], [REFERENCIA], [QTDE], [TP], [VLUNIT], [VLUNIT2]
            ', [VLTOTITEM], [DESCRAT], [TRIBUTO], [CONTROLE], [ITEM], [SITUACAO],
            '[Status] , [ENTRAT], [VLIPI], [SERIE1], [SERIE2], [SERIE3], [Desconto],
            '[plista] , [Comissao], [CMR], [icms], [bcomis], [csprod], [Linha], [Secao],
            '[VBUNIT] , [PERDESC], [SUBTRIBUT], [ICMPDV], [CODBARRA], [VLUNITALT], [VLTOTITEMA],
            '[VBUNITA]

            wNFitens = True
            WNF = 0
            wVlUnit = 0
            wVlUnit2 = 0
            wVlTotItem = 0
            WVLUNITAL = 0
            WVLTOTITEMAL = 0
            Wteste = ""
            WREFALTERNA = 0
            wPegaDescricaoAlternativa = "0"
            wValorMercadoriaAlternativa = 0
            wValorTotalItemAlternativa = 0
            wTipoMovimentacao = 12
            wReferencia = RsPegaItensPedi("Referencia")
            wQtde = RsPegaItensPedi("QTDE")
            WTP = 1
            wVlUnit = RsPegaItensPedi("VLunit")
            wVlUnit2 = RsPegaItensPedi("VlUnit2")
            wVlTotItem = wVlUnit * wQtde
            wValorMercadoriaAlternativa = 0
            wValorTotalItemAlternativa = 0
            WDESCRAT = IIf(IsNull(RsPegaItensPedi("Descrat")), "0", RsPegaItensPedi("Descrat"))
            wSituacao = "F"
            WSTATUS = "F"
            wItem = RsPegaItensPedi("Item")
            WSERIE1 = "123"
            WSERIE2 = "456"
            WSERIE3 = "789"
            WENTRAT = 0
            WVLIPI = 0
            Wdescontop = 0
            WCOMISSAO = 0
            WCMR = 0
            wICMS = 0
            WBCOMIS = 0
            WVBUNIT = 0
            WPERDESC = 0
            wLinha = IIf(IsNull(RsPegaItensPedi("Linha")), 0, RsPegaItensPedi("Linha"))
            wSecao = IIf(IsNull(RsPegaItensPedi("Secao")), 0, RsPegaItensPedi("Secao"))
            WSUBTRIBUT = IIf(IsNull(RsPegaItensPedi("SubTribut")), 0, RsPegaItensPedi("SubTribut"))
            wPLISTA = IIf(IsNull(RsPegaItensPedi("PLista")), 0, RsPegaItensPedi("PLista"))
            WTRIBUTO = IIf(IsNull(RsPegaItensPedi("Tributo")), 0, RsPegaItensPedi("Tributo"))
            wIcmPdv = IIf(IsNull(RsPegaItensPedi("ICMPDV")), 0, RsPegaItensPedi("ICMPDV"))
            wCodBarra = IIf(IsNull(RsPegaItensPedi("CodBarra")), "0", RsPegaItensPedi("CodBarra"))
            Call GravaNfItens
            RsPegaItensPedi.MoveNext
        Loop
    End If
    RsPegaItensPedi.Close
    
    
End Function



Function NotaTransferencia(ByVal NumeroPedido As Double)
    
    If VerificaItensPedido(NumeroPedido, "T") = True Then
        SQL = ""
        SQL = "select CT_SeqNota,CT_LOJA from Controle"
            Set RsPegaNumNote = rdoCNLoja.OpenResultset(SQL)
        wLoja = RsPegaNumNote("ct_loja")
        If Not RsPegaNumNote.EOF Then
            Call ExtraiSequenciaNotaTransferencia
            SQL = ""
            SQL = "Update NfCapa set NF = " & WNfTransferencia & " " _
                & "Where NumeroPed = " & NumeroPedido & ""
                rdoCNLoja.Execute (SQL)
                            
            SQL = ""
            SQL = "Update NfItens set NF=" & WNfTransferencia & ",TipoNota='T' " _
                & "Where Numeroped=" & NumeroPedido & " "
                rdoCNLoja.Execute (SQL)
                
            If EncerraVendaMigracao(NumeroPedido, " ", 0) = False Then
                SQL = ""
                SQL = "Delete * from NfCapa where NumeroPed=" & NumeroPedido & " and TipoNota='T'"
                    rdoCNLoja.Execute (SQL)
                
                SQL = ""
                SQL = "Delete * from NfItens where NumeroPed=" & NumeroPedido & " and TipoNota='T'"
                rdoCNLoja.Execute (SQL)
                
                MsgBox "Não foi possivel emitir a transferencia", vbCritical, "Aviso"
                
                Exit Function
            End If
                
            If Wserie = "CT" Then
                WNF = WNfTransferencia
                'Call AtualizaEstoque(WNfTransferencia, "CT", 0)
                EmiteNotafiscal WNF, Wserie
            Else
                'Call AtualizaEstoque(WNfTransferencia, "SN", 0)
                EmiteNotafiscal WNfTransferencia, PegaSerieNota
                'Call CriaArquivoNFTransferencia
            End If
        End If
    Else
        SQL = ""
        SQL = "Delete * from NfCapa where NumeroPed=" & NumeroPedido & " and TipoNota='T'"
            rdoCNLoja.Execute (SQL)
        
        SQL = ""
        SQL = "Delete * from NfItens where NumeroPed=" & NumeroPedido & " and TipoNota='T'"
        rdoCNLoja.Execute (SQL)
        
        MsgBox "Não foi possivel emitir a transferencia", vbCritical, "Aviso"
    End If
End Function


Function ProcessaRotinasDiarias()
    
    Dim rsPegaData As rdoResultset
    'frmAguarde.Show
    
    On Error Resume Next
    SQL = ""
    SQL = "Select CT_Loja,CT_TipoArquivo,CT_OnLine,CT_ECF from Controle"
        Set RSTipoControle = rdoCNLoja.OpenResultset(SQL)
            
    If Not RSTipoControle.EOF Then
        If RSTipoControle("CT_TipoArquivo") = 1 Then
            frmRotinasDiaria.lblRotinas.Caption = " Iniciando o dia"
            BeginTrans
                SQL = ""
                SQL = "Delete * from EstqLojaDBF where Situacao='P'"
                rdoCNLoja.Execute (SQL)
            CommitTrans
            BeginTrans
                SQL = ""
                SQL = "Delete * from ProduLjDBF where Situacao='P'"
                rdoCNLoja.Execute (SQL)
            CommitTrans
            Call ZeraVendaDia
            Call LimpaHoraAtualizacao
            BeginTrans
                frmRotinasDiaria.lblProcessos.Caption = "Atualizando Controle"
                SQL = ""
                SQL = "Update Controle set CT_TipoArquivo = 0"
                rdoCNLoja.Execute (SQL)
            CommitTrans
            frmRotinasDiaria.lblProcessos.Caption = "  BOAS VENDAS"
            
        ElseIf RSTipoControle("CT_TipoArquivo") = 2 Then
            frmRotinasDiaria.lblRotinas.Caption = " Fechando o dia"
            Screen.MousePointer = 11
            frmAbrirFecharCaixa.Refresh
            frmAguarde.Refresh
            frmAguarde.ZOrder
            frmAguarde.Show
            
            'Call AtualizaLoja
'            Call ConfereMovimentoEstoque
            frmAguarde.lblMensagem.Caption = "Aguarde, Emitindo Redução Z"
            If Format(AchaDataCaixa, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
               If Trim(RSTipoControle("CT_ECF")) = "S" Then
                   ' Call LeituraZ
               End If
            End If
            frmRotinasDiaria.lblProcessos.Caption = "Atualizando Controle"
            BeginTrans
                SQL = ""
                SQL = "Update Controle set CT_TipoArquivo = 0"
                rdoCNLoja.Execute (SQL)
            CommitTrans
            '
            '----------------------Importando Movimento Dia-------------------------
            '
            If Trim(GLB_NumeroCaixa) = 1 Then
                
                frmAguarde.lblMensagem.Caption = "Aguarde, Gravando Movimento Diario"
                
                SQL = "Select Max(CT_Data) as Data from CtCaixa"
                    Set rsPegaData = rdoCNLoja.OpenResultset(SQL)
                '-------------------------Auditor do Estoque----------------------------
                AtualizaProcessoTela "Auditor do Estoque", 1, 1
                AtualizaProcessoTela "Exclui Produto Deletado", 2, 1
                AtualizaProcessoTela "Giro do Estoque", 3, 1
                SQL = ""
                SQL = "ProcAuditorEstoqueLoja '" & Format(rsPegaData("Data"), "mm/dd/yyyy") & "',2 "
                    rdoCNLoja.Execute (SQL)
                If ComparaEstoqueAnterior = False Then
                    MsgBox "Não foi possivel fazer o giro do estoque favor avisar o CPD", vbCritical, "Erro Girando o Estoque"
                Else
                    frmAguarde.Refresh
                    AtualizaProcessoTela "Auditor do Estoque", 0, 2
                    AtualizaProcessoTela "Exclui Produto Deletado", 1, 2
                    AtualizaProcessoTela "Giro do Estoque", 2, 2
                    frmAguarde.Refresh
                End If
                '-----------------------------------------------------------------------
                '----------------------Salva Tabelas no banco historico-----------------
                AtualizaProcessoTela "Salvando Tabelas", 4, 1
                SQL = ""
                SQL = "ProcSalvaTabelas"
                    rdoCnLojaBach.Execute (SQL)
                AtualizaProcessoTela "Salvando Tabelas", 3, 2
                frmAguarde.Refresh
                Esperar 1
                frmAguarde.Refresh
                '------------------------Fecha Caixa Central----------------------------
                'frmAguarde.lblMensagem.Caption = "Aguarde, Fechando Caixa na Central"
                'AtualizaProcessoTela "Fechamento do Caixa Central", 1, 1
                'frmFechamentoLoja.Refresh
                'frmAguarde.Visible = False
                'frmAguarde.Refresh
                'Screen.MousePointer = 0
                'frmFechamentoLoja.Show 1
                'Screen.MousePointer = 11
                'frmAguarde.Visible = True
                'frmAguarde.Refresh
                'Esperar 1
                'If FechaLojaCentral(RSTipoControle("CT_Loja")) = True Then
                '    AtualizaProcessoTela "Fechamento do Caixa Central", 1, 2
                'End If
                '-----------------------------------------------------------------------
                '-------------------------Cria Relatorio codigo zero--------------------
                'AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "A"
                'CriaRelatorioCodigoZero rsPegaData("Data"), RSTipoControle("CT_Loja")
                '-----------------------------------------------------------------------
                '-------------------------Auditor do estoque----------------------------
                'AtualizaProcessoTela "Auditor do Estoque", 2, 1
                'AuditorEstoque rsPegaData("data")
                'AtualizaProcessoTela "Auditor do Estoque", 2, 2
                '-----------------------------------------------------------------------
                '-------------------------Giro do Estoque-------------------------------
                'AtualizaProcessoTela "Giro do Estoque", 3, 1
                'Call AtualizaEstoqueAnterior
                'If ComparaEstoqueAnterior = False Then
                '    Call AtualizaEstoqueAnterior
                '    If ComparaEstoqueAnterior = False Then
                '        MsgBox "Não foi possivel fazer o giro do estoque favor avisar o CPD", vbCritical, "Erro Girando o Estoque"
                '    Else
                '        AtualizaProcessoTela "Giro do Estoque", 3, 2
                '    End If
                'Else
                '    AtualizaProcessoTela "Giro do Estoque", 3, 2
                'End If

                '-----------------------------------------------------------------------
                'frmAguarde.lblMensagem.Caption = "Aguarde, Fechamento Central"
                'FechamentoCaixaCentral rsPegaData("Data"), RSTipoControle("CT_Loja")
                '------------------------Atualiza Produto Bc2000------------------------
                'frmAguarde.lblMensagem.Caption = "Aguarde, Atualizando Produto Bc2000"
                'AtualizaProcessoTela "Atualização do Produto Balcao2000", 4, 1
                'AtualizaProdutoBc2000
                'AtualizaProcessoTela "Atualização do Produto Balcao2000", 4, 2
                '-----------------------------------------------------------------------
                '------------------------Atualiza Produto Balcao------------------------
                'frmAguarde.lblMensagem.Caption = "Aguarde, Atualizando Produto DBF"
                'AtualizaProcessoTela "Atualização do Produto Balcão", 5, 1
                'AtualizaProdutoDBF
                'AtualizaProcessoTela "Atualização do Produto Balcão", 5, 2
                '----------------------------------------------------------------------
                '------------------------Acerta estoque do dbf-------------------------
                'AtualizaProcessoTela "Atualização do Estoque Balcão", 6, 1
                'AcertaEstoqueDBF rsPegaData("data")
                'AtualizaProcessoTela "Atualização do Estoque Balcão", 6, 2
                'frmAguarde.PrbContador.Visible = False
                '----------------------------------------------------------------------
                '------------------------Compactando banco de dados--------------------
                'AtualizaProcessoTela "Compactação do Banco de Dados Balcao2000", 7, 1
                'frmAguarde.lblMensagem.Caption = "Aguarde, Compactando Banco"
                'CompactaBanco "MovimentoCaixa", "MC_Data < '" & Format(DateAdd("m", -3, Date), "mm/dd/yyyy") & "'"
                'CompactaBanco "MovimentacaoEstoque", "ME_DataMovimento < '" & Format(DateAdd("m", -3, Date), "mm/dd/yyyy") & "'"
                'CompactaBanco "NfItens", "DataEmi < '" & Format(DateAdd("m", -3, Date), "mm/dd/yyyy") & "'"
                'CompactaBanco "NfCapa", "DataEmi < '" & Format(DateAdd("m", -3, Date), "mm/dd/yyyy") & "'"
                'AtualizaProcessoTela "Compactação do Banco de Dados Balcao2000", 7, 2
                '---------------------------------------------------------------------
                '-------------------------Fazendo Backup Loja.mdb---------------------
                'AtualizaProcessoTela "Backup do Banco de Dados Balcao2000", 8, 1
                'BackupLoja "BackupLoja.MDB", Mid(WbancoAccess, 1, Len(WbancoAccess) - 8), WbancoDbf & "SalvaBc2000\"
                'frmAguarde.lblMensagem.Caption = "Fim do Fechamento"
                'frmAguarde.aniBackup.Visible = False
                'AtualizaProcessoTela "Backup do Banco de Dados Balcao2000", 8, 2
                '--------------------------------------------------------------------
            End If
            frmRotinasDiaria.lblProcessos.Caption = "  BOA NOITE"
            'Unload frmAguarde
            MsgBox "Fechamento finalizado com sucesso", vbInformation, "Aviso"
            frmAguarde.fraProcessos.Enabled = False
            Screen.MousePointer = 0
        Else
            frmRotinasDiaria.lblProcessos.Caption = "ESTA ROTINA JA FOI PROCESSADA"
            Exit Function
        End If
    Else
        MsgBox "ERRO NO PROCESSAMENTO DE ROTINAS DIARIAS", vbCritical, "ATENÇÃO"
        Exit Function
    End If


End Function


Sub LimpaHoraAtualizacao()
    ' frmRotinasDiaria.lblProcessos.Caption = "Hora Atualização"
    SQL = ""
    SQL = "Update HoraAtualizacao set HA_Situacao='A', " _
        & "HA_Status='A',HA_HoraInicio='00:00',HA_HoraFim='00:00' " _
        & "Where HA_Sequencia=1 "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update HoraAtualizacao set HA_Situacao='A', " _
        & "HA_Status='E',HA_HoraInicio='00:00',HA_HoraFim='00:00' " _
        & "Where HA_Sequencia > 1 "
        rdoCNLoja.Execute (SQL)
    
End Sub

Sub AtualizaLoja()
    frmRotinasDiaria.lblProcessos.Caption = "Atualizando Loja"
    SQL = ""
    SQL = "Update MovimentoCaixa set MC_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MovimentoBancario set MB_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MovimentacaoEstoque set ME_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update CTcaixa set CT_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update DivergenciaEstoque set DE_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
    
    SQL = ""
    SQL = "Update EstoqueLoja set EL_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MetadeVendas set MT_Loja = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update NfCapa set LojaOrigem = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update NfItens set LojaOrigem = '" & RSTipoControle("CT_Loja") & "' "
        rdoCNLoja.Execute (SQL)
End Sub

Sub ZeraVendaDia()
    frmRotinasDiaria.lblProcessos.Caption = "Atualizando Vendas"
    SQL = ""
    SQL = "Update Vende set VE_MargemVenda = 0, VE_TotalVenda = 0"
    rdoCNLoja.Execute (SQL)
End Sub


Sub ExtraiSequenciaNotaTransferencia()

    Dim WnovaSeqNota As Long
     
     SQL = ""
     SQL = "Select * from controle"
     Set RsDados = rdoCNLoja.OpenResultset(SQL)
     
     If Not RsDados.EOF Then
        If Wserie = "CT" Then
            WnumeroNotaDbf = 0
            WnovaSeqNota = 0
            
            WnumeroNotaDbf = RsDados("CT_SeqCT") + 1
            WnovaSeqNota = WnumeroNotaDbf
            WNfTransferencia = WnumeroNotaDbf
        
            SQL = "update controle set CT_SeqCT= " & WnovaSeqNota & ""
            rdoCNLoja.Execute (SQL)
        Else
            WnumeroNotaDbf = 0
            WnovaSeqNota = 0
            
            WnumeroNotaDbf = RsDados("CT_SeqNota") + 1
            WnovaSeqNota = WnumeroNotaDbf
            WNfTransferencia = WnumeroNotaDbf
        
            SQL = "update controle set CT_SeqNota= " & WnovaSeqNota & ""
            rdoCNLoja.Execute (SQL)
        End If
             
     End If
End Sub



Sub ProcessaListaPreco()

    SQL = "Select * from ListaPrecoCapa " _
        & "where LC_DataVigencia='" & Format(Date, "dd/mm/yyyy")

End Sub


Function CopiaNfCapa(ByVal Data As String)

    Dim rsCopiaNfCapa As rdoResultset
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "Where DataEmi='" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and TipoNota not in ('PA','R','R2') and NF > 0 " _
        & "and Serie not in('R2','RC','S1','S2') order by Nf"
        Set rsCopiaNfCapa = rdoCNLoja.OpenResultset(SQL)
    If Not rsCopiaNfCapa.EOF Then
        Do While Not rsCopiaNfCapa.EOF
            SQL = ""
            SQL = "Insert into NfCapa (NUMEROPED, DATAEMI, VENDEDOR, VLRMERCADORIA, DESCONTO, " _
                & "SUBTOTAL, LOJAORIGEM, TIPONOTA, CONDPAG, AV, CLIENTE, CODOPER, DATAPAG, PGENTRA, LOJAT, QTDITEM, PEDCLI, TM, PESOBR, PESOLQ, VALFRETE, FRETECOBR, OUTRALOJA, OUTROVEND, NF, " _
                & "TOTALNOTA, NATOPERACAO, DATAPED, BASEICMS, ALIQICMS, VLRICMS, SERIE, HORA, TOTALIPI, ECF, NUMEROSF, NOMCLI, FONECLI, CGCCLI, INSCRICLI, ENDCLI, UFCLIENTE, MUNICIPIOCLI, BAIRROCLI, " _
                & "CEPCLI, PESSOACLI, REGIAOCLI, CFOAUX, AnexoAUx, PAGINANF, ECFNF, Carimbo1, Carimbo2, Carimbo3, Carimbo4, CustoMedioLiquido, VendaLiquida, MargemContribuicao, ValorTotalCodigoZero, " _
                & "TotalNotaAlternativa, ValorMercadoriaAlternativa, SituacaoEnvio, VendedorLojaVenda, LojaVenda) " _
                & "Values (" & rsCopiaNfCapa("NUMEROPED") & ", '" & Format(rsCopiaNfCapa("DATAEMI"), "dd/mm/yyyy") & "', " & rsCopiaNfCapa("VENDEDOR") & ", " & ConverteVirgula(rsCopiaNfCapa("VLRMERCADORIA")) & ", " & ConverteVirgula(rsCopiaNfCapa("DESCONTO")) & ", " _
                & "" & ConverteVirgula(rsCopiaNfCapa("SUBTOTAL")) & ", '" & rsCopiaNfCapa("LOJAORIGEM") & "', '" & rsCopiaNfCapa("TIPONOTA") & "', '" & rsCopiaNfCapa("CONDPAG") & "', " & rsCopiaNfCapa("AV") & ", " & rsCopiaNfCapa("CLIENTE") & ", " & rsCopiaNfCapa("CODOPER") & ", '" & Format(rsCopiaNfCapa("DATAPAG"), "dd/mm/yyyy") & "', " & ConverteVirgula(rsCopiaNfCapa("PGENTRA")) & ", '" & rsCopiaNfCapa("LOJAT") & "', " _
                & "" & rsCopiaNfCapa("QTDITEM") & ", " & rsCopiaNfCapa("PEDCLI") & ", " & rsCopiaNfCapa("TM") & ", " & ConverteVirgula(rsCopiaNfCapa("PESOBR")) & ", " & ConverteVirgula(rsCopiaNfCapa("PESOLQ")) & ", " & ConverteVirgula(rsCopiaNfCapa("VALFRETE")) & ", " & ConverteVirgula(rsCopiaNfCapa("FRETECOBR")) & ", '" & rsCopiaNfCapa("OUTRALOJA") & "', " & rsCopiaNfCapa("OUTROVEND") & ", " & rsCopiaNfCapa("NF") & ", " & ConverteVirgula(rsCopiaNfCapa("TOTALNOTA")) & ", " & rsCopiaNfCapa("NATOPERACAO") & ", '" & Format(IIf(IsNull(rsCopiaNfCapa("DATAPED")), rsCopiaNfCapa("DataEmi"), rsCopiaNfCapa("DATAPED")), "dd/mm/yyyy") & "', " _
                & "" & ConverteVirgula(rsCopiaNfCapa("BASEICMS")) & ", " & ConverteVirgula(rsCopiaNfCapa("ALIQICMS")) & ", " & ConverteVirgula(rsCopiaNfCapa("VLRICMS")) & ", '" & rsCopiaNfCapa("SERIE") & "', '" & IIf(IsNull(rsCopiaNfCapa("HORA")), "00:00", rsCopiaNfCapa("HORA")) & "', " & ConverteVirgula(rsCopiaNfCapa("TOTALIPI")) & ", " & rsCopiaNfCapa("ECF") & ", " & rsCopiaNfCapa("NUMEROSF") & ", '" & rsCopiaNfCapa("NOMCLI") & "', '" & rsCopiaNfCapa("FONECLI") & "', '" & rsCopiaNfCapa("CGCCLI") & "', '" & rsCopiaNfCapa("INSCRICLI") & "', '" & rsCopiaNfCapa("ENDCLI") & "', '" & rsCopiaNfCapa("UFCLIENTE") & "', " _
                & "'" & rsCopiaNfCapa("MUNICIPIOCLI") & "', '" & rsCopiaNfCapa("BAIRROCLI") & "', '" & rsCopiaNfCapa("CEPCLI") & "', " & rsCopiaNfCapa("PESSOACLI") & ", " & rsCopiaNfCapa("REGIAOCLI") & ", '" & rsCopiaNfCapa("CFOAUX") & "', '" & rsCopiaNfCapa("AnexoAUx") & "', " & rsCopiaNfCapa("PAGINANF") & ", " & rsCopiaNfCapa("ECFNF") & ", '" & rsCopiaNfCapa("Carimbo1") & "', ' " & rsCopiaNfCapa("Carimbo2") & "', '" & rsCopiaNfCapa("Carimbo3") & "', '" & rsCopiaNfCapa("Carimbo4") & "', " & ConverteVirgula(rsCopiaNfCapa("CustoMedioLiquido")) & ", " _
                & "" & ConverteVirgula(rsCopiaNfCapa("VendaLiquida")) & ", " & ConverteVirgula(rsCopiaNfCapa("MargemContribuicao")) & ", " & ConverteVirgula(rsCopiaNfCapa("ValorTotalCodigoZero")) & ", " & ConverteVirgula(rsCopiaNfCapa("TotalNotaAlternativa")) & ", " & ConverteVirgula(rsCopiaNfCapa("ValorMercadoriaAlternativa")) & ", '" & rsCopiaNfCapa("SituacaoEnvio") & "', " & rsCopiaNfCapa("VendedorLojaVenda") & ",'" & rsCopiaNfCapa("LojaVenda") & "')"
                dbMovDia.Execute (SQL)
            rsCopiaNfCapa.MoveNext
        Loop
    End If


End Function

Function CopiaNfItens(ByVal Data As String)
    
    Dim RsCopiaNfItens As rdoResultset
    Dim wDescricaoAlternativa  As String

    SQL = ""
    SQL = "Select * from NfItens " _
        & "Where DataEmi='" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and TipoNota not in ('PA','R','R2') and NF > 0 " _
        & "and Serie not in('R2','RC','S1','S2') order by Nf"
        Set RsCopiaNfItens = rdoCNLoja.OpenResultset(SQL)
    If Not RsCopiaNfItens.EOF Then
        Do While Not RsCopiaNfItens.EOF
            If RsCopiaNfItens("DescricaoAlternativa") = "" Then
                wDescricaoAlternativa = "0"
            Else
                wDescricaoAlternativa = IIf(IsNull(RsCopiaNfItens("DescricaoAlternativa")), 0, RsCopiaNfItens("DescricaoAlternativa"))
            End If
            SQL = "Insert into NfItens (NUMEROPED, DATAEMI, REFERENCIA, QTDE, VLUNIT, VLUNIT2, VLTOTITEM, DESCRAT, ICMS, ITEM, VLIPI, DESCONTO, PLISTA, COMISSAO, VALORICMS, BCOMIS, CSPROD, LINHA, SECAO, VBUNIT, ICMPDV, CODBARRA, NF, SERIE, LOJAORIGEM, CLIENTE, VENDEDOR, ALIQIPI, TIPONOTA, REDUCAOICMS, BASEICMS, TIPOMOVIMENTACAO, DETALHEIMPRESSAO, SERIEPROD1, " _
                & "SERIEPROD2, CustoMedioLiquido, VendaLiquida, MargemContribuicao, EncargosVendaLiquida, EncargosCustoMedioLiquido, PrecoUnitAlternativa, ValorMercadoriaAlternativa, ReferenciaAlternativa, SituacaoEnvio, DescricaoAlternativa)" _
                & "Values (" & RsCopiaNfItens("NUMEROPED") & ", '" & Format(RsCopiaNfItens("DATAEMI"), "dd/mm/yyyy") & "', '" & RsCopiaNfItens("REFERENCIA") & "', " & RsCopiaNfItens("QTDE") & ", " & ConverteVirgula(RsCopiaNfItens("VLUNIT")) & ", " & ConverteVirgula(RsCopiaNfItens("VLUNIT2")) & ", " & ConverteVirgula(RsCopiaNfItens("VLTOTITEM")) & ", " _
                & "" & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("DESCRAT")), 0, RsCopiaNfItens("DESCRAT"))) & ", " & ConverteVirgula(RsCopiaNfItens("ICMS")) & ", " & RsCopiaNfItens("ITEM") & ", " & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("VLIPI")), 0, RsCopiaNfItens("VlIPI"))) & ", " & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("DESCONTO")), 0, RsCopiaNfItens("DESCONTO"))) & ", " & ConverteVirgula(RsCopiaNfItens("PLISTA")) & ", " _
                & "" & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("COMISSAO")), 0, RsCopiaNfItens("COMISSAO"))) & ", " & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("VALORICMS")), 0, RsCopiaNfItens("VALORICMS"))) & ", " & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("BCOMIS")), 0, RsCopiaNfItens("BCOMIS"))) & ", " & RsCopiaNfItens("CSPROD") & ", " & RsCopiaNfItens("LINHA") & ", " & RsCopiaNfItens("SECAO") & ", " & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("VBUNIT")), 0, RsCopiaNfItens("VBUNIT"))) & ", " _
                & "" & ConverteVirgula(IIf(IsNull(RsCopiaNfItens("ICMPDV")), 0, RsCopiaNfItens("ICMPDV"))) & ", '" & RsCopiaNfItens("CODBARRA") & "', " & RsCopiaNfItens("NF") & ", '" & RsCopiaNfItens("SERIE") & "', '" & RsCopiaNfItens("LOJAORIGEM") & "', " & RsCopiaNfItens("CLIENTE") & ", " & RsCopiaNfItens("VENDEDOR") & ", " & ConverteVirgula(RsCopiaNfItens("ALIQIPI")) & ", '" & RsCopiaNfItens("TIPONOTA") & "', " & ConverteVirgula(RsCopiaNfItens("REDUCAOICMS")) & ", " _
                & "" & ConverteVirgula(RsCopiaNfItens("BASEICMS")) & ", " & RsCopiaNfItens("TIPOMOVIMENTACAO") & ", '" & RsCopiaNfItens("DETALHEIMPRESSAO") & "', '" & RsCopiaNfItens("SERIEPROD1") & "', '" & RsCopiaNfItens("SERIEPROD2") & "', " & ConverteVirgula(RsCopiaNfItens("CustoMedioLiquido")) & ", " & ConverteVirgula(RsCopiaNfItens("VendaLiquida")) & " , " _
                & "" & ConverteVirgula(RsCopiaNfItens("MargemContribuicao")) & ", " & ConverteVirgula(RsCopiaNfItens("EncargosVendaLiquida")) & ", " & ConverteVirgula(RsCopiaNfItens("EncargosCustoMedioLiquido")) & ", " & ConverteVirgula(RsCopiaNfItens("PrecoUnitAlternativa")) & ", " & ConverteVirgula(RsCopiaNfItens("ValorMercadoriaAlternativa")) & ", '" & RsCopiaNfItens("ReferenciaAlternativa") & "', '" & RsCopiaNfItens("SituacaoEnvio") & "', '" & wDescricaoAlternativa & "')"
                dbMovDia.Execute (SQL)
            RsCopiaNfItens.MoveNext
        Loop
    End If
    
End Function


Function Cliptografia(ByRef ValorClipt As String)

    Dim Ret As String
    Dim CharLido As String
    Dim Maximo As Long
    Dim i As Long
    
    Ret = ""
    Maximo = Len(ValorClipt)
    
    For i = 1 To Maximo
        CharLido = UCase(Mid(ValorClipt, i, 1))
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

Function CriaRelatorioCodigoZero(ByVal Data As String, ByVal Loja As String)

    Dim RsCriaRelCodZero As rdoResultset
    Dim wValorTotalNota As Double
    Dim wNumeroLinha As Integer
    Dim wTotal As Double
    Dim wTotalGeral As Double
    Dim wUltimaSerie As String
    Dim Serie As String
    
    frmAguarde.lblMensagem.Caption = "Aguarde, Emitindo Codigo Zero"
    
    'SM = RR
    '00 = RO
    '
    '-----------------------Set a Impressora--------------------------------
    '
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora do sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    '***********************************************************************
    
    
    SQL = ""
    SQL = "Select NF,TotalNota,LojaOrigem,Serie,TotalNotaAlternativa from NfCapa " _
        & "Where LojaOrigem='" & Loja & "' " _
        & "and DataEmi='" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and Serie in ('00','SM') " _
        & "and TipoNota='V' " _
        & "order by Serie,NF"
    Set RsCriaRelCodZero = rdoCNLoja.OpenResultset(SQL)
    If Not RsCriaRelCodZero.EOF Then
        'Pagina = 0
        wUltimaSerie = RsCriaRelCodZero("Serie")
        CabecalhoRelCodZero Data, Loja
        wValorTotalNota = 0
        wTotal = 0
        wTotalGeral = 0
        Do While Not RsCriaRelCodZero.EOF
                        
            If Trim(wUltimaSerie) <> Trim(RsCriaRelCodZero("Serie")) Then
                Printer.Print ""
                Printer.FontBold = True
                Printer.Print Space(18) & "TOTAL       RO    " & Right(Space(14) & Format(wTotal, "0.00"), 14)
                wTotal = 0
                Printer.FontBold = False
                Printer.Print
            End If
            If Trim(RsCriaRelCodZero("Serie")) = "SM" Then
                Serie = "RO"
                wValorTotalNota = RsCriaRelCodZero("TotalNota") - RsCriaRelCodZero("TotalNotaAlternativa")
                wTotal = wTotal + wValorTotalNota
            Else
                Serie = "RR"
                wValorTotalNota = RsCriaRelCodZero("TotalNota")
                wTotal = wTotal + wValorTotalNota
            End If
            wNumeroLinha = wNumeroLinha + 1
            If wNumeroLinha <= 50 Then
                Printer.Print Space(18) & Left(RsCriaRelCodZero("NF") & Space(12), 12) _
                    & Left(Serie & Space(10), 10) _
                    & Right(Space(10) & Format(wValorTotalNota, "0.00"), 10)
            Else
                Printer.NewPage
                CabecalhoRelCodZero Data, Loja
                Printer.Print Space(18) & Left(RsCriaRelCodZero("NF") & Space(12), 12) _
                    & Left(Serie & Space(10), 10) _
                    & Right(Space(10) & Format(wValorTotalNota, "0.00"), 10)
                wNumeroLinha = 0
            End If
            wTotalGeral = wTotalGeral + wValorTotalNota
            wUltimaSerie = RsCriaRelCodZero("Serie")
            RsCriaRelCodZero.MoveNext
        Loop
        If wTotal > 0 Then
            Printer.Print ""
            Printer.FontBold = True
            Printer.Print Space(18) & "TOTAL       RR    " & Right(Space(14) & Format(wTotal, "0.00"), 14)
            wTotal = 0
        End If
        Printer.Print ""
        Printer.FontBold = True
        Printer.Print Space(18) & "TOTAL GERAL       " & Right(Space(14) & Format(wTotalGeral, "0.00"), 14)
        Printer.EndDoc
    'Else
        'MsgBox "Nao Exite Movimento de Codigo zero / 00 para este dia", vbInformation, "Aviso"
    End If
End Function

Function CabecalhoRelCodZero(ByVal Data As String, ByVal Loja As String)

    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    'Pagina = Pagina + 1
    
    Printer.FontName = "ARIAL"
    Printer.FontBold = False
    Printer.FontSize = 9
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 5
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Line (8, 10)-(199, 10)
    Printer.FontBold = True
    'Printer.Print Tab(10); "DE MEO COMERCIAL IMP. LTDA"; Tab(118); "PÁGINA: "; 1 'Pagina
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(27); "R E L A T O R I O     D E     E S T A D O         " & Format(Data, "dd/mm/yyyy") & "   LOJA  " & Loja; Tab(118); Format(Date, "DD/mm/YYYY")
    Printer.Print ""
    Printer.Print ""
    
    
    Printer.CurrentY = Printer.CurrentY - 1
    
    '(COLUNA,LINHA)
'    Printer.Line (8, 10)-(8, 37)
'    Printer.Line (199, 10)-(199, 37)
    Printer.Line (8, 270)-(199, 270)
    Printer.Line (8, 26)-(199, 26)
    
    Printer.Print ""
    Printer.Print Tab(27); "DC                          ESTADO                              FRETE  "
    
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.FontName = "COURIER NEW"
    Printer.Print
End Function

Function PegaNumeroCFControle() As Double

    Dim rsPegaNumeroCF As rdoResultset
    
    SQL = ""
    SQL = "Select (CT_UltimoCupom + 1) as NumeroCupom from ControleECF " _
        & "where CT_Ecf=" & Val(glb_ECF) & ""
        Set rsPegaNumeroCF = rdoCNLoja.OpenResultset(SQL)
    If Not rsPegaNumeroCF.EOF Then
        SQL = ""
        SQL = "Update ControleECF set CT_UltimoCupom=" & rsPegaNumeroCF("NumeroCupom") & " " _
            & "where CT_Ecf=" & Val(glb_ECF) & ""
            rdoCNLoja.Execute (SQL)
    
        PegaNumeroCFControle = rsPegaNumeroCF("NumeroCupom")
    End If

End Function

Sub AtualizaNumeroCupom()

    SQL = ""
    SQL = "Update controleEcf set ct_ultimocupom= CT_UltimoCupom + 1 " _
        & "where CT_Ecf=" & Val(glb_ECF) & ""
           rdoCNLoja.Execute (SQL)

End Sub


Function VerificaControleEcf(ByVal NumeroECF As Integer, ByVal Loja As String)
    
    Dim rsVerificaControleEcf As rdoResultset
    Dim rsVerificaCT As rdoResultset
    Dim rsOperador As rdoResultset

    SQL = ""
    SQL = "Select * from ControleEcf " _
        & "where CT_Ecf=" & NumeroECF
        Set rsVerificaControleEcf = rdoCNLoja.OpenResultset(SQL)
                                    
    If rsVerificaControleEcf.EOF Then
        SQL = ""
        SQL = "Insert into ControleEcf (CT_Ecf,CT_QtdeEcf,CT_UltimoCupom,CT_SituacaoCupomFiscal,CT_SituacaoCaixa,CT_PegaPedido) " _
            & "Values(" & NumeroECF & ",0,0,'F','F','N') "
            rdoCNLoja.Execute (SQL)
        SQL = ""
        SQL = "Select * from CTCaixa " _
            & "where CT_NumeroEcf=" & glb_ECF & ""
            Set rsVerificaCT = rdoCNLoja.OpenResultset(SQL)
        If rsVerificaCT.EOF Then
            SQL = ""
            SQL = "Select CT_Operador from CTCaixa where CT_Operador > 0 "
                Set rsOperador = rdoCNLoja.OpenResultset(SQL)
            If Not rsOperador.EOF Then
                SQL = "insert into CtCaixa (CT_NumeroECF,CT_Loja,CT_Data,CT_HoraInicial,CT_HoraFinal,CT_Operacoes,CT_Controle,CT_Situacao,CT_Operador) " _
                    & "Values(" & glb_ECF & ",'" & Loja & "','" & Format(Date, "dd/mm/yyyy") & "', " _
                    & "'" & Format(Time, "hh:mm") & "', '" & Format(Time, "hh:mm") & "',0,0,'P'," & rsOperador("CT_Operador") & ")"
                    rdoCNLoja.Execute (SQL)
            End If
        End If
    Else
        If rsVerificaControleEcf("CT_SituacaoCupomFiscal") = "A" Then
     '*** Desabilitado 12/2009       Retorno = Bematech_FI_CancelaCupom()
     '*** Desabilitado 12/2009       MsgBox "Exite Cupom Aberto", vbInformation, "Aviso"
      '*** Desabilitado 12/2009      Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      '*** Desabilitado 12/2009      If Retorno = 1 Then
      '*** Desabilitado 12/2009          MsgBox "Atenção cupom sendo cancelado", vbInformation, "Problemas com Cupom"
      '*** Desabilitado 12/2009          Call AtualizaNumeroCupom
      '*** Desabilitado 12/2009          SQL = ""
      '*** Desabilitado 12/2009          SQL = "Update ControleEcf set CT_SituacaoCupomFiscal='F' "
      '*** Desabilitado 12/2009              rdoCNLoja.Execute (SQL)
      '*** Desabilitado 12/2009      End If
        End If
    End If
    
End Function


Function ComparaDataVersao(ByVal Versao As String, ByVal VersaoNova As String) As Boolean
    
    Dim DataExeNovo As Date
    Dim DataExe As Date
    
    On Error Resume Next
    DataExeNovo = Format(FileDateTime(VersaoNova), "dd/mm/yyyy hh:mm:ss")
    DataExe = Format(FileDateTime(Versao), "dd/mm/yyyy hh:mm:ss")
    
    If DataExeNovo = 0 Or DataExe = 0 Then
        MsgBox "Erro na troca de versão tente mais tarde", vbCritical, "Erro"
        ComparaDataVersao = False
    Else
        If DataExeNovo > DataExe Then
            ComparaDataVersao = True
        Else
            ComparaDataVersao = False
        End If
    End If
    

End Function


Function VerificaCaixaAberto(ByVal Situacao As String)
    Dim rsVerCaixa As rdoResultset
    Dim rsPegaPedido As rdoResultset
    
    SQL = ""
    SQL = "Update ControleEcf set CT_SituacaoCaixa='" & Situacao & "' " _
        & "where CT_ECF=" & Val(glb_ECF) & ""
        rdoCNLoja.Execute (SQL)
    
    If Situacao = "A" Then
        SQL = ""
        SQL = "Select CT_PegaPedido from ControleECF " _
            & "where CT_SituacaoCaixa='A' and CT_PegaPedido='S'"
            Set rsVerCaixa = rdoCNLoja.OpenResultset(SQL)
        If rsVerCaixa.EOF Then
            SQL = ""
            SQL = "Update ControleEcf set CT_PegaPedido='S' " _
                & "where CT_ECF=" & glb_ECF
            rdoCNLoja.Execute (SQL)
        End If
    ElseIf Situacao = "F" Then 'Fecha O Caixa que Pega Pedido
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido='N' " _
            & "where CT_ECF=" & Val(glb_ECF) & ""
            rdoCNLoja.Execute (SQL)
        
        SQL = ""
        SQL = "Select CT_PegaPedido from ControleECF " _
            & "where CT_SituacaoCaixa='A' and CT_PegaPedido='S'"
            Set rsVerCaixa = rdoCNLoja.OpenResultset(SQL)
        If rsVerCaixa.EOF Then
            SQL = "Select CT_PegaPedido,CT_ECF from ControleECF " _
                & "where CT_SituacaoCaixa='A' and CT_ECf<>" & Val(glb_ECF) & ""
                Set rsPegaPedido = rdoCNLoja.OpenResultset(SQL)
            If Not rsPegaPedido.EOF Then
                SQL = ""
                SQL = "Update ControleEcf set CT_PegaPedido='S' " _
                    & "where CT_ECF=" & rsPegaPedido("CT_ECF")
                rdoCNLoja.Execute (SQL)
            End If
        End If
    ElseIf Situacao = "S" Then  'Muda O Caixa que pega os pedidos
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido = 'N' "
            rdoCNLoja.Execute (SQL)
        
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido='S',CT_SituacaoCaixa='A' " _
            & "where CT_ecf=" & Val(glb_ECF) & ""
            rdoCNLoja.Execute (SQL)
    End If
End Function


Function CriaNotaCredito(ByVal Nf As Double, ByVal Serie As String, ByVal NfDev As Double, ByVal SerieDev As String, ByVal DataDev As String, ByVal ValorNotaCredito As Double, ByVal NotaCredito As Double, ByVal ReImpressao As Boolean)
    Dim rsDadosNfCapa As rdoResultset
    Dim rsVerLoja As rdoResultset
    Dim rsDataEmiDevol As rdoResultset
    Dim Linha1 As String
    Dim wTotalNota As Double
    Dim wValorExtenso As String
    Dim wDataEmiDevolucao As Date
    
    'Printer.Line (0, 10)-(199, 10)
    'Printer.Line (0, 10)-(0, 100)
    'Printer.Line (199, 10)-(199, 100)
    
    'Printer.EndDoc
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "where Nf=" & NfDev & " " _
        & "and Serie='" & SerieDev & "'"
    Set rsDadosNfCapa = rdoCNLoja.OpenResultset(SQL)
    If Not rsDadosNfCapa.EOF Then
        If ReImpressao = True Then
            SQL = ""
            SQL = "Select DataEmi From NfCapa Where NF = " & rsDadosNfCapa("NfDevolucao") & " and " _
                & "Serie = '" & rsDadosNfCapa("SerieDevolucao") & "' and Lojaorigem = '" & AchaLojaControle & "'"
            Set rsDataEmiDevol = rdoCNLoja.OpenResultset(SQL)
            
            If rsDataEmiDevol.EOF Then
                rsDataEmiDevol.Close
                
                MsgBox "Irei conectar na Retaguarda para localizar a nota." & Chr(10) & "Pois a mesma não foi encontrado no BANCO LOCAL", vbInformation + vbOKOnly
                
                Conexao.Close
                
                If ConectaODBC(Conexao, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = True Then
                    SQL = ""
                    SQL = "Select DataEmi From NfCapa Where NF = " & rsDadosNfCapa("NfDevolucao") & " and " _
                        & "Serie = '" & rsDadosNfCapa("SerieDevolucao") & "' and Lojaorigem = '" & AchaLojaControle & "'"
                    Set rsDataEmiDevol = Conexao.OpenResultset(SQL)
                    
                    If rsDataEmiDevol.EOF Then
                        wDataEmiDevolucao = Date
                    Else
                        wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                    End If
                    rsDataEmiDevol.Close
                    Conexao.Close
                End If
            Else
                wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                rsDataEmiDevol.Close
            End If
        Else
            wDataEmiDevolucao = frmDevolucaoVenda.mskDTEmi
        End If
            
        SQL = ""
        SQL = "Select CT_Loja,CT_Razao,CT_NCredito,CT_NovaRazao,Lojas.* from Controle,Lojas " _
            & "where LO_Loja=CT_Loja"
        Set rsVerLoja = rdoCNLoja.OpenResultset(SQL)
        If Not rsVerLoja.EOF Then
            If Serie = "SM" Then
                wTotalNota = rsDadosNfCapa("TotalNotaAlternativa")
            Else
                wTotalNota = ValorNotaCredito
            End If
            'wValorExtenso = PassaExtenso(wTotalNota)
            For i = 1 To 4
                'Printer.Line Step(2, 10)-(10, 100), , B
                'Printer.Line (2, 10)-(199, 10)
                Printer.ScaleMode = vbMillimeters
                Printer.FontName = "Romam"
                Printer.FontSize = 9
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                
                If Len(Trim(rsVerLoja("CT_NovaRazao"))) < 1 Then
                    If ReImpressao = False Then
                        Printer.Print Space(2) & rsVerLoja("CT_Razao")
                    Else
                        Printer.Print Space(2) & rsVerLoja("CT_Razao") & Space(84) & "RE-IMPRESSAO"
                    End If
                Else
                    If ReImpressao = False Then
                        Printer.Print Space(2) & Trim(Mid(rsVerLoja("CT_NovaRazao"), 20, Len(rsVerLoja("CT_NovaRazao"))))
                    Else
                        Printer.Print Space(2) & Trim(Mid(rsVerLoja("CT_NovaRazao"), 20, Len(rsVerLoja("CT_NovaRazao")))) & Space(84) & "RE-IMPRESSAO"
                    End If
                End If
                
                Printer.Print Space(2) & Left(rsVerLoja("LO_Endereco") & Space(30), 30) _
                    & "    -    " & rsVerLoja("LO_Cep") & "   -   " & rsVerLoja("LO_Municipio") _
                    & Right(Space(72) & "NOTA DE CREDITO", 72)
                Printer.Print Space(2) & "FONE : " & "(" & Right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" _
                        & Left(rsVerLoja("LO_Telefone") & Space(10), 10) & " -  " _
                        & "FAX : " & "(" & Right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" & Left(rsVerLoja("LO_Telefone") & Space(10), 10)
                Printer.Print Space(2) & "C.G.C : " & Left(rsVerLoja("LO_CGC") & Space(25), 25) & "INSCR.EST. : " & rsVerLoja("LO_InscricaoEstadual")
                Printer.Print Space(140) & "NUM.  " & Right(String(9, "0") & NotaCredito, 9) & Right(Space(10) & i & "a.VIA", 10)
                Printer.Print Space(2) & "A"
                Printer.Print Space(2) & rsDadosNfCapa("NomCli")
                Printer.Print Space(2) & Left(rsDadosNfCapa("EndCli") & Space(130), 130) & Left("DATA : " & rsDadosNfCapa("DataEmi") & Space(18), 18)
                Printer.Print Space(2) & rsDadosNfCapa("MunicipioCli") & "  -   " & rsDadosNfCapa("UfCliente")
                Printer.Print Space(2) & "FONE : " & rsDadosNfCapa("FoneCli")
                Printer.Print Space(2) & "EFETUAMOS NESTA DATA EM SUA CONTA CORRENTE O SEGUINTE LANÇAMENTO:"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print Space(40) & "HISTORICO" & Space(40) & "| DEBITO" & Space(30) & "| CREDITO"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(2) & "PELO RECEBIMENTO DA MERCADORIA EM DEVOLUÇÃO"
                Printer.Print Space(2) & "CONFORME NF " & NfDev & " SERIE " & SerieDev & " DE " & rsDadosNfCapa("DataEmi")
                Printer.Print Space(2) & "NO VALOR DE R$          " & Format(wTotalNota, "###,###,###0.00")
                Printer.Print
                Printer.Print Space(2) & "REFERENTE NF " & Nf & " - " & Serie & " DE " & wDataEmiDevolucao
                Printer.Print Space(2) & "DA LOJA " & rsDadosNfCapa("LojaOrigem") & Space(140) & Format(wTotalNota, "###,###,###0.00")
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(140) & "ATENCIOSAMENTE"
                Printer.Print
                Printer.Print Space(120) & "_______________________________________"
                If Len(Trim(rsVerLoja("CT_NovaRazao"))) < 1 Then
                    Printer.Print Space(120) & rsVerLoja("CT_Razao")
                Else
                    Printer.Print Space(120) & Trim(Mid(rsVerLoja("CT_NovaRazao"), 20, Len(rsVerLoja("CT_NovaRazao"))))
                End If
                Printer.Print
                Printer.Print
                If i = 2 Then
                    Printer.NewPage
                End If
            Next i
            Printer.EndDoc
        End If
    End If


End Function


Function CriaNotaCreditoBola(ByVal Nf As Double, ByVal Serie As String, ByVal NfDev As Double, ByVal SerieDev As String, ByVal DataDev As String, ByVal ValorNotaCredito As Double, ByVal NotaCredito As Double, ByVal ReImpressao As Boolean)
    Dim rsDadosNfCapa As rdoResultset
    Dim rsVerLoja As rdoResultset
    Dim rsDataEmiDevol As rdoResultset
    Dim Linha1 As String
    Dim wTotalNota As Double
    Dim wValorExtenso As String
    Dim wDataEmiDevolucao As Date
    'Printer.Line (0, 10)-(199, 10)
    'Printer.Line (0, 10)-(0, 100)
    'Printer.Line (199, 10)-(199, 100)
    
    'Printer.EndDoc
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "where Nf=" & NfDev & " " _
        & "and Serie='" & Serie & "' and Lojaorigem = '" & AchaLojaControle & "'"
    Set rsDadosNfCapa = rdoCNLoja.OpenResultset(SQL)
    If Not rsDadosNfCapa.EOF Then
        If ReImpressao = True Then
            SQL = ""
            SQL = "Select DataEmi From NfCapa Where NF = " & rsDadosNfCapa("NfDevolucao") & " and " _
                & "Serie = '" & rsDadosNfCapa("SerieDevolucao") & "' and Lojaorigem = '" & AchaLojaControle & "'"
            Set rsDataEmiDevol = rdoCNLoja.OpenResultset(SQL)
            
            If rsDataEmiDevol.EOF Then
                rsDataEmiDevol.Close
                
                MsgBox "Irei conectar na Retaguarda para localizar a nota." & Chr(10) & "Pois a mesma não foi encontrado no BANCO LOCAL", vbInformation + vbOKOnly
                
                If ConectaODBC(Conexao, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = True Then
                    SQL = ""
                    SQL = "Select DataEmi From NfCapa Where NF = " & rsDadosNfCapa("NfDevolucao") & " and " _
                        & "Serie = '" & rsDadosNfCapa("SerieDevolucao") & "' and Lojaorigem = '" & AchaLojaControle & "'"
                    Set rsDataEmiDevol = Conexao.OpenResultset(SQL)
                    
                    If rsDataEmiDevol.EOF Then
                        wDataEmiDevolucao = Date
                    Else
                        wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                    End If
                    rsDataEmiDevol.Close
                    Conexao.Close
                End If
            Else
                wDataEmiDevolucao = rsDataEmiDevol("DataEmi")
                rsDataEmiDevol.Close
            End If
        Else
            wDataEmiDevolucao = frmDevolucaoVenda.mskDTEmi
        End If
        
        SQL = ""
        SQL = "Select CT_Loja,CT_Razao,CT_NCredito,Lojas.* from Controle,Lojas " _
            & "where LO_Loja=CT_Loja"
        Set rsVerLoja = rdoCNLoja.OpenResultset(SQL)
        If Not rsVerLoja.EOF Then
            If Serie = "SM" Then
                wTotalNota = rsDadosNfCapa("TotalNota") - rsDadosNfCapa("TotalNotaAlternativa")
            Else
                wTotalNota = ValorNotaCredito
            End If
            'wValorExtenso = PassaExtenso(wTotalNota)
            For i = 1 To 2
                'Printer.Line Step(2, 10)-(10, 100), , B
                'Printer.Line (2, 10)-(199, 10)
                Printer.ScaleMode = vbMillimeters
                Printer.FontName = "Romam"
                Printer.FontSize = 9
                Printer.Print Space(2) & " CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO CO  "
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                
                If ReImpressao = False Then
                    Printer.Print Space(2) & "LOJA "; rsVerLoja("CT_Loja")
                Else
                    Printer.Print Space(2) & "LOJA "; rsVerLoja("CT_Loja") & Space(123) & "RE-IMPRESSAO"
                End If
                
                Printer.Print Space(140) & "NUM.  " & Right(String(9, "0") & NotaCredito, 9) & Right(Space(10) & i & "a.VIA", 10)
                Printer.Print Space(2) & "A"
                Printer.Print Space(2) & rsDadosNfCapa("NomCli")
                Printer.Print Space(2) & Left(rsDadosNfCapa("EndCli") & Space(130), 130) & Left("DATA : " & rsDadosNfCapa("DataEmi") & Space(18), 18)
                Printer.Print Space(2) & rsDadosNfCapa("MunicipioCli") & "  -   " & rsDadosNfCapa("UfCliente")
                Printer.Print Space(2) & "FONE : " & rsDadosNfCapa("FoneCli")
                Printer.Print Space(2) & "EFETUAMOS NESTA DATA EM SUA CONTA CORRENTE O SEGUINTE LANÇAMENTO:"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print Space(40) & "HISTORICO" & Space(40) & "| DEBITO" & Space(30) & "| CREDITO"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(2) & "PELO RECEBIMENTO DA MERCADORIA EM DEVOLUÇÃO"
                Printer.Print Space(2) & "CONFORME NF " & NfDev & " SERIE " & SerieDev & " DE " & rsDadosNfCapa("DataEmi")
                Printer.Print Space(2) & "NO VALOR DE R$          " & Format(wTotalNota, "###,###,###0.00")
                Printer.Print
                Printer.Print Space(2) & "REFERENTE NF " & Nf & " - " & Serie & " DE " & wDataEmiDevolucao
                Printer.Print Space(2) & "DA LOJA " & rsDadosNfCapa("LojaOrigem") & Space(140) & Format(wTotalNota, "###,###,###0.00")
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(130) & "ATENCIOSAMENTE"
                Printer.Print
                Printer.Print Space(110) & "_______________________________________"
                Printer.Print
                Printer.Print
            Next i
            Printer.EndDoc
        End If
    End If


End Function


Function ExtraiNumeroNotaCredito() As Double
    Dim rsNumeroNotaCredito As rdoResultset
    
        SQL = ""
        SQL = "Select (CT_NCredito+1) as NotaCredito from Controle"
            Set rsNumeroNotaCredito = rdoCNLoja.OpenResultset(SQL)
        If Not rsNumeroNotaCredito.EOF Then
            ExtraiNumeroNotaCredito = rsNumeroNotaCredito("NotaCredito")
            SQL = ""
            SQL = "Update Controle set CT_NCredito=" & rsNumeroNotaCredito("NotaCredito")
            rdoCNLoja.Execute (SQL)
        End If
    
End Function

Function AchaLojaControle() As String
    
    Dim ControleLoja As rdoResultset
    
    Set ControleLoja = rdoCNLoja.OpenResultset("Select CT_Loja from Controle")
       
    AchaLojaControle = ControleLoja("CT_Loja")
       
    ControleLoja.Close
   
End Function


Public Function ExtraiSeqPedidoDbfDev() As Double

     Dim WNovoSeqPed As Long
     Dim rdoNumPed As rdoResultset
     
'        WnumeroPedidoDbf = 0
'        WNovoSeqPed = 0
'        WnumeroPed = 0
'
'        Set DBFBanco = Workspaces(0).OpenDatabase(WbancoDbf, False, False, "DBase IV")
'
'        Set RsDadosDbf = DBFBanco.OpenRecordset("Select * from controle")
'
'        WnumeroPedidoDbf = RsDadosDbf("NumPed") + 1
'        WNovoSeqPed = WnumeroPedidoDbf
'        WnumeroPed = WnumeroPedidoDbf
'
'        BeginTrans
'
'        SQL = "update controle.dbf set NumPed= " & WNovoSeqPed & ""
'        DBFBanco.Execute (SQL)
'
'        CommitTrans
'
'        DBFBanco.Close

    SQL = ""
    SQL = "Select CT_NumPed as Pedido from Controle "
        Set rdoNumPed = rdoCNLoja.OpenResultset(SQL)
    If Not rdoNumPed.EOF Then
        WnumeroPed = rdoNumPed("Pedido")
        SQL = ""
        SQL = "update Controle set CT_NumPed=CT_NumPed + 1 "
            rdoCNLoja.Execute (SQL)
        ExtraiSeqPedidoDbfDev = WnumeroPed
    End If

End Function


Function LiberaSenha(ByVal Usuario As String, ByVal Senha As String) As Boolean
    Dim rsLiberaSenha As rdoResultset
    
    SQL = "Select Us_Usuario from Usuario " _
        & "where US_Usuario='" & Usuario & "' " _
        & "and US_Senha='" & Senha & "' " _
        & "and Us_TipoUsuario in (2,4)"
    Set rsLiberaSenha = rdoCNLoja.OpenResultset(SQL)
    If Not rsLiberaSenha.EOF Then
        LiberaSenha = True
    Else
        LiberaSenha = False
    End If
    
End Function


Function AtualizaProcessoFechamento(ByRef Tabela As String, ByRef Campo As String, ByVal Processo As String)

    SQL = ""
    SQL = "Update " & Tabela & " set " & Campo & "='" & Processo & "'"
        rdoCNLoja.Execute (SQL)

End Function

Function VerificaItensPedido(ByVal Pedido As Double, ByVal TipoNota As String) As Boolean
    Dim rsVerItensPed As rdoResultset
    Dim rsVerProdPed As rdoResultset
    Dim wVerProduto  As Boolean
    
    VerificaItensPedido = True
    SQL = ""
    SQL = "Select DISTINCT(Referencia) from NfItens " _
        & "where NumeroPed = " & Pedido & " " _
        & "and TipoNota ='" & TipoNota & "' "
    Set rsVerItensPed = rdoCNLoja.OpenResultset(SQL)
    Do While Not rsVerItensPed.EOF
        SQL = ""
        SQL = "Select PR_Referencia from Produto " _
            & "where PR_referencia='" & rsVerItensPed("Referencia") & "' "
            Set rsVerProdPed = rdoCNLoja.OpenResultset(SQL)
        If rsVerProdPed.EOF Then
            VerificaItensPedido = False
            If MsgBox("Referencia " & Trim(rsVerItensPed("Referencia")) & "  do pedido " & Val(Pedido) & " não encontrada, " & Chr(10) & " Deseja cadastrar essa referencia agora", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                Screen.MousePointer = 0
                On Error Resume Next
                GravaProduto rsVerItensPed("Referencia")
                If Err.Number = 0 Then
                    VerificaItensPedido = True
                End If
            Else
                MsgBox "Você só podera concluir esse pedido depois que gravar esta referencia", vbCritical, "Atenção"
            End If
        End If
        rsVerItensPed.MoveNext
    Loop

End Function


Function BackupLoja(ByVal NomeArquivo As String, ByVal Caminho As String, ByVal CaminhoSalva As String)
    Dim dbBackup As Database
    Dim rsBkEstoqueLoja As rdoResultset
    Dim rsBackupProduto As rdoResultset
    Dim rsBackupControle As rdoResultset
    Dim rsBackupNfItens As rdoResultset
    Dim rsBackupNfCapa As rdoResultset
    Dim rsBackupMovCaixa As rdoResultset
    Dim NomeSalva As String
    
    frmAguarde.lblMensagem.Caption = "Aguarde, Fazendo Backup"
    frmAguarde.PrbContador.Visible = False
    frmAguarde.aniBackup.Visible = True
    frmAguarde.aniBackup.Play
    
    '
    '------------------------Backup EstoqueLoja
    '
    On Error Resume Next
    NomeSalva = "BackupLoja" & Format(Date, "ddmmyy") & ".mdb"
    DBLoja.Close
    FileCopy Caminho & "Loja.mdb", CaminhoSalva & NomeSalva
    If Err.Number = 53 Then
        MsgBox "Não foi possivel encontrar a pasta " & Caminho, vbCritical, "Erro Fazendo Backup"
    End If
    Set DBLoja = OpenDatabase(Caminho & "Loja.MDB")
    
    
End Function


Function CompactaBanco(ByVal Tabela As String, ByVal Condicao As String)

    SQL = ""
    SQL = "Delete * from " & Tabela & " " _
        & "where " & Condicao & ""
        rdoCNLoja.Execute (SQL)

End Function


Function AtualizaProdutoDBF()
    Dim rsComparaProdDBF As rdoResultset
    Dim rsCompProduto As rdoResultset
    Dim Contador As Integer
    
    
    
    SQL = ""
    SQL = "Select * from Produ order by OrCodPro"
        Set rsComparaProdDBF = DBFBanco.OpenRecordset(SQL)
    If Not rsComparaProdDBF.EOF Then
        rsComparaProdDBF.MoveLast
        frmAguarde.PrbContador.Visible = True
        frmAguarde.Refresh
        frmAguarde.PrbContador.Max = rsComparaProdDBF.AbsolutePosition
        rsComparaProdDBF.MoveFirst
        Contador = 0
        Do While Not rsComparaProdDBF.EOF
            Contador = Contador + 1
            frmAguarde.PrbContador.Value = Contador
            SQL = ""
            SQL = "Select PR_Referencia from Produto " _
                & "where PR_Referencia='" & rsComparaProdDBF("OrCodPro") & "'"
            Set rsCompProduto = rdoCNLoja.OpenResultset(SQL)
            If rsCompProduto.EOF Then
                SQL = "insert into ProduLjDBF (DESCRICAO,ORCODPRO,ALIQIPI,CODIPI,CONTROLE,TRIBUTO,VENVAR1,CLASSFISC,UNIDADE, " _
                    & "PRECUS1,PROMO,BCOMIS,CSPROD,LINHA,SECAO,FORNECEDOR,TIPO,PESO,PAG,SUBTRIBUT,ICMPDV,CODBARRA,SITUACAO) " _
                    & "Values('" & rsComparaProdDBF("Descricao") & "', '" & rsComparaProdDBF("OrCodPro") & "', " & rsComparaProdDBF("AliqIPI") & ", " & rsComparaProdDBF("CodIPI") & ", " & rsComparaProdDBF("controle") & ", " _
                    & "" & rsComparaProdDBF("Tributo") & ", " & ConverteVirgula(rsComparaProdDBF("VenVar1")) & ", '" & rsComparaProdDBF("ClassFisc") & "', '" & rsComparaProdDBF("Unidade") & "', " & ConverteVirgula(Format(rsComparaProdDBF("PreCus1"), "0.00")) & ", " _
                    & "1, 0, " & rsComparaProdDBF("CsProd") & ", " & rsComparaProdDBF("Linha") & ", " & rsComparaProdDBF("Secao") & ", " & rsComparaProdDBF("Fornecedor") & ", '" & rsComparaProdDBF("Tipo") & "', " & ConverteVirgula(Format(rsComparaProdDBF("Peso"), "0.000")) & ", " & rsComparaProdDBF("Pag") & ", '" & IIf(IsNull(rsComparaProdDBF("SubTribut")), "N", rsComparaProdDBF("SubTribut")) & "', " _
                    & "" & ConverteVirgula(Format(rsComparaProdDBF("IcmPdv"), "0.00")) & ", '" & rsComparaProdDBF("CodBarra") & "', 'E')"
                    rdoCNLoja.Execute (SQL)
            End If
            rsComparaProdDBF.MoveNext
        Loop
    End If
    
End Function

Function AtualizaProdutoBc2000()
    Dim RsPegaDadosProduto As rdoResultset
    Dim RsComparaEstoque As rdoResultset
    Dim wReferenciaProduto As String * 7
    Dim wFornecedorProduto As String * 4
    Dim Contador As Integer
    Dim LojaControle As String * 5
    
    
    LojaControle = AchaLojaControle
    
    'SQL = ""
    'SQL = "Delete * from EstoqueLoja " _
        & "where EL_Referencia not in (Select PR_Referencia from Produto)"
    'rdocnloja.Execute (SQL)
    
    SQL = ""
    SQL = "Delete  * from EstoqueLoja " _
        & "where EL_Referencia in (Select PR_Referencia from Produto where PR_Situacao='E')"
    rdoCNLoja.Execute (SQL)
    
    SQL = ""
    SQL = "Delete * from Produto where PR_Situacao='E'"
        rdoCNLoja.Execute (SQL)
    
    
    SQL = ""
    SQL = "Select PR_Referencia,PR_CodigoFornecedor from Produto"
        Set RsPegaDadosProduto = rdoCNLoja.OpenResultset(SQL)
            
    If Not RsPegaDadosProduto.EOF Then
        RsPegaDadosProduto.MoveLast
        frmAguarde.PrbContador.Visible = True
        frmAguarde.Refresh
        frmAguarde.PrbContador.Value = 0
        frmAguarde.PrbContador.Max = RsPegaDadosProduto.AbsolutePosition
        RsPegaDadosProduto.MoveFirst
        Do While Not RsPegaDadosProduto.EOF
            Contador = Contador + 1
            frmAguarde.PrbContador.Value = Contador
            wReferenciaProduto = IIf(IsNull(RsPegaDadosProduto("PR_Referencia")), "0", RsPegaDadosProduto("PR_Referencia"))
            wFornecedorProduto = IIf(IsNull(RsPegaDadosProduto("PR_CodigoFornecedor")), 0, RsPegaDadosProduto("PR_CodigoFornecedor"))
            SQL = ""
            SQL = "Select EL_Referencia from EstoqueLoja " _
                & "where EL_Referencia = '" & wReferenciaProduto & "'"
                Set RsComparaEstoque = rdoCNLoja.OpenResultset(SQL)
            If RsComparaEstoque.EOF Then
                SQL = ""
                SQL = "insert into EstoqueLoja (EL_Loja, EL_Referencia, EL_CodigoFornecedor, EL_Estoque, " _
                    & "EL_VendaMes, EL_EstoqueAnterior, EL_UltimaVenda, EL_EntradaMes, EL_SaidaMes) " _
                    & "Values('" & LojaControle & "', '" & wReferenciaProduto & "', " & wFornecedorProduto & ", " _
                    & "0,0,0,0,0,0)"
                    rdoCNLoja.Execute (SQL)
            End If
             RsPegaDadosProduto.MoveNext
        Loop
    End If

End Function

'Function ComparaNFCentralLoja(ByVal Data As String, ByVal Loja As String) As Boolean
'    Dim rdoCompNFCentral As rdoResultset
'    Dim rsCompNfLoja As rdoResultset
'
'    On Error Resume Next
'Voltar:
'    Conexao.Close
'    If ConectaODBC(Conexao, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = True Then
'        SQL = ""
'        SQL = "Select sum(convert(decimal(6,2),TotalNota)) as TotalCentral,Serie,TipoNota from NfCapa " _
'            & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' and LojaOrigem='" & Loja & "' and NF>0" _
'            & "group by Serie,TipoNota order by Serie,TipoNota"
'            Set rdoCompNFCentral = Conexao.OpenResultset(SQL)
'        Do While Not rdoCompNFCentral.EOF
'            SQL = ""
'            SQL = "Select sum(TotalNota) as TotalLoja from NfCapa " _
'                & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' and LojaOrigem='" & Loja & "' " _
'                & "and Serie = '" & rdoCompNFCentral("Serie") & "' " _
'                & "and TipoNota = '" & rdoCompNFCentral("TipoNota") & "' " _
'                & "Having sum(TotalNota) <> " & ConverteVirgula(rdoCompNFCentral("TotalCentral")) & ""
'                Set rsCompNfLoja = rdocnloja.OpenResultset (SQL)
'            If Not rsCompNfLoja.EOF Then
'                ReenviaNotaFiscal Data, rdoCompNFCentral("Serie"), rdoCompNFCentral("TipoNota")
'                'rdoCompNFCentral.Close
'                'rsCompNfLoja.Close
'                'GoTo Voltar
'            End If
'            rdoCompNFCentral.MoveNext
'        Loop
'    Else
'        If MsgBox("Não foi possivel conectar-se ao servidor, Deseja tentar novamente", vbCritical + vbYesNo, "ERRO") = vbYes Then
'            GoTo Voltar
'        End If
'    End If
'    ComparaNFCentralLoja = True
'
'End Function
'
'Function ComparaMovCaixaCentralLoja(ByVal Data As String, ByVal Loja As String) As Boolean
'    Dim rdoCompMovCaixaCentral As rdoResultset
'    Dim rsCompMovCaixaLoja As rdoResultset
'
'Voltar:
'    SQL = ""
'    SQL = "Select sum(convert(decimal(6,2),MC_Valor)) as TotalCentral,MC_Grupo from MovimentoCaixa " _
'        & "where MC_Loja='" & Loja & "' and MC_Data='" & Format(Data, "mm/dd/yyyy") & "' " _
'        & "and MC_Grupo between 20101 and 20108 " _
'        & "group by MC_Grupo order by MC_Grupo"
'    Set rdoCompMovCaixaCentral = Conexao.OpenResultset(SQL)
'    Do While Not rdoCompMovCaixaCentral.EOF
'        SQL = ""
'        SQL = "Select sum(MC_Valor) as TotalLoja from MovimentoCaixa " _
'            & "where MC_Loja='" & Loja & "' " _
'            & "and MC_Data='" & Format(Data, "mm/dd/yyyy") & "' " _
'            & "and MC_Grupo=" & rdoCompMovCaixaCentral("MC_Grupo") & " " _
'            & "Having sum(MC_Valor) <> " & ConverteVirgula(rdoCompMovCaixaCentral("TotalCentral")) & ""
'        Set rsCompMovCaixaLoja = rdocnloja.OpenResultset (SQL)
'        If Not rsCompMovCaixaLoja.EOF Then
'            ReenviaMovimentoCaixa Data, rsCompMovCaixaLoja("MC_Data")
'            'rdoCompMovCaixaCentral.Close
'            'rsCompMovCaixaLoja.Close
'            'GoTo Voltar
'        End If
'        rdoCompMovCaixaCentral.MoveNext
'    Loop
'    ComparaMovCaixaCentralLoja = True
'
'End Function
'
'Function ReenviaMovimentoCaixa(ByVal Data As String, ByVal Grupo As String)
'
''
''------------------------------------Atualiza Movimento Caixa------------------------------
''
'    SQL = "Select * from MovimentoCaixa " _
'        & "Where MC_Data='" & Format(Data, "mm/dd/yyyy") & "' " _
'        & "and MC_Grupo = " & Grupo & " "
'        Set RsPegaMovCaixaEnvio = rdocnloja.OpenResultset (SQL)
'        If Not RsPegaMovCaixaEnvio.EOF Then
'            Do While Not RsPegaMovCaixaEnvio.EOF
'                On Error Resume Next
'                    Wserie = IIf(IsNull(RsPegaMovCaixaEnvio("MC_Serie")), "0", RsPegaMovCaixaEnvio("MC_Serie"))
'                    WBomPara = IIf(IsNull(RsPegaMovCaixaEnvio("MC_BOMPARA")), Format(Date, "mm/dd/yyyy"), RsPegaMovCaixaEnvio("MC_BOMPARA"))
'                    wNumeroCheque = IIf(IsNull(RsPegaMovCaixaEnvio("MC_NumeroCheque")), 0, RsPegaMovCaixaEnvio("MC_NumeroCheque"))
'                    SQL = "insert into MovimentoCaixa " _
'                        & "(MC_NumeroECF,MC_CodigoOperador,MC_Loja,MC_Data, " _
'                        & "MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_Banco, " _
'                        & "MC_Agencia,MC_ContaCorrente,MC_BomPara,MC_Parcelas, " _
'                        & "MC_Remessa,MC_Sequencia,MC_SituacaoEnvio,MC_NumeroCheque) " _
'                        & "Values(" & RsPegaMovCaixaEnvio("MC_NumeroECF") & "," & RsPegaMovCaixaEnvio("MC_CodigoOperador") & ",'" & RsPegaMovCaixaEnvio("MC_Loja") & "','" & Format(RsPegaMovCaixaEnvio("MC_Data"), "mm/dd/yyyy") & "', " _
'                        & "" & RsPegaMovCaixaEnvio("MC_Grupo") & "," & RsPegaMovCaixaEnvio("MC_Documento") & ",'" & Wserie & "'," & ConverteVirgula(RsPegaMovCaixaEnvio("MC_Valor")) & "," & RsPegaMovCaixaEnvio("MC_Banco") & ", " _
'                        & "'" & RsPegaMovCaixaEnvio("MC_Agencia") & "'," & RsPegaMovCaixaEnvio("MC_ContaCorrente") & ",'" & Format(WBomPara, "mm/dd/yyyy") & "'," & RsPegaMovCaixaEnvio("MC_Parcelas") & ", " _
'                        & "" & RsPegaMovCaixaEnvio("MC_Remessa") & "," & RsPegaMovCaixaEnvio("MC_Sequencia") & ",'A'," & wNumeroCheque & " )"
'                        Conexao.Execute (SQL)
'                If Err.Number = 0 Then
'                    CommitTrans
'                Else
'                    Rollback
'                End If
'                RsPegaMovCaixaEnvio.MoveNext
'            Loop
'            RsPegaMovCaixaEnvio.Close
'        End If
'
'End Function
'
'Function ReenviaNotaFiscal(ByVal Data As String, ByVal Serie As String, ByVal TipoNota As String)
'    Dim RsPegaNotaEnvio As rdoResultset
'    Dim WdataPag As String
'    Dim wDataPed As String
'    Dim wInscriCli As String
'    Dim wBairroCli As String
'    Dim wCepCli As String
'    Dim WCFOAux As String
'    Dim WAnexoAux As String
'    Dim wCarimbo1 As String
'    Dim wCarimbo2 As String
'    Dim wCarimbo3 As String
'    Dim wCarimbo4 As String
'    Dim wLojaVenda As String
'
''
''------------------------------------------ENVIA NFCAPA----------------------------------
''
'
'    SQL = ""
'    SQL = "Select * from NfCapa " _
'        & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
'        & "and TipoNota = '" & TipoNota & "'  " _
'        & "and NF > 0 and Serie ='" & Serie & "' Order by NF"
'    Set RsPegaNotaEnvio = rdocnloja.OpenResultset (SQL)
'
'    If Not RsPegaNotaEnvio.EOF Then
'        Do While Not RsPegaNotaEnvio.EOF
'            On Error Resume Next
'            BeginTrans
'            WdataPag = IIf(IsNull(RsPegaNotaEnvio("DataPag")), Format(Date, "MM/DD/YYYY"), RsPegaNotaEnvio("DataPag"))
'            wDataPed = IIf(IsNull(RsPegaNotaEnvio("dataPed")), Format(Date, "MM/DD/YYYY"), RsPegaNotaEnvio("DataPed"))
'            wInscriCli = IIf(IsNull(RsPegaNotaEnvio("InscriCli")), 0, RsPegaNotaEnvio("InscriCli"))
'            wBairroCli = IIf(IsNull(RsPegaNotaEnvio("BairroCli")), "0", RsPegaNotaEnvio("BairroCli"))
'            wCepCli = IIf(IsNull(RsPegaNotaEnvio("CepCli")), 0, RsPegaNotaEnvio("CepCli"))
'            WCFOAux = IIf(IsNull(RsPegaNotaEnvio("CfoAux")), 0, RsPegaNotaEnvio("CfoAux"))
'            WAnexoAux = IIf(IsNull(RsPegaNotaEnvio("AnexoAux")), 0, RsPegaNotaEnvio("AnexoAux"))
'            wCarimbo1 = IIf(IsNull(RsPegaNotaEnvio("Carimbo1")), "0", RsPegaNotaEnvio("Carimbo1"))
'            wCarimbo2 = IIf(IsNull(RsPegaNotaEnvio("Carimbo2")), "0", RsPegaNotaEnvio("Carimbo2"))
'            wCarimbo3 = IIf(IsNull(RsPegaNotaEnvio("Carimbo3")), "0", RsPegaNotaEnvio("Carimbo3"))
'            wCarimbo4 = IIf(IsNull(RsPegaNotaEnvio("Carimbo4")), "0", RsPegaNotaEnvio("Carimbo4"))
'            wLojaVenda = IIf(IsNull(RsPegaNotaEnvio("LojaVenda")), "0", RsPegaNotaEnvio("LojaVenda"))
'
'
'            SQL = ""
'            SQL = "Insert  into NfCapa " _
'            & "(NUMEROPED,DATAEMI,VENDEDOR,VLRMERCADORIA, DESCONTO, " _
'            & "SUBTOTAL,LOJAORIGEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG, " _
'            & "PGENTRA,LOJAT,QTDITEM,PEDCLI,TM,PESOBR,PESOLQ,VALFRETE,FRETECOBR, " _
'            & "OUTRALOJA,OUTROVEND,NF,TOTALNOTA,NATOPERACAO,DATAPED,BASEICMS, " _
'            & "ALIQICMS,VLRICMS,SERIE,HORA,TOTALIPI,ECF,NUMEROSF,NOMCLI,FONECLI, " _
'            & "CGCCLI,INSCRICLI,ENDCLI,UFCLIENTE,MUNICIPIOCLI,BAIRROCLI,CEPCLI, " _
'            & "PESSOACLI,REGIAOCLI,CFOAUX,AnexoAUx,PAGINANF,ECFNF,Carimbo1,Carimbo2, " _
'            & "Carimbo3,Carimbo4,CustoMedioLiquido,VendaLiquida,MargemContribuicao, " _
'            & "ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa, " _
'            & "SituacaoEnvio,VendedorLojaVenda,LojaVenda) " _
'            & "Values(" & RsPegaNotaEnvio("NUMEROPED") & ",'" & Format(RsPegaNotaEnvio("DATAEMI"), "mm/dd/yyyy") & "'," & RsPegaNotaEnvio("VENDEDOR") & "," & ConverteVirgula(RsPegaNotaEnvio("VLRMERCADORIA")) & ", " & ConverteVirgula(RsPegaNotaEnvio("DESCONTO")) & ", " _
'            & "" & ConverteVirgula(RsPegaNotaEnvio("SUBTOTAL")) & ",'" & RsPegaNotaEnvio("LOJAORIGEM") & "','" & RsPegaNotaEnvio("TIPONOTA") & "','" & RsPegaNotaEnvio("CONDPAG") & "'," & ConverteVirgula(RsPegaNotaEnvio("AV")) & "," & RsPegaNotaEnvio("CLIENTE") & "," & RsPegaNotaEnvio("CODOPER") & ",'" & WdataPag & "', " _
'            & "" & ConverteVirgula(RsPegaNotaEnvio("PGENTRA")) & ",'" & RsPegaNotaEnvio("LOJAT") & "'," & RsPegaNotaEnvio("QTDITEM") & "," & RsPegaNotaEnvio("PEDCLI") & "," & RsPegaNotaEnvio("TM") & "," & ConverteVirgula(RsPegaNotaEnvio("PESOBR")) & "," & ConverteVirgula(RsPegaNotaEnvio("PESOLQ")) & "," & ConverteVirgula(RsPegaNotaEnvio("VALFRETE")) & "," & ConverteVirgula(RsPegaNotaEnvio("FRETECOBR")) & ", " _
'            & "'" & RsPegaNotaEnvio("OUTRALOJA") & "'," & RsPegaNotaEnvio("OUTROVEND") & "," & RsPegaNotaEnvio("NF") & "," & ConverteVirgula(RsPegaNotaEnvio("TOTALNOTA")) & "," & RsPegaNotaEnvio("NATOPERACAO") & ",'" & Format(wDataPed, "mm/dd/yyyy") & "'," & ConverteVirgula(RsPegaNotaEnvio("BASEICMS")) & ", " _
'            & "" & ConverteVirgula(RsPegaNotaEnvio("ALIQICMS")) & "," & ConverteVirgula(RsPegaNotaEnvio("VLRICMS")) & ",'" & RsPegaNotaEnvio("SERIE") & "','" & Format(RsPegaNotaEnvio("HORA"), "hh:mm") & "'," & ConverteVirgula(RsPegaNotaEnvio("TOTALIPI")) & "," & RsPegaNotaEnvio("ECF") & "," & RsPegaNotaEnvio("NUMEROSF") & ",'" & RsPegaNotaEnvio("NOMCLI") & "' ,'" & RsPegaNotaEnvio("FONECLI") & "', " _
'            & "'" & RsPegaNotaEnvio("CGCCLI") & "','" & wInscriCli & "','" & RsPegaNotaEnvio("ENDCLI") & "','" & RsPegaNotaEnvio("UFCLIENTE") & "','" & RsPegaNotaEnvio("MUNICIPIOCLI") & "','" & wBairroCli & "','" & wCepCli & "', " _
'            & "" & RsPegaNotaEnvio("PESSOACLI") & "," & RsPegaNotaEnvio("REGIAOCLI") & ",'" & WCFOAux & "','" & WAnexoAux & "'," & RsPegaNotaEnvio("PAGINANF") & "," & RsPegaNotaEnvio("ECFNF") & ",'" & wCarimbo1 & "','" & wCarimbo2 & "', " _
'            & "'" & wCarimbo3 & "','" & wCarimbo4 & "'," & ConverteVirgula(RsPegaNotaEnvio("CustoMedioLiquido")) & "," & ConverteVirgula(RsPegaNotaEnvio("VendaLiquida")) & "," & ConverteVirgula(RsPegaNotaEnvio("MargemContribuicao")) & ", " _
'            & "" & ConverteVirgula(RsPegaNotaEnvio("ValorTotalCodigoZero")) & "," & ConverteVirgula(RsPegaNotaEnvio("TotalNotaAlternativa")) & "," & ConverteVirgula(RsPegaNotaEnvio("ValorMercadoriaAlternativa")) & ", " _
'            & "'A'," & RsPegaNotaEnvio("VendedorLojaVenda") & ",'" & wLojaVenda & "') "
'            Conexao.Execute (SQL)
'
'            If Err.Number = 0 Then
'                CommitTrans
'            Else
'                Rollback
'            End If
'            RsPegaNotaEnvio.MoveNext
'        Loop
'    End If
'
'End Function
'
'Function ReenviaItens(ByVal Data As String, ByVal Serie As String, ByVal TipoNota As String)
'    Dim RsPegaItensEnvio As String
'    Dim wVLICMS As Double
'    Dim WVBUNIT As Double
'    Dim wIcmPdv As Double
'    Dim wCodBarra As Double
'    Dim wDetalheImpressao As String
'    Dim WDESCRAT As String
'    Dim WVLIPI As Double
'    Dim wDesconto As Double
'    Dim wREDUCAOICMS As Double
'
''
''------------------------------------------ENVIA NFITENS----------------------------------
''
'
'    SQL = ""
'    SQL = "Select * from NfItens " _
'        & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
'        & "and TipoNota = '" & TipoNota & "'  " _
'        & "and NF > 0 AND Serie ='" & Serie & "' order by NF"
'    Set RsPegaItensEnvio = rdocnloja.OpenResultset (SQL)
'
'    If Not RsPegaItensEnvio.EOF Then
'        Do While Not RsPegaItensEnvio.EOF
'            wVLICMS = IIf(IsNull(RsPegaItensEnvio("VALORICMS")), 0, RsPegaItensEnvio("VALORICMS"))
'            WVBUNIT = IIf(IsNull(RsPegaItensEnvio("VBUNIT")), 0, RsPegaItensEnvio("VBUNIT"))
'            wIcmPdv = IIf(IsNull(RsPegaItensEnvio("ICMPDV")), 0, RsPegaItensEnvio("ICMPDV"))
'            wCodBarra = IIf(IsNull(RsPegaItensEnvio("CODBARRA")), 0, RsPegaItensEnvio("CODBARRA"))
'            wDetalheImpressao = IIf(IsNull(RsPegaItensEnvio("DETALHEIMPRESSAO")), "0", RsPegaItensEnvio("DETALHEIMPRESSAO"))
'            WDESCRAT = IIf(IsNull(RsPegaItensEnvio("DESCRAT")), 0, RsPegaItensEnvio("DESCRAT"))
'            WVLIPI = IIf(IsNull(RsPegaItensEnvio("VLIPI")), 0, RsPegaItensEnvio("VLIPI"))
'            wDesconto = IIf(IsNull(RsPegaItensEnvio("DESCONTO")), 0, RsPegaItensEnvio("DESCONTO"))
'            wREDUCAOICMS = IIf(IsNull(RsPegaItensEnvio("REDUCAOICMS")), 0, RsPegaItensEnvio("REDUCAOICMS"))
'
'            On Error Resume Next
'            SQL = "Insert into NfItens " _
'                & "(NUMEROPED,DATAEMI,REFERENCIA,QTDE,VLUNIT,VLUNIT2, " _
'                & "VLTOTITEM,DESCRAT,ICMS,ITEM,VLIPI,DESCONTO,PLISTA,COMISSAO, " _
'                & "VALORICMS,BCOMIS,CSPROD,LINHA,SECAO,VBUNIT,ICMPDV,CODBARRA, " _
'                & "NF,SERIE,LOJAORIGEM,CLIENTE,VENDEDOR,ALIQIPI,TIPONOTA,REDUCAOICMS, " _
'                & "BASEICMS,TIPOMOVIMENTACAO,DETALHEIMPRESSAO,CustoMedioLiquido,VendaLiquida, " _
'                & "MargemContribuicao,EncargosVendaLiquida,EncargosCustoMedioLiquido, " _
'                & "PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,SituacaoEnvio) " _
'                & "Values (" & ConverteVirgula(RsPegaItensEnvio("NUMEROPED")) & ",'" & Format(RsPegaItensEnvio("DATAEMI"), "mm/dd/yyyy") & "','" & RsPegaItensEnvio("REFERENCIA") & "'," & RsPegaItensEnvio("QTDE") & "," & ConverteVirgula(RsPegaItensEnvio("VLUNIT")) & "," & ConverteVirgula(RsPegaItensEnvio("VLUNIT2")) & ", " _
'                & "" & ConverteVirgula(RsPegaItensEnvio("VLTOTITEM")) & "," & ConverteVirgula(WDESCRAT) & "," & ConverteVirgula(RsPegaItensEnvio("ICMS")) & "," & RsPegaItensEnvio("ITEM") & "," & ConverteVirgula(WVLIPI) & "," & ConverteVirgula(wDesconto) & "," & ConverteVirgula(RsPegaItensEnvio("PLISTA")) & "," & ConverteVirgula(RsPegaItensEnvio("COMISSAO")) & ", " _
'                & "" & ConverteVirgula(wVLICMS) & "," & ConverteVirgula(RsPegaItensEnvio("BCOMIS")) & "," & RsPegaItensEnvio("CSPROD") & "," & RsPegaItensEnvio("LINHA") & "," & RsPegaItensEnvio("SECAO") & "," & ConverteVirgula(WVBUNIT) & "," & ConverteVirgula(wIcmPdv) & "," & wCodBarra & ", " _
'                & "" & RsPegaItensEnvio("NF") & ",'" & RsPegaItensEnvio("SERIE") & "','" & RsPegaItensEnvio("LOJAORIGEM") & "'," & RsPegaItensEnvio("CLIENTE") & "," & RsPegaItensEnvio("VENDEDOR") & "," & ConverteVirgula(RsPegaItensEnvio("ALIQIPI")) & ",'" & RsPegaItensEnvio("TIPONOTA") & "'," & ConverteVirgula(RsPegaItensEnvio("REDUCAOICMS")) & ", " _
'                & "" & ConverteVirgula(RsPegaItensEnvio("BASEICMS")) & "," & RsPegaItensEnvio("TIPOMOVIMENTACAO") & ",'" & wDetalheImpressao & "'," & ConverteVirgula(RsPegaItensEnvio("CustoMedioLiquido")) & "," & ConverteVirgula(RsPegaItensEnvio("VendaLiquida")) & ", " _
'                & "" & ConverteVirgula(RsPegaItensEnvio("MargemContribuicao")) & "," & ConverteVirgula(RsPegaItensEnvio("EncargosVendaLiquida")) & "," & ConverteVirgula(RsPegaItensEnvio("EncargosCustoMedioLiquido")) & ", " _
'                & "" & ConverteVirgula(RsPegaItensEnvio("PrecoUnitAlternativa")) & "," & ConverteVirgula(RsPegaItensEnvio("ValorMercadoriaAlternativa")) & ",'" & RsPegaItensEnvio("ReferenciaAlternativa") & "','A')"
'                Conexao.Execute (SQL)
'            If Err.Number = 0 Then
'                CommitTrans
'            Else
'                Rollback
'            End If
'            RsPegaItensEnvio.MoveNext
'        Loop
'    End If
'
'End Function

'Function FechamentoCaixaCentral(ByVal Data As String, ByVal Loja As String) As Boolean
'
'    If ComparaNFCentralLoja(Data, Loja) = True Then
'        If ComparaMovCaixaCentralLoja(Data, Loja) = True Then
'            SQL = ""
'            SQL = "Update Loja set LO_SituacaoCaixa='F', LO_ProcessaNota='T' " _
'                & "where LO_Loja='" & Loja & "'"
'                Conexao.Execute (SQL)
'            FechamentoCaixaCentral = True
'        Else
'            FechamentoCaixaCentral = False
'        End If
'    Else
'        FechamentoCaixaCentral = False
'    End If
'
'End Function

'Function VerificaComunicacao()
'    Dim rsVerConexao As rdoResultset
'
' On Error Resume Next
'    SQL = ""
'    SQL = "Select top 1 EC_Referencia from EstoqueCentralSQL"
'        Set rsVerConexao = rdoCnLoja.OpenResultset(SQL)
'    If Err.Number = 0 Then
'        mdiBalcao.AniComunicacao.Open "C:\sistemas\Balcao2000\ComunicacaoOK.avi"
'        mdiBalcao.AniComunicacao.ToolTipText = "Comunicação com " & LCase(GLB_Servidor) & " esta funcionando"
'        rsVerConexao.Close
'    Else
'        mdiBalcao.AniComunicacao.Open "C:\sistemas\Balcao2000\ComunicacaoFalhou.avi"
'        mdiBalcao.AniComunicacao.ToolTipText = "Comunicação com " & LCase(GLB_Servidor) & " não esta funcionando"
'    End If
'    mdiBalcao.AniComunicacao.Play
'    Err.Clear
'
'End Function




Function VerificaExeRodando(ByVal CaminhoExe As String) As Boolean

    Dim hSnapshot As Long, lRet As Long, P As PROCESSENTRY32
    Dim NomeExecucao As String
    Dim NomeRodando As String
    
    VerificaExeRodando = False
    P.dwSize = Len(P)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
    If hSnapshot Then
        lRet = Process32First(hSnapshot, P)
        Do While lRet
            NomeRodando = Left$(P.szExeFile, InStr(P.szExeFile, Chr$(0)) - 1)
            NomeExecucao = ""
            If Len(NomeRodando) > Len(CaminhoExe) Then
                NomeExecucao = Mid(NomeRodando, Len(NomeRodando) - (Len(CaminhoExe) - 1), Len(CaminhoExe))
            End If
            If UCase(CaminhoExe) = UCase(NomeExecucao) Then
                VerificaExeRodando = True
                Exit Function
            End If
            lRet = Process32Next(hSnapshot, P)
        Loop
        lRet = CloseHandle(hSnapshot)
    End If
    
End Function


Function LiberaBc2000(ByVal TipoLiberacao As Integer)
    
    If TipoLiberacao = 1 Then 'enibe os menus
        mdiBalcao.mnuProcedimentos.Visible = False
        mdiBalcao.MnuSegurança.Visible = False
        mdiBalcao.MnuOutros.Visible = False
        mdiBalcao.mnuLiberaMenu.Visible = True
        mdiBalcao.mnuBloqueiaMenu.Visible = False
    ElseIf TipoLiberacao = 2 Then 'exibe todos os menus
        mdiBalcao.mnuProcedimentos.Visible = True
        mdiBalcao.MnuSegurança.Visible = True
        mdiBalcao.MnuOutros.Visible = True
        mdiBalcao.mnuBloqueiaMenu.Visible = True
        mdiBalcao.mnuLiberaMenu.Visible = False
    End If
        
End Function


Function VerProcessoFechamento() As String
    Dim rsSeqFechamento As rdoResultset
    
    SQL = "Select CT_SeqFechamento from Controle where CT_SeqFechamento not in ('EF','P') "
        Set rsSeqFechamento = rdoCNLoja.OpenResultset(SQL)
    If Not rsSeqFechamento.EOF Then
        VerProcessoFechamento = rsSeqFechamento("CT_SeqFechamento")
    Else
        VerProcessoFechamento = ""
    End If
        
End Function

Function AtualizaProcessoTela(ByVal NomeList As String, ByVal NumeroProcesso As Integer, ByVal Tipo As Integer)
    
    If Tipo = 1 Then
        frmAguarde.lstProcessos.AddItem NomeList
        frmAguarde.lstProcessos.Selected(NumeroProcesso) = False
    ElseIf Tipo = 2 Then
        frmAguarde.lstProcessos.Selected(NumeroProcesso) = True
    End If
    frmAguarde.Refresh
    frmAguarde.lstProcessos.Refresh
    
End Function

Function FechaLojaCentral(ByVal Loja As String) As Boolean
        
    Err.Clear
    On Error Resume Next
    Conexao.Close
    If ConectaODBC(Conexao, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = True Then
        Err.Clear
        On Error Resume Next
        BeginTrans
        SQL = ""
        SQL = "Update Loja set LO_SituacaoCaixa='F',LO_ProcessaNota='T',LO_ProcessaEstoque='S',LO_ProcessaAuditorEstoque='S' " _
            & "where LO_Loja='" & Loja & "'"
            Conexao.Execute (SQL)
        If Err.Number = 0 Then
            CommitTrans
            FechaLojaCentral = True
        Else
            Rollback
            FechaLojaCentral = False
        End If
        Conexao.Close
    Else
        FechaLojaCentral = False
    End If
            
        
End Function

Function AchaDataCaixa() As String
    Dim rsPegaData As rdoResultset
    
    SQL = "Select Max(CT_Data) as Data from CtCaixa"
        Set rsPegaData = rdoCNLoja.OpenResultset(SQL)
    If Not rsPegaData.EOF Then
        AchaDataCaixa = rsPegaData("Data")
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

Function RateiaDesconto(ByVal ValorTotal As Double, ByVal Desconto As Double) As Double
    
    RateiaDesconto = Format((Desconto / ValorTotal) * 100, "###,###,###,##0.000000")

End Function



Function DescricaoOperacao(ByVal Descricao As String)

'    mdiBalcao.stbBarra.Panels.Item(1).Text = Descricao

End Function



Function ImprimirCotacao(ByVal Pedido As Double, ByVal Vendedor As String, ByVal NomeVend As String, ByVal CondPag As Integer)
    Dim rdoPedido As rdoResultset
    Dim VarImp As String
    Dim Pagina As Integer
    Dim SubTotal As Double
    Dim Total As Double
    Dim Desconto As Double
    Dim Linhas As Integer
    Dim Descricao As String
    Dim Referencia As String
    Dim VlUnit As Double
    Dim VlUnit2 As Double
    Dim VlTotItem As Double
    Dim DescontoItem As Double
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    Screen.MousePointer = 11
    Pagina = 1
    SubTotal = 0
    Total = 0
    Desconto = 0
    Linhas = 7
    CabecalhoCotacao Pedido, Pagina
    SQL = ""
    SQL = "Select Referencia,Qtde,VlUnit,VlUnit2,DescricaoAlternativa,ReferenciaAlternativa,PrecoUnitAlternativa,ValorMercadoriaAlternativa,PR_descricao from NfItens,Produto " _
        & "where PR_Referencia=Referencia and NumeroPed=" & Pedido & ""
    Set rdoPedido = rdoCNLoja.OpenResultset(SQL)
    If Not rdoPedido.EOF Then
        Do While Not rdoPedido.EOF
            If rdoPedido("DescricaoAlternativa") <> "0" Then
                Descricao = rdoPedido("DescricaoAlternativa")
            Else
                Descricao = rdoPedido("PR_Descricao")
            End If
            If rdoPedido("ReferenciaAlternativa") <> "0" Then
                Referencia = rdoPedido("ReferenciaAlternativa")
            Else
                Referencia = rdoPedido("Referencia")
            End If
            Linhas = Linhas + 1
            If rdoPedido("ValorMercadoriaAlternativa") > 0 Then
                VlUnit = rdoPedido("PrecoUnitAlternativa")
                VlUnit2 = rdoPedido("ValorMercadoriaAlternativa")
                SubTotal = SubTotal + VlUnit2
                Total = Total + VlUnit2
                DescontoItem = 0
            Else
                VlUnit = rdoPedido("VlUnit")
                VlUnit2 = rdoPedido("VlUnit2")
                SubTotal = SubTotal + (rdoPedido("VLUnit") * rdoPedido("Qtde"))
                Total = Total + rdoPedido("VlUnit2")
                Desconto = Desconto + ((rdoPedido("VLUnit") * rdoPedido("Qtde")) - rdoPedido("VlUnit2"))
                DescontoItem = (rdoPedido("VlUnit") * rdoPedido("Qtde")) - rdoPedido("VLUnit2")
            End If
            VarImp = Left(Referencia & Space(11), 11) _
                & Left(Descricao & Space(42), 42) _
                & Right(Space(6) & rdoPedido("Qtde"), 6) _
                & Right(Space(12) & Format(VlUnit, "###,###,###,##0.00"), 12) _
                & Right(Space(10) & Format(DescontoItem, "###,###,###,##0.00"), 10) _
                & Right(Space(12) & Format(VlUnit2, "###,###,###,##0.00"), 12)
            Printer.Print VarImp
            If Linhas = 62 Then
                Printer.Print "_________________________________________________________________________________________________________________________"
                Printer.NewPage
                CabecalhoCotacao Pedido, Pagina + 1
                Linhas = 8
            End If
            rdoPedido.MoveNext
        Loop
        If Desconto = 0 Then
            SubTotal = Total
        End If
        FinalizaCotacao Linhas, SubTotal, Desconto, Total, Vendedor, NomeVend, CondPag
    Else
        Screen.MousePointer = 0
        MsgBox "Impossivel imprimir cotação", vbExclamation, "Aviso"
        Exit Function
    End If
    Screen.MousePointer = 0
    
End Function


Function CabecalhoCotacao(ByVal Pedido As Double, ByVal Pagina As Integer)
    Dim rdoPedido As rdoResultset
    Dim VarImp As String

    SQL = ""
    SQL = "Select NfCapa.*,Lojas.* from NfCapa,Lojas " _
        & "where NumeroPed=" & Pedido & " and Lo_Loja=LojaOrigem "
    Set rdoPedido = rdoCNLoja.OpenResultset(SQL)
    
    If Not rdoPedido.EOF Then
        Printer.ScaleMode = vbMillimeters
        Printer.ForeColor = "0"
        Printer.FontSize = 8
        Printer.FontName = "draft 20cpi"
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.DrawWidth = 3
        Printer.FontName = "COURIER NEW"
        Printer.FontSize = 10#
        Printer.Print "COTACAO DE VENDA" & Space(35) & "NUMERO: " & Right(String(6, "0") & Pedido, 6) & Space(3) & "Data: " & Format(Date, "dd/mm/yyyy") & Space(5) & "PAG: " & Pagina
        Printer.Print "_________________________________________________________________________________________________________________________"
        VarImp = Left(rdoPedido("LO_Razao") & Space(45), 45) & Left(rdoPedido("NomCli") & Space(60), 60)
        Printer.Print VarImp
        VarImp = Left(rdoPedido("LO_Endereco") & Space(45), 45) & Left(rdoPedido("EndCli") & Space(60), 60)
        Printer.Print VarImp
        VarImp = Left(Right(String(7, "0") & rdoPedido("LO_CEP"), 7) & " - " & rdoPedido("LO_Bairro") & "  -  " & rdoPedido("LO_Municipio") & " - " & rdoPedido("LO_UF") & Space(45), 45)
        VarImp = VarImp & Left(rdoPedido("CepCli") & " - " & rdoPedido("BairroCli") & "  -  " & rdoPedido("MunicipioCli") & " - " & rdoPedido("UfCliente") & Space(60), 60)
        Printer.Print VarImp
        VarImp = Left("Telefone " & rdoPedido("LO_Telefone") & Space(45), 45) & Left("Telefone " & rdoPedido("FoneCli") & Space(60), 60)
        Printer.Print VarImp
        Printer.Print "_________________________________________________________________________________________________________________________"
        Printer.Print ""
        Printer.Print "REFERENCIA DESCRICAO                                   QTDE  PRECO UNIT  DESCONTO  PRECO TOTAL"
        Printer.Print ""
        rdoPedido.Close
    End If

    
    

End Function


Function FinalizaCotacao(ByVal Linhas As Integer, ByVal SubTotal As Double, ByVal Desconto As Double, ByVal Total As Double, ByVal Vendedor As String, ByVal NomeVend As String, ByVal CondPag As Integer)
    Dim rdoDescPag As rdoResultset
    Dim Desc As String
    
    Desc = ""
    SQL = ""
    SQL = "Select CP_Condicao from CondicaoPagamento " _
        & "where CP_Codigo=" & CondPag & ""
    Set rdoDescPag = rdoCNLoja.OpenResultset(SQL)
    If Not rdoDescPag.EOF Then
        Desc = rdoDescPag("CP_Condicao")
        rdoDescPag.Close
    End If
    For Linhas = Linhas To 56
        Printer.Print ""
    Next
    
    Printer.Print "_________________________________________________________________________________________________________________________"
    Printer.Print "COND PAGTO : " & Left(Desc & Space(42), 42) & "SUB-TOTAL" & Space(9) & "DESCONTO" & Space(9) & "TOTAL"
    Printer.Print "VALIDADE   : " & Left(Format(Date, "dd/mm/yyyy") & Space(12), 12) & Space(24) & Right(Space(15) & Format(SubTotal, "###,###,###,##0.00"), 15) _
        & Space(2) & Right(Space(14) & Format(Desconto, "###,###,###,##0.00"), 14) & Right(Space(15) & Format(Total, "###,###,###,##0.00"), 15)
    Printer.Print "VENDEDOR   : " & Vendedor & " - " & NomeVend
    Printer.Print "_________________________________________________________________________________________________________________________"
    Printer.EndDoc
    
End Function


Function PegaSerieNota() As String
    Dim rdoSerie As rdoResultset
    
    SQL = ""
    SQL = "Select CT_SerieNota from Controle"
    Set rdoSerie = rdoCNLoja.OpenResultset(SQL)
    If Not rdoSerie.EOF Then
        PegaSerieNota = rdoSerie("CT_SerieNota")
    End If
    rdoSerie.Close

End Function


Function ComparaHoraConexao(ByRef rdoHoraConexao As rdoResultset, ByVal Data As String) As Boolean
    
    Data = Date & " " & Time
    
    SQL = ""
    SQL = "Select CT_HoraConexao From Controle"
    Set rdoHoraConexao = rdoCnLojaBach.OpenResultset(SQL)
    If Not rdoHoraConexao.EOF Then
        If DateDiff("n", rdoHoraConexao("CT_HoraConexao"), Data) > 5 Then
            ComparaHoraConexao = True
        Else
            ComparaHoraConexao = False
        End If
    End If

End Function

Function SituacaoFechaRetaguarda() As Boolean

    SQL = ""
    SQL = "Select * from FechamentoRetaguarda Where FR_SituacaoFechamento = 'T' " _
        & "and FR_DataFechamento = '" & Format(DateAdd("d", -1, Date), "mm/dd/yyyy") & "'"
    Set FechaRetaguarda = rdoCnLojaBach.OpenResultset(SQL)
    
    If Not FechaRetaguarda.EOF Then
        SituacaoFechaRetaguarda = True
    Else
        SituacaoFechaRetaguarda = False
    End If
    
    FechaRetaguarda.Close

End Function

Function ConectaOdbcCD(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
           
    On Error GoTo ConexaoErro
    
    With RdoVar
        Servidor = Glb_ServidorLocal
        WBANCO = "LojaCPD"
        'Usuario = "sa"
        'Senha = "jeda36"
    
        .Connect = "Dsn=" & Trim(Servidor) & ";" _
                & "Server=" & Trim(Servidor) & ";" _
                & "DataBase=" & Trim(WBANCO) & ";" _
                & "MaxBufferSize=512;" _
                & "PageTimeout=5;" _
                & "UID=" & Usuario & ";" _
                & "PWD=" & Senha & ";"
    
        .LoginTimeout = 10
        .CursorDriver = rdUseClientBatch
        .EstablishConnection rdDriverNoPrompt
    End With
    
    ConectaOdbcCD = True
    Wconectou = True
    Exit Function
    
ConexaoErro:

    ConectaOdbcCD = False
    Wconectou = False

End Function

Function TraduzMes(ByVal Mes As Integer) As String

    Select Case Mes
        Case 1, -11: TraduzMes = "Janeiro"
        Case 2, -10: TraduzMes = "Fevereiro"
        Case 3, -9: TraduzMes = "Março"
        Case 4, -8: TraduzMes = "Abril"
        Case 5, -7: TraduzMes = "Maio"
        Case 6, -6: TraduzMes = "Junho"
        Case 7, -5: TraduzMes = "Julho"
        Case 8, -4: TraduzMes = "Agosto"
        Case 9, -3: TraduzMes = "Setembro"
        Case 10, -2: TraduzMes = "Outubro"
        Case 11, -1: TraduzMes = "Novembro"
        Case 12, 0: TraduzMes = "Dezembro"
    End Select
    
End Function

Sub VerificaSituacaoOnLine()

    SQL = ""
    SQL = "Select CT_BancosOnLine from Controle"
        Set rsLoja = rdoCNLoja.OpenResultset(SQL)
    GLB_BancosOnline = rsLoja("CT_BancosOnLine")

End Sub

Function ConsisteNota(ByRef Nota As Integer, ByRef Serie As String)

    Dim rdoVlNota As rdoResultset

    SQL = ""
    SQL = "Select Sum(VlTotItem) as TotItem, TotalNota, FreteCobr " _
        & "From NfCapa, NfItens " _
        & "where NfItens.Nf = NfCapa.Nf and NfItens.Serie = NfCapa.Serie and " _
        & "NfItens.DataEmi = NfCapa.DataEmi and NfCapa.Nf = " & Nota & " and NfCapa.Serie = '" & Serie & "' and " _
        & "NfCapa.LojaOrigem = '" & Trim(wLoja) & "' " _
        & "Group by TotalNota"
    Set rdoVlNota = rdoCNLoja.OpenResultset(SQL)
    
    If Not rdoVlNota.EOF Then
        If Format(rdoVlNota("TotItem") + rdoVlNota("FreteCobr"), "#,##0.00") <> Format(rdoVlNota("TotalNota"), "#,##0.00") Then
            SQL = ""
            SQL = "Update NfCapa Set TotalNota = " & ConverteVirgula(Format(rdoVlNota("TotItem") + rdoVlNota("FreteCobr"), "#,##0.00")) & " " _
                & "NfCapa.Nf = " & Nota & " and NfCapa.Serie = '" & Serie & "' and " _
                & "NfCapa.LojaOrigem = '" & Trim(wLoja) & "' "
            
            rdoCNLoja.Execute (SQL)
        End If
    End If
    
    rdoVlNota.Close
    
End Function

Function ConsistePedido(ByRef Pedido As Double)

    Dim rdoVlNota As rdoResultset
    Dim rdoVlItem As rdoResultset
    
    SQL = ""
    SQL = "Select Sum(VlUnit2) as TotItem " _
        & "From NfItens " _
        & "where NumeroPed = " & Pedido & " "
    Set rdoVlItem = rdoCNLoja.OpenResultset(SQL)
    
    SQL = ""
    SQL = "Select TotalNota,FreteCobr " _
        & "From NfCapa " _
        & "where NumeroPed = " & Pedido & " "
    Set rdoVlNota = rdoCNLoja.OpenResultset(SQL)
    
    If (Not rdoVlNota.EOF) And (Not rdoVlItem.EOF) Then
        If Format(rdoVlItem("TotItem") + rdoVlNota("FreteCobr"), "#,##0.00") <> Format(rdoVlNota("TotalNota"), "#,##0.00") Then
            SQL = ""
            SQL = "Update NfCapa Set TotalNota = " & ConverteVirgula(Format(rdoVlItem("TotItem") + rdoVlNota("FreteCobr"), "#,##0.00")) & ", VlrMercadoria = " & ConverteVirgula(Format(rdoVlItem("TotItem"), "#,##0.00")) & " " _
                & "Where NfCapa.NumeroPed = " & Pedido & " "
            
            rdoCNLoja.Execute (SQL)
        End If
    End If
    
    rdoVlNota.Close
    rdoVlItem.Close
    
End Function


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
Sub DadosECF()
'  Screen.MousePointer = 11
'
'  SQL = "Select * from ItensVenda " _
'      & "Where ITV_NumeroPedido = " & txtPedido.Text
'      Set RsDados = rdoCNLoja.OpenResultset(SQL)
'      If Not RsDados.EOF Then
'           If RsDados("ITV_Item") = 1 Then
'
'             ' substituir pela dll que pega o numero do cupom
''
''              NroPedido = rdocontrole("CTS_NumeroPedido") + 1
''              SQL = "Update ControleSistema set CTS_NumeroPedido = " & NroPedido
''              rdoCNLoja.Execute SQL, rdExecDirect
'
'              SQL = "Select max(ITV_NotaFiscal) as ITV_Documento from ItensVenda"
'              Set rdoMax = rdoCNLoja.OpenResultset(SQL)
'              If IsNull(rdoMax("ITV_Documento")) Then
'              NroNotaFiscal = 1
'           Else
'              NroNotaFiscal = (rdoMax("ITV_Documento") + 1)
'           End If
'      End If
'
'        Do While Not RsDados.EOF
'          On Error Resume Next
'        SQL = "Insert Into ItensVenda (ITV_Loja," _
'               & "ITV_NotaFiscal," _
'               & "ITV_Serie," _
'               & "ITV_Item," _
'               & "ITV_CodigoProduto," _
'               & "ITV_Data," _
'               & "ITV_Quantidade," _
'               & "ITV_PrecoUnitario," _
'               & "ITV_TipoNota," _
'               & "ITV_Vendedor," _
'               & "ITV_Desconto," _
'               & "ITV_Protocolo," _
'               & "ITV_Situacao, " _
'               & "ITV_NumeroPedido) " _
'               & "Values ('" & txtLoja.Text & "'," & NroNotaFiscal & ",'00'," _
'               & RsDados("ITV_Item") & ",'" & RsDados("ITV_Referencia") _
'               & "','" & Format(Date, "mm/dd/yyyy") & "'," _
'               & RsDados("ITV_Quantidade") & "," _
'               & ConverteVirgula(Format(RsDados("ITV_PrecoVenda"), "0.00")) _
'               & ",'" & RsDados("ITV_TipoPedido") & "'," _
'               & RsDados("ITV_Vendedor") & "," & RsDados("ITV_Desconto") _
'               & "," & frmControlaCaixa.lblNroCaixa.Caption & ",'A', " & txtPedido.Text & ")"
'                  rdoCNLoja.Execute SQL, rdExecDirect
'                  RsDados.MoveNext
'                Loop
'             End If
'             Screen.MousePointer = 0

          Screen.MousePointer = 11

        '
        'substituir pela dll que pega o numero do cupom
        '
        SQL = "Select * from ControleSistema "
        Set rdocontrole = rdoCNLoja.OpenResultset(SQL)
        'rdocontrole.CursorLocation = adUseClient
        'rdocontrole.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        
        NroNotaFiscal = rdocontrole("CTS_Numero00") + 1
        rdocontrole.Close
         
        'Pedido = NroPedido
        
        rdoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        
        SQL = "Update ControleSistema set CTS_Numero00 =" & NroNotaFiscal
        
        rdoCNLoja.Execute SQL
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
     
            
'================Antiga Adilson ==============================
'         SQL = "Select max(ITV_NotaFiscal) as ITV_Documento from ItensVenda"
'                Set rdoMax = rdoCNLoja.OpenResultset(SQL)
'                If IsNull(rdoMax("ITV_Documento")) Then
'                   NroNotaFiscal = 1
'                Else
'                   NroNotaFiscal = (rdoMax("ITV_Documento") + 1)
'                End If

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




