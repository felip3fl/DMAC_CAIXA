Attribute VB_Name = "ModCaixa"
Option Explicit
Public DBLoja As Database
Public DBFBanco As Database

Public PegaDadosCaixa As Recordset
Public VerificaCaixa As Recordset
Public RdoGravaDados As Recordset
Public NFCapaDBF As Recordset
Public NFItemDBF As Recordset
Public RsPegaNumNote As Recordset
Public rsPegaLoja As Recordset
Public RsVerificaPedido As Recordset
Public ISQL As Recordset
Public RSTipoControle As Recordset
Public RsPegaItensPedi As Recordset
Public RsPegaItensEspeciais As Recordset

Global GLB_TrocaVersao As String
Global GLB_Versao As String
Global GLB_VersaoNova As String
Global TipoPedido As String
Global NomeImpressora As Printer
Global Wimpressora As String
Global ContaImpressora As Integer
Global wSair As Boolean
Global GLB_Banco As String


Global GLB_Usuario As String
Global GLB_Senha As String
Global GLB_Servidor As String
Global GLB_NumeroCaixa As String

Global dbMovDia As Database
Global RsDados As Recordset
Global RsAtivaLojaOnLine As Recordset
Global RsDadosDbf As Recordset
Global RsCapaNF As Recordset
Global RsItensNF As Recordset
Global RsICMSInter As Recordset
Global RsUsuario As Recordset
Global RsSelecionaMovCaixa As Recordset
Global RsSelecionaMovBanco As Recordset
Global RsSelecionaMovEstoque As Recordset
Global RsSelecionaDivEstoque As Recordset
Global RsCarimbo As Recordset
Global RsPegaControleMigracao As Recordset
Global RsDescProduto As Recordset
Global RsPegaDescricaoAlternativa As Recordset
Global RsNumeroECF As Recordset


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
Global WNF As String
Global wLoja As String * 5
Global WRazao As String
Global WCGC As String
Global WIest As String
Global Wendereco As String
Global wbairro As String
Global WMunicipio As String
Global westado As String
Global WCep As String
Global WFone As String
Global WNumeroCupom As String * 6
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
Global wRecebeCarimboAnexo As String
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
Global Wusuario As String
Global WNatureza As String
Global WVendedor As String
Global wCarimbo4 As String
Global wStr1, wStr2, wStr3, wStr4, wStr5, wStr6, wStr7 As String
Global wStr8, wStr9, wStr10, wStr11, wStr12, wStr13, wStr15, wStr16, wStr17, wStr18, wStr19, wStr20, wStr21 As String
Global Wcondicao As String * 30
Global wVerificaLojaOnLine As String * 1
Global arquivo As String
Global BUFFER As String
Global wDescricao As String
Global wLinhaCarimbo As String
Global wReferenciaEspecial As String
Global wSerieProd1 As String
Global wSerieProd2 As String
Global wCarimbo5 As String

Global wPedidoCliente  As Double

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

Global wUltimoItem As Integer

Global WCOMISSAO As Integer
Global WCODOPER As Integer
Global WQTDITEM As Integer
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
Global wpagina As Integer
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
Global condpagto As Double
Global av As Double
Global Cliente As String
Global NATOPER As Double
Global datapag As String
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
Global pessoa As String
Global ufcli As String
Global cepcli As String
Global bairrocli As String
Global tiponota As String
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
Global wCarimbo1 As String
Global wCarimbo2 As String
Global wCarimbo3 As String
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
Global wTelefone As String
Dim RsTravaCaixa As Recordset




Sub Main()
    
    
    
    Call AbilitarCaixa

    'Call Shell("C:\TrocaVersao.bat")
    'Call Shell("C:\pkunzip c:\versao~1\bc2000.zip")
    'FileCopy "C:\bc2000.exe", "C:\Versao\bc2000.exe"
    
    
    
    If ComparaDataVersao(GLB_Versao, GLB_VersaoNova) = True Then
        Call Shell(GLB_TrocaVersao, vbNormalFocus)
        End
    End If
    
    
    Set DBLoja = Workspaces(0).OpenDatabase(WbancoAccess)
    
    Set DBFBanco = Workspaces(0).OpenDatabase(WbancoDbf, False, False, "DBase IV")
    
    '
    '------------------------------Verifica situação do sistema-----------------------
    '
    SQL = ""
    SQL = "Select CT_SituacaoCaixa from ControleECF where CT_SituacaoCaixa='T' and CT_ECF=" & Val(glb_ECF) & ""
        Set RsTravaCaixa = DBLoja.OpenRecordset(SQL)
    If Not RsTravaCaixa.EOF Then
        MsgBox "O sistema está travado por falta de conexão com a central, " _
            & "tire nota manual e entre em contato com o Fernando Alfano para destravar o sistema", vbCritical, "Aviso"
        frmLogin.Show
        Exit Sub
    End If
    '**********************************************************************************

    If wMigracao <> "" Then
       SQL = ""
       SQL = "Select CT_ControleMigracao from Controle where CT_ControleMigracao = 'F' "
           Set RsPegaControleMigracao = DBLoja.OpenRecordset(SQL)
    
       If Not RsPegaControleMigracao.EOF Then
           AbreMigracao = Shell(wMigracao, 1) 'Abre Migracao.exe
           SQL = ""
           SQL = "Update Controle set CT_ControleMigracao = 'A'"
           DBLoja.Execute (SQL)
       End If
    'Else
       'MsgBox "Não Foi Possivel Conectar-se a Migração de Arquivos, Problemas com o Caminho de Conexão " & wMigracao, vbExclamation, "Atenção"
       'Exit Sub
    End If
    
    If wAtualizaCentral <> "" Then
        SQL = ""
        SQL = "Select CT_OnLine from Controle where CT_AtualizacaoCentral = 'F' "
            Set RsAtivaLojaOnLine = DBLoja.OpenRecordset(SQL)
        If Not RsAtivaLojaOnLine.EOF Then
            AbreAtualizacaoOnLine = Shell(wAtualizaCentral, 1) 'Abre AtualizaCental.exe
            SQL = ""
            SQL = "Update Controle set CT_AtualizacaoCentral = 'A' "
                DBLoja.Execute (SQL)
        End If
    'Else
       'MsgBox "Não Foi Possivel Conectar-se em Atualiza Computador Central, Problemas com o Caminho de Conexão " & wMigracao, vbExclamation, "Atenção"
       'Exit Sub
    End If
    
    If GLB_NumeroCaixa = 1 Then
        SQL = ""
        SQL = "Update Controle set CT_Balcao = 'A'"
            DBLoja.Execute (SQL)
    End If
    mdiBalcao.Show

End Sub

Public Function ValidaAbertura()

    Set VerificaCaixa = DBLoja.OpenRecordset("Select max(Ct_Data) as DataMov,max(Ct_Sequencia) as Seq from CTCaixa where CT_NumeroECF = " & Val(glb_ECF) & "")

    If Not VerificaCaixa.EOF Then
       Set PegaDadosCaixa = DBLoja.OpenRecordset("Select * from CTCaixa where ct_data= # " & Format(VerificaCaixa("datamov"), "mm/dd/yyyy") & "# and Ct_Sequencia = " & VerificaCaixa("Seq") & " and CT_NumeroEcf=" & Val(glb_ECF) & "   ")

       If Not PegaDadosCaixa.EOF Then
          If PegaDadosCaixa("ct_Data") = Date And PegaDadosCaixa("ct_Situacao") = "A" Then
             frmCaixa.Show
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
    End If


End Function

Public Function ValidaFechamento()
    Set VerificaCaixa = DBLoja.OpenRecordset("Select max(Ct_Data) as datamov,max(Ct_Sequencia) as Seq from CTCaixa where CT_NumeroECF = " & glb_ECF & "")
    
    If Not VerificaCaixa.EOF Then
       Set PegaDadosCaixa = DBLoja.OpenRecordset("Select * from CTCaixa where ct_data= # " & Format(VerificaCaixa("datamov"), "mm/dd/yyyy") & "# and Ct_Sequencia = " & VerificaCaixa("Seq") & "  and CT_NumeroECF = " & glb_ECF & "")
       
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
           & "TipoNota='V', DataEmi = #" & Format(Date, "MM/DD/YYYY") & "# Where numeroped = " & frmCaixa.txtPedido.Text & ""
       DBLoja.Execute (SQL)
   Else
       SQL = "Update NfItens set nf = " & WNF & ",Serie= '" & Wserie & "'," _
           & "TipoNota='V',DataEmi = #" & Format(Date, "MM/DD/YYYY") & "# Where numeroped = " & WnumeroPed & ""
       DBLoja.Execute (SQL)
   End If
       
End Function


Public Function Atualizanfcapa()

    If frmCaixa.txtPedido.Text <> "" Then
       SQL = "Update NfCapa set nf = " & WNF & ",Serie= '" & Wserie & "', " _
           & "TipoNota='V',hora= #" & Time & " # , DataEmi = #" & Format(Date, "MM/DD/YYYY") & "# Where numeroped = " & frmCaixa.txtPedido.Text & ""
       DBLoja.Execute (SQL)
    Else
       SQL = "Update NfCapa set nf = " & WNF & ",Serie= '" & Wserie & "', " _
           & "TipoNota='V',hora= #" & Time & " #,DataEmi = #" & Format(Date, "MM/DD/YYYY") & "# Where numeroped = " & WnumeroPed & ""
       DBLoja.Execute (SQL)
    End If
'    If Wserie = "SM" And frmCaixa.cmdCodigoZero.Caption = "CO" Then
'        SQL = "Update EvDesDBF set NotaFis = " & WNF & " where NumPed = " & frmCaixa.txtPedido.Text & " "
'            DBLoja.Execute (SQL)
'    End If
End Function

Public Function EmiteCupom()

    SQL = ""
    SQL = "Select ct_ecf,ct_loja from controle"
    
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then

       wLoja = Trim(RsDados("Ct_LOJA"))
       
       If Trim(RsDados("Ct_ECF")) = "S" Then
          WEmiteCupom = True
       End If
          
    End If

End Function


Public Function DadosLoja()

    SQL = ""
    SQL = "Select CT_Loja,CT_Razao,Lojas.* from lojas,Controle where lo_loja=CT_Loja"

    Set RsDados = DBLoja.OpenRecordset(SQL)

    If Not RsDados.EOF Then

       WRazao = RsDados("CT_Razao")
       Wendereco = RsDados("lo_ENDERECO")
       wbairro = RsDados("lo_bairro")
       WCGC = RsDados("lo_CGC")
       WIest = RsDados("lo_INSCRICAOESTADUAL")
       WMunicipio = RsDados("lo_MUNICIPIO")
       westado = RsDados("lo_UF")
       WCep = RsDados("lo_CEP")
       WFone = RsDados("lo_TELEFONE")
       wLoja = RsDados("CT_Loja")
    
    End If


End Function


Public Function ExtraiNumeroCupom()

      wSair = False
      WNF = 0
      WNumeroCupom = 0
      For i = 0 To 5
        Retorno = Bematech_FI_NumeroCupom(WNumeroCupom)
        If Retorno <> 1 Or WNumeroCupom > 0 Then
            Exit For
        End If
      Next i
      Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      WNF = WNumeroCupom
      
     If TipoPedido <> "CriaPedido" Then
        If WNF > 0 Then
           SQL = ""
           SQL = "Update controleECF set ct_ultimocupom= " & WNF & " " _
                & "where CT_Ecf=" & Val(glb_ECF) & ""
           DBLoja.Execute (SQL)
        Else
           wECFNF = 2
           MsgBox "O sistema vai emitir nota fiscal", vbExclamation, "Aviso"
           wSair = True
        End If
      Else
        If Retorno = 1 Then
            wSair = True
        End If
      End If

End Function

Public Function ExtraiSeqPedido()
   
   SQL = "Select * from Controle "
   Set RsDados = DBLoja.OpenRecordset(SQL)
   
   WnumeroPed = RsDados("CT_NumPed") + 1
   
   MsgBox "Extrai nº imp. Fiscal"
   SQL = "Update Controle set CT_Numped=CT_Numped + 1"
   DBLoja.Execute (SQL)
   
End Function

Public Function ExtraiSeqNota()

    SQL = ""
    SQL = "Select Ct_SeqNota from controle"
    Set RsDados = DBLoja.OpenRecordset(SQL)

    If Not RsDados.EOF Then
       WNF = RsDados("Ct_SeqNota") + 1
       SQL = "Update controle set Ct_SeqNota=" & WNF + 1 & " "
       DBLoja.Execute (SQL)
    Else
       MsgBox "Erro ao atualizar sequencia de nota"
    End If

End Function


Public Function GravaNFCapa()
       
       If Trim(wVendedorLojaVenda) = "" Or Trim(wVendedorLojaVenda) = "0" Then
            wVendedorLojaVenda = WVendedor
       End If
       If Trim(wLojaVenda) = "" Then
            wLojaVenda = wLoja
       End If
       If UCase(Glb_SeriePed) <> "S1" And UCase(Glb_SeriePed) <> "D1" Then
            If Wserie <> "SN" And Wserie <> "CT" Then
                Wserie = ""
                Glb_Nf = 0
            End If
       Else
            Wserie = UCase(Glb_SeriePed)
       End If
       On Error Resume Next
       BeginTrans
       SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
            & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
            & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
            & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
            & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF)" _
            & "Values (" & WnumeroPed & ", #" & Format(Wdata, "mm/dd/yyyy") & "#," & WVendedor & ", " _
            & "" & ConverteVirgula(WTotPedido) & "," & ConverteVirgula(WTotPedido) & "," & ConverteVirgula(Format(Wdescontop, "0.00")) & ", " _
            & "" & ConverteVirgula(wSubTotal) & ",'" & wLoja & "'," & WQTDITEM & ", " _
            & "'" & WTipoNota & "','" & wCondPag & "'," & Wav & "," & WCliente & ", " _
            & "" & WCODOPER & ",#" & Format(WdataPag, "dd/mm/yyyy") & "#," & ConverteVirgula(WPGENTRA) & ", " _
            & "'" & Wlojat & "'," & WPesoBr & "," & WPesoLq & ", " _
            & "" & ConverteVirgula(WFRETECOBR) & "," & ConverteVirgula(WFRETECOBR) & ",'" & WOutraLoja & "'," & WOUTROVEND & ", " _
            & "'" & Wserie & "','" & WUF & "','" & WNOMCLI & "','" & WENDCLI & "','" & WCGCCLI & "','" & WMUNCLI & "', " _
            & "" & wPessoa & ",'" & WFone & "',0,'" & WIest & "','" & wbairro & "'," _
            & "'" & WCep & "','" & WDESCRIPAG & "','A'," & ConverteVirgula(Format(wValorTotalCodigoZero, "0.00")) & "," & ConverteVirgula(Format(wTotalNotaAlternativa, "0.00")) & "," & ConverteVirgula(Format(wTotalNotaAlternativa, "0.00")) & ",'" & wCarimbo3 & "','" & WCFOAux & "','" & wLojaVenda & "','" & wVendedorLojaVenda & "', " & wPedidoCliente & "," & Val(glb_ECF) & "," & Glb_Nf & ")"
        
       DBLoja.Execute (SQL)
       If Err.Number = 0 Then
            CommitTrans
            wNfCapa = True
       Else
            Rollback
       End If
            
End Function


Public Sub GravaNfItens()
    On Error Resume Next
    WNF = Glb_Nf
    BeginTrans
    SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
        & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
        & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
        & "Values (" & WnumeroPed & ", #" & Format(Wdata, "mm/dd/yyyy") & "#,'" & Trim(wReferencia) & "', " _
        & "" & wQtde & ", " & ConverteVirgula(wVlUnit) & ", " & ConverteVirgula(wVlUnit2) & ", " _
        & "" & ConverteVirgula(wVlTotItem) & "," & ConverteVirgula(WDESCRAT) & "," _
        & "" & wItem & "," & wLinha & "," & wSecao & "," & WCSPROD & ", " _
        & "" & ConverteVirgula(wPLISTA) & "," & WTRIBUTO & "," & ConverteVirgula(wIcmPdv) & ", " _
        & "'" & wCodBarra & "'," & WNF & ", '" & Wserie & "'," & WCliente & ", " _
        & "'" & WTipoNota & "'," & WVendedor & ",'" & wLoja & "'," & wTipoMovimentacao & ",'A'," & ConverteVirgula(wValorMercadoriaAlternativa) & "," & ConverteVirgula(wValorTotalItemAlternativa) & ",'" & WREFALTERNA & "','" & wPegaDescricaoAlternativa & "','" & wSerieProd1 & "' , '" & wSerieProd2 & "')"
     
     DBLoja.Execute (SQL)
     If Err.Number = 0 Then
        CommitTrans
        wNFitens = True
     Else
        Rollback
     End If
    
End Sub

Public Function EmiteNotafiscal()
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "COTACAO/RESUMO" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next

    
    wNotaTransferencia = False
    wpagina = 1
    If Wserie <> "CT" Then
        WNatureza = "VENDAS"
    Else
        WNatureza = "TRANSFERENCIA"
    End If
    'Temporario = "C:\NOTASVB\"
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
        & "Where NfCapa.nf= " & WNF & " " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
        
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then
           
      Call Cabecalho
      
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & WNF & " order by nfitens.item"

      Set RsdadosItens = DBLoja.OpenRecordset(SQL)

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
                          & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                          & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "########0.00"), 12) & Space(1) _
                          & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "########0.00"), 15) & Space(1) _
                          & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)
            
            Else
                     
                   wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), "0", RsdadosItens("DescricaoAlternativa"))
                   If wPegaDescricaoAlternativa <> "0" Then
                         wDescricao = wPegaDescricaoAlternativa
                   Else
                         wDescricao = Trim(RsdadosItens("pr_descricao"))
                   End If
                   
                   wStr16 = ""
                   wStr16 = Left$(RsdadosItens("pr_referencia") & Space(10), 10) _
                         & Space(2) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
                         & Space(16) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                         & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
                         & "" & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                         & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                         & Right$(Space(12) & Format(RsdadosItens("vlunit"), "########0.00"), 12) & Space(1) _
                         & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "########0.00"), 12) & Space(1) _
                         & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)

                                  
            End If
                      
                      'On Error Resume Next
                      Printer.Print Space(2) & wStr16
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
                            Printer.Print ""
                         Loop
                         RsdadosItens.MoveNext
                         wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                         Printer.Print wStr13
                         Printer.Print ""
                         Printer.Print ""
                         'Printer.Print Chr(18)  'Finaliza Impressão
                         'Close #Notafiscal
                         
                         wConta = 0
                         wpagina = wpagina + 1
                         'FileCopy Temporario & NomeArquivo, "S:\notasvb\" & NomeArquivo
'                         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
                         Printer.EndDoc
                         Cabecalho
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
'    Set RsDados = DBLoja.OpenRecordset(SQL)
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
'           Wentrada = Format(RsDados("Pgentra"), "########0.00")
'           wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
'        End If
'
'        wStr1 = Space(2) & Left$(Format(wStr17) & Space(50), 50) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(2) & "X" & Space(31) & Left$(Format(RsDados("nf"), "###,###"), 7)
'        wStr2 = Space(2) & Left$(Format(wStr18) & Space(50), 50) & Left$(Format(Trim(WMunicipio), ">") & Space(15), 15) & Space(29) & Left$(Trim(westado), 2)
'        wStr3 = Space(2) & Left$(Format(wStr19) & Space(50), 50) & "(011)" & Left$(Trim(Format(WFone, "####-####")), 9) & "/(011)" & Left$(Format(WFone, "####-####"), 9) & Space(11) & Left$(Format((WCep), "#####-###"), 9)
'        wStr4 = Space(2) & Left$(Format(wStr19) & Space(100), 100) & Left$(Trim(Format(WCGC, "###,###,###")), 10) & "/" & Format(Mid((WCGC), 11, 5), "####-##")
'        wStr5 = Space(44) & Trim(WNatureza) & Space(24) & Left$(RsDados("CFOAUX"), 10) & Space(27) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
'        wStr6 = Space(44) & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(17) & Left$(Trim(Format(RsDados("CGCCLI"), "###,###,###")), 10) & "/" & Right$(Format(RsDados("CGCCLI"), "####-##"), 7) & Space(5) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr7 = Space(44) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(16) & Left$(RsDados("CEPCLI"), 11) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
'        wStr8 = Space(44) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(43) & Left$(Trim(RsDados("UFCLIENTE")), 9) & Space(14) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
'
'
''        wStr6 = Space(40) & Left$(Format(Trim(rdorsExtra2("em_descricao")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(rdorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & Right$(Format(rdorsExtra2("lo_cgc"), "####-##"), 7) & Space(5) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr7 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(15), 15) & Space(32) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
''        wStr8 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(43) & Left$(Trim(rdorsExtra2("lo_uf")), 9) & Space(14) & Left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
'
'        wStr9 = Space(4) & Right$(Space(12) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
'        wStr10 = Space(67) & Right(Space(12) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
'        wStr11 = Space(2) & "                          "
'        wStr12 = Space(2) & "                                                     "
'        wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(13) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
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
'          Set RsdadosItens = DBLoja.OpenRecordset(SQL)
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
'                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("vlunit"), "########0.00"), 12) & Space(2) _
'                             & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "########0.00"), 15) & Space(2) _
'                             & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)
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

    
End Function



Private Sub FinalizaNota()
       If wNotaTransferencia = False Then
         If wReferenciaEspecial <> "" Then
             SQL = ""
             SQL = "Select * from CarimbosEspeciais " _
                & "where CE_Referencia='" & wReferenciaEspecial & "'"
                Set RsPegaItensEspeciais = DBLoja.OpenRecordset(SQL)
                
             If Not RsPegaItensEspeciais.EOF Then
                i = 0
        
                If RsPegaItensEspeciais("CE_Linha1") <> "" Then
                    wConta = wConta + 7
                    'Print #Notafiscal, ""
                    If Trim(RsPegaItensEspeciais("CE_Linha5")) = "" Then
                        Printer.Print Space(15) & "______________________________________________________________"
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Printer.Print Space(17) & "___________________________________     ____/____/______   "
                        Printer.Print Space(17) & "            Assinatura                        Data         "
                        'Print #Notafiscal, Space(15) & "____________________________________________________________"
                    Else
                        Printer.Print Space(15) & "______________________________________________________________"
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha2"), 60)
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha3"), 60)
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha4"), 60)
                        Printer.Print Space(16) & Right(RsPegaItensEspeciais("CE_Linha5"), 60)
                        Printer.Print Space(17) & "___________________________________     ____/____/______   "
                        Printer.Print Space(17) & "            Assinatura                        Data         "
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
        Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(120), 120)
     ElseIf RsDados("Carimbo1") <> "" And RsDados("Desconto") <> 0 Then
        Printer.Print Space(1) & Left(RsDados("Carimbo1") & Space(110), 110) & Left("Desconto" & Space(12), 12) & Left(Format(RsDados("Desconto"), "0.00") & Space(10), 10)
     ElseIf RsDados("Carimbo1") <> "" Then
        Printer.Print Space(1) & RsDados("Carimbo1")
     ElseIf RsDados("Desconto") <> 0 And Wsm = False Then
        Printer.Print Space(90) & "Desconto" & Space(13) & Format(RsDados("Desconto"), "0.00")
     Else
        Printer.Print ""
     End If
     If RsDados("Carimbo2") <> "" Then
        Printer.Print Space(4) & RsDados("Carimbo2")
     End If
     
     wConta = wConta + 1
     
     If (IIf(IsNull(RsDados("Carimbo5")), "", RsDados("Carimbo5"))) <> "" Then
        Printer.Print Space(4) & RsDados("Carimbo5")
     Else
        Printer.Print ""
     End If
        
     Do While wConta < 12
        wConta = wConta + 1
        Printer.Print ""
     Loop

     If Wsm = True Then
        Printer.Print ""
        Printer.Print ""
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
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        Printer.Print wStr9
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "########0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        Printer.Print wStr10
     Else
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(36) & Right$(Space(15) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
        Printer.Print wStr9
        Printer.Print ""
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "########0.00"), 12) & Space(50) & Right(Space(12) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
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
     wStr13 = Space(90) & "Lj " & RsDados("LojaOrigem") & Space(14) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
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
     wTotalNotaTransferencia = RsDados("VlrMercadoria")
     If wReemissao = False Then
        SQL = "Select * from CtCaixa order by CT_Data desc"
           Set rsPegaLoja = DBLoja.OpenRecordset(SQL)
        If Not rsPegaLoja.EOF Then
           If WNatureza = "TRANSFERENCIAS" Then
               SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & rsPegaLoja("ct_operador") & "','" & rsPegaLoja("ct_loja") & "', " _
                   & " #" & Format(rsPegaLoja("ct_data"), "mm/dd/yyyy") & "#, " & 20109 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   DBLoja.Execute (SQL)
           'ElseIf WNatureza = "DEVOLUCAO" Then
               'SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & RsPegaLoja("ct_operador") & "','" & RsPegaLoja("ct_loja") & "', " _
                   & " #" & Format(RsPegaLoja("ct_data"), "mm/dd/yyyy") & "#, " & 20201 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   'DBLoja.Execute (SQL)
           End If
        End If
    End If
       
End Sub


Private Sub Finaliza()

    flg = 0
    wlin = 99
    Screen.MousePointer = 0

End Sub


Public Function ExtraiLoja()

    SQL = "Select Ct_loja From controle"
    Set RsDados = DBLoja.OpenRecordset(SQL)
    If Not RsDados.EOF Then
       wLoja = RsDados("Ct_Loja")
    Else
       MsgBox "Problemas no controle"
       Exit Function
    End If

End Function
    
    
Public Function ExtraiDataMovimento()
    
    SQL = "Select max(Ct_Data) as WdataMax From Ctcaixa"
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then
       Wdata = RsDados("WdataMax")
       
       'Wdata = Mid(Isql("WdataMax"), 1, 2) & Mid(Isql("WdataMax"), 4, 2) & Mid(Isql("WdataMax"), 7, 2)
       'Wdate = Format(Wdata, "dd,mm,yyyy")
    Else
       MsgBox "Problemas no Ctcaixa"
       Exit Function
    End If
    
End Function


Public Function ExtraiSeqNotaControle()

     Dim WnovaSeqNota As Long
        
     
     SQL = ""
     SQL = "Select * from controle"
     Set RsDados = DBLoja.OpenRecordset(SQL)
     
     If Not RsDados.EOF Then
        WnumeroNotaDbf = 0
        WnovaSeqNota = 0
        
        WnumeroNotaDbf = RsDados("CT_SeqNota") + 1
        WnovaSeqNota = WnumeroNotaDbf
        WNF = WnumeroNotaDbf
    
        SQL = "update controle set CT_SeqNota= " & WnovaSeqNota & ""
        DBLoja.Execute (SQL)
        
         
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



Function EncerraVenda(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double)
        
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
        wCFOItem = 0
        wUltimoItem = 0
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wCarimbo5 = ""
'
'  --------------------------------- CALCULO DO ICMS ------------------------------------------------------------------------
'
        
        SQL = "Select nfcapa.*, Estados.* from nfcapa, Estados " _
              & "where nfcapa.numeroped = " & NumeroDocumento & "" _
              & "And nfcapa.ufCliente = Estados.UF_Estado"
              Set RsCapaNF = DBLoja.OpenRecordset(SQL)
        
        If Not RsCapaNF.EOF Then
               wConfereCodigoZero = IIf(IsNull(RsCapaNF("ValorTotalCodigoZero")), 0, RsCapaNF("ValorTotalCodigoZero"))
               If Trim(RsCapaNF("TipoNota")) = "E" Then
                    wPessoa = RsCapaNF("PessoaCli")
               End If
               If RsCapaNF("TM") <> 1 Then
                    wECFNF = 2
                    wNumeroECF = 2
                    wChaveICMS = RsCapaNF("UF_Regiao") & wPessoa
                    If RsCapaNF("UFCliente") = "SP" Then
                        If wPessoa = 2 Then
                           If WEmiteCupom = True Then
                              wECFNF = 1
                              'SQL = ""
                              'SQL = "Select CT_NumeroECF from Controle"
                              '  Set RsNumeroECF = DBLoja.OpenRecordset(SQL)
                              'If Not RsNumeroECF.EOF Then
                              wNumeroECF = glb_ECF
                              'Else
                              '  MsgBox "Nenhum Numero de ecf encontrado", vbCritical, "Atenção"
                              'End If
                           Else
                              wNumeroECF = glb_ECF
                              wECFNF = 2
                           End If
                        End If
                    End If
               End If
               If RsCapaNF("Serie") <> "S1" And RsCapaNF("Serie") <> "D1" Then
                    wECFNF = wECFNF
                    Wserie = ""
                    WNF = ""
               Else
                    Wserie = RsCapaNF("Serie")
                    wECFNF = 2
                    WNF = RsCapaNF("NF")
                    WEmiteCupom = False
               End If
        Else
            MsgBox "Nota não encontrada", vbInformation, "Atenção"
            Exit Function
        End If
                    
          SQL = "Select produto.*, nfitens.* from produto,nfitens " _
              & "where nfitens.numeroped = " & NumeroDocumento & "" _
              & "and pr_referencia = nfitens.referencia order by NfItens.Item"
              Set RsItensNF = DBLoja.OpenRecordset(SQL)
          
          If Not RsItensNF.EOF Then
             Do While Not RsItensNF.EOF
               If RsCapaNF("TM") <> 1 Then
                     wChaveICMSItem = wChaveICMS
                    If RsItensNF("PR_substituicaotributaria") = "S" Then
                        wSubstituicaoTributaria = 1
                    Else
                        wSubstituicaoTributaria = 0
                    End If
                    'If RsCapaNF("UFCliente") <> "SP" Then
                        wChaveICMSItem = wChaveICMSItem & RsItensNF("pr_icmssaida") & RsItensNF("pr_codigoreducaoicms") & wSubstituicaoTributaria
                        Call AcharICMSInterEstadual
                        GLB_AliquotaAplicadaICMS = RsICMSInter("IE_icmsAplicado")
                        GLB_AliquotaICMS = RsICMSInter("IE_IcmsDestino")
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
                        GLB_ValorCalculadoICMS = Format(((RsItensNF("ValorMercadoriaAlternativa") * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                        GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
                        If GLB_TotalICMSCalculado > 0 Then
                            If RsICMSInter("IE_BasedeReducao") = 0 Then
                                If GLB_AliquotaAplicadaICMS = 0 Then
                                    GLB_BasedeCalculoICMS = 0
                                Else
                                    GLB_BasedeCalculoICMS = RsItensNF("ValorMercadoriaAlternativa")
                                End If
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
                    WAnexoAux = ""
                    If RsItensNF("pr_codigoreducaoicms") <> 0 Then
                        WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "##0")
                    End If
                    
                    If RsCapaNF("TipoNota") <> "E" Then
                        If wCFOItem = 5102 Or wCFOItem = 6102 Then
                            wCFO1 = wCFOItem
                        ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                            wCFO2 = wCFOItem
                        End If
                    Else
                        If wCFOItem = 5102 Or wCFOItem = 6102 Then
                            wCFO1 = wCFOItem
                            If wCFO1 = 5102 Then
                                wCFO1 = 1202 'Devolucao dentro estado
                            Else
                                wCFO1 = 2202 'Devolucao p/ fora do estado
                            End If
                        ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                            wCFO2 = wCFOItem
                            If wCFO1 = 5405 Then
                                wCFO1 = 1202 'Devolucao dentro estado
                            Else
                                wCFO1 = 2202 'Devolucao p/ fora do estado
                            End If
                        End If

                    End If
               Else
                    wVerificaTM = True
               End If
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
                'If RsCapaNF("TM") <> 1 Then
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
                                    wCarimbo5 = Trim(wCarimbo5) & Trim(RsItensNF("Item")) & "-" & Trim(RsItensNF("SerieProd1")) & "/" & Trim(RsItensNF("SerieProd2")) & ","
                                End If
                            Else
                                If RsItensNF("SerieProd2") <> "0" Then
                                    wCarimbo5 = Trim(wCarimbo5) & Trim(RsItensNF("Item")) & "-" & Trim(RsItensNF("SerieProd1")) & ","
                                End If
                            End If
                        End If
                    Else
                        wCarimbo5 = Trim(RsCapaNF("Carimbo5"))
                    End If
                    
                    SQL = "UPDATE nfitens set baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & ", " _
                    & "Comissao = " & ConverteVirgula(Format(wComissaoVenda, "#0.00")) & ", " _
                    & "IcmPDV = " & ConverteVirgula(Format(RsICMSInter("IE_icmsdestino"), "0.00")) & ", " _
                    & "Bcomis = " & ConverteVirgula(RsItensNF("PR_PercentualComissao")) & " , " _
                    & "DetalheImpressao = '" & wDetalheImpressao & "' " _
                    & " where nfitens.numeroped = " & NumeroDocumento & "" _
                    & " and Referencia = '" & RsItensNF("PR_Referencia") & "' and Item=" & RsItensNF("Item") & ""
                    DBLoja.Execute (SQL)
                    
                    
                    'wUltimoItem = RsItensNF("Item")
                'End If
                    
                    SQL = ""
                    SQL = "Update nfitens set Comissao = " & ConverteVirgulaNegativa(Format(wComissaoVenda, "#0.00")) & " " _
                        & " where nfitens.numeroped = " & NumeroDocumento & "" _
                        & " and Referencia = '" & RsItensNF("PR_Referencia") & "'"
                        DBLoja.Execute (SQL)
                
                RsItensNF.MoveNext
             Loop
        End If
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
'
        If RsCapaNF("TM") <> 1 Then
            SQL = "Select CA_Descricao from CarimboNotaFiscal where CA_CodigoCarimbo = 1"
                Set RsCarimbo = DBLoja.OpenRecordset(SQL)
                
                If Not RsCarimbo.EOF Then
                    If wAnexo1 <> "" Or wAnexo2 <> "" Then
                        wPegaCarimboNF = RsCarimbo("CA_Descricao")
                        wRecebeCarimboAnexo = Mid(wPegaCarimboNF, 1, 79) & wAnexo1 & Mid(wPegaCarimboNF, 80, Len(wPegaCarimboNF)) & wAnexo2
                    Else
                        wRecebeCarimboAnexo = ""
                    End If
                End If
             
             'wUltimoItem = ((wUltimoItem / 10) + 0.9)
             If Trim(wCarimbo5) <> "" And RsCapaNF("TipoNota") <> "E" Then
                wCarimbo5 = "SERIE(S) ITEM(S):" & wCarimbo5
             End If
             SQL = "UPDATE nfcapa set baseicms = " & ConverteVirgula(GLB_BaseTotalICMS) & ", " _
                & "Vlricms = " & ConverteVirgula(GLB_TotalICMSCalculado) & ", " _
                & "cfoaux = '" & wCFO1 & wCFO2 & "', " _
                & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", " _
                & "Pessoacli = " & wPessoa & ", " _
                & "ECFNF = " & glb_ECF & ", " _
                & "Regiaocli = " & RsCapaNF("UF_Regiao") & ", " _
                & "Carimbo1 = '" & wRecebeCarimboAnexo & "', " _
                & "Carimbo5 = '" & Trim(wCarimbo5) & "' " _
                & "where nfcapa.numeroped = " & NumeroDocumento & ""
                DBLoja.Execute (SQL)
                            
                'GravaSequenciaLeitura 5, NumeroDocumento, SerieDocumento
        End If
' -------------------------------------- ATUALIZA MARGEM DE VENDA ---------------------------------------------------
'
                SQL = "UPDATE vende set VE_totalvenda = VE_totalVenda + " & ConverteVirgula(wSomaVenda) & ", " _
                & "VE_MargemVenda = VE_MargemVenda + " & ConverteVirgula(wSomaMargem) & " " _
                & "where VE_Codigo = " & RsCapaNF("Vendedor") & " "
                DBLoja.Execute (SQL)


'
' -------------------------------------- ATUALIZA CONTORLE DE OPERAÇÂO ---------------------------------------------------
'

                SQL = "UPDATE CTcaixa set Ct_operacoes = Ct_operacoes + 1 " _
                & "where ct_situacao = 'A' "
                DBLoja.Execute (SQL)
          
     
            
    
End Function


Sub AcharICMSInterEstadual()
    
    SQL = "SELECT * from IcmsInterEstadual where IE_Codigo = " & wChaveICMSItem
    Set RsICMSInter = DBLoja.OpenRecordset(SQL)
    
    If RsICMSInter.EOF Then
        MsgBox "ICMS inter estadual não encontrado", vbInformation, "Aviso"
        'Exit Sub
    End If
        
End Sub


Function GravaSequenciaLeitura(ByVal CodigoTabela As Double, SequenciaGravacao As Double, ByVal SerieDocumento As String)
  
'  SQL = "Insert into SequenciaGravacao(SL_CodigoTabela, SL_sequenciaGravacao, SL_Serie, SL_Situacao,SL_Data) " _
'        & "Values (" & CodigoTabela & ", " & SequenciaGravacao & ", " _
'        & "'" & SerieDocumento & "', 'A',#" & Format(Date, "DD/MM/YYYY") & "#) "
'  DBLoja.Execute (SQL)

End Function


Sub SelecionaMovimentoCaixa()
    'SQL = ""
    'SQL = "Select MC_Sequencia,MC_Remessa from MovimentoCaixa " _
        & "where MC_Remessa = 9 "
    'Set RsSelecionaMovCaixa = DBLoja.OpenRecordset(SQL)
    
    'If Not RsSelecionaMovCaixa.EOF Then
        'Do While Not RsSelecionaMovCaixa.EOF
            'GravaSequenciaLeitura 4, RsSelecionaMovCaixa("MC_Sequencia"), 0
            'RsSelecionaMovCaixa.MoveNext
        'Loop
    'End If
            
    'SQL = ""
    'SQL = "Update MovimentoCaixa set MC_Remessa = 0 " _
         & "Where MC_Remessa = " & 9 & " and MC_Grupo <> " & 10101 & " and MC_Grupo <> " & 30101
    'DBLoja.Execute (SQL)
    
    'SQL = "Update MovimentoCaixa set MC_Remessa = 1 " _
         & "Where MC_Remessa = " & 9 & " and MC_Grupo = " & 10101 & " or MC_Grupo = " & 30101
    'DBLoja.Execute (SQL)
End Sub

Sub SelecionaMovimentoBancario()
    SQL = ""
    SQL = "Select MB_Sequencia,MB_TipoMovimentacao from MovimentoBancario " _
        & "Where MB_TipoMovimentacao = 9 "
        Set RsSelecionaMovBanco = DBLoja.OpenRecordset(SQL)
    
    If Not RsSelecionaMovBanco.EOF Then
        Do While Not RsSelecionaMovBanco.EOF
            GravaSequenciaLeitura 3, RsSelecionaMovBanco("MB_Sequencia"), 0
            RsSelecionaMovBanco.MoveNext
        Loop
    End If

    SQL = "Update MovimentoBancario set MB_TipoMovimentacao = 0 " _
         & "Where MB_TipoMovimentacao = " & 9
    DBLoja.Execute (SQL)

End Sub

    
Sub SelecionaMovimentoEstoque()

    SQL = ""
    SQL = "Select ME_Sequencia,ME_Situacao from MovimentacaoEstoque " _
        & "Where ME_Situacao = '9' "
    Set RsSelecionaMovEstoque = DBLoja.OpenRecordset(SQL)
    
    If Not RsSelecionaMovEstoque.EOF Then
        Do While Not RsSelecionaMovEstoque.EOF
            GravaSequenciaLeitura 2, RsSelecionaMovEstoque("ME_Sequencia"), 0
            RsSelecionaMovEstoque.MoveNext
        Loop
    End If
    
    SQL = "Update MovimentacaoEstoque set ME_Situacao = 0 " _
         & "Where ME_Situacao = '" & 9 & "' "
    DBLoja.Execute (SQL)
    
End Sub



Sub SelecionaDivergenciaEstoque()
    SQL = ""
    SQL = "Select DE_Sequencia from DivergenciaEstoque order by DE_sequencia desc"
    Set RsSelecionaDivEstoque = DBLoja.OpenRecordset(SQL)
    
    If Not RsSelecionaDivEstoque.EOF Then
        GravaSequenciaLeitura 1, RsSelecionaDivEstoque("DE_Sequencia"), 0
    End If
End Sub


Function EncerraVendaMigracao(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double)

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
Dim RsCarimbo As Recordset
Dim wPegaCarimboNF As String
Dim wRecebeCarimboAnexo As String
Dim wConfereCodigoZero As String
Dim wECFNF As Double
Dim wPessoa As Double

Dim wSubstituicaoTributaria As Double
Dim wAnexoIten As String
Dim WAnexoAux As String

        
        
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
              Set RsCapaNF = DBLoja.OpenRecordset(SQL)
        
            
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
              Set RsItensNF = DBLoja.OpenRecordset(SQL)
          
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
                        Call AcharICMSInterEstadual
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
                    WAnexoAux = ""
                    If RsItensNF("pr_codigoreducaoicms") <> 0 Then
                        WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "##0")
                    End If
                    
                    If wCFOItem = 5102 Or wCFOItem = 6102 Then
                        wCFO1 = wCFOItem
                    ElseIf wCFOItem = 5405 Or wCFOItem = 6405 Then
                        wCFO2 = wCFOItem
                    End If
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
'        Set RsCapaNf = dbloja.OpenRecordset(SQL)
        
        
'        SQL = ""
'        SQL = "Select produto.*, nfitens.* from produto,nfitens " _
'              & "where nfitens.numeroped = " & NumeroDocumento & "" _
'              & "and nfitens.serie = 'SN' " _
'              & "and pr_referencia = nfitens.referencia "
'        Set RsItensNf = dbloja.OpenRecordset(SQL)
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
            & "Comissao = " & ConverteVirgula(Format(wComissaoVenda, "#0.00")) & ", " _
            & "Icms = " & RsICMSInter("IE_icmsdestino") & ", " _
            & "Bcomis = " & ConverteVirgula(RsItensNF("PR_PercentualComissao")) & " , " _
            & "DetalheImpressao = '" & wDetalheImpressao & "' " _
            & " where nfitens.numeroped = " & NumeroDocumento & "" _
            & " and Referencia = '" & RsItensNF("PR_Referencia") & "'"
            DBLoja.Execute (SQL)
                    
                    
            'wUltimoItem = RsItensNF("Item")
           
           
'        SQL = "UPDATE nfitens set DetalheImpressao = '" & wDetalheImpressao & "' " _
'            & " where nfitens.numeroped = " & NumeroDocumento & "" _
'            & " and Referencia = '" & RsItensNf("PR_Referencia") & "'"
'            dbloja.Execute (SQL)
            
        
           
           'If RsCapaNf("CODOPER") <> "522" Then
            
'                wComissaoVenda = 0
'                wComissaoVenda = (RsItensNf("VLUNIT2") * RsItensNf("pr_percentualcomissao") / 100)
'
'                wSomaVenda = wSomaVenda + RsItensNf("VLUNIT2")
'                wSomaMargem = wSomaMargem + (RsItensNf("VLUNIT2") - (RsItensNf("pr_precocusto1") * RsItensNf("qtde")))
'
'                SQL = ""
'                SQL = "Update nfitens set Comissao = " & ConverteVirgula(Format(wComissaoVenda, "#0.00")) & " " _
'                    & " where nfitens.numeroped = " & NumeroDocumento & "" _
'                    & " and Referencia = '" & RsItensNf("PR_Referencia") & "'"
'                dbloja.Execute (SQL)
        
' -------------------------------------- ATUALIZA MARGEM DE VENDA ---------------------------------------------------

                SQL = "UPDATE vende set VE_totalvenda = VE_TotalVenda + " & ConverteVirgula(wSomaVenda) & ", " _
                    & "VE_MargemVenda = VE_MargemVenda + " & ConverteVirgula(wSomaMargem) & " " _
                    & "where VE_Codigo = " & RsCapaNF("Vendedor") & " "
                DBLoja.Execute (SQL)
           
           'else
           'End If
            
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
'
                SQL = "Select CA_Descricao from CarimboNotaFiscal where CA_CodigoCarimbo = 1"
                Set RsCarimbo = DBLoja.OpenRecordset(SQL)
                
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
                    DBLoja.Execute (SQL)
                
                
'                SQL = "UPDATE nfcapa set " _
'                    & "Paginanf = " & ConverteVirgula(wUltimoItem) & ", " _
'                    & "Carimbo1 = '" & wRecebeCarimboAnexo & "' " _
'                    & "where nfcapa.numeroped = " & NumeroDocumento & ""
'                dbloja.Execute (SQL)
           
           'End If
        

' -------------------------------------- ATUALIZA CONTROLE DE OPERAÇÂO ---------------------------------------------------
         

           SQL = "UPDATE CTcaixa set Ct_operacoes = Ct_operacoes + 1 " _
               & "where ct_situacao = 'A' "
           DBLoja.Execute (SQL)
          
           RsItensNF.MoveNext
      Loop
    End If

End Function


Public Sub LeituraZ()

    Retorno = Bematech_FI_ReducaoZ("", "")
    Call VerificaRetornoImpressora("", "", "Redução Z")
    If Retorno = 1 Then
        Call AtualizaNumeroCupom
    End If
    
End Sub




Public Sub EmiteCodigoZero()
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "CODIGO ZERO" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
   
    Printer.Print
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 10cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Screen.MousePointer = 11
    SQL = ""
    SQL = "Select NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI," _
        & "NFCAPA.UFCLIENTE,NFCapa.Vendedor,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT2,NFITENS.VLUNIT,NFITENS.DescricaoAlternativa," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
        & "From NFCAPA INNER JOIN NFITENS " _
        & "on (NfCapa.nf=Nfitens.nf)  " _
        & "Where NfCapa.nf= " & WNF & " and NfCapa.Serie = '" & Wserie & "' and NfItens.Serie=NfCapa.Serie " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
    Set RsDados = DBLoja.OpenRecordset(SQL)


    If Not RsDados.EOF Then
        WVendedor = RsDados("Vendedor")
        wTotalPed = Format(RsDados("TotalNota"), "0.00")
        wDesconto = 0
        wDesconto = RsDados("Desconto")
        Printer.Print
        Call CabecalhoCodigoZero
        Do While Not RsDados.EOF
            wDescricao = ""
            wPegaDescricaoAlternativa = "0"
            wPegaDescricaoAlternativa = IIf(IsNull(RsDados("DescricaoAlternativa")), "0", RsDados("DescricaoAlternativa"))
            
            SQL = ""
            SQL = "Select PR_Descricao from Produto where PR_Referencia = '" & RsDados("Referencia") & "'"
                Set RsDescProduto = DBLoja.OpenRecordset(SQL)
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
        Printer.Print "  Valor Recebido R$                    " & Format(wTotalPed, "0.00")
        Printer.Print "  Troco   R$                           0,00"
        Printer.Print "________________________________________________"
        Printer.Print
        Printer.Print "Vendedor  " & WVendedor
        If Trim(wLoja) <> "800" Then
            Printer.Print "DE MEO a mais de 106 anos vendendo qualidade"
        Else
            Printer.Print "DM Motores sempre o melhor preço "
        End If
        Printer.Print "------------------------------------------------"
        Printer.Print "          AGRADECEMOS A PREFERENCIA"
        Printer.Print "------------------------------------------------"
        
        Printer.EndDoc
    End If
    Screen.MousePointer = 0


End Sub



Sub CabecalhoCodigoZero()
    SQL = ""
    SQL = "Select Lojas.*,CT_Loja,CT_SeqC0,CT_Razao from Lojas,Controle where LO_Loja = CT_Loja"
        Set rsPegaLoja = DBLoja.OpenRecordset(SQL)
    If Not rsPegaLoja.EOF Then
        
        
        Printer.ScaleMode = vbMillimeters
        Printer.ForeColor = "0"
        Printer.FontSize = 8
        Printer.FontName = "draft 10cpi"
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.DrawWidth = 3
        
        Printer.Print "      " & rsPegaLoja("CT_Razao")
        Printer.Print "      CNPJ " & rsPegaLoja("LO_CGC") & "  IE " & rsPegaLoja("LO_InscricaoEstadual")
        Printer.Print "      " & rsPegaLoja("LO_Endereco")
        Printer.Print "      Telefone : (" & rsPegaLoja("LO_DDD") & ")" & rsPegaLoja("LO_Telefone")
        Printer.Print "      " & Format(Date, "DD/MM/YYYY") & "   " & Format(Time, "hh:mm") & " NUMERO: " & wPegaSequenciaCO
        Printer.Print
        Printer.Print "    CONTROLE INTERNO-CUPOM SEM VALOR FISCAL"
        Printer.Print "================================================"
        Printer.Print " CÓDIGO                           DESCRIÇÃO"
        Printer.Print "   QTDxUNITARIO                       VALOR(R$)"
        Printer.Print "________________________________________________"
    End If
        
        
End Sub

Public Sub ExtraiSequenciaCodigoZero()
    SQL = ""
    SQL = "Select Lojas.*,CT_Loja,CT_SeqC0 from Lojas,Controle where LO_Loja = CT_Loja"
        Set rsPegaLoja = DBLoja.OpenRecordset(SQL)
    If Not rsPegaLoja.EOF Then
        wPegaSequenciaCO = Val(rsPegaLoja("CT_SeqC0") + 1)
        SQL = ""
        SQL = "Update Controle set CT_SeqC0 = " & wPegaSequenciaCO
            DBLoja.Execute (SQL)
    End If
End Sub


Public Sub EmiteNotaFiscalSM()
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "NOTA FISCAL" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next

    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 10cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Screen.MousePointer = 11
    wlin = 99
        
    WNatureza = "VENDAS"
            
    Call DadosLoja
            
    SQL = ""
    SQL = "Select NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NfCapa.TotalNotaAlternativa,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.DataPag,NfCapa.CondPag," _
        & "NFCAPA.UFCLIENTE,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
        & "From NFCAPA INNER JOIN NFITENS " _
        & "on (NfCapa.nf=Nfitens.nf) " _
        & "Where NfCapa.nf= " & WNF & " " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
        
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then
        If RsDados("CondPag") = 85 Then
            wCarimbo4 = RsDados("DataPag")
        Else
            wCarimbo4 = RsDados("Carimbo4")
        End If
        
        tmporient = Printer.Orientation
        wConta = 0
        wChave = 0
        wReduz = 0
        wStr15 = ""
        wStr17 = ""
        wStr18 = ""
        wStr19 = ""
        wStr20 = ""
        
        If Val(RsDados("CONDPAG")) = 1 Then
           Wcondicao = "Avista"
        ElseIf Val(RsDados("CONDPAG")) = 3 Then
           Wcondicao = "Financiada"
        ElseIf Val(RsDados("CONDPAG")) > 3 Then
           Wcondicao = "Faturada " & wCarimbo4
        End If
        
        wStr17 = "Pedido        : " & RsDados("NUMEROPED")
        wStr18 = "Vendedor      : " & RsDados("VENDEDOR")
        wStr19 = "Cond. Pagto   : " & Trim(Wcondicao)
        
        If RsDados("Pgentra") <> 0 Then
           Wentrada = Format(RsDados("Pgentra"), "########0.00")
           wStr20 = "Entrada       : " & Wentrada
        End If
        
        wStr1 = Space(2) & Left$(Format(wStr17) & Space(50), 50) & Left$(Format(Trim(Wendereco), ">") & Space(30), 30) & Space(7) & Left$(Format(Trim(wbairro), ">") & Space(18), 15) & Space(2) & "X" & Space(31) & Left$(Format(RsDados("nf"), "###,###"), 7)
        wStr2 = Space(2) & Left$(Format(wStr18) & Space(50), 50) & Left$(Format(Trim(WMunicipio), ">") & Space(15), 15) & Space(29) & Left$(Trim(westado), 2)
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(50), 50) & "(011)" & Left$(Trim(Format(WFone, "###-####")), 8) & "/(011)" & Left$(Format(WFone, "###-####"), 8) & Space(11) & Left$(Format((WCep), "#####-###"), 8)
        wStr4 = Space(2) & Left$(Format(wStr20) & Space(100), 100) & Left$(Trim(Format(WCGC, "###,###,###")), 10) & "/" & Format(Mid((WCGC), 11, 5), "####-##")
        wStr5 = Space(40) & Trim(WNatureza) & Space(15) & Left$(RsDados("CFOAUX"), 10) & Space(40) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        wStr6 = Space(40) & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(RsDados("CGCCLI"), "###,###,###")), 10) & "/" & Right$(Format(RsDados("CGCCLI"), "####-##"), 7) & Space(5) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        wStr7 = Space(40) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(12) & Left$(RsDados("CEPCLI") & Space(16), 16) & Space(4) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        wStr8 = Space(40) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(43) & Left$(Trim(RsDados("UFCLIENTE")), 9) & Space(14) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
       

'        wStr6 = Space(40) & Left$(Format(Trim(rdorsExtra2("em_descricao")), ">") & Space(50), 50) & Space(21) & Left$(Trim(Format(rdorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & Right$(Format(rdorsExtra2("lo_cgc"), "####-##"), 7) & Space(5) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
'        wStr7 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(15), 15) & Space(32) & Left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
'        wStr8 = Space(40) & Left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(43) & Left$(Trim(rdorsExtra2("lo_uf")), 9) & Space(14) & Left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
               
        wStr9 = Space(4) & Right$(Space(12) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        wStr10 = Space(67) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        wStr11 = Space(2) & "                          "
        wStr12 = Space(2) & "                                                     "
        wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(13) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                  
        Printer.ScaleMode = vbMillimeters
        Printer.ForeColor = "0"
        Printer.FontSize = 8
        Printer.FontName = "draft 10cpi"
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.DrawWidth = 3
        wpagina = 1
        
        Call Cabecalho
          
          SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
              & "produto.pr_classefiscal,produto.pr_unidade, " _
              & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.ReferenciaAlternativa, " _
              & "nfitens.vlunit,nfitens.vltotitem,NfItens.ValorMercadoriaAlternativa,NfItens.PrecoUnitAlternativa,nfitens.icms,nfitens.detalheImpressao " _
              & "from produto,nfitens " _
              & "where produto.pr_referencia=nfitens.referencia " _
              & "and nfitens.nf = " & WNF & ""
              
          Set RsdadosItens = DBLoja.OpenRecordset(SQL)
    
          If Not RsdadosItens.EOF Then
             Do While Not RsdadosItens.EOF
                      wPegaDescricaoAlternativa = IIf(IsNull(RsDados("Referencia")), "0", RsDados("Referencia"))
                    SQL = ""
                    SQL = "Select Desc from EvDesDBF where NotaFis = " & WNF & " " _
                        & "And Ref = '" & wPegaDescricaoAlternativa & "'"
                        Set RsPegaDescricaoAlternativa = DBLoja.OpenRecordset(SQL)
                    wPegaDescricaoAlternativa = ""
                    If Not RsPegaDescricaoAlternativa.EOF Then
                        wPegaDescricaoAlternativa = IIf(IsNull(RsPegaDescricaoAlternativa("Desc")), 0, RsPegaDescricaoAlternativa("Desc"))
                    End If
                      
                      wStr16 = ""
                      wStr16 = Space(6) & Left$(RsdadosItens("ReferenciaAlternativa") & Space(8), 8) _
                             & Space(2) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
                             & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                             & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
                             & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                             & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "########0.00"), 12) & Space(2) _
                             & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "########0.00"), 15) & Space(2) _
                             & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)
                   
                             Printer.Print wStr16
                   
                   If RsdadosItens("DetalheImpressao") = "D" Then
                             wConta = wConta + 1
                             RsdadosItens.MoveNext
                   ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                             wConta = 0
                             RsdadosItens.MoveNext
                             Printer.NewPage
                             wpagina = wpagina + 1
                             Call Cabecalho
                   ElseIf RsdadosItens("DetalheImpressao") = "T" Then
                             wConta = wConta + 1
                             RsdadosItens.MoveNext
                             Call FinalizaNota
                   Else
                             wConta = wConta + 1
                             RsdadosItens.MoveNext
                   End If
             Loop
          End If
   Else
           MsgBox "Impossivel imprimir nota fiscal", vbCritical, "Error"
           Call Finaliza
   End If
    
   
    flg = 0
    wlin = 99
    Screen.MousePointer = 0

End Sub


Public Sub Discador()


    Wconectou = False
    
    On Error Resume Next
   

    Set RdoDados = Conexao.OpenResultset("Select LO_Loja from Loja where Lo_loja='315'", Options:=rdExecDirect)
    
    
    If Err.Number = 40071 Then
       rtn = Shell("rundll32.exe rnaui.dll,RnaDial " & "VicNet", 1)
       'HandlerWindow = FindWindow("#32770", "Conectar a")
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
        Exit Function
    
ConexaoErro:

    ConectaODBC = False
    Wconectou = False

End Function


Public Function CriaArquivoNF()
    wNotaTransferencia = False
    wpagina = 1
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
        
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then
      
      If Glb_NfDevolucao = True And RsDados("Serie") = "SM" Then
            Wsm = True
      End If
      
      Call CabecalhoArq
            
      
      
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,NfItens.IcmPdv,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa,nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & WNF & " and NfItens.Serie = '" & Wserie & "' order by nfitens.item"

      Set RsdadosItens = DBLoja.OpenRecordset(SQL)

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
                          & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                          & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "########0.00"), 12) & Space(1) _
                          & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "########0.00"), 15) & Space(1) _
                          & Right$(Space(2) & Format(RsdadosItens("IcmPdv"), "#0"), 2)
            
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
                         & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                         & Right$(Space(12) & Format(RsdadosItens("vlunit"), "########0.00"), 12) & Space(1) _
                         & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "########0.00"), 15) & Space(1) _
                         & Right$(Space(2) & Format(RsdadosItens("IcmPdv"), "#0"), 2)

                                  
            End If
                      
                      'On Error Resume Next
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
                         wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                         Print #NotaFiscal, wStr13
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, Chr(18) 'Finaliza Impressão
                         Close #NotaFiscal
                         wConta = 0
                         wpagina = wpagina + 1
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
'         FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If
End Function
              
Private Sub CabecalhoArq()
        
        Dim wCgcCliente As String
        
        NomeArquivo = "nf" & Trim(RsDados("NF")) & wpagina & ".txt"
        
        NotaFiscal = FreeFile
        
        Open Temporario & NomeArquivo For Output Access Write As #NotaFiscal
        
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
           Wentrada = Format(RsDados("Pgentra"), "########0.00")
           wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
        End If
        If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
            wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
        End If
        
        WCGC = Right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "##,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        
        wStr1 = Space(120) & wpagina & "/" & RsDados("PAGINANF") 'Inicio Impressão
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
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & "(011)" & Left$(Trim(Format(WFone, "####-####")), 9) & "/(011)" & Left$(Format(WFone, "####-####"), 9) & Space(11) & Left$(Format((WCep), "#####-###"), 9)
        Print #NotaFiscal, wStr3
        wStr4 = Space(2) & Left$(Format(wStr20) & Space(100), 100) & Left$(Trim(Format(WCGC, "###,###,###")), 19) '& Format(Mid((WCGC), 11, 5), "####-##")
        Print #NotaFiscal, wStr4
        'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
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
            wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "##,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
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
                Set RsPegaItensEspeciais = DBLoja.OpenRecordset(SQL)
                
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
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        Print #NotaFiscal, wStr9
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "########0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("TotalNotaAlternativa"), "########0.00"), 12)
        Print #NotaFiscal, wStr10
     Else
        wStr9 = Right$(Space(2) & Format(RsDados("BaseICMS"), "########0.00"), 12) & Space(1) & Right$(Space(12) & Format(RsDados("VLRICMS"), "########0.00"), 12) & Space(38) & Right$(Space(15) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
        Print #NotaFiscal, wStr9
        wStr10 = Right(Space(2) & Format(Space(12) & RsDados("FreteCobr"), "########0.00"), 12) & Space(53) & Right(Space(12) & Format(RsDados("VlrMercadoria"), "########0.00"), 12)
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
'     FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\FlagShip\exe\" & NomeArquivo
     wTotalNotaTransferencia = RsDados("VlrMercadoria")
     If wReemissao = False Then
        SQL = "Select * from CtCaixa order by CT_Data desc"
           Set rsPegaLoja = DBLoja.OpenRecordset(SQL)
        If Not rsPegaLoja.EOF Then
           If WNatureza = "TRANSFERENCIAS" Then
               SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(" & Val(glb_ECF) & ",'" & rsPegaLoja("ct_operador") & "','" & rsPegaLoja("ct_loja") & "', " _
                   & " #" & Format(rsPegaLoja("ct_data"), "mm/dd/yyyy") & "#, " & 20109 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   DBLoja.Execute (SQL)
           'ElseIf WNatureza = "DEVOLUCAO" Then
               'SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                   & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(1,'" & RsPegaLoja("ct_operador") & "','" & RsPegaLoja("ct_loja") & "', " _
                   & " #" & Format(RsPegaLoja("ct_data"), "mm/dd/yyyy") & "#, " & 20201 & "," & WNfTransferencia & ",'SN', " _
                   & "" & ConverteVirgula(Format(wTotalNotaTransferencia, "###,###0.00")) & ", " _
                   & "0,0,0,0,0,9,'A')"
                   'DBLoja.Execute (SQL)
           End If
        End If
    End If
End Sub
           

Function Cabecalho()
    Dim wCgcCliente As String

'    'Printer.PrintQuality = vbPRPQDraft
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 7#
    
    Wcondicao = "            "
    Wav = "          "
    If RsDados("CondPag") = 85 Then
        wCarimbo4 = Format(RsDados("DataPag"), "dd/mm/yyyy")
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
    
    If Trim(WNatureza) = "TRANSFERENCIA" Then
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
       Wentrada = Format(RsDados("Pgentra"), "########0.00")
       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
    End If
    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
        wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
    End If
    
    
    'wLinha2 = Space(2) & Left(rsDadosCliente("CLI_RazaoSocial") & Space(100), 100) _
            & Left(rsDadosCliente("CLI_Cnpj") & Space(18), 18) _
            & Left(Format(Date, "dd/mm/yyyy") & Space(10), 10)
    
    'Printer.FontSize = 8
    
    WCGC = Right(String(14, "0") & WCGC, 14)
    WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "##,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
     wStr1 = Space(110) & wpagina & "/" & RsDados("PAGINANF") 'Inicio Impressão
    Printer.Print wStr1
    'Print #Notafiscal, wStr1
    wStr1 = Space(2) & Left(Format(wStr17) & Space(40), 40) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(10), 10) & Space(5) & "X" & Space(22) & Left(Format(RsDados("nf"), "###,###"), 7)
    Printer.Print wStr1
    wStr2 = Space(2) & Left(Format(wStr18) & Space(40), 40) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
    Printer.Print wStr2
    If Wserie = "CT" Then
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & Space(29) & "(011)" & Left$(Trim(Format(WFone, "####-####")), 9) & "/(011)" & Left$(Format(WFone, "####-####"), 9) & Space(5) & Left$(Format((WCep), "#####-###"), 9)
    Else
        wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & "(011)" & Left$(Trim(Format(WFone, "####-####")), 9) & "/(011)" & Left$(Format(WFone, "####-####"), 9) & Space(5) & Left$(Format((WCep), "#####-###"), 9)
    End If
    Printer.Print wStr3
    If Wserie = "CT" Then
        wStr4 = ""
    Else
        wStr4 = Space(2) & Left(Format(wStr20) & Space(40), 40) & Space(50) & Left(Trim(Format(WCGC, "###,###,###")), 19) '& Format(Mid((WCGC), 11, 5), "####-##")
    End If
    Printer.Print wStr4
    Printer.Print ""
    'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
    If Wserie = "CT" Then
        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) '& Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
    Else
        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        ElseIf Trim(Wav) <> "" Then
            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        Else
            wStr5 = Space(36) & Format(Trim(WNatureza), ">") & Space(19) & Left$(RsDados("CFOAUX"), 10) & Space(21) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
        End If
    End If
    Printer.Print wStr5
    'Print #Notafiscal, ""
    Printer.Print ""
    Printer.Print ""
    If Mid(RsDados("CLIENTE"), 1, 5) <> "99999" Then
        wCgcCliente = Right(String(14, "0") & Trim(RsDados("CGCCLI")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "##,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
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
        If wStr6 <> "" Then
            wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Trim(wCgcCliente), 19) & Space(11) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr6 = Space(36) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(45), 45) & Left$(Trim(wCgcCliente), 19) & Space(10) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        End If
    End If
    
    Printer.Print wStr6
    If Wserie = "CT" Then
        wStr7 = Space(2) & Left(wStr7 & Space(34), 34) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Space(14) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
    Else
        wStr7 = Space(2) & Left(wStr7 & Space(34), 34) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(42), 42) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(25), 25) & Left$(RsDados("CEPCLI"), 11) & Space(5) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
    End If
    Printer.Print ""
    Printer.Print wStr7
    If Wserie = "CT" Then
        wStr8 = ""
    Else
        wStr8 = Space(36) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
    End If
    'Printer.Print ""
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


Function CabecalhoNovo()

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
       Wentrada = Format(RsDados("Pgentra"), "########0.00")
       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
    End If
    If (IIf(IsNull(RsDados("PedCli")), 0, RsDados("PedCli"))) <> 0 Then
        wStr7 = "Ped. Cliente    : " & Trim(RsDados("PedCli"))
    End If
    
    
    'wLinha2 = Space(2) & Left(rsDadosCliente("CLI_RazaoSocial") & Space(100), 100) _
            & Left(rsDadosCliente("CLI_Cnpj") & Space(18), 18) _
            & Left(Format(Date, "dd/mm/yyyy") & Space(10), 10)
    
    
    Printer.Print Space(20) & wpagina & "/" & RsDados("PAGINANF")  'Inicio Impressão
    'Printer.Print wStr1
    'Print #Notafiscal, wStr1
    
    
    wStr1 = Space(2) & Left(Format(wStr17) & Space(40), 40) & Left(Format(Trim(Wendereco), ">") & Space(34), 34) & Left(Format(Trim(wbairro), ">") & Space(15), 15) & Space(7) & "X" & Space(20) & Left(Format(RsDados("nf"), "###,###"), 7)
    Printer.Print wStr1
    wStr2 = Space(2) & Left(Format(wStr18) & Space(40), 40) & Left(Format(Trim(WMunicipio)) & Space(15), 15) & Space(24) & Left$(Trim(westado), 2)
    Printer.Print wStr2
    wStr3 = Space(2) & Left$(Format(wStr19) & Space(40), 40) & "(011)" & Left$(Trim(Format(WFone, "####-####")), 9) & "/(011)" & Left$(Format(WFone, "####-####"), 9) & Space(5) & Left$(Format((WCep), "#####-###"), 9)
    Printer.Print wStr3
    Printer.Print ""
    wStr4 = Space(2) & Left(Format(wStr20) & Space(20), 20) & Space(20) & Left(Trim(Format(WCGC, "###,###,###")), 19) '& Format(Mid((WCGC), 11, 5), "####-##")
    Printer.Print wStr4
    Printer.Print ""
    'wStr5 = Space(44) & Trim(WNatureza) & Space(22) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
    If Trim(WNatureza) = "TRANSFERENCIAS" Then
        wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(16) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
    ElseIf Trim(Wav) <> "" Then
        wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
    Else
        wStr5 = Space(34) & Format(Trim(WNatureza), ">") & Space(25) & Left$(RsDados("CFOAUX"), 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
    End If
    Printer.Print wStr5
    'Print #Notafiscal, ""
    Printer.Print ""
    If wStr6 <> "" Then
        wStr6 = Space(2) & wStr6 & Space(8) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(6) & Left$(Trim(RsDados("CGCCLI")), 19) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
    Else
        wStr6 = Space(34) & Left$(Format(Trim(RsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & Left$(Format(Trim(RsDados("NOMCLI")), ">") & Space(50), 50) & Space(11) & Left$(Trim(RsDados("CGCCLI")), 19) & Space(6) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
    End If
    Printer.Print wStr6
    Printer.Print ""
    wStr7 = Space(2) & Left(wStr7 & Space(32), 32) & Left$(Format(Trim(RsDados("ENDCLI")), ">") & Space(40), 40) & Space(7) & Left$(Format(Trim(RsDados("BAIRROCLI")), ">") & Space(15), 15) & Space(19) & Left$(RsDados("CEPCLI"), 11) & Space(3) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
    Printer.Print wStr7
    wStr8 = Space(34) & Left$(Format(Trim(RsDados("MUNICIPIOCLI")), ">") & Space(15), 15) & Space(19) & Left$(Format(Trim(RsDados("FONECLI"))) & Space(15), 15) & Space(8) & Left$(Trim(RsDados("UFCLIENTE")), 2) & Space(5) & Left$(Trim(Format(RsDados("INSCRICLI"), "###,###,###,###")), 15)
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


Function EmiteNotaTransferencia()
    
    Set DBFBanco = OpenDatabase(WbancoDbf, False, False, "DBase IV")
        
    Set NFCapaDBF = DBFBanco.OpenRecordset("SELECT * FROM nfcapa.dbf " & _
                          "WHERE sitnf='A' and TipoVenda in ('T','E') and Serie not in('CT')")
    
    Wserie = ""
    Processado = "P"
    wNotaDevolucao = False
    X = NFCapaDBF.RecordCount
    
    If Not NFCapaDBF.EOF Then
        
        Nota = 0

            Do While Not NFCapaDBF.EOF
                If Nota <> NFCapaDBF("notafis") Then
                    Nota = 0
                    Nota = NFCapaDBF("notafis")
                    
                    Set NFItemDBF = DBFBanco.OpenRecordset("SELECT * FROM nfitem.dbf " & _
                                    "WHERE notafis=" & NFCapaDBF("notafis") & " and " & _
                                    "sitnf='A'")
                    If Not NFItemDBF.EOF Then
                        GravaNFCapaDBFAccess
                        GravaNFItemDBFAccess
                        If wNotaDevolucao = False Then
                            Call EncerraVendaMigracao(Pedido, " ", 0)
                        Else
                            QuebraNotaDevolucao Nota
                        End If
                        Call AtualizaEstoque(Nota, Wserie, 0)
                        
                        If WTipoNota = "T" Or WTipoNota = "E" Then
                           SQL = ""
                           SQL = "select CT_SeqNota,CT_LOJA from Controle"
                           Set RsPegaNumNote = DBLoja.OpenRecordset(SQL)
                              wLoja = RsPegaNumNote("ct_loja")
                              If Not RsPegaNumNote.EOF Then
                                 Call ExtraiSequenciaNotaTransferencia
                                 SQL = ""
                                 SQL = "Update NfCapa set NF = " & WNfTransferencia & " " _
                                     & "Where NumeroPed = " & Pedido & " and NF= " & Pedido & ""
                                 DBLoja.Execute (SQL)
                        
                                 SQL = ""
                                 SQL = "Update NfItens set NF=" & WNfTransferencia & " " _
                                     & "Where Numeroped=" & Pedido & " and Nf=" & Pedido & ""
                                 DBLoja.Execute (SQL)
                                 Call CriaArquivoNFTransferencia
                              End If
                       End If
                        
                    Else
                        DBFBanco.Execute "Update nfcapa.dbf set " & _
                                         "sitnf='V' WHERE notafis=" & Nota & "", dbFailOnError
                        GoTo continua
                        NFItemDBF.Close
                    End If
                End If
continua:
                NFCapaDBF.MoveNext
            Loop
 
        NFCapaDBF.Close
        NFItemDBF.Close
 Else
    MsgBox "Nenhuma nota encontrada", vbInformation, "Aviso"
    NFCapaDBF.Close
 End If
    
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
    condpagto = IIf(IsNull(NFCapaDBF("condpagto")), 0, NFCapaDBF("condpagto"))
    av = IIf(IsNull(NFCapaDBF("AV")), 0, NFCapaDBF("AV"))
    Cliente = IIf(IsNull(NFCapaDBF("CLIENTE")), 0, NFCapaDBF("CLIENTE"))
    NATOPER = IIf(IsNull(NFCapaDBF("NATOPER")), 0, NFCapaDBF("NATOPER"))
    datapag = IIf(IsNull(NFCapaDBF("DATAPAG")), 0, NFCapaDBF("DATAPAG"))
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
    pessoa = IIf(IsNull(NFCapaDBF("pessoa")), 0, NFCapaDBF("pessoa"))
    ufcli = IIf(IsNull(NFCapaDBF("uf")), "SP", NFCapaDBF("uf"))
    cepcli = IIf(IsNull(NFCapaDBF("cep")), "", NFCapaDBF("cep"))
    bairrocli = IIf(IsNull(NFCapaDBF("bairro")), "", NFCapaDBF("bairro"))
    
    
    
    wValorTotalCodigoZero = 0
    If wValorTotalMercadoriaAlternativa > 0 Then
        wValorTotalCodigoZero = Val(TotNota - wValorTotalMercadoriaAlternativa)
    End If
    WTipoNota = IIf(IsNull(NFCapaDBF("TIPOVENDA")), "", NFCapaDBF("TIPOVENDA"))
    
    If NFCapaDBF("serie") = "RS" Then
        tiponota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "RC" Then
        tiponota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "RA" Then
        tiponota = "RA"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    ElseIf NFCapaDBF("serie") = "R2" Then
        tiponota = "RE"
        PgEntra = 0: PEDCLI = 0: PesoBr = 0: PesoLq = 0: ValFrete = 0: FreteCobr = 0
        BASEICM = 0: PORICM = 0: VLRICM = 0: Hora = 0: TOTIPI = 0: Desconto = 0
    Else
        
        tiponota = NFCapaDBF("TipoVenda")
        If tiponota = "E" Then
            wNotaDevolucao = True
        Else
            wNotaDevolucao = False
        End If
    End If
    
    If IsNull(NFCapaDBF("datapag")) Then
        datapag = "00:00:00"
    Else
        datapag = NFCapaDBF("datapag")
    End If
    If IsNull(NFCapaDBF("condpagto")) Then
        condpagto = 0
    Else
        condpagto = NFCapaDBF("condpagto")
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
    If Trim(NATOPER) = 132 Then
        wCfoAuxDev = 1202
    ElseIf Trim(NATOPER) = 232 Then
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
          "(" & Pedido & ", #" & Format(DataEmi, "MM/DD/YYYY") & "#, " & CODVEND & ", " & ConverteVirgula(VLRMERC) & " , " & _
          "" & ConverteVirgula(Desconto) & " ," & ConverteVirgula(SubTotal) & " ,'" & wLoja & "', '" & tiponota & "', " & _
          "" & condpagto & " , " & av & " , " & Cliente & " , " & NATOPER & ", " & wCfoAuxDev & " , " & _
          "#" & datapag & "# , " & ConverteVirgula(PgEntra) & " , '" & lojat & "' , " & TOTITENS & " , " & _
          "" & PEDCLI & " , 1," & ConverteVirgula(PesoBr) & " , " & ConverteVirgula(PesoLq) & " , " & ConverteVirgula(ValFrete) & " , " & _
          "" & ConverteVirgula(FreteCobr) & " ,  '" & OUTRALOJA & "' , " & OUTROVEND & " , " & notafis & " , " & _
          "" & ConverteVirgula(TotNota) & ", " & ConverteVirgula(BASEICM) & ", " & ConverteVirgula(PORICM) & " ," & ConverteVirgula(VLRICM) & " , " & _
          "'" & Wserie & "' ,#" & Format(Hora, "hh:mm") & "#," & ConverteVirgula(TOTIPI) & ", '" & nomecli & "','" & fonecli & "','" & cgccli & "', '" & endcli & "','" & ufcli & "', " & _
          "'" & muncli & "'," & pessoa & "," & Val(glb_ECF) & ", " & numerosf & ", " & _
          "" & ConverteVirgula(wValorTotalCodigoZero) & "," & ConverteVirgula(wTotalNotaAlternativa) & ", " & ConverteVirgula(wValorTotalMercadoriaAlternativa) & ",'A','" & cepcli & "','" & muncli & "','" & Val(glb_ECF) & "','" & wCarimbo1 & "','" & wCarimbo2 & "','" & wCarimbo3 & "') "
    
    DBLoja.Execute (SQL)
    
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
                    tiponota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "RC" Then
                    tiponota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "RA" Then
                    tiponota = "RA"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                ElseIf NFItemDBF("serie") = "R2" Then
                    tiponota = "RE"
                    Desconto = 0: plista = 0: Linha = 0: Secao = 0: csprod = 0: CODVEND = 0
                    aliqipi = 0
                Else
                    If tipomov = 12 Then
                        tiponota = "T"
                    ElseIf tipomov = 23 Then
                        tiponota = "E"
                    End If
                End If
                
                
                BeginTrans
                
                SQL = "Insert INTO nfitens " & _
                      "(numeroped,dataemi,referencia,qtde,vlunit,vlunit2,vltotitem," & _
                      "item,vlipi,desconto,plista,comissao,icms,bcomis,csprod,linha,secao," & _
                      "nf,serie,lojaorigem,cliente,vendedor,aliqipi,tiponota,tipomovimentacao, " & _
                      "ValorMercadoriaAlternativa,PrecoUnitAlternativa,ReferenciaAlternativa,SituacaoEnvio) " & _
                      "Values " & _
                      "(" & PedidoItem & ",#" & Format(DataEmi, "MM/DD/YYYY") & "#,'" & Referencia & "'," & Quant & "," & _
                      "" & ConverteVirgula(PrecoUni) & "," & ConverteVirgula(wVlUnit2) & "," & ConverteVirgula(valormerc) & "," & Item & "," & ConverteVirgula(vlripi) & "," & ConverteVirgula(Desconto) & "," & _
                      "" & ConverteVirgula(plista) & "," & Comissao & "," & icms & "," & bcomis & "," & csprod & "," & _
                      "" & Linha & "," & Secao & "," & notafis & ",'" & Wserie & "','" & wLoja & "'," & _
                      "" & Cliente & "," & CODVEND & "," & aliqipi & ",'" & tiponota & "'," & tipomov & ", " & _
                      "" & ConverteVirgula(wValorMercadoriaAlternativa) & "," & ConverteVirgula(wPrecoUnitarioAlternativa) & "," & ConverteVirgula(wReferenciaAlternativa) & ",'A') "
                
                DBLoja.Execute (SQL)
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

    Dim WZERO As Double


    wpagina = 1
    wNotaTransferencia = True
    If WTipoNota = "T" Then
       WNatureza = "TRANSFERENCIAS"
    Else
       WNatureza = "DEVOLUCAO"
    End If
    
    Temporario = "C:\NOTASVB\"
            
    Call DadosLoja
            
    SQL = ""
    SQL = "Select NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.Carimbo5,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV,NFCAPA.Carimbo3,NFCAPA.Carimbo2,NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF,NFCAPA.LOJAT, " _
        & "NFCAPA.CLIENTE,NFCAPA.FONECLI,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
        & "NFCAPA.UFCLIENTE,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
        & "NFITENS.VLTOTITEM,NFITENS.ICMS " _
        & "From NFCAPA INNER JOIN NFITENS " _
        & "on (NfCapa.nf=Nfitens.nf) " _
        & "Where NfCapa.nf= " & WNfTransferencia & " " _
        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "'"
        
    Set RsDados = DBLoja.OpenRecordset(SQL)
    
    If Not RsDados.EOF Then
      
      Call CabecalhoArq
            
      SQL = "Select produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.detalheImpressao " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & WNfTransferencia & " order by nfitens.item"

      Set RsdadosItens = DBLoja.OpenRecordset(SQL)

      If Not RsdadosItens.EOF Then
         wConta = 0
         Do While Not RsdadosItens.EOF

               
               If Wsm = True Then
                    wPegaDescricaoAlternativa = IIf(IsNull(RsDados("Referencia")), "0", RsDados("Referencia"))
                      wStr16 = ""
                      wStr16 = Left$(RsdadosItens("ReferenciaAlternativa") & Space(8), 8) _
                             & Space(2) & Left$(Format(Trim(wPegaDescricaoAlternativa), ">") & Space(38), 38) _
                             & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                             & Space(10), 10) & Space(2) & Left$(Trim(wCodIPI), 1) & Left$(Trim(wCodTri), 1) _
                             & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                             & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                             & Right$(Space(12) & Format(RsdadosItens("PrecoUnitAlternativa"), "########0.00"), 12) & Space(1) _
                             & Right$(Space(12) & Format((RsdadosItens("PrecoUnitAlternativa") * RsdadosItens("QTDE")), "########0.00"), 15) & Space(1) _
                             & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)
               
               Else
               
                      WZERO = 0
                      wStr16 = ""
                      wStr16 = Left$(RsdadosItens("pr_referencia") & Space(8), 8) _
                            & Space(2) & Left$(Format(Trim(RsdadosItens("pr_descricao")), ">") & Space(38), 38) _
                            & Space(25) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                            & Space(10), 10) & Space(2) & Left$(Trim(WZERO), 1) & Left$(Trim(WZERO), 1) _
                            & "  " & Space(2) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                            & Space(5) & Right$(Space(6) & Format(RsdadosItens("QTDE"), "#####0"), 6) & Space(2) _
                            & Right$(Space(12) & Format(RsdadosItens("vlunit"), "########0.00"), 12) & Space(1) _
                            & Right$(Space(12) & Format(RsdadosItens("VlTotItem"), "########0.00"), 15) & Space(1) _
                            & Right$(Space(2) & Format(RsdadosItens("pr_icmssaida"), "#0"), 2)


               End If
               
                      
                      
                      Print #NotaFiscal, wStr16
                      
                      If RsdadosItens("DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                         Do While wConta < 21
                            wConta = wConta + 1
                            Print #NotaFiscal, ""
                         Loop
                         RsdadosItens.MoveNext
                         wStr13 = Space(95) & "Lj " & RsDados("LojaOrigem") & Space(16) & Right$(Space(7) & Format(RsDados("Nf"), "###,###"), 7)
                         Print #NotaFiscal, wStr13
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, ""
                         Print #NotaFiscal, Chr(18) 'Finaliza Impressão
                         Close #NotaFiscal
                         wConta = 0
                         wpagina = wpagina + 1
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
         'FileCopy Temporario & NomeArquivo, "\\DEMEOLINUX\Notas" & NomeArquivo
         
    
    End If
End Function


Function QuebraNotaDevolucao(ByVal wNumeroNota As Double)
    wQuantdadeTotalItem = 0
    wUltimoItem = 0
    SQL = ""
    SQL = "Select NfCapa.QtdItem,NfItens.Item from NfCapa,NfItens " _
        & "where Nfcapa.Nf=" & wNumeroNota & " " _
        & "and NfItens.Nf=NfCapa.NF order by NfItens.Item "
        Set RsCapaNF = DBLoja.OpenRecordset(SQL)
    
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
                DBLoja.Execute (SQL)
        
            RsCapaNF.MoveNext
        Loop
    
    SQL = ""
    SQL = "Update NfCapa set PaginaNF=" & wUltimoItem & " "
        DBLoja.Execute (SQL)
    
    End If
End Function


Sub BuscaTransferencia()


Dim Wtexto As String
Dim WAtualizados As String
Dim Wmaximo As Integer
Dim i As Integer
Dim Conta As Integer
Dim matArquivos() As String
Dim warquivo As String
Dim qtde As Integer

    Wtexto = WcaminhoTextos
    WAtualizados = WcaminhoTextosAtu
    
    warquivo = Dir(Wtexto)
    
        Do While warquivo <> ""
            If Mid(warquivo, 1, 2) = "tr" Or Mid(warquivo, 1, 2) = "ct" Then
                If Mid(warquivo, 1, 2) = "ct" Then
                     Wserie = "CT"
                End If
                wNumPed = Mid(warquivo, 3, Len(warquivo) - 6)
                wNfCapa = False
                wNFitens = False
                arquivo = FreeFile
                SQL = ""
                SQL = "Select NumeroPed from NfCapa where NumeroPed=" & Mid(warquivo, 3, Len(warquivo) - 6) & " "
                    Set RsVerificaPedido = DBLoja.OpenRecordset(SQL)
                If RsVerificaPedido.EOF Then
                    Open Wtexto & warquivo For Input Access Read As #arquivo
        
                    Do While Not EOF(arquivo)
                        Line Input #arquivo, BUFFER
                        If Mid(BUFFER, 1, 3) = "000" Then
                            Call AtualizaCapaTransf
                        ElseIf Mid(BUFFER, 1, 3) <> "PRO" Then
                            Call AtualizaItensTransf
                        End If
                    Loop
                    If wNfCapa = False Or wNFitens = False Then
                        PegaItensPedTransf False, ""
                    End If
                    wReemissao = False
                    NotaTransferencia Mid(warquivo, 3, Len(warquivo) - 6)
                End If
                Close #arquivo
                FileCopy Wtexto & warquivo, WAtualizados & warquivo
                Kill Wtexto & warquivo
        
            End If
            warquivo = Dir()
        Loop
        
        

End Sub


Sub AtualizaCapaTransf()

    SQL = "select  CT_Loja from Controle"
        Set rsPegaLoja = DBLoja.OpenRecordset(SQL)
    If Not rsPegaLoja.EOF Then
        wLoja = rsPegaLoja("CT_Loja")
    End If
        
        
    wVendedorLojaVenda = 0
    wLojaVenda = ""
    WTotPedido = 0
    wSubTotal = 0
    wTotalNotaAlternativa = 0
    wValorTotalCodigoZero = 0
    WnumeroPed = Mid(BUFFER, 4, 8)
    WCliente = Mid(BUFFER, 12, 6)
    WNomeCliente = Mid(BUFFER, 39, 18)
    WVendedor = Mid(BUFFER, 207, 3)
    WCOMISSAO = "7"
    wTotalNotaAlternativa = 0
    wValorTotalCodigoZero = 0
    wCarimbo3 = ""
    wPedidoCliente = 0
    Wdata = Format(Date, "dd/mm/yyyy")
    If Trim(Mid(BUFFER, 249, 14)) <> "" Then
       WTotPedido = ConverteVirgula2(Mid(BUFFER, 249, 14))
    End If
    If Trim(Mid(BUFFER, 235, 14)) <> "" Then
       wSubTotal = ConverteVirgula2(Mid(BUFFER, 235, 14))
    End If
    
    Wlojat = ConverteVirgula2(Mid(BUFFER, 331, 5))
    If Trim(Mid(BUFFER, 336, 15)) <> "" Then
        wCarimbo3 = Mid(BUFFER, 342, 15)
    End If
    Wdescontop = Format(wSubTotal - WTotPedido, "0.00")
    'Wlojat = "999"
    
    
    
    If Mid(BUFFER, 267, 3) <> "" Then
       WCODOPER = Mid(BUFFER, 267, 3)
       WCFOAux = WCODOPER
       If WCFOAux = 522 Then
            WCFOAux = 5152
       End If
       WPGENTRA = ConverteVirgula2(Mid(BUFFER, 271, 14))
       'WCONDPAG = Mid(BUFFER, 286, 2)
       wCondPag = 0
       WDESCRIPAG = Mid(BUFFER, 289, 13)
    End If
    
    WQTDITEM = Mid(BUFFER, 224, 3)
    
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
    
    WNOMCLI = Mid(BUFFER, 18, 39)
    WENDCLI = Mid(BUFFER, 57, 39)
    WENDENTCLI = Mid(BUFFER, 57, 39)
    wbairro = Mid(BUFFER, 96, 16)
    WMUNCLI = Mid(BUFFER, 112, 21)
    WUF = Mid(BUFFER, 133, 2)
    WREGIAO = Mid(BUFFER, 96, 16)
    WCep = Mid(BUFFER, 135, 9)
    WIest = Mid(BUFFER, 157, 16)
    WCGCCLI = Mid(BUFFER, 174, 18)
    WDDD = 0
    WFone = Mid(BUFFER, 192, 9)
    If IsDate(Mid(BUFFER, 304, 14)) = True Then
        WdataPag = Mid(BUFFER, 305, 14)
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
    
    wReferencia = Mid(BUFFER, 4, 13)
    wQtde = Val(Mid(BUFFER, 49, 6))
    WTP = 1
    If Trim(Mid(BUFFER, 57, 14)) <> "" Then
       wVlUnit = ConverteVirgula2(Mid(BUFFER, 57, 14))
       wVlUnit2 = ConverteVirgula2(Mid(BUFFER, 85, 14))
       wVlTotItem = wVlUnit * wQtde
    End If
    
    WDESCRAT = ConverteVirgula2(Mid(BUFFER, 99, 14))
    wSituacao = "F"
    WSTATUS = "F"
    wItem = Val(Mid(BUFFER, 1, 3))
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
    Set ISQL = DBLoja.OpenRecordset(SQL)
    
    If Not ISQL.EOF Then
       wLinha = ISQL("Pr_LINHA")
       wSecao = ISQL("Pr_SECAO")
       WSUBTRIBUT = ISQL("PR_SubstituicaoTributaria")
       wPLISTA = ISQL("pr_precovenda1")
       WTRIBUTO = ISQL("pr_icmssaida")
       wIcmPdv = ISQL("pr_icmssaida")
       wCodBarra = ISQL("pr_codigobarra")
    Else
       MsgBox "Produto não encontrado", vbCritical, "Atenção"
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
    
    SQL = ""
    SQL = "select CT_SeqNota,CT_LOJA from Controle"
        Set RsPegaNumNote = DBLoja.OpenRecordset(SQL)
    wLoja = RsPegaNumNote("ct_loja")
    If Not RsPegaNumNote.EOF Then
        Call ExtraiSequenciaNotaTransferencia
        SQL = ""
        SQL = "Update NfCapa set NF = " & WNfTransferencia & " " _
            & "Where NumeroPed = " & NumeroPedido & ""
            DBLoja.Execute (SQL)
                        
        SQL = ""
        SQL = "Update NfItens set NF=" & WNfTransferencia & " " _
            & "Where Numeroped=" & NumeroPedido & " "
            DBLoja.Execute (SQL)
            
        Call EncerraVendaMigracao(NumeroPedido, " ", 0)
        If Wserie = "CT" Then
            WNF = WNfTransferencia
            Call AtualizaEstoque(WNfTransferencia, "CT", 0)
            Call EmiteNotafiscal
        Else
            Call AtualizaEstoque(WNfTransferencia, "SN", 0)
            Call CriaArquivoNFTransferencia
        End If
    End If
End Function


Public Sub ProcessaRotinasDiarias()
    
    Dim rsPegaData As Recordset
    'frmAguarde.Show
    
    On Error Resume Next
    SQL = ""
    SQL = "Select CT_Loja,CT_TipoArquivo,CT_OnLine,CT_ECF from Controle"
        Set RSTipoControle = DBLoja.OpenRecordset(SQL)
            
    If Not RSTipoControle.EOF Then
        If RSTipoControle("CT_TipoArquivo") = 1 Then
            frmRotinasDiaria.lblRotinas.Caption = " Iniciando o dia"
            BeginTrans
                SQL = ""
                SQL = "Delete * from EstqLojaDBF where Situacao='P'"
                DBLoja.Execute (SQL)
            CommitTrans
            BeginTrans
                SQL = ""
                SQL = "Delete * from ProduLjDBF where Situacao='P'"
                DBLoja.Execute (SQL)
            CommitTrans
            Call ZeraVendaDia
            Call LimpaHoraAtualizacao
            BeginTrans
                frmRotinasDiaria.lblProcessos.Caption = "Atualizando Controle"
                SQL = ""
                SQL = "Update Controle set CT_TipoArquivo = 0"
                DBLoja.Execute (SQL)
            CommitTrans
            frmRotinasDiaria.lblProcessos.Caption = "  BOAS VENDAS"
            
        ElseIf RSTipoControle("CT_TipoArquivo") = 2 Then
            frmRotinasDiaria.lblRotinas.Caption = " Fechando o dia"
            Screen.MousePointer = 11
            frmAbrirFecharCaixa.Refresh
            frmAguarde.Refresh
            frmAguarde.ZOrder
            frmAguarde.Show
            Call AtualizaLoja
'            Call ConfereMovimentoEstoque
            Call AtualizaEstoqueAnterior
            If Format(Wdata, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
               If Trim(RSTipoControle("CT_ECF")) = "S" Then
                    Call LeituraZ
               End If
            End If
            frmRotinasDiaria.lblProcessos.Caption = "Atualizando Controle"
            BeginTrans
                SQL = ""
                SQL = "Update Controle set CT_TipoArquivo = 0"
                DBLoja.Execute (SQL)
            CommitTrans
            '
            '----------------------Importando Movimento Dia-------------------------
            '
            If Trim(GLB_NumeroCaixa) = 1 Then
            
                FileCopy "C:\MovDia.mdb", Trim(Mid(WbancoAccess, 1, Len(WbancoAccess) - 8)) & "MovDia.mdb"
                
                Set dbMovDia = OpenDatabase(Trim(Mid(WbancoAccess, 1, Len(WbancoAccess) - 8)) & "MovDia.mdb")
                SQL = "Select Max(CT_Data) as Data from CtCaixa"
                    Set rsPegaData = DBLoja.OpenRecordset(SQL)
                If Not rsPegaData.EOF Then
                    CopiaNfCapa rsPegaData("data")
                    CopiaNfItens rsPegaData("data")
                End If
                
                dbMovDia.Close
                FileCopy Trim(Mid(WbancoAccess, 1, Len(WbancoAccess) - 8)) & "MovDia.mdb", "S:\Envio\MovDia.mdb"
                        
                CriaRelatorioCodigoZero rsPegaData("Data"), RSTipoControle("CT_Loja")
            
            End If
            AuditorEstoque rsPegaData("data")
            AcertaEstoqueDBF rsPegaData("data")
            frmRotinasDiaria.lblProcessos.Caption = "  BOA NOITE"
            Unload frmAguarde
            Screen.MousePointer = 0
        Else
            frmRotinasDiaria.lblProcessos.Caption = "ESTA ROTINA JA FOI PROCESSADA"
            Exit Sub
        End If
    Else
        MsgBox "ERRO NO PROCESSAMENTO DE ROTINAS DIARIAS", vbCritical, "ATENÇÃO"
        Exit Sub
    End If


End Sub


Sub LimpaHoraAtualizacao()
    ' frmRotinasDiaria.lblProcessos.Caption = "Hora Atualização"
    SQL = ""
    SQL = "Update HoraAtualizacao set HA_Situacao='A', " _
        & "HA_Status='A',HA_HoraInicio='00:00',HA_HoraFim='00:00' " _
        & "Where HA_Sequencia=1 "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update HoraAtualizacao set HA_Situacao='A', " _
        & "HA_Status='E',HA_HoraInicio='00:00',HA_HoraFim='00:00' " _
        & "Where HA_Sequencia > 1 "
        DBLoja.Execute (SQL)
    
End Sub

Sub AtualizaLoja()
    frmRotinasDiaria.lblProcessos.Caption = "Atualizando Loja"
    SQL = ""
    SQL = "Update MovimentoCaixa set MC_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MovimentoBancario set MB_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MovimentacaoEstoque set ME_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update CTcaixa set CT_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update DivergenciaEstoque set DE_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
    
    SQL = ""
    SQL = "Update EstoqueLoja set EL_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update MetadeVendas set MT_Loja = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update NfCapa set LojaOrigem = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
        
    SQL = ""
    SQL = "Update NfItens set LojaOrigem = '" & RSTipoControle("CT_Loja") & "' "
        DBLoja.Execute (SQL)
End Sub

Sub ZeraVendaDia()
    frmRotinasDiaria.lblProcessos.Caption = "Atualizando Vendas"
    SQL = ""
    SQL = "Update Vende set VE_MargemVenda = 0, VE_TotalVenda = 0"
    DBLoja.Execute (SQL)
End Sub


Sub ExtraiSequenciaNotaTransferencia()

    Dim WnovaSeqNota As Long
     
     SQL = ""
     SQL = "Select * from controle"
     Set RsDados = DBLoja.OpenRecordset(SQL)
     
     If Not RsDados.EOF Then
        If Wserie = "CT" Then
            WnumeroNotaDbf = 0
            WnovaSeqNota = 0
            
            WnumeroNotaDbf = RsDados("CT_SeqCT") + 1
            WnovaSeqNota = WnumeroNotaDbf
            WNfTransferencia = WnumeroNotaDbf
        
            SQL = "update controle set CT_SeqCT= " & WnovaSeqNota & ""
            DBLoja.Execute (SQL)
        Else
            WnumeroNotaDbf = 0
            WnovaSeqNota = 0
            
            WnumeroNotaDbf = RsDados("CT_SeqNota") + 1
            WnovaSeqNota = WnumeroNotaDbf
            WNfTransferencia = WnumeroNotaDbf
        
            SQL = "update controle set CT_SeqNota= " & WnovaSeqNota & ""
            DBLoja.Execute (SQL)
        End If
             
     End If
End Sub



Sub ProcessaListaPreco()

    SQL = "Select * from ListaPrecoCapa " _
        & "where LC_DataVigencia=#" & Format(Date, "dd/mm/yyyy")

End Sub


Function CopiaNfCapa(ByVal Data As String)

    Dim rsCopiaNfCapa As Recordset
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "Where DataEmi=#" & Format(Data, "mm/dd/yyyy") & "# " _
        & "and TipoNota not in ('PA','R','R2') and NF > 0 " _
        & "and Serie not in('R2','RC','S1','S2') order by Nf"
        Set rsCopiaNfCapa = DBLoja.OpenRecordset(SQL)
    If Not rsCopiaNfCapa.EOF Then
        Do While Not rsCopiaNfCapa.EOF
            SQL = ""
            SQL = "Insert into NfCapa (NUMEROPED, DATAEMI, VENDEDOR, VLRMERCADORIA, DESCONTO, " _
                & "SUBTOTAL, LOJAORIGEM, TIPONOTA, CONDPAG, AV, CLIENTE, CODOPER, DATAPAG, PGENTRA, LOJAT, QTDITEM, PEDCLI, TM, PESOBR, PESOLQ, VALFRETE, FRETECOBR, OUTRALOJA, OUTROVEND, NF, " _
                & "TOTALNOTA, NATOPERACAO, DATAPED, BASEICMS, ALIQICMS, VLRICMS, SERIE, HORA, TOTALIPI, ECF, NUMEROSF, NOMCLI, FONECLI, CGCCLI, INSCRICLI, ENDCLI, UFCLIENTE, MUNICIPIOCLI, BAIRROCLI, " _
                & "CEPCLI, PESSOACLI, REGIAOCLI, CFOAUX, AnexoAUx, PAGINANF, ECFNF, Carimbo1, Carimbo2, Carimbo3, Carimbo4, CustoMedioLiquido, VendaLiquida, MargemContribuicao, ValorTotalCodigoZero, " _
                & "TotalNotaAlternativa, ValorMercadoriaAlternativa, SituacaoEnvio, VendedorLojaVenda, LojaVenda) " _
                & "Values (" & rsCopiaNfCapa("NUMEROPED") & ", #" & Format(rsCopiaNfCapa("DATAEMI"), "dd/mm/yyyy") & "#, " & rsCopiaNfCapa("VENDEDOR") & ", " & ConverteVirgula(rsCopiaNfCapa("VLRMERCADORIA")) & ", " & ConverteVirgula(rsCopiaNfCapa("DESCONTO")) & ", " _
                & "" & ConverteVirgula(rsCopiaNfCapa("SUBTOTAL")) & ", '" & rsCopiaNfCapa("LOJAORIGEM") & "', '" & rsCopiaNfCapa("TIPONOTA") & "', '" & rsCopiaNfCapa("CONDPAG") & "', " & rsCopiaNfCapa("AV") & ", " & rsCopiaNfCapa("CLIENTE") & ", " & rsCopiaNfCapa("CODOPER") & ", #" & Format(rsCopiaNfCapa("DATAPAG"), "dd/mm/yyyy") & "#, " & ConverteVirgula(rsCopiaNfCapa("PGENTRA")) & ", '" & rsCopiaNfCapa("LOJAT") & "', " _
                & "" & rsCopiaNfCapa("QTDITEM") & ", " & rsCopiaNfCapa("PEDCLI") & ", " & rsCopiaNfCapa("TM") & ", " & ConverteVirgula(rsCopiaNfCapa("PESOBR")) & ", " & ConverteVirgula(rsCopiaNfCapa("PESOLQ")) & ", " & ConverteVirgula(rsCopiaNfCapa("VALFRETE")) & ", " & ConverteVirgula(rsCopiaNfCapa("FRETECOBR")) & ", '" & rsCopiaNfCapa("OUTRALOJA") & "', " & rsCopiaNfCapa("OUTROVEND") & ", " & rsCopiaNfCapa("NF") & ", " & ConverteVirgula(rsCopiaNfCapa("TOTALNOTA")) & ", " & rsCopiaNfCapa("NATOPERACAO") & ", #" & Format(IIf(IsNull(rsCopiaNfCapa("DATAPED")), rsCopiaNfCapa("DataEmi"), rsCopiaNfCapa("DATAPED")), "dd/mm/yyyy") & "#, " _
                & "" & ConverteVirgula(rsCopiaNfCapa("BASEICMS")) & ", " & ConverteVirgula(rsCopiaNfCapa("ALIQICMS")) & ", " & ConverteVirgula(rsCopiaNfCapa("VLRICMS")) & ", '" & rsCopiaNfCapa("SERIE") & "', #" & IIf(IsNull(rsCopiaNfCapa("HORA")), "00:00", rsCopiaNfCapa("HORA")) & "#, " & ConverteVirgula(rsCopiaNfCapa("TOTALIPI")) & ", " & rsCopiaNfCapa("ECF") & ", " & rsCopiaNfCapa("NUMEROSF") & ", '" & rsCopiaNfCapa("NOMCLI") & "', '" & rsCopiaNfCapa("FONECLI") & "', '" & rsCopiaNfCapa("CGCCLI") & "', '" & rsCopiaNfCapa("INSCRICLI") & "', '" & rsCopiaNfCapa("ENDCLI") & "', '" & rsCopiaNfCapa("UFCLIENTE") & "', " _
                & "'" & rsCopiaNfCapa("MUNICIPIOCLI") & "', '" & rsCopiaNfCapa("BAIRROCLI") & "', '" & rsCopiaNfCapa("CEPCLI") & "', " & rsCopiaNfCapa("PESSOACLI") & ", " & rsCopiaNfCapa("REGIAOCLI") & ", '" & rsCopiaNfCapa("CFOAUX") & "', '" & rsCopiaNfCapa("AnexoAUx") & "', " & rsCopiaNfCapa("PAGINANF") & ", " & rsCopiaNfCapa("ECFNF") & ", '" & rsCopiaNfCapa("Carimbo1") & "', ' " & rsCopiaNfCapa("Carimbo2") & "', '" & rsCopiaNfCapa("Carimbo3") & "', '" & rsCopiaNfCapa("Carimbo4") & "', " & ConverteVirgula(rsCopiaNfCapa("CustoMedioLiquido")) & ", " _
                & "" & ConverteVirgula(rsCopiaNfCapa("VendaLiquida")) & ", " & ConverteVirgula(rsCopiaNfCapa("MargemContribuicao")) & ", " & ConverteVirgula(rsCopiaNfCapa("ValorTotalCodigoZero")) & ", " & ConverteVirgula(rsCopiaNfCapa("TotalNotaAlternativa")) & ", " & ConverteVirgula(rsCopiaNfCapa("ValorMercadoriaAlternativa")) & ", '" & rsCopiaNfCapa("SituacaoEnvio") & "', " & rsCopiaNfCapa("VendedorLojaVenda") & ",'" & rsCopiaNfCapa("LojaVenda") & "')"
                dbMovDia.Execute (SQL)
            rsCopiaNfCapa.MoveNext
        Loop
    End If


End Function

Function CopiaNfItens(ByVal Data As String)
    
    Dim RsCopiaNfItens As Recordset
    Dim wDescricaoAlternativa  As String

    SQL = ""
    SQL = "Select * from NfItens " _
        & "Where DataEmi=#" & Format(Data, "mm/dd/yyyy") & "# " _
        & "and TipoNota not in ('PA','R','R2') and NF > 0 " _
        & "and Serie not in('R2','RC','S1','S2') order by Nf"
        Set RsCopiaNfItens = DBLoja.OpenRecordset(SQL)
    If Not RsCopiaNfItens.EOF Then
        Do While Not RsCopiaNfItens.EOF
            If RsCopiaNfItens("DescricaoAlternativa") = "" Then
                wDescricaoAlternativa = "0"
            Else
                wDescricaoAlternativa = IIf(IsNull(RsCopiaNfItens("DescricaoAlternativa")), 0, RsCopiaNfItens("DescricaoAlternativa"))
            End If
            SQL = "Insert into NfItens (NUMEROPED, DATAEMI, REFERENCIA, QTDE, VLUNIT, VLUNIT2, VLTOTITEM, DESCRAT, ICMS, ITEM, VLIPI, DESCONTO, PLISTA, COMISSAO, VALORICMS, BCOMIS, CSPROD, LINHA, SECAO, VBUNIT, ICMPDV, CODBARRA, NF, SERIE, LOJAORIGEM, CLIENTE, VENDEDOR, ALIQIPI, TIPONOTA, REDUCAOICMS, BASEICMS, TIPOMOVIMENTACAO, DETALHEIMPRESSAO, SERIEPROD1, " _
                & "SERIEPROD2, CustoMedioLiquido, VendaLiquida, MargemContribuicao, EncargosVendaLiquida, EncargosCustoMedioLiquido, PrecoUnitAlternativa, ValorMercadoriaAlternativa, ReferenciaAlternativa, SituacaoEnvio, DescricaoAlternativa)" _
                & "Values (" & RsCopiaNfItens("NUMEROPED") & ", #" & Format(RsCopiaNfItens("DATAEMI"), "dd/mm/yyyy") & "#, '" & RsCopiaNfItens("REFERENCIA") & "', " & RsCopiaNfItens("QTDE") & ", " & ConverteVirgula(RsCopiaNfItens("VLUNIT")) & ", " & ConverteVirgula(RsCopiaNfItens("VLUNIT2")) & ", " & ConverteVirgula(RsCopiaNfItens("VLTOTITEM")) & ", " _
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

    Dim RsCriaRelCodZero As Recordset
    Dim wValorTotalNota As Double
    Dim wNumeroLinha As Integer
    Dim wTotal As Double
    Dim wTotalGeral As Double
    Dim wUltimaSerie As String
    
    
    '
    '-----------------------Set a Impressora--------------------------------
    '
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "COTACAO/RESUMO" Then
            ' Seta impressora do sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    '***********************************************************************
    
    
    SQL = ""
    SQL = "Select NF,TotalNota,LojaOrigem,Serie,TotalNotaAlternativa from NfCapa " _
        & "Where LojaOrigem='" & Loja & "' " _
        & "and DataEmi=#" & Format(Data, "mm/dd/yyyy") & "# " _
        & "and Serie in ('00','SM') " _
        & "and TipoNota='V' " _
        & "order by Serie,NF"
    Set RsCriaRelCodZero = DBLoja.OpenRecordset(SQL)
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
                Printer.Print Space(18) & "TOTAL       00    " & Right(Space(14) & Format(wTotal, "0.00"), 14)
                wTotal = 0
                Printer.FontBold = False
                Printer.Print
            End If
            If Trim(RsCriaRelCodZero("Serie")) = "SM" Then
                wValorTotalNota = RsCriaRelCodZero("TotalNota") - RsCriaRelCodZero("TotalNotaAlternativa")
                wTotal = wTotal + wValorTotalNota
            Else
                wValorTotalNota = RsCriaRelCodZero("TotalNota")
                wTotal = wTotal + wValorTotalNota
            End If
            wNumeroLinha = wNumeroLinha + 1
            If wNumeroLinha <= 50 Then
                Printer.Print Space(18) & Left(RsCriaRelCodZero("NF") & Space(12), 12) _
                    & Left(RsCriaRelCodZero("Serie") & Space(10), 10) _
                    & Right(Space(10) & Format(wValorTotalNota, "0.00"), 10)
            Else
                Printer.NewPage
                CabecalhoRelCodZero Data, Loja
                Printer.Print Space(18) & Left(RsCriaRelCodZero("NF") & Space(12), 12) _
                    & Left(RsCriaRelCodZero("Serie") & Space(10), 10) _
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
            Printer.Print Space(18) & "TOTAL       SM    " & Right(Space(14) & Format(wTotal, "0.00"), 14)
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
    Printer.Print Tab(27); "R E L A T O R I O   D E  C O D I G O   Z E R O / 00         " & Format(Data, "dd/mm/yyyy") & "   LOJA  " & Loja; Tab(118); Format(Date, "DD/mm/YYYY")
    Printer.Print ""
    Printer.Print ""
    
    
    Printer.CurrentY = Printer.CurrentY - 1
    
    '(COLUNA,LINHA)
'    Printer.Line (8, 10)-(8, 37)
'    Printer.Line (199, 10)-(199, 37)
    Printer.Line (8, 270)-(199, 270)
    Printer.Line (8, 26)-(199, 26)
    
    Printer.Print ""
    Printer.Print Tab(27); "N                            S                                    V    "
    
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.FontName = "COURIER NEW"
    Printer.Print
End Function

Function PegaNumeroCFControle() As Double

    Dim rsPegaNumeroCF As Recordset
    
    SQL = ""
    SQL = "Select (CT_UltimoCupom + 1) as NumeroCupom from ControleECF " _
        & "where CT_Ecf=" & Val(glb_ECF) & ""
        Set rsPegaNumeroCF = DBLoja.OpenRecordset(SQL)
    If Not rsPegaNumeroCF.EOF Then
        SQL = ""
        SQL = "Update ControleECF set CT_UltimoCupom=" & rsPegaNumeroCF("NumeroCupom") & " " _
            & "where CT_Ecf=" & Val(glb_ECF) & ""
            DBLoja.Execute (SQL)
    
        PegaNumeroCFControle = rsPegaNumeroCF("NumeroCupom")
    End If

End Function

Sub AtualizaNumeroCupom()

    SQL = ""
    SQL = "Update controleEcf set ct_ultimocupom= CT_UltimoCupom + 1 " _
        & "where CT_Ecf=" & Val(glb_ECF) & ""
           DBLoja.Execute (SQL)

End Sub


Function VerificaControleEcf(ByVal NumeroECF As Integer, ByVal Loja As String)
    
    Dim rsVerificaControleEcf As Recordset
    Dim rsVerificaCT As Recordset
    Dim rsOperador As Recordset

    SQL = ""
    SQL = "Select * from ControleEcf " _
        & "where CT_Ecf=" & NumeroECF
        Set rsVerificaControleEcf = DBLoja.OpenRecordset(SQL)
    If rsVerificaControleEcf.EOF Then
        SQL = ""
        SQL = "Insert into ControleEcf (CT_Ecf,CT_QtdeEcf,CT_UltimoCupom,CT_SituacaoCupomFiscal,CT_SituacaoCaixa,CT_PegaPedido) " _
            & "Values(" & NumeroECF & ",0,0,'F','F','N') "
            DBLoja.Execute (SQL)
        SQL = ""
        SQL = "Select * from CTCaixa " _
            & "where CT_NumeroEcf=" & glb_ECF & ""
            Set rsVerificaCT = DBLoja.OpenRecordset(SQL)
        If rsVerificaCT.EOF Then
            SQL = ""
            SQL = "Select CT_Operador from CTCaixa where CT_Operador > 0 "
                Set rsOperador = DBLoja.OpenRecordset(SQL)
            If Not rsOperador.EOF Then
                SQL = "insert into CtCaixa (CT_NumeroECF,CT_Loja,CT_Data,CT_HoraInicial,CT_HoraFinal,CT_Operacoes,CT_Controle,CT_Situacao,CT_Operador) " _
                    & "Values(" & glb_ECF & ",'" & Loja & "',#" & Format(Date, "dd/mm/yyyy") & "#, " _
                    & "#" & Format(Time, "hh:mm") & "#, #" & Format(Time, "hh:mm") & "#,0,0,'P'," & rsOperador("CT_Operador") & ")"
                    DBLoja.Execute (SQL)
            End If
        End If
    Else
        If rsVerificaControleEcf("CT_SituacaoCupomFiscal") = "A" Then
            Retorno = Bematech_FI_CancelaCupom()
            MsgBox "Exite Cupom Aberto", vbInformation, "Aviso"
            Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
            If Retorno = 1 Then
                MsgBox "Atenção cupom sendo cancelado", vbInformation, "Problemas com Cupom"
                Call AtualizaNumeroCupom
                SQL = ""
                SQL = "Update ControleEcf set CT_SituacaoCupomFiscal='F' "
                    DBLoja.Execute (SQL)
            End If
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
    Dim rsVerCaixa As Recordset
    Dim rsPegaPedido As Recordset
    
    SQL = ""
    SQL = "Update ControleEcf set CT_SituacaoCaixa='" & Situacao & "' " _
        & "where CT_ECF=" & Val(glb_ECF) & ""
        DBLoja.Execute (SQL)
    
    If Situacao = "A" Then
        SQL = ""
        SQL = "Select CT_PegaPedido from ControleECF " _
            & "where CT_SituacaoCaixa='A' and CT_PegaPedido='S'"
            Set rsVerCaixa = DBLoja.OpenRecordset(SQL)
        If rsVerCaixa.EOF Then
            SQL = ""
            SQL = "Update ControleEcf set CT_PegaPedido='S' " _
                & "where CT_ECF=" & glb_ECF
            DBLoja.Execute (SQL)
        End If
    ElseIf Situacao = "F" Then 'Fecha O Caixa que Pega Pedido
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido='N' " _
            & "where CT_ECF=" & Val(glb_ECF) & ""
            DBLoja.Execute (SQL)
        
        SQL = ""
        SQL = "Select CT_PegaPedido from ControleECF " _
            & "where CT_SituacaoCaixa='A' and CT_PegaPedido='S'"
            Set rsVerCaixa = DBLoja.OpenRecordset(SQL)
        If rsVerCaixa.EOF Then
            SQL = "Select CT_PegaPedido,CT_ECF from ControleECF " _
                & "where CT_SituacaoCaixa='A' and CT_ECf<>" & Val(glb_ECF) & ""
                Set rsPegaPedido = DBLoja.OpenRecordset(SQL)
            If Not rsPegaPedido.EOF Then
                SQL = ""
                SQL = "Update ControleEcf set CT_PegaPedido='S' " _
                    & "where CT_ECF=" & rsPegaPedido("CT_ECF")
                DBLoja.Execute (SQL)
            End If
        End If
    ElseIf Situacao = "S" Then  'Muda O Caixa que pega os pedidos
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido = 'N' "
            DBLoja.Execute (SQL)
        
        SQL = ""
        SQL = "Update ControleEcf set CT_PegaPedido='S',CT_SituacaoCaixa='A' " _
            & "where CT_ecf=" & Val(glb_ECF) & ""
            DBLoja.Execute (SQL)
    End If
End Function


Function CriaNotaCredito(ByVal NF As Double, ByVal Serie As String, ByVal NfDev As Double, ByVal SerieDev As String, ByVal DataDev As String, ByVal ValorNotaCredito As Double, ByVal NotaCredito As Double)
    Dim rsDadosNfCapa As Recordset
    Dim rsVerLoja As Recordset
    Dim Linha1 As String
    Dim wTotalNota As Double
    Dim wValorExtenso As String
    
    'Printer.Line (0, 10)-(199, 10)
    'Printer.Line (0, 10)-(0, 100)
    'Printer.Line (199, 10)-(199, 100)
    
    'Printer.EndDoc
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "COTACAO/RESUMO" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "where Nf=" & NF & " " _
        & "and Serie='" & Serie & "'"
    Set rsDadosNfCapa = DBLoja.OpenRecordset(SQL)
    If Not rsDadosNfCapa.EOF Then
        SQL = ""
        SQL = "Select CT_Loja,CT_Razao,CT_NCredito,Lojas.* from Controle,Lojas " _
            & "where LO_Loja=CT_Loja"
        Set rsVerLoja = DBLoja.OpenRecordset(SQL)
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
                Printer.Print Space(2) & rsVerLoja("CT_Razao")
                Printer.Print Space(2) & Left(rsVerLoja("LO_Endereco") & Space(30), 30) _
                    & "    -    " & rsVerLoja("LO_Cep") & "   -   " & rsVerLoja("LO_Municipio") _
                    & Right(Space(103) & "NOTA DE CREDITO", 103)
                Printer.Print Space(2) & "FONE : " & "(" & Right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" _
                        & Left(rsVerLoja("LO_Telefone") & Space(10), 10) & " -  " _
                        & "FAX : " & "(" & Right(String(3, "0") & rsVerLoja("LO_DDD"), 3) & ")" & Left(rsVerLoja("LO_Telefone") & Space(10), 10)
                Printer.Print Space(2) & "C.G.C : " & Left(rsVerLoja("LO_CGC") & Space(25), 25) & "INSCR.EST. : " & rsVerLoja("LO_InscricaoEstadual")
                Printer.Print Space(190) & "NUM.  " & Right(String(9, "0") & NotaCredito, 9) & Right(Space(10) & i & "a.VIA", 10)
                Printer.Print Space(2) & "A"
                Printer.Print Space(2) & rsDadosNfCapa("NomCli")
                Printer.Print Space(2) & Left(rsDadosNfCapa("EndCli") & Space(190), 190) & Left("DATA : " & Date & Space(16), 16)
                Printer.Print Space(2) & rsDadosNfCapa("MunicipioCli") & "  -   " & rsDadosNfCapa("UfCliente")
                Printer.Print Space(2) & "FONE : " & rsDadosNfCapa("FoneCli")
                Printer.Print Space(2) & "EFETUAMOS NESTA DATA EM SUA CONTA CORRENTE O SEGUINTE LANÇAMENTO:"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print Space(60) & "HISTORICO" & Space(60) & "| DEBITO" & Space(30) & "| CREDITO"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(2) & "PELO RECEBIMENTO DA MERCADORIA EM DEVOLUÇÃO"
                Printer.Print Space(2) & "CONFORME NF " & NfDev & " SERIE " & SerieDev & " DE " & DataDev
                Printer.Print Space(2) & "NO VALOR DE R$          " & Format(wTotalNota, "###,###,###0.00")
                Printer.Print
                Printer.Print Space(2) & "REFERENTE NF " & rsDadosNfCapa("NF") & " - " & rsDadosNfCapa("SERIE") & " DE " & rsDadosNfCapa("DataEmi")
                Printer.Print Space(2) & "DA LOJA " & rsDadosNfCapa("LojaOrigem") & Space(170) & Format(wTotalNota, "###,###,###0.00")
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(170) & "ATENCIOSAMENTE"
                Printer.Print
                Printer.Print Space(150) & "_______________________________________"
                Printer.Print Space(150) & rsVerLoja("CT_Razao")
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


Function CriaNotaCreditoBola(ByVal NF As Double, ByVal Serie As String, ByVal NfDev As Double, ByVal SerieDev As String, ByVal DataDev As String, ByVal ValorNotaCredito As Double, ByVal NotaCredito As Double)
    Dim rsDadosNfCapa As Recordset
    Dim rsVerLoja As Recordset
    Dim Linha1 As String
    Dim wTotalNota As Double
    Dim wValorExtenso As String
    'Printer.Line (0, 10)-(199, 10)
    'Printer.Line (0, 10)-(0, 100)
    'Printer.Line (199, 10)-(199, 100)
    
    'Printer.EndDoc
    
    For Each NomeImpressora In Printers
        If Trim(NomeImpressora.DeviceName) = "COTACAO/RESUMO" Then
            ' Seta impressora no sistema
            Set Printer = NomeImpressora
            Exit For
        End If
    Next
    
    SQL = ""
    SQL = "Select * from NfCapa " _
        & "where Nf=" & NF & " " _
        & "and Serie='" & Serie & "'"
    Set rsDadosNfCapa = DBLoja.OpenRecordset(SQL)
    If Not rsDadosNfCapa.EOF Then
        SQL = ""
        SQL = "Select CT_Loja,CT_Razao,CT_NCredito,Lojas.* from Controle,Lojas " _
            & "where LO_Loja=CT_Loja"
        Set rsVerLoja = DBLoja.OpenRecordset(SQL)
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
                Printer.Print Space(2) & "LOJA "; rsVerLoja("CT_Loja")
                Printer.Print Space(190) & "NUM.  " & Right(String(9, "0") & NotaCredito, 9) & Right(Space(10) & i & "a.VIA", 10)
                Printer.Print Space(2) & "A"
                Printer.Print Space(2) & rsDadosNfCapa("NomCli")
                Printer.Print Space(2) & Left(rsDadosNfCapa("EndCli") & Space(190), 190) & Left("DATA : " & Date & Space(16), 16)
                Printer.Print Space(2) & rsDadosNfCapa("MunicipioCli") & "  -   " & rsDadosNfCapa("UfCliente")
                Printer.Print Space(2) & "FONE : " & rsDadosNfCapa("FoneCli")
                Printer.Print Space(2) & "EFETUAMOS NESTA DATA EM SUA CONTA CORRENTE O SEGUINTE LANÇAMENTO:"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print Space(60) & "HISTORICO" & Space(60) & "| DEBITO" & Space(30) & "| CREDITO"
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(2) & "PELO RECEBIMENTO DA MERCADORIA EM DEVOLUÇÃO"
                Printer.Print Space(2) & "CONFORME NF " & NfDev & " SERIE " & SerieDev & " DE " & DataDev
                Printer.Print Space(2) & "NO VALOR DE R$          " & Format(wTotalNota, "###,###,###0.00")
                Printer.Print
                Printer.Print Space(2) & "REFERENTE NF " & rsDadosNfCapa("NF") & " - " & rsDadosNfCapa("SERIE") & " DE " & rsDadosNfCapa("DataEmi")
                Printer.Print Space(2) & "DA LOJA " & rsDadosNfCapa("LojaOrigem") & Space(170) & Format(wTotalNota, "###,###,###0.00")
                Printer.Print Space(2) & "___________________________________________________________________________________________________________________"
                Printer.Print
                Printer.Print Space(170) & "ATENCIOSAMENTE"
                Printer.Print
                Printer.Print Space(150) & "_______________________________________"
                Printer.Print
                Printer.Print
            Next i
            Printer.EndDoc
        End If
    End If


End Function


Function ExtraiNumeroNotaCredito() As Double
    Dim rsNumeroNotaCredito As Recordset
    
        SQL = ""
        SQL = "Select (CT_NCredito+1) as NotaCredito from Controle"
            Set rsNumeroNotaCredito = DBLoja.OpenRecordset(SQL)
        If Not rsNumeroNotaCredito.EOF Then
            ExtraiNumeroNotaCredito = rsNumeroNotaCredito("NotaCredito")
            SQL = ""
            SQL = "Update Controle set CT_NCredito=" & rsNumeroNotaCredito("NotaCredito")
            DBLoja.Execute (SQL)
        End If
    
End Function

Function AchaLojaControle() As String
    
    Dim ControleLoja As Recordset
    
    Set ControleLoja = DBLoja.OpenRecordset("Select CT_Loja from Controle")
       
    AchaLojaControle = ControleLoja("CT_Loja")
       
    ControleLoja.Close
   
End Function

Function LiberaSenha(ByVal Usuario As String, ByVal Senha As String) As Boolean
    Dim rsLiberaSenha As Recordset
    
    SQL = "Select Us_Usuario from Usuario " _
        & "where US_Usuario='" & Usuario & "' " _
        & "and US_Senha='" & Senha & "' " _
        & "and Us_TipoUsuario in (2,4)"
    Set rsLiberaSenha = DBLoja.OpenRecordset(SQL)
    If Not rsLiberaSenha.EOF Then
        LiberaSenha = True
    Else
        LiberaSenha = False
    End If
    
End Function
