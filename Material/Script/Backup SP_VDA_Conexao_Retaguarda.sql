USE [DMAC]
GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Conexao_Retaguarda]    Script Date: 28/04/2014 15:57:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



/*
Procedure que gera toda movimentação Gerencial/Fiscal/Estoque 
a partir da tabela NFCapa

Obs.
    Esta Procedure Baseia-se na situacão da tabela NFCapa
    a partir desta situação encadeia-se a atualização da retaguarda


-- exec SP_VDA_Conexao_Retaguarda '116'
select * from historicoloja where hl_data='2012/06/20' 
Update Loja set LO_Conexao='S' Where LO_Loja = '116'

 SP_VDA_Atualiza_HistoricoLoja_Venda '315','2012/06/20'
*/

ALTER     Procedure [dbo].[SP_VDA_Conexao_Retaguarda]
                  @LojaPar char(05)
As

Begin
--NFCapa
	Declare @W_Critica		            Char(1),
                @W_TotalItens	     	Float,
                @W_ValorCapa            Float,
                @W_ContaCliente         Float,
                @W_ItensItem            Float,
                @W_LojaDestino          char(05),
                @W_ObsCritica           Char(5000),
		        @W_ValorItens		    Float,
		        @W_ItensCapa		    Int,
		        @W_DataDuplicata	    Char(10),
                @SQL                    Char(4000),
                @DataMovimento          char(10),
                @Loja                   char(05), 
                @NomeServidor           char(50),
                @NomeServidorDestino    char(50),
                @QtdeTotal              Float,	
--NFCapa
                @NC_NUMEROPED           numeric,
                @NC_DATAEMI             datetime,
                @NC_VENDEDOR            int,
                @NC_VLRMERCADORIA       float,
                @NC_DESCONTO            float,
                @NC_LOJAORIGEM          char(05),
                @NC_TIPONOTA            char(02),
                @NC_CONDPAG             char(04),
                @NC_AV                  int,
                @NC_CLIENTE             Numeric,
                @NC_CODOPER             char(05),
                @NC_DATAPAG             datetime,
                @NC_PGENTRA             float,
                @NC_QTDITEM             int,
                @NC_PEDCLI              numeric,
                @NC_TM                  int,
                @NC_PESOBR              float,
                @NC_PESOLQ              float,
                @NC_VALFRETE            float,
                @NC_FRETECOBR           float,
                @NC_OUTRALOJA           char(05),
                @NC_OUTROVEND           int,
                @NC_NF                  numeric,
                @NC_TOTALNOTA           float,
                @NC_DATAPED             datetime,
                @NC_BASEICMS            float,
                @NC_ALIQICMS            float,
                @NC_VLRICMS             float,
                @NC_SERIE               char(02),
                @NC_HORA                dateTime,
                @NC_TOTALIPI            float,
                @NC_PAGINANF            int,
                @NC_ValorTotalCodigoZero float,
                @NC_TotalNotaAlternativa float,
                @NC_ValorMercadoriaAlternativa float,
                @NC_VendedorLojaVenda          int, 
                @NC_LojaVenda                  char(05),
                @NC_NotaCredito                int,
                @NC_NfDevolucao                numeric,
                @NC_SerieDevolucao             char(02),
                @NC_EmiteDataSaida             char(1),
                @NC_Protocolo                  Numeric,
                @NC_NroCaixa                   Int,
                @NC_ModalidadeVenda            char(15),
                @NC_Parcelas                   int,
                @NC_TipoTransporte             char(60),
                @NC_ECF                        int,
                @NC_CPFNFP                     char(14),
                @NC_SituacaoProcesso           char(1),

--NFItens 
                @NI_NUMEROPED                  Numeric,
                @NI_DATAEMI                    datetime,
                @NI_REFERENCIA                 char(07),
                @NI_QTDE                       int,
                @NI_VLUNIT                     float,
                @NI_VLTOTITEM                  float,
                @NI_ICMS                       float,
                @NI_ITEM                       int,
                @NI_VLIPI                      float,
                @NI_DESCONTO                   float,
                @NI_PLISTA                     float,
                @NI_VALORICMS                  float,
                @NI_CODBARRA                   char(14),
                @NI_NF                         numeric,
                @NI_SERIE                      char(02),
                @NI_LOJAORIGEM                 char(05),
                @NI_ALIQIPI                    float,
                @NI_TIPONOTA                   char(02),
                @NI_BASEICMS                   float,
                @NI_DETALHEIMPRESSAO           char(02),
                @NI_PrecoUnitAlternativa       float,
                @NI_ValorMercadoriaAlternativa float,
                @NI_ReferenciaAlternativa      char(07),
                @NI_DescricaoAlternativa       char(040),
                @NI_EstoqueAntes               int,
                @NI_EstoqueDepois              int,
              
--NFItens
                @AE_Loja                       char(05),
                @AE_Referencia                 char(07),
                @AE_QtdeTotal                  int,

				@SQL1		                   Char(4000),
				@SQL2		                   Char(4000),
				@SQL3		                   Char(4000),
				@SQL4		                   Char(4000),
				@SQL5		                   Char(4000)
                

       Select @DataMovimento = (Select convert(char(10),getdate(),111))

       Declare Temp_Lojas insensitive cursor for
               Select LO_Loja,LO_NomeServidor
               from Loja where LO_Conexao = 'S' and LO_Loja = @LojaPar                
               
       Open Temp_Lojas

       Fetch Next From Temp_Lojas Into
       @Loja,@NomeServidor
       While @@Fetch_Status = 0  
        Begin

          select @SQL=' '
          Exec SP_VDA_Conexao_Cancelamento_NF @Loja, @DataMovimento
          
          Select @SQL = 'Insert Into NFCapa   
                         Select * from ' + LTrim(Rtrim(@NomeServidor)) + 'NFCapa as L where L.LojaOrigem =' + '''' +
                         LTrim(Rtrim(@Loja)) + ''''  
                      + ' and L.DataProcesso =' + '''' + @DataMovimento + '''' + ' and tiponota not in(''PA'',''PD'')
                          and situacaoProcesso = ' + '''A''' + ' and L.NF is not null and  not exists
                         (select * from NFCapa as C where C.NF = L.NF and C.Serie = L.Serie 
                          and c.lojaorigem=l.lojaorigem)'

          print('1 - ' + @sql)
          Execute (@SQL)

          select @SQL=' '
          Select @SQL = 'Insert Into NFItens   
                         Select * from ' + LTrim(Rtrim(@NomeServidor)) + 'NFItens as L where L.LojaOrigem =' + '''' +
                         LTrim(Rtrim(@Loja)) + ''''  
                      + ' and L.DataProcesso =' + '''' + @DataMovimento + '''' + ' and tiponota not in(''PA'',''PD'')
                          and situacaoProcesso = ' + '''A''' + ' and L.NF is not null and  not exists
                         (select * from NFItens as C where C.NF = L.NF and C.Serie = L.Serie 
                          and c.lojaorigem=l.lojaorigem)'

 
		  print('2 - ' + @sql)
          Execute (@SQL)


          select @SQL=' '
          Select @SQL = 'Insert Into Movimentocaixa   
                         Select * from ' + LTrim(Rtrim(@NomeServidor)) + 'Movimentocaixa  as L where L.MC_Loja =' + '''' +
                         LTrim(Rtrim(@Loja)) + ''''  
                      + ' and mc_situacaoenvio = ' + '''A''' + ' and  
                         L.MC_Documento is not null and not exists
                          (select * from Movimentocaixa as C
                          where C.MC_documento = L.MC_documento and C.mc_Serie = L.mc_Serie 
                          and c.mc_loja=l.mc_loja)'

		  print('3 - ' + @sql)
          Execute (@SQL)

          select @SQL=' '
          Select @SQL = 'Insert Into CarimboNotaFiscal   
                         Select * from ' + LTrim(Rtrim(@NomeServidor)) + 'CarimboNotaFiscal as L where L.CNF_Loja =' + '''' +
                         LTrim(Rtrim(@Loja)) + ''''  
                      + ' and L.CNF_DataProcesso =' + '''' + @DataMovimento + '''' + ' 
                          and CNF_situacaoProcesso = ' + '''A''' + ' and L.CNF_NF is not null and  not exists
                         (select * from CarimboNotaFiscal as C where C.CNF_NF = L.CNF_NF and C.CNF_Serie = L.CNF_Serie 
                          and c.CNF_loja=l.CNF_loja)'

         Execute (@SQL)

		 print('4 - ' + 'Loja set LO_Conexao= n Where LO_Loja = @Loja')
         Update Loja set LO_Conexao='N' Where LO_Loja = @Loja
         
         Fetch Next From Temp_Lojas Into
         @Loja,@NomeServidor

      end
      close TemP_Lojas
      Deallocate TemP_Lojas

      Declare TemP_NFC Insensitive Cursor For
 		Select 	
                NUMEROPED,DATAEMI,VENDEDOR,VLRMERCADORIA,
                DESCONTO,LOJAORIGEM,TIPONOTA,CONDPAG,AV,
                CLIENTE,CODOPER,DATAPAG,PGENTRA,QTDITEM,PEDCLI,
                TM,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,
                OUTROVEND,NF,TOTALNOTA,DATAPED,BASEICMS,ALIQICMS,
                VLRICMS,SERIE,HORA,TOTALIPI,PAGINANF,ValorTotalCodigoZero,
                TotalNotaAlternativa,ValorMercadoriaAlternativa,VendedorLojaVenda, 
                LojaVenda,NotaCredito,NfDevolucao,SerieDevolucao,EmiteDataSaida,
                Protocolo,NroCaixa,ModalidadeVenda,Parcelas,TipoTransporte,
                ECF,CPFNFP,SituacaoProcesso
		From   NFCapa 
		Where  Dataemi = @DataMovimento and SituacaoProcesso = 'A' 
		Order by DataProcesso

	Open  TemP_NFC 

 	Fetch Next From  TemP_NFC  into
                @NC_NUMEROPED,@NC_DATAEMI,@NC_VENDEDOR,@NC_VLRMERCADORIA,
                @NC_DESCONTO,@NC_LOJAORIGEM,@NC_TIPONOTA,@NC_CONDPAG,@NC_AV,
                @NC_CLIENTE,@NC_CODOPER,@NC_DATAPAG,@NC_PGENTRA,@NC_QTDITEM,@NC_PEDCLI,
                @NC_TM,@NC_PESOBR,@NC_PESOLQ,@NC_VALFRETE,@NC_FRETECOBR,@NC_OUTRALOJA,
                @NC_OUTROVEND,@NC_NF,@NC_TOTALNOTA,@NC_DATAPED,@NC_BASEICMS,@NC_ALIQICMS,
                @NC_VLRICMS,@NC_SERIE,@NC_HORA,@NC_TOTALIPI,@NC_PAGINANF,@NC_ValorTotalCodigoZero,
                @NC_TotalNotaAlternativa,@NC_ValorMercadoriaAlternativa,@NC_VendedorLojaVenda, 
                @NC_LojaVenda,@NC_NotaCredito,@NC_NfDevolucao,@NC_SerieDevolucao,@NC_EmiteDataSaida,
                @NC_Protocolo,@NC_NroCaixa,@NC_ModalidadeVenda,@NC_Parcelas,@NC_TipoTransporte,
                @NC_ECF,@NC_CPFNFP,@NC_SituacaoProcesso

        While (@@Fetch_status = 0) 
	   Begin

 -------------------------------------------------------------------
 -- Valida  Campos da Nota Fiscal
 -------------------------------------------------------------------


       		Select @W_Critica = 'N'
       
       		Select @W_ItensItem = (Select count(*) from NFItens
                              		Where 	LOjaOrigem =@NC_LOJAORIGEM and
                                    		NF = @NC_NF and
                                    		Serie = @NC_Serie and DataEmi = @NC_DataEmi)
       
       		Select @W_ValorCapa =(@NC_TOTALNOTA - @NC_FRETECOBR )
	      
		Select @W_ValorItens = (Select Convert(Decimal(8,2),sum(Vltotitem-Desconto))
				From NFItens
				Where LojaOrigem = @NC_LojaOrigem and NF = @NC_NF and
				      Serie = @NC_Serie and DataEmi = @NC_DataEmi)

		If @NC_CONDPAG > 3
	           Begin
			Select @W_DataDuplicata = Convert(Char(10),@NC_DataEmi,101)
	                Select 	@W_ContaCliente = (Select count(*)
	                				From 	Cliente
	                				Where 	CE_CodigoCliente= @NC_Cliente)
		   End


	        If @W_ValorCapa <> @W_ValorItens
 	           Begin 
	                Select @W_Critica = 'S'
	                Select @W_ObsCritica = 'Valor Capa <> Valor Itens'
	           End
         
	        If @W_ItensCapa <> @W_ItensItem
	          Begin 
	                Select @W_Critica = 'S'
	                Select @W_ObsCritica =  @W_ObsCritica + 'Qtde. Itens Capa <> Qtde. Itens' 
	          End

	        IF @W_ContaCliente = 0
	          Begin
	                Select @W_Critica = 'S'
	                Select @W_ObsCritica = @W_ObsCritica + 'CLiente Não Encontrado' 
	          End
  
		IF @W_Critica = 'S'
                     
		   Begin
			Update NFCapa set SituacaoProcesso ='C',CriticaProcesso = @W_ObsCritica
			Where LOjaOrigem = @NC_LojaOrigem and NF = @NC_NF and Serie = @NC_Serie 
		   End

 -------------------------------------------------------------------
 --   Atualiza Estoque
 -------------------------------------------------------------------
            select  @W_Critica ='N'   
	        IF @W_Critica= 'N' --Aguardar definição
	       
                   Begin
                   
 		     Update NFcapa set SituacaoProcesso = 'P',CriticaProcesso = ' '
		     Where LOjaOrigem = @NC_LojaOrigem and NF = @NC_NF and Serie = @NC_Serie 
 
 		     Update NFItens set SituacaoProcesso = 'P'
		     Where LOjaOrigem = @NC_LojaOrigem and NF = @NC_NF and Serie = @NC_Serie 

 		     Update MovimentoCaixa set MC_SituacaoEnvio = 'P'
		     Where MC_Loja = @NC_LojaOrigem and MC_Documento = @NC_NF and MC_Serie = @NC_Serie

 		     Update CarimboNotaFiscal set CNF_SituacaoProcesso = 'P'
		     Where CNF_Loja = @NC_LojaOrigem and CNF_NF = @NC_NF and CNF_Serie = @NC_Serie 
-------------------------------------------------------------------------------------------------------------------------
--                   SOMA QUANTIDADE DA MESMA REFERENCIA NA NOTA FISCAL                                                --
-------------------------------------------------------------------------------------------------------------------------
              
--
                     Declare Temp_AtualizaEstoque insensitive cursor for
                     Select LojaOrigem,Referencia,Sum(QTDE) From NFItens 
                     Where LojaOrigem=@NC_LojaOrigem and NF=@NC_NF and Serie=@NC_Serie and DataEmi=@NC_DATAEMI
                     Group by LojaOrigem,Referencia               
       
                     Open Temp_AtualizaEstoque

                     Fetch Next From Temp_AtualizaEstoque Into
                     @AE_Loja,@AE_Referencia,@AE_QtdeTotal
                     While @@Fetch_Status = 0  
                       Begin
         
        
                       IF Rtrim(Ltrim(@NC_TipoNota)) = 'T' or (@NC_TipoNota = 'S ')  
	                      Begin
			/*			insert into il_trans (datahora,tipo) values (GETDATE(),
							'SP_VDA_Conexao_Retaguarda ' + rtrim(ltrim(convert(char(5),@AE_Loja))) + ',' + 
							rtrim(ltrim(convert(char(10),@AE_Referencia))) + ',' + 
							rtrim(ltrim(convert(char(5),@AE_QtdeTotal))) + 'Estoque retaguarda - loja envio')
	   */
 		                   Update Estoque set Es_Estoque = (ES_Estoque - @AE_QtdeTotal) 
	                       Where  ES_Loja = rtrim(lTrim(@AE_Loja)) and
	                              ES_Referencia = @AE_Referencia 
	         
                     
	  	      IF rtrim(Ltrim(@NC_TipoNota)) ='T'
		         Begin
	                 select @W_LojaDestino =  @NC_Cliente
	                 Select @NomeservidorDestino =(Select ltrim(rtrim(LO_NomeServidor)) 
													from Loja where LO_Loja = @Loja)
          
 			
		         Update Estoque set Es_Estoque = (ES_estoque + @AE_QtdeTotal)
	                   Where  ES_Loja = ltrim(rtrim(@W_LojaDestino)) and
	                   ES_Referencia = @AE_Referencia 
			               
	/*/		            
                  Select @SQL = 'Update ' + LTrim(Rtrim(@NomeServidordestino)) + 
                       		   'EstoqueLoja set El_Estoque = (El_Estoque + ' + 
					           rtrim(ltrim(convert(char(5),@AE_QtdeTotal))) + 
					           ')Where El_Loja = ' + '''' + ltrim(rtrim(@W_LojaDestino)) + '''' + 
					           ' and El_Referencia = ' + '''' + ltrim(rtrim(@AE_Referencia)) + ''''
	
			  Execute (@SQL)  */      
 	 			/* exec SP_VDA_Conexao_Retaguarda '135'
 	 			
 	 			   Select @SQL = 'Update ' + LTrim(Rtrim(@NomeServidor)) + 
                       		   'EstoqueLoja set El_Estoque = (El_Estoque + ' + 
					           rtrim(ltrim(convert(char(5),@AE_QtdeTotal))) + ') 
	                		   Where El_Loja = ' + '''' + ltrim(rtrim(@W_LojaDestino)) + '''' + ' and
	                		   El_Referencia = ' + '''' + @AE_Referencia + ''''*/

              /*   Select @SQL= 'Exec SP_Est_Transferencia_destino ' + '''' 
                             + rtrim(Ltrim(@W_LojaDestino)) + '''' + ',' +  '''' + @AE_Referencia + ''''
                             + ',' +  Rtrim(Ltrim(convert(char(5),@AE_QtdeTotal)))  
               execute (@SQL) 
               
               insert into il_trans (datahora,tipo) values (GETDATE(),rtrim(ltrim(@SQL)))         		  
*/

		     End  
	           End
	                 
                  /* IF @NC_TipoNota = 'E ' or (@NC_TipoNota = 'ET')     
	              Begin
	          
		      Update Estoque set Es_Estoque = (ES_Estoque + @AE_QtdeTotal)
	              Where  ES_Loja = @AE_Loja and
		             ES_Referencia = @AE_Referencia  
                   End*/
             
                   IF @NC_TipoNota = 'V '
		      Begin
		       
		       Update Estoque set Es_Estoque = (ES_Estoque - @AE_QtdeTotal),
			                 Es_Venda = (ES_Venda + @AE_QtdeTotal)
		       Where  ES_Loja = @AE_Loja and
		             ES_Referencia = @AE_Referencia 
                   End
                   Fetch Next From Temp_AtualizaEstoque Into
                   @AE_Loja,@AE_Referencia,@AE_QtdeTotal

                 end
                 close Temp_AtualizaEstoque
                 Deallocate Temp_AtualizaEstoque
    
 -----------------------------------------------------------------------------
 --print 'Cria  Capa e Item NFVenda / HistoricoLoja / Titulos a Receber/Carimbo NF'
 -----------------------------------------------------------------------------

		   If @NC_TipoNota <> 'C'
		     Begin

                    
			 Exec SP_VDA_Cria_Capa_Item_NFVenda 

                        Exec SP_VDA_Cria_Carimbo_NF
			                 @NC_LojaOrigem,@NC_NF,@NC_Serie
                    
	            Exec SP_VDA_Atualiza_HistoricoLoja_Venda @NC_LojaOrigem,@NC_DataEmi
                      


--                     If @NC_CondPag > 3
--			Begin

--                           Exec SP_FIN_Cria_Titulos_Receber
--                                @NC_DataEmi,
--				@NC_LojaOrigem,
--                                @NC_NF,
--				@NC_Serie
					
--		       End

-- Comissão
                   Exec SP_Atualiza_Comissoes_Nota_Fiscal @NC_DataEmi,@NC_DataEmi
         End

      End


--************************************************************************************************
--*  
--*   Falta Margem
--************************************************************************************************

 	Fetch Next From  TemP_NFC  into
                @NC_NUMEROPED,@NC_DATAEMI,@NC_VENDEDOR,@NC_VLRMERCADORIA,
                @NC_DESCONTO,@NC_LOJAORIGEM,@NC_TIPONOTA,@NC_CONDPAG,@NC_AV,
                @NC_CLIENTE,@NC_CODOPER,@NC_DATAPAG,@NC_PGENTRA,@NC_QTDITEM,@NC_PEDCLI,
                @NC_TM,@NC_PESOBR,@NC_PESOLQ,@NC_VALFRETE,@NC_FRETECOBR,@NC_OUTRALOJA,
                @NC_OUTROVEND,@NC_NF,@NC_TOTALNOTA,@NC_DATAPED,@NC_BASEICMS,@NC_ALIQICMS,
                @NC_VLRICMS,@NC_SERIE,@NC_HORA,@NC_TOTALIPI,@NC_PAGINANF,@NC_ValorTotalCodigoZero,
                @NC_TotalNotaAlternativa,@NC_ValorMercadoriaAlternativa,@NC_VendedorLojaVenda, 
                @NC_LojaVenda,@NC_NotaCredito,@NC_NfDevolucao,@NC_SerieDevolucao,@NC_EmiteDataSaida,
                @NC_Protocolo,@NC_NroCaixa,@NC_ModalidadeVenda,@NC_Parcelas,@NC_TipoTransporte,
                @NC_ECF,@NC_CPFNFP,@NC_SituacaoProcesso
	End

   -- Exec SP_VDA_Conexao_Cancelamento_NF '315','2012/07/03'
      Exec SP_SomaConsolidado

 -----------------------------------------------------------------------------
 print ('')
 print ('INSERINDO NOTA FISCAL ELETRONICA')
 -----------------------------------------------------------------------------

--print ('teste' + @sql)

print ('SELECT count(tm) FROM nfcapa where tm = 88 and DATAEMI = ''' + @DataMovimento + '''
and tiponota not in(''PA'',''PD'') and LOJAORIGEM = ''' + @loja + '''
and serie = ''NE''')

WHILE (SELECT count(tm) FROM nfcapa where tm = 88 and DATAEMI =  @DataMovimento
and tiponota not in('PA','PD') and LOJAORIGEM = @loja 
and serie = 'NE') > 0

  BEGIN
	DECLARE @NF AS char(10)
	--DECLARE @Loja AS char(10)
	DECLARE @Serie as char(10)

	select top 1 @NF = nf, @Loja = lojaorigem, @serie = serie 
	FROM nfcapa 
	where tm = 88 and DATAEMI = @DataMovimento
	and tiponota not in('PA','PD') 
	and LOJAORIGEM = @loja
	and serie = 'NE'

	update nfcapa set tm = 99 
	where tm = 88 and DATAEMI = @DataMovimento 
	and nf = @NF and serie = @serie and lojaorigem = @loja
	and dataemi =  @DataMovimento
	and serie = 'NE'

	set @sql = 'SP_VDA_Cria_NFe ''' + rtrim(@loja) + '''' + ', ' +  '''' + rtrim(@NF) + '''' + ', ' + '''' +  rtrim('NE') + '''' + ', ' + '''' + rtrim('@CARIMBO') + ''''
	print (@sql)
	Execute (@SQL)
	
  END

 -----------------------------------------------------------------------------
	
	If @@ERROR <> 0
	   Begin	
	   	Rollback Transaction		
	   	Return
	   End

	Close TemP_NFC
	Deallocate TemP_NFC
End
/*
drop table Temp_nfc
       exec SP_VDA_Conexao_Retaguarda '116'
update loja set lo_conexao='S' where lo_loja  in('333')
delete loja where lo_loja not in ('354','271','316','CONSO','CD')
select qtditem,nf,serie,situacaoprocesso,cliente,* from nfcapa  where dataemi='2011/09/24'
delete nfitens where dataemi='2011/09/24'
update nfcapa set situacaoprocesso='A' where dataemi='2011/09/22' and nf=12
select es_venda,es_estoque,* from estoque where es_estoque <>1000
select * from capanfvenda
select * from fin_titulos
select * from loja
update estoque set es_estoque=1000,ES_VENDA=0
insert into controlefec
(CV_DataMovimento,CV_PastaEnvioLoja,CV_PastaRecebeLoja,
       CV_PastaDestinoNaLoja,CV_UltimoGiro,CV_DataEstoqueBaixo,
       CV_PastaInternet,CV_Empresa,CV_HoraManutencao,
       CV_SituacaoFechamento,CV_Duplex,CV_DataUltimoAuditor,
       CV_ProcessandoFecMes) values  
       ('2011/01/01','2011/01/01','','',0,0,'',1,0,'A',0,0,'A')

cv_datamovimento) values ('2011/09/24')
select * from controlefec
select convert(gatedate()
update loja set lo_conexao='S' where lo_loja='116'
select  lo_conexao,*from loja
select * from nfcapa
SELECT * FROM NFITENS
SELECT * FROM ESTOQUE WHERE ES_ESTOQUE<>1000
DELETE NFCAPA
DELETE NFITENS
delete itemnfvenda
delete capanfvenda
select * from estoque

update estoque set es_estoque=0 where es_transito is null null
select * from itemnfvenda
select NF,SERIE,TIPONOTA,* from nfitens where DATAEMI='2012/06/05' order by NF
select NF,SERIE,TIPONOTA,* from [svglobo2].[dmac_loja].[dbo].nfcapa where DATAEMI='2012/06/05' order by NF
select NF,SERIE,TIPONOTA,referencia,qtde 
from [svglobo2].[dmac_loja].[dbo].nfitens 
where DATAEMI='2012/06/05' and referencia ='2590920' group by  NF,SERIE,TIPONOTA,referencia having count(*) > 1   order by NF
select * from  estoque where es_referencia='2590920'
select * from historicoloja

exec SP_VDA_Conexao_Retaguarda '28'
*/



