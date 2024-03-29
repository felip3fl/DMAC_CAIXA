USE [DMAC_LOJA_BACKUP_2]
GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Cria_Cupom]    Script Date: 15/04/2016 14:57:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



ALTER PROCEDURE [dbo].[SP_VDA_Cria_Cupom]
	@pedido	 numeric

AS
	declare

		@espacoP as char(4),

		--IDE
		@ide_NSU numeric,
		@ide_numerocaixa char(3),

		--DACFE
		@dacfe_impressora varchar(100),
		@dacfe_retornasp char(1),
		@dacfe_imprTef char(1),
		@dacfe_tipoImpr char(1),
		@dacfe_IMPRSOTEF char(3),

		--EMIT
		@emit_cnpj char(14),
		@emit_IE char(12),
		@emit_INDRATISSQN char(1),

		--DEST
		@dest_cgc char(14),

		--ICMSTOT
		@icmstot_VDESCSUBTOT numeric(10,2),
		@icmstot_VACRESSUBTOT numeric(10,2),
		@icmstot_VTOTTRIB numeric(10,2),

		--INFADIC
		@infadic_infclp varchar(5000),
		@infadic_infclpTEMP varchar(1000),

		--PAG
		@pag_TPAG char(2),
		@pag_VPAG numeric(10,2),
		@pag_CADMC char(3),

		--PROD
		@prod_cprod varchar(60),
		@prod_xprod varchar(120),
		@prod_ncm varchar(8),
		@prod_cfop varchar(4),
		@prod_ucom varchar(6),
		@prod_qcom float,
		@prod_vuncom numeric(10,2),
		@prod_indregra char(1),

		--ICMS
		@icms_orig char(1),
		@icms_cst char(2),
		@icms_picms float,

		--PIS
		@pis_cst char(2),
		@pis_vbc numeric(10,2),
		@pis_ppis numeric(10,2),

		--COFIN
		@cofins_cst char(2),
		@cofins_vbc numeric(10,2),
		@cofins_pcofins numeric(10,2)



Begin


--insert into sat_capa (cp_nsu, cp_caixa, cp_impressora, cp_tipoImpr, cp_cnpjLoja, cp_inscricaoLoja, cp_cnpjCliente, cp_cpfCliente, cp_NomeCliente)
--	 select           @nsu,   @caixa, @impressora, 2, lo_cgc, LO_InscricaoEstadual,(case when len(CGCCLI ) = 14 then CGCCLI else '' end), (case when len(CGCCLI)=11 then CGCCLI else '' end), (case when len(NOMCLI) > 2 then NOMCLI else '' end) 
--	   from loja,nfcapa 
--	  where lo_loja= lojaorigem 

--insert into sat_itens (ci_nsu, ci_referencia, ci_decricao, ci_ncm, ci_cfop, ci_ucom, ci_qcom, ci_vuncom, ci_vdesconto, ci_origem, ci_cst, ci_pIcms, ci_ppis, ci_pcofins)
--	 select           @nsu, REFERENCIA, pr_Descricao, pr_ClasseFiscal, CFOP, pr_Unidade, QTDE, VLUNIT, DESCONTO, SUBSTRING(pr_cst,1,1), icms, VALORICMS,  PisCofins , 0
--	  from nfitens,produto
--	 where pr_referencia = referencia

	--(case when len(CGCCLI ) = 14 then CGCCLI else '' end), (case when len(CGCCLI)=11 then CGCCLI else '' end), (case when len(NOMCLI) > 2 then NOMCLI else '' end) 

	delete sat_nf  where snf_pedido = @pedido
	select @espacoP = '    '

			--@icmstot_VDESCSUBTOT float,
		--@icmstot_VACRESSUBTOT float,
		--@icmstot_VTOTTRIB float,

	select 
	@ide_NSU = CONCAT(right(year(dataemi),2),REPLICATE('0', 2 - LEN(month(dataemi))) + RTrim(month(dataemi)), NUMEROPED),
	@ide_numerocaixa = REPLICATE('0', 3 - LEN(nroCaixa)) + RTrim(nroCaixa), 
	@emit_cnpj = LO_CGC,  
	@emit_IE = LO_InscricaoEstadual,
	@dest_cgc = CPFNFP,
	@icmstot_VDESCSUBTOT = desconto,
	@icmstot_VACRESSUBTOT = 0,
	@icmstot_VTOTTRIB = ((TOTALNOTA * 26.25) / 100)
	from nfcapa, loja
	where 
	lo_loja= lojaorigem 
	and numeroped = @pedido

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[IDE]','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NSU','=',@ide_NSU,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NUMEROCAIXA','=',@ide_numerocaixa,@pedido)		

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	SELECT 
		@dacfe_impressora = cs_dacfe,
		@dacfe_retornasp = '3',
		@dacfe_imprTef = 'IMPRTEF',
		@dacfe_tipoImpr = '2',
		@dacfe_IMPRSOTEF = 'não'
	FROM CONTROLESERIE where cs_nroCaixa = @ide_numerocaixa
	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DACFE]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRESSORA','=',@dacfe_impressora,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRESSORA','=',@dacfe_impressora,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'RETORNASP','=',@dacfe_retornasp,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRTEF','','',@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'TIPOIMPR','=',@dacfe_tipoImpr,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRSOTEF','=',@dacfe_IMPRSOTEF,@pedido)					

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--DESENV
	select @emit_cnpj = '61099008000141'
	select @emit_IE = '111111111111'
	select @emit_INDRATISSQN = 'N'

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[EMIT]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CNPJ','=',@emit_cnpj,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IE','=',@emit_IE,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INDRATISSQN','=',@emit_INDRATISSQN,@pedido)		
				

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if len(@dest_cgc) > 1
	begin
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DEST]','','',@pedido)							

		if len(@dest_cgc) >= 14 						
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CNPJ','=',@dest_cgc,@pedido)					

		if len(@dest_cgc) < 14
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CPF','=',@dest_cgc,@pedido)					
	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMSTOT]','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VDESCSUBTOT','=',@icmstot_VDESCSUBTOT,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VACRESSUBTOT','=',@icmstot_VACRESSUBTOT,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VTOTTRIB','=',@icmstot_VTOTTRIB,@pedido)		

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	select @infadic_infclp = ''
	select @infadic_infclpTEMP = ''

	Create	Table #TEMPInfAdic	(  
	infadic_infclp varchar(5000))
	Insert Into #TEMPInfAdic(infadic_infclp)
	Select CNF_Carimbo from CarimboNotaFiscal where CNF_NumeroPed = @pedido
	Declare curInfAdic Insensitive Cursor For
	Select 	infadic_infclp from #TEMPInfAdic
	Open curInfAdic
	Fetch Next From curInfAdic Into @infadic_infclpTEMP                                  
	While @@Fetch_Status=0
	Begin

		--if (select count(CNF_Carimbo) from CarimboNotaFiscal where CNF_NumeroPed = @pedido) > 0
		select @infadic_infclp = @infadic_infclpTEMP + ' - '


	Fetch Next From curInfAdic Into 
	@infadic_infclpTEMP      

	End

	Close curInfAdic
	Deallocate curInfAdic
	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[INFADIC]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INFCLP','=',@infadic_infclp,@pedido)		


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Create	Table #TEMPPag	(  
	pag_TPAG char(2),
	pag_VPAG numeric(10,2),
	pag_CADMC char(3))
	Insert Into #TEMPPag(pag_TPAG, pag_VPAG, pag_CADMC)
		Select MO_TipoPag, MC_Valor, MC_Agencia
		From MovimentoCaixa, Modalidade where MC_Pedido = @pedido AND 	MO_Grupo = MC_Grupo AND 	MO_Grupo < 11000
	Declare curPag Insensitive Cursor For
	Select 	pag_TPAG, pag_VPAG, pag_CADMC from #TEMPPag
	Open curPag
	Fetch Next From curPag Into @pag_TPAG, @pag_VPAG, @pag_CADMC                                  
	While @@Fetch_Status=0
	Begin

		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PAG]','','',@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'TPAG','=',@pag_TPAG,@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VPAG','=',@pag_VPAG,@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CADMC','=',@pag_CADMC,@pedido)		

	Fetch Next From curPag Into 
	@pag_TPAG, @pag_VPAG, @pag_CADMC      

	End

	Close curPag
	Deallocate curPag
	
	--select * from Modalidade

	--Tipo de TPAG
	--01 – Dinheiro;
	--02 – Cheque.
	--03 – Cartão de Crédito;
	--04 – Cartão de Débito;
	--05 – Crédito Loja;
	--10 – Vale Alimentação;
	--11 – Vale Refeição;
	--12 – Vale Presente;
	--13 – Vale Combustível;
	--99 – Outros.

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Create	Table #TempItens	(  
	prod_cprod varchar(60),	prod_xprod varchar(120), prod_ncm varchar(8), prod_cfop varchar(4), prod_ucom varchar(6), prod_qcom float, 
	prod_vuncom numeric(10,2), prod_indregra char(1), icms_cst char(2), icms_orig CHAR(1), icms_picms float, 
	pis_vbc numeric(10,2), pis_cst char(2), pis_ppis numeric(10,2), 
	cofins_vbc numeric(10,2), cofins_cst char(2), cofins_pcofins numeric(10,2))

	Insert Into #TempItens(prod_cprod, prod_xprod, prod_ncm, prod_cfop, prod_ucom, prod_qcom, prod_vuncom, prod_indregra, icms_cst, 
	icms_orig, icms_picms, 
	pis_vbc, pis_cst, pis_ppis, 
	cofins_vbc,cofins_cst,cofins_pcofins)

	Select pr_referencia, pr_descricao, pr_classeFiscal, cfop, 'PC', qtde, vlunit, 'A',  REPLICATE('0', 2 - LEN(CSTICMS)) + RTrim(CSTICMS),  '0',  ICMSAplicado, 
	VLUNIT2,'01', 1.65, 
	VLUNIT2,'01',7.60
	
	From produtoloja,nfitens where  pr_referencia = referencia and numeroped = @pedido
	Declare CurItens Insensitive Cursor For
	Select 	prod_cprod,prod_xprod,prod_ncm,
	prod_cfop,prod_ucom,prod_qcom,
	prod_vuncom,prod_indregra,icms_cst,
	icms_orig,icms_picms,pis_vbc, pis_cst, pis_ppis, 
	cofins_vbc,cofins_cst,cofins_pcofins from #TempItens
	Open CurItens
	Fetch Next From CurItens Into @prod_cprod, 	@prod_xprod, @prod_ncm, 
	@prod_cfop, @prod_ucom, @prod_qcom, 
	@prod_vuncom, @prod_indregra, @icms_cst, 
	@icms_orig, @icms_picms,	
	@pis_vbc, @pis_cst, @pis_ppis, 
	@cofins_vbc, @cofins_cst , @cofins_pcofins                          
	While @@Fetch_Status=0
	Begin
	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DET]','','',@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PROD]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CPROD','=',@prod_cprod,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'XPROD','=',@prod_xprod,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NCM','=',@prod_ncm,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CFOP','=',@prod_cfop,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'UCOM','=',@prod_ucom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'QCOM','=',@prod_qcom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VUNCOM','=',@prod_vuncom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INDREGRA','=',@prod_indregra,@pedido)	
		
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[IMPOSTO]','','',@pedido)							

		IF @icms_cst = '40' or @icms_cst = '41' or @icms_cst = '60'
		begin
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMS40]','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'ORIG','=',@icms_orig,@pedido)
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@icms_cst,@pedido)
		end

		IF @icms_cst = '00' or @icms_cst = '20' or @icms_cst = '90'
		begin
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMS00]','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'ORIG','=',@icms_orig,@pedido)
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@icms_cst,@pedido)
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'PICMS','=',@icms_picms,@pedido)
		end

		--select @pis_vbc = (@pis_vbc * @pis_ppis) / 100
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PISALIQ]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@pis_cst,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VBC','=',@pis_vbc,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'PPIS','=',@pis_ppis,@pedido)
		
		--select @cofins_vbc = (@cofins_vbc * @cofins_pcofins) / 100
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[COFINSALIQ]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@cofins_cst,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VBC','=',@cofins_vbc,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'pcofins','=',@cofins_pcofins,@pedido)


	Fetch Next From CurItens Into 
	@prod_cprod, @prod_xprod, @prod_ncm, @prod_cfop, @prod_ucom, @prod_qcom, @prod_vuncom, @prod_indregra, 
	@icms_cst,@icms_orig,@icms_picms,@pis_vbc, @pis_cst, @pis_ppis, @cofins_vbc, @cofins_cst , @cofins_pcofins  

	End

	Close CurItens
	Deallocate CurItens

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--select * from nfitens
	
END



/*
sp_help SAT_NF
	SELECT * FROM SAT_NF
	TRUNCATE TABLE SAT_NF

	SELECT MC_Pedido, COUNT(*) FROM MOVIMENTOCAIXA WHERE MC_data > '2016/02/01' AND MC_serie = 'CF4' GROUP BY MC_Pedido HAVING COUNT(*) > 2

	SELECT TOP 1 * FROM NFCAPA 

	exec SP_VDA_Cria_Cupom '67012'
	select * from carimbonotafiscal where cnf_numeroped = 67012

	select * from nfcapa where dataemi = '2016/03/07' and serie = 'CF3' AND cpfnfp <> ''

*/

