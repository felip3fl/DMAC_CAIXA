/*

CREATE TABLE NFE_fat(
eLoja char(5),
eNF char(10),
eSerie char(2),
Situacao char(1),
nFat varchar(60),
vOrig float,
vDesc float,
vLiq float
)

CREATE TABLE NFE_dup(
eLoja char(5),
eNF char(10),
eSerie char(2),
Situacao char(1),
nDuo varchar(60),
dVend datetime,
vDup float
)

drop table nfe_cobr



CREATE TABLE [dbo].[CodigoOperacao](
	[CF_CodigoOperacao] [smallint] NOT NULL,
	[CF_CodigoOperacaoAux] [smallint] NOT NULL,
	[CF_Descricao] [varchar](60) NOT NULL,
	[CF_EntradaSaida] [char](1) NOT NULL,
	[CF_Transferencia] [char](1) NOT NULL,
	[CF_SimplesRemessa] [char](1) NOT NULL,
	[CF_Interestadual] [char](1) NOT NULL,
	[CF_Importacao] [char](1) NOT NULL,
	[CF_Devolucao] [char](1) NOT NULL,
	[CF_CodigoTributo] [char](2) NOT NULL,
	[CF_TipoCodigo] [varchar](3) NULL,
	[CF_CodigoOperacaoNovo] [smallint] NOT NULL DEFAULT (0),
 CONSTRAINT [PKCF_CodigoOperacao] PRIMARY KEY CLUSTERED 
(
	[CF_CodigoOperacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]


CREATE TABLE [dbo].[CondicaoPagto](
	[CP_CodigoCondicao] [smallint] NOT NULL,
	[CP_Descricao] [varchar](50) NOT NULL,
	[CP_QuantidadeParcelas] [smallint] NOT NULL,
	[CP_TipoCondicao] [char](2) NOT NULL,
	[CP_Parcelas] [varchar](25) NOT NULL,
	[CP_VendaCompra] [char](1) NOT NULL,
	[CP_Intervalo] [smallint] NULL CONSTRAINT [DF__CondicaoP__CP_In__16451E08]  DEFAULT (0),
 CONSTRAINT [PKCP_CondPagto] PRIMARY KEY CLUSTERED 
(
	[CP_CodigoCondicao] ASC,
	[CP_VendaCompra] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]



CREATE TABLE [dbo].[Duplicata](
	[DP_Loja] [char](5) NOT NULL,
	[DP_NotaFiscal] [int] NOT NULL,
	[DP_Serie] [char](2) NOT NULL,
	[DP_Sequencia] [tinyint] NOT NULL,
	[DP_CodigoCliente] [int] NOT NULL,
	[DP_DataEmissao] [datetime] NOT NULL,
	[DP_Vendedor] [smallint] NOT NULL,
	[DP_Banco] [smallint] NOT NULL,
	[DP_DocumentoBancario] [varchar](10) NULL,
	[DP_ValorDuplicata] [float] NOT NULL,
	[DP_DataVencimento] [datetime] NOT NULL,
	[DP_NotaCredito] [smallint] NULL,
	[DP_Abatimento] [float] NOT NULL CONSTRAINT [DF__Duplicata__DP_Ab__1D072A30]  DEFAULT (0),
	[DP_Desconto] [float] NOT NULL CONSTRAINT [DF__Duplicata__DP_De__1DFB4E69]  DEFAULT (0),
	[DP_Despesas] [float] NOT NULL CONSTRAINT [DF__Duplicata__DP_De__1EEF72A2]  DEFAULT (0),
	[DP_Juros] [float] NOT NULL CONSTRAINT [DF__Duplicata__DP_Ju__1FE396DB]  DEFAULT (0),
	[DP_ValorPago] [float] NOT NULL CONSTRAINT [DF__Duplicata__DP_Va__20D7BB14]  DEFAULT (0),
	[DP_DataPagamento] [datetime] NULL,
	[DP_DataBaixa] [datetime] NULL,
	[DP_DataCartorio] [datetime] NULL,
	[DP_Historico] [varchar](250) NULL,
	[DP_TipoPagamento] [char](2) NULL,
	[DP_Agrupamento] [int] NULL,
	[DP_Situacao] [char](1) NOT NULL CONSTRAINT [DF__Duplicata__DP_Si__21CBDF4D]  DEFAULT (''''A''''),
 CONSTRAINT [PKDP_Duplicata] PRIMARY KEY CLUSTERED 
(
	[DP_Loja] ASC,
	[DP_NotaFiscal] ASC,
	[DP_Serie] ASC,
	[DP_Sequencia] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]



*/


CREATE   Procedure [dbo].[Sp_Cria_Duplicatas]
	@Loja		Char(5),
	@DataInicial	Char(20),
	@DataFinal	Char(20)

As


Declare		@LojaVenda		VarChar(5),
		@Serie			VarChar(2),
		@SerieAlm		VarChar(2),
		@NotaFiscal		Int,
		@SerieLida		Char(2),
		@LojaOrigem		Char(5),
		@Cliente		Int,
		@DataEmissao		DateTime,
		@CodigoVendedor		SmallInt,
		@CondicaoPagamento	SmallInt,
		@DataPagamento		DateTime,
		@PagamentoEntrada	Float,
		@TotalNota		Float,
		@Descricao		VarChar(60),
		@QuantidadeParcelas	SmallInt,
		@TipoCondicao		Char(2),
		@TipoPessoa		Char(1),
		@PagamentoCarteira	Char(1),
		@MaiorCompra		Float,
		@DataUltimaCompra	DateTime,
		@Inicio			SmallInt,
		@DataBase		DateTime,
		@ValorParcela		Float,
		@Sequencia		SmallInt,
		@ParcelaInt		Int,
		@Dias			SmallInt,
		@Situacao	Char(1)


Begin


	Begin Transaction

	If @Loja = ''''ALM01''''
	  Begin
		Select	@LojaVenda = ''''353'''',
			--@Serie = ''''S2'''',
			@SerieAlm = ''''''''
	  End
	Else
	  Begin
		If @Loja = ''''353''''
		  Begin
			Select	@LojaVenda = @Loja,
				--@Serie = ''''%'''',
				@SerieAlm = ''''S2''''
		  End

		Else
		  Begin
			
			Select	@LojaVenda = @Loja,
				--@Serie = ''''%'''',
				@SerieAlm = ''''''''
		  End
	  End

	--Select @serieNota	

	--print (''''DUPLICATA ETAPA 1'''')

		Declare curVenda Insensitive Cursor For
		
				Select  NF,
			Serie,
			LojaOrigem,
			Cliente,
			DataEmi,
			Vendedor,
			CondPag,
			DataPag,
			PgEntra,
			TotalNota,
			CP_Parcelas,
			CP_QuantidadeParcelas,
			CP_TipoCondicao,
			CE_TipoPessoa,
			CE_PagamentoCarteira,
			CE_MaiorCompra,
			CE_DataUltimaCompra
		From	NFCapa,
			CondicaoPagto,
			FIN_Cliente,
			CodigoOperacao
		Where	CondPag = CP_CodigoCondicao and
			Cliente = CE_CodigoCliente and
			CondPag > 3 and
			DataEmi between @DataInicial and @DataFinal and
			CF_TipoCodigo = ''''V'''' and 
			CP_VendaCompra = ''''V'''' and
			SituacaoEnvio =''''A'''' and
			TipoNota <> ''''C'''' and 
			LojaOrigem = @Loja and
			Serie = ''''NE''''



	Open curVenda
	Fetch Next From curVenda Into
		@NotaFiscal,
		@SerieLida,
		@LojaOrigem,
		@Cliente,
		@DataEmissao,
		@CodigoVendedor,
		@CondicaoPagamento,
		@DataPagamento,
		@PagamentoEntrada,
		@TotalNota,
		@Descricao,
		@QuantidadeParcelas,
		@TipoCondicao,
		@TipoPessoa,
		@PagamentoCarteira,
		@MaiorCompra,
		@DataUltimaCompra


	--print (''''DUPLICATA'''')
	While @@FETCH_STATUS = 0
	  Begin

		Delete 	Duplicata 
		Where	DP_Loja = @LojaOrigem and
			DP_NotaFiscal = @NotaFiscal and
			DP_Serie = @SerieLida

		If @@Error <> 0 Goto Desfaz


		Select  @Inicio = 1,
			@Sequencia = 0,
			@DataBase = (Case @TipoCondicao
					When ''''DD'''' Then @DataEmissao
					When ''''DL'''' Then DateAdd(day, 1, @DataEmissao)
					When ''''DI'''' Then @DataPagamento
				     End
				    )

		If @PagamentoEntrada > 0
		  Begin

			Select 	@TotalNota = @TotalNota - @PagamentoEntrada,
				@Sequencia = @Sequencia + 1
			

			Insert Into Duplicata (
				DP_Loja,
				DP_NotaFiscal,
				DP_Serie,
				DP_Sequencia,
				DP_CodigoCliente,
				DP_DataEmissao,
				DP_Vendedor,
				DP_Banco,
				DP_ValorDuplicata,
				DP_DataVencimento,
				DP_Situacao
			)
			Select	@LojaOrigem,
				@NotaFiscal,
				@SerieLida,
				@Sequencia,
				(Case
					When @CondicaoPagamento = 2 Then 999998
					When @CondicaoPagamento = 3 Then 999997
					Else @Cliente
				 End
				),
				@DataEmissao,
				@CodigoVendedor,
				(Case
					When @LojaOrigem = ''''800'''' Then 315
					Else 800
				 End
				),
				@PagamentoEntrada,
				@DataEmissao,
				''''W''''

			If @@Error <> 0 Goto Desfaz

		  End


		Select 	@ParcelaInt = Abs(Ceiling(-(@TotalNota / @QuantidadeParcelas) * 100))

		Select 	@ValorParcela = @ParcelaInt / 100.0

		--print (''''WHILE'''')
		--PRINT @Descricao
		
		While @QuantidadeParcelas > 0
		  Begin
			Select 	@Sequencia = @Sequencia + 1,
				@Dias = Convert(SmallInt, SubString(@Descricao, @Inicio, 3))

			If @QuantidadeParcelas = 1
			  Begin
				Select @ValorParcela = @TotalNota
			  End

			--print (''''INSERIR DUPLICATA'''')
			Insert Into Duplicata (
				DP_Loja,
				DP_NotaFiscal,
				DP_Serie,
				DP_Sequencia,
				DP_CodigoCliente,
				DP_DataEmissao,
				DP_Vendedor,
				DP_Banco,
				DP_ValorDuplicata,
				DP_DataVencimento,
				DP_Situacao
			)
			Select	@LojaOrigem,
				@NotaFiscal,
				@SerieLida,
				@Sequencia,
				(Case
					When @CondicaoPagamento = 2 Then 999998
					When @CondicaoPagamento = 3 Then 999997
					Else @Cliente
				 End
				),
				@DataEmissao,
				@CodigoVendedor,
				(Case
					When @CondicaoPagamento = 2 Then 997
					When @CondicaoPagamento = 3 Then 998
					When @CondicaoPagamento > 3 and @PagamentoCarteira = ''''S'''' and @LojaOrigem <> ''''800'''' Then 800
					When @CondicaoPagamento > 3 and @TipoPessoa = ''''U'''' and @LojaOrigem <> ''''800'''' Then 802
					When @CondicaoPagamento > 3 and @TipoPessoa = ''''A'''' and @LojaOrigem <> ''''800'''' Then 804
					When @CondicaoPagamento > 3 and @LojaOrigem = ''''800'''' Then 314
					When @LojaOrigem = ''''85'''' then 422
                                        Else 422
				 End
				),
				@ValorParcela,
				DateAdd(day, @Dias, @DataBase),
				''''W''''

				If @@Error <> 0 Goto Desfaz


			Select 	@QuantidadeParcelas = @QuantidadeParcelas - 1,
				@Inicio = @Inicio + 3,
				@TotalNota = @TotalNota - @ValorParcela

		  End



		Fetch Next From curVenda Into
			@NotaFiscal,
			@SerieLida,
			@LojaOrigem,
			@Cliente,
			@DataEmissao,
			@CodigoVendedor,
			@CondicaoPagamento,
			@DataPagamento,
			@PagamentoEntrada,
			@TotalNota,
			@Descricao,
			@QuantidadeParcelas,
			@TipoCondicao,
			@TipoPessoa,
			@PagamentoCarteira,
			@MaiorCompra,
			@DataUltimaCompra

	  End


	Close curVenda
	Deallocate curVenda


Final:



	Commit Transaction

	Return(0)


Desfaz:

	Close curVenda
	Deallocate curVenda


	Rollback Transaction

	Return(1)


End





USE [DMAC_LOJA]
GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_Processos_Venda]    Script Date: 05/07/2016 14:43:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




ALTER           Procedure [dbo].[SP_Atualiza_Processos_Venda]
                          @NumeroPedido integer,
                          @NumeroNF     integer,
                          @Protocolo    integer,
                          @NumeroCaixa integer
As
Declare @data                   char(10),
	    @condPag                nvarchar(8),
        @serie					char(3)
		
SELECT @data = convert(char(10),getdate(),111)
select top 1 @serie = serie, @condPag = CONDPAG from NFCapa where NUMEROPED = @NumeroPedido

 IF EXISTS (select * from sysobjects where name = '#TempAtualiza_Saida_Estoque ' and upper(xtype) = 'U')
        Drop Table #TempAtualiza_Saida_Estoque  

Begin Transaction
    Create Table  #TempAtualiza_Saida_Estoque 
          (TPS_CodigoProduto   Char(16),
           TPS_Quantidade      Numeric)
          
    Insert Into #TempAtualiza_Saida_Estoque (TPS_CodigoProduto,TPS_Quantidade)
           select Referencia ,sum(Qtde)
           from NFItens  Where NumeroPed=@NumeroPedido and TipoNota in ('PA','TA') 
           group by Referencia

      if substring((select top 1 CAST(CODOPER as varchar(10)) from NFCapa where NumeroPed = @NumeroPedido),1,1) = '1'
            Update EstoqueLoja set EL_Estoque =(EL_Estoque + TPS_Quantidade)
                     From EstoqueLoja,#TempAtualiza_Saida_Estoque 
                     Where EL_Referencia = TPS_CodigoProduto collate SQL_latin1_general_cp1_ci_as

      if substring((select top 1 CAST(CODOPER as varchar(10)) from NFCapa where NumeroPed = @NumeroPedido),1,1) <> '1'
            Update EstoqueLoja set EL_Estoque =(EL_Estoque - TPS_Quantidade)
                     From EstoqueLoja,#TempAtualiza_Saida_Estoque 
                     Where EL_Referencia = TPS_CodigoProduto collate SQL_latin1_general_cp1_ci_as           

    Update NfItens set 
    NF = @NumeroNF,
    TipoNota=(select CASE TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end), 
    DATAEMI = @data, 
    dataprocesso = @data  
    Where Numeroped = @NumeroPedido
       
    Update NfCapa set 
    NF = @NumeroNF,
    TipoNota=(select CASE TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end),
    hora=  convert(varchar(10),getdate(),108),
    Protocolo = @Protocolo,
    NroCaixa =  @NumeroCaixa,
    DATAEMI = @data, 
    dataprocesso = @data
    Where numeroped = @NumeroPedido
      
    Update CarimboNotaFiscal set 
    CNF_NF = @NumeroNF,
    CNF_DataProcesso = @data
    Where CNF_NumeroPed = @NumeroPedido      
    
    Update movimentocaixa set
    MC_Documento = @NumeroNF,
    MC_DataProcesso = @data,
    MC_TipoNota=(select CASE MC_TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end)
    where MC_Pedido = @NumeroPedido
    
    EXEC SP_Atualiza_Cliente_NFCAPA_Local @NumeroNF,@serie

	if @condPag > 3
		EXEC Sp_Cria_Duplicatas @NumeroNF, @data,@data

   If @@Error <> 0 
      Begin
         Rollback Transaction
      End 
   Else 
      Begin  
         Commit Transaction
      End
      
    
    

 










