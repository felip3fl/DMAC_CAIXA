USE [DMAC_Loja]
GO
/****** Object: StoredProcedure [dbo].[SP_Atualiza_Processos_Venda] Script Date: 19/07/2016 15:35:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




ALTER Procedure [dbo].[SP_Atualiza_Processos_Venda]
@NumeroPedido integer,
@NumeroNF integer,
@Protocolo integer,
@NumeroCaixa integer
As
Declare @data char(10),
@condPag nvarchar(8),
@serie char(3),
@loja varchar(5)

SELECT @data = convert(char(10),getdate(),111)
select top 1 @serie = serie, @condPag = CONDPAG from NFCapa where NUMEROPED = @NumeroPedido

IF EXISTS (select * from sysobjects where name = '#TempAtualiza_Saida_Estoque ' and upper(xtype) = 'U')
Drop Table #TempAtualiza_Saida_Estoque

Begin Transaction
Create Table #TempAtualiza_Saida_Estoque
(TPS_CodigoProduto Char(16),
TPS_Quantidade Numeric)

Insert Into #TempAtualiza_Saida_Estoque (TPS_CodigoProduto,TPS_Quantidade)
select Referencia ,sum(Qtde)
from NFItens Where NumeroPed=@NumeroPedido and TipoNota in ('PA','TA')
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
hora= convert(varchar(10),getdate(),108),
Protocolo = @Protocolo,
NroCaixa = @NumeroCaixa,
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

select @loja = (select top 1 CTS_Loja from ControleSistema)
if @condPag > 3
EXEC Sp_Cria_Duplicatas @loja, @data,@data

If @@Error <> 0
Begin
Rollback Transaction
End
Else
Begin
Commit Transaction
End




