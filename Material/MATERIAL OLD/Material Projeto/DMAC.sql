Select * from loja

select * from controle
update controle set ct_loja = 271

--ALTERAR CONTROLE PARA CONTROLESISTEMA
SELECT * FROM LOJAS
SELECT lo_Razao, * FROM DROP_LOJAS


--DROP TABLE NFCAPA
SELECT * FROM NFCAPA

--DROP TABLE NFItens
SELECT * FROM NFITENS

truncate table MovimentoCaixa
truncate table lembreme
truncate table fornecedor
select * from fornecedor

truncate table estoqueLoja
SELECT * FROM estoqueLoja
INSERT INTO EstoqueLoja(EL_Loja, EL_Referencia, EL_CodigoFornecedor, EL_Estoque,  EL_EstoqueAnterior) SELECT es_loja, es_referencia, fo_codigoFornecedor, es_estoque, es_estoqueAnterior
from demeo..estoque, demeo..Fornecedor, demeo..produto where es_loja = '271' and FO_CodigoFornecedor = PR_CodigoFornecedor and PR_Referencia = ES_Referencia

truncate table DivergenciaEstoque
truncate table ControleCaixa

select * from produtoBarras

select * from produtoLoja
truncate table produtoLoja

select * from dmac..produtoLoja
select * from estoqueLoja



insert into produtoloja select pr_Referencia , pr_CodigoFornecedor , pr_Descricao, pr_Classe , pr_Bloqueio , pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , pr_ICMSSaida, pr_CodigoReducaoICMS , pr_CustoMedio1 , pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador , pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv  , pr_HoraManutencao, pr_CodigoProdutoNoFornecedor , pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , pr_ICMSEntrada , pr_IcmPdvEntrada, pr_CST , pr_GarantiaEstendida , pr_GarantiaFabricante , pr_IndicePreco
from dmac..produto



Select (CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_ICMSSaida ELSE PR_ICMSSaidaIva End) as IcmsSaida,(CASE WHEN PR_SubstituicaoTributaria = 'N' THEN PR_IcmPdv ELSE PR_ICMSPDVSaidaIva End) as IcmsPdv,PRB_CodigoBarras,PR_Referencia,PR_Descricao,PR_PrecoVenda1,EL_Estoque,PR_Classe,PR_Bloqueio,PR_SubstituicaoTributaria,LPR_Linha,LPR_Descricao From ProdutoLoja, Produtobarras, EstoqueLoja, LinhaProduto Where EL_Referencia=PR_Referencia and PR_Descricao Like 'FUR%'  and PR_Situacao not in('E') and PRB_Referencia = PR_Referencia and (Case When PR_LinhaProduto IS NULL Then  '990100' Else PR_LinhaProduto End) = LPR_Linha and PRB_TipoCodigo = 'D' Order By PR_CodigoFornecedor,PR_Descricao
Alter Database DMAC_LOJA Collate sql_latin1_general_cp1_ci_as



update ControleSistema set cts_loja = '271', cts_numeronf = 1, cts_numerone = 1, CTS_NumeroPedido = 1

