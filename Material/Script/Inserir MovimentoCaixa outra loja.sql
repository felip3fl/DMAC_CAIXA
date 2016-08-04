select * from MovimentoCaixa where mc_data = '2014/11/03'
delete MovimentoCaixa where mc_data = '2014/11/03'

sp_help MovimentoCaixa

insert into MovimentoCaixa(MC_NumeroECF, MC_CodigoOperador, MC_Loja, MC_Data, MC_Grupo, MC_SubGrupo, MC_Documento, MC_Serie, MC_Valor, MC_Banco, MC_Agencia, MC_ContaCorrente, MC_NumeroCheque, MC_BomPara, MC_Parcelas, MC_Remessa, MC_SituacaoEnvio, MC_ControleAVR, MC_DataBaixaAVR, MC_Protocolo, MC_NroCaixa, MC_GrupoAuxiliar, MC_Situacao, MC_Pedido, MC_DataProcesso, MC_TipoNota, MC_SequenciaTEF) 
select MC_NumeroECF, MC_CodigoOperador, MC_Loja, MC_Data, MC_Grupo, MC_SubGrupo, MC_Documento, MC_Serie, MC_Valor, MC_Banco, MC_Agencia, MC_ContaCorrente, MC_NumeroCheque, MC_BomPara, MC_Parcelas, MC_Remessa, MC_SituacaoEnvio, MC_ControleAVR, MC_DataBaixaAVR, '104', MC_NroCaixa, MC_GrupoAuxiliar, MC_Situacao, MC_Pedido, MC_DataProcesso, MC_TipoNota, MC_SequenciaTEF from [dmac28].[dmac_loja].[dbo].movimentocaixa where mc_data = '2014/11/03'


select * from ControleCaixa order by CTR_DataInicial desc
SELECT MC_GrupoAuxiliar,MO_Descricao,SUM(MC_Valor) as Valor FROM MOVIMENTOCAIXA,MODALIDADE WHERE Mo_GRUPO=MC_GrupoAuxiliar AND MC_GRUPOAUXILIAR LIKE '30%' and MC_DATA = 2014/11/04 GROUP BY MC_GrupoAuxiliar,MO_DESCRICAO order by MC_GrupoAuxiliar