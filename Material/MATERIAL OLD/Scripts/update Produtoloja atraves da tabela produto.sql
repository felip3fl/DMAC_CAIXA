update produtoloja set pr_precovenda1 = demeo..produto.pr_precovenda1
from produtoloja, demeo..produto
where produtoloja.PR_referencia = demeo..produto.PR_Referencia collate sql_latin1_general_cp1_ci_as

select * from produtoloja