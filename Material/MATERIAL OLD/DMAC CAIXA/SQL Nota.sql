select * from estoqueloja where el_referencia = '1165454'

--delete movimentocaixa where mc_data = '2010/07/28'

--delete nfitens where dataemi = '2010/07/28'

select nf,serie,tiponota,* from nfcapa where dataemi = '2010/07/28' order by numeroped
select nf,serie,tiponota,* from nfitens where dataemi = '2010/07/28'  order by numeroped
select mc_documento,mc_serie,* from movimentocaixa where mc_data = '2010/07/28'

select * from estoqueloja where el_referencia = '1165454'

--delete movimentocaixa  where mc_documento in (220,10649)
