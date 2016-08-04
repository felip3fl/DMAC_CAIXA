exec SP_ALERTA_MovimentoCaixa '2015/06/22','2015/06/22','48','S',''
SELECT * FROM dmac48.dmac_loja.dbo.MOVIMENTOCAIXA WHERE MC_Valor = 660 AND  MC_DATA = '2015/06/22'
SELECT nf,* FROM dmac48.dmac_loja.dbo.nfcapa WHERE totalnota = 660 AND  dataemi = '2015/06/22'
SELECT nf,* FROM nfcapa WHERE totalnota = 660 AND  dataemi = '2015/06/22'


select MC_Documento,* from MovimentoCaixa where mc_data = '2015/06/13' and MC_Loja = '48' order by MC_Documento
select MC_Documento,* from dmac48.dmac_loja.dbo.MovimentoCaixa where mc_data = '2015/06/13' and MC_Loja = '48' order by MC_Documento

select sum(totalnota) from nfcapa where dataemi = '2015/06/22' and lojaorigem = '48' and tiponota = 'V'
select sum(totalnota) from dmac48.dmac_loja.dbo.nfcapa where dataemi = '2015/06/22' and lojaorigem = '48' and tiponota = 'V'

select * from nfcapa where DATAEMI = '2015/06/13' and LOJAORIGEM = '48' and nf = 10205
select serie, * from dmac48.dmac_loja.dbo.nfcapa where totalnota = 17 and DATAEMI = '2015/06/13' and LOJAORIGEM = '48' and nf = 10205


Update Loja set LO_Conexao='S' Where LO_Loja = '48'
exec SP_VDA_Conexao_Retaguarda '48'

update dmac48.dmac_loja.dbo.nfcapa set tiponota = 'V', DATAPROCESSO = '2015/06/22' where 
DATAEMI = '2015/06/13' and LOJAORIGEM = '48' and serie = 'CF' and nf = 10205

update dmac48.dmac_loja.dbo.nfitens set tiponota = 'V', DATAPROCESSO = '2015/06/22' where 
DATAEMI = '2015/06/13' and LOJAORIGEM = '48' and serie = 'CF' and nf = 10205

update dmac48.dmac_loja.dbo.movimentocaixa set mc_tiponota = 'V', mc_DATAPROCESSO = '2015/06/22', mc_protocolo = '551' where 
mc_DATA = '2015/06/13' and mc_LOJA= '48' and mc_serie = 'CF' and mc_documento = 10205



select * from capanfvenda where VC_NotaFiscal = '5496' and vc_lojaorigem = '48'
SELECT MC_DOCUMENTO, * FROM MOVIMENTOCAIXA WHERE MC_DOCUMENTO = 5496 and MC_Data = '2015/06/22' AND MC_Grupo = '20101' AND MC_Loja = '48'  ORDER BY MC_Documento
SELECT MC_DOCUMENTO, * FROM dmac48.dmac_loja.dbo.MOVIMENTOCAIXA WHERE MC_DOCUMENTO = 3936
SELECT MC_DOCUMENTO, * FROM MOVIMENTOCAIXA WHERE mc_valor = 23.9 and mc_loja = '48' and  MC_Data = '2015/06/22'
SELECT NF, * FROM dmac48.dmac_loja.dbo.NFCAPA WHERE DATAEMI = '2015/06/22' AND LOJAORIGEM = '48' AND SERIE = 'CF1'  AND NF = 3936 ORDER BY NF
SELECT NF, * FROM NFCAPA WHERE DATAEMI = '2015/06/22' AND LOJAORIGEM = '48' AND SERIE = 'CF1' and tiponota = 'V' AND NF = 3936 ORDER BY NF

select * from dmac48.dmac_loja.dbo.nfcapa where nf = 4887

