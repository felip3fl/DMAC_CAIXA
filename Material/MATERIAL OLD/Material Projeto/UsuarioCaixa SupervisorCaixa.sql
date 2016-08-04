use dmac_loja

declare @tipoUsuario char(1)
declare @nome char(25)
declare @senha char(6)
declare @situacao char(1)

set @nome = 'Felipe'
set @senha = '0000'
set @tipoUsuario = 'S'		-- [O = Operador] [S = Supervisor]
set @situacao = 'S'			-- [A = Ativo] [S = Supervisor]

--select top 100 * from UsuarioCaixa where usu_nome = @nome

insert into UsuarioCaixa 
(USU_Codigo,usu_nome, usu_tipoUsuario, USU_Senha, USU_Situacao) values 
((select max(usu_codigo) from UsuarioCaixa)+1,@nome,@tipoUsuario,@senha,@situacao)

--Select top 100 * from UsuarioCaixa
--delete UsuarioCaixa where usu_nome = @nome

/*

sp_help UsuarioCaixa

*/