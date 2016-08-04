USE [DMAC]
GO
/****** Object:  Trigger [dbo].[TR_Startata_Conexao_Loja]    Script Date: 25/04/2014 15:42:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER  TRIGGER [dbo].[TR_Startata_Conexao_Loja]

ON [dbo].[Loja]

FOR Update 
AS 

Declare @LojaPar char(05)
 
BEGIN

	If Update(lo_conexao) 
	   Begin
             if (Select top 1 lo_conexao from inserted) = 'S'
                Begin
                  Select @LojaPar=(Select top 1 LO_Loja from Inserted)
				  --print @lojapar
                  Exec SP_VDA_Conexao_Retaguarda @LojaPar
                End 
        End
END

/*

update loja set lo_conexao = 'S' where lo_loja = '28'
Exec SP_Est_Transferencia_destino                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          


*/