USE [DMAC_LOJA_BACKUP_2]
GO
/****** Object:  Table [dbo].[SAT_NF]    Script Date: 15/04/2016 15:16:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[SAT_NF](
	[snf_Sequencia] [int] IDENTITY(1,1) NOT NULL,
	[snf_Descricao] [char](18) NULL,
	[snf_Sinal] [char](1) NULL,
	[snf_Dados] [varchar](2000) NULL,
	[snf_pedido] [numeric](18, 0) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
