USE [DMAC_Loja]
GO
/****** Object:  Table [dbo].[FIN_Estado]    Script Date: 06/01/2016 11:37:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
drop table FIN_Estado
CREATE TABLE [dbo].[FIN_Estado](
	[UF_Estado] [nvarchar](2) NOT NULL,
	[UF_Nome] [nvarchar](25) NULL,
	[UF_Regiao] [int] NULL,
	[UF_ICMSInterno] [float] NULL,
	[UF_ICMSInterEstadual] [float] NULL,
	[UF_ICMSInterImport] [float] NULL,
	[UF_ICMSDifal] [float] NULL,
	[UF_ICMSDifalImportado] [float] NULL,
	[UF_FECP] [float] NULL,
	[UF_Participacao] [float] NULL,
PRIMARY KEY CLUSTERED 
(
	[UF_Estado] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'AC', N'Acre', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'AL', N'Alagoas', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'AM', N'Amazonas', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'AP', N'Amapá', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'BA', N'Bahia', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'CE', N'Ceará', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'DF', N'Distrito Federal', 3, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'ES', N'Espirito Santo', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'GO', N'Goiàs', 3, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'MA', N'Maranhão', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'MG', N'Minas Gerais', 2, 18, 12, 4, 6, 14, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'MS', N'Mato Grosso do Sul', 3, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'MT', N'Mato Grosso', 3, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'PA', N'Pará', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'PB', N'Paraiba', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'PE', N'Pernanbuco', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'PI', N'Piaui', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'PR', N'Paraná', 1, 18, 12, 4, 6, 14, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'RJ', N'Rio de Janeiro', 2, 19, 12, 4, 7, 15, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'RN', N'Rio Grande do Norte', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'RO', N'Rondônia', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'RR', N'Roraima', 4, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'RS', N'Rio Grande do Sul', 1, 17, 12, 4, 5, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'SC', N'Santa Catarina', 1, 17, 12, 4, 5, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'SE', N'Sergipe', 5, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'SP', N'São Paulo', 9, 17, 7, 4, 10, 13, 2, 40)
INSERT [dbo].[FIN_Estado] ([UF_Estado], [UF_Nome], [UF_Regiao], [UF_ICMSInterno], [UF_ICMSInterEstadual], [UF_ICMSInterImport], [UF_ICMSDifal], [UF_ICMSDifalImportado], [UF_FECP], [UF_Participacao]) VALUES (N'TO', N'Tocantins', 3, 17, 7, 4, 10, 13, 2, 40)
