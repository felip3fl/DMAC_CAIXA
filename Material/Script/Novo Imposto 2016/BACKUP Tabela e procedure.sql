USE [DMAC_Loja]
GO
/****** Object:  Table [dbo].[NFE_ESTRUTURA]    Script Date: 29/12/2015 14:10:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE NFE_ESTRUTURA
CREATE TABLE [dbo].[NFE_ESTRUTURA](
	[ETR_Sequencia] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[ETR_Rotulo] [nvarchar](255) NULL,
	[ETR_Campo] [nvarchar](255) NULL,
	[ETR_Tabela_DE] [nvarchar](255) NULL,
	[ETR_Campo_DE] [nvarchar](255) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_prod]    Script Date: 29/12/2015 14:10:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
DROP TABLE NFe_prod
CREATE TABLE [dbo].[NFe_prod](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[H_nItem] [int] NULL,
	[I_cProd] [char](60) NULL,
	[I_cEAN] [char](15) NULL,
	[I_xProd] [char](120) NULL,
	[I_NCM] [char](10) NULL,
	[I_EXTIPI] [char](3) NULL,
	[I_CFOP] [char](4) NULL,
	[I_uCom] [char](6) NULL,
	[I_qCom] [float] NULL,
	[I_vUnCom] [float] NULL,
	[I_vProd] [float] NULL,
	[I_cEANTrib] [char](14) NULL,
	[I_uTrib] [char](6) NULL,
	[I_qTrib] [float] NULL,
	[I_vUnTrib] [float] NULL,
	[I_vFrete] [float] NULL,
	[I_vSeg] [float] NULL,
	[I_vDesc] [float] NULL,
	[I_vOutro] [float] NULL,
	[I_indTot] [char](1) NULL,
	[N_origICMS] [char](1) NULL,
	[N_CSTICMS] [char](2) NULL,
	[N_modBCICMS] [char](1) NULL,
	[N_vBCICMS] [float] NULL,
	[N_pRedBCICMS] [float] NULL,
	[N_pICMS] [float] NULL,
	[N_vICMS] [float] NULL,
	[N_modBCST] [char](1) NULL,
	[N_pMVAST] [float] NULL,
	[N_pRedBCST] [float] NULL,
	[N_vBCST] [float] NULL,
	[N_pICMSST] [float] NULL,
	[N_vICMSST] [float] NULL,
	[O_cIEnq] [char](5) NULL,
	[O_CNPJProd] [char](14) NULL,
	[O_cSelo] [char](60) NULL,
	[O_qSelo] [char](12) NULL,
	[O_cEnq] [char](3) NULL,
	[O_CSTIPI] [char](2) NULL,
	[O_vBCIPI] [float] NULL,
	[O_qUnid] [float] NULL,
	[O_vUnid] [float] NULL,
	[O_pIPI] [float] NULL,
	[O_vIPI] [float] NULL,
	[O_CSTIPINT] [char](2) NULL,
	[P_vBCII] [float] NULL,
	[P_vDespAdu] [float] NULL,
	[P_vII] [float] NULL,
	[P_vIOF] [float] NULL,
	[Q_CSTPIS] [char](2) NULL,
	[Q_vBCPIS] [float] NULL,
	[Q_pPIS] [float] NULL,
	[Q_qBCProdPIS] [float] NULL,
	[Q_vAliqProdPIS] [float] NULL,
	[Q_vPIS] [float] NULL,
	[R_vBCPISST] [float] NULL,
	[R_pPISST] [float] NULL,
	[R_qBCProdPISST] [float] NULL,
	[R_vAliqProdPISST] [float] NULL,
	[R_vPISST] [float] NULL,
	[S_CSTCOFINS] [char](2) NULL,
	[S_vBCCOFINS] [float] NULL,
	[S_pCOFINS] [float] NULL,
	[S_qBCProdCOFINS] [float] NULL,
	[S_vAliqProdCOFINS] [float] NULL,
	[S_vCOFINS] [float] NULL,
	[T_vBCCOFINSST] [float] NULL,
	[T_pCOFINSST] [float] NULL,
	[T_qBCProdCOFINSST] [float] NULL,
	[T_vAliqProdCOFINSST] [float] NULL,
	[T_vCOFINSST] [float] NULL,
	[U_vBCISSQN] [float] NULL,
	[U_vAliqISSQN] [float] NULL,
	[U_vISSQN] [float] NULL,
	[U_cMunFGISSQN] [char](7) NULL,
	[U_cListServ] [char](4) NULL,
	[U_cSitTrib] [char](1) NULL,
	[V_infAdProd] [char](500) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[NFe_total]    Script Date: 29/12/2015 14:10:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
DROP TABLE NFe_total
CREATE TABLE [dbo].[NFe_total](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[vBCICMS] [float] NULL,
	[vICMS] [float] NULL,
	[vBCST] [float] NULL,
	[vST] [float] NULL,
	[vProd] [float] NULL,
	[vFrete] [float] NULL,
	[vSeg] [float] NULL,
	[vDesc] [decimal](8, 2) NULL,
	[vII] [float] NULL,
	[vIPI] [float] NULL,
	[vCOFINS] [float] NULL,
	[vOutro] [float] NULL,
	[vNF] [float] NULL,
	[vServ] [float] NULL,
	[vBCISSQ] [float] NULL,
	[vISS] [float] NULL,
	[vPIS] [float] NULL,
	[vCOFINsISSQ] [float] NULL,
	[vRetPIS] [float] NULL,
	[vRetCOFINS] [float] NULL,
	[vRetCSLL] [float] NULL,
	[vBCIRRF] [float] NULL,
	[vIRRF] [float] NULL,
	[vBCRetPrev] [float] NULL,
	[vRetPrev] [float] NULL,
	[vVICMSDESON] [float] NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
SET IDENTITY_INSERT [dbo].[NFE_ESTRUTURA] ON 

INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(1 AS Numeric(18, 0)), N'IDE               ', N'[IDE]', N'Nfe_ide             ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(2 AS Numeric(18, 0)), N'IDE               ', N'    CUF', N'Nfe_ide             ', N'cUF               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(3 AS Numeric(18, 0)), N'IDE               ', N'    NATOP', N'Nfe_ide             ', N'natOp             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(4 AS Numeric(18, 0)), N'IDE               ', N'    CNF', N'Nfe_ide             ', N'nNF               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(5 AS Numeric(18, 0)), N'IDE               ', N'    INDPAG', N'Nfe_ide             ', N'indPag            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(6 AS Numeric(18, 0)), N'IDE               ', N'    MOD', N'Nfe_ide             ', N'mod               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(7 AS Numeric(18, 0)), N'IDE               ', N'    NNF', N'Nfe_ide             ', N'nNF               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(8 AS Numeric(18, 0)), N'IDE               ', N'    SERIE', N'Nfe_ide             ', N'serie             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(9 AS Numeric(18, 0)), N'IDE               ', N'    DHEMI', N'Nfe_ide             ', N'dEmi              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(10 AS Numeric(18, 0)), N'IDE               ', N'    DHSAIENT', N'Nfe_ide             ', N'dSaiEnt           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(11 AS Numeric(18, 0)), N'IDE               ', N'    TPNF', N'Nfe_ide             ', N'tpNF              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(12 AS Numeric(18, 0)), N'IDE', N'    IDDEST', N'Nfe_ide', N'IDDEST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(13 AS Numeric(18, 0)), N'IDE               ', N'    CMUNFG', N'Nfe_ide             ', N'cMunFG            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(14 AS Numeric(18, 0)), N'IDE               ', N'    TPIMP', N'Nfe_ide             ', N'tpImp             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(15 AS Numeric(18, 0)), N'IDE               ', N'    TPEMIS', N'Nfe_ide             ', N'tpEmis            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(16 AS Numeric(18, 0)), N'IDE', N'    INDFINAL', N'Nfe_ide', N'INDFINAL')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(17 AS Numeric(18, 0)), N'IDE', N'    INDPRES', N'Nfe_ide', N'INDPRES')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(18 AS Numeric(18, 0)), N'IDE               ', N'    FINNFE', N'Nfe_ide             ', N'finNFe            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(19 AS Numeric(18, 0)), N'IDE               ', N'    PROCEMI', N'Nfe_ide             ', N'procEmi           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(20 AS Numeric(18, 0)), N'IDE               ', N'    VERPROC', N'Nfe_ide             ', N'verProc           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(21 AS Numeric(18, 0)), N'DANFE', N'[DANFE]', N'Nfe_controle', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(22 AS Numeric(18, 0)), N'DANFE', N'    IMPRESSORA', N'Nfe_controle', N'danfe_IMPRESSORA')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(23 AS Numeric(18, 0)), N'DANFE', N'    RETORNARESP', N'Nfe_controle', N'danfe_RETORNARESP')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(24 AS Numeric(18, 0)), N'EMAIL', N'[EMAIL]', N'Nfe_controle', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(25 AS Numeric(18, 0)), N'EMAIL', N'    DESTINATARIO', N'Nfe_controle', N'email_DESTINATARIO')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(26 AS Numeric(18, 0)), N'EMAIL', N'    ASSUNTO', N'Nfe_controle', N'email_ASSUNTO')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(27 AS Numeric(18, 0)), N'EMAIL', N'    MENSAGEM', N'Nfe_controle', N'email_MENSAGEM')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(28 AS Numeric(18, 0)), N'EMAIL', N'    EMAILEMITENTE', N'Nfe_controle', N'email_EMAILEMITENTE')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(29 AS Numeric(18, 0)), N'EMAIL', N'    NOMEEMITENTE', N'Nfe_controle', N'email_NOMEEMITENTE')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(30 AS Numeric(18, 0)), N'EMAIL', N'    ANEXOPDF', N'Nfe_controle', N'email_ANEXOPDF')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(31 AS Numeric(18, 0)), N'EMAIL', N'    ANEXOXML', N'Nfe_controle', N'email_ANEXOXML')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(32 AS Numeric(18, 0)), N'EMAIL', N'    ANEXOPROTOCOLO', N'Nfe_controle', N'email_ANEXOPROTOCOLO')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(33 AS Numeric(18, 0)), N'EMAIL', N'    ANEXOADICIONAL', N'Nfe_controle', N'email_ANEXOADICIONAL')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(34 AS Numeric(18, 0)), N'EMAIL', N'    COMPACTADO', N'Nfe_controle', N'email_COMPACTADO')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(35 AS Numeric(18, 0)), N'EMAIL', N'    RETORNARESP', N'Nfe_controle', N'email_RETORNARESP')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(36 AS Numeric(18, 0)), N'NFREF', N'[NFREF]', N'Nfe_ide', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(37 AS Numeric(18, 0)), N'NFREF', N'    REFNFE', N'Nfe_ide', N'refNFE')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(38 AS Numeric(18, 0)), N'EMIT              ', N'[EMIT]', N'Nfe_emit            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(39 AS Numeric(18, 0)), N'EMIT              ', N'    CNPJ', N'Nfe_emit            ', N'CNPJ              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(40 AS Numeric(18, 0)), N'EMIT              ', N'    XNOME', N'Nfe_emit            ', N'xNome             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(41 AS Numeric(18, 0)), N'EMIT              ', N'    IE', N'Nfe_emit            ', N'IE                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(42 AS Numeric(18, 0)), N'EMIT              ', N'    CRT', N'Nfe_emit            ', N'CRT               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(43 AS Numeric(18, 0)), N'ENDEREMIT         ', N'[ENDEREMIT]', N'Nfe_emit            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(44 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    XLGR', N'Nfe_emit            ', N'xLgr              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(45 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    NRO', N'Nfe_emit            ', N'nro               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(46 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    XCPL', N'Nfe_emit            ', N'xCpl              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(47 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    XBAIRRO', N'Nfe_emit            ', N'xBairro           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(48 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    CMUN', N'Nfe_emit            ', N'cMun              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(49 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    XMUN', N'Nfe_emit            ', N'xMun              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(50 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    UF', N'Nfe_emit            ', N'UF                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(51 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    CEP', N'Nfe_emit            ', N'CEP               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(52 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    CPAIS', N'Nfe_emit            ', N'cPais             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(53 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    XPAIS', N'Nfe_emit            ', N'xPais             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(54 AS Numeric(18, 0)), N'ENDEREMIT         ', N'    FONE', N'Nfe_emit            ', N'fone              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(55 AS Numeric(18, 0)), N'DEST              ', N'[DEST]', N'Nfe_dest            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(56 AS Numeric(18, 0)), N'DEST              ', N'    CNPJ', N'Nfe_dest            ', N'CNPJ              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(57 AS Numeric(18, 0)), N'DEST              ', N'    CPF', N'Nfe_dest            ', N'CPF')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(58 AS Numeric(18, 0)), N'DEST              ', N'    XNOME', N'Nfe_dest            ', N'xNome             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(59 AS Numeric(18, 0)), N'DEST', N'    INDIEDEST', N'Nfe_dest', N'INDIEDEST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(60 AS Numeric(18, 0)), N'DEST              ', N'    IE', N'Nfe_dest            ', N'IE                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(61 AS Numeric(18, 0)), N'DEST              ', N'    Isuf', N'Nfe_dest            ', N'ISUF')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(62 AS Numeric(18, 0)), N'ENDERDEST         ', N'[ENDERDEST]', N'Nfe_dest            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(63 AS Numeric(18, 0)), N'ENDERDEST         ', N'    XLGR', N'Nfe_dest            ', N'xLgr              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(64 AS Numeric(18, 0)), N'ENDERDEST         ', N'    NRO', N'Nfe_dest            ', N'Nro               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(65 AS Numeric(18, 0)), N'ENDERDEST         ', N'    XCPL', N'Nfe_dest            ', N'xCpl              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(66 AS Numeric(18, 0)), N'ENDERDEST         ', N'    XBAIRRO', N'Nfe_dest            ', N'xBairro           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(67 AS Numeric(18, 0)), N'ENDERDEST         ', N'    CMUN', N'Nfe_dest            ', N'cMun              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(68 AS Numeric(18, 0)), N'ENDERDEST         ', N'    XMUN', N'Nfe_dest            ', N'xMun              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(69 AS Numeric(18, 0)), N'ENDERDEST         ', N'    UF', N'Nfe_dest            ', N'UF                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(70 AS Numeric(18, 0)), N'ENDERDEST         ', N'    CEP', N'Nfe_dest            ', N'CEP               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(71 AS Numeric(18, 0)), N'ENDERDEST         ', N'    CPAIS', N'Nfe_dest            ', N'cPais             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(72 AS Numeric(18, 0)), N'ENDERDEST         ', N'    XPAIS', N'Nfe_dest            ', N'xPais             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(73 AS Numeric(18, 0)), N'ENDERDEST         ', N'    FONE', N'Nfe_dest            ', N'fone              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(74 AS Numeric(18, 0)), N'TRANSP            ', N'[TRANSP]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(75 AS Numeric(18, 0)), N'TRANSP            ', N'    MODFRETE', N'Nfe_transp          ', N'modFrete          ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(76 AS Numeric(18, 0)), N'TRANSPORTA        ', N'[TRANSPORTA]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(77 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    CNPJ', N'Nfe_transp          ', N'CNPJ              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(78 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    XNOME', N'Nfe_transp          ', N'xNome             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(79 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    IE', N'Nfe_transp          ', N'IE                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(80 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    XENDER', N'Nfe_transp          ', N'xEnder            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(81 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    XMUN', N'Nfe_transp          ', N'xMun              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(82 AS Numeric(18, 0)), N'TRANSPORTA        ', N'    UF', N'Nfe_transp          ', N'UF                ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(83 AS Numeric(18, 0)), N'VEICTRANSP        ', N'[VEICTRANSP]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(84 AS Numeric(18, 0)), N'VEICTRANSP        ', N'    PLACA', N'Nfe_transp          ', N'placa             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(85 AS Numeric(18, 0)), N'VEICTRANSP        ', N'    UF', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(86 AS Numeric(18, 0)), N'VEICTRANSP        ', N'    RNTC', N'Nfe_transp          ', N'RNTC              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(87 AS Numeric(18, 0)), N'REBOQUE           ', N'[REBOQUE]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(88 AS Numeric(18, 0)), N'REBOQUE           ', N'    PLACA', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(89 AS Numeric(18, 0)), N'REBOQUE           ', N'    UF', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(90 AS Numeric(18, 0)), N'REBOQUE           ', N'    RNTC', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(91 AS Numeric(18, 0)), N'REBOQUE           ', N'[REBOQUE]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(92 AS Numeric(18, 0)), N'REBOQUE           ', N'    PLACA', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(93 AS Numeric(18, 0)), N'REBOQUE           ', N'    UF', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(94 AS Numeric(18, 0)), N'REBOQUE           ', N'    RNTC', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(95 AS Numeric(18, 0)), N'VOL               ', N'[VOL]', N'Nfe_transp          ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(96 AS Numeric(18, 0)), N'VOL               ', N'    NVOL', N'Nfe_transp          ', N'nVol              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(97 AS Numeric(18, 0)), N'VOL               ', N'    QVOL', N'Nfe_transp          ', N'qVol              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(98 AS Numeric(18, 0)), N'VOL               ', N'    ESP', N'Nfe_transp          ', N'esq               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(99 AS Numeric(18, 0)), N'VOL               ', N'    MARCA', N'Nfe_transp          ', N'marca             ')
GO
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(100 AS Numeric(18, 0)), N'VOL               ', N'    PESOL', N'Nfe_transp          ', N'pesoL             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(101 AS Numeric(18, 0)), N'VOL               ', N'    PESOB', N'Nfe_transp          ', N'pesoB             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(102 AS Numeric(18, 0)), N'ICMSTOT           ', N'[ICMSTOT]', N'Nfe_total           ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(103 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VBC', N'Nfe_total           ', N'vBCICMS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(104 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VICMS', N'Nfe_total           ', N'vICMS             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(105 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VBCST', N'Nfe_total           ', N'vBCST             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(106 AS Numeric(18, 0)), N'ICMSTOT', N'    VICMSDESON', N'Nfe_total', N'vVICMSDESON')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(107 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VST', N'Nfe_total           ', N'vST               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(108 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VPROD', N'Nfe_total           ', N'vProd             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(109 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VFRETE', N'Nfe_total           ', N'vFrete            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(110 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VSEG', N'Nfe_total           ', N'vSeg              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(111 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VDESC', N'Nfe_total           ', N'vDesc             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(112 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VII', N'Nfe_total           ', N'vII               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(113 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VIPI', N'Nfe_total           ', N'vIPI              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(114 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VPIS', N'Nfe_total           ', N'vPIS              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(115 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VCOFINS', N'Nfe_total           ', N'vCOFINS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(116 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VOUTRO', N'Nfe_total           ', N'vOutro            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(117 AS Numeric(18, 0)), N'ICMSTOT           ', N'    VNF', N'Nfe_total           ', N'vNF               ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(118 AS Numeric(18, 0)), N'INFADIC           ', N'[INFADIC]', N'Nfe_infAdic         ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(119 AS Numeric(18, 0)), N'INFADIC           ', N'    INFADFISCO', N'Nfe_infAdic         ', N'infAdFisco')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(120 AS Numeric(18, 0)), N'INFADIC           ', N'    INFCPL', N'Nfe_infAdic         ', N'infCpl            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(121 AS Numeric(18, 0)), N'OBSCONT           ', N'[OBSCONT]', N'Nfe_infAdic         ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(122 AS Numeric(18, 0)), N'OBSCONT           ', N'    XCAMPO', N'Nfe_infAdic         ', N'xCampoCont        ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(123 AS Numeric(18, 0)), N'OBSCONT           ', N'    XTEXTO', N'Nfe_infAdic         ', N'xTextoCont        ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(124 AS Numeric(18, 0)), N'FAT               ', N'[FAT]', N'Nfe_cobr            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(125 AS Numeric(18, 0)), N'FAT               ', N'    NFAT', N'Nfe_cobr            ', N'nFat              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(126 AS Numeric(18, 0)), N'FAT               ', N'    VORIG', N'Nfe_cobr            ', N'vOrig             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(127 AS Numeric(18, 0)), N'FAT               ', N'    VLIQ', N'Nfe_cobr            ', N'vLiq              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(128 AS Numeric(18, 0)), N'DUP               ', N'[DUP]', N'Nfe_cobr            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(129 AS Numeric(18, 0)), N'DUP               ', N'    NDUP', N'Nfe_cobr            ', N'nDup              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(130 AS Numeric(18, 0)), N'DUP               ', N'    DVENC', N'Nfe_cobr            ', N'dVend             ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(131 AS Numeric(18, 0)), N'DUP               ', N'    VDUP', N'Nfe_cobr            ', N'vDup              ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(132 AS Numeric(18, 0)), N'                  ', N'--', N'                    ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(133 AS Numeric(18, 0)), N'PROD              ', N'[PROD]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(134 AS Numeric(18, 0)), N'PROD              ', N'    CPROD', N'Nfe_prod            ', N'I_cProd           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(135 AS Numeric(18, 0)), N'PROD              ', N'    XPROD', N'Nfe_prod            ', N'I_xProd           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(136 AS Numeric(18, 0)), N'PROD              ', N'    NCM', N'Nfe_prod            ', N'I_NCM')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(137 AS Numeric(18, 0)), N'PROD              ', N'    CFOP', N'Nfe_prod            ', N'I_CFOP            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(138 AS Numeric(18, 0)), N'PROD              ', N'    UCOM', N'Nfe_prod            ', N'I_uCom            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(139 AS Numeric(18, 0)), N'PROD              ', N'    QCOM', N'Nfe_prod            ', N'I_qCom            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(140 AS Numeric(18, 0)), N'PROD              ', N'    VUNCOM', N'Nfe_prod            ', N'I_vUnCom          ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(141 AS Numeric(18, 0)), N'PROD              ', N'    VPROD', N'Nfe_prod            ', N'I_vProd           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(142 AS Numeric(18, 0)), N'PROD              ', N'    UTRIB', N'Nfe_prod            ', N'I_uTrib           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(143 AS Numeric(18, 0)), N'PROD              ', N'    QTRIB', N'Nfe_prod            ', N'I_qTrib           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(144 AS Numeric(18, 0)), N'PROD              ', N'    VUNTRIB', N'Nfe_prod            ', N'I_vUnTrib         ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(145 AS Numeric(18, 0)), N'PROD              ', N'    VFRETE', N'Nfe_prod            ', N'I_vFrete')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(146 AS Numeric(18, 0)), N'PROD              ', N'    VSEG', N'Nfe_prod            ', N'I_vSeg')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(147 AS Numeric(18, 0)), N'PROD              ', N'    VDESC', N'Nfe_prod            ', N'I_vDesc')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(148 AS Numeric(18, 0)), N'PROD              ', N'    INDTOT', N'Nfe_prod            ', N'I_indTot')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(149 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(150 AS Numeric(18, 0)), N'ICMS00            ', N'[ICMS00]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(151 AS Numeric(18, 0)), N'ICMS00            ', N'    CST', N'Nfe_prod            ', N'N_CSTICMS         ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(152 AS Numeric(18, 0)), N'ICMS00            ', N'    ORIG', N'Nfe_prod            ', N'N_origICMS        ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(153 AS Numeric(18, 0)), N'ICMS00            ', N'    MODBC', N'Nfe_prod            ', N'N_modBCICMS       ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(154 AS Numeric(18, 0)), N'ICMS00            ', N'    VBC', N'Nfe_prod            ', N'N_VBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(155 AS Numeric(18, 0)), N'ICMS00            ', N'    PICMS', N'Nfe_prod            ', N'N_pICMS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(156 AS Numeric(18, 0)), N'ICMS00            ', N'    VICMS', N'Nfe_prod            ', N'N_vICMS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(157 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(158 AS Numeric(18, 0)), N'ICMS10', N'[ICMS10]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(159 AS Numeric(18, 0)), N'ICMS10', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(160 AS Numeric(18, 0)), N'ICMS10', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(161 AS Numeric(18, 0)), N'ICMS10', N'    MODBC', N'Nfe_prod', N'N_modBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(162 AS Numeric(18, 0)), N'ICMS10', N'    VBC', N'Nfe_prod', N'N_VBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(163 AS Numeric(18, 0)), N'ICMS10', N'    PICMS', N'Nfe_prod', N'N_pICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(164 AS Numeric(18, 0)), N'ICMS10', N'    VICMS', N'Nfe_prod', N'N_vICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(165 AS Numeric(18, 0)), N'ICMS10', N'    MODBCST', N'Nfe_prod', N'n_modBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(166 AS Numeric(18, 0)), N'ICMS10', N'    VBCST', N'Nfe_prod', N'n_vBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(167 AS Numeric(18, 0)), N'ICMS10', N'    PICMSST', N'Nfe_prod', N'n_pICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(168 AS Numeric(18, 0)), N'ICMS10', N'    VICMSST', N'Nfe_prod', N'n_vICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(169 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(170 AS Numeric(18, 0)), N'ICMS20            ', N'[ICMS20]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(171 AS Numeric(18, 0)), N'ICMS20', N'    CST', N'Nfe_prod            ', N'N_CSTICMS         ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(172 AS Numeric(18, 0)), N'ICMS20            ', N'    ORIG', N'Nfe_prod            ', N'N_origICMS        ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(173 AS Numeric(18, 0)), N'ICMS20', N'    MODBC', N'Nfe_prod            ', N'N_modBCICMS       ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(174 AS Numeric(18, 0)), N'ICMS20', N'    PREDBC', N'Nfe_prod            ', N'ROUND(100-((N_vBCICMS*100)/I_vProd),2,0) ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(175 AS Numeric(18, 0)), N'ICMS20', N'    VBC', N'Nfe_prod            ', N'N_VBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(176 AS Numeric(18, 0)), N'ICMS20', N'    PICMS', N'Nfe_prod            ', N'N_pICMS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(177 AS Numeric(18, 0)), N'ICMS20', N'    VICMS', N'Nfe_prod            ', N'N_vICMS           ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(178 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(179 AS Numeric(18, 0)), N'ICMS30', N'[ICMS30]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(180 AS Numeric(18, 0)), N'ICMS30', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(181 AS Numeric(18, 0)), N'ICMS30', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(182 AS Numeric(18, 0)), N'ICMS30', N'    MODBCST', N'Nfe_prod', N'n_modBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(183 AS Numeric(18, 0)), N'ICMS30', N'    VBCST', N'Nfe_prod', N'n_vBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(184 AS Numeric(18, 0)), N'ICMS30', N'    PICMSST', N'Nfe_prod', N'n_pICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(185 AS Numeric(18, 0)), N'ICMS30', N'    VICMSST', N'Nfe_prod', N'n_vICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(186 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(187 AS Numeric(18, 0)), N'ICMS40', N'[ICMS40]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(188 AS Numeric(18, 0)), N'ICMS40', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(189 AS Numeric(18, 0)), N'ICMS40', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(190 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(191 AS Numeric(18, 0)), N'ICMS51', N'[ICMS51]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(192 AS Numeric(18, 0)), N'ICMS51', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(193 AS Numeric(18, 0)), N'ICMS51', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(194 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(195 AS Numeric(18, 0)), N'ICMS60            ', N'[ICMS60]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(196 AS Numeric(18, 0)), N'ICMS60', N'    CST', N'Nfe_prod            ', N'N_CSTICMS         ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(197 AS Numeric(18, 0)), N'ICMS60            ', N'    ORIG', N'Nfe_prod            ', N'N_origICMS        ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(198 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(199 AS Numeric(18, 0)), N'ICMS70', N'[ICMS70]', N'Nfe_prod', N'')
GO
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(200 AS Numeric(18, 0)), N'ICMS70', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(201 AS Numeric(18, 0)), N'ICMS70', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(202 AS Numeric(18, 0)), N'ICMS70', N'    MODBC', N'Nfe_prod', N'N_modBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(203 AS Numeric(18, 0)), N'ICMS70', N'    PREDBC', N'Nfe_prod', N'ROUND(100-((N_vBCICMS*100)/I_vProd),2,0)')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(204 AS Numeric(18, 0)), N'ICMS70', N'    VBC', N'Nfe_prod', N'N_VBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(205 AS Numeric(18, 0)), N'ICMS70', N'    PICMS', N'Nfe_prod', N'N_pICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(206 AS Numeric(18, 0)), N'ICMS70', N'    VICMS', N'Nfe_prod', N'N_vICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(207 AS Numeric(18, 0)), N'ICMS70', N'    MODBCST', N'Nfe_prod', N'n_modBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(208 AS Numeric(18, 0)), N'ICMS70', N'    VBCST', N'Nfe_prod', N'n_vBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(209 AS Numeric(18, 0)), N'ICMS70', N'    PICMSST', N'Nfe_prod', N'n_pICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(210 AS Numeric(18, 0)), N'ICMS70', N'    VICMSST', N'Nfe_prod', N'n_vICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(211 AS Numeric(18, 0)), N'IMPOSTO', N'[IMPOSTO]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(212 AS Numeric(18, 0)), N'ICMS90', N'[ICMS90]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(213 AS Numeric(18, 0)), N'ICMS90', N'    ORIG', N'Nfe_prod', N'N_origICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(214 AS Numeric(18, 0)), N'ICMS90', N'    CST', N'Nfe_prod', N'N_CSTICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(215 AS Numeric(18, 0)), N'ICMS90', N'    MODBC', N'Nfe_prod', N'N_modBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(216 AS Numeric(18, 0)), N'ICMS90', N'    VBC', N'Nfe_prod', N'N_VBCICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(217 AS Numeric(18, 0)), N'ICMS90', N'    PICMS', N'Nfe_prod', N'N_pICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(218 AS Numeric(18, 0)), N'ICMS90', N'    VICMS', N'Nfe_prod', N'N_vICMS')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(219 AS Numeric(18, 0)), N'ICMS90', N'    MODBCST', N'Nfe_prod', N'n_modBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(220 AS Numeric(18, 0)), N'ICMS90', N'    VBCST', N'Nfe_prod', N'n_vBCST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(221 AS Numeric(18, 0)), N'ICMS90', N'    PICMSST', N'Nfe_prod', N'n_pICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(222 AS Numeric(18, 0)), N'ICMS90', N'    VICMSST', N'Nfe_prod', N'n_vICMSST')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(223 AS Numeric(18, 0)), N'IPI', N'[IPI]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(224 AS Numeric(18, 0)), N'IPI', N'    CENQ', N'Nfe_prod', N'O_cEnq')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(225 AS Numeric(18, 0)), N'IPITRIB', N'[IPITRIB]', N'Nfe_prod', N'')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(226 AS Numeric(18, 0)), N'IPITRIB', N'    CST', N'Nfe_prod', N'O_CSTIPI')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(227 AS Numeric(18, 0)), N'IPITRIB', N'    VBC', N'Nfe_prod', N'O_vBCIPI')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(228 AS Numeric(18, 0)), N'IPITRIB', N'    PIPI', N'Nfe_prod', N'O_pIPI')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(229 AS Numeric(18, 0)), N'IPITRIB', N'    VIPI', N'Nfe_prod', N'O_vIPI')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(230 AS Numeric(18, 0)), N'PISALIQ           ', N'[PISALIQ]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(231 AS Numeric(18, 0)), N'PISALIQ           ', N'    CST', N'Nfe_prod            ', N'Q_CSTPIS          ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(232 AS Numeric(18, 0)), N'PISALIQ           ', N'    VBC', N'Nfe_prod            ', N'Q_vBCPIS          ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(233 AS Numeric(18, 0)), N'PISALIQ           ', N'    PPIS', N'Nfe_prod            ', N'Q_pPis            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(234 AS Numeric(18, 0)), N'PISALIQ           ', N'    VPIS', N'Nfe_prod            ', N'Q_vPIS            ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(235 AS Numeric(18, 0)), N'COFINSALIQ        ', N'[COFINSALIQ]', N'Nfe_prod            ', N'                  ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(236 AS Numeric(18, 0)), N'COFINSALIQ        ', N'    CST', N'Nfe_prod            ', N'S_CSTCOFINS       ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(237 AS Numeric(18, 0)), N'COFINSALIQ        ', N'    VBC', N'Nfe_prod            ', N'S_vBCCOFINS       ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(238 AS Numeric(18, 0)), N'COFINSALIQ        ', N'    PCOFINS', N'Nfe_prod            ', N'S_pCOFINS         ')
INSERT [dbo].[NFE_ESTRUTURA] ([ETR_Sequencia], [ETR_Rotulo], [ETR_Campo], [ETR_Tabela_DE], [ETR_Campo_DE]) VALUES (CAST(239 AS Numeric(18, 0)), N'COFINSALIQ        ', N'    VCOFINS', N'Nfe_prod            ', N'S_vCOFINS         ')
SET IDENTITY_INSERT [dbo].[NFE_ESTRUTURA] OFF
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8237      ', N'NE', NULL, 1, N'1160031                                                     ', N'               ', N'ESM ANG 4.1/2 GWS7-115 ET 3885E1 220V                                                                                   ', N'84672999  ', N'   ', N'5409', N'PC    ', 1, 215.658, 215.66, N'              ', N'PC    ', 1, 215.658, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 215.658, 0, 0, N'  ', 0, 0, 0, 0, N'01', 215.66, 1.65, 0, 0, 3.55839, 0, 0, 0, 0, 0, N'01', 215.66, 7.6, 0, 0, 16.39016, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8237      ', N'NE', NULL, 2, N'1160284                                                     ', N'               ', N'LIX OSC GSS23AE 1070 7E0  220V                                                                                          ', N'84672999  ', N'   ', N'5409', N'PC    ', 1, 246.7752, 246.78, N'              ', N'PC    ', 1, 246.7752, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 246.7752, 0, 0, N'  ', 0, 0, 0, 0, N'01', 246.78, 1.65, 0, 0, 4.07187, 0, 0, 0, 0, 0, N'01', 246.78, 7.6, 0, 0, 18.75528, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8238      ', N'NE', NULL, 1, N'1160472                                                     ', N'               ', N'FUR PAR GSR 7-14E 1447 0E0 220V                                                                                         ', N'84672992  ', N'   ', N'5405', N'PC    ', 1, 199, 199, N'              ', N'PC    ', 1, 199, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 199, 0, 0, N'  ', 0, 0, 0, 0, N'01', 199, 1.65, 0, 0, 3.2835, 0, 0, 0, 0, 0, N'01', 199, 7.6, 0, 0, 15.124, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8238      ', N'NE', NULL, 2, N'0600300                                                     ', N'               ', N'PTA PHIL 2 N°02 GRANEL DW2022B                                                                                          ', N'82079000  ', N'   ', N'5405', N'PC    ', 2, 3.3, 6.6, N'              ', N'PC    ', 2, 3.3, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 2, 3.3, 0, 0, N'  ', 0, 0, 0, 0, N'01', 6.6, 1.65, 0, 0, 0.1089, 0, 0, 0, 0, 0, N'01', 6.6, 7.6, 0, 0, 0.5016, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8238      ', N'NE', NULL, 3, N'1160335                                                     ', N'               ', N'BROCA MULTICONSTRU 596 055 8MMX120                                                                                      ', N'82075011  ', N'   ', N'5405', N'PC    ', 1, 17, 17, N'              ', N'PC    ', 1, 17, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 17, 0, 0, N'  ', 0, 0, 0, 0, N'01', 17, 1.65, 0, 0, 0.2805, 0, 0, 0, 0, 0, N'01', 17, 7.6, 0, 0, 1.292, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8238      ', N'NE', NULL, 4, N'1580457                                                     ', N'               ', N'BROCA AC 3P 06MM 5301060000                                                                                             ', N'82075011  ', N'   ', N'5405', N'PC    ', 1, 5.9, 5.9, N'              ', N'PC    ', 1, 5.9, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 5.9, 0, 0, N'  ', 0, 0, 0, 0, N'01', 5.9, 1.65, 0, 0, 0.09735, 0, 0, 0, 0, 0, N'01', 5.9, 7.6, 0, 0, 0.4484, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8239      ', N'NE', NULL, 1, N'7650001                                                     ', N'               ', N'KIT DE ACESSORIOS C/118PÇ F0018                                                                                         ', N'82079000  ', N'   ', N'5409', N'PC    ', 1, 72.4, 72.4, N'              ', N'PC    ', 1, 72.4, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 72.4, 0, 0, N'  ', 0, 0, 0, 0, N'01', 72.4, 1.65, 0, 0, 1.1946, 0, 0, 0, 0, 0, N'01', 72.4, 7.6, 0, 0, 5.5024, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8240      ', N'NE', NULL, 1, N'7150003                                                     ', N'               ', N'INVER SLD 150A IE-6150/I 220V                                                                                           ', N'85153900  ', N'   ', N'5102', N'PC    ', 1, 822, 822, N'              ', N'PC    ', 1, 822, 0, 0, 42, 0, N'1', N'0', N'20', N'2', 381.34, 1, 18, 68.64, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 822, 0, 0, N'  ', 0, 0, 0, 0, N'01', 780, 1.65, 0, 0, 12.87, 0, 0, 0, 0, 0, N'01', 780, 7.6, 0, 0, 59.28, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8241      ', N'NE', NULL, 1, N'1583851                                                     ', N'               ', N'MOTO ESMERIL 360W 110V ECCOFER                                                                                          ', N'84609090  ', N'   ', N'5152', N'PC    ', 1, 82.65, 82.65, N'              ', N'PC    ', 1, 82.65, 0, 0, 0, 0, N'1', N'0', N'20', N'2', 60.62, 1, 12, 7.27, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 82.65, 0, 0, N'  ', 0, 0, 0, 0, N'01', 82.65, 1.65, 0, 0, 1.363725, 0, 0, 0, 0, 0, N'01', 82.65, 7.6, 0, 0, 6.2814, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8242      ', N'NE', NULL, 1, N'0618119                                                     ', N'               ', N'SERRA RAP 14 2200W D28720 220V                                                                                          ', N'84615090  ', N'   ', N'5102', N'PC    ', 1, 699, 699, N'              ', N'PC    ', 1, 699, 0, 0, 39, 0, N'1', N'0', N'20', N'2', 322.67, 1, 18, 58.08, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 699, 0, 0, N'  ', 0, 0, 0, 0, N'01', 660, 1.65, 0, 0, 10.89, 0, 0, 0, 0, 0, N'01', 660, 7.6, 0, 0, 50.16, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8243      ', N'NE', NULL, 1, N'4330023                                                     ', N'               ', N'LAV AP HD 585 PROF 220V 1974185                                                                                         ', N'84243010  ', N'   ', N'5405', N'PC    ', 1, 1319, 1319, N'              ', N'PC    ', 1, 1319, 0, 0, 24, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 1319, 0, 0, N'  ', 0, 0, 0, 0, N'01', 1295, 1.65, 0, 0, 21.3675, 0, 0, 0, 0, 0, N'01', 1295, 7.6, 0, 0, 98.42, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8244      ', N'NE', NULL, 1, N'3730357                                                     ', N'               ', N'GUINCHO ELET 100/200KG 75583                                                                                            ', N'84251310  ', N'   ', N'5102', N'PC    ', 1, 539, 539, N'              ', N'PC    ', 1, 539, 0, 0, 0, 0, N'1', N'0', N'20', N'2', 395.3, 1, 12, 47.44, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 539, 0, 0, N'  ', 0, 0, 0, 0, N'01', 539, 1.65, 0, 0, 8.8935, 0, 0, 0, 0, 0, N'01', 539, 7.6, 0, 0, 40.964, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8245      ', N'NE', NULL, 1, N'6800984                                                     ', N'               ', N'ROCAD GAS 1,9CV DM 300 3F FS 160 0038                                                                                   ', N'84678900  ', N'   ', N'5405', N'PC    ', 1, 1542, 1542, N'              ', N'PC    ', 1, 1542, 0, 0, 52.17, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 1542, 0, 0, N'  ', 0, 0, 0, 0, N'01', 1489.83, 1.65, 0, 0, 24.582195, 0, 0, 0, 0, 0, N'01', 1489.83, 7.6, 0, 0, 113.22708, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8245      ', N'NE', NULL, 2, N'6801006                                                     ', N'               ', N'CJ CORTE TRIMCUT 41-2 40037102104                                                                                       ', N'84679900  ', N'   ', N'5405', N'PC    ', 1, 54, 54, N'              ', N'PC    ', 1, 54, 0, 0, 1.83, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 54, 0, 0, N'  ', 0, 0, 0, 0, N'01', 52.17, 1.65, 0, 0, 0.860805, 0, 0, 0, 0, 0, N'01', 52.17, 7.6, 0, 0, 3.96492, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8246      ', N'NE', NULL, 1, N'6000106                                                     ', N'               ', N'GER GAS 2500W MONO TG2800CXE P/E                                                                                        ', N'85013120  ', N'   ', N'5102', N'PC    ', 1, 2325, 2325, N'              ', N'PC    ', 1, 2325, 0, 0, 125, 0, N'1', N'0', N'0 ', N'2', 2200, 0, 12, 264, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 2325, 0, 0, N'  ', 0, 0, 0, 0, N'01', 2200, 1.65, 0, 0, 36.3, 0, 0, 0, 0, 0, N'01', 2200, 7.6, 0, 0, 167.2, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8247      ', N'NE', NULL, 1, N'0520028                                                     ', N'               ', N'APARADOR FIO NYLON M1000 220V PLUS                                                                                      ', N'84672999  ', N'   ', N'5405', N'PC    ', 1, 222, 222, N'              ', N'PC    ', 1, 222, 0, 0, 22, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 222, 0, 0, N'  ', 0, 0, 0, 0, N'01', 200, 1.65, 0, 0, 3.3, 0, 0, 0, 0, 0, N'01', 200, 7.6, 0, 0, 15.2, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8248      ', N'NE', NULL, 1, N'1780541                                                     ', N'               ', N'APARADOR FIO NYLON 700W 110V GAM700                                                                                     ', N'84672999  ', N'   ', N'5405', N'PC    ', 1, 129, 129, N'              ', N'PC    ', 1, 129, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 129, 0, 0, N'  ', 0, 0, 0, 0, N'01', 129, 1.65, 0, 0, 2.1285, 0, 0, 0, 0, 0, N'01', 129, 7.6, 0, 0, 9.804, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_prod] ([eLoja], [eNF], [eSerie], [Situacao], [H_nItem], [I_cProd], [I_cEAN], [I_xProd], [I_NCM], [I_EXTIPI], [I_CFOP], [I_uCom], [I_qCom], [I_vUnCom], [I_vProd], [I_cEANTrib], [I_uTrib], [I_qTrib], [I_vUnTrib], [I_vFrete], [I_vSeg], [I_vDesc], [I_vOutro], [I_indTot], [N_origICMS], [N_CSTICMS], [N_modBCICMS], [N_vBCICMS], [N_pRedBCICMS], [N_pICMS], [N_vICMS], [N_modBCST], [N_pMVAST], [N_pRedBCST], [N_vBCST], [N_pICMSST], [N_vICMSST], [O_cIEnq], [O_CNPJProd], [O_cSelo], [O_qSelo], [O_cEnq], [O_CSTIPI], [O_vBCIPI], [O_qUnid], [O_vUnid], [O_pIPI], [O_vIPI], [O_CSTIPINT], [P_vBCII], [P_vDespAdu], [P_vII], [P_vIOF], [Q_CSTPIS], [Q_vBCPIS], [Q_pPIS], [Q_qBCProdPIS], [Q_vAliqProdPIS], [Q_vPIS], [R_vBCPISST], [R_pPISST], [R_qBCProdPISST], [R_vAliqProdPISST], [R_vPISST], [S_CSTCOFINS], [S_vBCCOFINS], [S_pCOFINS], [S_qBCProdCOFINS], [S_vAliqProdCOFINS], [S_vCOFINS], [T_vBCCOFINSST], [T_pCOFINSST], [T_qBCProdCOFINSST], [T_vAliqProdCOFINSST], [T_vCOFINSST], [U_vBCISSQN], [U_vAliqISSQN], [U_vISSQN], [U_cMunFGISSQN], [U_cListServ], [U_cSitTrib], [V_infAdProd]) VALUES (N'271  ', N'8248      ', N'NE', NULL, 2, N'4390266                                                     ', N'               ', N'REFIL NYLON 1,8MM 100GRS 216-3                                                                                          ', N'39169010  ', N'   ', N'5405', N'PC    ', 1, 19, 19, N'              ', N'PC    ', 1, 19, 0, 0, 0, 0, N'1', N'0', N'60', N'2', 0, 0, 0, 0, N'0', 0, 0, 0, 0, 0, N'     ', N'              ', N'                                                            ', N'            ', N'999', N'50', 0, 1, 19, 0, 0, N'  ', 0, 0, 0, 0, N'01', 19, 1.65, 0, 0, 0.3135, 0, 0, 0, 0, 0, N'01', 19, 7.6, 0, 0, 1.444, 0, 0, 0, 0, 0, 0, 0, 0, N'       ', N'    ', N' ', N'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    ')
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8237      ', N'NE', NULL, 0, 0, 0, 0, 462.44, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 35.14544, 0, 462.44, 0, 0, 0, 7.63026, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8238      ', N'NE', NULL, 0, 0, 0, 0, 228.5, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 17.366, 0, 228.5, 0, 0, 0, 3.77025, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8239      ', N'NE', NULL, 0, 0, 0, 0, 72.4, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 5.5024, 0, 72.4, 0, 0, 0, 1.1946, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8240      ', N'NE', NULL, 381.34, 68.64, 0, 0, 822, 0, 0, CAST(42.00 AS Decimal(8, 2)), 0, 0, 59.28, 0, 780, 0, 0, 0, 12.87, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8241      ', N'NE', NULL, 60.62, 7.27, 0, 0, 82.65, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 6.2814, 0, 82.65, 0, 0, 0, 1.363725, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8242      ', N'NE', NULL, 322.67, 58.08, 0, 0, 699, 0, 0, CAST(39.00 AS Decimal(8, 2)), 0, 0, 50.16, 0, 660, 0, 0, 0, 10.89, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8243      ', N'NE', NULL, 0, 0, 0, 0, 1319, 0, 0, CAST(24.00 AS Decimal(8, 2)), 0, 0, 98.42, 0, 1295, 0, 0, 0, 21.3675, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8244      ', N'NE', NULL, 395.3, 47.44, 0, 0, 539, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 40.964, 0, 539, 0, 0, 0, 8.8935, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8245      ', N'NE', NULL, 0, 0, 0, 0, 1596, 0, 0, CAST(54.00 AS Decimal(8, 2)), 0, 0, 117.192, 0, 1542, 0, 0, 0, 25.443, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8246      ', N'NE', NULL, 2200, 264, 0, 0, 2325, 0, 0, CAST(125.00 AS Decimal(8, 2)), 0, 0, 167.2, 0, 2200, 0, 0, 0, 36.3, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8247      ', N'NE', NULL, 0, 0, 0, 0, 222, 0, 0, CAST(22.00 AS Decimal(8, 2)), 0, 0, 15.2, 0, 200, 0, 0, 0, 3.3, 0, 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[NFe_total] ([eLoja], [eNF], [eSerie], [Situacao], [vBCICMS], [vICMS], [vBCST], [vST], [vProd], [vFrete], [vSeg], [vDesc], [vII], [vIPI], [vCOFINS], [vOutro], [vNF], [vServ], [vBCISSQ], [vISS], [vPIS], [vCOFINsISSQ], [vRetPIS], [vRetCOFINS], [vRetCSLL], [vBCIRRF], [vIRRF], [vBCRetPrev], [vRetPrev], [vVICMSDESON]) VALUES (N'271  ', N'8248      ', N'NE', NULL, 0, 0, 0, 0, 148, 0, 0, CAST(0.00 AS Decimal(8, 2)), 0, 0, 11.248, 0, 148, 0, 0, 0, 2.442, 0, 0, 0, 0, 0, 0, 0, 0, 0)
/****** Object:  StoredProcedure [dbo].[SP_VDA_Cria_NFe]    Script Date: 29/12/2015 14:10:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
exec SP_VDA_Cria_NFe 'CD','3590','NE',''

SELECT nf,* FROM nfcapa where dataemi = '2015/09/21' nf = 3590

select * from fin_cliente  where ce_codigocliente = 192

select * from nfe_controle

DROP TABLE NFE_ESTRUTURA


*/

--exec SP_VDA_Cria_NFe 'CD','3661','NE',''

ALTER PROCEDURE [dbo].[SP_VDA_Cria_NFe]

	@Loja		Char(5),
	@NF		    Numeric,
	@Serie		Char(2),
    @Carimbo    varchar(MAX)

AS

	DECLARE	@SQL        	char(4000),
			@CondPagto		Char(2),
			@CondPagtoNF	Char(2),
			@Parcelas       Char(2),
			@NroNF_NFe		Char(10),
			@Referencia		Char(7),
			@UFCliente		Char(2),
			@IDDEST			char(1),
			@finNFe			char(1),
            @CEPCliente     Char(8),
            @NomeServidor   char(40),
            @Cliente        char(6),
			@ClienteT       char(6),
			@IE				char(13),
			@Pessoa         char(1),
			@TipoEmissao    Char(1),
			@QtdeVolume     float,
			@TotalFrete     numeric,
			@PercFrete		float,
			@DiferencaFrete float,
			@Item			numeric,
			@tiponota		char(4),
			@Operacao		char(60),
			@cfop			numeric(18,0),
            @Hora           char(12),
            @Chave          char(8),
            @UFLoja         char(2),
			@EntradaSaida   char(1)


                 
BEGIN

	exec sp_delete_nfe @loja, @nf, @Serie
	delete NFE_NFLojas 
	 where NFL_NroNFE = @nf
	
	Select @tiponota = (Select top 1 tiponota 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie)

	
	-- -- ACERTOS NFCAPA -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	update NfItens set 
		   VALORICMS = round(((BASEICMS * ICMSAplicado) / 100),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'
	
	update NfCapa set 
		   vlrICMS = round((select SUM(VALORICMS) as total 
						      from NfItens 
						     where nf = @nf 
						       and serie = @Serie),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'       

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


	print 'OK 1'
	Select @CondPagtoNF = (Select TOP 1 CondPag 
	                         from NFcapa 
	                        where LojaOrigem = @Loja 
	                          And NF = @NF 
	                          And Serie = @Serie)

	Select @Parcelas = (Select TOP 1 CP_parcelas 
	                      from CondicaoPagamento 
	                     Where CP_Codigo = @CondPagtoNF)
	                     
	SELECT @cfop = (Select TOP 1 CODOPER 
	                        from NFcapa 
	                       where LojaOrigem = @Loja 
	                         And NF = @NF 
	                         And Serie = @Serie)
	
	--Update ControleSup set CS_NumeroNFe = (CS_NumeroNFe + 1)
	
	Select @NroNF_NFe = @NF
	
	print 'OK 2'
	Select @UFCliente = (select ce_Estado 
	                       from NFCapa,FIN_cliente 
	                      where ce_codigocliente = cliente 
	                        and lojaorigem = @Loja 
	                        and NF = @Nf 
	                        and Serie = @serie)

	Select @Pessoa = (Select CE_TipoPessoa 
	                    from NFCapa,FIN_cliente 
	                   where ce_codigocliente = cliente 
	                     and lojaorigem = @Loja 
	                     and NF = @Nf 
	                     and Serie = @serie)
               

    Select @CEPCliente = (Select replicate('0',8 - len(CE_Cep)) + CE_Cep 
                            from NFCapa,FIN_cliente 
                           where ce_codigocliente = cliente 
                             and lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)

	print 'OK 3'
    Select @QtdeVolume = (Select sum(qtde) 
                            from nfItens 
                           where LojaOrigem = @Loja 
                             And NF = @NF 
                             And Serie = @Serie)

	--select @EntradaSaida = (Select top 1 substring(codoper,1,1) from nfcapa where LojaOrigem = @Loja And NF = @NF And Serie = @Serie)
	
	print 'OK 3-1'	
	Select @Cliente = (Select top 1 cliente 
	                     from nfcapa 
	                    where LojaOrigem = @Loja 
	                      And NF = @NF 
	                      And Serie = @Serie 
	                      and tiponota <> 'T')
	
	print 'OK 3-2'
	Select @ClienteT = (Select top 1 lojat 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie 
	                       and TIPONOTA = 'T')
	
	print 'OK 3-3'	
    Select @TotalFrete = (Select fretecobr 
                            from NFCapa 
                           where lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)
    
    print 'OK 3-4'    
    Select @PercFrete = (Select ((fretecobr * 100)/ vlrmercadoria) 
                           from NFCapa 
                          where lojaorigem = @Loja 
                            and NF = @Nf 
                            and Serie = @serie)
	print 'OK 3-5'	
	select @DiferencaFrete = (select ( @TotalFrete - (sum(((vltotitem - desconto) * @PercFrete) / 100))) 
	                            from NFitens
		                       where lojaorigem = @Loja 
		                         and NF = @Nf 
		                         and Serie = @serie)
	print 'OK 3-6'	                         
	Select @Item = (select top 1 Item 
	                  from nfitens 
	                 where lojaorigem = @Loja 
	                   and NF = @Nf 
	                   and Serie = @serie 
	                 order by Item)
    
    print 'OK 3-7'
    Select @UFLoja = (select distinct substring(convert(nvarchar(9),lo_codigoMunicipio),1,2)
                        from Loja,nfcapa 
                       where lojaorigem = @Loja 
                         and NF = @Nf 
                         and Serie = @serie 
                         and lojaorigem = lo_loja)
                         
    Select @Hora = CONVERT(varchar,GETDATE(),114)
    Select @Hora = replace(@Hora,':','')
    Select @Chave = substring(@hora,5,2) + substring(@hora,3,2) + substring(@hora,1,2) + substring(@hora,8,2)
      

-- SELECT Name + REPLICATE('*', 20 - LEN(Name)) FROM Employee
--	update nfcapa set fonecli = replace(fonecli,'-','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set fonecli = replace(fonecli,' ','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'.','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'(','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,')','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set cepcli = ' ' where LojaOrigem = @Loja And NF = @NF And Serie = @Serie And len(cepcli)<7
	print 'OK 4'
	
	Update nfitens set 
	       CSTICMS = 60 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'S' 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF 
	   and @tiponota <> 'S'
	
	print ('Update nfitens set CSTICMS = 60')

	print 'OK 5'
	
	Update nfitens set 
	       CSTICMS = 20 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms > 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'
	print ('Update nfitens set CSTICMS = 20')

	print 'OK 6'
	
	Update nfitens set 
	       CSTICMS = 00 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms = 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'
	   
	select @IDDEST = '1'

	if @Tiponota NOT IN ('E') 
		BEGIN

	IF @UFCliente = 'SP'
	   BEGIN
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
					       CFOP = 5102 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
			           and pr_substituicaoTributaria = 'N' 
			           and LojaOrigem = @Loja 
			           AND Serie = @Serie 
			           AND NF = @NF
			           and @tiponota <> 'S'
			           print ('Update nfitens set CFOP = 5102')
			           
				end
			IF @pessoa = 'F' or @pessoa = 'U' or @pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
						   CFOP = 5405 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 5405')
					
				end
		  --print @tiponot
		END

	IF @UFCliente <> 'SP'
		BEGIN
			set @IDDEST = '2'
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6404 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6404')
				end 
				
			IF @pessoa = 'F' or @pessoa = 'U' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6108 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6108')
				end
				
			IF @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6102 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6102')
				end
		END

	END

	IF rtrim(ltrim(@tiponota)) = 'T'
		Begin
			set @IDDEST = '1'
			Update nfitens set 
			       CFOP = 5409 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'S' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
			print ('Update nfitens set CFOP = 5409 transferencia ST')

			Update nfitens set 
			       CFOP = 5152 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'N' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
		end
	
			
		--update NFItens set 
		--       CFOP = (select codoper 
		--                 from NFCapa 
		--                where LojaOrigem = @Loja 
		--                  AND Serie = @Serie 
		--                  AND NF = @NF )	
		--  from NFItens 
		-- where LojaOrigem = @Loja 
		--   AND Serie = @Serie 
		--   AND NF = @NF
		
	Update nfcapa set codoper = (select top 1 CFOP from nfitens where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF) 
	from nfcapa where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF			
				
	print 'NF'
	print @CondPagtoNF
	
	If @CondPagtoNF = 1
	   Begin
		Select @CondPagto = 0
	   End
	   
	If @CondPagtoNF = 3
	   Begin
		Select @CondPagto = 2
	   End
	   
	If @CondPagtoNF between 4 and 199 
	   Begin
		Select @CondPagto = 1
	   End
	   
	If @CondPagtoNF = 2 or @CondPagtoNF >= 200 
           Begin		
                Select @CondPagto = 2
           End

	select @Operacao = (select top 1 cn_descricaooperacao 
	                      from codigooperacaonovo, NFCapa 
	                     where codoper = cn_codigooperacaonovo 
	                       and LojaOrigem = @Loja 
	                       AND Serie = @Serie 
	                       AND NF = @NF)
	
	if LTrim(Rtrim(@Operacao)) = ''	   
	Begin
		Select @Operacao = 'Venda.'
	End
	  
	/*
	FINNFE
	1 – NF-e normal
	2 – NF-e complementar
	3 – NF-e de ajuste
	4 – Devolução de mercadoria
	*/

	SET @finNFe = '1'
	if  @Tiponota <> 'E' 
	select @entradaSaida =  '1'

	if @cfop = '5202' or @cfop = '5411' or @cfop = '5553' or @cfop = '5909'  or @cfop = '6202' or @cfop = '6411' or @cfop = '6913' 
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end
	
	if @cfop = '1202' or @cfop = '1411' or @cfop = '2202' 
	begin
		select @entradaSaida = '0'
		select @finNFe = '4'
	end	
	
	if @cfop = '5918'  
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end	


	set @IE = (select top 1 ce_inscricaoEstadual 
	             from FIN_Cliente, NFCapa 
	            where cliente = CE_CodigoCliente 
	              and NF = @NF 
	              and serie = @Serie 
	              and LOJAORIGEM = @Loja)

	if @Pessoa = 'F' or @Pessoa = 'U' 
	begin
		set @pessoa = '9'
		set @IE = ''
	end 
	
	if @Pessoa = 'J' or @Pessoa = 'O' 
	begin
		set @pessoa = '1'	
		
		if @IE = 'ISENTO'
		begin
			set @pessoa = '9'	
			set @IE = ''	
		end 
		
	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	Select @SQL = 'INSERT INTO NFe_ide (eLoja,eNF,eSerie,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,
	tpNF,cMunFG,tpImp,tpEmis,cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust,IDDEST,INDFINAL,INDPRES,refNFe) Select LojaOrigem AS eLoja,nf AS eNF,
	Serie as eSerie,'+''''+ LTrim(RTrim(@UFLoja))+'''' +' AS cUF,'+ LTrim(Rtrim(@NF)) +' As cNF,
	' + '''' + LTrim(RTrim(@Operacao)) + '''' + ' as natop,
	'+ @CondPagto +' As indPag,'+ '''55''' +' AS mod,'+'''1'''+' As serie,
	' + ''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' AS nNF,dataemi As dEmi,DataEmi As dSaiEnt,
	Hora as hSaiEnt,' + '' + @entradaSaida + '' + ' As tpNF,LO_CodigoMunicipio As cMunFG,' + '''1''' + ' As tpImp,
	' + '''1''' + ' As tpEmis,'+ ''' ''' +' As cDV,' +'''2'''+ ' As tpAmb,' + '''' + @finNFe + '''' + ' As finNFe,
	' + '''3''' + ' As procEmi,'+ '''2.0.0''' +' As verProc,getdate() As dhCont,
	' + '''Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz''' + ' As xJust, 
	''' + @IDDEST + ''' as IDDEST,''1'' as INDFINAL,''1'' as INDPRES, ChaveNFeDevolucao
	FROM NFCapa (NOLOCK), Loja (NOLOCK) 
	WHERE LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)
	
--select * from NFe_ide where eNF = '2049'


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		'''' as email_DESTINATARIO,'''' as email_ASSUNTO,'''' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'
	ELSE
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		ce_email as email_DESTINATARIO,''Nota Fiscal Eletrônica ' + LTrim(Rtrim(@NF)) + ' - '' + LO_NomeFantasia as email_ASSUNTO,''Olá '' + ltrim(rtrim(CE_Razao)) + '' 
		Você está recebendo uma cópia da DANFE e o arquivo XML'' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), fin_Cliente, Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and cliente = CE_CodigoCliente and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	
	

	Select @SQL = 'INSERT INTO NFe_emit(eLoja,eNF,eSerie,CNPJ,xNome,xFant,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,
	CEP,cPais,xPais,fone,IE,IEST,IM,CNAE,CRT) SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
	LO_CGC As CNPJ,LO_razao As xNome,LO_NomeFantasia As xFant,
	Lo_Endereco As xLgr,Lo_numero As nro,'''' As xCpl,LO_Bairro As xBairro,
	LO_CodigoMunicipio As cMun,LO_Municipio As xMun,LO_UF As UF,LO_CEP As CEP, 
	'+ '''1058''' +' As cPais, '+'''Brasil'''+' As xPais,LO_DDD + LO_Telefone As fone,
	LO_InscricaoEstadual As IE,'+''' '''+' As IEST,'+''' '''+' As IM,'+''' '''+' As CNAE, 
	'+'''3'''+' As CRT
	FROM Loja (NOLOCK), NFCapa (NOLOCK) WHERE LojaOrigem = LO_loja And 
	LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+ '''' + @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print (@SQL)
	Exec (@SQL)

	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
		(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
		lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
		lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
		lo_cep as CEP,
		''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
		lo_inscricaoEstadual as IE,
		'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
		FROM loja (nolock)
		WHERE lo_loja = '+''''+ @ClienteT +''''
	else
	--IF rtrim(ltrim(@tiponota)) = 'E'
	--	Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
	--	xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
	--	(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
	--	lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
	--	lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
	--	lo_cep as CEP,
	--	''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
	--	lo_inscricaoEstadual as IE,
	--	'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
	--	FROM loja (nolock)
	--	WHERE lo_loja = '+''''+ @Loja +''''
	--ELSE 
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,CPF,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email, INDIEDEST)SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
		(Case When len(CE_CGC) = 14 Then CE_CGC else '+''' '''+' end),
		(Case When len(CE_CGC) = 11 Then CE_CGC else '+''' '''+' end),
		CE_Razao As xNome,CE_Endereco As xLgr,CE_numero As nro,CE_Complemento As xCpl,
		CE_bairro As xBairro,CE_CodigoMunicipio As cMun,CE_Municipio As xMun,CE_Estado As UF,
		'+''''+ LTrim(Rtrim(@CEPCliente)) +''''+' as CEP,
		' + '''1058''' + ' As cPais,'+'''Brasil'''+' AS xPais,CE_telefone As fone,
		''' + @IE + ''' as IE,
		CE_InscricaoEstadualSuframa As ISUF,CE_email as Email, ''' + @pessoa +  ''' as INDIEDEST
		FROM NFCapa (NOLOCK),fin_Cliente (nolock)
		WHERE cliente = CE_CodigoCliente AND
		LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+
		' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF));		
	


	--Print @SQL-- select ce_cgc,* from fin_cliente where ce_codigocliente = 60046
	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--select * from nfe_estrutura
	Select @SQL = 'INSERT INTO NFe_prod (eLoja,eNF,eSerie,H_nItem,I_cProd,I_cEAN,I_xProd,I_NCM,I_EXTIPI,I_CFOP,
	I_uCom,I_qCom,I_vUnCom,I_vProd,I_cEANTrib,I_uTrib,I_qTrib,I_vUnTrib,I_vFrete,I_vSeg,I_vDesc,I_vOutro,
	I_indTot,N_origICMS,N_CSTICMS,N_modBCICMS,N_vBCICMS,N_pRedBCICMS,N_pICMS,N_vICMS,N_modBCST,N_pMVAST,
	N_pRedBCST,N_vBCST,N_pICMSST,N_vICMSST,O_cIEnq,O_CNPJProd,O_cSelo,O_qSelo,O_cEnq,O_CSTIPI,
	O_vBCIPI,O_qUnid,O_vUnid,O_pIPI,O_vIPI,O_CSTIPINT,P_vBCII,P_vDespAdu,P_vII,P_vIOF,Q_CSTPIS,
	Q_vBCPIS,Q_pPIS,Q_qBCProdPIS,Q_vAliqProdPIS,Q_vPIS,R_vBCPISST,R_pPISST,R_qBCProdPISST,
	R_vAliqProdPISST,R_vPISST,S_CSTCOFINS,S_vBCCOFINS,S_pCOFINS,S_qBCProdCOFINS,S_vAliqProdCOFINS,
	S_vCOFINS,T_vBCCOFINSST,T_pCOFINSST,T_qBCProdCOFINSST,T_vAliqProdCOFINSST,T_vCOFINSST,
	U_vBCISSQN,U_vAliqISSQN,U_vISSQN,U_cMunFGISSQN,U_cListServ,U_cSitTrib,V_infAdProd) 
	SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,ITEM As H_nItem,Referencia As I_cProd,
	'+''' '''+' As I_cEAN,PR_Descricao As I_xProd,PR_ClasseFiscal As I_NCM,'+''' '''+' As I_EXTIPI,
	CFOP As I_CFOP,PR_Unidade As I_uCom,QTDE As I_qCom,VLUnit As I_vUnCom,
	VLTotItem As I_vProd,'+''' '''+' As I_cEANTrib,PR_UNIDADE AS I_uTrib,QTDE aS I_qTrib,
	VLUnit as I_vUnTrib,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ((((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) end),

	'+'''0'''+' as I_vSeg,desconto as I_vDesc,'+'''0'''+' as I_vOutro, 
	'+'''1'''+' I_indTot,'+ '''0''' +' as N_origICMS,CSTICMS as N_CSTICMS,'+ '''2''' +' as N_modBCICMS,
	baseicms as N_vBCICMS,PR_codigoReducaoICMS as N_pRedBCICMS,ICMSAplicado as N_pICMS,
	ValorICMS as N_vICMS,'+'''0'''+' as N_modBCST,'+'''0'''+' as N_pMVAST,'+'''0'''+' as N_pRedBCST,
	'+'''0'''+' as N_vBCST,'+'''0'''+' as N_pICMSST,'+'''0'''+' as N_vICMSST,
	'+''' '''+' as O_cIEnq,'+''' '''+' as O_CNPJProd,'+''' '''+' as O_cSelo,'+''' '''+' as O_qSelo,
	'+'''999'''+' as O_cEnq,'+'''50'''+' as O_CSTIPI, baseIPI as O_vBCIPI, qtde as O_qUnid,
	vlUnit as O_vUnid, aliqIPI as O_pIPI, vlIpi as O_vIPI,'+''' '''+' as O_CSTIPINT,
	'+'''0'''+' as P_vBCII,'+'''0'''+' as P_vDespAdu,'+'''0'''+' as P_vII,
	'+'''0'''+' as P_vIOF,'+'''01'''+' as Q_CSTPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end), 

	'+'''1.65'''+' as Q_pPIS,'+'''0'''+' as Q_qBCProdPIS,'+'''0'''+' as Q_vAliqProdPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 1.65)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 1.65)/100) end) as Q_vPIS,

	'+'''0'''+' as R_vBCPISST,'+'''0'''+' as R_pPISST,'+'''0'''+' as R_qBCProdPISST,
	'+'''0'''+' as R_vAliqProdPISST,'+'''0'''+' as R_vPISST,'+'''01'''+' as S_CSTCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end),  

	'+'''7.60'''+' as S_pCOFINS,'+'''0'''+' as S_qBCProdCOFINS,'+'''0'''+' as S_vAliqProdCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 7.60)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 7.60)/100) end),

	'+'''0'''+' as T_vBCCOFINSST,'+'''0'''+' as T_pCOFINSST,
	'+'''0'''+' as T_qBCProdCOFINSST,'+'''0'''+' as T_vAliqProdCOFINSST,
	'+'''0'''+' as T_vCOFINSST,'+'''0'''+' as U_vBCISSQN,'+'''0'''+' as U_vAliqISSQN,
	'+'''0'''+' as U_vISSQN,'+''' '''+' as U_cMunFGISSQN,'+''' '''+' as U_cListServ,
	'+''' '''+' as U_cSitTrib,'+''' '''+' as V_infAdProd 
	FROM produtoloja (NOLOCK), NFItens (NOLOCK) 
	WHERE PR_Referencia = Referencia AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+ 
	' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF)) +' Order by H_nItem'

	--select * from nfe_prod whe

	Print @SQL 
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_total (eLoja,eNF,eSerie,vBCICMS,vICMS,vBCST,vST,vProd,vFrete,vSeg,vDesc,vII,vIPI,
	vCOFINS,vOutro,vNF,vServ,vBCISSQ,vISS,vPIS,vCOFINSISSQ,vRetPIS,vRetCOFINS,vRetCSLL,vBCIRRF,
	vIRRF,vBCRetPrev,vRetPrev,vVICMSDESON)Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,

	(Case When baseicms is null Then 0 else baseicms end), 

	VlrICMS AS vICMS,0 as vBCST,0 as vST,
	vlrmercadoria as vProd,Fretecobr as vFrete,'+''' 0''' + ' as vSeg,Desconto as vDesc,
	'+ '''0''' +' as vII,totalipi as vIPI,(((Totalnota-totalipi) * 7.60)/100) as vCOFINS,0 as vOutro,
	TotalNota as vNF,'+ '''0''' +' as vServ,'+ '''0''' +' as vBCISSQ,'+ '''0''' +' as vISS,
	(((Totalnota - totalipi) * 1.65)/100) as vPIS,'+'''0'''+' as vCOFINSISSQ,'+ '''0''' +' as vRetPIS,
	'+ '''0''' +' as vRetCOFINS,'+ '''0''' +' as vRetCSLL,'+ '''0''' +' as vBCIRRF,
	'+ '''0''' +' as vIRRF,'+ '''0''' +' as vBCRetPrev,'+ '''0''' +' as vRetPrev, ''0'' as vVICMSDESON from NFCapa(Nolock) 
	Where LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+'''' + @Serie + ''''+
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL -- baseicms as vBCICMS,
	Exec (@SQL)
	
	--select * from nfe_prod where enf = 3796
	--select * from NFItens where nf = 3796
	
	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_transp (eLoja,eNF,eSerie,modFrete,CNPJ,CPF,xNome,IE,xEnder,xMun,UF,vServ,vBCRet,pICMSRet,
	vICMSRet,CFOP,cMunFG,placa,UFveic,RNTC,qVol,esq,marca,nVol,pesoL,pesoB,nLacres)
	Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,TipoFrete as modFrete,'+''' '''+' As CNPJ,
	'+''' '''+' as CPF,'+''' '''+' as xNome,'+''' '''+' as IE,'+''' '''+' as xEnder,
	'+''' '''+' as xMun,'+''' '''+' as UF,'+ '''0'''+' as vServ,'+ '''0''' +' as vBCRet,
	'+ '''0''' +' as pICMSRet,'+ '''0''' +' as vICMSRet,'+''' '''+' as CFOP,'+''' '''+' as cMunFG,
	'+''' '''+' as placa,'+''' '''+' as UFveic,'+''' '''+' as RNTC,
	volume as qVol,'+'''VOLUME(S)'''+' as esq,'+''' '''+' as marca,
	'+ '''0''' +' as nVol,pesolq as pesoL,pesobr as pesoB,
	'+ '''0''' +' as nLacres FROM Loja(NOLOCK), NFCapa (NOLOCK)
	Where lojaOrigem = LO_loja And LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	declare @descricao varchar(max)
	--declare @sequencia int
	--declare @sequenciaMaxima int

	SET @Carimbo = ''
	Declare Temp_Carimbo insensitive cursor for
			Select rtrim(LTRIM(CNF_Carimbo))
			  from CarimboNotaFiscal 
			 where CNF_Loja = @Loja 
			   and cnf_serie = @Serie 
			   and CNF_NF = @NF
			 order by CNF_TipoCarimbo desc, CNF_Sequencia 
	Open Temp_Carimbo
	Fetch Next From Temp_Carimbo Into @Descricao
	While @@Fetch_Status = 0  
		Begin

		set @Carimbo = @Carimbo + @Descricao + '  -  '
			Fetch Next From Temp_Carimbo Into @Descricao
		end
	close Temp_Carimbo
	Deallocate Temp_Carimbo

	--set @Carimbo = left(@Carimbo,len(@Carimbo)-2)

	Select @SQL = 'insert into NFe_infAdic (eLoja,eNF,eSerie,infAdFisco,infCpl,xCampoCont,
	xTextoCont,xCampoFisco,xTextoFisco,nProc,indProc) Select LojaOrigem as eLoja,
	NF as eNF,Serie as eSerie,'+''' '''+' as infAdFisco,''PEDIDO: '''+' + RTrim(LTrim(Convert(VarChar(10),numeroped)))+ 
	'+''', VENDEDOR: '''+' + RTrim(LTrim(Convert(VarChar(10),Vendedor)))+'+''', COND PAGTO: '''+' + 
	(Case When (RTrim(LTrim(cp_condicao))) is Null Then '+''' '''+' else cp_condicao end) + '+'''  -  ' + @Carimbo + '''' + ''+' as infcpl,
	'+'''E-MAIL'''+' as xCampoCont, Upper(LO_EmaiLoja) as xTextoCont,
	'+''' '''+' as xCampoFisco,'+''' '''+' as xTextoFisco,
	'+''' '''+' as nProc,'+''' '''+' as indProc from nfCapa(nolock),condicaopagamento(nolock),Loja(nolock)
	where cp_codigo = condpag and cp_id = 1 AND LojaOrigem = LO_Loja AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)
	
END



GO
