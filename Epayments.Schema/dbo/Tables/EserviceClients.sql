CREATE TABLE [dbo].[EserviceClients](
	[EserviceClientId] [int] IDENTITY(1,1) NOT NULL,
	[Gid] [uniqueidentifier] NOT NULL,
	[Alias] [nvarchar](100) NOT NULL,
	[AisName] [nvarchar](100) NOT NULL,
	[ServiceName] [nvarchar](200) NULL,
	[ServiceDescription] [nvarchar](max) NULL,
	[AccountBank] [nvarchar](100) NOT NULL,
	[AccountBIC] [nvarchar](50) NOT NULL,
	[AccountIBAN] [nvarchar](50) NOT NULL,
	[ClientId] [nvarchar](100) NOT NULL,
	[SecretKey] [nvarchar](200) NOT NULL,
	[VposClientId] [int] NULL,
	[DskVposMerchantId] [nvarchar](100) NULL,
	[DskVposMerchantPassword] [nvarchar](max) NULL,
	[IsAuthPassAuthorized] [bit] NOT NULL,
	[IsActive] [bit] NOT NULL,
	[BoricaVposTerminalId] [nvarchar](100) NULL,
	[BoricaVposMerchantId] [nvarchar](100) NULL,
	[BoricaVposRequestSignCertFileName] [nvarchar](max) NULL,
	[BoricaVposRequestSignCertPassword] [nvarchar](max) NULL,
	[BoricaVposRequestSignCertValidTo] [datetime2](7) NULL,
	[BoricaVposRequestSignCertExpMailSend] [bit] NULL,
	[DepartmentId] [int] NULL,
	[FiBankVposAccessKeystoreFilename] [nvarchar](max) NULL,
	[FiBankVposAccessKeystorePassword] [nvarchar](max) NULL,
	[IsEpayVposEnabled] [bit] NOT NULL,
	[BoricaVposRequestSignCertExpHideAdminMsg] [bit] NULL,
	[BoricaEtLogCertificateName] [nvarchar](max) NULL,
	[BoricaEtLogCertificatePassword] [nvarchar](max) NULL,
	[BoricaEtLogCertificateValidTo] [datetime2](7) NULL,
	[BoricaEtLogUsername] [nvarchar](max) NULL,
	[BoricaEtLogPassword] [nvarchar](max) NULL,
	[AggregateToParent] [bit] NOT NULL,
	[DistributionTypeId] [int] NOT NULL,
	[ParentId] [int] NULL,
	[DeliveryAdminstrationId] [nvarchar](50) NULL,
	[DeliveryAdministrationGuid] [nvarchar](max) NULL,
	[IsBoricaVposEnabled] [bit] NOT NULL,
 CONSTRAINT [PK_EserviceClients] PRIMARY KEY CLUSTERED 
(
	[EserviceClientId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Gid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceClients]  WITH CHECK ADD  CONSTRAINT [FK_EserviceClients_Departments] FOREIGN KEY([DepartmentId])
REFERENCES [dbo].[Departments] ([DepartmentId])
GO

ALTER TABLE [dbo].[EserviceClients] CHECK CONSTRAINT [FK_EserviceClients_Departments]
GO
ALTER TABLE [dbo].[EserviceClients]  WITH CHECK ADD  CONSTRAINT [FK_EserviceClients_DistributionTypes] FOREIGN KEY([DistributionTypeId])
REFERENCES [dbo].[DistributionTypes] ([DistributionTypeId])
GO

ALTER TABLE [dbo].[EserviceClients] CHECK CONSTRAINT [FK_EserviceClients_DistributionTypes]
GO
ALTER TABLE [dbo].[EserviceClients]  WITH CHECK ADD  CONSTRAINT [FK_EserviceClients_EserviceClients] FOREIGN KEY([ParentId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[EserviceClients]  WITH CHECK ADD  CONSTRAINT [FK_EserviceClients_VposClients] FOREIGN KEY([VposClientId])
REFERENCES [dbo].[VposClients] ([VposClientId])
GO

ALTER TABLE [dbo].[EserviceClients] CHECK CONSTRAINT [FK_EserviceClients_VposClients]
GO
ALTER TABLE [dbo].[EserviceClients] ADD  DEFAULT ((0)) FOR [AggregateToParent]
GO
ALTER TABLE [dbo].[EserviceClients] ADD  DEFAULT ((1)) FOR [DistributionTypeId]
GO

GO
ALTER TABLE [dbo].[EserviceClients] ADD  DEFAULT ((0)) FOR [IsBoricaVposEnabled]